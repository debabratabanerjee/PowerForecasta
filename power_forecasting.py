import os
import pandas as pd
import pulp
import logging
import traceback
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import tkinter as tk
from tkinter import messagebox, filedialog
from tkcalendar import DateEntry
import sys


# Constants for better readability
DATE_FORMAT = "%d-%m-%Y"
DEMAND_CONVERSION_FACTOR = 1000 * 0.25  # For MW to kWh conversion

def resource_path(relative_path):
    """ Get absolute path to resource, works for both development and PyInstaller packaged apps """
    try:
        base_path = sys._MEIPASS  # PyInstaller temporary folder
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# Update file paths to use the resource_path function
DEMAND_FILE = resource_path("Demand.xlsx")
GENERATOR_FILE = resource_path("generator_mod.xlsx")
GRID_COST_FILE = resource_path("Rate.xlsx")
OA_FILE = resource_path("OA.xlsx")
BANK_FILE = resource_path("Bank.xlsx")
GENERATOR_HISTORICAL_DATA_DIR = resource_path("generator_data/")
MUST_RUN_TYPE = "Must run"
AVAILABLE_TYPE = "Available"
OUTPUT_DIR = resource_path("output")
NUM_WORKERS = 8

# Create a timestamped log file
timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
LOG_FILE = os.path.join("logs", f'power_optimization_log_{timestamp}.txt')

# Ensure necessary directories exist
os.makedirs("logs", exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Initialize logging with the timestamped log file
logging.basicConfig(filename=LOG_FILE, level=logging.INFO,
                    format='%(asctime)s - %(message)s', filemode='w')
logger = logging.getLogger()

# Helper function to validate file existence
def validate_file_exists(filepath):
    if not os.path.exists(filepath):
        logger.error(f"File not found: {filepath}")
        raise FileNotFoundError(f"File not found: {filepath}")

# Preload all generator data into memory
def preload_generator_data():
    generator_data = {}
    for filename in os.listdir(GENERATOR_HISTORICAL_DATA_DIR):
        if filename.endswith(".xlsx"):
            generator_code = filename.replace(".xlsx", "")
            file_path = os.path.join(GENERATOR_HISTORICAL_DATA_DIR, filename)
            df = pd.read_excel(file_path)

            if 'Date' not in df.columns:
                logger.error(f"Date column missing in file: {filename}")
                continue  # Skip this file

            df['Date'] = pd.to_datetime(df['Date'], format=DATE_FORMAT, errors='coerce')
            generator_data[generator_code] = df

    return generator_data

# Preload grid cost, Open Access, and Bank data into memory
def preload_additional_data():
    validate_file_exists(GRID_COST_FILE)
    validate_file_exists(OA_FILE)
    validate_file_exists(BANK_FILE)

    df_grid_cost = pd.read_excel(GRID_COST_FILE)
    df_grid_cost['Date'] = pd.to_datetime(df_grid_cost['Date'], format=DATE_FORMAT, errors='coerce')

    df_oa = pd.read_excel(OA_FILE)
    df_oa['Date'] = pd.to_datetime(df_oa['Date'], format=DATE_FORMAT, errors='coerce')

    df_bank = pd.read_excel(BANK_FILE)
    df_bank['Date'] = pd.to_datetime(df_bank['Date'], format=DATE_FORMAT, errors='coerce')

    return df_grid_cost, df_oa, df_bank

# Validate generator data
def validate_generator_data(df):
    if df[['available_power', 'variable_cost']].isnull().values.any():
        missing_values = df[df[['available_power', 'variable_cost']].isnull()]
        logger.error("Missing values in the 'available_power' or 'variable_cost' columns.")
        logger.error(missing_values)
        raise ValueError("Invalid generator data. Check the log file for details.")

# Adjust demand with Open Access
def adjust_demand_with_open_access(date, block, demand_mw, df_oa):
    row = df_oa[df_oa['Date'] == pd.to_datetime(date)]
    if row.empty:
        return demand_mw * DEMAND_CONVERSION_FACTOR, 0  # Convert MW to kWh, OA = 0
    block_column = block if block in row.columns else str(block)
    if block_column not in row.columns:
        return demand_mw * DEMAND_CONVERSION_FACTOR, 0  # Convert MW to kWh, OA = 0
    oa_value_mw = row[block_column].values[0]
    adjusted_demand_mw = demand_mw - oa_value_mw
    adjusted_demand_kwh = adjusted_demand_mw * DEMAND_CONVERSION_FACTOR  # Convert MW to kWh
    logger.info(f"Adjusted demand after Open Access for {date} Block {block}: {adjusted_demand_kwh} kWh")
    return adjusted_demand_kwh, oa_value_mw

# Adjust demand with Banked energy
def adjust_demand_with_bank(date, block, demand_kwh, df_bank):
    row = df_bank[df_bank['Date'] == pd.to_datetime(date)]
    if row.empty:
        return demand_kwh, 0
    block_column = block if block in row.columns else str(block)
    if block_column not in row.columns:
        return demand_kwh, 0
    bank_value_mw = row[block_column].values[0]
    bank_value_kwh = bank_value_mw * DEMAND_CONVERSION_FACTOR  # Convert MW to kWh
    adjusted_demand = demand_kwh + bank_value_kwh
    logger.info(f"Adjusted demand after Bank for {date} Block {block}: {adjusted_demand} kWh (Bank adjustment: {bank_value_kwh} kWh)")
    return adjusted_demand, bank_value_mw

# Load grid cost for the given block
def load_grid_cost(date, block, df_grid_cost):
    row = df_grid_cost[df_grid_cost['Date'] == pd.to_datetime(date)]
    if row.empty:
        raise ValueError(f"No grid cost data found for date {date}")
    block_column = block if block in row.columns else str(block)
    
    if block_column not in row.columns:
        raise ValueError(f"No grid cost data found for block {block} on date {date}")
    
    return row[block_column].values[0]

def setup_optimization_problem(available_generators):
    prob = pulp.LpProblem("Maximize_Generator_Use", pulp.LpMaximize)
    gen_power_vars = {}
    for _, g in available_generators.iterrows():
        gen_power_vars[g["name"]] = pulp.LpVariable(g["name"], lowBound=0, upBound=0)

    # Specify the path to the local CBC solver
    solver_path = resource_path("solvers/cbc.exe")
    pulp.LpSolverDefault = pulp.COIN_CMD(path=solver_path)

    return prob, gen_power_vars

# Preload generator and grid data
preloaded_generators = preload_generator_data()
df_grid_cost, df_oa, df_bank = preload_additional_data()

# Optimization function for demand fulfillment
def optimize_power_for_demand(demand_kwh, date, block, generators, prob, gen_power_vars):
    must_run_power, must_run_cost, must_run_details = calculate_must_run_power(generators, date, block)

    # Adjust demand after must-run generators
    remaining_demand = demand_kwh - must_run_power
    logger.info(f"Block {block} on {date}: Original Demand: {demand_kwh} kWh, After Must-Run Generators: {remaining_demand} kWh")

    if remaining_demand < 0:
        logger.error(f"Invalid remaining demand: {remaining_demand} kWh")
        return None

    try:
        grid_cost = load_grid_cost(date, block, df_grid_cost)
        logger.info(f"Grid Cost for Block {block} on {date}: {grid_cost} INR/kWh")
    except Exception as e:
        logger.error(f"Error loading grid cost: {e}")
        return None

    total_cost = must_run_cost
    exchange_quantity = 0
    available_gen_demand_met = 0

    if remaining_demand > 0:
        available_gen_demand_met, available_gen_cost, available_gen_details = optimize_available_generators(
            remaining_demand, date, block, generators, prob, gen_power_vars
        )
        total_cost += available_gen_cost
        logger.info(f"Available Generator Demand Met: {available_gen_demand_met} kWh, Cost: {available_gen_cost} INR")

        if available_gen_demand_met < remaining_demand:
            exchange_quantity = remaining_demand - available_gen_demand_met
            total_cost += exchange_quantity * grid_cost
            logger.info(f"Using Exchange Grid for {exchange_quantity} kWh at cost {exchange_quantity * grid_cost} INR")

    return {
        "Date": date,
        "Block": block,
        "Demand": demand_kwh,
        "Total Demand Met": demand_kwh,
        "Must-Run Generators Used (kWh)": must_run_power,
        "Available Generators Used (kWh)": available_gen_demand_met,
        "Grid Consumption (kWh)": exchange_quantity,
        "Total Cost": total_cost,
        "Grid Rate (INR/kWh)": grid_cost,
        "OA Used (MW)": None,
        "Bank Adjustment (MW)": None,
        "Must-Run Generator Details": must_run_details,
        "Available Generator Details": available_gen_details
    }

# Calculate must-run power (preload generator data)
def calculate_must_run_power(generators, date, block):
    must_run_power = 0
    must_run_cost = 0
    must_run_details = []

    for _, g in generators[generators['Type of Plant'] == MUST_RUN_TYPE].iterrows():
        generator_code = g["Code"]
        generator_data = preloaded_generators.get(generator_code)
        if generator_data is not None:
            power_mw = generator_data.loc[generator_data['Date'] == date, block].values[0]
            power_kwh = power_mw * DEMAND_CONVERSION_FACTOR  # Convert MW to kWh
            must_run_power += power_kwh
            must_run_cost += power_kwh * g['variable_cost']
            must_run_details.append(f"{generator_code} | {power_kwh} kWh | {power_kwh * g['variable_cost']:.2f} INR")

    return must_run_power, must_run_cost, "\n".join(must_run_details)

# Optimize available generators (maximize their use)
def optimize_available_generators(remaining_demand, date, block, generators, prob, gen_power_vars):
    available_generators = generators[generators['Type of Plant'] == AVAILABLE_TYPE]
    for _, g in available_generators.iterrows():
        generator_code = g["Code"]
        generator_data = preloaded_generators.get(generator_code)
        if generator_data is not None:
            power_mw = generator_data.loc[generator_data['Date'] == date, block].values[0]
            power_kwh = power_mw * DEMAND_CONVERSION_FACTOR  # Convert MW to kWh
            gen_power_vars[g["name"]].upBound = power_kwh

    prob.setObjective(pulp.lpSum(gen_power_vars[g['name']] for _, g in available_generators.iterrows()))
    prob += pulp.lpSum(gen_power_vars[g['name']] for _, g in available_generators.iterrows()) <= remaining_demand
    prob.solve()

    if pulp.LpStatus[prob.status] == "Optimal":
        available_gen_demand_met = sum(gen_power_vars[g["name"]].varValue for _, g in available_generators.iterrows())
        available_gen_cost = sum(gen_power_vars[g["name"]].varValue * g["variable_cost"] for _, g in available_generators.iterrows())
        available_gen_details = [f"{g['Code']} | {gen_power_vars[g['name']].varValue:.2f} kWh | {gen_power_vars[g['name']].varValue * g['variable_cost']:.2f} INR" for _, g in available_generators.iterrows()]
        return available_gen_demand_met, available_gen_cost, "\n".join(available_gen_details)
    else:
        logger.error("No feasible solution found for available generators.")
        return 0, 0, ""

# Main function to run optimization for a block
def process_block(demand, date, block, df_generators):
    try:
        demand_after_oa, oa_value_mw = adjust_demand_with_open_access(date, block, demand, df_oa)
        demand_after_bank, bank_value_mw = adjust_demand_with_bank(date, block, demand_after_oa, df_bank)

        logger.info(f"Starting optimization for demand after adjustments: {demand_after_bank} kWh on {date} Block {block}")

        prob, gen_power_vars = setup_optimization_problem(df_generators)

        result = optimize_power_for_demand(demand_after_bank, date, block, df_generators, prob, gen_power_vars)
        result["OA Used (MW)"] = oa_value_mw
        result["Bank Adjustment (MW)"] = bank_value_mw
        return result
    except Exception as e:
        logger.error(f"Error processing block {block} on {date}: {traceback.format_exc()}")
        return None

# Main function with parallel execution
def run_normal(start_date_str, end_date_str, start_block, end_block):
    validate_file_exists(DEMAND_FILE)
    validate_file_exists(GENERATOR_FILE)

    df_demand = pd.read_excel(DEMAND_FILE)
    df_generators = pd.read_excel(GENERATOR_FILE)
    validate_generator_data(df_generators)

    start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
    end_date = datetime.strptime(end_date_str, "%Y-%m-%d")

    df_filtered = df_demand[(df_demand['Date'] >= start_date) & (df_demand['Date'] <= end_date)]

    results = []
    futures = []

    with ThreadPoolExecutor(max_workers=NUM_WORKERS) as executor:
        for idx, row in df_filtered.iterrows():
            date = row["Date"]
            for block in range(start_block, end_block + 1):
                demand_mw = row[block]
                futures.append(executor.submit(process_block, demand_mw, date, block, df_generators))

        for future in as_completed(futures):
            result = future.result()
            if result:
                results.append(result)

    # Convert results into a DataFrame and structure as per the required output
    df_results = pd.DataFrame(results)

    # Calculate new columns based on the results
    df_results['Net Demand (kWh)'] = df_results['Demand'] - df_results['OA Used (MW)'] * DEMAND_CONVERSION_FACTOR + df_results['Bank Adjustment (MW)'] * DEMAND_CONVERSION_FACTOR
    df_results['Grid Rate (INR/kWh)'] = df_results['Grid Rate (INR/kWh)']

    # Reorder and rename columns to match the required format
    df_results = df_results[[
        "Date", "Block", "Demand", 
        "OA Used (MW)", "Bank Adjustment (MW)", "Net Demand (kWh)", 
        "Must-Run Generators Used (kWh)", "Must-Run Generator Details",
        "Available Generators Used (kWh)", "Available Generator Details",
        "Grid Consumption (kWh)", "Grid Rate (INR/kWh)", "Total Cost"
    ]]

    # Save the results to Excel in the output directory
    output_file_path = os.path.join(OUTPUT_DIR, f"power_optimization_results_ui_{start_date_str}_to_{end_date_str}_block_{start_block}_to_{end_block}.xlsx")
    df_results.to_excel(output_file_path, index=False)

    logger.info(f"Results saved to {output_file_path}")

# Custom run function for a custom date and block range
def custom_run(start_date, end_date, start_block, end_block):
    run_normal(start_date, end_date, start_block, end_block)

# Tkinter UI for date and block range picker
def run_ui():
    def run_custom():
        start_date = start_entry.get_date().strftime("%Y-%m-%d")
        end_date = end_entry.get_date().strftime("%Y-%m-%d")
        start_block = int(start_block_entry.get())
        end_block = int(end_block_entry.get())
        custom_run(start_date, end_date, start_block, end_block)
        messagebox.showinfo("Info", f"Custom run from {start_date} to {end_date}, Blocks {start_block} to {end_block} complete!")

    root = tk.Tk()
    root.title("Power Forecasting Optimization")

    start_label = tk.Label(root, text="Start Date")
    start_label.pack(pady=5)
    start_entry = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2)
    start_entry.pack(pady=5)

    end_label = tk.Label(root, text="End Date")
    end_label.pack(pady=5)
    end_entry = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2)
    end_entry.pack(pady=5)

    start_block_label = tk.Label(root, text="Start Block (1-96)")
    start_block_label.pack(pady=5)
    start_block_entry = tk.Entry(root)
    start_block_entry.pack(pady=5)

    end_block_label = tk.Label(root, text="End Block (1-96)")
    end_block_label.pack(pady=5)
    end_block_entry = tk.Entry(root)
    end_block_entry.pack(pady=5)

    custom_run_button = tk.Button(root, text="Run Custom Date and Block Range", command=run_custom)
    custom_run_button.pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    try:
        run_ui()
    except Exception as e:
        logger.critical(f"Unhandled exception: {traceback.format_exc()}")
        messagebox.showerror("Fatal Error", f"An unexpected error occurred: {e}")
        sys.exit(1)
