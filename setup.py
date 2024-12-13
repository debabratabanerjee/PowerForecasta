from cx_Freeze import setup, Executable
import os

# List of additional files to be included
include_files = [
    ("solve rs/cbc.exe", "solvers/cbc.exe"),
    ("Demand.xlsx", "Demand.xlsx"),
    ("generator_mod.xlsx", "generator_mod.xlsx"),
    ("Rate.xlsx", "Rate.xlsx"),
    ("OA.xlsx", "OA.xlsx"),
    ("Bank.xlsx", "Bank.xlsx"),
    ("generator_data/", "generator_data/"),
    ("output/", "output/")
]

# Packages to include
packages = ["pandas", "numpy", "tkinter", "tkcalendar", "pulp", "babel.numbers"]

# Setup configuration for cx_Freeze
setup(
    name="PowerForecasting",
    version="1.0",
    description="Power Forecasting Application",
    options={
        "build_exe": {
            "packages": packages,
            "include_files": include_files,
        }
    },
    executables=[Executable("power_forecasting.py", base=None)]
)
