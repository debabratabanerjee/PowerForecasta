# Power Forecasting Optimization Application

Developed by: Debabrata Banerjee, Software Engineer
contact: info@slideway.dev

## Overview

The Power Forecasting Optimization Application is a proprietary tool that enables efficient optimization of power generation to meet specific demand blocks. It calculates the most cost-effective way to meet power demand using available generators, including "must-run" and optional generators. The application helps to reduce costs while ensuring power availability.

## Features

[1]. Optimizes power generation to meet the demand using the most cost-effective generators.
[2]. Uses both "must-run" and "available" generators for maximum efficiency.
[3]. Incorporates Open Access (OA) and banked energy adjustments to demand.
[4]. Provides detailed logging of each step and result.
[5]. Outputs results in an easy-to-read Excel file.

## Requirements

[1]. Python 3.11+
[2]. Required libraries:
    [2.1]. pandas
    [2.2]. pulp
    [2.3]. tkinter
    [2.4]. tkcalendar
    [2.5]. babel
    [2.6]. numpy
    [2.7]. cx_Freeze

All these dependencies are included in the executable version created using cx_Freeze.

## Installation & Setup

[1]. Download the executable from the release folder.
[2]. Ensure that the necessary Excel files (Demand.xlsx, generator_mod.xlsx, Rate.xlsx, OA.xlsx, Bank.xlsx) are placed in the root directory.
[3]. Ensure the generator_data directory contains the generator-specific Excel files.
[4]. The application also requires the cbc.exe solver for optimization, which is bundled with the application in the solvers folder.

## Usage Instructions

### Running the Application

[1]. Starting the Application:
    [1.1]. When you run the application, a command prompt (terminal window) will open, and the application will take about 2 minutes to load the necessary data into memory.
    [1.2]. You can track the progress of the data loading through the command prompt.

[2]. Changing or Adding Data:
    [2.1]. You can change values in the existing Excel files (like generator_mod.xlsx or files in the generator_data folder).
    [2.2]. You can also add new Excel files to the generator_data directory before running the application.
    [2.3]. Ensure any new Excel files are formatted consistently with the existing ones.

[3]. Running the Optimization:
    [3.1]. After the data has been loaded, a UI will appear, allowing you to input:
        [3.1.1]. Start Date and End Date: The range of dates for which you want to run the optimization.
        [3.1.2]. Start Block and End Block: The range of time blocks (1-96) for which you want to run the optimization.
    [3.2]. Click the "Run Custom Date and Block Range" button to run the optimization.

[4]. Output & Logs:
    [4.1]. Output Folder: Results will be saved as an Excel file in the output folder. The filename will be based on the selected date and block range.
    [4.2]. Logs Folder: If any errors occur or additional information is required, refer to the log files in the logs folder for detailed error reports or process tracking.

### Important Notes

[1]. Command Prompt Logs: Keep an eye on the command prompt to see real-time terminal logs, which will give you information about data loading, processing, and any potential issues.
[2]. Error Handling: If errors occur during optimization (for example, missing data in an Excel file), the logs will contain detailed error messages. Ensure that all necessary columns (like Date, available_power, variable_cost, etc.) are present in each Excel file.

## Troubleshooting

If the application fails to open or stops unexpectedly, check the following:
[1]. Ensure all Excel files are correctly placed and formatted.
[2]. Check the logs for detailed error messages.
[3]. Ensure you allow enough time for the data to load (up to 2 minutes) before the UI becomes visible.

## FAQ

[1]. Can I add new generators by adding Excel files?
    A: Yes, you can add new generator data by placing new Excel files in the generator_data directory before running the application.

[2]. How long does it take to load the application?
    A: The application typically takes around 2 minutes to load all the necessary data into memory before showing the user interface.

[3]. Where are the results stored?
    A: The results are saved in the output folder in an Excel format after the optimization process completes.

## License

This application is proprietary software and may not be distributed, copied, or modified without explicit permission.