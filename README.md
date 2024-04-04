# JD-HR
Weekly billing calculation for labor outsourcing company

## Overview
This script automates the integration of attendance and payroll data for HR reporting. It encompasses several key functions, including data transformation, merging, and calculation.

Data Transformation and Renaming: Converts date formats, renames columns for clarity, and maps location codes.
Integration by Day of the Week: Pivots attendance data to analyze hours worked each day of the week and calculates weekly totals.
Invoice Data Transformation: Standardizes payroll invoice data, calculates rates, and amounts considering markup rates.
Data Merge and Calculation: Merges attendance and payroll data, calculates salary and bonus amounts considering markup rates.
Agency Subtotals Addition: Aggregates totals for each agency, including weekly totals, hours, and salary.
Total Calculation: Calculates overall totals, including subtotals and grand totals.

## Additional Information
In the payroll data, you can find employee identifiers (employeeNo, Full Name), departmental details (Department, Position), and rate information (Original Pay Rate, Markup Rate, Updated Base Pay Rate, OT Rate, Bonus Rate, etc.). In the attendance data, each entry records an employee's daily check-in/out times.
Due to differences in overtime hour calculation and variations in the original data formats between the US East and US West regions, the script has been divided into two separate scripts.
Please ensure you use the appropriate script corresponding to your region to ensure accurate calculations and processing of the data.

## Instructions
Prepare Data: Ensure both attendance and payroll data are formatted correctly and compatible with the script.
Execute Script: Run the script in a suitable Python environment, providing the necessary data inputs.
Review Output: Verify the generated Excel file.
Customization: Modify the script as needed to accommodate specific data requirements or additional functionalities.

## Using Google Colab
Utilize Google Colab for seamless execution of the script without the need for downloading, installing, and configuring environments.
US East: [Colab Notebook Link](https://colab.research.google.com/drive/1si7ckvqH-zMf4Y_rXo53wL5qoDB57drN?usp=sharing)
US West: [Colab Notebook Link](https://colab.research.google.com/drive/1HshjH2X15EfsJZV6iVHGpCVftKcPgI4F?usp=sharing)
