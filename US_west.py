import pandas as pd

def transform_and_rename_columns(input_df):
    # Change date format
    input_df['checkDate'] = pd.to_datetime(input_df['checkDate'])
    input_df['date'] = input_df['checkDate'].dt.date


    # Change column names
    selected_cols = ['warehouseNo', 'groupName', 'employeeNo', 'companyShortName', 'lastName', 'firstName', 'allTime', 'date']
    result_df = input_df[selected_cols].rename(columns={
        'employeeNo': 'USID #',
        'companyShortName': 'Agency',
        'lastName': "Workers' Last Name",
        'firstName': "Workers' First Name",
        'groupName': "GROUP Name",
        'warehouseNo': 'Building'
    })

    # Convert to lowercase
    result_df["Workers' Last Name"] = result_df["Workers' Last Name"].str.lower()
    result_df["Workers' First Name"] = result_df["Workers' First Name"].str.lower()
    result_df['FullName'] = result_df["Workers' First Name"] + ' ' + result_df["Workers' Last Name"]
    result_df['FullName'] = result_df['FullName'].str.replace(' ', '')

        # Define the mapping for location values
    location_mapping = {
    'C0000000578': 'CA1',
    'C0000000579': 'CA2',
    'C0000002520': 'CA3'
    }

    # Apply the mapping to the Location column
    result_df['Building'] = result_df['Building'].map(location_mapping)

    return result_df


# Integration function: Pivot by day of the week
def integrate_by_day_of_week(input_df):

    pivot_df = input_df.pivot_table(index=['USID #', 'Agency', 'Building', 'GROUP Name', "Workers' Last Name", "Workers' First Name", "FullName"],
                                    columns='date', values='allTime', aggfunc='sum', fill_value=0)

    # Calculate weekly totals and hours
    pivot_df['Weekly Total'] = pivot_df.sum(axis=1)

    pivot_df = pivot_df.loc[pivot_df['Weekly Total'] != 0]

    pivot_df['Reg Hrs'] = 0
    pivot_df['OT Hrs'] = 0

    for idx, row in pivot_df.iterrows():
        consecutive_days = 0 # Track consecutive working days
        for col in pivot_df.columns[:-3]: # Iterate through each day
            hours = row[col]

            if hours > 0: # If the worker worked that day
                consecutive_days += 1
                if consecutive_days < 7: # For the first 6 days
                    reg_hours = min(hours, 8)
                    ot_hours = max(hours - 8, 0)
                else: # For the 7th consecutive working day
                    reg_hours = 0 # The first 8 hours are considered as overtime
                    ot_hours = hours

                pivot_df.at[idx, 'Reg Hrs'] += reg_hours
                pivot_df.at[idx, 'OT Hrs'] += ot_hours
            else:
                consecutive_days = 0

    pivot_df['OT Hrs'] += pivot_df.apply(lambda row: row['Reg Hrs'] - 40 if row['Reg Hrs'] >= 40 else 0, axis=1)
    pivot_df['Reg Hrs'] = pivot_df['Reg Hrs'].apply(lambda x: min(x, 40))

    pivot_df.reset_index(inplace=True)

    return pivot_df

def transform_invoice(input_df):

    input_df['Full Name'] = input_df['Full Name'].str.lower()
    input_df['Full Name'] = input_df['Full Name'].str.replace(' ', '')
    input_df['Markup Rate'] = input_df['Markup Rate'].astype(float)
    input_df['Original Overtime Rate'] = input_df['Original Pay Rate'] * 1.5
    input_df['OT Rate'] = input_df['Updated Base Pay Rate']*1.5

    return input_df

# join attendance and invoice
def merge_sheets(attendance, pay):

    pay_df = transform_invoice(pay)
    
    final_df_for_hr = attendance.merge(pay_df[['employeeNo', "Full Name", 'Department', 'Position', 'Updated Base Pay Rate', 'OT Rate', 'Markup Rate',
                                               'Original Pay Rate', 'Original Overtime Rate',  'Base Markup', 'BaseOT Markup', 'Bonus Markup', 'Bonus OT Markup']],
                                             left_on=["USID #"], 
                                             right_on=['employeeNo'], how='left')

    final_df_for_hr['Salary'] = (final_df_for_hr['Reg Hrs'] * final_df_for_hr['Original Pay Rate']) + \
                                (final_df_for_hr['OT Hrs'] * final_df_for_hr['Original Overtime Rate'])
    final_df_for_hr['Bonus'] = (final_df_for_hr['Reg Hrs'] * final_df_for_hr['Bonus Markup']) + \
                                (final_df_for_hr['OT Hrs'] * final_df_for_hr['Bonus OT Markup'])
    final_df_for_hr['Payable Salary after markup'] = (final_df_for_hr['Reg Hrs'] * final_df_for_hr['Base Markup']) + \
                                (final_df_for_hr['OT Hrs'] * final_df_for_hr['BaseOT Markup']) + final_df_for_hr['Bonus']

    
    final_df_for_hr.drop(columns=['employeeNo', "FullName", 'Full Name', 'Original Pay Rate', 'Original Overtime Rate',  'Base Markup', 'BaseOT Markup', 'Bonus Markup', 'Bonus OT Markup'], inplace=True)

    desired_column_order = ['Building', 'Department', 'Agency', "Workers' Last Name", "Workers' First Name", 'Position']

    final_df_for_hr = final_df_for_hr[desired_column_order + [col for col in final_df_for_hr.columns if col not in desired_column_order]]

    return final_df_for_hr

# Agency subtotals function
def add_agency_subtotals(input_df):

    agency_totals = input_df.groupby('Agency')[['Weekly Total', 'Reg Hrs', 'OT Hrs', 'Salary', 'Payable Salary after markup', 'Bonus']].sum()

    # Add agency and "Subtotal" to output
    agency_totals['Agency'] = agency_totals.index
    agency_totals['Building'] = "Subtotal"

    # Concatenate the original df and agency subtotals
    subtotal_df = pd.concat([input_df, agency_totals], ignore_index=True)
    subtotal_df.sort_values(by=['Agency', 'Department'], inplace=True)

    return subtotal_df

# Calculate totals
def add_totals(input_df):

    # Calculate totals base on subtotals
    total_row = input_df[input_df['Building'] == 'Subtotal'].sum(numeric_only=True)
    total_row['Building'] = 'Total'
    total_dict = total_row.to_dict()
    final_df = pd.concat([input_df, pd.DataFrame([total_dict])], ignore_index=True)

    return final_df

def main():

    
    # Read the attendance spreadsheet
    attendance = pd.read_excel(r"path\to\attendance.xlsx")  # Please replace "path\to\attendance.xlsx" with the actual path to the attendance spreadsheet
    # Read the pay rate spreadsheet
    pay = pd.read_excel(r"path\to\pay_rate.xlsx")  # Please replace "path\to\pay_rate.xlsx" with the actual path to the pay rate spreadsheet
    # Specify the name of the exported Excel file
    excel_output_for_hr = "output.xlsx" # Please replace "output.xlsx" with the desired name of the output Excel file

    attendance_df_for_hr = (attendance.pipe(transform_and_rename_columns)
                            .pipe(integrate_by_day_of_week))

    final_df_for_hr = merge_sheets(attendance_df_for_hr, pay)
    final_df_for_hr = (final_df_for_hr.pipe(add_agency_subtotals)
                        .pipe(add_totals))

    final_df_for_hr.to_excel(excel_output_for_hr, index=False, engine='openpyxl')


if __name__ == '__main__':
    main()