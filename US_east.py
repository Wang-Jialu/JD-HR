import pandas as pd

# Preprocess invoice sheet
def transform_invoice(input_df):


    input_df['Full Name'] = input_df['Full Name'].str.lower()
    input_df['Full Name'] = input_df['Full Name'].str.replace(' ', '')
    input_df['Markup Rate'] = input_df['Markup Rate'].astype(float)
    input_df['Original Overtime Rate'] = input_df['Original Pay Rate'] * 1.5

    return input_df

def transform_and_rename_columns(input_df):

    input_df["Workers' Full Name"] = input_df["Workers' Full Name"].str.lower()
    input_df["Workers' Full Name"] = input_df["Workers' Full Name"].str.replace(' ', '')

    input_df['Reg Hrs'] = input_df['Weekly Total'].apply(lambda x: x if x <= 40 else 40)
    input_df['OT Hrs'] = input_df['Weekly Total'] - input_df['Reg Hrs']


    return input_df

# join attendance and invoice
def merge_sheets(attendance, pay, building):

    pay = pay[pay['Building'] == building]

    pay_df = transform_invoice(pay)
    
    final_df_for_hr = attendance.merge(pay_df[["Full Name", 'Department', 'Position', 'Updated Base Pay Rate', 'OT Rate', 'Markup Rate',
                                               'Original Pay Rate', 'Original Overtime Rate',  'Pay Rate w/ Markup', 'OT w/ Markup', 'Bonus w/ Markup', 'Bonus OT w/ Markup']],
                                             left_on=["Workers' Full Name"], 
                                             right_on=['Full Name'], how='left')

    final_df_for_hr['Salary'] = (final_df_for_hr['Reg Hrs'] * final_df_for_hr['Original Pay Rate']) + \
                                (final_df_for_hr['OT Hrs'] * final_df_for_hr['Original Overtime Rate'])
    final_df_for_hr['Bonus'] = (final_df_for_hr['Reg Hrs'] * final_df_for_hr['Bonus w/ Markup']) + \
                                (final_df_for_hr['OT Hrs'] * final_df_for_hr['Bonus OT w/ Markup'])
    final_df_for_hr['Payable Salary after markup'] = (final_df_for_hr['Reg Hrs'] * final_df_for_hr['Pay Rate w/ Markup']) + \
                                (final_df_for_hr['OT Hrs'] * final_df_for_hr['OT w/ Markup']) + final_df_for_hr['Bonus']

    
    final_df_for_hr.drop(columns=["Workers' Full Name", 'Full Name', 'Original Pay Rate', 'Original Overtime Rate',  'Pay Rate w/ Markup', 'OT w/ Markup', 'Bonus w/ Markup', 'Bonus OT w/ Markup'], inplace=True)

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

    # Specify the building
    building = '' # (e.g., 'NJ1')
    # Read the attendance spreadsheet
    attendance = pd.read_excel(r"path\to\attendance.xlsx")  # Please replace "path\to\attendance.xlsx" with the actual path to the attendance spreadsheet
    # Read the pay rate spreadsheet
    pay = pd.read_excel(r"path\to\pay_rate.xlsx")  # Please replace "path\to\pay_rate.xlsx" with the actual path to the pay rate spreadsheet
    # Specify the name of the exported Excel file
    excel_output_for_hr = "output.xlsx" # Please replace "output.xlsx" with the desired name of the output Excel file


    attendance_df_for_hr = transform_and_rename_columns(attendance)

    final_df_for_hr = merge_sheets(attendance_df_for_hr, pay, building)
    final_df_for_hr = (final_df_for_hr.pipe(add_agency_subtotals)
                       .pipe(add_totals))
    
    final_df_for_hr.to_excel(excel_output_for_hr, index=False, engine='openpyxl')



if __name__ == '__main__':
    main()