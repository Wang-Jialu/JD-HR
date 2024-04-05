import pandas as pd
from decimal import Decimal, ROUND_HALF_UP

# Preprocess invoice sheet
def transform_invoice(input_df):

    input_df['Full Name'] = input_df['Full Name'].str.lower()
    input_df['Full Name'] = input_df['Full Name'].str.replace(' ', '')
    input_df['Pay Rate'] = input_df['Pay Rate'].replace('[\$,]', '', regex=True).astype(float)
    input_df['Markup Rate'] = input_df['Markup Rate'].astype(float)
    input_df['Reg after markup'] = input_df.apply(lambda x: Decimal(str(x['Pay Rate'] * (x['Markup Rate'] + 1)))\
                                                  .quantize(Decimal("0.01"), rounding=ROUND_HALF_UP), axis=1)
    input_df['Overtime Rate'] = input_df['Pay Rate'] * 1.5
    input_df['OT after markup'] = (input_df['Overtime Rate'] * (input_df['Markup Rate'] + 1)).round(2)
    input_df['OT after markup'] = input_df.apply(lambda x: Decimal(str(x['Overtime Rate'] * (x['Markup Rate'] + 1)))\
                                                 .quantize(Decimal("0.01"), rounding=ROUND_HALF_UP), axis=1)

    return input_df

# join attendance and invoice
def merge_sheets(attendance, pay):

    pay_df = transform_invoice(pay)

    attendance['FullName'] = attendance["Worker's First Name"] + ' ' + attendance["Worker's Last Name"]
    attendance['FullName'] = attendance['FullName'].str.replace(' ', '')
    attendance['FullName'] = attendance['FullName'].str.lower()
    attendance['Reg'] = attendance['Weekly Total'].apply(lambda x: x if x <= 40 else 40)
    attendance['OT'] = attendance['Weekly Total'] - attendance['Reg']
    
    final_df_for_hr = attendance.merge(pay_df[['Department', 'Position', 'Pay Rate', 'Overtime Rate', 'Markup Rate', 'Reg after markup', 'OT after markup', 'Full Name']],
                                             left_on=["FullName"], 
                                             right_on=['Full Name'], how='left')
    
    final_df_for_hr.drop(columns=['FullName', 'Full Name'], inplace=True)

    final_df_for_hr['Salary'] = (final_df_for_hr['Reg'] * final_df_for_hr['Pay Rate']) + \
                                (final_df_for_hr['OT'] * final_df_for_hr['Overtime Rate'])
    final_df_for_hr['Bill after Markup'] = (final_df_for_hr['Reg after markup'].astype(float) * final_df_for_hr['Reg']) + \
                                (final_df_for_hr['OT after markup'].astype(float) * final_df_for_hr['OT'])
    
    final_df_for_hr['Bill after Markup'] = final_df_for_hr['Bill after Markup'].apply(
    lambda x: Decimal(str(x)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
)

    


    return final_df_for_hr

# Building subtotals function
def add_building_subtotals(input_df):

    building_totals = input_df.groupby('Building')[['Weekly Total', 'Reg', 'OT', 'Salary', 'Bill after Markup']].sum()

    # Add agency and "Subtotal" to output
    building_totals['Building'] = building_totals.index
    building_totals['Agency'] = "Subtotal"

    # Concatenate the original df and agency subtotals
    subtotal_df = pd.concat([input_df, building_totals], ignore_index=True)
    subtotal_df.sort_values(by=['Building', "Department"], inplace=True)

    return subtotal_df

# Calculate totals
def add_totals(input_df):

    # Calculate totals base on subtotals
    total_row = input_df[input_df['Agency'] == 'Subtotal'].sum(numeric_only=True)
    total_row['Agency'] = 'Total'
    total_dict = total_row.to_dict()
    final_df = pd.concat([input_df, pd.DataFrame([total_dict])], ignore_index=True)

    return final_df

def main():

    # Read excel
    attendance = pd.read_excel(r"path\to\attendance.xlsx")
    pay = pd.read_excel(r"path\to\pay_rate.xlsx")

    # Name of the exported excel file 
    excel_output_for_hr = "output.xlsx"

    final_df_for_hr = merge_sheets(attendance, pay)
    final_df_for_hr = (final_df_for_hr.pipe(add_building_subtotals)
                       .pipe(add_totals))
    
    final_df_for_hr.to_excel(excel_output_for_hr, index=False, engine='openpyxl')



if __name__ == '__main__':
    main()