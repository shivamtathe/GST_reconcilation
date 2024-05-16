import streamlit as st
import pandas as pd
import io
from fuzzywuzzy import process, fuzz
import xlsxwriter

def load_data(purchase_data, gstr2a_data):
    return pd.read_csv(purchase_data), pd.read_csv(gstr2a_data)

def get_best_match(name, choices, scorer=fuzz.partial_ratio, cutoff=75):
    best_match = process.extractOne(name, choices, scorer=scorer)
    return best_match[0] if best_match and best_match[1] > cutoff else name

def reconcile_data(purchase_data, gstr2a_data):
    primary_keys = ['Invoice_Number', 'Tax_Rate', 'Taxable_Amount', 'CGST', 'SGST', 'IGST']
    additional_keys = ['Party_Name', 'GSTIN', 'Invoice_Date']
    reconciliation = pd.merge(
        purchase_data, gstr2a_data, 
        on=primary_keys, 
        how='outer', 
        suffixes=('', '_from_gstr2a'), 
        indicator=True
    )
    merge_labels = {'both': 'Matched', 'left_only': 'Not in GSTR2A', 'right_only': 'Not in Books'}
    reconciliation['_merge'] = reconciliation['_merge'].map(merge_labels)
    for key in additional_keys:
        if key + '_from_gstr2a' in reconciliation.columns:
            reconciliation[key] = reconciliation[key].fillna(reconciliation[key + '_from_gstr2a'])
            reconciliation.drop(columns=[key + '_from_gstr2a'], inplace=True)
    final_columns_order = additional_keys + primary_keys + ['_merge']
    reconciliation = reconciliation[final_columns_order].drop_duplicates()
    return reconciliation

def create_pivot_summary(purchase_data, gstr2a_data):
    gstr2a_parties = gstr2a_data['Party_Name'].tolist()
    purchase_data['Party_Name'] = purchase_data['Party_Name'].apply(get_best_match, args=(gstr2a_parties,))
    book_aggregate = purchase_data.groupby('Party_Name')[['Taxable_Amount', 'CGST', 'SGST', 'IGST']].sum().reset_index()
    gstr2a_aggregate = gstr2a_data.groupby('Party_Name')[['Taxable_Amount', 'CGST', 'SGST', 'IGST']].sum().reset_index()
    comparison = pd.merge(book_aggregate, gstr2a_aggregate, on='Party_Name', how='outer', suffixes=('_books', '_gstr2a'))
    for col in ['Taxable_Amount', 'CGST', 'SGST', 'IGST']:
        comparison[col + '_diff'] = comparison[col + '_books'].fillna(0) - comparison[col + '_gstr2a'].fillna(0)
    return comparison

def generate_excel(reconciliation_data, pivot_summary):
    # Change header of column J to "Remarks"
    reconciliation_data.columns.values[9] = 'Remarks'
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        reconciliation_data.to_excel(writer, sheet_name='Reconciliation Details', index=False)
        pivot_summary.to_excel(writer, sheet_name='Pivot Summary', index=False)
        
        workbook = writer.book
        
        # Header format
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#D7E4BC',
            'border': 1
        })
        
        # Reconciliation Details sheet
        worksheet_reconciliation = writer.sheets['Reconciliation Details']
        
        # Make headers bold and set column width based on header name
        for col_num, value in enumerate(reconciliation_data.columns.values):
            worksheet_reconciliation.write(0, col_num, value, header_format)
            column_width = len(value) + 2
            worksheet_reconciliation.set_column(col_num, col_num, column_width)
        
        # Pivot Summary sheet
        worksheet_pivot = writer.sheets['Pivot Summary']
        
        # Headers for the Pivot Summary based on provided example
        pivot_headers = ['Party Name', 'Taxable Amount (Books)', 'CGST (Books)', 'SGST (Books)', 'IGST (Books)',
                         'Taxable Amount (GSTR2A)', 'CGST (GSTR2A)', 'SGST (GSTR2A)', 'IGST (GSTR2A)',
                         'Taxable Amount Diff', 'CGST Diff', 'SGST Diff', 'IGST Diff']
        
        # Make headers bold and set column width based on header name
        for col_num, header in enumerate(pivot_headers):
            worksheet_pivot.write(0, col_num, header, header_format)
            column_width = len(header) + 2
            worksheet_pivot.set_column(col_num, col_num, column_width)

    output.seek(0)  # Reset the buffer
    return output

def app():
    st.title("GST Reconciliation Tool")
    uploaded_file_purchase = st.file_uploader("Choose a file for purchase data", type='csv')
    uploaded_file_gsmart = st.file_uploader("Choose a file for GSTR-2A data", type='csv')
    if uploaded_file_purchase is not None and uploaded_file_gsmart is not None:
        purchase_data, gsmart_data = load_data(uploaded_file_purchase, uploaded_file_gsmart)
        reconciliation_data = reconcile_data(purchase_data, gsmart_data)
        pivot_summary = create_pivot_summary(purchase_data, gsmart_data)
        st.write("Reconciliation Data", reconciliation_data)
        st.write("Pivot Summary", pivot_summary)
        export_button = st.button("Generate Report")
        if export_button:
            output = generate_excel(reconciliation_data, pivot_summary)
            st.download_button(
                label="Download Excel report",
                data=output,
                file_name="reconciliation_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    app()
