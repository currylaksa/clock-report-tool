import streamlit as st
import pandas as pd
from io import BytesIO

# Page Configuration
st.set_page_config(page_title="Eastern Region Clock Report", layout="centered")

st.title("Eastern Region Clock Report Processor")
st.markdown("""
**Instructions:**
1. Upload the daily **Clock Detail Report** (Excel file).
2. The system will process **ECNB** and **ECMW**.
3. It generates a **Pivot-style** view (grouped and sorted) just like the manual report.
4. **Original sheets are preserved.**
""")

def create_pivot_view(df_input, group_cols):
    """
    Simulates a Pivot Table "Tabular View" by sorting and hiding repeated labels.
    """
    # 1. Sort the data strictly by the grouping order
    # Use fillna to handle empty cells gracefully before sorting
    df_sorted = df_input.fillna("").sort_values(by=group_cols).copy()
    
    # 2. Create a display version where we hide duplicates (Masking)
    # We convert to string to ensure we can write empty strings
    df_display = df_sorted.astype(str).copy()
    
    # We iterate through the columns to mask duplicates, but ONLY if the parent column is also a duplicate.
    # Logic: If Company is same as above, hide Company. 
    #        If Company AND Name are same as above, hide Name.
    #        If Company AND Name AND Account are same as above, hide Account.
    
    # Initialize a tracker for the previous row
    prev_row = {col: None for col in group_cols}
    
    # We need to modify df_display index by index. 
    # It's faster to do this via a list of lists for display purposes.
    formatted_rows = []
    
    for _, row in df_sorted.iterrows():
        current_row = []
        is_parent_same = True # Assumption starts true, breaks if a parent differs
        
        for col in group_cols:
            val = row[col]
            # Check if this value matches the previous row AND the parent hierarchy was also the same
            if is_parent_same and val == prev_row[col]:
                current_row.append("") # Hide it (Pivot look)
            else:
                current_row.append(val) # Show it
                is_parent_same = False # Break the chain for child columns
            
            # Update previous row tracker
            prev_row[col] = val
            
        formatted_rows.append(current_row)
        
    return df_sorted, pd.DataFrame(formatted_rows, columns=group_cols)

# 1. File Upload
uploaded_file = st.file_uploader("Upload your clockreport file (xlsx)", type="xlsx")

if uploaded_file:
    try:
        # Load ALL sheets to preserve them
        all_sheets = pd.read_excel(uploaded_file, sheet_name=None)
        
        source_sheet_name = "Clock Detail Report"
        if source_sheet_name not in all_sheets:
            st.error(f"Error: The file must contain a sheet named '{source_sheet_name}'.")
            st.stop()
            
        df_source = all_sheets[source_sheet_name]
        
        # Cleanup Headers
        df_source.columns = df_source.columns.astype(str).str.strip()
        
        # Prepare Output
        output = BytesIO()
        
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # --- FORMATS ---
            # Header: Bold, Light Blue, Border
            header_fmt = workbook.add_format({
                'bold': True, 'text_wrap': False, 'valign': 'vcenter', 'align': 'center',
                'fg_color': '#D9E1F2', 'border': 1
            })
            
            # Data: Border, align left
            data_fmt = workbook.add_format({'border': 1, 'align': 'left'})

            # Bold Data: Border, align left, Bold (for Main Categories)
            data_bold_fmt = workbook.add_format({'border': 1, 'align': 'left', 'bold': True})
            
            # Orange Highlight (Duplicate DU IDs)
            orange_fmt = workbook.add_format({'bg_color': '#FFC000', 'font_color': '#000000', 'border': 1})
            
            # 1. Write Original Sheets
            for sheet_name, df in all_sheets.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            categories = ["ECNB", "ECMW"]
            
            for category in categories:
                # Check columns
                if len(df_source.columns) < 9:
                    st.error("Error: File has fewer than 9 columns.")
                    st.stop()

                # Filter Data (Column I / Index 8)
                mask = df_source.iloc[:, 8].astype(str).str.contains(category, case=False, na=False)
                df_filtered = df_source[mask]
                
                # Write Data Sheet (Raw Data)
                df_filtered.to_excel(writer, sheet_name=f"Data {category}", index=False)
                
                # --- PIVOT SIMULATION ---
                pivot_cols = ["Company", "Name", "Account", "DU ID"]
                
                # Check keys
                missing = [c for c in pivot_cols if c not in df_filtered.columns]
                if missing:
                    st.error(f"Missing columns: {missing}")
                    st.stop()
                
                # Get the Raw Sorted Data (for logic) and Display Data (for visuals)
                df_raw_pivot = df_filtered[pivot_cols].drop_duplicates()
                df_sorted, df_display = create_pivot_view(df_raw_pivot, pivot_cols)
                
                pivot_sheet_name = f"Pivot {category}"
                worksheet = workbook.add_worksheet(pivot_sheet_name)
                writer.sheets[pivot_sheet_name] = worksheet
                
                # Write Headers (Row 3, Index 2)
                for col_num, val in enumerate(pivot_cols):
                    worksheet.write(2, col_num, val, header_fmt)
                
                # Write Display Data (Row 4, Index 3)
                # We iterate row by row to write and apply format
                for row_idx, row_data in df_display.iterrows():
                    # We need to check DU ID for duplication logic.
                    # DU ID is the last column (Index 3 in pivot_cols).
                    # We check the RAW sorted dataframe for the actual value to detect duplicates.
                    actual_du_id = df_sorted.iloc[row_idx]["DU ID"]
                    
                    # Check if this DU ID appears more than once in the whole filtered list
                    # (Note: Logic from script: highlight if DU ID is duplicate in the PIVOT list)
                    is_duplicate = len(df_sorted[df_sorted["DU ID"] == actual_du_id]) > 1
                    
                    excel_row = row_idx + 3
                    
                    for col_idx, cell_value in enumerate(row_data):
                        # Determine Format
                        cell_fmt = data_fmt
                        
                        # 1. Check for Duplicate DU ID (Column 3)
                        if col_idx == 3 and is_duplicate:
                            cell_fmt = orange_fmt
                        # 2. Bold the Company Name (Column 0) if visible (easier to view)
                        elif col_idx == 0 and cell_value != "":
                            cell_fmt = data_bold_fmt

                        worksheet.write(excel_row, col_idx, cell_value, cell_fmt)

                # Set Column Widths (Visuals)
                worksheet.set_column(0, 0, 40) # Company (Wide)
                worksheet.set_column(1, 1, 30) # Name (Medium)
                worksheet.set_column(2, 2, 20) # Account
                worksheet.set_column(3, 3, 25) # DU ID
                
                # --- Summary Table (at G3) ---
                summary = df_filtered.groupby("Company")["Name"].nunique().reset_index()
                summary.columns = ["Company", "Count of Name"]
                
                # Write Summary Headers
                worksheet.write("G3", "Company", header_fmt)
                worksheet.write("H3", "Count of Name", header_fmt)
                
                # Write Summary Data
                last_row = 3
                for idx, row in summary.iterrows():
                    last_row = idx + 3
                    worksheet.write(last_row, 6, row["Company"], data_fmt)
                    worksheet.write(last_row, 7, row["Count of Name"], data_fmt)
                
                # Write Grand Total Row
                total_row = last_row + 1
                total_count = summary["Count of Name"].sum()
                worksheet.write(total_row, 6, "Grand Total", header_fmt)
                worksheet.write(total_row, 7, total_count, header_fmt)
                
                # Summary Widths
                worksheet.set_column(6, 6, 40) # G
                worksheet.set_column(7, 7, 15) # H

        output.seek(0)
        st.success("Processing Complete!")
        
        st.download_button(
            label="Download Processed Excel File",
            data=output,
            file_name="Processed_ClockReport.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"An error occurred: {e}")
