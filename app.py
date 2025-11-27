import streamlit as st
import pandas as pd
from io import BytesIO

# Title of the Web App
st.title("Eastern Region Clock Report Processor")
st.write("Upload the daily Clock Report to process it automatically.")

# 1. User Uploads File
uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file:
    st.success("File uploaded successfully! Processing...")
    
    # Load the Excel file (No need for installed Excel)
    # We use openpyxl engine for xlsx files
    try:
        # Load the source sheet
        df_source = pd.read_excel(uploaded_file, sheet_name="Clock Detail Report")
        
        # Create an in-memory buffer to save the new file
        output = BytesIO()
        
        # We use ExcelWriter to write multiple sheets
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            
            # Write the original sheet back (optional, but good for reference)
            df_source.to_excel(writer, sheet_name="Clock Detail Report", index=False)
            
            categories = ["ECNB", "ECMW"]
            
            for category in categories:
                # --- Step 1: Filter Data ---
                # Filter Column I (Index 8 in 0-based Python) or by Name if known. 
                # Assuming Column I name matches header. Let's filter by string content in all columns if unsure, 
                # but better to assume the 9th column (index 8) if headers are standard.
                # Adjust 'Column_I_Name' to the actual header name of Column I.
                # For now, we search all string columns for the category to be safe, or use the 9th column.
                
                target_col = df_source.columns[8] # 9th Column
                
                # Filter rows where the category exists in the target column
                mask = df_source[target_col].astype(str).str.contains(category, na=False)
                df_cat_data = df_source[mask]
                
                # Write 'Data [Category]' sheet
                data_sheet_name = f"Data {category}"
                df_cat_data.to_excel(writer, sheet_name=data_sheet_name, index=False)
                
                # --- Step 2: Create "Pivot" Table Data ---
                # Your PS script creates a Pivot with: Company, Name, Account, DU ID (Rows) and No Values.
                # This is effectively a list of unique combinations.
                pivot_cols = ["Company", "Name", "Account", "DU ID"]
                
                # specific_pivot = df_cat_data[pivot_cols].drop_duplicates()
                # Use strict column names from your file. If headers differ, update 'pivot_cols'.
                # We will attempt to find these columns case-insensitively if needed.
                try:
                    df_pivot = df_cat_data[pivot_cols].drop_duplicates()
                except KeyError:
                    st.error(f"Error: Could not find columns {pivot_cols}. Please check file headers.")
                    st.stop()
                
                pivot_sheet_name = f"Pivot {category}"
                df_pivot.to_excel(writer, sheet_name=pivot_sheet_name, index=False, startrow=2) # Start at Row 3 (Index 2)
                
                # Access the workbook and worksheet to apply formatting
                workbook = writer.book
                worksheet = writer.sheets[pivot_sheet_name]
                
                # --- Feature 1: Highlight Duplicate DU IDs ---
                # Find duplicates in the result
                dupe_ids = df_pivot[df_pivot.duplicated(subset=['DU ID'], keep=False)]['DU ID'].unique()
                
                # Define Orange format
                orange_format = workbook.add_format({'bg_color': '#FFC000'}) # Orange-ish
                
                # Apply format. Iterate through rows in the Excel sheet
                # Data starts at row 3 (Excel row 4), so we check from there.
                # df_pivot index is not reliable for row number in excel, we enumerate.
                du_id_col_idx = pivot_cols.index("DU ID")
                
                for row_num, row_data in enumerate(df_pivot.values):
                    du_id = row_data[du_id_col_idx]
                    if du_id in du_id_col_idx:
                        # This logic needs to check if the ID is in the duplicate list
                        pass
                
                # Simpler approach: Conditional Formatting on the column range
                # Calculate range e.g., D4:D100
                row_count = len(df_pivot)
                if row_count > 0:
                    start_row = 3 # Excel Row 4 (0-based index 3)
                    end_row = start_row + row_count - 1
                    # DU ID is the 4th column (D)
                    worksheet.conditional_format(start_row, 3, end_row, 3, 
                                                 {'type': 'duplicate', 'format': orange_format})

                # --- Feature 2: Summary Table (Unique Names per Company) ---
                # Calculate unique name count per company
                summary = df_cat_data.groupby('Company')['Name'].nunique().reset_index()
                summary.columns = ['Company', 'Count of Name']
                
                # Write Summary Table at G3 (Row 2, Col 6)
                # Write headers
                worksheet.write('G3', 'Company')
                worksheet.write('H3', 'Count of Name')
                
                # Write data
                for i, row in summary.iterrows():
                    worksheet.write(i + 3, 6, row['Company'])
                    worksheet.write(i + 3, 7, row['Count of Name'])
                    
        # 3. Download Button
        # Reset buffer position
        output.seek(0)
        
        st.download_button(
            label="Download Processed Report",
            data=output,
            file_name=f"processed_clock_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"An error occurred: {e}")