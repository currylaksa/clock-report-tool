import streamlit as st
import pandas as pd
from io import BytesIO

# Page Configuration
st.set_page_config(page_title="Eastern Region Clock Report", layout="centered")

st.title("Eastern Region Clock Report Processor")
st.markdown("""
**Instructions:**
1. Upload the daily **Clock Detail Report** (Excel file).
2. The system will process **ECNB** and **ECMW** categories.
3. **Original sheets are preserved.**
4. Download the formatted report.
""")

# 1. File Upload
uploaded_file = st.file_uploader("Upload your clockreport file (xlsx)", type="xlsx")

if uploaded_file:
    try:
        # Load ALL sheets to preserve them (sheet_name=None reads all)
        all_sheets = pd.read_excel(uploaded_file, sheet_name=None)
        
        # Verify the source sheet exists
        source_sheet_name = "Clock Detail Report"
        if source_sheet_name not in all_sheets:
            st.error(f"Error: The file must contain a sheet named '{source_sheet_name}'.")
            st.stop()
            
        df_source = all_sheets[source_sheet_name]
        
        # CLEANUP Headers: Strip whitespace (e.g., "Company " -> "Company")
        df_source.columns = df_source.columns.astype(str).str.strip()
        
        # Prepare Output
        output = BytesIO()
        
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # --- DEFINING FORMATS (To make it look like the PowerShell version) ---
            # Header Format: Bold, Light Blue background, Border
            header_fmt = workbook.add_format({
                'bold': True,
                'text_wrap': False,
                'valign': 'top',
                'fg_color': '#D9E1F2', # Light blue
                'border': 1
            })
            
            # Data Format: Border
            data_fmt = workbook.add_format({'border': 1})
            
            # Orange Highlight Format (for Duplicates)
            orange_fmt = workbook.add_format({'bg_color': '#FFC000', 'font_color': '#000000'})
            
            # 1. Write ALL Original Sheets back to the file
            for sheet_name, df in all_sheets.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # 2. Process Categories
            categories = ["ECNB", "ECMW"]
            
            for category in categories:
                # --- Step 1: Filter Data ---
                if len(df_source.columns) < 9:
                    st.error("Error: File has fewer than 9 columns.")
                    st.stop()

                # Filter (Column I / Index 8)
                mask = df_source.iloc[:, 8].astype(str).str.contains(category, case=False, na=False)
                df_filtered = df_source[mask]
                
                # Write Data Sheet
                data_sheet_name = f"Data {category}"
                df_filtered.to_excel(writer, sheet_name=data_sheet_name, index=False)
                
                # --- Step 2: Create Pivot-Like Sheet ---
                pivot_cols = ["Company", "Name", "Account", "DU ID"]
                
                # Check columns
                missing = [c for c in pivot_cols if c not in df_filtered.columns]
                if missing:
                    st.error(f"Missing columns: {missing}")
                    st.stop()
                    
                # Create unique list (Pivot data)
                df_pivot = df_filtered[pivot_cols].drop_duplicates()
                
                pivot_sheet_name = f"Pivot {category}"
                
                # Write data with header format
                # We write the dataframe but turn off the default header so we can write our own styled one
                df_pivot.to_excel(writer, sheet_name=pivot_sheet_name, index=False, startrow=2, header=False)
                
                worksheet = writer.sheets[pivot_sheet_name]
                
                # --- MANUAL FORMATTING TO MATCH IMAGES ---
                
                # A. Write Styled Headers (Row 3, Index 2)
                for col_num, value in enumerate(df_pivot.columns.values):
                    worksheet.write(2, col_num, value, header_fmt)
                    
                # B. Apply Data Border to all cells
                # (Looping is expensive, so we just apply column width mostly, but let's try to be clean)
                # Setting column widths is the most important part for readability!
                worksheet.set_column('A:A', 25) # Company
                worksheet.set_column('B:B', 20) # Name
                worksheet.set_column('C:C', 15) # Account
                worksheet.set_column('D:D', 15) # DU ID
                
                # C. Highlight Duplicates (Orange)
                if len(df_pivot) > 0:
                    start_row = 3 # Excel Row 4
                    end_row = start_row + len(df_pivot) - 1
                    worksheet.conditional_format(start_row, 3, end_row, 3,
                                                 {'type': 'duplicate', 'format': orange_fmt})

                # --- Feature 3: Summary Table (at G3) ---
                summary = df_filtered.groupby("Company")["Name"].nunique().reset_index()
                summary.columns = ["Company", "Count of Name"]
                
                # Write Summary Headers (G3, H3)
                worksheet.write("G3", "Company", header_fmt)
                worksheet.write("H3", "Count of Name", header_fmt)
                
                # Write Summary Data
                for idx, row in summary.iterrows():
                    worksheet.write(idx + 3, 6, row["Company"], data_fmt)
                    worksheet.write(idx + 3, 7, row["Count of Name"], data_fmt)
                
                # Set Summary Column Widths
                worksheet.set_column('G:G', 25)
                worksheet.set_column('H:H', 15)

        output.seek(0)
        
        st.success("Processing Complete!")
        
        st.download_button(
            label="Download Processed Excel File",
            data=output,
            file_name=f"Processed_ClockReport.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"An error occurred: {e}")
