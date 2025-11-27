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
3. Download the finished report.
""")

# 1. File Upload
uploaded_file = st.file_uploader("Upload your clockreport file (xlsx)", type="xlsx")

if uploaded_file:
    try:
        # Load the file
        df_source = pd.read_excel(uploaded_file, sheet_name="Clock Detail Report")
        
        # CLEANUP: Strip whitespace from headers (e.g., "Company " -> "Company")
        df_source.columns = df_source.columns.astype(str).str.strip()
        
        # Prepare the Output Buffer
        output = BytesIO()
        
        # Use XlsxWriter engine
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # Define Orange Format for duplicates
            format_orange = workbook.add_format({'bg_color': '#FFC000', 'font_color': '#000000'})
            
            categories = ["ECNB", "ECMW"]
            
            for category in categories:
                # --- Step 1: Filter Data ---
                # Safety Check: Ensure file has enough columns (Index 8 is the 9th column)
                if len(df_source.columns) < 9:
                    st.error("Error: The uploaded file has fewer than 9 columns. Please check the file format.")
                    st.stop()
                
                # Filter Column I (Index 8) for the category
                mask = df_source.iloc[:, 8].astype(str).str.contains(category, case=False, na=False)
                df_filtered = df_source[mask]
                
                # Write Data Sheet
                data_sheet_name = f"Data {category}"
                df_filtered.to_excel(writer, sheet_name=data_sheet_name, index=False)
                
                # --- Step 2: Create Pivot Table Data ---
                pivot_cols = ["Company", "Name", "Account", "DU ID"]
                
                # Check if columns exist
                missing_cols = [col for col in pivot_cols if col not in df_filtered.columns]
                if missing_cols:
                    st.error(f"Error: Missing columns in file: {missing_cols}")
                    st.stop()
                    
                # Create pivot data (unique rows)
                df_pivot = df_filtered[pivot_cols].drop_duplicates()
                
                pivot_sheet_name = f"Pivot {category}"
                # Write to Excel starting at Row 3 (Index 2)
                df_pivot.to_excel(writer, sheet_name=pivot_sheet_name, index=False, startrow=2)
                
                worksheet = writer.sheets[pivot_sheet_name]
                
                # --- Feature 1: Highlight Duplicates ---
                # Apply conditional formatting to DU ID column (Column D, Index 3)
                if len(df_pivot) > 0:
                    start_row = 3  # Excel Row 4 (0-based index is 3)
                    end_row = start_row + len(df_pivot) - 1
                    
                    # Apply formatting to range D4:D[End]
                    worksheet.conditional_format(start_row, 3, end_row, 3,
                                                 {'type': 'duplicate',
                                                  'format': format_orange})
                
                # --- Feature 2: Summary Table (Unique Counts) ---
                summary = df_filtered.groupby("Company")["Name"].nunique().reset_index()
                summary.columns = ["Company", "Count of Name"]
                
                # Write headers at G3
                worksheet.write("G3", "Company")
                worksheet.write("H3", "Count of Name")
                
                # Write data starting at G4
                for idx, row in summary.iterrows():
                    worksheet.write(idx + 3, 6, row["Company"])       # G is col 6
                    worksheet.write(idx + 3, 7, row["Count of Name"]) # H is col 7

        # Save and Seek
        output.seek(0)
        
        st.success("Processing Complete!")
        
        # 2. Download Button
        st.download_button(
            label="Download Processed Excel File",
            data=output,
            file_name=f"Processed_ClockReport.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        # Show detailed error for debugging
        st.error(f"An error occurred: {e}")
