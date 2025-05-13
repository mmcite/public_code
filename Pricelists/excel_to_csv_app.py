import streamlit as st
import pandas as pd
import os
from pathlib import Path
import base64
import io

def clean_number(value):
    try:
        return int(float(str(value).replace(',', '.')))
    except (ValueError, TypeError):
        return 0

def process_file(input_file, sheet_name, columns, output_file, filter_column=None, min_length=3):
    try:
        # Read the Excel file
        df = pd.read_excel(input_file, sheet_name=sheet_name)
        
        # Select specified columns
        if columns:
            available_columns = df.columns.tolist()
            valid_columns = [col for col in columns if col in available_columns]
            
            if not valid_columns:
                return f"Error: None of the specified columns {columns} found in the file. Available columns: {available_columns}"
            
            df = df[valid_columns]
        
        # Filter rows based on minimum length of a column value if specified
        if filter_column and filter_column in df.columns:
            df = df[df[filter_column].astype(str).str.len() >= min_length]
        
        # Clean numeric columns (all except first column)
        for col in df.columns[1:]:
            df[col] = df[col].apply(clean_number)
        
        # Save to CSV
        csv_data = df.to_csv(sep=';', index=False, header=False)
        return csv_data, f"Successfully processed file"
    except Exception as e:
        return None, f"Error processing file: {str(e)}"

def get_download_link(csv_data, filename, link_text="Download CSV file"):
    b64 = base64.b64encode(csv_data.encode()).decode()
    href = f'<a href="data:file/csv;base64,{b64}" download="{filename}">{link_text}</a>'
    return href

def main():
    st.title("Excel to CSV Converter - Pricelist")
    st.write("Convert Excel files to CSV with custom sheet and column selection")
    
    # Setup output directory
    output_dir = Path('output')
    output_dir.mkdir(exist_ok=True)
    
    # File upload
    st.header("1. Upload Excel File")
    uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        try:
            # Load and display sheet names
            xls = pd.ExcelFile(uploaded_file)
            sheet_names = xls.sheet_names
            
            st.header("2. Select Sheet")
            selected_sheet = st.selectbox("Choose a sheet", sheet_names, index=0 if 'ceník' in sheet_names else 0)
            
            # Load and display columns
            if selected_sheet:
                # Reset file pointer to beginning
                uploaded_file.seek(0)
                df_preview = pd.read_excel(uploaded_file, sheet_name=selected_sheet, nrows=5)
                
                st.header("3. Select Columns")
                st.write("Preview of first 5 rows:")
                st.dataframe(df_preview)
                
                all_columns = df_preview.columns.tolist()
                
                # Preselect common columns based on your existing code
                default_columns = []
                common_first_columns = ['Artikl/Article', 'SKU']
                common_price_columns = ['nakup cena CZK', 'cena CZK (nákup)', 'nákup cena CZK', 
                                       'nákup (CZK)', 'nakup CZK', 'cost (CZK)']
                common_other_columns = ['SPODNÍ STAVBY', 'MONTÁŽ', 'SPODNÍ STAVBY (nákup)', 
                                       'MONTÁŽ (nákup)', 'unit (USD)', 'unit (CAD)', 'unit (MXN)', 'PRICE (EUR)']
                
                # Try to find and preselect common columns
                for col in all_columns:
                    if col in common_first_columns:
                        default_columns.append(col)
                        break
                
                for col in all_columns:
                    if col in common_price_columns:
                        default_columns.append(col)
                        break
                
                for col in common_other_columns:
                    if col in all_columns:
                        default_columns.append(col)
                
                selected_columns = st.multiselect(
                    "Select columns to include in CSV", 
                    all_columns,
                    default=default_columns if default_columns else all_columns[:min(3, len(all_columns))]
                )
                
                # Filter options
                st.header("4. Filter Options")
                filter_enabled = st.checkbox("Filter rows by minimum text length", value=True)
                
                filter_column = None
                min_length = 3
                
                if filter_enabled:
                    filter_column = st.selectbox(
                        "Select column to filter by length", 
                        selected_columns,
                        index=0 if selected_columns and selected_columns[0] in common_first_columns else 0
                    )
                    min_length = st.number_input("Minimum length", min_value=1, value=3)
                
                # Output options
                st.header("5. Output Options")
                output_filename = st.text_input("Output filename", value=uploaded_file.name.replace('.xlsx', '.csv'))
                
                # Process button
                if st.button("Convert to CSV"):
                    if not selected_columns:
                        st.error("Please select at least one column")
                    else:
                        # Reset file pointer to beginning
                        uploaded_file.seek(0)
                        
                        csv_data, result_message = process_file(
                            uploaded_file, 
                            selected_sheet, 
                            selected_columns, 
                            None,  # No output file path needed
                            filter_column if filter_enabled else None,
                            min_length
                        )
                        
                        if csv_data:
                            st.success(result_message)
                            st.markdown(get_download_link(csv_data, output_filename), unsafe_allow_html=True)
                            
                            # Also save to output directory if needed
                            output_file = output_dir / output_filename
                            with open(output_file, 'w', encoding='utf-8') as f:
                                f.write(csv_data)
                            st.info(f"File also saved to {output_file}")
                        else:
                            st.error(result_message)
                
        except Exception as e:
            st.error(f"Error loading file: {str(e)}")

if __name__ == '__main__':
    main()