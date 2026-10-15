import os
import glob
import pandas as pd
from tkinter import Tk, filedialog

def create_consolidated_report():
    root = Tk()
    root.withdraw()
    folder_path = filedialog.askdirectory(title="Select Folder Containing Excel and CSV Files")
    
    if not folder_path:
        print("No folder selected. Operation cancelled.")
        return

    if not folder_path.endswith('/'):
        folder_path += '/'

    dict_count = {}
    dict_bytes = {}
    file_total = 0
    file_done = 0
    processed_sheets = 0

    search_patterns = [os.path.join(folder_path, "*.xlsx"), 
                       os.path.join(folder_path, "*.xls"), 
                       os.path.join(folder_path, "*.csv"),
                       os.path.join(folder_path, "*.XLSX"),
                       os.path.join(folder_path, "*.XLS"),
                       os.path.join(folder_path, "*.CSV")]
    
    all_files = []
    for pattern in search_patterns:
        all_files.extend(glob.glob(pattern))

    unique_files = list(set(all_files))
    valid_files = []
    
    for file_path in unique_files:
        file_name = os.path.basename(file_path)
        if "consolidated report" not in file_name.lower():
            valid_files.append(file_path)
            file_total += 1

    if file_total == 0:
        print("No valid Excel or CSV files were found in the selected folder.")
        return

    print(f"\n--- Scanning {file_total} files ---")

    for file_path in valid_files:
        file_name = os.path.basename(file_path)
        file_done += 1

        base_name = os.path.splitext(file_name)[0].lower()
        is_custom_file = "custom" in base_name

        file_type_label = "[CUSTOM FILE]" if is_custom_file else "[Standard File]"
        print(f"\nProcessing {file_type_label} ({file_done}/{file_total}): {file_name}")

        try:
            if file_name.lower().endswith('.csv'):
                df_csv = pd.read_csv(file_path, header=None, sep=None, engine='python', encoding='utf-8-sig', on_bad_lines='skip')
                dfs = {file_name: df_csv}
            else:
                dfs = pd.read_excel(file_path, sheet_name=None, header=None)

            for sheet_name, df in dfs.items():
                if is_custom_file:
                    if df.shape[1] >= 3:
                        data_rows = df.iloc[3:]
                        sheet_rows_added = 0
                        for _, row in data_rows.iterrows():
                            error_name = str(row.iloc[0]).strip()
                            if error_name and error_name.lower() != 'nan' and error_name != '':
                                try:
                                    row_count = float(row.iloc[1]) if pd.notnull(row.iloc[1]) else 0.0
                                except ValueError:
                                    row_count = 0.0

                                try:
                                    row_bytes = float(row.iloc[2]) if pd.notnull(row.iloc[2]) else 0.0
                                except ValueError:
                                    row_bytes = 0.0

                                dict_count[error_name] = dict_count.get(error_name, 0.0) + row_count
                                dict_bytes[error_name] = dict_bytes.get(error_name, 0.0) + row_bytes
                                sheet_rows_added += 1
                        processed_sheets += 1
                    else:
                        print(f"     ERROR: Custom file has fewer than 3 columns.")
                else:
                    header_row_idx = -1
                    count_col = -1
                    size_col = -1

                    for idx, row in df.iterrows():
                        row_strs = [str(val).strip().lower() for val in row.values]
                        if "count" in row_strs and "size" in row_strs:
                            header_row_idx = idx
                            count_col = row_strs.index("count")
                            size_col = row_strs.index("size")
                            break

                    if header_row_idx != -1 and count_col != -1 and size_col != -1:
                        data_rows = df.iloc[header_row_idx + 1:]
                        row_count_val = 0.0
                        row_bytes_val = 0.0

                        for _, row in data_rows.iterrows():
                            try:
                                c_val = float(row.iloc[count_col]) if pd.notnull(row.iloc[count_col]) else 0.0
                                row_count_val += c_val
                            except ValueError:
                                pass

                            try:
                                s_val = float(row.iloc[size_col]) if pd.notnull(row.iloc[size_col]) else 0.0
                                row_bytes_val += s_val
                            except ValueError:
                                pass

                        if "ocr" in str(sheet_name).upper() or "text" in str(sheet_name).upper():
                            error_name = "OCR Exceptions"
                        else:
                            error_name = str(sheet_name)

                        if "protected" not in error_name.lower() and "encrypted" not in error_name.lower():
                            dict_count[error_name] = dict_count.get(error_name, 0.0) + row_count_val
                            dict_bytes[error_name] = dict_bytes.get(error_name, 0.0) + row_bytes_val
                            processed_sheets += 1

        except Exception as e:
            print(f"Error processing file {file_name}: {e}")

    report_data = []
    for key in dict_count:
        size_mb = dict_bytes[key] / (1024 ** 2)
        report_data.append({
            "Error Name": key,
            "Count": dict_count[key],
            "Size (MB)": size_mb
        })

    report_df = pd.DataFrame(report_data)

    if not report_df.empty:
        total_count = report_df["Count"].sum()
        total_size = report_df["Size (MB)"].sum()
        
        total_row = pd.DataFrame([{
            "Error Name": "Grand Total",
            "Count": total_count,
            "Size (MB)": total_size
        }])
        report_df = pd.concat([report_df, total_row], ignore_index=True)

    output_path = os.path.join(folder_path, "Consolidated report.xlsx")
    report_df.to_excel(output_path, index=False)

    print("\nConsolidated report has been created successfully.")
    print(f"Files processed: {file_done}")
    print(f"Sheets processed: {processed_sheets}")
    print(f"Location: {output_path}")

if __name__ == "__main__":
    create_consolidated_report()