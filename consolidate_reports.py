import os
import pandas as pd
from tkinter import Tk, filedialog

def create_consolidated_report():
    root = Tk()
    root.withdraw()
    folder_path = filedialog.askdirectory(title="Select Folder Containing Excel and CSV Files")
    
    if not folder_path:
        print("No folder selected. Operation cancelled.")
        return

    # Normalize folder path cleanly
    folder_path = os.path.abspath(folder_path) + os.sep
    print(f"\nScanning directory: {folder_path}")

    dict_count = {}
    dict_bytes = {}
    file_total = 0
    file_done = 0
    processed_sheets = 0

    try:
        all_items = os.listdir(folder_path)
    except Exception as e:
        print(f"Error accessing directory: {e}")
        return

    print(f"Total items found in directory: {len(all_items)}")

    valid_files = []
    for item in all_items:
        item_lower = item.lower()
        # Check if it's an excel or csv file and not the output report
        if item_lower.endswith(('.xlsx', '.xls', '.csv')) and "consolidated report" not in item_lower:
            full_path = os.path.join(folder_path, item)
            valid_files.append(full_path)
            file_total += 1
            print(f"  [ACCEPTED] Found valid file: {item}")
        else:
            print(f"  [SKIPPED] Ignored item: {item}")

    if file_total == 0:
        print("\n[WARNING] 0 valid files were identified. Make sure your CSV and Excel files are directly in the selected folder.")
        return

    print(f"\n--- Starting Processing of {file_total} files ---")

    for file_path in valid_files:
        file_name = os.path.basename(file_path)
        file_done += 1

        base_name = os.path.splitext(file_name)[0].lower()
        is_custom_file = "custom" in base_name

        file_type_label = "[CUSTOM FILE]" if is_custom_file else "[Standard File]"
        print(f"\nProcessing {file_type_label} ({file_done}/{file_total}): {file_name}")

        try:
            if file_name.lower().endswith('.csv'):
                # Read CSV file safely
                df_csv = pd.read_csv(file_path, header=None, sep=None, engine='python', encoding='utf-8-sig', on_bad_lines='skip')
                dfs = {file_name: df_csv}
            else:
                dfs = pd.read_excel(file_path, sheet_name=None, header=None)

            for sheet_name, df in dfs.items():
                print(f"  -> Sheet/Source '{sheet_name}' shape: {df.shape}")
                
                if is_custom_file:
                    # Custom files: Column A (0) = Name, Column B (1) = Count, Column C (2) = Size
                    if df.shape[1] >= 3:
                        sheet_rows_added = 0
                        for _, row in df.iterrows():
                            error_name = str(row.iloc[0]).strip()
                            if not error_name or error_name.lower() in ['nan', 'none', '', 'error name', 'name', 'warning name']:
                                continue
                                
                            try:
                                row_count = float(row.iloc[1]) if pd.notnull(row.iloc[1]) else 0.0
                            except ValueError:
                                continue

                            try:
                                row_bytes = float(row.iloc[2]) if pd.notnull(row.iloc[2]) else 0.0
                            except ValueError:
                                row_bytes = 0.0

                            dict_count[error_name] = dict_count.get(error_name, 0.0) + row_count
                            dict_bytes[error_name] = dict_bytes.get(error_name, 0.0) + row_bytes
                            sheet_rows_added += 1

                        print(f"     Successfully parsed {sheet_rows_added} custom rows.")
                        processed_sheets += 1
                    else:
                        print(f"     ERROR: File has fewer than 3 columns (found {df.shape[1]}).")
                else:
                    # Standard files logic
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
                    else:
                        print(f"     Standard file missing 'Count' or 'Size' headers. Skipped.")

        except Exception as e:
            print(f"  [CRITICAL ERROR] Processing file {file_name}: {e}")

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

    print("\n==========================================")
    print("Consolidated report has been created successfully.")
    print(f"Files processed: {file_done}")
    print(f"Sheets processed: {processed_sheets}")
    print(f"Location: {output_path}")
    print("==========================================")

if __name__ == "__main__":
    create_consolidated_report()
