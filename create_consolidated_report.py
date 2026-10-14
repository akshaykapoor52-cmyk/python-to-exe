import os
import glob
import tkinter as tk
from tkinter import filedialog, messagebox
import openpyxl

def find_header_column(ws, header_name):
    # ws is an openpyxl worksheet
    max_col = ws.max_column
    target = header_name.strip().lower()
    for c in range(1, max_col + 1):
        cell_val = ws.cell(row=1, column=c).value
        if cell_val is not None:
            if str(cell_val).strip().lower() == target:
                return c
    return 0

def create_consolidated_report(selected_folder=""):
    folder_path = selected_folder.strip()
    
    if not folder_path:
        root = tk.Tk()
        root.withdraw()
        folder_path = filedialog.askdirectory(title="Select Folder Containing Excel Files")
        if not folder_path:
            return
            
    if not folder_path.endswith("/") and not folder_path.endswith("\"):
        folder_path += os.path.sep
        
    print("Preparing...")
    
    # Find all excel files
    pattern = os.path.join(folder_path, "*.xls*")
    all_files = glob.glob(pattern)
    
    excel_files = []
    for f in all_files:
        base_name = os.path.basename(f)
        if base_name.lower() not in ["consolidated report.xlsx", "consolidated report.xlsm"]:
            excel_files.append(f)
            
    file_total = len(excel_files)
    if file_total == 0:
        messagebox.showwarning("No Files Found", f"No Excel files were found in:\n{folder_path}")
        return
        
    dict_count = {}
    dict_bytes = {}
    file_done = 0
    processed_sheets = 0
    
    for file_path in excel_files:
        file_done += 1
        file_name = os.path.basename(file_path)
        print(f"Processing file {file_done} of {file_total}: {file_name}")
        
        try:
            # Open workbook read-only using openpyxl or pandas/openpyxl
            wb = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                
                count_col = find_header_column(ws, "Count")
                size_col = find_header_column(ws, "Size")
                
                if count_col > 0 and size_col > 0:
                    # To find max row efficiently in openpyxl read_only mode
                    # Note: read_only ws.max_row might sometimes be None or inaccurate if not calculated,
                    # let's iterate rows or compute max row.
                    sheet_count = 0.0
                    sheet_bytes = 0.0
                    
                    row_idx = 0
                    for row in ws.iter_rows(min_row=2, values_only=True):
                        row_idx += 2 # placeholder or track row
                        # let's check values at count_col - 1 and size_col - 1 (0-indexed)
                        if len(row) >= max(count_col, size_col):
                            c_val = row[count_col - 1]
                            s_val = row[size_col - 1]
                            
                            if isinstance(c_val, (int, float)):
                                sheet_count += float(c_val)
                            elif c_val is not None:
                                try:
                                    sheet_count += float(str(c_val).replace(",", ""))
                                except ValueError:
                                    pass
                                    
                            if isinstance(s_val, (int, float)):
                                sheet_bytes += float(s_val)
                            elif s_val is not None:
                                try:
                                    sheet_bytes += float(str(s_val).replace(",", ""))
                                except ValueError:
                                    pass
                                    
                    if "ocr" in sheet_name.lower() or "text" in sheet_name.lower():
                        error_name = "OCR Exceptions"
                    else:
                        error_name = sheet_name
                        
                    dict_count[error_name] = dict_count.get(error_name, 0.0) + sheet_count
                    dict_bytes[error_name] = dict_bytes.get(error_name, 0.0) + sheet_bytes
                    processed_sheets += 1
            wb.close()
        except Exception as e:
            print(f"Error processing {file_name}: {e}")
            
    print("Building consolidated report...")
    
    report_wb = openpyxl.Workbook()
    report_ws = report_wb.active
    report_ws.title = "Consolidated Report"
    
    report_ws.cell(row=1, column=1, value="Error Name")
    report_ws.cell(row=1, column=2, value="Count")
    report_ws.cell(row=1, column=3, value="Size (MB)")
    
    # Styling header
    header_font = openpyxl.styles.Font(bold=True)
    header_alignment = openpyxl.styles.Alignment(horizontal="center")
    header_fill = openpyxl.styles.PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")
    
    for col in range(1, 4):
        cell = report_ws.cell(row=1, column=col)
        cell.font = header_font
        cell.alignment = header_alignment
        cell.fill = header_fill
        
    output_row = 2
    thin_border = openpyxl.styles.Border(
        left=openpyxl.styles.Side(style='thin'),
        right=openpyxl.styles.Side(style='thin'),
        top=openpyxl.styles.Side(style='thin'),
        bottom=openpyxl.styles.Side(style='thin')
    )
    
    for key in dict_count:
        report_ws.cell(row=output_row, column=1, value=key)
        report_ws.cell(row=output_row, column=2, value=dict_count[key])
        report_ws.cell(row=output_row, column=3, value=dict_bytes[key] / (1024.0 ** 2))
        
        for col in range(1, 4):
            report_ws.cell(row=output_row, column=col).border = thin_border
            
        report_ws.cell(row=output_row, column=2).number_format = "#,##0"
        report_ws.cell(row=output_row, column=3).number_format = "#,##0.00"
        output_row += 1
        
    # Grand Total row
    report_ws.cell(row=output_row, column=1, value="Grand Total")
    report_ws.cell(row=output_row, column=2, value=f"=SUM(B2:B{output_row - 1})")
    report_ws.cell(row=output_row, column=3, value=f"=SUM(C2:C{output_row - 1})")
    
    for col in range(1, 4):
        cell = report_ws.cell(row=output_row, column=col)
        cell.font = header_font
        cell.border = thin_border
        
    report_ws.cell(row=output_row, column=2).number_format = "#,##0"
    report_ws.cell(row=output_row, column=3).number_format = "#,##0.00"
    
    # Auto-fit columns
    for col in report_ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        report_ws.column_dimensions[col_letter].width = max(max_len + 3, 12)
        
    output_path = os.path.join(folder_path, "Consolidated report.xlsx")
    report_wb.save(output_path)
    print("Completed successfully.")
    
    messagebox.showinfo(
        "Report Complete",
        f"Consolidated report has been created successfully.\n\nFiles processed: {file_done}\nSheets processed: {processed_sheets}\n\nLocation:\n{output_path}"
    )

if __name__ == "__main__":
    create_consolidated_report()
