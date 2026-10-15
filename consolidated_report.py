import glob
import os
import sys
import ctypes
import win32com.client as win32


def find_header_column(ws, header_name):
  last_col = ws.Cells(1, ws.Columns.Count).End(-4159).Column  # -4159 is xlToLeft

  for c in range(1, last_col + 1):
    cell_value = str(ws.Cells(1, c).Value or "").strip().lower()
    if cell_value == header_name.strip().lower():
      return c

  return 0


def show_msgbox(text, title, style=0):
  # Native Windows MessageBox using ctypes (0 = MB_OK, 16 = MB_ICONERROR, 48 = MB_ICONWARNING, 64 = MB_ICONINFORMATION)
  return ctypes.windll.user32.MessageBoxW(0, text, title, style)


def create_consolidated_report(selected_folder=""):
  xl = None
  wb = None
  report_wb = None

  try:
    folder_path = selected_folder.strip()

    # Use gencache to ensure full early-bound COM support for PyInstaller
    xl = win32.gencache.EnsureDispatch("Excel.Application")

    if not folder_path:
      file_dialog = xl.FileDialog(3)  # 3 is msoFileDialogFolderPicker
      file_dialog.Title = "Select Folder Containing Excel Files"
      if file_dialog.Show() != -1:
        return
      folder_path = file_dialog.SelectedItems(1)

    if not folder_path.endswith("\\") and not folder_path.endswith("/"):
      folder_path = folder_path + "\\"

    dict_count = {}
    dict_bytes = {}

    # Find Excel files using glob
    search_pattern = os.path.join(folder_path, "*.xls*")
    all_files = glob.glob(search_pattern)
    
    excel_files = []
    for f in all_files:
      fname = os.path.basename(f)
      if (
          fname.lower() != "consolidated report.xlsx"
          and fname.lower() != "consolidated report.xlsm"
      ):
        excel_files.append(f)

    file_total = len(excel_files)

    if file_total == 0:
      show_msgbox(
          "No Excel files were found in:\r\n" + folder_path,
          "No Files Found",
          48,
      )
      return

    xl.ScreenUpdating = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False

    file_done = 0
    processed_sheets = 0

    for file_path in excel_files:
      file_done += 1
      fname = os.path.basename(file_path)

      wb = xl.Workbooks.Open(
          Filename=file_path, ReadOnly=True, UpdateLinks=False
      )

      for ws in wb.Worksheets:
        count_col = find_header_column(ws, "Count")
        size_col = find_header_column(ws, "Size")

        if count_col > 0 and size_col > 0:
          last_row_count = ws.Cells(ws.Rows.Count, count_col).End(-4162).Row  # -4162 is xlUp
          last_row_size = ws.Cells(ws.Rows.Count, size_col).End(-4162).Row
          last_row = max(last_row_count, last_row_size)

          sheet_count = 0.0
          sheet_bytes = 0.0

          for r in range(2, last_row + 1):
            val_count = ws.Cells(r, count_col).Value
            val_size = ws.Cells(r, size_col).Value

            if val_count is not None and str(val_count).replace(".", "", 1).isdigit():
              sheet_count += float(val_count)
            if val_size is not None and str(val_size).replace(".", "", 1).isdigit():
              sheet_bytes += float(val_size)

          ws_name = str(ws.Name)
          if "OCR" in ws_name.upper() or "TEXT" in ws_name.upper():
            error_name = "OCR Exceptions"
          else:
            error_name = ws_name

          if error_name not in dict_count:
            dict_count[error_name] = sheet_count
            dict_bytes[error_name] = sheet_bytes
          else:
            dict_count[error_name] += sheet_count
            dict_bytes[error_name] += sheet_bytes

          processed_sheets += 1

      wb.Close(SaveChanges=False)
      wb = None

    report_wb = xl.Workbooks.Add()
    report_ws = report_wb.Worksheets(1)
    report_ws.Name = "Consolidated Report"

    report_ws.Cells(1, 1).Value = "Error Name"
    report_ws.Cells(1, 2).Value = "Count"
    report_ws.Cells(1, 3).Value = "Size (MB)"

    header_range = report_ws.Range("A1:C1")
    header_range.Font.Bold = True
    header_range.HorizontalAlignment = -4108  # xlCenter
    header_range.Interior.ColorIndex = 15

    output_row = 2
    for key in dict_count:
      report_ws.Cells(output_row, 1).Value = key
      report_ws.Cells(output_row, 2).Value = dict_count[key]
      report_ws.Cells(output_row, 3).Value = dict_bytes[key] / (1024.0**2)
      output_row += 1

    report_ws.Columns("A:C").AutoFit()
    report_ws.Columns("B").NumberFormat = "#,##0"
    report_ws.Columns("C").NumberFormat = "#,##0.00"

    if output_row > 2:
      report_ws.Range(f"A1:C{output_row - 1}").Borders.LineStyle = 1  # xlContinuous

    report_ws.Cells(output_row, 1).Value = "Grand Total"
    report_ws.Cells(output_row, 2).Formula = f"=SUM(B2:B{output_row - 1})"
    report_ws.Cells(output_row, 3).Formula = f"=SUM(C2:C{output_row - 1})"

    total_range = report_ws.Range(f"A{output_row}:C{output_row}")
    total_range.Font.Bold = True
    total_range.Borders.LineStyle = 1

    report_wb.SaveAs(
        Filename=folder_path + "Consolidated report.xlsx",
        FileFormat=51,  # xlOpenXMLWorkbook
    )
    report_wb.Close(SaveChanges=True)
    report_wb = None

    xl.ScreenUpdating = True
    xl.DisplayAlerts = True
    xl.EnableEvents = True

    show_msgbox(
        "Consolidated report has been created successfully.\r\n\r\nFiles processed: "
        + str(file_done)
        + "\r\nSheets processed: "
        + str(processed_sheets)
        + "\r\n\r\nLocation:\r\n"
        + folder_path
        + "Consolidated report.xlsx",
        "Report Complete",
        64,
    )

  except Exception as err:
    if wb is not None:
      wb.Close(SaveChanges=False)
    if report_wb is not None:
      report_wb.Close(SaveChanges=False)
    if xl is not None:
      xl.ScreenUpdating = True
      xl.DisplayAlerts = True
      xl.EnableEvents = True
    show_msgbox(
        "An error occurred:\r\n\r\n" + str(err),
        "Consolidated Report Tool",
        16,
    )


if __name__ == "__main__":
  create_consolidated_report()
