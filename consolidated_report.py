import os
import sys
import win32com.client


def find_header_column(ws, header_name):
  last_col = ws.Cells(1, ws.Columns.Count).End(-4159).Column  # -4159 is xlToLeft

  for c in range(1, last_col + 1):
    cell_value = str(ws.Cells(1, c).Value or "").strip().lower()
    if cell_value == header_name.strip().lower():
      return c

  return 0


def create_consolidated_report(selected_folder=""):
  xl = None
  wb = None
  report_wb = None

  try:
    folder_path = selected_folder.strip()

    # Initialize Excel application
    xl = win32com.client.Dispatch("Excel.Application")

    if not folder_path:
      file_dialog = xl.FileDialog(3)  # 3 is msoFileDialogFolderPicker
      file_dialog.Title = "Select Folder Containing Excel Files"
      if file_dialog.Show() != -1:
        return
      folder_path = file_dialog.SelectedItems(1)

    if not folder_path.endswith("\\"):
      folder_path = folder_path + "\\"

    dict_count = {}
    dict_bytes = {}

    # Find total files
    file_total = 0
    file_name = xl.Dir(folder_path & "*.xls*")
    while file_name != "":
      if (
          file_name.lower() != "consolidated report.xlsx"
          and file_name.lower() != "consolidated report.xlsm"
      ):
        file_total += 1
      file_name = xl.Dir()

    if file_total == 0:
      xl.MsgBox(
          "No Excel files were found in:\r\n" + folder_path,
          48,
          "No Files Found",
      )
      return

    xl.ScreenUpdating = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False

    file_done = 0
    processed_sheets = 0
    file_name = xl.Dir(folder_path & "*.xls*")

    while file_name != "":
      if (
          file_name.lower() != "consolidated report.xlsx"
          and file_name.lower() != "consolidated report.xlsm"
      ):
        file_done += 1
        file_path = folder_path + file_name

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

      file_name = xl.Dir()

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

    xl.MsgBox(
        "Consolidated report has been created successfully.\r\n\r\nFiles processed: "
        + str(file_done)
        + "\r\nSheets processed: "
        + str(processed_sheets)
        + "\r\n\r\nLocation:\r\n"
        + folder_path
        + "Consolidated report.xlsx",
        64,
        "Report Complete",
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
      xl.MsgBox(
          "An error occurred:\r\n\r\n" + str(err),
          16,
          "Consolidated Report Tool",
      )


if __name__ == "__main__":
  create_consolidated_report()