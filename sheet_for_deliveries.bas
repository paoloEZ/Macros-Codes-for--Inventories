Attribute VB_Name = "sheet_for_deliveries"
Sub CopyAndSafeSheetForDeliveries()
    Dim wbNew As Workbook
    Dim wsOrigin As Worksheet, wsDestiny As Worksheet
    Dim route As String
    
    ' Set the documents folder path
    route = CreateObject("WScript.Shell").SpecialFolders("MyDocuments") & "\"
    
    ' Create a new Excel workbook
    Set wbNew = Workbooks.Add
    
    ' Set the source sheet (Inventory and the table inventory_table in this case)
    Set wsOrigin = ThisWorkbook.Sheets("Inventory")
    ThisWorkbook.Sheets("Inventory").ListObjects("inventory_table").Range.Copy
    
    ' Set the target sheet to the new workbook
    Set wsDestiny = wbNew.Sheets(1)
    
    ' Paste the data into the destination workbook, excluding the columns
    wsDestiny.Paste Destination:=wsDestiny.Range("A1")
    Application.CutCopyMode = False
    
    ' Delete the BOX CODE, SALES PRICE and LOCATION columns on the destination sheet
    wsDestiny.Range("B:B,E:E,G:G").EntireColumn.Delete
    
    ' Save the new workbook in the documents folder with the name "sheet_for_deliveries + CurrentDate
    wbNew.SaveAs route & "sheet_for_deliveries_" & Format(Now(), "yyyymmdd") & ".xlsx"
    
    ' Close the new workbook without saving changes to the original workbook
    wbNew.Close SaveChanges:=False
    
    MsgBox "The order form has been successfully copied and saved in the Documents folder.", vbInformation
End Sub


