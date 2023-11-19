Attribute VB_Name = "copy_inventory_sheet"
Sub CopyAndSaveInventory()
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
    
    ' Set the destination sheet in the new workbook
    Set wsDestiny = wbNew.Sheets(1)
    
    ' Paste the data into the destination sheet
    wsDestiny.Paste Destination:=wsDestiny.Range("A1")
    
    ' Save the new workbook in the documents folder with the name: "Inventory + CurrentDate"
    wbNew.SaveAs route & "Inventory_" & Format(Now(), "yyyymmdd_hhmmss") & ".xlsx"
    
    ' Close the new workbook without saving changes to the original workbook
    wbNew.Close SaveChanges:=False
    
    MsgBox "the inventory has successfully copied and saved in the documents folder ", vbInformation
End Sub

