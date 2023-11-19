Attribute VB_Name = "Add_new_product"
Sub UpdateInventoryFromNewProduct()
    Dim wsInventory As Worksheet
    Dim wsNewProduct As Worksheet
    Dim wsInventory As Range
    Dim NewProduct As Range
    Dim code As Variant
    Dim quantities As Variant
    
    ' Set the Inventory and NewProduct worksheets
    Set wsInventory = ThisWorkbook.Sheets("Inventory")
    Set wsNewProduct = ThisWorkbook.Sheets("NewProduct")
    
    ' Define the data range on the Inventory sheet
    Set inventory = wsInventory.Range("A2:A" & wsInventory.Cells(Rows.Count, "A").End(xlUp).Row)
    
    ' Loop through each row on the NewProduct sheet
    For Each Newproduct In wsNewProduct.UsedRange.Rows
        code = Newproduct.Cells(1, 1).Value ' Code in column A of NewProduct
        quantities = NewProduct.Cells(1, 8).Value ' Quantity in column H of NewProduct
        
        ' Find the code in the Inventory
        Dim found As Range
        Set found = inventory.Find(What:=code, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not found Is Nothing Then
            ' If the code exists in Inventory, add the quantity of NewProduct
            found.Offset(0, 7).Value = found.Offset(0, 7).Value + quantities
        Else
            ' If the code does not exist in Inventory, add a new row and copy the data from NewProduct
            Dim LastRow As Long
            LastRow = wsInventory.Cells(Rows.Count, "A").End(xlUp).Row + 1
            wsInventory.Cells(LastRow, 1).Resize(, 8).Value = NewProduct.Resize(, 8).Value
        End If
    Next Newproduct
    
    ' Finish the process
    MsgBox "Inventory update from NewProduct has been completed.", vbInformation
End Sub




