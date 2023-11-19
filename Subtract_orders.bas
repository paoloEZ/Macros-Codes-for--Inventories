Attribute VB_Name = "Subtract_orders"
Sub SubtractQuantityOrders()
    Dim wsInventory As Worksheet
    Dim wsOrders As Worksheet
    Dim inventory As Range
    Dim Order As Range
    
    ' Set Inventory and Orders Worksheets
    Set wsInventory = ThisWorkbook.Sheets("Inventory")
    Set wsOrders = ThisWorkbook.Sheets("Orders")
    
    ' Define the data range on the Inventory sheet
    Set inventory = wsInventory.Range("A10:A" & wsInventory.Cells(Rows.Count, "A").End(xlUp).Row)
    
    ' Cycle through each row on the Orders shee
    For Each order In wsOrders.Range("A2:A" & wsOrders.Cells(Rows.Count, "A").End(xlUp).Row)
        ' Find the order code on the Inventory sheet
        Dim code As LongLong
        code = order.Value
        
        Dim quantityOrdered As Long
        quantityOrdered = order.Offset(0, 1).Value ' Quantity ordered in column B of Orders
        
        ' Perform the search in the Inventory range
        Dim found As Range
        Set found = inventory.Find(What:=code, LookIn:=xlValues, LookAt:=xlWhole)
        
        ' If the code was found in Inventory, subtract the quantity ordered
        If Not found Is Nothing Then
            Dim quantityInventory As Long
            quantityInventory = found.Offset(0, 7).Value ' Quantity in column H of Inventory
            found.Offset(0, 7).Value = quantityInventory - quantityOrdered ' Subtract the ordered quantity
            
            ' Clear the ordered quantity on the Orders sheet
            order.Offset(0, 1).ClearContents
        End If
    Next order
    
    ' Finish the process
    MsgBox "Inventory has been updated.", vbInformation
End Sub



