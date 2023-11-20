# Macros Codes for Inventories 
Here we have 4 Excel Macros to automate inventory management processes, fueron disenados para una empresa coreana que tiene sucursales en Bolivia con el proposito de manejar de forma mas practica informes y automatizar procesos del almacen de productos.

In this project we work entirely with the Visual Basic for Applications language, commonly known as VBA.

### Repository Structure:
This repository has 4 different Macros with functionalities designated for specific tasks according to the requirements of the requesting company, here is a small description of the functions of each one.
1. Add_new_product: We work with a sheet or table named "inventory", within the same book we create a sheet or also a table named "New Product", each product within these sheets is assigned a code that works as a primary key, When our VBA code comes into action, it looks for the products codes to match on both sheets and in this way the quantity in "New Product" is added to the quantity in "Inventory". In the event that the product code does not exist in the inventory, then copy the entire "New Product" row and add it to "Inventory".
2. Subtract_orders: We continue with the same sheet "Inventory" which is where all these processes are applied, basically it is the opposite of the previous macros, in this case we work with another sheet called "Orders", from which the quantities will be subtracted in inventory". Once the subtraction is done, the quantity records in "Orders" are cleared. If the record is not cleared then it means that the code for that product is not in our inventory.
3. copy_inventory_sheet: Make a copy of "inventory" in a new Excel document, with the name Inventory plus the date and time in which this copy was made, this to share with senior staff or colleagues who need to take a quick look .
4. sheet_for_deliveries: Make a copy of "inventory" excluding certain fields that are not necessary for store managers, the purpose is to remove irrelevant information that could be confused or misinterpreted in stores or branches. For example, store managers could confuse the product code with the code of the box where it is located. (Believe me, it happens more than you would think.)

### Conclusions:
With the development of these frameworks, several tasks were automated that facilitated the work of the company's warehouse manager.
Any observations and contributions are welcome, I hope they are useful to the data analysis community.

### Contact and Contributions:
Author: Paolo Eguino 
E-Mail: paoloeguinozerain@gmail.com
