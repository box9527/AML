VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "主要交易對手分析"
   ClientHeight    =   4640
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   6660
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

















Private Sub UserForm_Initialize()
    ' Clear the existing items in the list

    Exit Sub
    ' Add column headers
    ListBox1.ColumnCount = 3
    ListBox1.ColumnWidths = "80;80;80"
    'ListBox1.ListHeaderStyle = fmListHeaderTruncated
    ListBox1.ColumnHeads = Array("Column 1", "Column 2", "Column 3")
    ' Add data to the list
    ListBox1.AddItem Array("Item 1A", "Item 1B", "Item 1C")
    ListBox1.AddItem Array("Item 2A", "Item 2B", "Item 2C")
    ListBox1.AddItem Array("Item 3A", "Item 3B", "Item 3C")
    ' Add more items as needed
End Sub
