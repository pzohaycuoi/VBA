Sub Catdulieu()

'Rename data2.xlsx with Data file
'Rename Destination2.xls with Destination file
'All file must be open because we are using "Windows" here!
'Checked - all working as inteded with looping error

'Loop Until Data file is Empty
Do Until IsEmpty(Windows("data2.xlsx").ActiveSheet.Range("$A$4:$H$1932")) '!- Loop not working as intended - need to fix
'Create new Sheet from the First Sheet -> new sheet created is current active sheet in Destination file
    Windows("Destination2.xlsx").Activate
    Sheets("01").Select
    Sheets("01").Copy After:=Sheets(1)
'Copy company's name from Data file to "Kinh Gui" part of Destination file
    Windows("data2.xlsx").Activate
    ActiveSheet.Range("E4").Copy
    Windows("Destination2.xlsx").Activate
    ActiveSheet.Range("F6").PasteSpecial Paste:=xlPasteValues
'Copy First MST row from Data file to seller's MST cell in Destination file
    Windows("data2.xlsx").Activate
    ActiveSheet.Range("F4").Copy
    Windows("Destination2.xlsx").Activate
    ActiveSheet.Range("$K$8").PasteSpecial Paste:=xlPasteValues
'Using the MST just copied in Destination file to filter in Data file
    Windows("data2.xlsx").ActiveSheet.Range("$A$4:$H$2600").AutoFilter Field:=6, Criteria1:=Workbooks("Destination2.xlsx").ActiveSheet.Range("$K$8")
'Copy non-blank cells from Data file and paste value into Destination file
    Windows("data2.xlsx").ActiveSheet.Range("$A$4:$H$2600").SpecialCells(xlVisible).Copy
    Windows("Destination2.xlsx").ActiveSheet.Range("$A$13").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
'Delete all the copied cell
    Application.DisplayAlerts = False
    Windows("data2.xlsx").ActiveSheet.Range("$A$4:$H$2600").SpecialCells(xlVisible).Delete
'Un-filter in Data file => Prepare for next loop
    Windows("data2.xlsx").ActiveSheet.Range("$A$3:$H$2600").AutoFilter Field:=6
'Delete all blank row in Destination file
    Windows("Destination2.xlsx").ActiveSheet.Range("$C$13:$C$3320").SpecialCells(xlCellTypeBlanks).EntireRow.Delete


'Numbering in Destination file
    Dim lastRow As Long, counter As Long
    Dim cell As Range
    Dim ws As Worksheet
    Set ws = Windows("Destination2.xlsx").ActiveSheet

    lastRow = ws.Range("D" & ws.Rows.Count).End(xlUp).Row
    counter = 1
    For Each cell In ws.Range("A13:A" & lastRow)
        cell.Value = counter
        counter = counter + 1
    Next cell


'Re-name sheet with MST from seller's MST cell
    Windows("Destination2.xlsx").ActiveSheet.Name = Windows("Destination2.xlsx").ActiveSheet.Range("$F$13")


'Copy data from destination to XMHD list
    Dim wsSource As Worksheet: Set wsSource = Workbooks("Destination2.xlsx").ActiveSheet
    Dim wsDest As Worksheet: Set wsDest = Workbooks("ListingXMHD.xlsx").Sheets("So theo doi XMHD")

    countCheckPaste = wsSource.Range("D13", wsSource.Range("D13").End(xlDown)).Rows.Count 'Checked - Can go live
    If countCheckPaste = 65524 Or countCheckPaste = 1048564 Then countCheckPaste = 1
    wsDest.Range("H900").End(xlUp).Offset(1) = countCheckPaste

    SubtotalHD = wsSource.Range("G13").End(xlDown).Value 'Checked - Can go live
    wsDest.Range("I900").End(xlUp).Offset(1) = SubtotalHD

    Seller = wsSource.Range("E13") 'Checked - Can go live
    wsDest.Range("F900").End(xlUp).Offset(1) = Seller

    SellerAndMst = wsSource.Range("F13") 'Checked - Can go live
    wsDest.Range("G900").End(xlUp).Offset(1) = SellerAndMst

Loop


'Not working because the loop can't break
    Windows("ListingXMHD.xlsx").ActiveSheet.Range("$F$13:$F$296").SpecialCells(xlCellTypeBlanks).EntireRow.Delete

End Sub



