# Excelfiles

Pull

Sub PullData()
    Dim SourceWorkbook As Workbook
    Dim DestinationWorkbook As Workbook
    Dim SourceWorksheet As Worksheet
    Dim DestinationWorksheet As Worksheet
    Dim LastRow As Long
    
    ' Open the source workbook
    Set SourceWorkbook = Workbooks.Open("C:\Users\Ken\Downloads\Main.xlsx")
    ' Set the source worksheet
    Set SourceWorksheet = SourceWorkbook.Worksheets("Sheet1")
    
    ' Open the destination workbook
    Set DestinationWorkbook = ThisWorkbook ' Assumes the macro is in the destination workbook
    ' Set the destination worksheet
    Set DestinationWorksheet = DestinationWorkbook.Worksheets("Pull-Send")
    
    ' Find the last row in the source worksheet
    LastRow = SourceWorksheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Copy data from source to destination
    SourceWorksheet.Range("A1:C" & LastRow).Copy DestinationWorksheet.Range("A1")
    
    ' Close the source workbook
    SourceWorkbook.Close SaveChanges:=False
End Sub 

Save

Sub SaveChanges()
    Dim SourceWorkbook As Workbook
    Dim DestinationWorkbook As Workbook
    Dim SourceWorksheet As Worksheet
    Dim DestinationWorksheet As Worksheet
    
    ' Open the source workbook
    Set SourceWorkbook = ThisWorkbook
    ' Set the source worksheet
    Set SourceWorksheet = SourceWorkbook.Worksheets("Pull-Send")
    
    ' Open the destination workbook
    Set DestinationWorkbook = Workbooks.Open("C:\Users\Ken\Downloads\Main.xlsx") ' Assumes the macro is in the destination workbook
    ' Set the destination worksheet
    Set DestinationWorksheet = DestinationWorkbook.Worksheets("Sheet1")
    
    ' Copy edited data from destination to source
    SourceWorksheet.Range("A1:D100").Copy DestinationWorksheet.Range("A1")
    
    ' Save changes in the source workbook
    DestinationWorkbook.Save
    
    ' Close the source workbook
    DestinationWorkbook.Close ' SaveChanges:=True
End Sub
