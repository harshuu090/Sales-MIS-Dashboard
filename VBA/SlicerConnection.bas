Attribute VB_Name = "Module1"

Sub SlicerConnection()

If Sheet1.Range("A1").Value = True Then

    ActiveSheet.Shapes.Range(Array("Region")).Select
    ActiveWorkbook.SlicerCaches("Slicer_Region").PivotTables.AddPivotTable ( _
        ActiveSheet.PivotTables("PivotTable1"))
        
Else

    ActiveWorkbook.SlicerCaches("Slicer_Region").PivotTables.RemovePivotTable ( _
        ActiveSheet.PivotTables("PivotTable1"))
        
End If


End Sub
