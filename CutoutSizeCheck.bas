Attribute VB_Name = "Module1"
Sub SizeCheck()
Attribute SizeCheck.VB_ProcData.VB_Invoke_Func = " \n14"
'
' SizeCheck Macro
'

'
    Sheets("CutOut Image").Select
    Range("I2").Select
    ActiveSheet.Pictures.Insert( _
        "L:\temp\Plan Checking Temp\image2021-04-16-111118-1.jpg").Select
    Selection.ShapeRange.ZOrder msoSendToBack
    
    Sheets("Block Coordinates").Select
    Range("B4:C4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("CutOut Coordinates").Select
    Range("A3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Plan Data").Select
    Cells.Select
    Selection.Columns.AutoFit
    Range("B18").Select
    Selection.Copy
    Sheets("CutOut Coordinates").Select
    Range("J6").Select
    ActiveSheet.Paste
    
    Sheets("CutOut Image").Select
    ActiveSheet.Shapes.Range(Array("Straight Arrow Connector 3")).Select
    Selection.ShapeRange.Height = 283.4645669291
    ActiveSheet.Shapes.Range(Array("Straight Arrow Connector 2")).Select
    Selection.ShapeRange.Height = 283.4645669291
    ActiveSheet.Shapes.Range(Array("Straight Arrow Connector 2", _
        "Straight Arrow Connector 3")).Select
    Selection.ShapeRange.Align msoAlignCenters, msoFalse
    Selection.ShapeRange.Align msoAlignMiddles, msoFalse
    Selection.ShapeRange.Group.Select
    Selection.ShapeRange.IncrementLeft 98.25
    Selection.ShapeRange.IncrementTop 9.75
    
    Sheets("CutOut Image").Select
    ActiveSheet.ChartObjects("Chart 4").Activate
    ActiveSheet.Shapes("Chart 4").IncrementLeft 1022.25
    ActiveSheet.Shapes("Chart 4").IncrementTop -3.75
    ActiveSheet.ChartObjects("Chart 4").Activate
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.Axes(xlCategory).MinimumScale = -5
    ActiveChart.Axes(xlCategory).MaximumScale = 5
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MinimumScale = -5
    ActiveChart.Axes(xlValue).MaximumScale = 5
    ActiveChart.ChartArea.Select
    ActiveSheet.Shapes("Chart 4").Height = 283.4645669291
    ActiveSheet.Shapes("Chart 4").Width = 283.4645669291
    ActiveSheet.Shapes("Chart 4").IncrementLeft -505.5
    ActiveSheet.Shapes("Chart 4").IncrementTop 184.5
    
    
    ActiveSheet.Shapes.Range(Array("Picture 1")).Select
    Selection.ShapeRange.Align msoAlignCenters, msoFalse
    Selection.ShapeRange.Align msoAlignMiddles, msoFalse
    Selection.ShapeRange.IncrementLeft 485.25
    Selection.ShapeRange.IncrementTop 171.75

    ActiveSheet.Shapes.Range(Array("Straight Arrow Connector 2")).Select
    Selection.ShapeRange.IncrementLeft -6
    ActiveSheet.Shapes.Range(Array("Straight Arrow Connector 2", _
        "Straight Arrow Connector 3")).Select
    Range("V34").Select
    ActiveSheet.Shapes.Range(Array("Straight Arrow Connector 3")).Select
    Range("Y28").Select
    ActiveSheet.ChartObjects("Chart 4").Activate
    ActiveSheet.Shapes.Range(Array("Picture 6")).Select
End Sub
