def Macro1():
    rng = ActiveSheet.Range("A1").CurrentRegion
    
    ActiveSheet.ListObjects.Add(xlSrcRange, rng, "", xlYes).Name = "Tableau1"
    
    ActiveSheet.ListObjects("Tableau1").TableStyle = "TableStyleMedium27"
    
    Selection.HorizontalAlignment = xlCenter
    Selection.VerticalAlignment = xlBottom
    Selection.WrapText = False
    Selection.Orientation = 0
    Selection.AddIndent = False
    Selection.IndentLevel = 0
    Selection.ShrinkToFit = False
    Selection.MergeCells = False
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeLeft).ColorIndex = 0
    Selection.Borders(xlEdgeLeft).TintAndShade = 0
    Selection.Borders(xlEdgeLeft).Weight = xlThin
    
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).ColorIndex = 0
    Selection.Borders(xlEdgeTop).TintAndShade = 0
    Selection.Borders(xlEdgeTop).Weight = xlThin
    
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).ColorIndex = 0
    Selection.Borders(xlEdgeBottom).TintAndShade = 0
    Selection.Borders(xlEdgeBottom).Weight = xlThin
    
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).ColorIndex = 0
    Selection.Borders(xlEdgeRight).TintAndShade = 0
    Selection.Borders(xlEdgeRight).Weight = xlThin
    
    Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
    Selection.Borders(xlInsideVertical).ColorIndex = 0
    Selection.Borders(xlInsideVertical).TintAndShade = 0
    Selection.Borders(xlInsideVertical).Weight = xlThin
    
    Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    Selection.Borders(xlInsideHorizontal).ColorIndex = 0
    Selection.Borders(xlInsideHorizontal).TintAndShade = 0
    Selection.Borders(xlInsideHorizontal).Weight = xlThin
    
    Cells.Select
    Cells.EntireColumn.AutoFit()
