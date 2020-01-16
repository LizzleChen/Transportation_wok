Sub checkForFig()

    Dim iRows As Integer
    Dim iRowStart As Integer
    
    Dim iCols As Integer
    Dim iColStart As Integer
    Dim t As Integer
    Dim Excla As Integer
    
    
    ' RGB(180, 198, 231) Blue
    ' RGB(198, 224, 180) Green
    ' RGB(248, 203, 173) Pink
    

    iRows = 1000
    iRowStart = 10
    
    iCols = 700
    iColStart = 5
    t = 0
     For i = iRowStart To iRows
     For j = iColStart To iCols
    
    
    If Cells(i, j).Interior.Color = RGB(180, 198, 231) Or Cells(i, j).Interior.Color = RGB(248, 203, 173) Or Cells(i, j).Interior.Color = RGB(198, 224, 180) Then
    Excla = InStr(Cells(i, j).Formula, "!")
    CelRef = Mid(Cells(i, j).Formula, Excla + 1)
    
    If Cells(i, j).address(RowAbsolute:=False, ColumnAbsolute:=False) <> CelRef Then
    
    Cells(i, j).Interior.Color = RGB(255, 0, 0)
    t = t + 1
    
    End If
    End If

  Next
  Next
  MsgBox (t & " Error(s) Found!")
     
      
End Sub
