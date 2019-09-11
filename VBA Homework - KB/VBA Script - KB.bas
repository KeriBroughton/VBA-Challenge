Attribute VB_Name = "Module1"
Sub Multipleyearstock()
Dim Ticker_Letter As String
    Dim Summary_Table_Row2 As Integer
    Dim Volume As Double
    Dim Yearly_Change As Double
    
    Volume = 0
    Yearly_Change = 0
    Summary_Table_Row = 2
   
    
    For i = 2 To 760192
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    Ticker_Letter = Cells(i, 1).Value
    
    Volume = Volume + Cells(i, 7).Value
    Yearly_Change = Yearly_Change + Cells(i, 3).Value - Cells(i, 6).Value
    
    Range("I" & Summary_Table_Row).Value = Ticker_Letter
    Range("L" & Summary_Table_Row).Value = Volume
    Range("J" & Summary_Table_Row).Value = Yearly_Change
    
    Summary_Table_Row = Summary_Table_Row + 1
    
    
    Volume = 0
    Yearly_Change = 0

    Else

    Volume = Volume + Cells(i, 7).Value

    End If
Next i



End Sub


Sub Worksheet()

    Dim ws As Interior
    Dim i As Integer
    Dim k As Integer
    
        For Each ws In ActiveWorkbook.Sheets
    Next i
    Next ws
    

End Sub
