Attribute VB_Name = "Module1"
Sub CopyColouredFontTransactions()
    Dim TransIDField As Range
    Dim TransIDCell As Range
    Dim ATransWS As Worksheet
    Dim HTransWS As Worksheet
    Dim DestCol As String
    
    Set ATransWS = Worksheets("Courses")
    Set TransIDField = Union(ATransWS.Range("B2:B30"), ATransWS.Range("C2:C9"), ATransWS.Range("D2:D10"), ATransWS.Range("E2:E9"))
    Set HTransWS = Worksheets("Filler")
    
    ' Clear existing contents of "Filler" worksheet
    HTransWS.Cells.ClearContents
    
    For Each TransIDCell In TransIDField
        If Not TransIDCell.Font.Bold Then ' Check if font is not bold
            Select Case TransIDCell.Font.Color
                Case RGB(0, 204, 255)
                    DestCol = "A"
                Case RGB(192, 0, 0)
                    DestCol = "B"
                Case RGB(255, 0, 0)
                    DestCol = "C"
                Case RGB(255, 192, 0)
                    DestCol = "D"
                Case RGB(146, 208, 80)
                    DestCol = "E"
                Case RGB(0, 176, 80)
                    DestCol = "F"
                Case RGB(0, 112, 192)
                    DestCol = "G"
                Case RGB(112, 48, 160)
                    DestCol = "H"
                Case Else
                    DestCol = ""
            End Select
        
            If DestCol <> "" Then
                TransIDCell.Copy Destination:=HTransWS.Range(DestCol & Rows.Count).End(xlUp).Offset(1)
            End If
        End If
    Next TransIDCell
End Sub
Sub CheckForDuplicatesAndSubtractFromL10()
    Dim myRange As Range
    Dim myCell As Range
    Dim myValue As Variant
    Dim myDuplicateCells As New Collection
    Dim myThirdDigits As New Collection
    Dim i As Integer
    Dim subtractValue As Integer
    
    subtractValue = Range("L11").Value ' Get value from cell L10
    
    Set myRange = Range("B2:J7") ' Replace with your desired range
    
    For Each myCell In myRange
        myValue = myCell.Value
        If WorksheetFunction.CountIf(myRange, myValue) > 1 Then
            On Error Resume Next
            myDuplicateCells.Add myCell, CStr(myCell.Value) ' Add cell to collection of duplicates
            On Error GoTo 0
        End If
    Next myCell
    
    For i = 1 To myDuplicateCells.Count ' Loop through collection of duplicate cells
        myValue = myDuplicateCells.Item(i).Value
        myThirdDigits.Add Right(Left(myValue, Len(myValue) - 2), 1) ' Extract third digit from end of string
    Next i
    
    For i = 1 To myThirdDigits.Count ' Loop through collection of third digits
        subtractValue = subtractValue - myThirdDigits.Item(i) ' Subtract digit from value in L10
    Next i
    
    Range("L10").Value = subtractValue ' Input result into cell L10
End Sub
Sub RunMyMacro()
    ' Call your macro here
    CopyColouredFontTransactions
    CheckForDuplicatesAndSubtractFromL10
    ' Set the next time the macro should run
    Application.OnTime Now + TimeValue("00:01:00"), "RunMyMacro"
End Sub

