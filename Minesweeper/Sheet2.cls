VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public sStart As Single

Private Sub Worksheet_BeforeDoubleClick(ByVal rTarget As Range, bCancel As Boolean)

Dim bFirstMove As Boolean
Dim lColumn As Long
Dim lRow As Long
Dim rCell As Range
Dim rMyRange As Range
Dim vResponse As Variant

    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With
    
    bFirstMove = False
    lColumn = Sheet3.UsedRange.Columns.Count
    lRow = Sheet3.UsedRange.Rows.Count
    Set rMyRange = Sheet3.UsedRange.Offset(1, 1).Resize(lRow - 2, lColumn - 2)
    
    If rTarget.Address = "$K$1" Then
    
        SelectDifficulty.Show
        
    ElseIf Not Intersect(rTarget, Range(rMyRange.Address)) Is Nothing Then
    
        lRow = rTarget.Row
        lColumn = rTarget.Column
        
        If Application.WorksheetFunction.CountA(Sheet2.UsedRange) = 3 Then
            bFirstMove = True
        Else
            bFirstMove = False
        End If
        
        If rTarget.Value <> "O" Then
            If Sheet3.Cells(lRow, lColumn).Value = Chr(173) Then
                For Each rCell In Sheet2.UsedRange
                    rCell.Select
                    If rCell.Value = "O" Then
                        If Sheet3.Cells(rCell.Row, rCell.Column).Value <> Chr(173) Then
                            rCell.Value = Chr(173)
                            rCell.Interior.ColorIndex = 3
                            With rCell.Borders(xlDiagonalDown)
                                .LineStyle = xlContinuous
                                .Weight = xlThin
                            End With
                            With rCell.Borders(xlDiagonalUp)
                                .LineStyle = xlContinuous
                                .Weight = xlThin
                            End With
                        End If
                    End If
                Next
                        
                Application.ScreenUpdating = True
                rTarget.Font.Name = "Wingdings"
                rTarget.Value = Chr(173)
                rTarget.Interior.ColorIndex = 3
                'MsgBox "You Lose!!!!!!", vbCritical, "Game Over"
                vResponse = MsgBox("GAME OVER!" & vbNewLine & vbNewLine & _
                    "Play Again?", vbYesNo + vbQuestion, "New Game")
                bCancel = True
                If vResponse = vbYes Then
                    'SelectDifficulty.Show
                    With Sheet2
                        .Cells.Clear
                        With .Range("K1")
                            .Value = "J"
                            .Font.Name = "Wingdings"
                            .Interior.ColorIndex = 6
                            .HorizontalAlignment = xlCenter
                            .VerticalAlignment = xlCenter
                            .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium, _
                                ColorIndex:=xlColorIndexAutomatic
                        End With
                    End With
                    
                    If iGameLevel <> 0 Then
                        Select Case iGameLevel
                            Case 1
                                iGameWidth = 9
                                iGameHeight = 9
                                iMines = 10
                            Case 2
                                iGameWidth = 16
                                iGameHeight = 16
                                iMines = 40
                            Case 3
                                iGameWidth = 30
                                iGameHeight = 16
                                iMines = 99
                            Case 4
                                iGameHeight = Sheet4.Range("B2").Value
                                iGameWidth = Sheet4.Range("B3").Value
                                iMines = Sheet4.Range("B4").Value
                            Case Else
                        End Select
                    Else
                        iGameLevel = Sheet4.Range("B1").Value
                        iGameHeight = Sheet4.Range("B2").Value
                        iGameWidth = Sheet4.Range("B3").Value
                        iMines = Sheet4.Range("B4").Value
                    End If
                    
                    'SelectDifficulty.Hide
                    'Unload Me
                    
                    Call SetMines(iGameHeight, iGameWidth, iMines)
                    Call SetGrid(iGameHeight, iGameWidth)
                Else
                    Sheet2.Cells.Clear
                    Sheet1.Select
                End If
                GoTo Terminate
            End If
            
            If Sheet3.Cells(lRow, lColumn).Value = "" Then
                Call SweepEm(Sheet3.Cells(lRow, lColumn))
            Else
                rTarget.Value = Sheet3.Cells(lRow, lColumn)
                rTarget.Interior.ColorIndex = 15
            End If
        
        End If
        
        For Each rCell In Sheet2.UsedRange
            Select Case rCell
                Case 1
                    With rCell.Font
                        .ColorIndex = 5
                        .Bold = True
                    End With
                Case 2
                    With rCell.Font
                        .ColorIndex = 10
                        .Bold = True
                    End With
                Case 3
                    With rCell.Font
                        .ColorIndex = 3
                        .Bold = True
                    End With
                Case 4
                    With rCell.Font
                        .ColorIndex = 11
                        .Bold = True
                    End With
                Case 5
                    With rCell.Font
                        .ColorIndex = 9
                        .Bold = True
                    End With
                Case 6
                    With rCell.Font
                        .ColorIndex = 33
                        .Bold = True
                    End With
                Case 7
                    With rCell.Font
                        .ColorIndex = 1
                        .Bold = True
                    End With
                Case 8
                    With rCell.Font
                        .ColorIndex = 15
                        .Bold = True
                    End With
                Case Else
            End Select
        Next
        
        Call CheckWin(bFirstMove)

    End If
    
    bCancel = True
    Sheet2.Range("K1").Interior.ColorIndex = 6
    
Terminate:
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With

End Sub

Private Sub Worksheet_BeforeRightClick(ByVal rTarget As Range, bCancel As Boolean)

Dim lColumn As Long
Dim lRow As Long
Dim rMyRange As Range

    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With

    lColumn = Sheet3.UsedRange.Columns.Count
    lRow = Sheet3.UsedRange.Rows.Count
    Set rMyRange = Sheet3.UsedRange.Offset(1, 1).Resize(lRow - 2, lColumn - 2)
    
    If Not Intersect(rTarget, Range(rMyRange.Address)) Is Nothing Then
        If rTarget.Value = "" Then
            rTarget.Font.Name = "Wingdings"
            rTarget.Value = "O"
            
            If Sheet3.Range(rTarget.Address).Value = "" Then
                Sheet3.Range(rTarget.Address).Value = "O"
            End If
            
            rTarget.Interior.ColorIndex = 35
            Sheet2.Range("F1").Value = Sheet2.Range("F1").Value - 1
            bCancel = True
            
            GoTo Terminate
        End If
        
        If rTarget.Value = "O" Then
            If Sheet3.Range(rTarget.Address).Value = "O" Then
                Sheet3.Range(rTarget.Address).Value = ""
            End If
            
            Sheet2.Range("F1").Value = Sheet2.Range("F1").Value + 1
            rTarget.Value = ""
            rTarget.Font.Name = "Arial"
            rTarget.Interior.ColorIndex = xlNone
        End If

        bCancel = True
    End If
    
Terminate:
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With
    
End Sub

Sub CheckWin(bFirstMove As Boolean)

Dim lCountBlanks As Long
Dim lCountFlags As Long
Dim lLastColumn As Long
Dim lLastRow As Long
Dim lName As Long
Dim lTotalMines As Long
Dim rCell As Range
Dim rGrid As Range
Dim rMyRange As Range
Dim sFinish As Single
Dim vName As Variant

    lLastRow = Sheet3.Cells.Find(What:="*", After:=Sheet3.Cells(1), _
        LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, MatchCase:=False).Row - 1
    lLastColumn = Sheet3.Cells.Find(What:="*", After:=Sheet3.Cells(1), _
        LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, _
        SearchDirection:=xlPrevious, MatchCase:=False).Column - 1
    
    With Sheet2
        Set rGrid = .Range(.Cells(2, 2), .Cells(lLastRow, lLastColumn))
    End With
    
    lCountFlags = Application.WorksheetFunction.CountIf(rGrid, "=O")
    
    For Each rCell In rGrid
        If rCell.Value = "" And rCell.Interior.ColorIndex = xlNone Then
            lCountBlanks = lCountBlanks + 1
        End If
    Next rCell
    
    lTotalMines = _
        Application.WorksheetFunction.CountIf(Sheet3.Range(rGrid.Address), Chr(173))
    
    If lCountFlags + lCountBlanks = lTotalMines Then
        For Each rCell In rGrid
            If rCell.Interior.ColorIndex = xlNone And rCell.Value = "" Then
                rCell.Interior.ColorIndex = 35
                rCell.Font.Name = "Wingdings"
                rCell.Value = "O"
            End If
        Next
        
        Application.ScreenUpdating = True
        sFinish = Timer - Sheet4.Range("B5").Value
        'sSeconds = sFinish - (60 * (sFinish \ 60))
        'sMintues = sFinish \ 60
        sFinish = Format(sFinish, "0.00")
        
        vName = InputBox("You Win!" & vbNewLine & "You completed the game in " & sFinish & _
            " seconds." & vbNewLine & vbNewLine & "Please enter your name:", "Congratulations")
            
        With Sheet1
            Select Case iGameLevel
                Case 1
                    Set rMyRange = .Range("A1:C12")
                    lName = 2
                Case 2
                    Set rMyRange = .Range("E1:G12")
                    lName = 6
                Case 3
                    Set rMyRange = .Range("A14:C25")
                    lName = 2
                Case 4
                    Set rMyRange = .Range("E14:I25")
                    lName = 6
                Case Else
            End Select
            
            .Select
            
            If .ProtectContents = True Or .ProtectDrawingObjects = True Or _
                .ProtectScenarios = True Then
                .Unprotect
            End If
            
            lLastRow = rMyRange.Find(What:="*", After:=rMyRange(1), LookIn:=xlFormulas, _
                LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
                MatchCase:=False).Row
            
            .Cells(lLastRow, lName).Value = vName
            .Cells(lLastRow, lName + 1).Value = sFinish
            
            If iGameLevel = 4 Then
                .Cells(lLastRow, lName + 2).Value = iGameHeight - 1 & "x" & iGameWidth - 1
                .Cells(lLastRow, lName + 3).Value = iMines
                
                .Range(.Cells(lLastRow - 10, lName), .Cells(lLastRow, lName + 3)).Sort _
                    Key1:=.Cells(lLastRow - 10, lName + 1), Order1:=xlAscending, _
                    Header:=xlYes
            Else
                .Range(.Cells(lLastRow - 10, lName), .Cells(lLastRow, lName + 1)).Sort _
                    Key1:=.Cells(lLastRow - 10, lName + 1), Order1:=xlAscending, _
                    Header:=xlYes
            End If
            
            .Protect
            
        End With
        
        Exit Sub
    End If
    
    If bFirstMove = True Then
        sStart = Timer
        Sheet4.Range("B5").Value = sStart
    End If

End Sub
