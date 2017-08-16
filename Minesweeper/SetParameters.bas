Attribute VB_Name = "SetParameters"
Option Explicit

Sub SetGrid(iGameHeight As Integer, iGameWidth As Integer)

Dim iMine As Integer
Dim rBottomRow As Range
Dim rCell As Range
Dim rGrid As Range
Dim rLeftColumn As Range
Dim rMine As Range
Dim rMyRange As Range
Dim rRightColumn As Range
Dim rTopRow As Range

    Sheet2.Range(Sheet2.Cells(2, 1), Sheet2.Cells(Rows.Count, Columns.Count)).Clear
    Sheet2.Visible = True
    Sheet2.Select
    
    iGameHeight = iGameHeight + 1
    iGameWidth = iGameWidth + 1

    With Sheet2.Range(Sheet2.Cells(2, 2), Sheet2.Cells(iGameHeight, iGameWidth))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .ColumnWidth = 2
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
    End With

    ActiveWindow.DisplayGridlines = False
    
    With Sheet3
        Set rGrid = .Range(.Cells(2, 2), .Cells(iGameHeight, iGameWidth))
        Set rTopRow = .Range(.Cells(1, 1), .Cells(1, iGameWidth + 1))
        Set rBottomRow = _
            .Range(.Cells(iGameHeight + 1, 1), .Cells(iGameHeight + 1, _
            iGameWidth + 1))
        Set rLeftColumn = .Range(.Cells(1, 1), .Cells(iGameHeight + 1, 1))
        Set rRightColumn = _
            .Range(.Cells(1, iGameWidth + 1), .Cells(iGameHeight + 1, iGameWidth + 1))
    End With
    
    rTopRow.Value = "*"
    rBottomRow.Value = "*"
    rLeftColumn.Value = "*"
    rRightColumn.Value = "*"
    
    For Each rCell In rGrid
        If rCell.Value <> Chr(173) Then
            With Sheet3
                Set rMyRange = .Range(rCell.Offset(-1, -1), rCell.Offset(1, 1))
            End With
            For Each rMine In rMyRange
                If rMine.Value = Chr(173) Then
                    iMine = iMine + 1
                End If
            Next
            If iMine > 0 Then
                rCell.Value = iMine
            End If
            iMine = 0
        End If
    Next
    
    If Sheet2.Visible <> True Then
        Sheet2.Visible = True
    End If
    Sheet2.Select

End Sub

Sub SetMines(iGridHeight, iGridWidth, iTotalMines)

Dim iAcross As Integer
Dim iDown As Integer
Dim iMines As Integer
    
    Randomize
    
    Sheet3.Cells.Clear
    
    Do
        iDown = Int((iGridHeight) * Rnd) + 1
        iAcross = Int((iGridWidth) * Rnd) + 1
        If Sheet3.Cells(iDown + 1, iAcross + 1).Value = "" Then
            Sheet3.Cells(iDown + 1, iAcross + 1).Value = Chr(173)
            Sheet3.Cells(iDown + 1, iAcross + 1).Font.Name = "Wingdings"
            iMines = iMines + 1
        End If
    Loop While iMines < iTotalMines
    
    Sheet2.Range("A1").Value = "Mines Left:"
    Sheet2.Range("F1").HorizontalAlignment = xlCenter
    Sheet2.Range("F1").Value = iMines
End Sub

Sub ResetLeaderboard()

Dim rMyRange As Range

    With Sheet1
    
        If .ProtectContents = True Or .ProtectDrawingObjects = True _
            Or .ProtectScenarios = True Then
            .Unprotect
        End If
        
        Set rMyRange = .Range("B3:C12, F3:G12, B16:C25, F16:I25")
        
        rMyRange.Clear
        
        .Protect
    
    End With
        
End Sub
