Attribute VB_Name = "PlayGame"
Option Explicit

Public iGameLevel As Integer
Public iGameHeight As Integer
Public iGameWidth As Integer
Public iMines As Integer

Sub PlayGameNow()

    SelectDifficulty.Show

End Sub

Sub SweepEm(ByRef rTarget As Range)

Dim lLastColumn As Long
Dim lLastCount As Long
Dim lLastRow As Long
Dim lThisCount As Long
Dim rBuild As Range
Dim rCell As Range
Dim rCell2 As Range
Dim rFinal As Range

    Set rBuild = CheckAround(rTarget)

    lLastCount = 1
    lThisCount = 1
    
    For Each rCell2 In rBuild.Cells
        Set rBuild = Union(rBuild, CheckAround(rCell2))
    Next rCell2

    Set rFinal = StripDupeCells(rBuild)
    
    lThisCount = rFinal.Cells.Count
    
    Do While lThisCount > lLastCount
        lLastCount = lThisCount
        Set rBuild = rFinal
        Set rFinal = Nothing
        
        For Each rCell2 In rBuild
            Set rBuild = Union(rBuild, CheckAround(rCell2))
        Next rCell2
        Set rFinal = StripDupeCells(rBuild)
        lThisCount = rFinal.Cells.Count
    Loop
      
    For Each rCell In rFinal
        If IsNumeric(rCell.Offset(-1, 0).Value) And rCell.Offset(-1, 0).Value <> "" Then
            If Sheet2.Range(rCell.Offset(-1, 0).Address).Value <> "O" Then
                Sheet2.Range(rCell.Offset(-1, 0).Address).Value = _
                    rCell.Offset(-1, 0).Value
            End If
        End If
        If IsNumeric(rCell.Offset(1, 0).Value) And rCell.Offset(1, 0).Value <> "" Then
            If Sheet2.Range(rCell.Offset(1, 0).Address).Value <> "O" Then
                Sheet2.Range(rCell.Offset(1, 0).Address).Value = _
                    rCell.Offset(1, 0).Value
            End If
        End If
        If IsNumeric(rCell.Offset(0, -1).Value) And rCell.Offset(0, -1).Value <> "" Then
            If Sheet2.Range(rCell.Offset(0, -1).Address).Value <> "O" Then
                Sheet2.Range(rCell.Offset(0, -1).Address).Value = _
                    rCell.Offset(0, -1).Value
            End If
        End If
        If IsNumeric(rCell.Offset(0, 1).Value) And rCell.Offset(0, 1).Value <> "" Then
            If Sheet2.Range(rCell.Offset(0, 1).Address).Value <> "O" Then
                Sheet2.Range(rCell.Offset(0, 1).Address).Value = _
                    rCell.Offset(0, 1).Value
            End If
        End If
        If IsNumeric(rCell.Offset(1, 1).Value) And rCell.Offset(1, 1).Value <> "" Then
            If Sheet2.Range(rCell.Offset(1, 1).Address).Value <> "O" Then
                Sheet2.Range(rCell.Offset(1, 1).Address).Value = _
                    rCell.Offset(1, 1).Value
            End If
        End If
        If IsNumeric(rCell.Offset(1, -1).Value) And rCell.Offset(1, -1).Value <> "" Then
            If Sheet2.Range(rCell.Offset(1, -1).Address).Value <> "O" Then
                Sheet2.Range(rCell.Offset(1, -1).Address).Value = _
                    rCell.Offset(1, -1).Value
            End If
        End If
        If IsNumeric(rCell.Offset(-1, 1).Value) And rCell.Offset(-1, 1).Value <> "" Then
            If Sheet2.Range(rCell.Offset(-1, 1).Address).Value <> "O" Then
                Sheet2.Range(rCell.Offset(-1, 1).Address).Value = _
                    rCell.Offset(-1, 1).Value
            End If
        End If
        If IsNumeric(rCell.Offset(-1, -1).Value) And rCell.Offset(-1, -1).Value <> "" Then
            If Sheet2.Range(rCell.Offset(-1, -1).Address).Value <> "O" Then
                Sheet2.Range(rCell.Offset(-1, -1).Address).Value = _
                    rCell.Offset(-1, -1).Value
            End If
        End If
    Next
    
    With Sheet2
    
        lLastRow = Sheet3.Cells.Find(What:="*", After:=Sheet3.Cells(1), _
            LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
            SearchDirection:=xlPrevious, MatchCase:=False).Row
            
        lLastColumn = Sheet3.Cells.Find(What:="*", After:=Sheet3.Cells(1), _
            LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, _
            SearchDirection:=xlPrevious, MatchCase:=False).Column
   
        .Range(rFinal.Address).Interior.ColorIndex = 15

        .Range(rFinal.Offset(1, 1).Address).Interior.ColorIndex = 15
        .Range(rFinal.Offset(-1, -1).Address).Interior.ColorIndex = 15
        .Range(rFinal.Offset(-1, 1).Address).Interior.ColorIndex = 15
        .Range(rFinal.Offset(1, -1).Address).Interior.ColorIndex = 15
        .Range(rFinal.Offset(0, 1).Address).Interior.ColorIndex = 15
        .Range(rFinal.Offset(1, 0).Address).Interior.ColorIndex = 15
        .Range(rFinal.Offset(0, -1).Address).Interior.ColorIndex = 15
        .Range(rFinal.Offset(-1, 0).Address).Interior.ColorIndex = 15
        
        .Range("A:A").Interior.ColorIndex = xlNone
        .Range("1:1").Interior.ColorIndex = xlNone
        .Range(.Cells(lLastRow, 1), .Cells(lLastRow, _
            Columns.Count)).Interior.ColorIndex = xlNone
        .Range(.Cells(1, lLastColumn), .Cells(Rows.Count, _
            lLastColumn)).Interior.ColorIndex = xlNone
        
    End With
        
End Sub

Function CheckAround(rCell As Range) As Range

Dim rBuild As Range
    
    Set rBuild = rCell
       
    If rCell.Offset(-1, 0).Value = "" Then
        Set rBuild = Union(rBuild, rCell.Offset(-1, 0))
    End If
    
    If rCell.Offset(1, 0).Value = "" Then
        Set rBuild = Union(rBuild, rCell.Offset(1, 0))
    End If
    
    If rCell.Offset(0, -1).Value = "" Then
        Set rBuild = Union(rBuild, rCell.Offset(0, -1))
    End If
    
    If rCell.Offset(0, 1).Value = "" Then
        Set rBuild = Union(rBuild, rCell.Offset(0, 1))
    End If
    
    If rCell.Offset(1, 1).Value = "" Then
        Set rBuild = Union(rBuild, rCell.Offset(1, 1))
    End If
    
    If rCell.Offset(1, -1).Value = "" Then
        Set rBuild = Union(rBuild, rCell.Offset(1, -1))
    End If
    
    If rCell.Offset(-1, 1).Value = "" Then
        Set rBuild = Union(rBuild, rCell.Offset(-1, 1))
    End If

    If rCell.Offset(-1, -1).Value = "" Then
        Set rBuild = Union(rBuild, rCell.Offset(-1, -1))
    End If

    Set CheckAround = rBuild

End Function

Function StripDupeCells(rMyRange As Range) As Range
    
Dim rFinal As Range
Dim rCell As Range
     
    For Each rCell In rMyRange
        If rFinal Is Nothing Then
            Set rFinal = rCell
        ElseIf Intersect(rCell, rFinal) Is Nothing Then
            Set rFinal = Union(rFinal, rCell)
        End If
    Next rCell
     
    Set StripDupeCells = rFinal
     
End Function
