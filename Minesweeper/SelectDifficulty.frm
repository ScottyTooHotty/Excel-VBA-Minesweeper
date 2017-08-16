VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectDifficulty 
   Caption         =   "Select Difficulty"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3105
   OleObjectBlob   =   "SelectDifficulty.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelectDifficulty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelButton_Click()

    Unload Me

End Sub

Private Sub LevelBeginner_Click()

    iGameLevel = 1
    PlayNowButton.Enabled = True

End Sub

Private Sub LevelCustom_Click()

    CustomGame.WidthCustomtxt.Value = 60
    CustomGame.HeightCustomtxt.Value = 32
    CustomGame.MinesCustomtxt.Value = 400
    CustomGame.Show

    iGameLevel = 4

End Sub

Private Sub LevelExpert_Click()

    iGameLevel = 3
    PlayNowButton.Enabled = True

End Sub

Private Sub LevelIntermediate_Click()

    iGameLevel = 2
    PlayNowButton.Enabled = True

End Sub

Private Sub PlayNowButton_Click()

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
            iGameWidth = CustomGame.WidthCustomtxt.Value
            iGameHeight = CustomGame.HeightCustomtxt.Value
            iMines = CustomGame.MinesCustomtxt.Value
        Case Else
    End Select
    
    With Sheet4
        .Range("B1").Value = iGameLevel
        .Range("B2").Value = iGameHeight
        .Range("B3").Value = iGameWidth
        .Range("B4").Value = iMines
    End With
    
    SelectDifficulty.Hide
    Unload Me
    
    Call SetMines(iGameHeight, iGameWidth, iMines)
    Call SetGrid(iGameHeight, iGameWidth)

End Sub

