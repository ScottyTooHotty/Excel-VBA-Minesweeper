VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CustomGame 
   Caption         =   "Custom Game"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3105
   OleObjectBlob   =   "CustomGame.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CustomGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iGameHeight As Integer
Dim iGameWidth As Integer

Private Sub MinesCustomSpin_Change()

    MinesCustomtxt.Value = MinesCustomSpin.Value

End Sub

Private Sub CancelButton_Click()

    CustomGame.Hide
    SelectDifficulty.LevelCustom.Value = False
    
End Sub

Private Sub HeightCustomSpin_Change()

    HeightCustomtxt.Value = HeightCustomSpin.Value
    MinesCustomSpin.Max = (HeightCustomtxt.Value * WidthCustomtxt.Value) - 1

End Sub

Private Sub OkButton_Click()

    iGameWidth = WidthCustomtxt.Value
    iGameHeight = HeightCustomtxt.Value

    CustomGame.Hide
    SelectDifficulty.PlayNowButton.Enabled = True
    
End Sub

Private Sub WidthCustomSpin_Change()

    WidthCustomtxt.Value = WidthCustomSpin.Value
    MinesCustomSpin.Max = (HeightCustomtxt.Value * WidthCustomtxt.Value) - 1

End Sub
