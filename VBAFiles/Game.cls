VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Game"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Private TotalScore As Integer

Public Sub Roll(Value As Integer)
    TotalScore = TotalScore + Value
End Sub

Public Function Score()
    Score = TotalScore
End Function
