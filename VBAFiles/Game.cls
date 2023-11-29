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
Private Rolls(21) As Integer
Private CurrentRoll As Integer

Public Sub Class_Initialize()
    CurrentRoll = 0
End Sub


Public Sub Roll(Value As Integer)
    Rolls(CurrentRoll) = Value
    CurrentRoll = CurrentRoll + 1
End Sub

Public Function Score()
    Dim FrameIndex As Integer
    Dim Frame As Integer
    
    Score = 0
    FrameIndex = 0
   
    For Frame = 1 To 10
        If Rolls(FrameIndex) + Rolls(FrameIndex + 1) = 10 Then
            Score = Score + 10 + Rolls(FrameIndex + 2)
        Else
            Score = Score + Rolls(FrameIndex) + Rolls(FrameIndex + 1)
        End If
        FrameIndex = FrameIndex + 2
    Next Frame
End Function
