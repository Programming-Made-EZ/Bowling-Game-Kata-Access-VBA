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
        Select Case True
        Case IsStrike(FrameIndex)
            Score = Score + StrikeBonus(FrameIndex)
            FrameIndex = FrameIndex + 1
        Case IsSpare(FrameIndex)
            Score = Score + SpareBonus(FrameIndex)
            FrameIndex = FrameIndex + 2
        Case Else
            Score = Score + SumOfBallsInFrame(FrameIndex)
            FrameIndex = FrameIndex + 2
        End Select
    Next Frame
End Function

Private Function IsSpare(FrameIndex As Integer) As Boolean
    IsSpare = Rolls(FrameIndex) + Rolls(FrameIndex + 1) = 10
End Function

Private Function IsStrike(FrameIndex As Integer) As Boolean
    IsStrike = Rolls(FrameIndex) = 10
End Function

Private Function SpareBonus(FameIndex As Integer)
    SpareBonus = 10 + Rolls(FrameIndex + 2)
End Function

Private Function StrikeBonus(FrameIndex As Integer)
    StrikeBonus = 10 + Rolls(FrameIndex + 1) + Rolls(FrameIndex + 2)
End Function

Private Function SumOfBallsInFrame(FrameIndex As Integer)
    SumOfBallsInFrame = Rolls(FrameIndex) + Rolls(FrameIndex + 1)
End Function

