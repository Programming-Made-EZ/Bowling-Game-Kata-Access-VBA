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
Private Rolls(20) As Integer
Private CurrentRoll As Integer


Public Sub Roll(NumberOfPins As Integer)
    Rolls(CurrentRoll) = NumberOfPins
    CurrentRoll = CurrentRoll + 1
End Sub

Public Function Score() As Integer
    Dim Fram As Integer
    Dim RollIndex As Integer
    
    For Frame = 1 To 10
        If IsStrike(RollIndex) Then
            Score = Score + StrikeBonus(RollIndex)
            RollIndex = RollIndex + 1
        ElseIf IsSpare(RollIndex) Then
            Score = Score + SpareBonus(RollIndex)
            RollIndex = RollIndex + 2
        Else
            Score = Score + SumOfBallsInFrame(RollIndex)
            RollIndex = RollIndex + 2
        End If
        
    Next Frame
End Function

Private Function IsSpare(RollIndex As Integer) As Boolean
    IsSpare = Rolls(RollIndex) + Rolls(RollIndex + 1) = 10
End Function

Private Function IsStrike(RollIndex As Integer) As Boolean
    IsStrike = Rolls(RollIndex) = 10
End Function

Private Function SpareBonus(RollIndex As Integer) As Integer
    SpareBonus = 10 + Rolls(RollIndex + 2)
End Function

Private Function StrikeBonus(RollIndex As Integer) As Integer
    StrikeBonus = 10 + Rolls(RollIndex + 1) + Rolls(RollIndex + 2)
End Function

Private Function SumOfBallsInFrame(RollIndex As Integer) As Integer
    SumOfBallsInFrame = Rolls(RollIndex) + Rolls(RollIndex + 1)
End Function
