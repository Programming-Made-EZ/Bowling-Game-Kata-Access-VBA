Attribute VB_Name = "Game_Tests"
'@TestModule
'@Folder("Tests")

Option Compare Database

Option Explicit
Option Private Module

Private Assert As New Rubberduck.AssertClass
Private Game As Game

'@TestMethod
Private Sub Score_ShouldReturn0_WhenGutterballGame()
    Set Game = New Game
    RollMany 20, 0
    Assert.AreEqual 0, Game.Score
End Sub

'@TestMethod
Private Sub Score_ShouldReturn20_WhenAllOnePinRolls()
    Set Game = New Game
    RollMany 20, 1
    Assert.AreEqual 20, Game.Score
End Sub

'@TestMethod
Private Sub Score_ShouldCalculateSpare_When10PinsInOneFrame()
    Set Game = New Game
    RollSpare
    Game.Roll 3
    RollMany 17, 0
    Assert.AreEqual 16, Game.Score
End Sub

'@TestMethod
Private Sub Score_ShouldCalculateStrike_When10PinsInOneRoll()
    Set Game = New Game
    
    RollStrike
    Game.Roll 3
    Game.Roll 4
    RollMany 16, 0
    
    Assert.AreEqual 24, Game.Score
End Sub

Private Sub RollMany(Rolls As Integer, Pins As Integer)
    Dim i As Integer
    For i = 1 To Rolls
        Game.Roll Pins
    Next i
End Sub

Private Sub RollSpare()
    Game.Roll 5
    Game.Roll 5
End Sub

Private Sub RollStrike()
    Game.Roll 10
End Sub
