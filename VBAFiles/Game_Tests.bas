Attribute VB_Name = "Game_Tests"
'@TestModule
'@Folder("Tests")

Option Compare Database

Option Explicit
Option Private Module

Private Assert As New Rubberduck.AssertClass
Private Game As Game

'@TestInitialize
Private Sub TestInitialize()
    Set Game = New Game
End Sub

'@TestMethod
Private Sub Score_ShouldReturn0_WhenGutterballGame()
    Dim i As Integer
    
    For i = 1 To 20
        Game.Roll 0
    Next i
    
    Assert.AreEqual 0, Game.Score
End Sub

'@TestMethod
Private Sub Score_ShouldReturn20_WhenAllOnePinRolls()
    Dim i As Integer
    
    For i = 1 To 20
        Game.Roll 1
    Next i
    
    Assert.AreEqual 20, Game.Score
End Sub

'@TestMethod
Private Sub Foo()
    Dim i As Integer
    
    Game.Roll 5
    Game.Roll 5 'Spare
    Game.Roll 3
    RollMany 17, 0
    Assert.AreEqual 16, Game.Score
End Sub

'@TestMethod

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
