Attribute VB_Name = "Game_Tests"
'@TestModule
'@Folder("Tests")

Option Compare Database

Option Explicit
Option Private Module

    Private Assert As New Rubberduck.AssertClass

'@TestMethod
Private Sub Score_ShouldReturn0_WhenGutterballGame()
    Dim Game As New Game
    Dim i As Integer
    
    For i = 1 To 20
        Game.Roll 0
    Next i
    
    Assert.AreEqual 0, Game.Score
End Sub

'@TestMethod
Private Sub Score_ShouldReturn20_WhenAllOnePinRolls()
    Dim Game As New Game
    Dim i As Integer
    
    For i = 1 To 20
        Game.Roll 1
    Next i
    
    Assert.AreEqual 20, Game.Score
End Sub

'@TestMethod
Private Sub Foo()
    Dim Game As New Game
    Dim i As Integer
    
    Game.Roll 5
    Game.Roll 5 'Spare
    Game.Roll 3
        
    For i = 1 To 17
        Game.Roll 0
    Next i
    
    Assert.AreEqual 16, Game.Score
End Sub
