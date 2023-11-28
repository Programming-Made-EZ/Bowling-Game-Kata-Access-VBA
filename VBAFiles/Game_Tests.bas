Attribute VB_Name = "Game_Tests"
'@TestModule
'@Folder("Tests")

Option Compare Database

Option Explicit
Option Private Module

    Private Assert As New Rubberduck.AssertClass

'@TestMethod("Uncategorized")
Private Sub Score_ShouldReturn0_WhenGutterballGame()
    Dim Game As New Game
    Dim i As Integer
    
    For i = 1 To 20
        Game.Roll 0
    Next i
    Assert.AreEqual 0, Game.Score
    
End Sub
