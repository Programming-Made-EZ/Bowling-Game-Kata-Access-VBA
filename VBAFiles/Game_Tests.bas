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
Private Sub Score_ShouldReturn0_WhenGutterBallGame()
    'Given
    Dim i As Integer
    
    'When
    For i = 1 To 20
        Game.Roll 0
    Next i
    
    'Then
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
Private Sub Score_ShouldAddSpareBonus_WhenSpareIsRolled()
    Dim i As Integer
    
    Game.Roll 5
    Game.Roll 5
    Game.Roll 3
    
    For i = 1 To 17
        Game.Roll 0
    Next i
    
    Assert.AreEqual 16, Game.Score
End Sub

'@TestMethod
Private Sub Score_ShouldAddStrikeBonus_WhenStrikeIsRolled()
    Dim i As Integer
    
    Game.Roll 10
    Game.Roll 3
    Game.Roll 4
    
    For i = 1 To 17
        Game.Roll 0
    Next i
    
    Assert.AreEqual 24, Game.Score
    
End Sub

'@TestMethod
Private Sub Score_ShouldBe300_WhenAPerfectGame()
    Dim i As Integer
    
    For i = 1 To 12
        Game.Roll 10
    Next i
    
    Assert.AreEqual 300, Game.Score
End Sub
