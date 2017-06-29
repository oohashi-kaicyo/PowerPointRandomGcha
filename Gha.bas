Attribute VB_Name = "Module1"
Sub ChangeTarget(idx As Integer)
    Dim wait As Double
    wait = Timer
    While Timer < wait + 2
    Wend
    ActivePresentation.Slides(1).Shapes(idx).Fill.ForeColor.RGB = RGB(r, G, b)
End Sub

Function GenerateRandom()
Const low As Integer = 2
Const high As Integer = 10
Dim i As Integer
i = Int((high - low + 1) * Rnd + low)
GenerateRandom = i
End Function

Sub Gacha()

r = Int(0)
G = Int(0)
b = Int(0)
ChangeTarget GenerateRandom

End Sub
