Option Explicit

Sub calcPi()

Application.ScreenUpdating = False

Dim r(0 To 2801)
Dim i As Long
Dim k As Long
Dim b As Long
Dim d As Long
Dim c As Long
Dim oRng As Word.Range
Dim x As String

c = 0

Set oRng = ActiveDocument.Range
            
For i = 0 To 2800
    r(i) = 2000
Next i

For k = 2800 To 1 Step -14
    d = 0
    i = k
    Do While True
    
        d = d + r(i) * 10000
        b = 2 * i - 1
        
        r(i) = d Mod (b)
        d = d \ b
        i = i - 1
        If i = 0 Then Exit Do
        d = d * i
    Loop
    
    x = Format(Int((c + d \ 10000)), "#0000")
    oRng.InsertAfter (x)
    c = d Mod (10000)
Next k

Application.ScreenUpdating = True

End Sub