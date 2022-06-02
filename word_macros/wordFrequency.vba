Option Explicit

Sub WordFrequency()

    Application.ScreenUpdating = False

    Const maxwords = 9000          'Maximum unique words allowed
    Dim SingleWord As String       'Raw word pulled from doc
    Dim Words(maxwords) As String  'Array to hold unique words
    Dim Freq(maxwords) As Integer  'Frequency counter for unique words
    Dim WordNum As Integer         'Number of unique words
    Dim ByFreq As Boolean          'Flag for sorting order
    Dim ttlwds As Long             'Total words in the document
    Dim Excludes As String         'Words to be excluded
    Dim Found As Boolean           'Temporary flag
    Dim j, k, l, Temp As Integer   'Temporary variables
    Dim ans As String              'How user wants to sort results
    Dim tword As String            '
    Dim aword As Range
    Dim tmpName As String
    
    ' Set up excluded words
    Excludes = "[the][a][of][is][to][for][by][be][and][are]"

    ' Find out how to sort
    ByFreq = True
    ans = InputBox("Sort by WORD or by FREQ?", "Sort order", "WORD")
    If ans = "" Then End
    If UCase(ans) = "WORD" Then
        ByFreq = False
    End If
    
    Selection.HomeKey Unit:=wdStory
    System.Cursor = wdCursorWait
    WordNum = 0
    ttlwds = ActiveDocument.Words.Count

    ' Control the repeat
    For Each aword In ActiveDocument.Words
        SingleWord = Trim(LCase(aword))
        'Out of range?
        If SingleWord < "a" Or SingleWord > "z" Then
            SingleWord = ""
        End If
        'On exclude list?
        If InStr(Excludes, "[" & SingleWord & "]") Then
            SingleWord = ""
        End If
        If Len(SingleWord) > 0 Then
            Found = False
            For j = 1 To WordNum
                If Words(j) = SingleWord Then
                    Freq(j) = Freq(j) + 1
                    Found = True
                    Exit For
                End If
            Next j
            If Not Found Then
                WordNum = WordNum + 1
                Words(WordNum) = SingleWord
                Freq(WordNum) = 1
            End If
            If WordNum > maxwords - 1 Then
                j = MsgBox("Too many words.", vbOKOnly)
                Exit For
            End If
        End If
        ttlwds = ttlwds - 1
        StatusBar = "Remaining: " & ttlwds & ", Unique: " & WordNum
    Next aword

    ' Now sort it into word order
    For j = 1 To WordNum - 1
        k = j
        For l = j + 1 To WordNum
            If (Not ByFreq And Words(l) < Words(k)) _
              Or (ByFreq And Freq(l) > Freq(k)) Then k = l
        Next l
        If k <> j Then
            tword = Words(j)
            Words(j) = Words(k)
            Words(k) = tword
            Temp = Freq(j)
            Freq(j) = Freq(k)
            Freq(k) = Temp
        End If
        StatusBar = "Sorting: " & WordNum - j
    Next j

    ' Now write out the results
    tmpName = ActiveDocument.AttachedTemplate.FullName
    Documents.Add Template:=tmpName, NewTemplate:=False
    Selection.ParagraphFormat.TabStops.ClearAll
    With Selection
        For j = 1 To WordNum
            .TypeText Text:=Trim(Str(Freq(j))) _
              & vbTab & Words(j) & vbCrLf
        Next j
    End With
    System.Cursor = wdCursorNormal
    j = MsgBox("There were " & Trim(Str(WordNum)) & _
      " different words ", vbOKOnly, "Finished")
      
    Application.ScreenUpdating = True
    
End Sub