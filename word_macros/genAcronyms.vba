Option Explicit

Sub genAcronyms()

Application.ScreenUpdating = False
Dim StrTmp As String, StrAcronyms As String, i As Long, j As Long, k As Long, Rng As Range
StrAcronyms = "Acronym" & vbTab & "Term" & vbTab & vbCr
Dim oDoc_Source As Document
Dim oDoc_Target As Document
Dim oTable As Table
Dim tr As Range
Set oDoc_Source = ActiveDocument
Set oDoc_Target = Documents.Add
With oDoc_Source
  With .Range
    With .Find
      .ClearFormatting
      .Replacement.ClearFormatting
      .MatchWildcards = True
      .Wrap = wdFindStop
      .Text = "\([a-zA-Z0-9][a-z&A-Z&0-9]{1" & Application.International(wdListSeparator) & "}\)"
      .Replacement.Text = ""
      .Execute
    End With
    Do While .Find.Found = True
      StrTmp = Replace(Replace(.Text, "(", ""), ")", "")
      If (InStr(1, StrAcronyms, .Text, vbBinaryCompare) = 0) And (Not IsNumeric(StrTmp)) Then
        If .Words.First.Previous.Previous.Words(1).Characters.First = Right(StrTmp, 1) Then
          For i = Len(StrTmp) To 1 Step -1
            .MoveStartUntil Mid(StrTmp, i, 1), wdBackward
            .Start = .Start - 1
            If InStr(.Text, vbCr) > 0 Then
              .MoveStartUntil vbCr, wdForward
              .Start = .Start + 1
            End If
            If .Sentences.Count > 1 Then .Start = .Sentences.Last.Start
            If .Characters.Last.Information(wdWithInTable) = False Then
              If .Characters.First.Information(wdWithInTable) = True Then
                .Start = .Cells(.Cells.Count).Range.End + 1
              End If
            ElseIf .Cells.Count > 1 Then
              .Start = .Cells(.Cells.Count).Range.Start
            End If
          Next
        End If
        StrTmp = Replace(Replace(Replace(.Text, " (", "("), "(", "|"), ")", "")
        StrAcronyms = StrAcronyms & Split(StrTmp, "|")(1) & vbTab & Split(StrTmp, "|")(0) & vbCr
      End If
      .Collapse wdCollapseEnd
      .Find.Execute
    Loop
  End With
  StrAcronyms = Replace(Replace(Replace(StrAcronyms, " (", "("), "(", vbTab), ")", "")
  Set Rng = oDoc_Source.Range.Characters.Last
  With Rng
    .Collapse wdCollapseEnd
    .Style = "Normal"
    .Text = StrAcronyms
    Set oTable = .ConvertToTable(Separator:=vbTab, numrows:=.Paragraphs.Count, NumColumns:=2)
    With oTable
        Set tr = oDoc_Target.Range
        tr.Collapse wdCollapseEnd
        tr.FormattedText = oTable.Range.FormattedText
        tr.Collapse wdCollapseEnd
        tr.Text = vbCrLf
        .Delete
    End With
    .Collapse wdCollapseStart
    End With
  End With
oDoc_Target.Tables.Item(1).Rows(1).HeadingFormat = True
oDoc_Target.Tables.Item(1).Rows(1).Range.Font.Bold = True
Application.ScreenUpdating = True
End Sub