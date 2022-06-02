Option Explicit

' References Microsoft HTML Object Library
' References Microsoft Forms 2.0 Object Library

Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
        (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, _
        ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long


Sub copyWebsiteText()

    Application.ScreenUpdating = False
    Dim oDoc As Word.Document
    Dim clipboard As MSForms.DataObject
    Dim oRng As Word.Range
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim sLocalFilename As String
    sLocalFilename = Environ$("TMP") & "\urlmon.html"

    Dim sURL As String
    sURL = "https://github.com/harlan3/MS-Office-Macros/blob/main/copyWebsiteText_Word_Macro.txt"

    Dim bOk As Boolean
    bOk = (URLDownloadToFile(0, sURL, sLocalFilename, 0, 0) = 0)
    If bOk Then
        If fso.FileExists(sLocalFilename) Then

            Dim oHtml4 As MSHTML.IHTMLDocument4
            Set oHtml4 = New MSHTML.HTMLDocument

            Dim oHtml As MSHTML.HTMLDocument
            Set oHtml = oHtml4.createDocumentFromUrl(sLocalFilename, "")

            '* need to wait a little while the document parses
            '* because it is multithreaded
            While oHtml.readyState <> "complete"
                DoEvents  '* do not comment this out it is required to break into the code if in infinite loop
            Wend
            Debug.Assert oHtml.readyState = "complete"

            Set oRng = ActiveDocument.Range
            oRng.Text = oHtml.body.innerText
            If oRng.Characters.Last = Chr(13) Or oRng.Characters.Last = Chr(11) Then
                oRng.End = oRng.End - 1
            End If
            oRng.Text = Replace(Replace(oRng.Text, Chr(11), " "), Chr(13), " ")

            Set clipboard = New MSForms.DataObject
            clipboard.SetText (oRng.Text)
            clipboard.PutInClipboard
            
            Set oDoc = Documents.Add
            oDoc.Content.Paste
            
            oRng.Delete
            
        End If
    End If
    Application.ScreenUpdating = True
End Sub