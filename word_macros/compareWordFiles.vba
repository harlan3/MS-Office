Option Explicit

Sub compareWordFiles()

Dim strFolderA As String, strFolderB As String
Dim strFileSpec As String, strFileName As String
Dim objDocA As Word.Document, objDocB As Word.Document

strFolderA = "C:\test_folder\folderA\"
strFolderB = "C:\test_folder\folderB\"
'strFileSpec = "*.txt"
strFileSpec = "*.doc*"
strFileName = Dir(strFolderA & strFileSpec)

Do While strFileName <> vbNullString
    Set objDocA = Documents.Open(strFolderA & strFileName)
    Set objDocB = Documents.Open(strFolderB & strFileName)
    Application.CompareDocuments _
    OriginalDocument:=objDocA, _
    RevisedDocument:=objDocB, _
    Destination:=wdCompareDestinationNew
    objDocA.Close
    objDocB.Close
    strFileName = Dir
Loop

End Sub