VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'此类使用P. Havrda, Czech Republic编写的递归代码，完成当前目录及子目录的文件探索。
'此类是对Excel2007以后版本取消FileSearch类的替代方案，由于采用递归方式，受限于Excel本身
'参数栈的大小，对于过量的递归可能出现溢出，具体递归层数限制未经测试。
 
Dim pLookIn As String
Dim pSearchSubFolders As Boolean
Dim pFileName As String

Public FoundFiles As New Collection

Property Get Item(ByVal vID As Integer) As String
'Attribute Item.VB_UserMemID=0
    Item = FoundFiles.Item(vID)
End Property

Public Function NewEnum() As IUnknown
'Attribute NewEnum.VB_UserMemID=-4
  Set NewEnum = FoundFiles.[_NewEnum]
End Function

Public Property Get LookIn() As String
    LookIn = pLookIn
End Property

Public Property Let LookIn(value As String)
    pLookIn = value
End Property

Public Property Get SearchSubFolders() As Boolean
    SearchSubFolders = pSearchSubFolders
End Property

Public Property Let SearchSubFolders(value As Boolean)
    pSearchSubFolders = value
End Property

Public Property Get FileName() As String
    FileName = pFileName
End Property

Public Property Let FileName(value As String)
    pFileName = value
End Property



Private Sub FileSearchByHavrda(pFoundFiles As Collection, pPath As String, pMask As String, pIncludeSubdirectories As Boolean)
'
' Search files in Path and create FoundFiles list(collection) of file names(path included) accordant with Mask (search in subdirectories if enabled)
' 01.06.2009, Author: P. Havrda, Czech Republic
'

    Dim DirFile As String
    Dim CollectionItem As Variant
    Dim SubDirCollection As New Collection

    ' Add backslash at the end of path if not present
    pPath = Trim$(pPath)
    If Right$(pPath, 1) <> "\" Then pPath = pPath & "\"

    ' Searching files accordant with mask
    DirFile = Dir$(pPath & pMask)
    Do While DirFile <> ""
        pFoundFiles.Add pPath & DirFile  'add file name to list(collection)
        DirFile = Dir$    ' next file
    Loop

    ' Procedure exiting if searching in subdirectories isn't enabled
    If Not pIncludeSubdirectories Then Exit Sub

    ' Searching for subdirectories in path
    DirFile = Dir$(pPath & "*", vbDirectory)
    Do While DirFile <> ""
        ' Add subdirectory to local list(collection) of subdirectories in path
        If DirFile <> "." And DirFile <> ".." Then If ((GetAttr(pPath & DirFile) And vbDirectory) = 16) Then SubDirCollection.Add pPath & DirFile
        DirFile = Dir$    'next file
    Loop

    ' Subdirectories list(collection) processing
    For Each CollectionItem In SubDirCollection
        Call FileSearchByHavrda(pFoundFiles, CStr(CollectionItem), pMask, pIncludeSubdirectories)      ' Recursive procedure call
    Next

End Sub

Public Function Execute() As Long

    'Dim FileNameWithPath As Variant
    'Dim FoundFiles As New Collection    ' create a collection of filenames

    ' Filling a collection of filenames (search Excel files including subdirectories)
    Call FileSearchByHavrda(FoundFiles, pLookIn, pFileName, pSearchSubFolders)


    ' Print list to immediate debug window and as a message window
    'For Each FileNameWithPath In FoundFiles    ' cycle for list(collection) processing
    '    Debug.Print FileNameWithPath & Chr(13)
    '    MsgBox FileNameWithPath & Chr(13)
    'Next FileNameWithPath

    ' Print to immediate debug window and message if no file was found
    'If FoundFiles.Count = 0 Then
    '    Debug.Print "No file was found !"
    'End If

    Execute = FoundFiles.Count

End Function
