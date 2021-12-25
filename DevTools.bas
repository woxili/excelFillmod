Attribute VB_Name = "DevTools"
Option Explicit

Public Const cStrQuote As String = """"

Public Declare Function SetTimer& Lib "user32" (ByVal hwnd&, _
        ByVal nIDEvent&, ByVal uElapse&, ByVal lpTimerFunc&)

Public Declare Function KillTimer& Lib "user32" (ByVal hwnd&, _
                                                 ByVal nIDEvent&)
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Public Const NV_INPUTBOX As Long = &H5000

Public Function IFERROR(ByRef ToEvaluate As Variant, ByRef Default As Variant) As Variant
'ToEvaluate 封装IsError
    If IsError(ToEvaluate) Then
        IFERROR = Default
    Else
        IFERROR = ToEvaluate
    End If
End Function

Public Sub RegFunction(ByVal sFname As String, ByRef sDescript As String)
'注册自定义函数
    Application.MacroOptions macro:=sFname, Description:=sDescript, Category:=14
End Sub

Public Sub DeRegFunction(ByVal sFname As String)
'解除注册自定义函数
    Application.MacroOptions macro:=sFname, Description:=Empty, Category:=Empty
End Sub

Public Sub unMergeForColumns(argRg As Range)  '将合并单元格拆分，并且将其内容复制到每个折分开的单元格中，
  Dim i As Integer, r As Integer, c As Integer
  Dim v As Variant
  With argRg
    For i = 1 To .Rows.Count
      If .Cells(i, 1).MergeCells Then
         v = .Cells(i, 1)
         If IsNumeric(v) And VarType(v) = vbString Then
           v = "'" & v
         End If
         c = .Cells(i, 1).MergeArea.Rows.Count
         .Cells(i, 1).UnMerge
         For r = 2 To c
           .Cells(i + r - 1, 1) = v
         Next r
         i = i + c - 1
       End If
    Next i
  End With
End Sub

Sub magerColumnCells(argRg As Range) '将一列中，内容相同的连续单元格合并成一个
  Dim i As Integer, c As Integer
  Dim rg As Range, mgRg As Range
  Dim str As String
  
  str = ""
  i = -1
  Application.DisplayAlerts = False
  For Each rg In argRg
    If str = "" Then
      str = rg.value
      Set mgRg = rg
    End If
    
    If str = rg.value Then
      i = i + 1
    Else
      Set mgRg = mgRg.Resize(mgRg.Rows.Count + i)
      mgRg.Merge
      str = rg.value
      Set mgRg = rg
      i = 0
    End If
  Next rg
  Application.DisplayAlerts = True
End Sub

Sub magerColumnCellsForRef(argRg As Range, argRefRg As Range) '按照参考区域中连续相同行，将目标列中相同
  Dim i As Integer, d As Integer, ct As Integer, cl As Integer '的连续单元合并成一个单元
  Dim rg As Range, refRg As Range, mgRg As Range
  
  ct = 1
  Application.DisplayAlerts = False
  For d = 1 To argRefRg.Rows.Count
    Set mgRg = argRefRg.Cells(d, 1)
    If mgRg.MergeCells Then
      cl = mgRg.MergeArea.Rows.Count
      Set mgRg = argRg.Cells(d, 1)
      Set mgRg = mgRg.Resize(mgRg.Rows.Count + cl - 1)
      mgRg.Merge
      'mgRg = ct
      ct = ct + 1
      d = d + cl - 1
    Else
      argRg.Cells(d, 1) = ct
      ct = ct + 1
    End If
  Next d
  Application.DisplayAlerts = True
End Sub



Function isSheetExist(argName As String) As Boolean
   Dim ws As Worksheet
   On Error GoTo 0
   On Error GoTo lbbad
   Set ws = ThisWorkbook.Worksheets(argName)
   isSheetExist = True
   Exit Function
lbbad:
  isSheetExist = False
End Function

Sub CellsReplace(argWS As Worksheet, strSrc As String, strTo As String)
  argWS.Cells.Replace What:=strSrc, Replacement:=strTo, LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub

Function FillParam(fmtStr As String, ParamArray argValues() As Variant) As Variant
  Dim i As Integer
  Dim tmpStr As String
  tmpStr = fmtStr
  For i = 0 To UBound(argValues)
    If TypeName(argValues(i)) = "String" Then
        tmpStr = Replace$(tmpStr, "%" & Trim(str(i)), Chr(34) & Trim(argValues(i)) & Chr(34))
    Else
      tmpStr = Replace$(tmpStr, "%" & Trim(str(i)), argValues(i))
    End If
  Next i
  FillParam = tmpStr
End Function

Function FillParam2(ByRef fmtStr As String, ParamArray argValues() As Variant) As Variant
  Dim i As Integer
  Dim regx As Object
  Dim tmpParam As String, tmpValue, tmpStr As String
  
  Set regx = CreateObject("VBScript.Regexp")
  'Set regx = New RegExp
  regx.Global = True
  tmpStr = fmtStr
  For i = 0 To UBound(argValues)
    tmpParam = "%" & Trim(str(i))
    regx.Pattern = tmpParam
    If TypeName(argValues(i)) = "String" Then
        tmpValue = Chr(34) & Trim(argValues(i)) & Chr(34)
        fmtStr = regx.Replace(fmtStr, tmpValue)
    Else
      tmpValue = argValues(i)
      fmtStr = regx.Replace(fmtStr, tmpValue)
    End If
  Next i
  FillParam2 = tmpStr
  Set regx = Nothing
End Function

Function ClearNPCharInStr(pStr As String) As String
'消除字符串中的特殊字符

    Dim ErrArray As Variant
    Dim oArray As Variant
    
    ErrArray = Array(" ", "\", "/", "*", "?", "<", ">", "|", """", Chr(9), Chr(10), Chr(13))
    For Each oArray In ErrArray '遍历替换
        pStr = Replace(pStr, oArray, "")
     Next
     ClearNPCharInStr = pStr
End Function

Sub DpMsg(msg As String)
    Debug.Print "DP:==>" & msg
End Sub

Function GetFileNameApart(argFilename As String, Optional argApart As Integer = 0) As String '0:filename,1:path,2:ext
  Dim i As Integer
  GetFileNameApart = ""
  i = InStrRev(argFilename, "\")
  If argApart = 0 Then
    If (i = 0) Then
      GetFileNameApart = argFilename
      Exit Function
    Else
      GetFileNameApart = Mid$(argFilename, i + 1)
    End If
  ElseIf argApart = 1 Then
    If (i = 0) Then
      GetFileNameApart = ""
      Exit Function
    Else
      GetFileNameApart = Left$(argFilename, i - 1)
    End If
  ElseIf argApart = 2 Then
    i = InStrRev(argFilename, ".")
    If i = 0 Then
      GetFileNameApart = ""
      Exit Function
    Else
      GetFileNameApart = Right$(argFilename, Len(argFilename) - i)
    End If
  End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : GenSumInFlagColumn
' Author    : Hp
' Date      : 2021/3/18
' Purpose   : 对指定的数据区域，按照包含合并单元格的数据列分组，对指定列中的分组数据进行汇总，并
'           ：将汇总公式放入指定列中每个分组的第一行
' Arg       :
' Sample    : GenSumInMergeColumn Selection, 1, 24, 26
'---------------------------------------------------------------------------------------
'
Sub GenSumInMergeColumn(argDataRG As Range, argMergeCol As Integer, argRefCol As Integer, argFormulaCol As Integer)
  Dim i As Integer, c As Integer
  Dim rg As Range
  
  With argDataRG
    For i = 1 To .Cells.Rows.Count
      Set rg = .Cells(i, argFormulaCol)
      If .Cells(i, argMergeCol).MergeCells Then
        c = rg.MergeArea.Rows.Count
        rg.Formula = "=sum(" & rg.Offset(0, argRefCol - argFormulaCol).Resize(c, 1).Address & ")"
        i = i + c - 1
      Else
        rg.Formula = "=" & rg.Offset(0, argRefCol - argFormulaCol).Address
      End If
    Next i
  End With
End Sub

Sub StatusPrograssBar(argStart As Integer, argMax As Integer, argSkip As Integer, Optional argInfo As String)
  Dim sChar As String
  Dim c As Double
  Dim i As Integer
  c = 50 / argMax
  i = Round(c * argSkip, 0)
  sChar = String(i, "■") & String(50 - i, "□")
  Application.StatusBar = "当前进度:  " & CStr(argSkip) & "/" & CStr(argMax) & " :  " & sChar & "  " & argInfo
End Sub

Function CreateListObject(ByRef argWS As Worksheet, argFilename As Variant, argPath As Variant, argDestRg As String) As ListObject
  Set CreateListObject = argWS.ListObjects.Add(SourceType:=0, _
                        Source:="ODBC;DSN=Excel Files;DBQ=" & argFilename & ";DefaultDir=" & argPath & ";DriverId=1046;MaxBufferSize=2048;PageTimeout=5;", _
                        Destination:=Range(argDestRg))
End Function

Function CreateQueryTable(ByRef argWS As Worksheet, argFilename As String, argPath As String, argDestRg As String, argSql As String) As QueryTable
  Dim connectStr As Variant
  Dim mp As String
  connectStr = "ODBC;DBQ=" & argFilename & ";DefaultDir=" & argPath & ";Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};" & _
             "DriverId=1046;FIL=excel 12.0;MaxBufferSize=2048;MaxScanRows=8;PageTimeout=5;ReadOnly=0;" & _
             "SafeTransactions=0;Threads=3;UID=admin;UserCommitSync=Yes;"
  'connectStr = FillParam(excelConnectionStr, argFilename, argPath)
  With argWS
    Set CreateQueryTable = .QueryTables.Add(connectStr, Range(argDestRg), argSql)
  End With
End Function

Sub GetDataFromExcelQrySql(argExcelFileName As String, argSql As String, Optional argWS As Worksheet = Null, Optional argRange As String = "A1")
  Dim qrt As QueryTable
  Dim ws As Worksheet
  If argWS Is Nothing Then
    Set ws = ActiveSheet
  Else
    Set ws = argWS
  End If
  Set qrt = CreateQueryTable(ws, argExcelFileName, GetFileNameApart(argExcelFileName, 1), argRange, argSql)
    With qrt
    .PreserveFormatting = True
    .RefreshStyle = xlInsertDeleteCells
    .PreserveColumnInfo = True
    .BackgroundQuery = False
    .AdjustColumnWidth = True
    .SaveData = True
    .Refresh
  End With
End Sub

Sub SearchFilesInCurrentDir(argSearchFor As String, argDestRg As Excel.Range, Optional argSubfolders As Boolean = False)
    Dim sFile As String
    Dim sPath As String
    Dim i As Integer

    sPath = ActiveWorkbook.Path
    Dim fs As New FileSearch
    With fs
        .LookIn = sPath
        .SearchSubFolders = argSubfolders
        .FileName = argSearchFor
        If .Execute > 0 Then
            For i = 1 To .FoundFiles.Count
                sFile = .FoundFiles.Item(i)
                argDestRg.Cells(i, 1) = sFile
                argDestRg.Hyperlinks.Add Anchor:=argDestRg.Cells(i, 1), Address:=sFile, TextToDisplay:=sFile
            Next
        End If
    End With
End Sub

'Sub ConvertWps2DocxInCurrentDir()
'    Dim sFile As String
'    Dim sPath As String
'    Dim i As Integer
'    Dim doc As Word.Document
'
'    sPath = ThisDocument.Path
'    Dim fs As New FileSearch
'    With fs
'        .LookIn = sPath
'        .SearchSubFolders = False
'        .FileName = "*.wps"
'        If .Execute > 0 Then
'            For i = 1 To .FoundFiles.Count
'                sFile = .FoundFiles.Item(i)
'                Set doc = Documents.Open(sFile, , True)
'                sFile = Left$(sFile, InStrRev(sFile, ".") - 1) & ".docx"
'                doc.SaveAs FileName:=sFile, FileFormat:=12, AddToRecentFiles:=False
'                doc.Close False
'                Set doc = Nothing
'            Next
'        End If
'    End With
'End Sub


'=============Ged Function============
Function myFilter(argRg As Range, argMatch As String, Optional argType As Integer) As Integer
  Dim v As Variant, f As Variant
  v = argRg.Columns(, 1)
  f = Filter(argRg, argMatch, True, argType)
  myFilter = UBound(f)
End Function

Function myMatch(argStr As String, argRg As Range, Optional argType As Double = 0) As Integer
  Dim rg As Range
   myMatch = -1
   For Each rg In argRg
     If StrComp(argStr, rg.value) = 0 Then
       myMatch = rg.Row
       Exit Function
      End If
   Next rg
End Function


Sub replaceStrInRangeCol(argRg As Range, argSrcStr As String, argDestStr As String, Optional argMatchCase As Boolean = False)
    argRg.Replace What:=argSrcStr, Replacement:=argDestStr, LookAt:=xlWhole, _
        SearchOrder:=xlByColumns, MatchCase:=argMatchCase, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub

Sub dp(argStr As Variant)
  Debug.Print argStr
End Sub

Function FindStrInCol(argRg As Range, argCol As Integer, argStr As Variant) As Integer
  Dim rg As Range
  Set rg = argRg.Columns(argCol).Find(argStr, , , xlWhole, xlByColumns)
  If Not rg Is Nothing Then
    FindStrInCol = rg.Row
  Else
    FindStrInCol = -1
  End If
 End Function


Function copyFilterSelectData(argSrcData As Range, argDestRg As Range, argField As Integer, argCriteria As Variant) As Range
  Dim aV As Variant
  Dim v As Variant
    With argSrcData
        .AutoFilter Field:=argField, Criteria1:=argCriteria, Operator:=xlFilterValues
        .Worksheet.Activate
        .Select
        argDestRg.CurrentRegion.Clear
        Selection.SpecialCells(xlCellTypeVisible).Copy 'Destination:=argDestRg
        argDestRg.PasteSpecial xlPasteValues
        .Worksheet.ShowAllData
        Set copyFilterSelectData = argDestRg.CurrentRegion
    End With
End Function

Function FillParamFromVariant(fmtStr As String, argValues As Variant) As String
  Dim i As Integer
  Dim tmpStr As String
  tmpStr = fmtStr
  For i = 0 To UBound(argValues)
    If TypeName(argValues(i)) = "String" Then
        tmpStr = Replace(tmpStr, "%" & Trim(CStr(i)), cStrQuote & Trim(argValues(i)) & cStrQuote)
    Else
      tmpStr = Replace(tmpStr, "%" & Trim(CStr(i)), argValues(i))
    End If
  Next i
  FillParamFromVariant = tmpStr
End Function



'---------------------------------------------------------------------------------------
' Procedure : ClearUserDefStyles
' Author    : Hp
' Date      : 2021/9/4
' Purpose   : 清除指定工作薄中的所有非内置样式
'---------------------------------------------------------------------------------------
'
Sub ClearUserDefStyles(argWB As Workbook)
   On Error Resume Next
   Dim s As Excel.Style
   With argWB
     For Each s In .Styles
       If Not s.BuiltIn Then s.Delete
     Next s
   End With
End Sub

Function QuickSearch(argSearchFor As Variant, argRange As Range, Optional argCol As Integer = 1) As Integer
  Dim i As Integer, n As Integer, r As Integer
  QuickSearch = -1
  With argRange
    r = .Rows.Count
    n = r Mod 2
    If 1 = n Then
      If argSearchFor = .Cells(r, argCol) Then QuickSearch = r: Exit Function
    End If
    For i = 1 To (r \ 2)
     If argSearchFor = .Cells(i, argCol) Then
       QuickSearch = i: Exit Function
     ElseIf argSearchFor = .Cells(r \ 2 + i, argCol) Then
       QuickSearch = r \ 2 + i: Exit Function
     End If
    Next i
  End With
End Function

'---------------------------------------------------------------------------------------
' Procedure : rgxFind
' Author    : Hp
' Date      : 2021/12/17
' Purpose   :
'---------------------------------------------------------------------------------------
'
Function rgxFind(argDestStr As Variant, argPattern As String, Optional argGloab As Boolean = True, _
                 Optional argIgnoreCase As Boolean = False, Optional argMultipleLines As Boolean = False) As String
  'Dim regx As Object
  Dim regx As RegExp
  'Dim m As Object, mc As Object
  Dim m As Match, mc As MatchCollection
  Dim strTmp As String
  
  
  'Set regx = CreateObject("VBScript.Regexp")
  Set regx = New RegExp
  
  With regx
    .MultiLine = argMultipleLines
    .IgnoreCase = argIgnoreCase
    .Global = argGloab
    .Pattern = argPattern
    Set mc = .Execute(argDestStr)
    If mc.Count > 0 Then
      For Each m In mc
        strTmp = strTmp + m.Value & ","
      Next m
      rgxFind = Left$(strTmp, Len(strTmp) - 1)
    Else
      rgxFind = ""
    End If
  End With
End Function
'========================test=======================
Sub DemoSearchFiles()
    Dim sFile As String
    Dim sPath As String
    Dim i As Integer

    sPath = ActiveWorkbook.Path
    Dim fs As New FileSearch
    With fs
        .LookIn = sPath
        .SearchSubFolders = True
        .FileName = "*.xlsm"
        If .Execute > 0 Then
            For i = 1 To .FoundFiles.Count
                sFile = .FoundFiles.Item(i)
                ActiveSheet.Cells(i, 1) = sFile
            Next
        End If
    End With
End Sub
