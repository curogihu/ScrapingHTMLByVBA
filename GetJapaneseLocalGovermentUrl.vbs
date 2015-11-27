'Warning: it was created a few years ago. I didn't check whether it works properly until now.
Option Explicit

Const MAX_GEOGRAPHICAL_NAME As Long = 4

'Nov 26th, 2015 / I checked this url didn't work now and alternative one isn't existed.
'An utl of upper layer page, https://www.j-lis.go.jp/lasdec-archive/cms/17,11.html work until now.
Const MAIN_URL As String = "https://www.lasdec.or.jp/cms/1,0,69.html"

Const MAIN_SHEETNAME As String = "main"
Const WORK_SHEETNAME As String = "sub"

Sub mainProc()
  Application.DisplayAlerts = False
  Application.ScreeUpdating = False
  Worksheets(MAIN_SHEETNAME).Cells.Clear

  Call outputprefecture(MAIN_URL)
  Call editProc

  MsgBox ("This program have finished.")

  Application.DisplayAlerts = True
  Applicatin.ScreenUpdating = True
End Sub

Sub editProc()
  Dim tmpSheet As Worksheet
  Dim iLoop As Long

  With Worksheets(MAIN_SHEETNAME)

    If Not isExistSheet(WORK_SHEETNAME) Then
      ActiveWorkbook.Worksheets.Add(after:=Worksheets(Worksheets.Count)).Name = WORK_SHEETNAME
    End If

    Worksheets(WORK_SHEETNAME).Cells.Clear

    .Select

    .Cells(1, 1) = "Prefeture Name"
    .Cells(1, 2) = "Region Name1"
    .Cells(1, 3) = "Region Name2"
    .Cells(1, 4) = "Region Name3"
    .Cells(1, 5) = "Region Name4"
    .Cells(1, 6) = "URL"

    .Cells(1, 1).Interior.ColorIndex = 38
    .Cells(1, 2).Interior.ColorIndex = 36
    .Cells(1, 3).Interior.ColorIndex = 35
    .Cells(1, 4).Interior.ColorIndex = 34
    .Cells(1, 5).Interior.ColorIndex = 37
    .Cells(1, 6).Interior.ColorIndex = 39

    .Range(.Cells(1, 1), _
            .Cells(Rows.Count, 1).End(xlUp).Row, MAX_GEOGRAPHICAL_NAME + 2).Borders.LineStyle = True

    .Columns("A:F").Select
    .Columns("A:F").AdvancedFilter Action:=xlFilterInPlace, unique:=True

    'move unique data to sub sheet
    Selection.Copy Worksheets(WORK_SHEETNAME).Range("A1")

    'Without this condition, it may occurs error when there is unique data in main sheet
    If .FilterMode Then
      .ShowAllData
    End If

    .Cells.Clear

    Worksheets(WORK_SHEETNAME).Range("A:F").Copy Range("A1")
    Worksheets(WORK_SHEETNAME).Delete

    .Range("A1").Select
    .Columns("A:F").AutoFit

  End With

End Sub

Function isExistSheet(strSheetName As String) As Boolean
  Dim objSheet As Object

  isExistSheet = False

  For Each objSheet In ActiveWorkbook.Sheets

    If objSheet.Name = strSheetName Then
      isExistSheet = True
      Exit Function
    End If

  Next objSheet
End Function


Private Function outputPrefeture(Optional strURL As String = "", _
                                  Optional strRegionName As String = "", _
                                  Optional outputRowIndex As Long = 2, _
                                  Optional outputColumnIndex As Long = 1) As Long

  Dim strRetVal As String
  Dim re As regexp
  Dim mc As matchcollection
  Dim iLoop As Long

  Dim strTargetURL As String

  Dim tmpRegionName As String
  Dim linkFlg As Boolean
  Dim tmpURL As String

  If Not GetHtmlSource(strURL, strRetVal) Then
    MsgBox "it couldn't get any information", vbCritical
    Exit Function
  End If

  Set re = CreateObject("VBScript.RegExp")

  re.Pattern = "<area title=.*"
  re.Global = True

  Set mc = re.Execute(strRetVal)

  With Worksheets(MAIN_SHEETNAME)
    For iLoop = 0 To mc.Count - 1
      linkFlg = False

      strRegionName = GetRegionName(mc(iLoop))
      strTargetURL = GetUrl(mc(iLoop))

      'It has underlayer link when url head doesn't begin from "http"
      If Left(strTargetURL, 4) <> "http" Then
        strTargetURL = "https:www.lasdec.or.jp/cms/" & GetUrl(mc(iLoop))
        linkFlg = True
      End If

      If outputColumnIndex > 1 Then
        ' copy prefecture + region name
        .Range(.Cells(outputRowIndex - 1, 1), _
            .Cells(outputRowIndex - 1, outputColumnIndex)).Copy _
        Destination:=.Range(.Cells(outputRowIndex, 1), _
                              .Cells(outputRowIndex, outputColumnIndex))
      End If

      .Cells(outputRowIndex, outputColumnIndex) = strRegionName

      ' a prefecture column + four region columns + a URL column = six columns
      .Cells(outputRowIndex, MAX_GEOGRAPHICAL_NAME + 2) = strTargetURL

      outputRowIndex = outputRowIndex + 1

      If outputRowIndex Mod 500 = 0 Then
        MsgBox ("amount of outputline became " & outputRowIndex & ".")
      End If

      'when any underlayer are existed
      If linkFlg Then
        outputRowIndex = outputprefecture(strTargetURL, strRegionName, outputRowIndex, outputColumnIndex + 1)
      End If
    Next

  End With

  Set re = Nothing
  Set mc = Nothing

  outputprefecture = outpuRowIndex

End Function

Private Function GetRegionName(ByVal targetStr As String) As String
  Dim strTmp As String
  Dim foundIndex As Long

  strTmp = Replace(targetStr, "<area title=""", "")
  foundIndex = InStr(strTmp, """")

  If foundIndex = 0 Then
    GetRegionName = ""
    Exit Function
  End If

  GetRegionName = Trim(Left$(strTmp, foundIndex - 1))
End Function

Private Function GetUrl(ByVal targetStr As String) As String
  Dim foundPrefixIndex As Long
  Dim foundSuffixIndex As Long

  Dim strPrefix As String
  Dim strSuffix As String
  Dim strTmp As String

  strPrefix = "href=""./"
  strSuffix = """"

  'delete an unnecessary part of head
  foundPrefixIndex = InStr(targetStr, strPrefix)

  If foundPrefixIndex = 0 Then
    strPrefix = "href="""

    foundPrefixIndex = InStr(targetStr, strPrefix)

    If foundPrefixIndex = 0 Then
      GetUrl = ""
      Exit Function
    End If
  End If

  strTmp = Trim(Mid$(targetStr, foundPrefixIndex + Len(strPrefix)))
  foundSuffixIndex = InStr(strTmp, strSuffix)

  If foundSuffixIndex = 0 Then
    GetUrl = ""
    Exit Function
  End If

  'delete an unnecessary part of tail
  GetUrl = Trim(Left$(strTmpp, foundSuffixIndex - 1))
End Function

Private Function GetHtmlSource(ByVal strURL As String, ByRef strRetVal As String) As Boolean
  Dim oHttp As Object

  On Error Resume Next

  Set oHttp = CreateObject("MSXML2.XMLHTTP")

  If (Err.Number <> 0) Then
    Set oHttp = CreateObject("MSXML.XMLHTTPRequest")
  End If

  On Error GoTo 0

  If oHttp Is Nothing Then
    MsgBox "XMLHTTP object hadn't been created", vbCritical
    Exit Function
  End If

  oHttp.Open "GET", strURL, False

  'cache reset
  oHttp.setRequestHeader "If-Modified-Since", "Thu, 01 Jun 1970 00:00:00 GMT"

  oHttp.Send

  If (oHttp.Status < 200 Or oHttp.Status >= 300) Then Exit Function

  ' get html source
  strRetVal = oHttp.responseText

  Set oHttp = Nothing

  GetHtmlSource = True
  
End Function

