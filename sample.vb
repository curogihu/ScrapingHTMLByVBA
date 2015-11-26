Option Explicit

Const MAX_GEOGRAPHICAL_NAME as Long = 4
Const MAIN_URL as string = "https://www.lasdec.or.jp/cms/1,0,69.html"

Const MAIN_SHEETNAME as string = "main"
Const WORK_SHEETNAME as string = "sub"

Sub mainProc()
  Application.DisplayAlerts = false
  Application.ScreeUpdating = false
  Worksheets(MAIN_SHEETNAME).Cells.clear

  call outputPrefecture(MAIN_URL)
  call editProc

  Msgbox("This program have finished.")

  Application.DisplayAlerts = true
  Applicatin.ScreenUpdating = true
End Sub

Sub editProc()
  dim tmpSheet as Worksheet
  dim iLoop as long

  with worksheets(MAIN_SHEETNAME)

    if not isExistSheet(WORK_SHEETNAME) then
      Activeworkbook.worksheets.add(after:=worksheets(Worksheets.count)).name = WORK_SHEETNAME
    end if

    worksheets(WORK_SHEETNAME).cells.clear

    .select

    .cells(1, 1) = "Prefeture Name"
    .cells(1, 2) = "Region Name1"
    .cells(1, 3) = "Region Name2"
    .cells(1, 4) = "Region Name3"
    .cells(1, 5) = "Region Name4"
    .cells(1, 6) = "URL"

    .cells(1, 1).interior.colorIndex = 38
    .cells(1, 1).interior.colorIndex = 36
    .cells(1, 1).interior.colorIndex = 35
    .cells(1, 1).interior.colorIndex = 34
    .cells(1, 1).interior.colorIndex = 37
    .cells(1, 1).interior.colorIndex = 39

    .range(.cells(1, 1),
            .cells(Rows.count, 1).end(xlup).row, MAX_GEOGRAPHICAL_NAME + 2)).borders.linestyle = true

    .columns("A:F").select
    .columns("A:F").AdvancedFilter action:=xlFilterInPlace, unique:=true

    'move unique data to sub sheet
    selection.copy worksheets(WORK_SHEETNAME).range("A1")

    'Without this condition, it may occurs error when there is unique data in main sheet
    if .filterMode then
      .showalldata
    end if

    .cells.clear

    Worksheets(WORK_SHEETNAME).range("A:F").copy Range("A1")
    worksheets(WORK_SHEETNAME).delete

    .range("A1").select
    .columns("A:F").autofit

  end with

end sub

function isExistSheet(strSheetName as string) as boolean
  dim objSheet as object

  isExistSheet = false

  for each objSheet in activeworkbook.sheets

    if objSheet.name = strSheetName then
      isExistSheet = true
      exit function
    end if

  next objSheet
end function


private function outputPrefeture(optional strURL as string = "",
                                  optional strRegionName as string = "",
                                  optional outputRowIndex as long = 2,
                                  optional outputColumnIndex as Long = 1) as long

  dim strRetVal as string
  dim re as regexp
  dim mc as matchcollection
  dim iLoop as long

  dim strTargetURL as string

  dim tmpRegionName as string
  dim linkFlg as Boolean
  dim tmpURL as string

  if not GetHtmlSource(strURL, strRetVal) then
    msgbox "it couldn't get any information", vbcritical)
    exit function
  end if

  set re = createobject("VBScript.RegExp")

  re.pattern = "<area title=.*"
  re.Global = true

  set mc = re.execute(strRetVal)

  with worksheets(MAIN_SHEETNAME)
    for iLoop = 0 to mc.count - 1
      linkFlg = false

      strRegionName = getRegionName(mc(iLoop))
      strTargetURL = GetURL(mc(iLoop))

      'It has underlayer link when url head doesn't begin from "http"
      if left(strTargetURL, 4) <> "http" then
        strTargetURL = "https:www.lasdec.or.jp/cms/" & GetUrl(mc(iLoop))
        linkFlg = true
      end if

      if outputColumnIndex > 1 then
        // copy prefecture + region name
        .range(.cells(outputRowIndex - 1, 1),
                .cells(outputRowIndex - 1, outputColumnIndex)).copy _
        destination:= .Range(.cells(outputRowIndex, 1),
                              .cells(outputRowindex, outputColumnIndex))
      end if

      .cells(outputRowIndex, outputColumnIndex) = strRegionName

      ' a prefecture column + four region columns + a URL column = six columns
      .cells(outputRowindex, MAX_GEOGRAPHICAL_NAME + 2) = strTargetURL

      outputRowIndex = outputRowIndex + 1

      if outputRowIndex mod 500 = 0 then
        msgbox("amount of outputline became " & outputRowIndex & ".")
      end if

      'when any underlayer are existed
      if linkFlg then
        outputRowIndex = outputPrefecture(strTargetURL, strRegionName, outputRowIndex, outputColumnIndex + 1)
      end if
    next

  end with

  set re = nothing
  set mc = nothing

  outputprefecture = outpuRowIndex

end function

private function GetRegionName(ByVal targetStr as string) as string
  dim strTmp as string
  dim foundIndex as long

  strTmp = replace(targetStr, "<area title=""", "")
  foundIndex = instr(strTmp, """")

  if foundIndex = 0 then
    GetRegionName = ""
    Exit function
  end if

  GetRegionName = Trim(Left$(strTmp, foundIndex - 1))
end function

private function GetUrl(ByVal targetStr as string) as string
  dim foundPrefixIndex as long
  dim foundSuffixIndex as long

  dim strPrefix as string
  dim strSuffix as string
  dim strTmp as string

  strPrefix = "href=""./"
  strSuffix = """"

  'delete an unnecessary part of head
  foundPrefixIndex = instr(targetStr, strPrefix)

  if foundPrefixIndex = 0 then
    strPrefix = "href="""

    foundPrefixIndex = instr(targetStr, strPrefix)

    if foundPrefixIndex = 0 then
      GetUrl = ""
      exit function
    end if
  end if

  strTmp = trim(Mid$(targetStr, foundPrefixIndex + len(strPrefix)))
  foundSuffixIndex = instr(strTmp, strSuffix)

  if foundSuffixIndex = 0 then
    GetUrl = ""
    exit function
  end if

  'delete an unnecessary part of tail
  GetUrl = trim(left$(strTmpp, foundSuffixIndex - 1))
end function

private function GetHtmlSource(ByVal strURL as string, ByRef strRetVal as string) as boolean
  dim oHttp as Object

  on Error resume next

  set oHttp = Createobject("MSXML2.XMLHTTP")

  if(Err.number <> 0) then
    set oHttp = Createobject("MSXML.XMLHTTPRequest")
  end if

  on error goto 0

  if oHttp is nothing then
    msgbox "XMLHTTP object hadn't been created", vbcritical
    exit function
  end if

  oHttp.Open "GET", strURL, false

  'cache reset
  oHttp.setRequestHeader "If-Modified-Since", "Thu, 01 Jun 1970 00:00:00 GMT"

  oHttp.Send

  if(oHttp.Status < 200 or oHttp.status >= 300) then exit function

  ' get html source
  strRetVal = oHttp.responseText

  set oHttp = nothing

  GetHtmlSource = true
end function