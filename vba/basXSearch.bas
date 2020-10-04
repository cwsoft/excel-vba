Attribute VB_Name = "basXSearch"
Option Explicit
Option Compare Text
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' Module: basXSearch
' Extended search allowing to find matches from multiple search pattern(s) more easily.
' Note: xSearch by default is case insensitive defined via: "Option Compare Text"
'
' LICENSE: GNU General Public License 3.0
'
' @platform    Excel 2010 (Windows 7)
' @package     excel-vba (https://github.com/cwsoft/excel-vba)
' @author      cwsoft (http://cwsoft.de)
' @copyright   cwsoft
' @license     http://www.gnu.org/licenses/gpl-3.0.html
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' constants to enable debugging for this module
Private Const DEBUG_MODE = True
Private Const DISPLAY_ERRORS = True
''

'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' VBA ROUTINES - DON´T CHANGE ANYTHING BELOW THIS LINE UNLESS YOU KNOW WHAT YOU DO
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Sub errorHandler(err As ErrObject, Source As String)
   ' error handler for this module
   If err.Number = 0 Or Not DISPLAY_ERRORS Then Exit Sub
   
   MsgBox "An error occured in '" & Source & "'." & vbCr _
      & "Err " & CStr(err.Number) & ": " & err.Description & "." & vbCr & vbCr _
      & "Please contact the application author to fix the error." _
      , vbExclamation + vbOKOnly, "VBA Application Error"
End Sub

'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' VBA FUNCTIONS - DON´T CHANGE ANYTHING BELOW THIS LINE UNLESS YOU KNOW WHAT YOU DO
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Public Function xGetLastCell(ByVal wks) As Range
   ' returns cell at bottom right corner of the worksheet's real used range
   Dim lLastCol As Long, lLastRow As Long
   
   If DEBUG_MODE Then On Error GoTo 0 Else On Error GoTo errHandler
   
   ' create worksheet object if needed
   If TypeName(wks) <> "Worksheet" Then
      If TypeName(wks) = "Range" Then Set wks = ThisWorkbook.Sheets(wks.Value)
      If TypeName(wks) = "String" Then Set wks = ThisWorkbook.Sheets(wks)
   End If
   
   lLastCol = 1: lLastRow = 1
   With wks.UsedRange
      lLastCol = .Cells.Find(What:="*" _
         , After:=.Cells(1) _
         , SearchOrder:=xlByColumns _
         , SearchDirection:=xlPrevious _
         , SearchFormat:=False).Column
      
      lLastRow = .Cells.Find(What:="*" _
         , After:=.Cells(1) _
         , SearchOrder:=xlByRows _
         , SearchDirection:=xlPrevious _
         , SearchFormat:=False).Row
   End With
   
   Set xGetLastCell = wks.Cells(lLastRow, lLastCol)
   
errHandler:
   Call basXSearch.errorHandler(err, Source:="basXSearch.xGetLastCell")
End Function

Public Function xGetValue(ByVal wks, ByVal Row, ByVal Column, Optional ByVal CheckIfNumeric) As Variant
   ' returns a sheet value from a specified worksheet range
   Dim vValue
   
   If DEBUG_MODE Then On Error GoTo 0 Else On Error GoTo errHandler
   
   ' create worksheet object if needed
   If TypeName(wks) <> "Worksheet" Then
      If TypeName(wks) = "Range" Then Set wks = ThisWorkbook.Sheets(wks.Value)
      If TypeName(wks) = "String" Then Set wks = ThisWorkbook.Sheets(wks)
   End If
   
   If IsMissing(CheckIfNumeric) Then CheckIfNumeric = False
  
   vValue = wks.Cells(Row, Column).Value
   vValue = IIf(CheckIfNumeric, IIf(IsNumeric(vValue), vValue, "#NAN"), vValue)
   xGetValue = vValue

errHandler:
   Call basXSearch.errorHandler(err, Source:="basXSearch.xGetValue")
End Function

Public Function xSearch(ByVal wks, ByVal Search1, ByVal Column1, ParamArray OptionalArgs()) As Variant
   ' Search worksheet for user defined search pattern(s). Returns row number of first matching main pattern.
   ' The search itself is case insensitive. Wrap search pattern in \search\ to perform a partial search.
   ' Usage: xSearch("Sheetname", "Search1", Col1, [startRow, endRow], ["Search2", Col2, Offset2], ... ["SearchN", ColN, OffsetN])
  
   ' OBLIGATORY PARAMETERS:
   '   wks:                 string of worksheet object of sheet to be searched
   '   Search1:             first SearchPattern to be found
   '   Column1:             column number where first searchPattern is expected
  
   ' OPTIONAL PARAMETERS:
   '   [lStart, lEnd]       limits search range to start/end row (define or ommit both)
  
   '   Search2,             Obligatory further search pattern
   '   Column2,             Obligatory search column of previous search pattern
   '   Offset2              Optional row offset
   '   ...
   '   SearchN,             Obligatory further search pattern
   '   ColumnN,             Obligatory search column of previous search pattern
   '   OffsetN              Optional row offset
  
   Const NUMERIC_TYPES As String = "#Integer#Long#Single#Double#"
   Dim rngSearch As Range
   Dim vStart, vEnd, vSearchN, vColumnN, vOffsetN
   Dim i As Long, lFirst As Long, lLast As Long
   Dim j As Integer, iNbrOptionalArgs As Integer, iFirstSubPatternIndex As Integer, iStep As Integer
   Dim bAllSubMatchesFound As Boolean
   
   If DEBUG_MODE Then On Error GoTo 0 Else On Error GoTo errHandler
   '--------------------------------------------------------------------
   ' extract obligatory main search parameters
   '--------------------------------------------------------------------
   ' create worksheet object if needed
   If TypeName(wks) <> "Worksheet" Then
      If TypeName(wks) = "Range" Then Set wks = ThisWorkbook.Sheets(wks.Value)
      If TypeName(wks) = "String" Then Set wks = ThisWorkbook.Sheets(wks)
   End If
   
   Column1 = CLng(Column1)

   '--------------------------------------------------------------------
   ' extract optional sub search parameters and start/end rows
   '--------------------------------------------------------------------
   ' set default values for start/end row
   vStart = 1: vEnd = basXSearch.xGetLastCell(wks).Row
   
   ' ensure we have the right amount of optional parameters
   iFirstSubPatternIndex = 0
   iNbrOptionalArgs = UBound(OptionalArgs, 1) + 1
   If iNbrOptionalArgs > 0 Then
      ' sanity check of optional parameter count
      If Not ((iNbrOptionalArgs - 2) Mod 3 = 0 Or iNbrOptionalArgs Mod 3 = 0) Then GoTo errArgs
   
      ' extract optional start/end row
      If (iNbrOptionalArgs - 2) Mod 3 = 0 Then
         ' ensure first two optional args are numbers, no strings allowed
         If Not NUMERIC_TYPES Like "*" & TypeName(OptionalArgs(0)) & "*" Then GoTo errArgs
         If Not NUMERIC_TYPES Like "*" & TypeName(OptionalArgs(1)) & "*" Then GoTo errArgs
         vStart = OptionalArgs(0)
         vEnd = OptionalArgs(1)
         iFirstSubPatternIndex = 2
      End If
   End If
   
   '--------------------------------------------------------------------
   ' find all user specified search patterns
   ' Note: If needed the xSearch time could be reduced by 33% using:
   '  a) for each instead index based loop (only for forward search)
   '  b) VBA Array instead rngSearch loop (forward/backward search)
   '--------------------------------------------------------------------
   ' set search range for main and sub patterns
   Set rngSearch = wks.Range(wks.Cells(vStart, Column1), wks.Cells(vEnd, Column1))
   
   ' set forward/backward search depending on user preference
   lFirst = 1: lLast = rngSearch.Cells.Count: iStep = 1
   If vStart > vEnd Then
      lFirst = lLast: lLast = 1: iStep = -1
   End If

   ' search for main pattern matches
   For i = lFirst To lLast Step iStep
      If rngSearch(i).Value Like Search1 Then
         ' quit if no sub patterns are defined
         If UBound(OptionalArgs, 1) + 1 < 3 Then GoTo matchFound

         ' check for sub pattern matches relative to main pattern
         bAllSubMatchesFound = False
         For j = iFirstSubPatternIndex To UBound(OptionalArgs, 1) Step 3
            ' extract sub matches
            vSearchN = OptionalArgs(j)
            vColumnN = OptionalArgs(j + 1)
            vOffsetN = OptionalArgs(j + 2)

            ' check if sub pattern matches (skip remaining sub patterns if no match found)
            bAllSubMatchesFound = rngSearch(i).Offset(vOffsetN, vColumnN - rngSearch(i).Column).Value Like vSearchN
            If Not bAllSubMatchesFound Then Exit For
         Next j

         ' stop search if all sub pattern match
         If bAllSubMatchesFound Then GoTo matchFound
      End If
   Next
      
noMatchFound:
   xSearch = -1
   GoTo errHandler
  
matchFound:
   xSearch = rngSearch(i).Row
   GoTo errHandler
  
errArgs:
  xSearch = "#ARGS"
  Call err.Raise(Number:=1000, Source:="basXSearch.xSearch", Description:="Invalid arguments supplied to the xSearch function")
  GoTo errHandler
  
errHandler:
   Call basXSearch.errorHandler(err, Source:="basXSearch.xSearch")
End Function


