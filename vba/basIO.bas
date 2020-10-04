Attribute VB_Name = "basIO"
Option Explicit
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' Module: basIO
' VBA routines to import ASCII textfile(s) as sheet(s) using the given delimiters and to
' export an Excel worksheet as textfile in CSV (comma seperated value) format.
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

' public variable to store status of file import
Public gbFileImportStatus As Boolean
'''

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

Public Sub importFileToSheet(ByVal Sheetname As String _
   , Optional ByVal FileFilter = "All Files (*.*), *.*" _
   , Optional ByVal MultiSelection = False _
   , Optional ByVal Verbose = True _
   , Optional ByVal customDelimiter = "" _
   , Optional ByVal useDelTAB = True _
   , Optional ByVal useDelSPACE = True _
   , Optional ByVal useDelCOMMA = False _
   , Optional ByVal useDelSEMICOLON = False _
   , Optional ByVal useConsecutiveDelimiter = True _
   )
   ' imports ASCII textfile(s) to active workbook using given delimiters
   Dim wks As Worksheet
   Dim vFilesToOpen As Variant, vFileList As Variant
   Dim sWkbPath, sBaseName As String, sNewSheet As String
   Dim iAnswer As Integer, iIndex As Integer, iMaxIndex As Integer, iNbrImportedFiles As Integer

   ' initialize public variable which controls success of file import
   gbFileImportStatus = False
   With Application
      .Calculation = xlCalculationManual
      .ScreenUpdating = False
      .DisplayAlerts = False
   End With

   '**************************************************************
   ' set drive and path of actual workbook (fix 255 char problem)
   '**************************************************************
   On Error Resume Next
   sWkbPath = ThisWorkbook.Path
   If Mid(sWkbPath, 2, 2) = ":\" Then ChDrive Left(sWkbPath, 1)
   ChDir (sWkbPath)
   err.Clear
   If DEBUG_MODE Then On Error GoTo 0 Else On Error GoTo errHandler

   '**************************************************************
   ' get file(s) user want to import and store in array vFileList
   '**************************************************************
   vFilesToOpen = Application.GetOpenFilename(FileFilter:=FileFilter _
      , Title:="Choose textfile" & IIf(MultiSelection, "s", "") & " to import" _
      , MultiSelect:=MultiSelection)

   ' check if a file was specified by the user
   If MultiSelection Then
      ' only returns boolean if dialogue was canceled by the user
      If TypeName(vFilesToOpen) = "Boolean" Then GoTo ImportCanceled
      vFileList = vFilesToOpen

      ' find max SheetName index (e.g. IMPORT_N) to start with
      iMaxIndex = 0
      For Each wks In ThisWorkbook.Worksheets
         If InStr(wks.Name, Sheetname & "_") > 0 Then
            iIndex = CInt(WorksheetFunction.Substitute(wks.Name, Sheetname & "_", ""))
            If iIndex > iMaxIndex Then iMaxIndex = iIndex
         End If
      Next
   Else
      If vFilesToOpen = False Then GoTo ImportCanceled
      vFileList = Array(vFilesToOpen)
   End If

   '**************************************************************
   ' import file(s) specified by the user
   '**************************************************************
   For iIndex = IIf(MultiSelection, 1, 0) To UBound(vFileList)
      ' add index to each sheet when in multi select mode
      sNewSheet = Sheetname & IIf(MultiSelection, "_" & CStr(iIndex + iMaxIndex), "")

      ' check if a sheet with the same name already exists
      For Each wks In ThisWorkbook.Worksheets
         If wks.Name = sNewSheet Then
            iAnswer = MsgBox("Workbook: '" & CStr(ThisWorkbook.Name) & "' already has a sheet named: '" & CStr(sNewSheet) + "'!" & vbCr _
               & "Delete existing sheet and continue with file import?" & vbCr & vbCr _
               & "Press 'Yes' to delete existing sheet and to continue, 'No' to cancel the import." _
               , vbQuestion + vbYesNo, "Sheet with same name already exist")
            If iAnswer = vbNo Then GoTo errHandler Else wks.Delete
         End If
      Next

      ' open specified textfile in a new workbook using OpenText method
      Workbooks.OpenText Filename:=vFileList(iIndex) _
         , Origin:=xlWindows _
         , DataType:=xlDelimited _
         , ConsecutiveDelimiter:=useConsecutiveDelimiter _
         , Tab:=useDelTAB _
         , Semicolon:=useDelSEMICOLON _
         , Comma:=useDelCOMMA _
         , Space:=useDelSPACE _
         , Other:=(customDelimiter <> "") _
         , OtherChar:=customDelimiter

      ' in case file import was successful the imported file was opened in new workbook
      sBaseName = ActiveWorkbook.Name
      If sBaseName = ThisWorkbook.Name Then GoTo errHandler

      ' rename first sheet of imported workbook and move it to calling workbook
      With ActiveWorkbook
         .Sheets(1).Name = sNewSheet
         .Sheets(1).Move After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
      End With
   Next iIndex
   
   gbFileImportStatus = True

   ' return status messages if we are supposed to
   If Verbose Then
      If gbFileImportStatus Then
         MsgBox "Import of " & CStr(UBound(vFileList, 1)) & " textfile(s) completed!", vbInformation + vbOKOnly, "File import finished"
      Else
         MsgBox "Import of specified textfile(s) failed.", vbExclamation + vbOKOnly, "File import failed"
      End If
   End If
   
   GoTo errHandler

ImportCanceled:
   gbFileImportStatus = False
   If Verbose Then MsgBox "File import canceled by the user!", vbInformation + vbOKOnly, "File import canceled"

errHandler:
   With Application
      .Calculation = xlCalculationAutomatic
      .DisplayAlerts = True
      .ScreenUpdating = True
   End With

   Call basIO.errorHandler(err, Source:="basIO.importFileToSheet")
End Sub

Sub exportSheetToCSVFile(wks As Worksheet, Optional Verbose As Boolean = True)
   ' exports active sheet as CSV text file
   Dim wkb As Workbook
   Dim vFileName As Variant
   Dim sSheetName As String
   Dim iAnswer As Integer

   ' copy values of specified sheet
   If DEBUG_MODE Then On Error GoTo 0 Else On Error GoTo errHandler
   Application.ScreenUpdating = False
   With wks
      sSheetName = .Name
      .Cells.Copy
   End With

   ' create new workbook and paste data
   Set wkb = Workbooks.Add
   ActiveWorkbook.Worksheets(1).Range("A1").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=True

   ' show file open dialogue and get filename of file user wants to open
   vFileName = Application.GetSaveAsFilename(InitialFilename:=sSheetName _
      , FileFilter:="CSV Textfile (*.csv), *.csv""" _
      , Title:="Export sheet as CSV file")
   
   If vFileName = False Then
      If Verbose Then MsgBox "File export was canceled by the user.", vbExclamation + vbOKOnly, "File export canceled"
      GoTo errHandler
   End If

   With wkb
      .Saved = True
      .SaveAs Filename:=vFileName, FileFormat:=xlCSV, CreateBackup:=False
      Application.DisplayAlerts = False
      .Close
   End With
   
   If Verbose Then MsgBox "Sheet '" & sSheetName & "' sucessfully exported to '" & CStr(vFileName) & "'." _
      , vbInformation + vbOKOnly, "Export to CSV file complete"

errHandler:
   With Application
      .DisplayAlerts = True
      .ScreenUpdating = True
   End With
   
   Call basIO.errorHandler(err, Source:="basIO.exportSheetToCSVFile")
End Sub
