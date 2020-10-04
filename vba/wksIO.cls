VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wksIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' Sheet: IO-Demo
' Demo of the methods provided by the basIO module.
'
' LICENSE: GNU General Public License 3.0
'
' @platform    Excel 2010 (Windows 7)
' @package     excel-vba (https://github.com/cwsoft/excel-vba)
' @author      cwsoft (http://cwsoft.de)
' @copyright   cwsoft
' @license     http://www.gnu.org/licenses/gpl-3.0.html
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' SHEET EVENT HANDLER - DON´T CHANGE ANYTHING BELOW THIS LINE UNLESS YOU KNOW WHAT YOU DO
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub cmdImportTextfile_Click()
   ' import delimited textfile(s)
   Call basIO.importFileToSheet(Sheetname:="IMPORT", FileFilter:="Textfile (*.txt),.txt", MultiSelection:=True, Verbose:=True)
End Sub

Private Sub cmdExportSheetAsCSV_Click()
   ' store content of actual sheet as CSV file
   Call basIO.exportSheetToCSVFile(wks:=Me, Verbose:=True)
End Sub
