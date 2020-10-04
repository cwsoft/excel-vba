VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wksXSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' Sheet: basXSearch
' Demo of the methods provided by the basXSearch module.
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
Private Sub cmdSearch_Click()
   ' repeat xSearch tests using VBA code
   MsgBox basXSearch.xSearch(Me.Range("C16"), "tWo", 3, Me.Range("B17").Row, Me.Range("B24").Row)
   MsgBox basXSearch.xSearch("xSearch", "two", 3, Me.Range("B24").Row, Me.Range("B17").Row)
   MsgBox basXSearch.xSearch("xSearch", "Two", 3, "thReE", 4, 2)
   MsgBox basXSearch.xSearch("xSearch", "*em*", 4, "*es*", 2, -1)
   MsgBox basXSearch.xSearch("xSearch", "L0[1-5][SR]", 5)
   MsgBox basXSearch.xGetValue(Me.Range("C16"), xSearch("xSearch", "item", 4), 5)
   MsgBox basXSearch.xGetLastCell(Me.Range("C16")).Row
End Sub

