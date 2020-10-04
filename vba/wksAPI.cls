VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wksAPI"
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
' @requires    excel-app (https://github.com/cwsoft/excel-vba)
' @author      cwsoft (http://cwsoft.de)
' @copyright   cwsoft
' @license     http://www.gnu.org/licenses/gpl-3.0.html
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' SHEET EVENT HANDLER - DON´T CHANGE ANYTHING BELOW THIS LINE UNLESS YOU KNOW WHAT YOU DO
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub cmdBorderlessForm_Click()
   ' demo of a borderless, moveable user form
   Load frmAPI
   frmAPI.Show
End Sub
