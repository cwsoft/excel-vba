VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAPI 
   Caption         =   "Demo of the basAPI functions"
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   OleObjectBlob   =   "frmAPI.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' Form: frmDemoAPI
' Demo of an borderless but moveable user form using basAPI functions.
'
' LICENSE: GNU General Public License 3.0
'
' @platform    Excel 2010 (Windows 7)
' @package     excel-vba (https://github.com/cwsoft/excel-app)
' @requires    -
' @author      cwsoft (http://cwsoft.de)
' @copyright   cwsoft
' @license     http://www.gnu.org/licenses/gpl-3.0.html
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' FORM EVENT HANDLER -- DON´T CHANGE ANYTHING BELOW THIS LINE UNLESS YOU KNOW WHAT YOU DO
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub UserForm_Initialize()
   ' Remove border from this user form
   Call basAPI.removeBorder(form:=Me)
End Sub

Private Sub cmdClose_Click()
   ' Close user form
   Me.Hide
   Unload Me
End Sub

Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   ' Make borderless user form moveable via left mouse click inside user form
   If Button = xlPrimaryButton Then
      Call basAPI.moveFormLeftMouseButtonDown(form:=Me)
   End If
End Sub
