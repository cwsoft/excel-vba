Attribute VB_Name = "basAPI"
Option Explicit
Option Private Module
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' Module: basAPI
' Windows API functions to be used in VBA (e.g. to manipulate forms etc.)
'
' LICENSE: GNU General Public License 3.0
'
' @platform    Excel 2010 (Windows 7)
' @package     excel-app (https://github.com/cwsoft/excel-app)
' @requires    -
' @author      cwsoft (http://cwsoft.de)
' @copyright   cwsoft
' @license     http://www.gnu.org/licenses/gpl-3.0.html
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' Windows API calls to remove user form title bar
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

' Windows API calls to make borderless user forms moveable
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

' Windows API calls to find window handle by process ID
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwprocessid As Long) As Long

' data structure to store window dimensions
Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' VBA ROUTINES - DON´T CHANGE ANYTHING BELOW THIS LINE UNLESS YOU KNOW WHAT YOU DO
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Public Sub removeBorder(form As Object)
   ' removes border from VBA form objects via Windows API call
   Const GW_CHILD As Integer = 5
   Const B_REDRAW As Boolean = True
   Const FORM_CAPTION As String = "WindowWithoutBorder"
   Dim RectMain As RECT, RectChild As RECT, RectNew As RECT
   Dim hRegion As Variant
   Dim hMainWindow As Long, hChildWindow As Long
   Dim bStatus As Boolean
   
   ' set caption and remove border from given UserForm
   With form
      .Caption = FORM_CAPTION
      .BorderStyle = fmBorderStyleSingle
   End With
   
   ' get window handle of user form which caption matches FORM_CAPTION
   hMainWindow = FindWindow(vbNullString, FORM_CAPTION)
   
   ' get child window of the main window found
   hChildWindow = GetWindow(hMainWindow, GW_CHILD)
   
   ' extract dimensions of main and child window
   bStatus = GetWindowRect(hMainWindow, RectMain)
   bStatus = GetWindowRect(hChildWindow, RectChild)
   
   ' compute new window dimensions without border
   RectNew.Left = 2
   RectNew.Top = (RectChild.Top - RectMain.Top) - 1
   RectNew.Right = (RectMain.Right - RectMain.Left) - 2
   RectNew.Bottom = (RectMain.Bottom - RectMain.Top) - 2

   ' create a new region with the updated dimensions
   hRegion = CreateRectRgn(RectNew.Left, RectNew.Top, RectNew.Right, RectNew.Bottom)
   
   ' update window region to the new dimensions
   bStatus = SetWindowRgn(hMainWindow, hRegion, B_REDRAW)
End Sub

Public Sub moveFormLeftMouseButtonDown(form As Object)
   ' allows moving borderless user forms around
   ' API: http://msdn.microsoft.com/en-us/library/windows/desktop/ff468877%28v=vs.85%29.aspx
   Const HTCAPTION = 2
   Const WM_NCLBUTTONDOWN = &HA1
   Dim hMainWindow As Long
   
   Call ReleaseCapture
   
   ' get user form handle and trigger mouse button event
   hMainWindow = FindWindow(vbNullString, form.Caption)
   Call SendMessage(hMainWindow, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&)
End Sub

Public Function getPIDFromWindowHandle(ByVal lWindowHandle) As Long
   ' returns process ID for the window matching the given window handle
   Dim lPID As Long
   
   ' get process ID for given window handle
   GetWindowThreadProcessId lWindowHandle, lPID
   
   getPIDFromWindowHandle = lPID
End Function

Public Function getWindowHandleFromPID(ByVal lProcessID As Long) As Long
   ' returns parent window handle of the window matching the given process ID
   ' http://support.microsoft.com/kb/242308
   Const GW_HWNDNEXT = 2
   Dim tempHwnd As Long
   
   ' grab first window handle that Windows finds
   tempHwnd = FindWindow(vbNullString, vbNullString)
   
   ' loop until you find a match or there are no more window handles
   Do Until tempHwnd = 0
      ' check if no parent for this window
      If GetParent(tempHwnd) = 0 Then
         ' check for PID match
         If lProcessID = getPIDFromWindowHandle(tempHwnd) Then
            ' return window handle matching PID and exit loop
            getWindowHandleFromPID = tempHwnd
            Exit Do
         End If
      End If
   
      ' get next window handle
      tempHwnd = GetWindow(tempHwnd, GW_HWNDNEXT)
   Loop
End Function

Public Function getWindowCaption(ByVal lWindowHandle As Long) As String
   ' returns the caption of the window identified by given window handle
   Dim sBuffer As String
   Dim iNumChars As Integer
   
   If lWindowHandle = 0 Then Exit Function
   
   ' initialize buffer
   sBuffer = Space$(128)
      
   ' get caption of window
   iNumChars = GetWindowText(lWindowHandle, sBuffer, Len(sBuffer))
   
   ' Display window's caption
   getWindowCaption = Left$(sBuffer, iNumChars)
End Function

Public Function isWindowActive(sCaption) As Boolean
   ' checks if a window with the defined window caption is opened
   isWindowActive = FindWindow(vbNullString, sCaption)
End Function
