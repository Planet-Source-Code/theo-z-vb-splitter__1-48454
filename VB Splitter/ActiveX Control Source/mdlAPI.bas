Attribute VB_Name = "mdlAPI"
Attribute VB_Description = "A module to declare Windows API procedures, functions, types and constants"
'*******************************************************************************
'** File Name  : mdlAPI.bas                                                   **
'** Language   : Visual Basic 6.0                                             **
'** Author     : Theo Zacharias (theo_yz@yahoo.com)                           **
'** Description: A module to declare Windows API procedures, functions, types **
'**              and constants                                                **
'** Usage      : 1. Gets the height of the item in list box controls or       **
'**                 other controls that inherit it                            **
'**              2. Clips the mouse pointer                                   **
'** Note       : * Currently, one of my purpose to develop programs using VB  **
'**                is to explore its capability as deep as I could go. That's **
'**                why I try to minimize the usage of external                **
'**                libraries/components which do not come with VB.            **
'**              * These APIs routines is not significantly important for the **
'**                VB Splitter control. Without those, you only lose          **
'**                ClipCursor property and get flickering effect in           **
'**                LiveUpdate mode for list box control or other controls     **
'**                that inherit it                                            **
'** Last modified on September 10, 2003                                        **
'*******************************************************************************

Option Explicit

'--- Constant Declaration
Public Const LB_GETITEMHEIGHT = &H1A1

'--- Types Declaration
Public Type POINTAPI
  x As Long
  y As Long
End Type

Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

'---------------------------------------------------------------------------
' See MSDN Library for the complete specification of each API routine below
'---------------------------------------------------------------------------

' Purpose    : Confines the cursor to a rectangular area lpRect on the screen
Public Declare Sub ClipCursor Lib "user32" (lpRect As RECT)
Attribute ClipCursor.VB_Description = "Confines the cursor to a rectangular area lpRect on the screen"

' Purpose    : Frees the cursor to move anywhere on the screen
Public Declare Sub _
  ClipCursorClear Lib "user32" Alias "ClipCursor" _
    (Optional ByVal lpRect As Long = 0&)
Attribute ClipCursorClear.VB_Description = "Frees the cursor to move anywhere on the screen"

' Purpose    : Retrieves the cursor's position in screen coordinates
Public Declare Sub GetCursorPos Lib "user32" (lpPoint As POINTAPI)
Attribute GetCursorPos.VB_Description = "Retrieves the cursor's position in screen coordinates"

' Purpose    : Sends message wMsg to window hWnd
' Usage      : Gets the height of the item in list box control or other controls
'              that inherit it
Public Declare Function _
  SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, _
     ByVal wParam As Long, lParam As Long) As Long
Attribute SendMessage.VB_Description = "Sends message wMsg to window hWnd"
