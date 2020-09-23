Attribute VB_Name = "ModMain"
Option Explicit

'this module is simply to make the program have an XP-style

Public Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type

Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Public Const ICC_USEREX_CLASSES = &H200

Public Sub Main()
On Error Resume Next
Dim ticcRet As tagInitCommonControlsEx
With ticcRet
.lngSize = LenB(ticcRet)
.lngICC = ICC_USEREX_CLASSES
End With
InitCommonControlsEx ticcRet
On Error GoTo 0
frmMain.Show
End Sub


