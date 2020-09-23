VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGetFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Class Written By:  GDuncan
'              On:  8/11/98
'Description:       Class provides an interface to the Win32 Find functions
'Dependences:       Win32 platform

' '-- Usage Example
'Declarations
'Dim WithEvents clsFind As CFindFile
'Private Sub Search()
'  Dim asfiles As Variant
'  Dim lLoop As Long
'  Dim lCount As Long
'  Dim bResult as Boolean
'
'  Set clsFind = New CFindFile
'
'  clsFind.Path = "D:\" 'UNC Paths are supported
'  clsFind.FileSpec = "FOLDER.MDB" 'Wild Cards are also supported
'
'  bResult = clsFind.FindAll(asfiles)
'
'  If VarType(asfiles) = (vbArray + vbString) Then
'    lCount = UBound(asfiles)
'    For lLoop = 0 To lCount
'      Debug.Print asfiles(lLoop)
'    Next lLoop
'  End If
'End Sub
'
'Private Sub clsFind_BeginFindFiles()
'  'staStatusBar.Panels.Item("Status").Text = "Begining Search"
'End Sub
'
'Private Sub clsFind_EndFindFiles(FileCount As Long)
'  'staStatusBar.Panels.Item("Status").Text = "Found " & CStr(FileCount) & " Files"
'End Sub
'
'Private Sub clsFind_FoundFile(FileName As String, Cancel As Boolean)
'  'staStatusBar.Panels.Item("Status").Text = "Found File: " & FileName
'  'staStatusBar.Refresh
'End Sub
' '-- End of Usage Example

' If the function succeeds, the return value is a search handle
' used in a subsequent call to FindNextFile or FindClose

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long


' Rtns True (non zero) on succes, False on failure

Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long

' Rtns True (non zero) on succes, False on failure

Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long


Private Const MaxLFNPath = 260

'FindFirstFile failure rtn value

Private Const INVALID_HANDLE_VALUE = -1

' SearchDirs() constants

Private Const vbBackslash = "\"

Private Const vbAllFiles = "*.*"

Private Const vbKeyDot = 46

Private Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
  dwFileAttributes As Long
  ftCreationTime As FILETIME
  ftLastAccessTime As FILETIME
  ftLastWriteTime As FILETIME
  nFileSizeHigh As Long
  nFileSizeLow As Long
  dwReserved0 As Long
  dwReserved1 As Long
  cFileName As String * MaxLFNPath
  cShortFileName As String * 14
End Type

Private WFD As WIN32_FIND_DATA
Private hItem&
Private hFile&

'local variable(s) to hold property value(s)
Private msFileSpec        As String
Private msPath            As String
Private mbSearching       As Boolean
Private mbRecursive       As Boolean
Private mlRecursiveLevels As Integer
Private mbUseFileSpec     As Boolean
Private mlTotalDirs       As Long
Private mlTotalFiles      As Long

' *****************
' Events
' *****************

Public Event BeginFindFiles()
Public Event EndFindFiles(FileCount As Long)
Public Event FoundFile(filename As String, Cancel As Boolean)

' *****************
' Public Properties
' *****************

Public Property Let Path(ByVal vData As String)
    msPath = vData
End Property


Public Property Get Path() As String
  Path = msPath
End Property


Public Property Let Filespec(ByVal vData As String)
  msFileSpec = vData
End Property


Public Property Get Filespec() As String
  Filespec = msFileSpec
End Property


Public Property Let Searching(ByVal vData As Boolean)
    mbSearching = vData
End Property


Public Property Get Searching() As Boolean
    Searching = mbSearching
End Property

' *********************
' Public Functions/Subs
' *********************

Public Function FindAll(FileListArray As Variant) As Boolean
  Dim asfiles() As String
  ReDim asfiles(0)
  On Error GoTo eop_error
  
  If FindFiles(asfiles()) Then
    FileListArray = asfiles()
    FindAll = True
  Else
    FindAll = False
  End If

eop_error:
  Select Case err.Number
    Case Is > 0
      FindAll = False
      MsgBox "Error #" & err.Number & " has occured in clsGetFiles.FindAll", , "ERROR!!"
  End Select
End Function

' ******************
' Private Properties
' ******************

Public Property Let recursive(ByVal vData As Boolean)
  'Future Property
  mbRecursive = vData
End Property


Public Property Get recursive() As Boolean
  'Future Property
  recursive = mbRecursive
End Property


Public Property Let RecursiveLevels(ByVal vData As Long)
  'Future Property
  mlRecursiveLevels = vData
End Property


Public Property Get RecursiveLevels() As Long
  'Future Property
  RecursiveLevels = mlRecursiveLevels
End Property

' **********************
' Private Functions/Subs
' **********************

Private Function FindFiles(asFoundFiles() As String) As Boolean
  ' If we're running & we got a click, it's because DoEvents in
  ' either the SearchDirs() or SearchFileSpec() proc let it happen.
  ' Tell the proc to stop. Once SearchDirs() has un-recursed itself
  ' we'll finish off below where we left off...
  Dim drvbitmask&, maxpwr%, pwr%
  FindFiles = False
  If Searching Then
    Searching = False
    GoTo eop_exit
  End If
  On Error Resume Next
  
  ' A parsing routine could be implemented here for
  ' multiple file spec searches, i.e. "*.bmp,*.wmf", etc.
  ' See the MS KB article Q130860 for information on how
  ' FindFirstFile() does not handle the "?" wildcard char correctly !!
  If Len(Filespec) = 0 Then GoTo eop_exit
  
  If Len(Path) = 0 Then GoTo eop_exit
  
  mbSearching = True
  mbUseFileSpec = True
  
  RaiseEvent BeginFindFiles
  
  Call SearchDirs(Path, asFoundFiles(), 0)
  
  Searching = False
  mbUseFileSpec = False
  
  mlTotalFiles = UBound(asFoundFiles)
  RaiseEvent EndFindFiles(mlTotalFiles)
  
  FindFiles = True
    
eop_exit:
    
End Function

' This is were it all happens...

' You can use the values in returned in the
' WIN32_FIND_DATA structure to virtually obtain any
' information you want for a particular folder or group of files.

' This recursive procedure is similar to the Dir$ function
' example found in the VB3 help file...


Private Sub SearchDirs(CurPath$, asFoundFiles() As String, Deepness As Integer) ' curpath$ is passed w/ trailing "\"
  ' These can't be static!!! They must be
  ' re-allocated on each recursive call.
  Dim dirs%, dirbuf$(), i%
  
  ' This proc to be cancelled by the user.
  ' It's not necessary to have this in the loop
  ' below since the loop works so fast...  DoEvents
  
  If Not Searching Then GoTo eop_exit
  If mbRecursive And ((mlRecursiveLevels >= Deepness) Or mlRecursiveLevels = -1) Then
  ' This loop finds *every* subdir and file in the current dir
  hItem& = FindFirstFile(CurPath$ & vbAllFiles, WFD)
  
  If hItem& <> INVALID_HANDLE_VALUE Then
    Do
      ' Tests for subdirs only...
      If (WFD.dwFileAttributes And vbDirectory) Then
          
        ' If not a  "." or ".." DOS subdir...
        If Asc(WFD.cFileName) <> vbKeyDot Then
          
        ' If not a  "." or ".." DOS subdir...
          ' This is executed in the mnuFindFiles_Click()
          ' call though it isn't used...
          mlTotalDirs = mlTotalDirs + 1
          ' This is the heart of a recursive proc...
          ' Cache the subdirs of the current dir in the 1 based array.
          ' This proc calls itself below for each subdir cached in the array.
          ' (re-allocating the array only once every 10 itinerations improves speed)
          If (dirs% Mod 10) = 0 Then ReDim Preserve dirbuf$(dirs% + 10)
          dirs% = dirs% + 1
          dirbuf$(dirs%) = Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1)
        End If
      
        ' File size and attribute tests can be used here, i.e:
        ' ElseIf (WFD.dwFileAttributes And vbHidden) = False Then  'etc...
        
        ' Get a total file count for mnuFolderInfo_Click()
      ElseIf Not mbUseFileSpec Then
        mlTotalFiles = mlTotalFiles + 1
      End If
    
    ' Get the next subdir or file
    Loop While FindNextFile(hItem&, WFD)
        
    ' Close the search handle
    Call FindClose(hItem&)
  
  End If
  End If
  ' When UseFileSpec% is set mnuFindFiles_Click(),
  ' SearchFileSpec() is called & each folder must be
  ' searched a second time.
  If mbUseFileSpec Then
    Call SearchFileSpec(CurPath$, asFoundFiles())
  End If

  ' Recursively call this proc & iterate through each subdir cached above.
  If mbRecursive And ((mlRecursiveLevels >= Deepness) Or mlRecursiveLevels = -1) Then
  For i% = 1 To dirs%
    SearchDirs CurPath$ & dirbuf$(i%) & vbBackslash, asFoundFiles(), Deepness + 1
  Next i%
  End If
eop_exit:

End Sub


Public Sub SearchFileSpec(CurPath$, asFoundFiles() As String, Optional ByVal strfilespec As String, _
                          Optional fileattrib As VbFileAttribute = -1, Optional noworry As Boolean = False, Optional IncDir As Boolean = False)
  'curpath$ is passed w/ trailing "\"
  ' This procedure *only*  finds files in the
  ' current folder that match the FileSpec$
  Dim Cancel As Boolean, fcnt As Long, orgcnt As Long
  Dim sTempFile As String
  On Error GoTo 0
  If strfilespec = "" Then
    'defaults to this
    strfilespec = Filespec
  End If
  hFile& = FindFirstFile(CurPath$ & strfilespec, WFD)
  If hFile& <> INVALID_HANDLE_VALUE Then
    orgcnt = UBound(asFoundFiles)
    ReDim Preserve asFoundFiles(UBound(asFoundFiles) + 10)
    Do
      ' Use DoEvents here since we're loading a ListBox and
      ' there could be hundreds of files matching the FileSpec$
      If mbSearching = False And noworry = False Then GoTo eop_exit
      If WFD.dwFileAttributes = vbDirectory Then
        If IncDir Then
                fcnt = fcnt + 1
                If fcnt Mod 10 Then ReDim Preserve asFoundFiles(UBound(asFoundFiles) + 10)
                sTempFile = CurPath$ & Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1)
                asFoundFiles(fcnt + orgcnt) = sTempFile & "\"
        End If
      Else
            If fileattrib = -1 Then
                fcnt = fcnt + 1
                If fcnt Mod 10 Then ReDim Preserve asFoundFiles(UBound(asFoundFiles) + 10)
                sTempFile = CurPath$ & Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1)
                asFoundFiles(fcnt + orgcnt) = sTempFile
            ElseIf fileattrib And WFD.dwFileAttributes Then
                fcnt = fcnt + 1
                If fcnt Mod 10 Then ReDim Preserve asFoundFiles(UBound(asFoundFiles) + 10)
                sTempFile = CurPath$ & Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1)
                asFoundFiles(fcnt + orgcnt) = sTempFile
            End If
            Cancel = False
            RaiseEvent FoundFile(sTempFile, Cancel)
            mbSearching = Not Cancel
      End If
      ' Get the next file matching the FileSpec$
    Loop While FindNextFile(hFile&, WFD)
    ' Close the search handle
    ReDim Preserve asFoundFiles(orgcnt + fcnt) 'ends the trailing empty strings
  Call FindClose(hFile&)
  End If
Exit Sub
eop_exit:
Call FindClose(hFile&)
End Sub

Private Sub Class_Initialize()
mbRecursive = True
mlRecursiveLevels = -1
End Sub
