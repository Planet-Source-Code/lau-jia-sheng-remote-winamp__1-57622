Attribute VB_Name = "winampModule"
'*************************************************************
'******************* Winamp Control Module  ******************
'********************         by           *******************
'*********************   James Crasta     ********************
'**********************                  *********************
'***********************                **********************
'************************              ***********************
'***********************  Version 1.3   **********************
'**********************     Released     *********************
'*********************     06/19/2001     ********************
'********************                      *******************
'*************************************************************
' This module was created by me when i couldn't find any
' winamp module for VB that would do what I wanted so
' i made one using the winamp button code numbers found
' in the nsdn section of Winamp's site. The code is
' pretty self-explanatory, i have included comments to
' help you understand this.
'**************************************************************
' Any changes or updates may be made to this module and it
' may be used in any program you make, whether for profit or
' not, but if you are going to distribute the source please
' make sure that credit is made to me somewhere in the module,
' preferably the top.

Option Explicit
Public Const WM_COMMAND = 273
Public Const WM_USER = 1024
Public Const WM_WA_IPC = &H400
Public Const WM_COPYDATA = &H4A
Public Const DefWinampPath = ":\Program Files\Winamp\winamp.exe"


Public vWinampID As Long
Public winamppath As String
Public LastWinampCaption As String
Public LastTitle As String


'**************************
'****** API DECLARES ******
'**************************
' Find Winamp Window
Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As Long) As Long

' SendMessage To Window (waits for reply)
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal WndID As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

' PostMessage To Window (returns true/false)
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

' Gets Window Caption
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long

' Sends text messages using COPYDATASTRUCT
Public Declare Function SendMessageCDS Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As COPYDATASTRUCT) As Long

' Copies a string to its Address, used for SendMessageCDS
Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

' Gets Menu items , used for Shuffle / repeat status
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function GetMenuState Lib "user32" (ByVal hmenu As Long, ByVal wID As Long, ByVal wFlags As Long) As Long

'Drive type to scan just local HardDisks
'Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
'Public Const DRIVE_FIXED = 3



'Private Const MF_CHECKED As Long = 8

'**************************
'****** DEFINED TYPES *****
'**************************

Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As Long
End Type
'=============================


'###############################################
'######## POSTMESSAGE COMMAND CONSTANTS ########
'###############################################
'These are the constants for the functions
'***activate using WM_DO(constant_name)
'***i.e:  WM_DO(PREV_TRACK) <--will switch
'                              to previous track

'Constant declare                                       Function
'-----------------------------------------------------------------------------------------------------------------------------
Public Const PREV_TRACK As Long = 40044                 'Goes to previous track
Public Const NEXT_TRACK As Long = 40048                 'Goes to next track
Public Const PLAY_TRACK As Long = 40045                 'Begins playing
Public Const PAUSE_TRACK As Long = 40046                'Pauses playing
Public Const STOP_TRACK As Long = 40047                 'Stops winamp
Public Const FADEOUT_STOP As Long = 40147               'fades out and stops
Public Const STOP_AFTER_CURRENT As Long = 40157         'stops after current track (Problems have been reported with this function)
Public Const FORWARD_5_SEC As Long = 40148              'jumps forward in the song 5 seconds
Public Const BACK_5_SEC As Long = 40144                 'jumps backward in the song 5 seconds
Public Const GO_BEGINNING_PLAYLIST As Long = 40154      'jumps to the first song in the playlist
Public Const GO_END_PLAYLIST As Long = 40158            'jumps to the last song in the playlist
Public Const OPEN_FILE_DIALOG As Long = 40029           'displays the open file dialog
Public Const OPEN_URL_DIALOG As Long = 40155            'displays the Open URL dialog
Public Const OPEN_FILE_INFO_DIALOG As Long = 40188      'displays the file info dialog
Public Const DISP_ELAPSED_TIME As Long = 40037          'changes the time display to show elapsed time
Public Const DISP_REMAINING_TIME As Long = 40038        'changes the time display to show remaining time
Public Const PREFS_DIALOG As Long = 40012               'Brings up the preferences dialog
Public Const VIS_OPTIONS As Long = 40190                'Visualization options
Public Const VIS_PLUGIN_OPTIONS As Long = 40191         'Visualization plugin options
Public Const START_VIS = 40192                          'Begin Visualization
Public Const SHOW_ABOUT As Long = 40041                 'Shows the winamp about dialog box
Public Const TOGGLE_TITLE_AUTOSCROLL As Long = 40189    'Turns on/off the automatic scrolling
Public Const TOGGLE_ALWAYS_ON_TOP As Long = 40019       'Turns on/off Always on top
Public Const TOGGLE_WINDOWSHADE As Long = 40064         'Turns on/off windowshade
Public Const TOGGLE_PLAYLIST_WINDOWSHADE As Long = 40266 'Turns on/off windowshade in the playlist
Public Const TOGGLE_DOUBLESIZE As Long = 40165          'Turns on/off double size mode
Public Const TOGGLE_EQ As Long = 40036                  'Displays/hides the Equalizer
Public Const TOGGLE_PLAYLIST As Long = 40040            'Displays/hides the playlist
Public Const TOGGLE_WINAMP_VISIBLE As Long = 40258      'Displays/hides the main Winamp window
Public Const TOGGLE_MINIBROWSER As Long = 40298         'Displays/hides the minibrowser window
Public Const TOGGLE_EASYMOVE As Long = 40186            'Turns on/off the Easy Move feature
Public Const VOLUME_RAISE As Long = 40058               'Raises the volume a notch ***** use the wm_setvolume function for more precise volume control
Public Const VOLUME_LOWER As Long = 40059               'lowers the volume a notch ***** use the wm_setvolume function for more precise volume control
Public Const TOGGLE_REPEAT As Long = 40022              'Turns on/off repeat
Public Const TOGGLE_SHUFFLE As Long = 40023             'Turns on/off shuffle
Public Const JUMPTO_TIME_DIALOG As Long = 40193         'Shows the jump to time dialog
Public Const JUMPTO_FILE_DIALOG As Long = 40194         'Shows the jump to file dialog
Public Const OPEN_SKIN_SLECTOR As Long = 40219          'Shows the skin selection window
Public Const CONFIG_CURRENT_VIS As Long = 40221         'Configure current vis plugin
Public Const RELOAD_CURRENT_SKIN As Long = 40291            'Reload the current skin
Public Const WINAMP_EXIT As Long = 40001                'Exits Winamp


'###############################################
'######## SENDMESSAGE COMMAND CONSTANTS ########
'###############################################
'use variable = WM_GET(constant_name)
'example:  length = WM_GET(CLEAR_PLAYLIST)
Public Const WINAMP_VERSION As Byte = 0         'Retrieves the version of Winamp running. Version will be 0x20yx for 2.yx. This is a good way to determine if you did in fact find the right window, etc.
Public Const CLEAR_PLAYLIST As Byte = 1         'Clears Winamp 's internal playlist.
Public Const PLAYBACK_STATUS As Byte = 2        'Returns the status of playback. If 'ret' is 1, Winamp is playing. If 'ret' is 3, Winamp is paused. Otherwise, playback is stopped.
Public Const SONG_LENGTH As Byte = 3            'Returns the length of the song in seconds
Public Const SONG_POSITION As Byte = 4          'Returns the current position in the song, in milliseconds
Public Const SEEK_CURRENT_TRACK As Byte = 5     'Seeks within the current track. The offset is specified in 'data', in milliseconds.
Public Const WRITE_CURR_PLAYLIST As Byte = 6    'Writes out the current playlist to Winampdir\winamp.m3u, and returns the current position in the playlist.
Public Const SET_PLAYLIST_POS As Byte = 7       'Sets the playlist position to the position specified in tracks in 'data'.
Public Const PLAYLIST_LENGTH As Byte = 8        'Returns length of the current playlist, in tracks.
Public Const PLAYLIST_POSITION As Byte = 9      'Returns the position in the current playlist, in tracks (requires Winamp 2.05+).
Public Const SONG_SAMPLERATE  As Byte = 10      'Returns the sample rate of the song
Public Const SONG_BITRATE As Byte = 11          'Returns the bitrate of the song
Public Const SONG_NUMCHANNELS As Byte = 12      'Returns the number of channels in the song.
Public Const SONG_TITLE As Byte = 13            'Returns the title of the song.
Public AutoLoadWinAmp As Boolean

Public Property Get WinampID() As Long
FindWinamp
WinampID = vWinampID
End Property

Public Function FindWinampPath() As String
Dim i As Integer
If winamppath = "" Then
        For i = 67 To 90
            If Dir(Chr(i) & DefWinampPath) <> "" Then
                winamppath = Chr(i) & DefWinampPath
                FindWinampPath = winamppath
                Exit Function
            End If
        Next i
ElseIf Dir(winamppath) = "" Then
        For i = 67 To 90
            If Dir(Chr(i) & DefWinampPath) <> "" Then
                winamppath = Chr(i) & DefWinampPath
                FindWinampPath = winamppath
                Exit Function
            End If
        Next i
Else
    FindWinampPath = winamppath
End If
End Function
Public Function FindWinamp()
On Error GoTo err
    Dim i As Integer
    vWinampID = FindWindowA("Winamp v1.x", 0)
    If vWinampID = 0 And winamppath <> "" And Dir(winamppath) <> "" And AutoLoadWinAmp = True Then
        Dim deluseless As Double
        deluseless = Shell(winamppath)
        Dim cntr As Integer
        Do
            cntr = cntr + 1
            vWinampID = FindWindowA("Winamp v1.x", 0)
            DoEvents: DoEvents: DoEvents: DoEvents
            If cntr = 60000 Then
            MsgBox "Timeout!"
            Exit Function
            End If
        Loop Until WinampID <> 0
    ElseIf vWinampID = 0 And AutoLoadWinAmp Then
        For i = 67 To 90
            If Dir(Chr(i) & DefWinampPath) <> "" Then
                winamppath = Chr(i) & DefWinampPath
                FindWinamp
                Exit Function
            End If
        Next i
        'MsgBox "The winamp path provided is incorrect.  Please specify a correct path"
    End If
        Exit Function
err:
        Call WinampModErrorHandler
End Function

' These are the standalone single-function functions.
' They are pretty self-explanatory

'Sets the volume (Volume must be between 0 - 255)
Public Function WM_SetVolume(Volume As Byte) As Long
    On Error GoTo err

    WM_SetVolume = SendMessage(WinampID, WM_WA_IPC, Volume, 122)
    Exit Function
err:
    Call WinampModErrorHandler
End Function


'Jumps in the playlist to the position specified in tracks
Public Function WM_SetPlaylistPos(PosInPlaylist As Integer) As Long
    On Error GoTo err

    WM_SetPlaylistPos = SendMessage(WinampID, WM_WA_IPC, PosInPlaylist, 121)
    Exit Function
err:
    Call WinampModErrorHandler
End Function
'Sets the panning to a specified number, which can be between 0 (all left) and 255 (all right).
Public Function WM_SetPanning(Panning As Byte)
    On Error GoTo err
    If Panning >= 128 Then
        Panning = Panning - 128
    Else
        Panning = Panning + 128
    End If
    WM_SetPanning = SendMessage(WinampID, WM_WA_IPC, Panning, 123)
    Exit Function
err:
    Call WinampModErrorHandler
End Function
Public Function WM_SetSongPosition(Optional Seconds As Long, Optional Ms As Long)
'Sets the current position in the song
'Returns:
'0 if success
'1 if eof
'-1 if not playing
    WM_SetSongPosition = SendMessage(WinampID, WM_WA_IPC, (Seconds * 1000 + Ms), 106)
End Function

Public Function WM_AddPlaylist(filename As String)
'adds specified file/directory  to playlist.
    '=============================
    If Right(filename, 1) = "\" Then
        filename = Mid(filename, 1, Len(filename) - 1)
    End If
   
    Dim cds As COPYDATASTRUCT
    cds.dwData = 100
    cds.lpData = lstrcpy(filename, filename)
    cds.cbData = Len(filename) + 1

    SendMessageCDS WinampID, &H4A, 0&, cds
    '=============================
End Function


Public Function WM_GetShufflestatus() As Integer
'Gets the status of the shuffle button
'Returns:
'0 if off
'1 if on
'-1 if error
    On Error GoTo err
    Dim hmenu As Long, q As Long
    hmenu = GetSystemMenu(WinampID, 0)
    If GetMenuState(hmenu, TOGGLE_SHUFFLE, 0) = 8 Then
        WM_GetShufflestatus = 1
    Else
        WM_GetShufflestatus = 0
    End If
    Exit Function
err:
    WinampModErrorHandler
    'WM_GetShufflestatus = -1
End Function
Public Function WM_GetRepeatStatus() As Integer
'Gets the status of the repeat button
'Returns:
'0 if off
'1 if on
'-1 if error
    Dim hmenu As Long, q As Long
  hmenu = GetSystemMenu(WinampID, 0)
   If GetMenuState(hmenu, TOGGLE_REPEAT, 0) = 8 Then
        WM_GetRepeatStatus = 1
   Else
        WM_GetRepeatStatus = 0
   End If
End Function





' Handles errors and outputs them to the debug screen.
Private Sub WinampModErrorHandler()
    Debug.Print "There has been an error: "; err.Description; " and number "; err.Number
End Sub


Public Function WM_DO(cmnd As Long)
    On Error GoTo err
    Dim tmp As String
    Dim ret As Integer
    Dim isplay As Integer
    'isplay = SendMessage(WinampID, WM_USER, 0, 104)
    ret = PostMessage(WinampID, WM_COMMAND, cmnd, 0)
    Exit Function
err:
    Call WinampModErrorHandler
End Function

Public Function WM_GET(cmnd As Byte, Optional data As Long) As Variant
    On Error GoTo err
    Dim tmp As String
    Dim ret As Integer
    Dim isplay As Integer
    isplay = SendMessage(WinampID, WM_USER, 0, 104)
    Select Case cmnd
        Case WINAMP_VERSION ' returns current winamp version
            WM_GET = SendMessage(WinampID, WM_WA_IPC, 0, 0) ' returns winamp version in LONG format
        Case CLEAR_PLAYLIST ' clears all items in the playlist
            WM_GET = SendMessage(WinampID, WM_WA_IPC, 0, 101) ' does not return anything
        Case PLAYBACK_STATUS 'gets current playback status
            WM_GET = SendMessage(WinampID, WM_WA_IPC, 0, 104) ' returns 1 for playing, 3 for paused, 0 for stopped
        Case SONG_LENGTH
            WM_GET = SendMessage(WinampID, WM_WA_IPC, 1, 105) ' returns track length in seconds
        Case SONG_POSITION
            WM_GET = SendMessage(WinampID, WM_WA_IPC, 0, 105) ' returns position in the current track in milliseconds
        Case PLAYLIST_LENGTH
            WM_GET = SendMessage(WinampID, WM_WA_IPC, 0, 124) ' returns number of songs in playlist
        Case PLAYLIST_POSITION
            WM_GET = SendMessage(WinampID, WM_WA_IPC, 0, 125) ' returns the currently playing track number
        Case SONG_SAMPLERATE
            WM_GET = SendMessage(WinampID, WM_WA_IPC, 0, 126) ' returns the currently playing song's sample rate
        Case SONG_BITRATE
            WM_GET = SendMessage(WinampID, WM_WA_IPC, 1, 126) ' returns the currently playing song's bit rate
        Case SONG_NUMCHANNELS
            WM_GET = SendMessage(WinampID, WM_WA_IPC, 3, 126) ' returns the currently playing song's sample rate
        Case WRITE_CURR_PLAYLIST  'Writes out the current playlist to Winampdir\winamp.m3u, and returns the current position in the playlist.
                WM_GET = SendMessage(WinampID, WM_WA_IPC, 3, 120) ' writes the winamp.m3u file
        Case SONG_TITLE ' returns a string with the song title in it
            ' Since I couldnt find any API that
            ' would query winamp, i decided to
            ' read in the caption and trim it down
            ' to just the title.
            Dim strBuffer As String, lngtextlen As Long
            Let lngtextlen& = GetWindowTextLength(WinampID) 'gets the length of the caption
            Let strBuffer$ = String$(lngtextlen&, 0&) 'i dont know why this is necessary, i found it in someone else's API code
            Call GetWindowText(WinampID, strBuffer$, lngtextlen& + 1&) ' reads in the caption text
            If strBuffer$ = LastWinampCaption Then
                WM_GET = LastTitle
            Else
                LastWinampCaption = strBuffer
                If LCase(strBuffer$) Like "*[paused]" = False Then ' queries if the [Paused] string is there and removes it
                    strBuffer$ = Left(strBuffer$, Len(strBuffer) - 9)
                End If
                strBuffer$ = Mid(strBuffer$, 1, Len(strBuffer) - 8) ' removes the -Winamp
                Dim findDot As Integer
                findDot = InStr(1, strBuffer, ".") ' finds the dot in the number at the beginning
                LastTitle = Trim(Mid(strBuffer$, findDot + 1)) 'Returns the final title value
                WM_GET = LastTitle
            End If
    End Select
    Exit Function
err:
    Call WinampModErrorHandler
End Function

