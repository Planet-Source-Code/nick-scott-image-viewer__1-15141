VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form frmMedia 
   Caption         =   "Media Player"
   ClientHeight    =   6165
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8100
   Icon            =   "Frmmedia.frx":0000
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   8100
   StartUpPosition =   2  'CenterScreen
   Begin MediaPlayerCtl.MediaPlayer mPlayer 
      Height          =   6495
      Left            =   2805
      TabIndex        =   3
      Top             =   -480
      Width           =   5160
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   -1  'True
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   0
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   -1  'True
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   -1  'True
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   -1  'True
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -150
      WindowlessVideo =   0   'False
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H80000000&
      ForeColor       =   &H00000000&
      Height          =   2820
      Left            =   0
      Pattern         =   "*.wav;*.mid;*.avi;*.mpg;*.mpeg;*.mov"
      TabIndex        =   2
      Top             =   3370
      Width           =   2775
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H80000000&
      ForeColor       =   &H00000000&
      Height          =   3015
      Left            =   0
      TabIndex        =   1
      Top             =   340
      Width           =   2775
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2775
   End
   Begin VB.Menu mnuExit 
      Caption         =   "File"
      Begin VB.Menu mnuViewer 
         Caption         =   "&Image Viewer"
      End
   End
End
Attribute VB_Name = "frmMedia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    Drive1.Drive = "C:\"
    Dir1.Path = App.Path
    'Dir1.SetFocus
    
End Sub
Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive    ' Set directory path.
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path  ' Set file path.
End Sub

Private Sub File1_PathChange()
    Dir1.Path = File1.Path  ' Set Dir1 path.
End Sub

Private Sub File1_Click()
    Dim msg As String   ' Declare variables.
    Dim FileName
    FileName = File1.FileName
    mPlayer.FileName = Dir1.Path + "\" + File1.FileName
            
        frmMedia.Caption = "My Picture Viewer" + "   " + "Now Playing " + File1.FileName
    If Err Then
        frmMedia.Enabled = False
        msg = "Could not find the requested file."
        MsgBox msg  ' Display error message.

    End If
End Sub



Private Sub mnuViewer_Click()
    frmMedia.Hide
    Main.Show
    mPlayer.SelectionEnd = True
    mPlayer.FileName = ""
    Drive1.Refresh
    Dir1.Refresh
    File1.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Main
    Dim i As Integer

    'close all sub forms
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
    
End Sub
