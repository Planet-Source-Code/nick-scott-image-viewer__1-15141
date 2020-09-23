VERSION 5.00
Object = "{B8325759-F2AB-11D2-B1E6-9246AA68EB78}#2.0#0"; "ThumbBrowseP.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "fm20.dll"
Begin VB.Form Main 
   Caption         =   "Webblasters Image Viewer"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frm_mainmenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   7935
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "4:04 Viper1"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   17965
            TextSave        =   "2/10/00"
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picLogo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   8295
      Left            =   0
      ScaleHeight     =   8295
      ScaleWidth      =   375
      TabIndex        =   6
      Top             =   -240
      Width           =   375
   End
   Begin VB.ComboBox Select1 
      Height          =   315
      Left            =   840
      TabIndex        =   5
      Text            =   "*.bmp"
      Top             =   7200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin ThumbBrowseControl.ThumbBrowse tb 
      Height          =   7800
      Left            =   9960
      TabIndex        =   0
      Top             =   0
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   13758
      ThumbWidth      =   70
      ThumbBorder     =   5
      ColorLight      =   -2147483648
      ColorDark       =   -2147483648
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   720
      Top             =   6480
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8175
      Left            =   480
      ScaleHeight     =   8175
      ScaleWidth      =   2055
      TabIndex        =   1
      Top             =   0
      Width           =   2055
      Begin VB.FileListBox File1 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   4575
         Left            =   0
         Pattern         =   "*.JPG;*.BMP;*.GIF;*.DIB"
         TabIndex        =   4
         Top             =   3240
         Width           =   1965
      End
      Begin VB.DirListBox Dir1 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2790
         Left            =   0
         TabIndex        =   3
         Top             =   360
         Width           =   1965
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1935
      End
   End
   Begin MSForms.Image Image1 
      Height          =   7935
      Left            =   2520
      Top             =   0
      Width           =   7335
      BackColor       =   -2147483644
      BorderStyle     =   0
      SpecialEffect   =   1
      Size            =   "12938;13996"
      Picture         =   "frm_mainmenu.frx":030A
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      WindowList      =   -1  'True
      Begin VB.Menu mnuCut 
         Caption         =   "&Cut"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
         Shortcut        =   ^D
      End
      Begin VB.Menu ByBy 
         Caption         =   "&Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu Program 
      Caption         =   "&Add-Ins"
      Begin VB.Menu mnuSlide 
         Caption         =   "SlideShow"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuOrder 
         Caption         =   "ReOrder"
         Shortcut        =   ^O
      End
   End
   Begin VB.Menu mnuBMP 
      Caption         =   "&BMP"
   End
   Begin VB.Menu mnuGif 
      Caption         =   "&GIF"
   End
   Begin VB.Menu mnuJPG 
      Caption         =   "&JPG"
   End
   Begin VB.Menu mnuPlayer 
      Caption         =   "&Media Player"
   End
   Begin VB.Menu mnuThumbs 
      Caption         =   "&LoadThumbs"
   End
   Begin VB.Menu mnuClear 
      Caption         =   "&Clear Image"
   End
   Begin VB.Menu mnuSlides 
      Caption         =   "&SlideShow"
      Begin VB.Menu mnuStart 
         Caption         =   "&Start"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuStop 
         Caption         =   "&Stop"
         Shortcut        =   ^Y
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Index           =   0
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuCredits 
         Caption         =   "&Contents"
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu mnuPref 
      Caption         =   "Preferences"
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CRLF        As String
Dim CRLF_CRLF   As String
Dim iPic        As Byte
Dim curSelect   As StdPicture
Dim cL          As New cLogo

Public i As Integer
Private Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As Long '  only used if FOF_SIMPLEPROGRESS
End Type
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Const FO_DELETE = &H3
Private Const FOF_ALLOWUNDO = &H40

Private Sub Form_Resize()
    
    On Error Resume Next
    picLogo.Height = Me.ScaleHeight
    On Error GoTo 0
    cL.Draw
    
End Sub

Private Sub Command1_Click()
i = 0
' to disable drive, dir , file and combo box during slide show
Drive1.Enabled = False
Dir1.Enabled = False
File1.Enabled = False
Select1.Enabled = False
Timer2.Enabled = True
End Sub

Private Sub Command2_Click()

  Timer2.Enabled = False
       ' to enable drive, dir , file and combo box after slide show
  Drive1.Enabled = True
  Dir1.Enabled = True
  File1.Enabled = True
  Select1.Enabled = True
End Sub

Private Sub mnuCut_Click()
        Clipboard.Clear
        Clipboard.SetData Image1.Picture
        Image1.Picture = Nothing
End Sub

Private Sub mnuPaste_Click()
   Image1.Picture = Clipboard.GetData
End Sub

Private Sub mnuPref_Click()
frm_preferences.Show
End Sub

Private Sub mnuPrint_Click()

On Error GoTo PrintErr

Main.PrintForm

PrintErr:
    If Err.Number = 32755 Then
    Exit Sub
    
      
    End If
End Sub

Private Sub mnuStart_Click()
i = 0
' to disable drive, dir , file and combo box during slide show
Drive1.Enabled = False
Dir1.Enabled = False
File1.Enabled = False
Select1.Enabled = False
Timer2.Enabled = True
End Sub

Private Sub mnuStop_Click()

  Timer2.Enabled = False
       ' to enable drive, dir , file and combo box after slide show
  Drive1.Enabled = True
  Dir1.Enabled = True
  File1.Enabled = True
  Select1.Enabled = True
End Sub

Private Sub select1_Click()
   File1.Pattern = Trim(Select1.Text)
End Sub

Private Sub Form_Load()
    CRLF = vbCrLf
    CRLF_CRLF = CRLF & "." & CRLF
    iPic = 101
    cL.DrawingObject = picLogo
    cL.Caption = "                       Webblasters Software Inc./  ® Nick Scott"
    frmMedia.Hide
    'frmAbout.Hide
    frmBrowser.Hide
    FileRenamer.Hide
    Dir1.Path = App.Path
    'Image1.PictureSizeMode = fmPictureSizeModeStretch
    Image1.Picture = LoadPicture(App.Path & "\images\" & "Lab.jpg")
End Sub

Private Sub tmr_Timer()

    If iPic = 124 Then iPic = 101
    Set curSelect = LoadResPicture(iPic, vbResBitmap)
    iPic = iPic + 1
    
End Sub

Private Sub ByBy_Click()
Unload Me
End
End Sub

Private Sub cmdExit_Click()
Unload Me
End
End Sub



Private Sub mnuBMP_Click()
    File1.Pattern = "*.bmp"
    File1.Refresh
End Sub


Private Sub mnuCredits_Click()
Load frmBrowser
frmBrowser.WebBrowser1.Navigate App.Path & "\Help.htm"
frmBrowser.Show 'End Sub
End Sub

Private Sub mnuDelete_Click()
Dim nxtFile

Dim op As SHFILEOPSTRUCT

    With op
        .wFunc = FO_DELETE
        .pFrom = (Dir1.Path + "\" + File1.FileName)
        .fFlags = FOF_ALLOWUNDO
    End With
    SHFileOperation op
    File1.Refresh
End Sub

Private Sub mnuJPG_Click()
    File1.Pattern = "*.jpg"
    File1.Refresh
End Sub

Private Sub mnuGIF_Click()
    File1.Pattern = "*.gif"
    File1.Refresh
End Sub
Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub
Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub
Private Sub mnuClear_Click()
    'Image1.PictureSizeMode = fmPictureSizeModeStretch
    Image1.Picture = LoadPicture(App.Path & "\images\" & "Lab.jpg")
       
End Sub
Private Sub mnuOrder_Click()
    Main.Hide
    FileRenamer.Show
End Sub
Private Sub mnuPlayer_Click()
    Main.Hide
    frmMedia.Show
End Sub

   
Private Sub mnuThumbs_Click()
Dim lsPath As String
        
    tb.Cls
    For i = 1 To File1.ListCount - 1
        lsPath = File1.Path & "\" & File1.List(i)
        tb.AddThumb lsPath, File1.List(i), FileLen(lsPath), FileDateTime(lsPath)
        DoEvents
    Next
End Sub
Private Sub File1_Click()
    
    Dim msg As String   ' Declare variables.
    Dim FileName
    FileName = File1.FileName
    Image1.Picture = LoadPicture()
    Image1.Picture = LoadPicture(Dir1.Path + "\" + FileName)
        
        If Image1.Picture.Height > Image1.Height Then
            Image1.PictureSizeMode = fmPictureSizeModeZoom
        Else
            Image1.PictureSizeMode = fmPictureSizeModeClip
        End If
        
        If Image1.Picture.Width > Image1.Width Then
            Image1.PictureSizeMode = fmPictureSizeModeZoom
        Else
            Image1.PictureSizeMode = fmPictureSizeModeClip
        End If
        
        
        If Image1.Picture = LoadPicture() Then
        msg = "Could not find the requested file."
        MsgBox msg  ' Display error message.
    End If
End Sub


Private Sub tb_ThumbClick(Index As Long, ThumbPath As String, ThumbCaption As String, ThumbSize As Long, ThumbDate As Date, Width As Long, Height As Long, Planes As Long, Colors As Long)
   
    Set Image1.Picture = LoadPicture(ThumbPath)
    'Dim msg As String   ' Declare variables.
    'Dim FileName                                 'for file list click
    'FileName = File1.FileName
    'Image1.Picture = LoadPicture()
    'Image1.Picture = LoadPicture(Dir1.Path + "\" + FileName)
        
        If Image1.Picture.Height > Image1.Height Then
            Image1.PictureSizeMode = fmPictureSizeModeZoom
        Else
            Image1.PictureSizeMode = fmPictureSizeModeClip
        End If
        
        If Image1.Picture.Width > Image1.Width Then
            Image1.PictureSizeMode = fmPictureSizeModeZoom
        Else
            Image1.PictureSizeMode = fmPictureSizeModeClip
        End If
        
        
    
End Sub

Private Sub Timer2_Timer()
Image1.Picture = LoadPicture(File1.Path & "\" & File1.List(i))
If Image1.Picture.Height > Image1.Height Then
            Image1.PictureSizeMode = fmPictureSizeModeZoom
        Else
            Image1.PictureSizeMode = fmPictureSizeModeClip
        End If
        
        If Image1.Picture.Width > Image1.Width Then
            Image1.PictureSizeMode = fmPictureSizeModeZoom
        Else
            Image1.PictureSizeMode = fmPictureSizeModeClip
        End If
        
        
i = i + 1
If i = File1.ListCount Then
Timer2.Enabled = False
End If
End Sub
