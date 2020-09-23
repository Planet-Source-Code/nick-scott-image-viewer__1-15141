VERSION 5.00
Begin VB.Form frm_preferences 
   Caption         =   "Preferences"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4935
   Icon            =   "frm_preferences.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   4935
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fra_restrict_filetypes 
      Caption         =   "Restrict File Types To"
      Height          =   1455
      Left            =   2640
      TabIndex        =   4
      Top             =   120
      Width           =   2175
      Begin VB.CheckBox chk_allow_jpg 
         Caption         =   "JPG/Jpeg Files"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CheckBox chk_allow_gif 
         Caption         =   "GIF Files"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1455
      End
      Begin VB.CheckBox chk_allow_bmp 
         Caption         =   "Bitmap Files"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmd_cancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmd_save 
      Caption         =   "Save and Exit"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   1455
   End
   Begin VB.DirListBox dir_setdirectory 
      Height          =   3015
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
   Begin VB.DriveListBox drv_setdirectory 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
   Begin VB.Frame fra_defaultgfxdir 
      Caption         =   "Default Graphics Directory"
      Height          =   3735
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frm_preferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_cancel_Click()

Unload Me

End Sub

Private Sub cmd_save_Click()

FilePattern = ""

If chk_allow_bmp.Value = 1 Then
   FilePattern = FilePattern + "*.bmp"
End If

If chk_allow_gif.Value = 1 And FilePattern > "" Then
   FilePattern = FilePattern + ";*.gif"
Else
   If chk_allow_gif.Value = 1 And FilePattern = "" Then
      FilePattern = FilePattern + "*.gif"
   End If
End If

If chk_allow_jpg.Value = 1 And FilePattern > "" Then
   FilePattern = FilePattern + ";*.jpg;*.jpeg"
Else
   If chk_allow_jpg.Value = 1 And FilePattern = "" Then
      FilePattern = FilePattern + "*.jpg;*.jpeg"
   End If
End If

If FilePattern = "" Then
   FilePattern = "*.*"
End If

SaveSetting "GfxOrg", "Preferences", "Graphics Drive", drv_setdirectory.Drive
SaveSetting "GfxOrg", "Preferences", "Graphics Directory", dir_setdirectory.Path
SaveSetting "GfxOrg", "Preferences", "File Pattern", FilePattern
SaveSetting "GfxOrg", "Preferences", "Allow *.bmp", chk_allow_bmp.Value
SaveSetting "GfxOrg", "Preferences", "Allow *.gif", chk_allow_gif.Value
SaveSetting "GfxOrg", "Preferences", "Allow *.jpg", chk_allow_jpg.Value

frm_mainmenu.drv_graphics.Drive = drv_setdirectory.Drive
frm_mainmenu.dir_graphics.Path = dir_setdirectory.Path
frm_mainmenu.fil_graphics.Path = dir_setdirectory.Path
frm_mainmenu.fil_graphics.Pattern = FilePattern
frm_mainmenu.txt_filepattern.Text = FilePattern

Unload Me

End Sub

Private Sub drv_setdirectory_Change()

dir_setdirectory.Path = drv_setdirectory.Drive

End Sub

Private Sub Form_Load()

GfxDrvRegSet = GetSetting("GfxOrg", "Preferences", "Graphics Drive", "C:")
GfxDirRegSet = GetSetting("GfxOrg", "Preferences", "Graphics Directory", "C:\")
TypeBmpRegSet = GetSetting("GfxOrg", "Preferences", "Allow *.bmp", "1")
TypeGifRegSet = GetSetting("GfxOrg", "Preferences", "Allow *.gif", "1")
TypeJpgRegSet = GetSetting("GfxOrg", "Preferences", "Allow *.jpg", "1")

drv_setdirectory.Drive = GfxDrvRegSet
dir_setdirectory.Path = GfxDirRegSet

chk_allow_bmp.Value = TypeBmpRegSet
chk_allow_gif.Value = TypeGifRegSet
chk_allow_jpg.Value = TypeJpgRegSet

End Sub
