VERSION 5.00
Begin VB.Form FileRenamer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ReOrder"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7320
   Icon            =   "ReOrder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   372
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   488
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Select all      "
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2880
      TabIndex        =   16
      Top             =   3480
      Width           =   1230
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Select none "
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   17
      Top             =   3120
      Width           =   1230
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H8000000B&
      Caption         =   "ReOrder !"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   20
      Top             =   4800
      Width           =   1410
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Preview"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   480
      Picture         =   "ReOrder.frx":030A
      TabIndex        =   19
      Top             =   2160
      Width           =   1245
   End
   Begin VB.ListBox List1 
      BackColor       =   &H80000000&
      Height          =   2595
      Left            =   120
      TabIndex        =   18
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "File Operations"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1815
      Left            =   45
      TabIndex        =   6
      Top             =   0
      Width           =   4650
      Begin VB.CheckBox Check1 
         Caption         =   "Adapt counter after renaming"
         Height          =   240
         Left            =   1620
         TabIndex        =   22
         Top             =   1125
         Value           =   1  'Checked
         Width           =   2670
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   270
         Left            =   3645
         Max             =   100
         Min             =   1
         TabIndex        =   15
         Top             =   1440
         Value           =   1
         Width           =   420
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   135
         TabIndex        =   12
         Text            =   "1"
         Top             =   1440
         Width           =   1275
      End
      Begin VB.OptionButton Option1 
         Caption         =   "After"
         Height          =   240
         Index           =   2
         Left            =   2385
         TabIndex        =   11
         Top             =   675
         Width           =   1200
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Before"
         Height          =   240
         Index           =   1
         Left            =   1260
         TabIndex        =   10
         Top             =   675
         Width           =   1200
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000B&
         Caption         =   "Replace"
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   9
         Top             =   675
         Value           =   -1  'True
         Width           =   1200
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1665
         TabIndex        =   8
         Text            =   "Text2"
         Top             =   315
         Width           =   2760
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Begin counter"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   135
         TabIndex        =   21
         Top             =   1125
         Width           =   1275
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000009&
         X1              =   45
         X2              =   4680
         Y1              =   1005
         Y2              =   1005
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   45
         X2              =   4680
         Y1              =   990
         Y2              =   990
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         Height          =   270
         Left            =   3195
         TabIndex        =   14
         Top             =   1440
         Width           =   420
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Increase counter by:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1620
         TabIndex        =   13
         Top             =   1440
         Width           =   1545
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Change File name:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   90
         TabIndex        =   7
         Top             =   315
         Width           =   1365
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000004&
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6120
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   4920
      TabIndex        =   3
      Top             =   120
      Width           =   2295
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H80000000&
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   4920
      TabIndex        =   2
      Top             =   480
      Width           =   2295
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H80000000&
      ForeColor       =   &H00000000&
      Height          =   2625
      Left            =   4920
      MultiSelect     =   2  'Extended
      TabIndex        =   1
      ToolTipText     =   "Use mouse, shift and control to select"
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4920
      TabIndex        =   23
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pattern :"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6120
      TabIndex        =   0
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu mnuViewer 
         Caption         =   "&Image Viewer"
      End
   End
End
Attribute VB_Name = "FileRenamer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X, Y, t As Integer
Dim Counter As Long
Dim temp, temp2, Oldfile, Newfile As String

Private Sub All_Click()
For X = File1.ListCount - 1 To 0 Step -1
File1.Selected(X) = True
Next X
Selections
Text2.SetFocus
End Sub

Private Sub ByBy_Click()
Beep
End
End Sub

Private Sub Command1_Click() ' select all
For X = File1.ListCount - 1 To 0 Step -1
File1.Selected(X) = True
Next X
Selections
Text2.SetFocus
End Sub

Private Sub Command2_Click() ' select none
For X = File1.ListCount - 1 To 0 Step -1
File1.Selected(X) = False
Next X
Selections
Text2.SetFocus
End Sub

Private Sub Command3_Click() 'preview
Selections
If t = 0 Then
temp = MsgBox("First select 1 or more files !", vbOKOnly + vbExclamation, "Renamer")
Exit Sub
End If
Screen.MousePointer = 11
List1.Clear
Counter = Val(Text3.Text)
For X = 0 To File1.ListCount - 1
If File1.Selected(X) = True Then
    GetNames
    If Option1(0).Value = True Then 'replace
        List1.AddItem Text2 & Format(Str(Counter), "000") & temp2
        Counter = Counter + Val(Label5)
    End If
    If Option1(1).Value = True Then 'before
        List1.AddItem Text2 & temp & Format(Str(Counter), "000") & temp2
        Counter = Counter + Val(Label5)
    End If
    If Option1(2).Value = True Then 'after
        List1.AddItem temp & Text2 & Format(Str(Counter), "000") & temp2
        Counter = Counter + Val(Label5)
    End If
End If
Next X
Text2.SetFocus
Screen.MousePointer = 1
End Sub

Private Sub Command4_Click() 'rename
Selections
If t = 0 Then
temp = MsgBox("First select 1 or more files !", vbOKOnly + vbExclamation, "Renamer")
Exit Sub
End If
temp = MsgBox("You're about to change the selected filenames..." & vbCr & vbCr & vbCr & "Are you sure about this ?", vbOKCancel + vbQuestion, "File-Renamer")
Text2.SetFocus
If temp = vbCancel Then Exit Sub
List1.Clear
On Error GoTo Fout
Screen.MousePointer = 11
Counter = Val(Text3.Text)
For X = 0 To File1.ListCount - 1
If File1.Selected(X) = True Then
    Oldfile = File1.Path & "\" & File1.List(X)
    GetNames
    If Option1(0).Value = True Then 'replace
        Newfile = File1.Path & "\" & Text2 & Format(Str(Counter), "000") & temp2
        Name Oldfile As Newfile
        Counter = Counter + Val(Label5)
    End If
    If Option1(1).Value = True Then 'before
        Newfile = File1.Path & "\" & Text2 & temp & Format(Str(Counter), "000") & temp2
        Counter = Counter + Val(Label5)
    End If
    If Option1(2).Value = True Then 'after
        Newfile = File1.Path & "\" & temp & Text2 & Format(Str(Counter), "000") & temp2
        Counter = Counter + Val(Label5)
    End If
End If
Next X
File1.Refresh
Command2_Click
If Check1.Value = 1 Then Text3.Text = Str(Counter)
Screen.MousePointer = 1
Text2.SetFocus
Exit Sub
Fout:
File1.Refresh
Command2_Click
MsgBox ("File exists !" & vbCr & "Check your input...")
Text2.SetFocus
Screen.MousePointer = 1
End Sub

Private Sub Command5_Click() 'stop/restart scroll
If Timer1.Enabled = True Then
Timer1.Enabled = False
Command5.Caption = "Restart scroll"
Else
Timer1.Enabled = True
Command5.Caption = "Stop scroll"
End If
Text2.SetFocus
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
Selections
End Sub

Private Sub Drive1_Change()
On Error GoTo drivefout
temp = Drive1.Drive
Dir1.Path = Left(Drive1.Drive, 2) + "\"
Debug.Print Dir1.Path
Exit Sub
drivefout:
temp = MsgBox("The selected device is not ready!", vbOKOnly & vbCritical, "File-Renamer")
End Sub

Private Sub egfwegf_Click()

End Sub

Private Sub File1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Selections
Text2.SetFocus
End Sub

Private Sub Form_Activate()
Text2.SetFocus
End Sub

Private Sub Form_Load()
Drive1.Drive = "c:\"
Dir1.Path = "c:\"
File1.Path = "c:\"
Selections
Text1.Text = File1.Pattern
List1.Clear
Text2.Text = ""

End Sub

Private Sub Selections()
t = 0
For X = 0 To File1.ListCount - 1
If File1.Selected(X) = True Then
t = t + 1
End If
Next X
Label1.Caption = "Selected : " & t & "   "
Label8.Caption = "Total files : " & File1.ListCount & "   "
End Sub

Private Sub GetNames()
Dim tel As Integer
tel = 0
For Y = Len(File1.List(X)) To 1 Step -1
tel = tel + 1
If Mid(File1.List(X), Y, 1) = "." Then
    temp2 = Right(File1.List(X), tel)
    temp = Left(File1.List(X), Len(File1.List(X)) - tel)
Exit For
End If
Next Y
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = True
temp = MsgBox("Do you want to quit ?", vbOKCancel + vbQuestion, "ReOrder")
If temp = vbOK Then End
End Sub

Private Sub Go_Click()
frmAbout.Show
End Sub

Private Sub HScroll1_Change()
Label5.Caption = HScroll1.Value
Text2.SetFocus
End Sub

Private Sub Look_Click()
Selections
If t = 0 Then
temp = MsgBox("First select 1 or more files !", vbOKOnly + vbExclamation, "ReOrder")
Exit Sub
End If
Screen.MousePointer = 11
List1.Clear
Counter = Val(Text3.Text)
For X = 0 To File1.ListCount - 1
If File1.Selected(X) = True Then
    GetNames
    If Option1(0).Value = True Then 'replace
        List1.AddItem Text2 & Format(Str(Counter), "000") & temp2
        Counter = Counter + Val(Label5)
    End If
    If Option1(1).Value = True Then 'before
        List1.AddItem Text2 & temp & Format(Str(Counter), "000") & temp2
        Counter = Counter + Val(Label5)
    End If
    If Option1(2).Value = True Then 'after
        List1.AddItem temp & Text2 & Format(Str(Counter), "000") & temp2
        Counter = Counter + Val(Label5)
    End If
End If
Next X
Text2.SetFocus
Screen.MousePointer = 1
End Sub

Private Sub None_Click()
For X = File1.ListCount - 1 To 0 Step -1
File1.Selected(X) = False
Next X
Selections
Text2.SetFocus
End Sub

Private Sub Order_Click()
Selections
If t = 0 Then
temp = MsgBox("First select 1 or more files !", vbOKOnly + vbExclamation, "ReOrder")
Exit Sub
End If
temp = MsgBox("You're about to change the selected filenames..." & vbCr & vbCr & vbCr & "Are you sure about this ?", vbOKCancel + vbQuestion, "File-Renamer")
Text2.SetFocus
If temp = vbCancel Then Exit Sub
List1.Clear
On Error GoTo Fout
Screen.MousePointer = 11
Counter = Val(Text3.Text)
For X = 0 To File1.ListCount - 1
If File1.Selected(X) = True Then
    Oldfile = File1.Path & "\" & File1.List(X)
    GetNames
    If Option1(0).Value = True Then 'replace
        Newfile = File1.Path & "\" & Text2 & Format(Str(Counter), "000") & temp2
        Name Oldfile As Newfile
        Counter = Counter + Val(Label5)
    End If
    If Option1(1).Value = True Then 'before
        Newfile = File1.Path & "\" & Text2 & temp & Format(Str(Counter), "000") & temp2
        Counter = Counter + Val(Label5)
    End If
    If Option1(2).Value = True Then 'after
        Newfile = File1.Path & "\" & temp & Text2 & Format(Str(Counter), "000") & temp2
        Counter = Counter + Val(Label5)
    End If
End If
Next X
File1.Refresh
Command2_Click
If Check1.Value = 1 Then Text3.Text = Str(Counter)
Screen.MousePointer = 1
Text2.SetFocus
Exit Sub
Fout:
File1.Refresh
Command2_Click
MsgBox ("File exists !" & vbCr & "Check your input...")
Text2.SetFocus
Screen.MousePointer = 1

End Sub

Private Sub mnuViewer_Click()
    FileRenamer.Hide
    Main.Show
End Sub

Private Sub Text1_Change()
File1.Pattern = Text1.Text
Selections
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Unload Form1
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


