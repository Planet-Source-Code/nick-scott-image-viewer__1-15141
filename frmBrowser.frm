VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmBrowser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Image Viewer Help"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   8295
   Icon            =   "frmBrowser.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   8295
   StartUpPosition =   2  'CenterScreen
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8055
      ExtentX         =   14208
      ExtentY         =   10610
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   -1  'True
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
      End
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub SSCommand1_Click()
On Error GoTo oops
WebBrowser1.GoBack
oops:
Exit Sub
End Sub

Private Sub SSCommand2_Click()
On Error GoTo whoops
WebBrowser1.GoForward
whoops:
Exit Sub
End Sub

Private Sub SSCommand3_Click()
Unload frmBrowser
End Sub
