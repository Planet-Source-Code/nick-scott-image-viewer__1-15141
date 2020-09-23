VERSION 5.00
Begin VB.UserControl ThumbBrowse 
   Alignable       =   -1  'True
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   2610
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6480
   ScaleHeight     =   174
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   432
   ToolboxBitmap   =   "ThumbBrowse.ctx":0000
   Begin VB.VScrollBar VScroll 
      Height          =   2550
      Left            =   6105
      TabIndex        =   1
      Top             =   15
      Width           =   270
   End
   Begin VB.PictureBox picThumbPane 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1050
      Left            =   45
      ScaleHeight     =   70
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   153
      TabIndex        =   0
      Top             =   150
      Width           =   2295
   End
End
Attribute VB_Name = "ThumbBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Type Thumb
    ThumbPath As String
    ThumbCaption As String
    ThumbSize As Long
    ThumbDate As Date
    HdrWidth As Long
    HdrHeight As Long
    HdrPlanes As Long
    HdrColors As Long
End Type

Dim Thumbs() As Thumb

Dim mlColor As Long
Dim mlHeight As Long

Const m_def_ThumbBorder = 10
Const m_def_ThumbWidth = 80
Const m_def_ThumbHeight = 60
Const m_def_ColorLight = &HFFFFFF
Const m_def_ColorDark = &HE0E0E0

Dim m_ThumbBorder As Long
Dim m_ThumbWidth As Long
Dim m_ThumbHeight As Long
Dim m_ColorLight As OLE_COLOR
Dim m_ColorDark As OLE_COLOR

Event Click()
Event Change()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event Resize()
Event ThumbClick(Index As Long, ThumbPath As String, ThumbCaption As String, ThumbSize As Long, ThumbDate As Date, Width As Long, Height As Long, Planes As Long, Colors As Long)
Attribute ThumbClick.VB_Description = "This event fires when a thumbnail is clicked. The Width,Height,Planes and Colors parameters will only be returned for bitmaps. Otherwise those fields will be 0. "
Attribute ThumbClick.VB_MemberFlags = "200"

Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Scroll()

Public Sub AddThumb(psPath As String, psCaption As String, plSize As Long, pdDate As Date)
    Dim lsIndex As Long

    
    On Error Resume Next
    
    lsIndex = UBound(Thumbs) + 1
    ReDim Preserve Thumbs(lsIndex)
    
    Thumbs(lsIndex).ThumbPath = psPath
    Thumbs(lsIndex).ThumbCaption = psCaption
    Thumbs(lsIndex).ThumbSize = plSize
    Thumbs(lsIndex).ThumbDate = pdDate
        
    '*****************************************************************
    'Get Extended Information if type is BMP
    '*****************************************************************
    If UCase$(Right(psPath, 3)) = "BMP" Then
    
        Thumbs(lsIndex).HdrHeight = Nfo.Height
        Thumbs(lsIndex).HdrWidth = Nfo.Width
        Thumbs(lsIndex).HdrPlanes = Nfo.Planes
        Thumbs(lsIndex).HdrColors = Nfo.Colors
    End If
    
    X = 0
    Y = (lsIndex - 1) * (m_ThumbHeight + m_ThumbBorder)
    
    If mlColor = m_ColorLight Then
        mlColor = m_ColorDark
    Else
        mlColor = m_ColorLight
    End If
    
    picThumbPane.DrawStyle = vbSolid
    
    mlHeight = mlHeight + (m_ThumbHeight + m_ThumbBorder)
    picThumbPane.Height = mlHeight
    
    '*****************************************************************
    'Draw Thumb
    '*****************************************************************
    picThumbPane.Line (X, Y)-(picThumbPane.Width, Y + (m_ThumbHeight + m_ThumbBorder)), mlColor, BF
    picThumbPane.PaintPicture LoadPicture(psPath), X, Y + ((m_ThumbBorder) / 2), m_ThumbWidth, m_ThumbHeight, , , , , vbSrcCopy
    
    '*****************************************************************
    'Caption
    '*****************************************************************
    picThumbPane.CurrentX = m_ThumbWidth + 10
    picThumbPane.CurrentY = (Y + (m_ThumbHeight / 2)) - 5
    picThumbPane.Print psCaption
        
    '*****************************************************************
    'Size
    '*****************************************************************
    picThumbPane.CurrentX = m_ThumbWidth + 130
    picThumbPane.CurrentY = (Y + (m_ThumbHeight / 2)) - 5
    picThumbPane.Print Format(plSize, "###,### Bytes")
    
    '*****************************************************************
    'Date
    '*****************************************************************
    picThumbPane.CurrentX = m_ThumbWidth + 245
    picThumbPane.CurrentY = (Y + (m_ThumbHeight / 2)) - 5
    picThumbPane.Print pdDate
        
    If picThumbPane.Height > UserControl.ScaleHeight Then
        VScroll.Max = mlHeight - UserControl.ScaleHeight
        VScroll.Visible = True
    Else
        VScroll.Visible = False
    End If
    DoEvents
End Sub

Private Sub picThumbPane_Click()
    RaiseEvent Click
End Sub

Private Sub picThumbPane_Change()
    RaiseEvent Change
End Sub

Private Sub picThumbPane_DblClick()
    RaiseEvent DblClick
End Sub

Public Property Get hDC() As Long
Attribute hDC.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
    hDC = UserControl.hDC
End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = UserControl.hWnd
End Property

Private Sub picThumbPane_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub picThumbPane_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub picThumbPane_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub picThumbPane_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    Dim i As Long
    
    On Error Resume Next
    
    '*********************************************************
    'Enumerate elements
    '*********************************************************
    For i = 0 To UBound(Thumbs)
        '*********************************************************
        'Check Range
        '*********************************************************
        If Y >= (i - 1 * (m_ThumbHeight + m_ThumbBorder)) And _
           Y < (i * (m_ThumbHeight + m_ThumbBorder)) Then

            'picThumbPane.ToolTipText = Thumbs(i).ThumbPath
            RaiseEvent ThumbClick(i, Thumbs(i).ThumbPath, Thumbs(i).ThumbCaption, Thumbs(i).ThumbSize, Thumbs(i).ThumbDate, Thumbs(i).HdrWidth, Thumbs(i).HdrHeight, Thumbs(i).HdrPlanes, Thumbs(i).HdrColors)
            i = UBound(Thumbs)
        End If
    Next
End Sub

Private Sub UserControl_Initialize()
    ReDim Thumbs(0)
    mlHeight = 0
    VScroll.Min = 0
    VScroll.Max = 1
    VScroll.Visible = False
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next

    If UserControl.Width < 2000 Then UserControl.Width = 2000
    If UserControl.Height < 2000 Then UserControl.Height = 2000
    
    VScroll.Top = 0
    VScroll.Left = (UserControl.ScaleWidth - VScroll.Width)
    VScroll.Height = UserControl.ScaleHeight
    
    picThumbPane.Top = 0
    picThumbPane.Height = UserControl.ScaleHeight
    picThumbPane.Width = VScroll.Left
    
    If m_ThumbHeight = 0 And m_ThumbBorder = 0 Then
        VScroll.SmallChange = m_def_ThumbWidth + m_def_ThumbBorder
    Else
        VScroll.SmallChange = m_ThumbHeight + m_ThumbBorder
    End If
    VScroll.LargeChange = picThumbPane.ScaleHeight
    RaiseEvent Resize
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
End Sub

Private Sub UserControl_InitProperties()
    m_ThumbWidth = m_def_ThumbWidth
    m_ThumbHeight = m_def_ThumbHeight
    m_ThumbBorder = m_def_ThumbBorder
    m_ColorLight = m_def_ColorLight
    m_ColorDark = m_def_ColorDark
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    m_ThumbWidth = PropBag.ReadProperty("ThumbWidth", m_def_ThumbWidth)
    m_ThumbHeight = PropBag.ReadProperty("ThumbHeight", m_def_ThumbHeight)
    m_ThumbBorder = PropBag.ReadProperty("ThumbBorder", m_def_ThumbBorder)
    m_ColorLight = PropBag.ReadProperty("ColorLight", m_def_ColorLight)
    m_ColorDark = PropBag.ReadProperty("ColorDark", m_def_ColorDark)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ThumbWidth", m_ThumbWidth, m_def_ThumbWidth)
    Call PropBag.WriteProperty("ThumbHeight", m_ThumbHeight, m_def_ThumbHeight)
    Call PropBag.WriteProperty("ThumbBorder", m_ThumbBorder, m_def_ThumbBorder)
    Call PropBag.WriteProperty("ColorLight", m_ColorLight, m_def_ColorLight)
    Call PropBag.WriteProperty("ColorDark", m_ColorDark, m_def_ColorDark)
End Sub

Public Property Get ThumbWidth() As Long
Attribute ThumbWidth.VB_ProcData.VB_Invoke_Property = ";Scale"
    ThumbWidth = m_ThumbWidth
End Property

Public Property Let ThumbWidth(ByVal New_ThumbWidth As Long)
    m_ThumbWidth = New_ThumbWidth
    PropertyChanged "ThumbWidth"
End Property

Public Property Get ThumbHeight() As Long
Attribute ThumbHeight.VB_ProcData.VB_Invoke_Property = ";Scale"
    ThumbHeight = m_ThumbHeight
End Property

Public Property Let ThumbHeight(ByVal New_ThumbHeight As Long)
    m_ThumbHeight = New_ThumbHeight
    PropertyChanged "ThumbHeight"
End Property

Public Property Get ThumbBorder() As Long
    ThumbBorder = m_ThumbBorder
End Property

Public Property Let ThumbBorder(ByVal New_ThumbBorder As Long)
    m_ThumbBorder = New_ThumbBorder
    PropertyChanged "ThumbBorder"
End Property

Private Sub VScroll_Change()
    picThumbPane.Top = -VScroll.Value
End Sub

Private Sub VScroll_Scroll()
    RaiseEvent Scroll
    picThumbPane.Top = -VScroll.Value
End Sub

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = "StandardColor;Appearance"
Attribute BackColor.VB_UserMemId = -501
    BackColor = picThumbPane.BackColor
End Property

Public Sub Cls()
Attribute Cls.VB_Description = "Clears graphics and text generated at run time from a Form, Image, or PictureBox."
     ReDim Thumbs(0)
     picThumbPane.Cls
     VScroll.Visible = False
End Sub

Private Sub picThumbPane_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub picThumbPane_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Public Property Get ColorLight() As OLE_COLOR
Attribute ColorLight.VB_Description = "The light color to use for each row."
Attribute ColorLight.VB_ProcData.VB_Invoke_Property = "StandardColor;Appearance"
    ColorLight = m_ColorLight
End Property

Public Property Let ColorLight(ByVal New_ColorLight As OLE_COLOR)
    m_ColorLight = New_ColorLight
    PropertyChanged "ColorLight"
End Property

Public Property Get ColorDark() As OLE_COLOR
Attribute ColorDark.VB_Description = "The dark color to use for each row."
Attribute ColorDark.VB_ProcData.VB_Invoke_Property = "StandardColor;Appearance"
    ColorDark = m_ColorDark
End Property

Public Property Let ColorDark(ByVal New_ColorDark As OLE_COLOR)
    m_ColorDark = New_ColorDark
    PropertyChanged "ColorDark"
End Property

