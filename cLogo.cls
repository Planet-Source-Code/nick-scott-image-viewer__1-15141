VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
    left As Long
    top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long

Private Const LOGPIXELSX = 88    'pixels/inch lógicos en X
Private Const LOGPIXELSY = 90    'pixels/inch lógicos en Y

Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long

Private Const LF_FACESIZE = 32

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Const FW_NORMAL = 400
Private Const FW_BOLD = 700
Private Const FF_DONTCARE = 0
Private Const DEFAULT_PITCH = 0
Private Const DEFAULT_CHARSET = 1

Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long

Private Const CLR_INVALID = -1

Private Declare Function SetGraphicsMode Lib "gdi32" (ByVal hDC As Long, ByVal iMode As Long) As Long

Private Const GM_ADVANCED = 2

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long

'Ctes. lfQuality:
Private Const DEFAULT_QUALITY = 0           'Apariencia de la fuente es estabelcida a predeterminada
Private Const DRAFT_QUALITY = 1             'Apariencia es menos importante que PROOF_QUALITY.
Private Const PROOF_QUALITY = 2             'Mejor calidad de carácter
Private Const NONANTIALIASED_QUALITY = 3    'No se suavizan los bordes, aunque el sistema lo haga
Private Const ANTIALIASED_QUALITY = 4       'Se suaviza si el sistema lo hace

Private m_picThis As PictureBox
Private m_sCaption As String
Private m_bRGBStart(1 To 3) As Integer
Private m_oStartColor As OLE_COLOR
Private m_bRGBEnd(1 To 3) As Integer
Private m_oEndColor As OLE_COLOR
Public Property Let Caption(ByVal sCaption As String)
    m_sCaption = sCaption
End Property
Public Property Get Caption() As String
    Caption = m_sCaption
End Property
Public Property Let DrawingObject(ByRef picThis As PictureBox)
    Set m_picThis = picThis
End Property
Public Property Get StartColor() As OLE_COLOR
    StartColor = m_oStartColor
End Property
Public Property Let StartColor(ByVal oColor As OLE_COLOR)
    
    Dim lColor As Long
    
    If (m_oStartColor <> oColor) Then
        m_oStartColor = oColor
        OleTranslateColor oColor, 0, lColor
        m_bRGBStart(1) = lColor And &HFF&
        m_bRGBStart(2) = ((lColor And &HFF00&) \ &H100)
        m_bRGBStart(3) = ((lColor And &HFF0000) \ &H10000)
        If Not (m_picThis Is Nothing) Then
            Draw
        End If
    End If
    
End Property
Public Property Get EndColor() As OLE_COLOR
    EndColor = m_oEndColor
End Property
Public Property Let EndColor(ByVal oColor As OLE_COLOR)

    Dim lColor As Long
    
    If (m_oEndColor <> oColor) Then
        m_oEndColor = oColor
        OleTranslateColor oColor, 0, lColor
        m_bRGBEnd(1) = lColor And &HFF&
        m_bRGBEnd(2) = ((lColor And &HFF00&) \ &H100)
        m_bRGBEnd(3) = ((lColor And &HFF0000) \ &H10000)
        If Not (m_picThis Is Nothing) Then
            Draw
        End If
    End If
End Property
Public Sub Draw()
    
    Dim lHeight As Long, lWidth As Long
    Dim lYStep As Long
    Dim lY As Long
    Dim bRGB(1 To 3) As Integer
    Dim tLF As LOGFONT
    Dim hFnt As Long
    Dim hFntOld As Long
    Dim lR As Long
    Dim tR As RECT
    Dim rct As RECT
    Dim hBr As Long
    Dim hDC As Long
    Dim dR(1 To 3) As Double
    
    On Error GoTo DrawError

    hDC = m_picThis.hDC
    
    lHeight = m_picThis.Height \ Screen.TwipsPerPixelY
    rct.Right = m_picThis.Width \ Screen.TwipsPerPixelY
    
    'Establecer graduación de 255 pixeles
    lYStep = lHeight \ 255
    
    If (lYStep = 0) Then
        lYStep = 1
    End If
    
    rct.Bottom = lHeight
    
    LSet tR = rct
    
    bRGB(1) = m_bRGBStart(1)
    bRGB(2) = m_bRGBStart(2)
    bRGB(3) = m_bRGBStart(3)
    dR(1) = m_bRGBEnd(1) - m_bRGBStart(1)
    dR(2) = m_bRGBEnd(2) - m_bRGBStart(2)
    dR(3) = m_bRGBEnd(3) - m_bRGBStart(3)
        
    For lY = lHeight To 0 Step -lYStep
        
        'dibujar barra
        rct.top = rct.Bottom - lYStep
        hBr = CreateSolidBrush((bRGB(3) * &H10000 + bRGB(2) * &H100& + bRGB(1)))
        FillRect hDC, rct, hBr
        DeleteObject hBr
        rct.Bottom = rct.top
        
        'ajustar color
        bRGB(1) = m_bRGBStart(1) + dR(1) * (lHeight - lY) / lHeight
        bRGB(2) = m_bRGBStart(2) + dR(2) * (lHeight - lY) / lHeight
        bRGB(3) = m_bRGBStart(3) + dR(3) * (lHeight - lY) / lHeight
        
        'Debug.Print bRGB(1), (lHeight - lY) / lHeight
    Next lY
    
    pOLEFontToLogFont m_picThis.Font, hDC, tLF
    
    tLF.lfEscapement = 900
    hFnt = CreateFontIndirect(tLF)
    If (hFnt <> 0) Then
        hFntOld = SelectObject(hDC, hFnt)
        lR = TextOut(hDC, 0, lHeight - 16, m_sCaption, lstrlen(m_sCaption))
        SelectObject hDC, hFntOld
        DeleteObject hFnt
    End If
    
    m_picThis.Refresh
    
    Exit Sub
    
DrawError:
    Debug.Print "Problema: " & Err.Description
End Sub
Private Sub pOLEFontToLogFont(fntThis As StdFont, hDC As Long, tLF As LOGFONT)
    
    Dim sFont As String
    Dim iChar As Integer

    'Convertir OLE(StdFont)=>(LOGFONT) estructura
    With tLF
        sFont = fntThis.Name
        
        For iChar = 1 To Len(sFont)

        Next iChar
        
        .lfHeight = -MulDiv((fntThis.Size), (GetDeviceCaps(hDC, LOGPIXELSY)), 72)
        .lfItalic = fntThis.Italic
        If (fntThis.Bold) Then
            .lfWeight = FW_BOLD
        Else
            .lfWeight = FW_NORMAL
        End If
        
        .lfUnderline = fntThis.Underline
        .lfStrikeOut = fntThis.Strikethrough
        .lfCharSet = fntThis.Charset
        .lfQuality = ANTIALIASED_QUALITY
    End With

End Sub
Private Sub Class_Initialize()
    StartColor = &H0
    EndColor = vbButtonFace
End Sub
