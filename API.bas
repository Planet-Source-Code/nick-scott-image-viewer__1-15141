Attribute VB_Name = "API"
Option Explicit


Type SYSTEMTIME '16 bit
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

'api cacasoft
'Type TIME_ZONE_INFORMATION ' 172 Bytes
'Bias As Long
'StandardName(32) As Integer
'StandardDate As SYSTEMTIME
'StandardBias As Long
'DaylightName(32) As Integer
'DaylightDate As SYSTEMTIME
'DaylightBias As Long
'End Type

Type TIME_ZONE_INFORMATION  '32 bit
    Bias As Long
    StandardName(31) As Integer
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(31) As Integer
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type

Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation _
    As TIME_ZONE_INFORMATION) As Long
    

Public Declare Function timeGetTime Lib "winmm.dll" () As Long

Public Type MIME_DATA
    SMTP_SERVER As String
    SMTP_PORT   As Long
    SMTP_MAIL   As String
    SMTP_MAILTO As String
End Type

Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, _
        ByVal lpFile As String, ByVal lpParameters As String, ByVal _
            lpDirectory As String, ByVal nShowCmd As Long) As Long

'deshabilitar close
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nSize As Long, ByVal wFlags As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long

Const MF_BYPOSITION = &H400&
Const MF_REMOVE = &H1000&

'Cambiar de Icono
Public Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Public Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Any) As Long
Public Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Public Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

'constants
Public Const GCL_HCURSOR = -12
Global curSelect As StdPicture
Public Function GetLocalTZ(Optional ByRef strTZName As String) As Long

    Dim objTimeZone As TIME_ZONE_INFORMATION
    Dim lngResult As Long
    Dim i As Long
    
    lngResult = GetTimeZoneInformation&(objTimeZone)

    Select Case lngResult
        Case 0&, 1& 'hora estandar
            GetLocalTZ = -(objTimeZone.Bias + objTimeZone.StandardBias) * 60 'into minutes

            For i = 0 To 31
                If objTimeZone.StandardName(i) = 0 Then Exit For
                strTZName = strTZName & Chr(objTimeZone.StandardName(i))
            Next
        Case 2& 'dias salvan horas de luz
        
            GetLocalTZ = -(objTimeZone.Bias + objTimeZone.DaylightBias) * 60 'into minutes
            
            For i = 0 To 31
                If objTimeZone.DaylightName(i) = 0 Then Exit For
                strTZName = strTZName & Chr(objTimeZone.DaylightName(i))
            Next
    End Select

End Function

Sub DisableX(frm As Form)
     
    Dim hMenu As Long
    Dim nCount As Long
    
    hMenu = GetSystemMenu(frm.hWnd, 0)
    nCount = GetMenuItemCount(hMenu)

    Call RemoveMenu(hMenu, nCount - 1, MF_REMOVE Or MF_BYPOSITION)
    Call RemoveMenu(hMenu, nCount - 2, MF_REMOVE Or MF_BYPOSITION)

    DrawMenuBar frm.hWnd
     
    Const SC_SIZE = &HF000
    Const MF_BYCOMMAND = &H0
    
    hMenu = GetSystemMenu(frm.hWnd, 0)
    
    Call DeleteMenu(hMenu, SC_SIZE, MF_BYCOMMAND)

     
End Sub
