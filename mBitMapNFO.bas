Attribute VB_Name = "mBitMapNFO"
'************************************************************************************
'*  Vivid ThumbBrowse Control Version 1  (Bitmap extended information module)       *
'*       Written by Kelly S. Elias                                                  *
'*            eliask@cadvision.com                                                  *
'*                                                                                  *
'*  Notes                                                                           *
'*      - This was tested under VB6. (Should still work under VB5 SP3.              *
'*      - The extended infomation (planes, width, height and colors) will only work *
'*        for bitmaps. JPG's, and GIF's will return 0 in these fields.              *
'*        (If you know how to read there headers please let me know at the address  *
'*         below)                                                                   *
'*                                                                                  *
'*  Known Issues                                                                    *
'*      - Thumbnail size can be changed while other thumbs with a different         *
'*        size are already loaded. This causes problems with determining            *
'*        how to place new thumbs and which thumb was clicked. You should           *
'*        always ensure your program prevents this, by only allowing the control to *
'*        host one size of of thumbnail at a single time, and calling the CLS method*
'*        before you change thumb sizes.                                            *
'*      - The ThumbBrowse Control will not dynamically re-size at runtime when      *
'*        loaded with more then 1 screen of information.                            *
'*                                                                                  *
'*  If you have any questions or suggestions my e-mail address is above, let        *
'*  me know what you think.                                                         *
'*                                                                                  *
'*  Feel free to distribute this code freely, as long as all credit is given        *
'*  to the author and this entire comment block remains.  Feel free to modify       *
'*  the code in anyway you want. If you make some improvments to speed it up        *
'*  let me know.                                                                    *
'*                                                                                  *
'*  If you know how to read the header information of JPG and GIF files PLEASE let  *
'*  me know.
'*                                                                                  *
'*  Copyright Kelly S. Elias 1999, All Rights Reserved.                             *
'*  eliask@cadvision.com                                                            *
'************************************************************************************
  
  
  Option Explicit
  
  Private Type BITMAPINFOHEADER
    biSize            As Long
    biWidth           As Long
    biHeight          As Long
    biPlanes          As Integer
    biBitCount        As Integer
    biCompression     As Long
    biSizeImage       As Long
    biXPelsPerMeter   As Long
    biYPelsPerMeter   As Long
    biClrUsed         As Long
    biClrImportant    As Long
  End Type
  
  Private Type BITMAPFILEHEADER
    bfType            As Integer
    bfSize            As Long
    bfReserved1       As Integer
    bfReserved2       As Integer
    bfOffBits         As Long
  End Type
  
  Public Type BITMAPINFO
    Width             As Long
    Height            As Long
    Planes            As Long
    Colors            As Long
  End Type

Public Function GetBitmapInfo(psPath As String) As BITMAPINFO
  Dim f As Integer
  Dim tmp As String
  Dim FileHeader As BITMAPFILEHEADER
  Dim InfoHeader As BITMAPINFOHEADER
  Dim i As BITMAPINFO
  
  On Error Resume Next
 
  f = FreeFile
  Open psPath For Binary Access Read As #f
  Get #f, , FileHeader
  Get #f, , InfoHeader
  Close #f
      
  i.Width = InfoHeader.biWidth
  i.Height = InfoHeader.biHeight
  i.Planes = InfoHeader.biPlanes
  i.Colors = CLng((InfoHeader.biBitCount & 2 ^ InfoHeader.biBitCount) / 1000)
  
  GetBitmapInfo = i
End Function

