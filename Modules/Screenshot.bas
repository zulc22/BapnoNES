Attribute VB_Name = "Screenshot"
Option Explicit

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As picBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Public Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Public Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Public Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColor0 As Long
    bmiColor1 As Long
    bmiColor2 As Long
End Type

Public Type picBmp
    Size As Long
    Type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type
Public Function CreateBitmapPicture(ByVal hBmp As Long) As Picture
    Dim r As Long, Pic As picBmp, IPic As IPicture, IID_IDispatch As GUID
    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    With Pic
        .Size = Len(Pic)
        .Type = vbPicTypeBitmap
        .hBmp = hBmp
    End With
    r = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)
    Set CreateBitmapPicture = IPic
End Function
Public Function ExportBmp(PBox As PictureBox) As Picture
    Dim hDCMemory As Long, hBmp As Long, hBmpPrev As Long, r As Long, hdcSrc As Long, hPal As Long, hPalPrev As Long, RasterCapsScrn As Long, HasPaletteScrn As Long, PaletteSizeScrn As Long, WidthPix As Single, HeightPix As Single
    WidthPix = PBox.ScaleWidth
    HeightPix = PBox.ScaleHeight
    hdcSrc = GetWindowDC(PBox.hwnd)
    hDCMemory = CreateCompatibleDC(hdcSrc)
    hBmp = CreateCompatibleBitmap(hdcSrc, WidthPix, HeightPix)
    hBmpPrev = SelectObject(hDCMemory, hBmp)
    r = BitBlt(hDCMemory, 0, 0, WidthPix, HeightPix, hdcSrc, 0, 0, vbSrcCopy)
    hBmp = SelectObject(hDCMemory, hBmpPrev)
    r = DeleteDC(hDCMemory)
    r = ReleaseDC(PBox.hwnd, hdcSrc)
    Set ExportBmp = CreateBitmapPicture(hBmp)
End Function
