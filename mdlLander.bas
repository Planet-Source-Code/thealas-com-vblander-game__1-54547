Attribute VB_Name = "mdlLander"
Option Explicit

' All this api's are explained in MSDN, but you need to know at least some basics of c/c++ language
' Well.. if you want to make games, then you will have to learn it, it is a LOT better then VB.

' Copy DC's by using many advanced methods
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
' Select GDI object, basic GDI function
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
' Delete object to free memory
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
' Create DC that is compatible with device
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
' Delete DC from memory
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
' GetObject, used to retrieve BITMAP object form StdPicture's handle
Public Declare Function GetObjectA Lib "gdi32" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
' Simple api to create DIB section, nice stuff
Public Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As HBITMAPINFOHEADER, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
' Draws a pixel anywhere on the screen, DC that is...
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
' Gets any visible pixel's color
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
' Gets current mouse cursor position
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
' Get the miliseconds interval sense windows stared
Public Declare Function GetTickCount Lib "kernel32" () As Long
' If a certain key was pressed or released
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
' Change current position on some DC, used for lines mostly...
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
' Draw a line to a specified point, begins at the currently selected point (from MoveTo function)
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
' Creates a pen for drawing
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
' Fill and area (specified color) with currently selected color
Public Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
' Draw text on DC with currently selected font
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
' Change background mode, mostly for text operations
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long

' Color constants, for DIB section
Public Const DIB_PAL_COLORS = 1
Public Const DIB_RGB_COLORS = 0

' 20ms for each frame
Public Const TMR_INTERVAL = 20

' RECT structure
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

' Position, point...
Public Type POINTAPI
        x As Long
        y As Long
End Type

'// GDI objects
Public Type HBITMAPINFOHEADER
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

Public Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

' Holder for bitmap, hdc and mask stuffs...
Public Type HHOLDER
    hdc As Long
    hBMP As Long
    mhDC As Long
    mBMP As Long
End Type

' Some lander's data...
Public lmdGravity           As Single ' Hmmm... I dont think I've included an option to change gravity
Public lmdAccel             As Single

Public lmdMaxFuel           As Single
Public lmdFuel              As Single

'**********************************************************************
'*  To free memory
'**********************************************************************
Public Sub DeleteHolder(hld As HHOLDER)
    ' Select it and delete it
    DeleteObject SelectObject(hld.hdc, hld.hBMP)
    DeleteObject hld.hdc
    DeleteObject SelectObject(hld.mhDC, hld.mBMP)
    DeleteObject hld.mhDC
End Sub
