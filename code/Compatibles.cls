VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Compatibles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal Hwnd As Long, lpRect As RECT) As Long
Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal hHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal BLENDFUNCTION As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFOHEADER, ByVal un As Long, ByVal lplpVoid As Long, ByVal Handle As Long, ByVal dw As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal Hwnd As Long) As Long

Const GMEM_MOVEABLE As Long = &H2
Const AC_SRC_ALPHA = &H1
Const AC_SRC_OVER = &H0
Const DIB_RGB_Colors As Long = 0


Private Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

Private Type Size
    cx As Long
    cy As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Const ULW_ALPHA = &H2
Const GWL_EXSTYLE As Long = -20
Const WS_EX_LAYERED = &H80000

Public CompatibleDC As Long
Private DIB As Long
Private mWidth As Long, mHeight As Long, BlendFunc As BLENDFUNCTION

Public Property Get Height() As Long 'DC高度
Height = mHeight
End Property

Public Property Get Width() As Long 'DC宽度
Width = mWidth
End Property

'初始化，Hdc可以传入窗口的Hdc，width/height为欲创建DC的大小
Public Sub Create(hdc As Long, Width As Long, Height As Long)
    Dim bmHeader As BITMAPINFOHEADER, BlendFunc As BLENDFUNCTION
    
    mWidth = Width
    mHeight = Height
    
    With BlendFunc
        .AlphaFormat = AC_SRC_ALPHA
        .BlendFlags = 0
        .BlendOp = AC_SRC_OVER
        .SourceConstantAlpha = 255
    End With
    
    With bmHeader
        .biSize = Len(bmHeader)
        .biBitCount = 32
        .biHeight = Height
        .biWidth = Width
        .biPlanes = 1
        .biSizeImage = .biWidth * .biHeight * 4
    End With
    
    CompatibleDC = CreateCompatibleDC(hdc)
    DIB = CreateDIBSection(CompatibleDC, bmHeader, DIB_RGB_Colors, ByVal 0, 0, 0)
    DeleteObject SelectObject(CompatibleDC, DIB)
    
    If CompatibleDC = 0 Then Err.Raise 4040, , "Runtime Error" & vbCrLf & "Failed to create a compatible DC ."
End Sub

'将DC里的图像剪裁后贴入目标DC
'DstDC:目标DC
'X/Y:绘制在目标DC的坐标
'OffsetX/OffsetY：剪裁偏移坐标
'height/width：剪裁大小
'Alpha:混合时的整体透明度
Public Sub PresentWithClip(ByVal DstDC As Long, x As Long, y As Long, offsetX As Long, offsetY As Long, Width As Long, Height As Long, Optional Alpha As Byte = 255)
    Dim SrcSize As RECT, BlendFunLong As Long
    
    If CompatibleDC = 0 Then Err.Raise 4041, , "Runtime Error" & vbCrLf & "Present before create ."
    
    If offsetX + Width > mWidth Then Width = mWidth - offsetX
    If offsetY + Height > mHeight Then Height = mHeight - offsetY
    
    SrcSize.Right = Width
    SrcSize.Bottom = Height
    
    With BlendFunc
        .BlendOp = AC_SRC_OVER
        .AlphaFormat = AC_SRC_ALPHA
        .BlendFlags = 0
        .SourceConstantAlpha = Alpha
    End With
    
    RtlMoveMemory BlendFunLong, BlendFunc, 4
    AlphaBlend DstDC, x, y, SrcSize.Right, SrcSize.Bottom, CompatibleDC, offsetX, offsetY, SrcSize.Right, SrcSize.Bottom, BlendFunLong
    
End Sub

'将DC里的完整图像直接贴入目标DC
'DstDC:目标DC
'X/Y:绘制在目标DC的坐标
'Alpha:混合时的整体透明度
Public Sub Present(ByVal DstDC As Long, ByVal x As Long, ByVal y As Long, Optional Alpha As Byte = 255)
    Dim SrcSize As RECT, BlendFunLong As Long
    
    If CompatibleDC = 0 Then Err.Raise 4041, , "Runtime Error" & vbCrLf & "Present before create ."
    
    SrcSize.Right = mWidth
    SrcSize.Bottom = mHeight
    
    With BlendFunc
        .BlendOp = AC_SRC_OVER
        .AlphaFormat = AC_SRC_ALPHA
        .BlendFlags = 0
        .SourceConstantAlpha = Alpha
    End With
    
    RtlMoveMemory BlendFunLong, BlendFunc, 4
    AlphaBlend DstDC, x, y, SrcSize.Right, SrcSize.Bottom, CompatibleDC, SrcSize.Left, SrcSize.Top, SrcSize.Right, SrcSize.Bottom, BlendFunLong
    
End Sub

Sub Dispose()
    DeleteObject CompatibleDC
    DeleteObject DIB
End Sub

Private Sub Class_Initialize()
    AddMember Me
End Sub
