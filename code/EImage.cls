VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim mImage As Long, mImgSize As ESize, mSize As ESize, mPath As String
Public Property Get Hwnd() As Long
    Hwnd = mImage
End Property
Public Property Get Size() As ESize
    Size = mSize
End Property
Public Property Let Size(nSize As ESize)
    mSize = nSize
    AddRefImage Me
End Property
Public Property Get ImgSize() As ESize
    ImgSize = mImgSize
End Property
Public Property Get Path() As String
    Path = mPath
End Property
Public Property Let Path(ByVal nPath As String)
    If mImage <> 0 Then GdipDisposeImage mImage
    Dim Width As Long, Height As Long
    mPath = nPath
    GdipCreateBitmapFromFile StrPtr(nPath), mImage
    GdipGetImageWidth mImage, Width
    GdipGetImageHeight mImage, Height
    mImgSize.Width = Width: mImgSize.Height = Height
    AddRefImage Me
End Property
Public Sub Blur(Optional ByVal BlurRadius As Long = 40)
    Dim Effect As Long
    Dim p As BlurParams
    GdipCreateEffect2 GdipEffectType.Blur, Effect
    p.Radius = BlurRadius
    GdipSetEffectParameters Effect, p, LenB(p)
    GdipBitmapApplyEffect mImage, Effect, NewRectL(0, 0, mImgSize.Width, mImgSize.Height), 0, 0, 0
    AddRefImage Me
End Sub
Public Sub ClipCircle()
    Dim Path As Long, bmpGraph As Long, BMP As Long, BMP2 As Long, tempGraph As Long

    GdipCreateBitmapFromScan0 mImgSize.Width, mImgSize.Height, ByVal 0, PixelFormat32bppARGB, ByVal 0, BMP2
    GdipCreateBitmapFromScan0 mImgSize.Width * 2, mImgSize.Height * 2, ByVal 0, PixelFormat32bppARGB, ByVal 0, BMP

    GdipGetImageGraphicsContext BMP2, tempGraph
    GdipGetImageGraphicsContext BMP, bmpGraph
    GdipSetSmoothingMode tempGraph, SmoothingModeAntiAlias
    GdipSetSmoothingMode bmpGraph, SmoothingModeAntiAlias
    
    GdipCreatePath FillModeWinding, Path
    GdipAddPathEllipseI Path, 1, 1, mImgSize.Width * 2 - 2, mImgSize.Height * 2 - 2
    GdipSetClipPath bmpGraph, Path, CombineModeReplace
    
    GdipDrawImageRectI bmpGraph, mImage, 0, 0, mImgSize.Width * 2, mImgSize.Height * 2
    GdipResetClip bmpGraph
    
    GdipSetInterpolationMode tempGraph, InterpolationModeHighQualityBilinear
    
    GdipDrawImageRectI tempGraph, BMP, 0, 0, mImgSize.Width - 1, mImgSize.Height - 1

    GdipDeletePath Path
    GdipDeleteGraphics bmpGraph
    GdipDisposeImage BMP
    GdipDeleteGraphics tempGraph
    GdipDisposeImage mImage
    
    mImage = BMP2
    
    AddRefImage Me
End Sub
Private Sub Class_Initialize()
    AddMember Me
End Sub
Public Sub Dispose()
    If Hwnd <> 0 Then GdipDisposeImage Hwnd
End Sub
