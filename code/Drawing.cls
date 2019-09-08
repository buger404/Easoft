VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Drawing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim mSize As ESize, MGraphics As Long, mImage As Long, mPath As Long, mClipPath As Long
Public Sub Clear()
    GdipGraphicsClear MGraphics, 0
End Sub
Public Property Get Image() As Long
    Image = mImage
End Property
Public Property Get Hwnd() As Long
    Hwnd = MGraphics
End Property
Public Property Get Size() As ESize
    Size = mSize
End Property
Public Property Let Size(nSize As ESize)
    mSize = nSize
    GdipDeleteGraphics MGraphics
    GdipDisposeImage mImage
    GdipCreateBitmapFromScan0 mSize.Width, mSize.Height, ByVal 0, PixelFormat32bppARGB, ByVal 0, mImage
    GdipGetImageGraphicsContext mImage, MGraphics
    GdipSetSmoothingMode MGraphics, SmoothingModeAntiAlias
    GdipSetTextRenderingHint MGraphics, TextRenderingHintAntiAliasGridFit
End Property
Public Sub AddLine(Pos As EPosition, Pos2 As EPosition)
    GdipAddPathLine mPath, Pos.x, Pos.y, Pos2.x, Pos2.y
End Sub
Public Sub AddArc(Pos As EPosition, Size As ESize, Start As Long, Angle As Long)
    GdipAddPathArc mPath, Pos.x, Pos.y, Size.Width, Size.Height, Start, Angle
End Sub
Public Sub AddRect(Pos As EPosition, Size As ESize)
    GdipAddPathRectangle mPath, Pos.x, Pos.y, Size.Width, Size.Height
End Sub
Public Sub AddEllipse(Pos As EPosition, Size As ESize)
    GdipAddPathEllipse mPath, Pos.x, Pos.y, Size.Width, Size.Height
End Sub
Public Sub FillPath(Color As ColorMix)
    GdipClosePathFigure mPath
    GdipFillPath MGraphics, Color.brush, mPath
    GdipResetPath mPath
End Sub
Public Sub DrawPath(Color As ColorMix)
    GdipClosePathFigure mPath
    GdipDrawPath MGraphics, Color.pen, mPath
    GdipResetPath mPath
End Sub
Public Sub DrawShape(Shape As EShape, Pos As EPosition, Size As ESize, Radian As Long, Color As ColorMix)
    Select Case Shape
        Case EShape.Square
            Call DrawRect(Pos, Size, Color)
        Case EShape.Oval
            Call DrawEllipse(Pos, Size, Color)
        Case EShape.RoundRect
            Call DrawRoundRect(Pos, Size, Radian, Color)
        End Select
End Sub
Public Sub FillShape(Shape As EShape, Pos As EPosition, Size As ESize, Radian As Long, Color As ColorMix)
    Select Case Shape
        Case EShape.Square
            Call FillRect(Pos, Size, Color)
        Case EShape.Oval
            Call FillEllipse(Pos, Size, Color)
        Case EShape.RoundRect
            Call FillRoundRect(Pos, Size, Radian, Color)
        End Select
End Sub
Public Sub ClipShape(Shape As EShape, Pos As EPosition, Size As ESize, Radian As Long)
    Select Case Shape
        Case EShape.Square
            Call ClipRect(Pos, Size)
        Case EShape.Oval
            Call ClipEllipse(Pos, Size)
        Case EShape.RoundRect
            Call ClipRoundRect(Pos, Size, Radian)
        End Select
End Sub
Public Sub ResetClip()
    GdipResetPath mClipPath
    GdipResetClip MGraphics
End Sub
Public Sub ClipEllipse(Pos As EPosition, Size As ESize)
    GdipAddPathEllipse mClipPath, Pos.x, Pos.y, Size.Width - 1, Size.Height - 1
    GdipResetClip MGraphics
    GdipSetClipPath MGraphics, mClipPath, CombineModeReplace
End Sub
Public Sub ClipRect(Pos As EPosition, Size As ESize)
    GdipAddPathRectangle mClipPath, Pos.x, Pos.y, Size.Width, Size.Height
    GdipResetClip MGraphics
    GdipSetClipPath MGraphics, mClipPath, CombineModeReplace
End Sub
Public Sub ClipRoundRect(Pos As EPosition, Size As ESize, Radian As Long)
    If Radian > Size.Width Then Radian = Size.Width
    If Radian > Size.Height Then Radian = Size.Height
    GdipAddPathArc mClipPath, Pos.x, Pos.y, Radian, Radian, 180, 90
    GdipAddPathArc mClipPath, Pos.x + Size.Width - Radian - 1, Pos.y, Radian, Radian, 270, 90
    GdipAddPathArc mClipPath, Pos.x + Size.Width - Radian - 1, Pos.y + Size.Height - Radian - 1, Radian, Radian, 0, 90
    GdipAddPathArc mClipPath, Pos.x, Pos.y + Size.Height - Radian - 1, Radian, Radian, 90, 90
    GdipClosePathFigure mClipPath
    GdipResetClip MGraphics
    GdipSetClipPath MGraphics, mClipPath, CombineModeReplace
End Sub
Public Sub FillEllipse(Pos As EPosition, Size As ESize, Color As ColorMix)
    GdipSetSmoothingMode MGraphics, SmoothingModeAntiAlias
    GdipAddPathEllipse mPath, Pos.x, Pos.y, Size.Width - 1, Size.Height - 1
    GdipFillPath MGraphics, Color.brush, mPath
    GdipResetPath mPath
End Sub
Public Sub FillRect(Pos As EPosition, Size As ESize, Color As ColorMix)
    GdipSetSmoothingMode MGraphics, SmoothingModeDefault
    GdipAddPathRectangle mPath, Pos.x, Pos.y, Size.Width, Size.Height
    GdipFillPath MGraphics, Color.brush, mPath
    GdipResetPath mPath
End Sub
Public Sub FillRoundRect(Pos As EPosition, Size As ESize, Radian As Long, Color As ColorMix)
    GdipSetSmoothingMode MGraphics, SmoothingModeAntiAlias
    If Radian > Size.Width Then Radian = Size.Width
    If Radian > Size.Height Then Radian = Size.Height
    GdipAddPathArc mPath, Pos.x, Pos.y, Radian, Radian, 180, 90
    GdipAddPathArc mPath, Pos.x + Size.Width - Radian - 1, Pos.y, Radian, Radian, 270, 90
    GdipAddPathArc mPath, Pos.x + Size.Width - Radian - 1, Pos.y + Size.Height - Radian - 1, Radian, Radian, 0, 90
    GdipAddPathArc mPath, Pos.x, Pos.y + Size.Height - Radian - 1, Radian, Radian, 90, 90
    GdipClosePathFigure mPath
    GdipFillPath MGraphics, Color.brush, mPath
    GdipResetPath mPath
End Sub
Public Sub FillArc(Pos As EPosition, Size As ESize, Start As Long, Angle As Long, Color As ColorMix)
    GdipSetSmoothingMode MGraphics, SmoothingModeAntiAlias
    GdipAddPathArc mPath, Pos.x, Pos.y, Size.Width - 1, Size.Height - 1, Start, Angle
    GdipClosePathFigure mPath
    GdipFillPath MGraphics, Color.brush, mPath
    GdipResetPath mPath
End Sub
Public Sub FillArc2(Pos As EPosition, Size As ESize, Start As Long, Angle As Long, Color As ColorMix)
    GdipSetSmoothingMode MGraphics, SmoothingModeAntiAlias
    GdipAddPathLine mPath, Size.Width / 2, Size.Height / 2, Size.Width / 2, 0
    GdipAddPathArc mPath, Pos.x, Pos.y, Size.Width - 1, Size.Height - 1, Start, Angle
    GdipClosePathFigure mPath
    GdipFillPath MGraphics, Color.brush, mPath
    GdipResetPath mPath
End Sub
Public Sub DrawLine(Pos As EPosition, Pos2 As EPosition, Color As ColorMix)
    GdipSetSmoothingMode MGraphics, SmoothingModeAntiAlias
    GdipAddPathLine mPath, Pos.x, Pos.y, Pos2.x, Pos2.y
    GdipDrawPath MGraphics, Color.pen, mPath
    GdipResetPath mPath
End Sub
Public Sub DrawEllipse(Pos As EPosition, Size As ESize, Color As ColorMix)
    GdipSetSmoothingMode MGraphics, SmoothingModeAntiAlias
    GdipAddPathEllipse mPath, Pos.x, Pos.y, Size.Width - Color.Width, Size.Height - Color.Width
    GdipDrawPath MGraphics, Color.pen, mPath
    GdipResetPath mPath
End Sub
Public Sub DrawRect(Pos As EPosition, Size As ESize, Color As ColorMix)
    GdipSetSmoothingMode MGraphics, SmoothingModeDefault
    GdipAddPathRectangle mPath, Pos.x, Pos.y, Size.Width - Color.Width, Size.Height - Color.Width
    GdipDrawPath MGraphics, Color.pen, mPath
    GdipResetPath mPath
End Sub
Public Sub DrawRoundRect(Pos As EPosition, Size As ESize, Radian As Long, Color As ColorMix)
    GdipSetSmoothingMode MGraphics, SmoothingModeAntiAlias
    If Radian > Size.Width Then Radian = Size.Width
    If Radian > Size.Height Then Radian = Size.Height
    GdipAddPathArc mPath, Pos.x, Pos.y, Radian, Radian, 180, 90
    GdipAddPathArc mPath, Pos.x + Size.Width - Radian - 1, Pos.y, Radian, Radian, 270, 90
    GdipAddPathArc mPath, Pos.x + Size.Width - Radian - 1, Pos.y + Size.Height - Radian - 1, Radian, Radian, 0, 90
    GdipAddPathArc mPath, Pos.x, Pos.y + Size.Height - Radian - 1, Radian, Radian, 90, 90
    GdipClosePathFigure mPath
    GdipDrawPath MGraphics, Color.pen, mPath
    GdipResetPath mPath
End Sub
Public Sub DrawArc(Pos As EPosition, Size As ESize, Start As Long, Angle As Long, Color As ColorMix)
    GdipSetSmoothingMode MGraphics, SmoothingModeAntiAlias
    GdipAddPathArc mPath, Pos.x, Pos.y, Size.Width - Color.Width, Size.Height - Color.Width, Start, Angle
    GdipDrawPath MGraphics, Color.pen, mPath
    GdipResetPath mPath
End Sub
Public Function GetStringSize(ByVal Text As String, Font As Fonter) As ESize
    Dim mRect As RECTF
    GdipMeasureString Graphics, StrPtr(Text), Len(Text), Font.Hwnd, NewRectF(0, 0, 0, 0), StrFormat(2), mRect, 0, 0
    GetStringSize.Width = mRect.Right: GetStringSize.Height = mRect.Bottom
End Function
Public Sub DrawImageRect(Pos As EPosition, Size As ESize, Img As EImage)
    GdipDrawImageRect MGraphics, Img.Hwnd, Pos.x, Pos.y, Size.Width, Size.Height
End Sub
Public Sub DrawImage(Pos As EPosition, Img As EImage)
    GdipDrawImage MGraphics, Img.Hwnd, Pos.x, Pos.y
End Sub
Public Sub DrawEImage(Pos As EPosition, Size As ESize, Align As EAlign, Img As EImage)
    If Img.Hwnd = 0 Then Exit Sub
    Dim Width As Single, Height As Single, DrawX As Single, DrawY As Single
    Width = IIf(Img.Size.Width = 0, Img.ImgSize.Width, Img.Size.Width)
    Height = IIf(Img.Size.Height = 0, Img.ImgSize.Height, Img.Size.Height)
    
    If Align.Horizontal = OnLeft Then DrawX = Pos.x
    If Align.Horizontal = OnCenter Then DrawX = Pos.x + Size.Width / 2 - Width / 2
    If Align.Horizontal = OnRight Then DrawX = Pos.x + Size.Width - Width
    
    If Align.Vertical = OnTop Then DrawY = Pos.y
    If Align.Vertical = OnMiddle Then DrawY = Pos.y + Size.Height / 2 - Height / 2
    If Align.Vertical = OnBottom Then DrawY = Pos.y + Size.Height - Height
    
    GdipDrawImageRect MGraphics, Img.Hwnd, DrawX, DrawY, Width, Height
End Sub
Public Sub DrawString(Pos As EPosition, Size As ESize, ByVal Text As String, Align As EAlign, Font As Fonter, Color As ColorMix)
    If Text = "" Then Exit Sub

    Dim DrawX As Long, DrawY As Long, StrFor As Long
    
    If Align.Horizontal = OnLeft Then DrawX = Pos.x: StrFor = StrFormat(0)
    If Align.Horizontal = OnCenter Then DrawX = Pos.x + Size.Width / 2: StrFor = StrFormat(1)
    If Align.Horizontal = OnRight Then DrawX = Pos.x + Size.Width: StrFor = StrFormat(2)
    
    If Align.Vertical = OnTop Then DrawY = Pos.y
    If Align.Vertical = OnMiddle Then DrawY = Pos.y + Size.Height / 2 - Font.Size / 0.75 / 2
    If Align.Vertical = OnBottom Then DrawY = Pos.y + Size.Height - GetStringSize(Text, Font).Height
    
    GdipDrawString MGraphics, StrPtr(Text), -1, Font.Hwnd, NewRectF(DrawX, DrawY, 0, 0), StrFor, Color.brush
    
End Sub
Public Sub DrawEditString(Pos As EPosition, Size As ESize, ByVal Text As String, Align As EAlign, Font As Fonter, Color As ColorMix)
    If Text = "" Then Exit Sub

    Dim DrawX As Long, DrawY As Long, StrFor As Long
    
    If Align.Horizontal = OnLeft Then DrawX = Pos.x: StrFor = EditStrFormat(0)
    If Align.Horizontal = OnCenter Then DrawX = Pos.x + Size.Width / 2: StrFor = EditStrFormat(1)
    If Align.Horizontal = OnRight Then DrawX = Pos.x + Size.Width: StrFor = EditStrFormat(2)
    
    If Align.Vertical = OnTop Then DrawY = Pos.y
    If Align.Vertical = OnMiddle Then DrawY = Pos.y + Size.Height / 2 - Font.Size / 0.75 / 2
    If Align.Vertical = OnBottom Then DrawY = Pos.y + Size.Height - GetStringSize(Text, Font).Height
    
    GdipDrawString MGraphics, StrPtr(Text), -1, Font.Hwnd, NewRectF(DrawX, DrawY, 0, 0), StrFor, Color.brush
    
End Sub
Private Sub Class_Initialize()
    If StartEasoft = False Then Err.Raise 4046, , "Blackmouth's love" & vbCrLf & "Start Easoft first ."
    GdipCreateBitmapFromScan0 mSize.Width, mSize.Height, ByVal 0, PixelFormat32bppARGB, ByVal 0, mImage
    GdipGetImageGraphicsContext mImage, MGraphics
    GdipSetSmoothingMode MGraphics, SmoothingModeAntiAlias
    GdipSetTextRenderingHint MGraphics, TextRenderingHintAntiAliasGridFit
    GdipCreatePath FillModeWinding, mPath
    GdipCreatePath FillModeWinding, mClipPath
    AddMember Me
End Sub
Public Sub Dispose()
    GdipDeleteGraphics MGraphics
    GdipDisposeImage mImage
    GdipDeletePath mPath
    GdipDeletePath mClipPath
End Sub
