Attribute VB_Name = "BAS_Bitmaps"

'##################################################################################
'##################################################################################
'##                                                                              ##
'## 3D Digital Reality Engine V1.00, Pure VB Code, By KACI Lounes 2009           ##
'##                                                                              ##
'##################################################################################
'##################################################################################

'##################################################################################
'##################################################################################
'###
'###  MODULE      : BAS_Bitmaps.BAS
'###
'###  DESCRIPTION : A useful, portable, mini image-software destined to treating
'###                2D surfaces (or BitMaps), support only 8 & 24 Bits formats,
'###                including:
'###
'###                Base functions: creation, displaying to DCs, deleting.
'###                Bltting operations
'###                Primary-shapes drawing: lines circle, ellipses...
'###                Color filtering: (gamma, contrast...)
'###                Reconstruction: or resampling functions
'###
'###                   And much more...
'###
'##################################################################################
'##################################################################################

Option Explicit

Public Declare Function SetPixel Lib "GDI32.DLL" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function SetDIBits Lib "GDI32.DLL" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

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

Public Type BITMAPINFO
 bmiHeader As BITMAPINFOHEADER
End Type

'Minimum & maximum bitmaps dimension authorized
Global Const MinBitMapWidth As Integer = 10
Global Const MinBitMapHeight As Integer = 10
Global Const MaxBitMapWidth As Integer = 700
Global Const MaxBitMapHeight As Integer = 700

'The resolution of primary shapes
Global Const SamplesDensity As Integer = 50
Function BitMap2D_DrawCircle(TheBitmap As BitMap2D, ACircle As Circle2D, Color As ColorRGB, AntiAlias As Boolean)

 '##################################################################################
 '##################################################################################
 '###
 '###  FUNCTION    : BiMap2D_DrawCircle
 '###
 '###  DESCRIPTION : Draw a 2D circle on a given bitmap.
 '###
 '##################################################################################
 '##################################################################################

 If (BitMap2D_IsValid(TheBitmap) = False) Then Exit Function
 If (ColorIsValid(Color) = False) Then Exit Function
 If (ACircle.Radius < 5) Then Exit Function

 Dim AX%, AY%, BX%, BY%, OX1%, OY1%, OX2%, OY2%, StepAngle!, CurSample%

 With ACircle
  StepAngle = (Pi / (SamplesDensity * 0.5))
  For CurSample = 2 To SamplesDensity
   AX = (.Center.X + (Sin(StepAngle * (CurSample)) * .Radius))
   AY = (.Center.Y + (Cos(StepAngle * (CurSample)) * .Radius))
   BX = (.Center.X + (Sin(StepAngle * (CurSample - 1)) * .Radius))
   BY = (.Center.Y + (Cos(StepAngle * (CurSample - 1)) * .Radius))
   If (ClipLine(0, 0, CSng(TheBitmap.Dimensions.X), CSng(TheBitmap.Dimensions.Y), AX, AY, BX, BY, OX1, OY1, OX2, OY2) = True) Then
    BitMap2D_DrawLine TheBitmap, OX1, OY1, OX2, OY2, Color, AntiAlias
   End If
  Next CurSample
  AX = (.Center.X + (Sin(StepAngle) * .Radius))
  AY = (.Center.Y + (Cos(StepAngle) * .Radius))
  BX = (.Center.X + (Sin(StepAngle * SamplesDensity) * .Radius))
  BY = (.Center.Y + (Cos(StepAngle * SamplesDensity) * .Radius))
  If (ClipLine(0, 0, CSng(TheBitmap.Dimensions.X), CSng(TheBitmap.Dimensions.Y), AX, AY, BX, BY, OX1, OY1, OX2, OY2) = True) Then
   BitMap2D_DrawLine TheBitmap, OX1, OY1, OX2, OY2, Color, AntiAlias
  End If
 End With

End Function
Function BitMap2D_DrawEllipse(TheBitmap As BitMap2D, AnEllipse As Ellipse2D, Color As ColorRGB, AntiAlias As Boolean)

 '##################################################################################
 '##################################################################################
 '###
 '###  FUNCTION    : BiMap2D_DrawEllipse
 '###
 '###  DESCRIPTION : Draw an 2D ellipse on a given bitmap.
 '###
 '##################################################################################
 '##################################################################################

 If (BitMap2D_IsValid(TheBitmap) = False) Then Exit Function
 If (ColorIsValid(Color) = False) Then Exit Function
 If (Rect2DIsValid(Rect2DInput(AnEllipse.MinPoint.X, AnEllipse.MinPoint.Y, AnEllipse.MaxPoint.X, AnEllipse.MaxPoint.Y)) = False) Then Exit Function
 If (AnEllipse.MaxPoint.X - AnEllipse.MinPoint.X < 5) Then Exit Function
 If (AnEllipse.MaxPoint.Y - AnEllipse.MinPoint.Y < 5) Then Exit Function

 Dim CX%, CY%, AX!, AY!, BX!, BY!, OX1%, OY1%, OX2%, OY2%, DiffX%, DiffY%
 Dim RAX%, RAY%, RBX%, RBY%, StepAngle!, Sinus!, Cosinus!, CurSample%

 With AnEllipse
  CX = (AnEllipse.MinPoint.X + ((AnEllipse.MaxPoint.X - AnEllipse.MinPoint.X) * 0.5))
  CY = (AnEllipse.MinPoint.Y + ((AnEllipse.MaxPoint.Y - AnEllipse.MinPoint.Y) * 0.5))
  DiffX = (AnEllipse.MaxPoint.X - CX): DiffY = (AnEllipse.MaxPoint.Y - CY)
  If (.Angle <> 0) Then Sinus = Sin(.Angle): Cosinus = Cos(.Angle)
  StepAngle = (Pi / (SamplesDensity * 0.5))
  For CurSample = 2 To SamplesDensity
   AX = (Sin(StepAngle * (CurSample)) * DiffX)
   AY = (Cos(StepAngle * (CurSample)) * DiffY)
   BX = (Sin(StepAngle * (CurSample - 1)) * DiffX)
   BY = (Cos(StepAngle * (CurSample - 1)) * DiffY)
   If (.Angle = 0) Then
    RAX = (CX + AX): RAY = (CY + AY): RBX = (CX + BX): RBY = (CY + BY)
   Else
    RAX = ((AX * Cosinus) - (AY * Sinus)): RAY = ((AX * Sinus) + (AY * Cosinus))
    RBX = ((BX * Cosinus) - (BY * Sinus)): RBY = ((BX * Sinus) + (BY * Cosinus))
    RAX = (CX + RAX): RAY = (CY + RAY): RBX = (CX + RBX): RBY = (CY + RBY)
   End If
   If (ClipLine(0, 0, CSng(TheBitmap.Dimensions.X), CSng(TheBitmap.Dimensions.Y), RAX, RAY, RBX, RBY, OX1, OY1, OX2, OY2) = True) Then
    BitMap2D_DrawLine TheBitmap, OX1, OY1, OX2, OY2, Color, AntiAlias
   End If
  Next CurSample
  AX = (CX + (Sin(StepAngle) * DiffX))
  AY = (CY + (Cos(StepAngle) * DiffY))
  BX = (CX + (Sin(StepAngle * SamplesDensity) * DiffX))
  BY = (CY + (Cos(StepAngle * SamplesDensity) * DiffY))
  If (.Angle = 0) Then
   RAX = (CX + AX): RAY = (CY + AY): RBX = (CX + BX): RBY = (CY + BY)
  Else
   RAX = ((AX * Cosinus) - (AY * Sinus)): RAY = ((AX * Sinus) + (AY * Cosinus))
   RBX = ((BX * Cosinus) - (BY * Sinus)): RBY = ((BX * Sinus) + (BY * Cosinus))
   RAX = (CX + RAX): RAY = (CY + RAY): RBX = (CX + RBX): RBY = (CY + RBY)
  End If
  If (ClipLine(0, 0, CSng(TheBitmap.Dimensions.X), CSng(TheBitmap.Dimensions.Y), RAX, RAY, RBX, RBY, OX1, OY1, OX2, OY2) = True) Then
   BitMap2D_DrawLine TheBitmap, OX1, OY1, OX2, OY2, Color, AntiAlias
  End If
 End With

End Function
Function BitMap2D_DrawRectangle(TheBitmap As BitMap2D, ARectangle As Rectangle2D, Color As ColorRGB, AntiAlias As Boolean)

 '##################################################################################
 '##################################################################################
 '###
 '###  FUNCTION    : BiMap2D_DrawRectangle
 '###
 '###  DESCRIPTION : Draw an 2D rectangle on a given bitmap.
 '###
 '##################################################################################
 '##################################################################################

 If (BitMap2D_IsValid(TheBitmap) = False) Then Exit Function
 If (ColorIsValid(Color) = False) Then Exit Function
 If (Rect2DIsValid(Rect2DInput(ARectangle.MinPoint.X, ARectangle.MinPoint.Y, ARectangle.MaxPoint.X, ARectangle.MaxPoint.Y)) = False) Then Exit Function

 Dim CtX%, CtY%, AX!, AY!, BX!, BY!, CX!, CY!, DX!, Dy!, OX1%, OY1%, OX2%, OY2%
 Dim RAX%, RAY%, RBX%, RBY%, RCX%, RCY%, RDX%, RDY%, Sinus!, Cosinus!

 With ARectangle
  CtX = (ARectangle.MinPoint.X + ((ARectangle.MaxPoint.X - ARectangle.MinPoint.X) * 0.5))
  CtY = (ARectangle.MinPoint.Y + ((ARectangle.MaxPoint.Y - ARectangle.MinPoint.Y) * 0.5))
  If (.Angle <> 0) Then Sinus = Sin(.Angle): Cosinus = Cos(.Angle)
  AX = (ARectangle.MinPoint.X - CtX): AY = (ARectangle.MinPoint.Y - CtY)
  BX = (ARectangle.MaxPoint.X - CtX): BY = (ARectangle.MinPoint.Y - CtY)
  CX = (ARectangle.MaxPoint.X - CtX): CY = (ARectangle.MaxPoint.Y - CtY)
  DX = (ARectangle.MinPoint.X - CtX): Dy = (ARectangle.MaxPoint.Y - CtY)
  If (.Angle = 0) Then
   RAX = (CtX + AX): RAY = (CtY + AY): RBX = (CtX + BX): RBY = (CtY + BY)
   RCX = (CtX + CX): RCY = (CtY + CY): RDX = (CtX + DX): RDY = (CtY + Dy)
  Else
   RAX = ((AX * Cosinus) - (AY * Sinus)): RAY = ((AX * Sinus) + (AY * Cosinus))
   RBX = ((BX * Cosinus) - (BY * Sinus)): RBY = ((BX * Sinus) + (BY * Cosinus))
   RCX = ((CX * Cosinus) - (CY * Sinus)): RCY = ((CX * Sinus) + (CY * Cosinus))
   RDX = ((DX * Cosinus) - (Dy * Sinus)): RDY = ((DX * Sinus) + (Dy * Cosinus))
   RAX = (CtX + RAX): RAY = (CtY + RAY): RBX = (CtX + RBX): RBY = (CtY + RBY)
   RCX = (CtX + RCX): RCY = (CtY + RCY): RDX = (CtX + RDX): RDY = (CtY + RDY)
  End If
  If (ClipLine(0, 0, CSng(TheBitmap.Dimensions.X), CSng(TheBitmap.Dimensions.Y), RAX, RAY, RBX, RBY, OX1, OY1, OX2, OY2) = True) Then
   BitMap2D_DrawLine TheBitmap, OX1, OY1, OX2, OY2, Color, AntiAlias
  End If
  If (ClipLine(0, 0, CSng(TheBitmap.Dimensions.X), CSng(TheBitmap.Dimensions.Y), RBX, RBY, RCX, RCY, OX1, OY1, OX2, OY2) = True) Then
   BitMap2D_DrawLine TheBitmap, OX1, OY1, OX2, OY2, Color, AntiAlias
  End If
  If (ClipLine(0, 0, CSng(TheBitmap.Dimensions.X), CSng(TheBitmap.Dimensions.Y), RCX, RCY, RDX, RDY, OX1, OY1, OX2, OY2) = True) Then
   BitMap2D_DrawLine TheBitmap, OX1, OY1, OX2, OY2, Color, AntiAlias
  End If
  If (ClipLine(0, 0, CSng(TheBitmap.Dimensions.X), CSng(TheBitmap.Dimensions.Y), RDX, RDY, RAX, RAY, OX1, OY1, OX2, OY2) = True) Then
   BitMap2D_DrawLine TheBitmap, OX1, OY1, OX2, OY2, Color, AntiAlias
  End If
 End With

End Function
Function BitMap2D_DrawPolyLine(TheBitmap As BitMap2D, APolyLine As PolyLine2D, Color As ColorRGB, AntiAlias As Boolean)

 '##################################################################################
 '##################################################################################
 '###
 '###  FUNCTION    : BiMap2D_PolyLine
 '###
 '###  DESCRIPTION : Draw an 2D polyline (a set of connected 2D lines) on a given bitmap.
 '###
 '##################################################################################
 '##################################################################################

 If (BitMap2D_IsValid(TheBitmap) = False) Then Exit Function
 If (ColorIsValid(Color) = False) Then Exit Function

 Dim AX%, AY%, BX%, BY%, CurPoint&

 For CurPoint = 1 To UBound(APolyLine.Points())
  AX = APolyLine.Points(CurPoint).X
  AY = APolyLine.Points(CurPoint).Y
  BX = APolyLine.Points(CurPoint - 1).X
  BY = APolyLine.Points(CurPoint - 1).Y
  BitMap2D_DrawLine TheBitmap, AX, AY, BX, BY, Color, AntiAlias
 Next CurPoint

End Function
Function BitMap2D_DrawBezier(TheBitmap As BitMap2D, ABezier As Bezier2D, Color As ColorRGB, AntiAlias As Boolean)

 '##################################################################################
 '##################################################################################
 '###
 '###  FUNCTION    : BiMap2D_DrawBezier
 '###
 '###  DESCRIPTION : Draw a 2D bezier on given bitmap.
 '###
 '##################################################################################
 '##################################################################################

 If (BitMap2D_IsValid(TheBitmap) = False) Then Exit Function
 If (ColorIsValid(Color) = False) Then Exit Function

 Dim AX%, AY%, BX%, BY%, OX1%, OY1%, OX2%, OY2%, A!, B!, TheTimer!, Density!

 With ABezier
  Density = (1 / SamplesDensity)
  For TheTimer = 0 To (1 - Density) Step Density
   A = TheTimer: B = (1 - TheTimer)
   'Cubic interpolation for A:
   AX = ((.SPoint.X * (B ^ 3)) + (.EPoint.X * (A ^ 3)) + (((.CPoint1.X * 3) * (B ^ 2)) * A)) + (((.CPoint2.X * 3) * (A ^ 2)) * B)
   AY = ((.SPoint.Y * (B ^ 3)) + (.EPoint.Y * (A ^ 3)) + (((.CPoint1.Y * 3) * (B ^ 2)) * A)) + (((.CPoint2.Y * 3) * (A ^ 2)) * B)
   A = (TheTimer + Density): B = (1 - (TheTimer + Density))
   'Cubic interpolation for B:
   BX = ((.SPoint.X * (B ^ 3)) + (.EPoint.X * (A ^ 3)) + (((.CPoint1.X * 3) * (B ^ 2)) * A)) + (((.CPoint2.X * 3) * (A ^ 2)) * B)
   BY = ((.SPoint.Y * (B ^ 3)) + (.EPoint.Y * (A ^ 3)) + (((.CPoint1.Y * 3) * (B ^ 2)) * A)) + (((.CPoint2.Y * 3) * (A ^ 2)) * B)
   If (ClipLine(0, 0, CSng(TheBitmap.Dimensions.X), CSng(TheBitmap.Dimensions.Y), AX, AY, BX, BY, OX1, OY1, OX2, OY2) = True) Then
    BitMap2D_DrawLine TheBitmap, OX1, OY1, OX2, OY2, Color, AntiAlias
   End If
  Next TheTimer
 End With

End Function
Function BitMap2D_DrawPolyBezier(TheBitmap As BitMap2D, APolyBezier As PolyBezier2D, Color As ColorRGB, AntiAlias As Boolean)

 '##################################################################################
 '##################################################################################
 '###
 '###  FUNCTION    : BiMap2D_DrawPolyBezier
 '###
 '###  DESCRIPTION : Draw an 2D polybezier (a set of connected 2D beziers)
 '###                on a given bitmap.
 '###
 '##################################################################################
 '##################################################################################

 If (BitMap2D_IsValid(TheBitmap) = False) Then Exit Function
 If (ColorIsValid(Color) = False) Then Exit Function
 If (UBound(APolyBezier.CPoints()) <> UBound(APolyBezier.Points())) Then Exit Function

 Dim CurPoint&, CurBezier As Bezier2D

 For CurPoint = 1 To UBound(APolyBezier.Points())
  CurBezier.SPoint.X = APolyBezier.Points(CurPoint).X
  CurBezier.SPoint.Y = APolyBezier.Points(CurPoint).Y
  CurBezier.EPoint.X = APolyBezier.Points(CurPoint - 1).X
  CurBezier.EPoint.Y = APolyBezier.Points(CurPoint - 1).Y
  CurBezier.CPoint1.X = APolyBezier.CPoints(CurPoint).X
  CurBezier.CPoint1.Y = APolyBezier.CPoints(CurPoint).Y
  CurBezier.CPoint2.X = APolyBezier.CPoints(CurPoint - 1).X
  CurBezier.CPoint2.Y = APolyBezier.CPoints(CurPoint - 1).Y
  BitMap2D_DrawBezier TheBitmap, CurBezier, Color, AntiAlias
 Next CurPoint

End Function
Function BitMap2D_DrawLine(TheBitmap As BitMap2D, X1%, Y1%, X2%, Y2%, Color As ColorRGB, AntiAlias As Boolean)

 '##################################################################################
 '##################################################################################
 '###
 '###  FUNCTION    : BiMap2D_DrawLine
 '###
 '###  DESCRIPTION : Draw an 2D line on a given bitmap, an anti-alias algorithm
 '###                (remove the pixalization effect on the line) is selected
 '###                as an option.
 '###
 '##################################################################################
 '##################################################################################

 On Error Resume Next

 If (AntiAlias = True) Then

  'An anti-alias line drawing algorithm
  '====================================

  Dim DltX%, DltY%, S%, E%, LoopC%, Fixed%, DX!, Dy!, DYDX!, Alph!

  DltX = (X2 - X1): If (DltX < 0) Then DltX = (0 - DltX)
  DltY = (Y2 - Y1): If (DltY < 0) Then DltY = (0 - DltY)

  If ((DltX <> 0) Or (DltY <> 0)) Then
   If (DltX > DltY) Then
    If (Y2 > Y1) Then DYDX = -(DltY / DltX) Else DYDX = (DltY / DltX)
    If (X2 < X1) Then S = X2: E = X1: Dy = Y2 Else S = X1: E = X2: Dy = Y1: DYDX = -DYDX
    For LoopC = S To E
     Fixed = (Dy - 0.5): Alph = (Dy - Fixed)
     TheBitmap.Datas(0, LoopC, Fixed) = ((TheBitmap.Datas(0, LoopC, Fixed) * Alph) + (Color.R * (1 - Alph)))
     TheBitmap.Datas(1, LoopC, Fixed) = ((TheBitmap.Datas(1, LoopC, Fixed) * Alph) + (Color.G * (1 - Alph)))
     TheBitmap.Datas(2, LoopC, Fixed) = ((TheBitmap.Datas(2, LoopC, Fixed) * Alph) + (Color.B * (1 - Alph)))
     TheBitmap.Datas(0, LoopC, (Fixed + 1)) = ((TheBitmap.Datas(0, LoopC, (Fixed + 1)) * (1 - Alph)) + (Color.R * Alph))
     TheBitmap.Datas(1, LoopC, (Fixed + 1)) = ((TheBitmap.Datas(1, LoopC, (Fixed + 1)) * (1 - Alph)) + (Color.G * Alph))
     TheBitmap.Datas(2, LoopC, (Fixed + 1)) = ((TheBitmap.Datas(2, LoopC, (Fixed + 1)) * (1 - Alph)) + (Color.B * Alph))
     Dy = (Dy + DYDX)
    Next LoopC
   Else
    If (X2 > X1) Then DYDX = -(DltX / DltY) Else DYDX = (DltX / DltY)
    If (Y2 < Y1) Then S = Y2: E = Y1: DX = X2 Else S = Y1: E = Y2: DX = X1: DYDX = -DYDX
    For LoopC = S To E
     Fixed = (DX - 0.5): Alph = (DX - Fixed)
     TheBitmap.Datas(0, Fixed, LoopC) = ((TheBitmap.Datas(0, Fixed, LoopC) * Alph) + (Color.R * (1 - Alph)))
     TheBitmap.Datas(1, Fixed, LoopC) = ((TheBitmap.Datas(1, Fixed, LoopC) * Alph) + (Color.G * (1 - Alph)))
     TheBitmap.Datas(2, Fixed, LoopC) = ((TheBitmap.Datas(2, Fixed, LoopC) * Alph) + (Color.B * (1 - Alph)))
     TheBitmap.Datas(0, (Fixed + 1), LoopC) = ((TheBitmap.Datas(0, (Fixed + 1), LoopC) * (1 - Alph)) + (Color.R * Alph))
     TheBitmap.Datas(1, (Fixed + 1), LoopC) = ((TheBitmap.Datas(1, (Fixed + 1), LoopC) * (1 - Alph)) + (Color.G * Alph))
     TheBitmap.Datas(2, (Fixed + 1), LoopC) = ((TheBitmap.Datas(2, (Fixed + 1), LoopC) * (1 - Alph)) + (Color.B * Alph))
     DX = (DX + DYDX)
    Next LoopC
   End If
  Else
   Exit Function
  End If

 ElseIf (AntiAlias = False) Then

  'An integers-based (Bresenham) line drawing algorithm
  '====================================================

  Dim DDX%, DDY%, SDX%, SDY%, X%, Y%, PX%, PY%

  DDX = (X2 - X1): If (DDX < 0) Then SDX = -1 Else SDX = 1
  DDY = (Y2 - Y1): If (DDY < 0) Then SDY = -1 Else SDY = 1

  DDX = ((SDX * DDX) + 1): DDY = ((SDY * DDY) + 1): PX = X1: PY = Y1

  If (DDX >= DDY) Then
   Do While (X < DDX)
    TheBitmap.Datas(0, PX, PY) = Color.R
    TheBitmap.Datas(1, PX, PY) = Color.G
    TheBitmap.Datas(2, PX, PY) = Color.B
    Y = (Y + DDY): If (Y >= DDX) Then Y = (Y - DDX): PY = (PY + SDY)
    X = (X + 1): PX = (PX + SDX)
   Loop
  Else
   Do While (Y < DDY)
    TheBitmap.Datas(0, PX, PY) = Color.R
    TheBitmap.Datas(1, PX, PY) = Color.G
    TheBitmap.Datas(2, PX, PY) = Color.B
    X = (X + DDX): If (X >= DDY) Then X = (X - DDY): PX = (PX + SDX)
    Y = (Y + 1): PY = (PY + SDY)
   Loop
  End If

 End If

End Function
Function BitMap2D_Blt(SrcBitMap As BitMap2D, DestBitMap As BitMap2D, SrcRect As Rect2D, DestRect As Rect2D, BltFlag As BltFlags)

 '##################################################################################
 '##################################################################################
 '###
 '###  FUNCTION    : BiMap2D_Blt
 '###
 '###  DESCRIPTION : A big Bit-bLoc-Transfer function, works at any combinison.
 '###
 '##################################################################################
 '##################################################################################

 If (BitMap2D_IsValid(SrcBitMap) = False) Then Exit Function
 If (BitMap2D_IsValid(DestBitMap) = False) Then Exit Function
 If (Rect2DIsInsideRegion(SrcRect, SrcBitMap.Dimensions) = False) Then Exit Function
 If (Rect2DIsInsideRegion(DestRect, DestBitMap.Dimensions) = False) Then Exit Function
 With BltFlag
  If ((.ChannelFrom < 0) Or (.ChannelFrom > 2)) Then Exit Function
  If ((.ChannelTo < 0) Or (.ChannelTo > 2)) Then Exit Function
  If ((.AlphaFlag < 1) Or (.AlphaFlag > 3)) Then Exit Function
  If (.AlphaFlag = 2) Then
   If (BitMap2D_IsValid(.AlphaMask) = False) Then Exit Function
   If (Rect2DIsInsideRegion(.AlphaRect, .AlphaMask.Dimensions) = False) Then Exit Function
   If ((.ChannelAlphaFrom < 0) Or (.ChannelAlphaFrom > 2)) Then Exit Function
   If (BltFlag.Stretch = False) Then
    If (SrcBitMap.Dimensions.X < BltFlag.AlphaMask.Dimensions.X) Then Exit Function
    If (SrcBitMap.Dimensions.Y < BltFlag.AlphaMask.Dimensions.Y) Then Exit Function
   End If
  End If
  If (.Transparent = True) Then
   If (ColorIsValid(.TransColor) = False) Then Exit Function
  End If
 End With

 Dim OperationCode As String

 'Make the operation code :

 If (BltFlag.Transparent = False) Then
  OperationCode = "T0,"
 ElseIf (BltFlag.Transparent = True) Then
  OperationCode = "T1,"
 End If

 If (BltFlag.Stretch = False) Then
  OperationCode = OperationCode & "S0,"
 ElseIf (BltFlag.Stretch = True) Then
  OperationCode = OperationCode & "S1,"
 End If

 If (SrcBitMap.BitsDepth = 8) Then
  OperationCode = OperationCode & "Src08,"
 ElseIf (SrcBitMap.BitsDepth = 24) Then
  OperationCode = OperationCode & "Src24,"
 End If

 If (DestBitMap.BitsDepth = 8) Then
  OperationCode = OperationCode & "Dst08"
 ElseIf (DestBitMap.BitsDepth = 24) Then
  OperationCode = OperationCode & "Dst24"
 End If

 Dim CurX&, CurY&, Filtered As ColorRGB
 Dim SrcX&, SrcY&, DstX&, DstY&, AlphaX&, AlphaY&
 Dim ParamX!, ParamY!, CurU!, CurV!, ACurU!, ACurV!
 Dim AlphaMin!, AlphaMax!, AlphaCol As ColorRGB

 If (BltFlag.AlphaFlag = 1) Then 'Disable alpha blending
  Select Case OperationCode
   Case "T0,S0,Src08,Dst08":
    If (Rect2DDiagonalLength(SrcRect) > Rect2DDiagonalLength(DestRect)) Then
     For CurY = DestRect.Y1 To DestRect.Y2
      SrcY = (SrcRect.Y1 + (CurY - DestRect.Y1)): DstY = CurY
      For CurX = DestRect.X1 To DestRect.X2
       SrcX = (SrcRect.X1 + (CurX - DestRect.X1)): DstX = CurX
       DestBitMap.Datas(0, DstX, DstY) = SrcBitMap.Datas(0, SrcX, SrcY)
      Next CurX
     Next CurY
    Else
     For CurY = SrcRect.Y1 To SrcRect.Y2
      SrcY = CurY: DstY = (DestRect.Y1 + (CurY - SrcRect.Y1))
      For CurX = SrcRect.X1 To SrcRect.X2
       SrcX = CurX: DstX = (DestRect.X1 + (CurX - SrcRect.X1))
       DestBitMap.Datas(0, DstX, DstY) = SrcBitMap.Datas(0, SrcX, SrcY)
      Next CurX
     Next CurY
    End If
   Case "T0,S0,Src08,Dst24":
    If (Rect2DDiagonalLength(SrcRect) > Rect2DDiagonalLength(DestRect)) Then
     For CurY = DestRect.Y1 To DestRect.Y2
      SrcY = (SrcRect.Y1 + (CurY - DestRect.Y1)): DstY = CurY
      For CurX = DestRect.X1 To DestRect.X2
       SrcX = (SrcRect.X1 + (CurX - DestRect.X1)): DstX = CurX
       DestBitMap.Datas(BltFlag.ChannelTo, DstX, DstY) = SrcBitMap.Datas(0, SrcX, SrcY)
      Next CurX
     Next CurY
    Else
     For CurY = SrcRect.Y1 To SrcRect.Y2
      SrcY = CurY: DstY = (DestRect.Y1 + (CurY - SrcRect.Y1))
      For CurX = SrcRect.X1 To SrcRect.X2
       SrcX = CurX: DstX = (DestRect.X1 + (CurX - SrcRect.X1))
       DestBitMap.Datas(BltFlag.ChannelTo, DstX, DstY) = SrcBitMap.Datas(0, SrcX, SrcY)
      Next CurX
     Next CurY
    End If
   Case "T0,S0,Src24,Dst08":
    If (Rect2DDiagonalLength(SrcRect) > Rect2DDiagonalLength(DestRect)) Then
     For CurY = DestRect.Y1 To DestRect.Y2
      SrcY = (SrcRect.Y1 + (CurY - DestRect.Y1)): DstY = CurY
      For CurX = DestRect.X1 To DestRect.X2
       SrcX = (SrcRect.X1 + (CurX - DestRect.X1)): DstX = CurX
       DestBitMap.Datas(0, DstX, DstY) = SrcBitMap.Datas(BltFlag.ChannelFrom, SrcX, SrcY)
      Next CurX
     Next CurY
    Else
     For CurY = SrcRect.Y1 To SrcRect.Y2
      SrcY = CurY: DstY = (DestRect.Y1 + (CurY - SrcRect.Y1))
      For CurX = SrcRect.X1 To SrcRect.X2
       SrcX = CurX: DstX = (DestRect.X1 + (CurX - SrcRect.X1))
       DestBitMap.Datas(0, DstX, DstY) = SrcBitMap.Datas(BltFlag.ChannelFrom, SrcX, SrcY)
      Next CurX
     Next CurY
    End If
   Case "T0,S0,Src24,Dst24":
    If (Rect2DDiagonalLength(SrcRect) > Rect2DDiagonalLength(DestRect)) Then
     For CurY = DestRect.Y1 To DestRect.Y2
      SrcY = (SrcRect.Y1 + (CurY - DestRect.Y1)): DstY = CurY
      For CurX = DestRect.X1 To DestRect.X2
       SrcX = (SrcRect.X1 + (CurX - DestRect.X1)): DstX = CurX
       DestBitMap.Datas(0, DstX, DstY) = SrcBitMap.Datas(0, SrcX, SrcY)
       DestBitMap.Datas(1, DstX, DstY) = SrcBitMap.Datas(1, SrcX, SrcY)
       DestBitMap.Datas(2, DstX, DstY) = SrcBitMap.Datas(2, SrcX, SrcY)
      Next CurX
     Next CurY
    Else
     For CurY = SrcRect.Y1 To SrcRect.Y2
      SrcY = CurY: DstY = (DestRect.Y1 + (CurY - SrcRect.Y1))
      For CurX = SrcRect.X1 To SrcRect.X2
       SrcX = CurX: DstX = (DestRect.X1 + (CurX - SrcRect.X1))
       DestBitMap.Datas(0, DstX, DstY) = SrcBitMap.Datas(0, SrcX, SrcY)
       DestBitMap.Datas(1, DstX, DstY) = SrcBitMap.Datas(1, SrcX, SrcY)
       DestBitMap.Datas(2, DstX, DstY) = SrcBitMap.Datas(2, SrcX, SrcY)
      Next CurX
     Next CurY
    End If
   Case "T0,S1,Src08,Dst08":
    For CurY = DestRect.Y1 To DestRect.Y2
     ParamY = ((CurY - DestRect.Y1) / (DestRect.Y2 - DestRect.Y1))
     For CurX = DestRect.X1 To DestRect.X2
      ParamX = ((CurX - DestRect.X1) / (DestRect.X2 - DestRect.X1))
      CurU = (SrcRect.X1 + ((SrcRect.X2 - SrcRect.X1) * ParamX))
      CurV = (SrcRect.Y1 + ((SrcRect.Y2 - SrcRect.Y1) * ParamY))
      DestBitMap.Datas(0, CurX, CurY) = DoTexelFiltering(BltFlag.PixelFilter, SrcBitMap, CurU, CurV, False).R
     Next CurX
    Next CurY
   Case "T0,S1,Src08,Dst24":
    For CurY = DestRect.Y1 To DestRect.Y2
     ParamY = ((CurY - DestRect.Y1) / (DestRect.Y2 - DestRect.Y1))
     For CurX = DestRect.X1 To DestRect.X2
      ParamX = ((CurX - DestRect.X1) / (DestRect.X2 - DestRect.X1))
      CurU = (SrcRect.X1 + ((SrcRect.X2 - SrcRect.X1) * ParamX))
      CurV = (SrcRect.Y1 + ((SrcRect.Y2 - SrcRect.Y1) * ParamY))
      DestBitMap.Datas(BltFlag.ChannelTo, CurX, CurY) = DoTexelFiltering(BltFlag.PixelFilter, SrcBitMap, CurU, CurV, False).R
     Next CurX
    Next CurY
   Case "T0,S1,Src24,Dst08":
    For CurY = DestRect.Y1 To DestRect.Y2
     ParamY = ((CurY - DestRect.Y1) / (DestRect.Y2 - DestRect.Y1))
     For CurX = DestRect.X1 To DestRect.X2
      ParamX = ((CurX - DestRect.X1) / (DestRect.X2 - DestRect.X1))
      CurU = (SrcRect.X1 + ((SrcRect.X2 - SrcRect.X1) * ParamX))
      CurV = (SrcRect.Y1 + ((SrcRect.Y2 - SrcRect.Y1) * ParamY))
      If (BltFlag.ChannelFrom = 0) Then
       DestBitMap.Datas(0, CurX, CurY) = DoTexelFiltering(BltFlag.PixelFilter, SrcBitMap, CurU, CurV, False).R
      ElseIf (BltFlag.ChannelFrom = 1) Then
       DestBitMap.Datas(0, CurX, CurY) = DoTexelFiltering(BltFlag.PixelFilter, SrcBitMap, CurU, CurV, False).G
      ElseIf (BltFlag.ChannelFrom = 2) Then
       DestBitMap.Datas(0, CurX, CurY) = DoTexelFiltering(BltFlag.PixelFilter, SrcBitMap, CurU, CurV, False).B
      End If
     Next CurX
    Next CurY
   Case "T0,S1,Src24,Dst24":
    For CurY = DestRect.Y1 To DestRect.Y2
     ParamY = ((CurY - DestRect.Y1) / (DestRect.Y2 - DestRect.Y1))
     For CurX = DestRect.X1 To DestRect.X2
      ParamX = ((CurX - DestRect.X1) / (DestRect.X2 - DestRect.X1))
      CurU = (SrcRect.X1 + ((SrcRect.X2 - SrcRect.X1) * ParamX))
      CurV = (SrcRect.Y1 + ((SrcRect.Y2 - SrcRect.Y1) * ParamY))
      Filtered = DoTexelFiltering(BltFlag.PixelFilter, SrcBitMap, CurU, CurV, False)
      DestBitMap.Datas(0, CurX, CurY) = Filtered.R
      DestBitMap.Datas(1, CurX, CurY) = Filtered.G
      DestBitMap.Datas(2, CurX, CurY) = Filtered.B
     Next CurX
    Next CurY
   Case "T1,S0,Src08,Dst08":
    If (Rect2DDiagonalLength(SrcRect) > Rect2DDiagonalLength(DestRect)) Then
     For CurY = DestRect.Y1 To DestRect.Y2
      SrcY = (SrcRect.Y1 + (CurY - DestRect.Y1)): DstY = CurY
      For CurX = DestRect.X1 To DestRect.X2
       SrcX = (SrcRect.X1 + (CurX - DestRect.X1)): DstX = CurX
       If (SrcBitMap.Datas(0, SrcX, SrcY) <> BltFlag.TransColor.R) Then
        DestBitMap.Datas(0, DstX, DstY) = SrcBitMap.Datas(0, SrcX, SrcY)
       End If
      Next CurX
     Next CurY
    Else
     For CurY = SrcRect.Y1 To SrcRect.Y2
      SrcY = CurY: DstY = (DestRect.Y1 + (CurY - SrcRect.Y1))
      For CurX = SrcRect.X1 To SrcRect.X2
       SrcX = CurX: DstX = (DestRect.X1 + (CurX - SrcRect.X1))
       If (SrcBitMap.Datas(0, SrcX, SrcY) <> BltFlag.TransColor.R) Then
        DestBitMap.Datas(0, DstX, DstY) = SrcBitMap.Datas(0, SrcX, SrcY)
       End If
      Next CurX
     Next CurY
    End If
   Case "T1,S0,Src08,Dst24":
    If (Rect2DDiagonalLength(SrcRect) > Rect2DDiagonalLength(DestRect)) Then
     For CurY = DestRect.Y1 To DestRect.Y2
      SrcY = (SrcRect.Y1 + (CurY - DestRect.Y1)): DstY = CurY
      For CurX = DestRect.X1 To DestRect.X2
       SrcX = (SrcRect.X1 + (CurX - DestRect.X1)): DstX = CurX
       If (SrcBitMap.Datas(0, SrcX, SrcY) <> BltFlag.TransColor.R) Then
        DestBitMap.Datas(BltFlag.ChannelTo, DstX, DstY) = SrcBitMap.Datas(0, SrcX, SrcY)
       End If
      Next CurX
     Next CurY
    Else
     For CurY = SrcRect.Y1 To SrcRect.Y2
      SrcY = CurY: DstY = (DestRect.Y1 + (CurY - SrcRect.Y1))
      For CurX = SrcRect.X1 To SrcRect.X2
       SrcX = CurX: DstX = (DestRect.X1 + (CurX - SrcRect.X1))
       If (SrcBitMap.Datas(0, SrcX, SrcY) <> BltFlag.TransColor.R) Then
        DestBitMap.Datas(BltFlag.ChannelTo, DstX, DstY) = SrcBitMap.Datas(0, SrcX, SrcY)
       End If
      Next CurX
     Next CurY
    End If
   Case "T1,S0,Src24,Dst08":
    If (Rect2DDiagonalLength(SrcRect) > Rect2DDiagonalLength(DestRect)) Then
     For CurY = DestRect.Y1 To DestRect.Y2
      SrcY = (SrcRect.Y1 + (CurY - DestRect.Y1)): DstY = CurY
      For CurX = DestRect.X1 To DestRect.X2
       SrcX = (SrcRect.X1 + (CurX - DestRect.X1)): DstX = CurX
       If (SrcBitMap.Datas(0, SrcX, SrcY) <> BltFlag.TransColor.R) Then
        DestBitMap.Datas(0, DstX, DstY) = SrcBitMap.Datas(BltFlag.ChannelFrom, SrcX, SrcY)
       End If
      Next CurX
     Next CurY
    Else
     For CurY = SrcRect.Y1 To SrcRect.Y2
      SrcY = CurY: DstY = (DestRect.Y1 + (CurY - SrcRect.Y1))
      For CurX = SrcRect.X1 To SrcRect.X2
       SrcX = CurX: DstX = (DestRect.X1 + (CurX - SrcRect.X1))
       If (SrcBitMap.Datas(0, SrcX, SrcY) <> BltFlag.TransColor.R) Then
        DestBitMap.Datas(0, DstX, DstY) = SrcBitMap.Datas(BltFlag.ChannelFrom, SrcX, SrcY)
       End If
      Next CurX
     Next CurY
    End If
   Case "T1,S0,Src24,Dst24":
    If (Rect2DDiagonalLength(SrcRect) > Rect2DDiagonalLength(DestRect)) Then
     For CurY = DestRect.Y1 To DestRect.Y2
      SrcY = (SrcRect.Y1 + (CurY - DestRect.Y1)): DstY = CurY
      For CurX = DestRect.X1 To DestRect.X2
       SrcX = (SrcRect.X1 + (CurX - DestRect.X1)): DstX = CurX
       If (SrcBitMap.Datas(0, SrcX, SrcY) <> BltFlag.TransColor.R) Then
        DestBitMap.Datas(0, DstX, DstY) = SrcBitMap.Datas(0, SrcX, SrcY)
        DestBitMap.Datas(1, DstX, DstY) = SrcBitMap.Datas(1, SrcX, SrcY)
        DestBitMap.Datas(2, DstX, DstY) = SrcBitMap.Datas(2, SrcX, SrcY)
       End If
      Next CurX
     Next CurY
    Else
     For CurY = SrcRect.Y1 To SrcRect.Y2
      SrcY = CurY: DstY = (DestRect.Y1 + (CurY - SrcRect.Y1))
      For CurX = SrcRect.X1 To SrcRect.X2
       SrcX = CurX: DstX = (DestRect.X1 + (CurX - SrcRect.X1))
       If (SrcBitMap.Datas(0, SrcX, SrcY) <> BltFlag.TransColor.R) Then
        DestBitMap.Datas(0, DstX, DstY) = SrcBitMap.Datas(0, SrcX, SrcY)
        DestBitMap.Datas(1, DstX, DstY) = SrcBitMap.Datas(1, SrcX, SrcY)
        DestBitMap.Datas(2, DstX, DstY) = SrcBitMap.Datas(2, SrcX, SrcY)
       End If
      Next CurX
     Next CurY
    End If
   Case "T1,S1,Src08,Dst08":
    For CurY = DestRect.Y1 To DestRect.Y2
     ParamY = ((CurY - DestRect.Y1) / (DestRect.Y2 - DestRect.Y1))
     For CurX = DestRect.X1 To DestRect.X2
      ParamX = ((CurX - DestRect.X1) / (DestRect.X2 - DestRect.X1))
      CurU = (SrcRect.X1 + ((SrcRect.X2 - SrcRect.X1) * ParamX))
      CurV = (SrcRect.Y1 + ((SrcRect.Y2 - SrcRect.Y1) * ParamY))
      Filtered = DoTexelFiltering(BltFlag.PixelFilter, SrcBitMap, CurU, CurV, False)
      If (Filtered.R <> BltFlag.TransColor.R) Then DestBitMap.Datas(0, CurX, CurY) = Filtered.R
     Next CurX
    Next CurY
   Case "T1,S1,Src08,Dst24":
    For CurY = DestRect.Y1 To DestRect.Y2
     ParamY = ((CurY - DestRect.Y1) / (DestRect.Y2 - DestRect.Y1))
     For CurX = DestRect.X1 To DestRect.X2
      ParamX = ((CurX - DestRect.X1) / (DestRect.X2 - DestRect.X1))
      CurU = (SrcRect.X1 + ((SrcRect.X2 - SrcRect.X1) * ParamX))
      CurV = (SrcRect.Y1 + ((SrcRect.Y2 - SrcRect.Y1) * ParamY))
      Filtered = DoTexelFiltering(BltFlag.PixelFilter, SrcBitMap, CurU, CurV, False)
      If (Filtered.R <> BltFlag.TransColor.R) Then DestBitMap.Datas(BltFlag.ChannelTo, CurX, CurY) = Filtered.R
     Next CurX
    Next CurY
   Case "T1,S1,Src24,Dst08":
    For CurY = DestRect.Y1 To DestRect.Y2
     ParamY = ((CurY - DestRect.Y1) / (DestRect.Y2 - DestRect.Y1))
     For CurX = DestRect.X1 To DestRect.X2
      ParamX = ((CurX - DestRect.X1) / (DestRect.X2 - DestRect.X1))
      CurU = (SrcRect.X1 + ((SrcRect.X2 - SrcRect.X1) * ParamX))
      CurV = (SrcRect.Y1 + ((SrcRect.Y2 - SrcRect.Y1) * ParamY))
      Filtered = DoTexelFiltering(BltFlag.PixelFilter, SrcBitMap, CurU, CurV, False)
      If (BltFlag.ChannelFrom = 0) Then
       If (Filtered.R <> BltFlag.TransColor.R) Then DestBitMap.Datas(0, CurX, CurY) = Filtered.R
      ElseIf (BltFlag.ChannelFrom = 1) Then
       If (Filtered.G <> BltFlag.TransColor.G) Then DestBitMap.Datas(1, CurX, CurY) = Filtered.G
      ElseIf (BltFlag.ChannelFrom = 2) Then
       If (Filtered.B <> BltFlag.TransColor.B) Then DestBitMap.Datas(2, CurX, CurY) = Filtered.B
      End If
     Next CurX
    Next CurY
   Case "T1,S1,Src24,Dst24":
    For CurY = DestRect.Y1 To DestRect.Y2
     ParamY = ((CurY - DestRect.Y1) / (DestRect.Y2 - DestRect.Y1))
     For CurX = DestRect.X1 To DestRect.X2
      ParamX = ((CurX - DestRect.X1) / (DestRect.X2 - DestRect.X1))
      CurU = (SrcRect.X1 + ((SrcRect.X2 - SrcRect.X1) * ParamX))
      CurV = (SrcRect.Y1 + ((SrcRect.Y2 - SrcRect.Y1) * ParamY))
      Filtered = DoTexelFiltering(BltFlag.PixelFilter, SrcBitMap, CurU, CurV, False)
      If (ColorCompare(Filtered, BltFlag.TransColor) = False) Then
       DestBitMap.Datas(0, CurX, CurY) = Filtered.R
       DestBitMap.Datas(1, CurX, CurY) = Filtered.G
       DestBitMap.Datas(2, CurX, CurY) = Filtered.B
      End If
     Next CurX
    Next CurY
  End Select

 ElseIf (BltFlag.AlphaFlag = 2) Then 'Constant val alpha blending ////////
  AlphaMin = (1 / BltFlag.AlphaValueRed): AlphaMax = (1 - AlphaMin)
  Select Case OperationCode
   Case "T0,S0,Src08,Dst08":
    If (Rect2DDiagonalLength(SrcRect) > Rect2DDiagonalLength(DestRect)) Then
     For CurY = DestRect.Y1 To DestRect.Y2
      SrcY = (SrcRect.Y1 + (CurY - DestRect.Y1)): DstY = CurY
      For CurX = DestRect.X1 To DestRect.X2
       SrcX = (SrcRect.X1 + (CurX - DestRect.X1)): DstX = CurX
       DestBitMap.Datas(0, DstX, DstY) = ((DestBitMap.Datas(0, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(0, SrcX, SrcY) * AlphaMax))
      Next CurX
     Next CurY
    Else
     For CurY = SrcRect.Y1 To SrcRect.Y2
      SrcY = CurY: DstY = (DestRect.Y1 + (CurY - SrcRect.Y1))
      For CurX = SrcRect.X1 To SrcRect.X2
       SrcX = CurX: DstX = (DestRect.X1 + (CurX - SrcRect.X1))
       DestBitMap.Datas(0, DstX, DstY) = ((DestBitMap.Datas(0, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(0, SrcX, SrcY) * AlphaMax))
      Next CurX
     Next CurY
    End If
   Case "T0,S0,Src08,Dst24":
    If (Rect2DDiagonalLength(SrcRect) > Rect2DDiagonalLength(DestRect)) Then
     For CurY = DestRect.Y1 To DestRect.Y2
      SrcY = (SrcRect.Y1 + (CurY - DestRect.Y1)): DstY = CurY
      For CurX = DestRect.X1 To DestRect.X2
       SrcX = (SrcRect.X1 + (CurX - DestRect.X1)): DstX = CurX
       DestBitMap.Datas(BltFlag.ChannelTo, DstX, DstY) = ((DestBitMap.Datas(BltFlag.ChannelTo, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(0, SrcX, SrcY) * AlphaMax))
      Next CurX
     Next CurY
    Else
     For CurY = SrcRect.Y1 To SrcRect.Y2
      SrcY = CurY: DstY = (DestRect.Y1 + (CurY - SrcRect.Y1))
      For CurX = SrcRect.X1 To SrcRect.X2
       SrcX = CurX: DstX = (DestRect.X1 + (CurX - SrcRect.X1))
       DestBitMap.Datas(BltFlag.ChannelTo, DstX, DstY) = ((DestBitMap.Datas(BltFlag.ChannelTo, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(0, SrcX, SrcY) * AlphaMax))
      Next CurX
     Next CurY
    End If
   Case "T0,S0,Src24,Dst08":
    If (Rect2DDiagonalLength(SrcRect) > Rect2DDiagonalLength(DestRect)) Then
     For CurY = DestRect.Y1 To DestRect.Y2
      SrcY = (SrcRect.Y1 + (CurY - DestRect.Y1)): DstY = CurY
      For CurX = DestRect.X1 To DestRect.X2
       SrcX = (SrcRect.X1 + (CurX - DestRect.X1)): DstX = CurX
       DestBitMap.Datas(0, DstX, DstY) = ((DestBitMap.Datas(0, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(BltFlag.ChannelFrom, SrcX, SrcY) * AlphaMax))
      Next CurX
     Next CurY
    Else
     For CurY = SrcRect.Y1 To SrcRect.Y2
      SrcY = CurY: DstY = (DestRect.Y1 + (CurY - SrcRect.Y1))
      For CurX = SrcRect.X1 To SrcRect.X2
       SrcX = CurX: DstX = (DestRect.X1 + (CurX - SrcRect.X1))
       DestBitMap.Datas(0, DstX, DstY) = ((DestBitMap.Datas(0, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(BltFlag.ChannelFrom, SrcX, SrcY) * AlphaMax))
      Next CurX
     Next CurY
    End If
   Case "T0,S0,Src24,Dst24":
    If (Rect2DDiagonalLength(SrcRect) > Rect2DDiagonalLength(DestRect)) Then
     For CurY = DestRect.Y1 To DestRect.Y2
      SrcY = (SrcRect.Y1 + (CurY - DestRect.Y1)): DstY = CurY
      For CurX = DestRect.X1 To DestRect.X2
       SrcX = (SrcRect.X1 + (CurX - DestRect.X1)): DstX = CurX
       DestBitMap.Datas(0, DstX, DstY) = ((DestBitMap.Datas(0, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(0, SrcX, SrcY) * AlphaMax))
       DestBitMap.Datas(1, DstX, DstY) = ((DestBitMap.Datas(1, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(1, SrcX, SrcY) * AlphaMax))
       DestBitMap.Datas(2, DstX, DstY) = ((DestBitMap.Datas(2, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(2, SrcX, SrcY) * AlphaMax))
      Next CurX
     Next CurY
    Else
     For CurY = SrcRect.Y1 To SrcRect.Y2
      SrcY = CurY: DstY = (DestRect.Y1 + (CurY - SrcRect.Y1))
      For CurX = SrcRect.X1 To SrcRect.X2
       SrcX = CurX: DstX = (DestRect.X1 + (CurX - SrcRect.X1))
       DestBitMap.Datas(0, DstX, DstY) = ((DestBitMap.Datas(0, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(0, SrcX, SrcY) * AlphaMax))
       DestBitMap.Datas(1, DstX, DstY) = ((DestBitMap.Datas(1, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(1, SrcX, SrcY) * AlphaMax))
       DestBitMap.Datas(2, DstX, DstY) = ((DestBitMap.Datas(2, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(2, SrcX, SrcY) * AlphaMax))
      Next CurX
     Next CurY
    End If
   Case "T0,S1,Src08,Dst08":
    For CurY = DestRect.Y1 To DestRect.Y2
     ParamY = ((CurY - DestRect.Y1) / (DestRect.Y2 - DestRect.Y1))
     For CurX = DestRect.X1 To DestRect.X2
      ParamX = ((CurX - DestRect.X1) / (DestRect.X2 - DestRect.X1))
      CurU = (SrcRect.X1 + ((SrcRect.X2 - SrcRect.X1) * ParamX))
      CurV = (SrcRect.Y1 + ((SrcRect.Y2 - SrcRect.Y1) * ParamY))
      DestBitMap.Datas(0, CurX, CurY) = ((DestBitMap.Datas(0, CurX, CurY) * AlphaMin) + (DoTexelFiltering(BltFlag.PixelFilter, SrcBitMap, CurU, CurV, False).R * AlphaMax))
     Next CurX
    Next CurY
   Case "T0,S1,Src08,Dst24":
    For CurY = DestRect.Y1 To DestRect.Y2
     ParamY = ((CurY - DestRect.Y1) / (DestRect.Y2 - DestRect.Y1))
     For CurX = DestRect.X1 To DestRect.X2
      ParamX = ((CurX - DestRect.X1) / (DestRect.X2 - DestRect.X1))
      CurU = (SrcRect.X1 + ((SrcRect.X2 - SrcRect.X1) * ParamX))
      CurV = (SrcRect.Y1 + ((SrcRect.Y2 - SrcRect.Y1) * ParamY))
      DestBitMap.Datas(BltFlag.ChannelTo, CurX, CurY) = ((DestBitMap.Datas(BltFlag.ChannelTo, CurX, CurY) * AlphaMin) + (DoTexelFiltering(BltFlag.PixelFilter, SrcBitMap, CurU, CurV, False).R * AlphaMax))
     Next CurX
    Next CurY
   Case "T0,S1,Src24,Dst08":
    For CurY = DestRect.Y1 To DestRect.Y2
     ParamY = ((CurY - DestRect.Y1) / (DestRect.Y2 - DestRect.Y1))
     For CurX = DestRect.X1 To DestRect.X2
      ParamX = ((CurX - DestRect.X1) / (DestRect.X2 - DestRect.X1))
      CurU = (SrcRect.X1 + ((SrcRect.X2 - SrcRect.X1) * ParamX))
      CurV = (SrcRect.Y1 + ((SrcRect.Y2 - SrcRect.Y1) * ParamY))
      If (BltFlag.ChannelFrom = 0) Then
       DestBitMap.Datas(0, CurX, CurY) = ((DestBitMap.Datas(0, CurX, CurY) * AlphaMin) + (DoTexelFiltering(BltFlag.PixelFilter, SrcBitMap, CurU, CurV, False).R * AlphaMax))
      ElseIf (BltFlag.ChannelFrom = 1) Then
       DestBitMap.Datas(0, CurX, CurY) = ((DestBitMap.Datas(0, CurX, CurY) * AlphaMin) + (DoTexelFiltering(BltFlag.PixelFilter, SrcBitMap, CurU, CurV, False).G * AlphaMax))
      ElseIf (BltFlag.ChannelFrom = 2) Then
       DestBitMap.Datas(0, CurX, CurY) = ((DestBitMap.Datas(0, CurX, CurY) * AlphaMin) + (DoTexelFiltering(BltFlag.PixelFilter, SrcBitMap, CurU, CurV, False).B * AlphaMax))
      End If
     Next CurX
    Next CurY
   Case "T0,S1,Src24,Dst24":
    For CurY = DestRect.Y1 To DestRect.Y2
     ParamY = ((CurY - DestRect.Y1) / (DestRect.Y2 - DestRect.Y1))
     For CurX = DestRect.X1 To DestRect.X2
      ParamX = ((CurX - DestRect.X1) / (DestRect.X2 - DestRect.X1))
      CurU = (SrcRect.X1 + ((SrcRect.X2 - SrcRect.X1) * ParamX))
      CurV = (SrcRect.Y1 + ((SrcRect.Y2 - SrcRect.Y1) * ParamY))
      Filtered = DoTexelFiltering(BltFlag.PixelFilter, SrcBitMap, CurU, CurV, False)
      DestBitMap.Datas(0, CurX, CurY) = ((DestBitMap.Datas(0, CurX, CurY) * AlphaMin) + (Filtered.R * AlphaMax))
      DestBitMap.Datas(1, CurX, CurY) = ((DestBitMap.Datas(1, CurX, CurY) * AlphaMin) + (Filtered.G * AlphaMax))
      DestBitMap.Datas(2, CurX, CurY) = ((DestBitMap.Datas(2, CurX, CurY) * AlphaMin) + (Filtered.B * AlphaMax))
     Next CurX
    Next CurY
   Case "T1,S0,Src08,Dst08":
    If (Rect2DDiagonalLength(SrcRect) > Rect2DDiagonalLength(DestRect)) Then
     For CurY = DestRect.Y1 To DestRect.Y2
      SrcY = (SrcRect.Y1 + (CurY - DestRect.Y1)): DstY = CurY
      For CurX = DestRect.X1 To DestRect.X2
       SrcX = (SrcRect.X1 + (CurX - DestRect.X1)): DstX = CurX
       Filtered.R = ((DestBitMap.Datas(0, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(0, SrcX, SrcY) * AlphaMax))
       If (Filtered.R <> BltFlag.TransColor.R) Then
        DestBitMap.Datas(0, DstX, DstY) = Filtered.R
       End If
      Next CurX
     Next CurY
    Else
     For CurY = SrcRect.Y1 To SrcRect.Y2
      SrcY = CurY: DstY = (DestRect.Y1 + (CurY - SrcRect.Y1))
      For CurX = SrcRect.X1 To SrcRect.X2
       SrcX = CurX: DstX = (DestRect.X1 + (CurX - SrcRect.X1))
       Filtered.R = ((DestBitMap.Datas(0, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(0, SrcX, SrcY) * AlphaMax))
       If (Filtered.R <> BltFlag.TransColor.R) Then
        DestBitMap.Datas(0, DstX, DstY) = Filtered.R
       End If
      Next CurX
     Next CurY
    End If
   Case "T1,S0,Src08,Dst24":
    If (Rect2DDiagonalLength(SrcRect) > Rect2DDiagonalLength(DestRect)) Then
     For CurY = DestRect.Y1 To DestRect.Y2
      SrcY = (SrcRect.Y1 + (CurY - DestRect.Y1)): DstY = CurY
      For CurX = DestRect.X1 To DestRect.X2
       SrcX = (SrcRect.X1 + (CurX - DestRect.X1)): DstX = CurX
       Filtered.R = ((DestBitMap.Datas(BltFlag.ChannelTo, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(0, SrcX, SrcY) * AlphaMax))
       If (Filtered.R <> BltFlag.TransColor.R) Then
        DestBitMap.Datas(BltFlag.ChannelTo, DstX, DstY) = Filtered.R
       End If
      Next CurX
     Next CurY
    Else
     For CurY = SrcRect.Y1 To SrcRect.Y2
      SrcY = CurY: DstY = (DestRect.Y1 + (CurY - SrcRect.Y1))
      For CurX = SrcRect.X1 To SrcRect.X2
       SrcX = CurX: DstX = (DestRect.X1 + (CurX - SrcRect.X1))
       Filtered.R = ((DestBitMap.Datas(BltFlag.ChannelTo, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(0, SrcX, SrcY) * AlphaMax))
       If (Filtered.R <> BltFlag.TransColor.R) Then
        DestBitMap.Datas(BltFlag.ChannelTo, DstX, DstY) = Filtered.R
       End If
      Next CurX
     Next CurY
    End If
   Case "T1,S0,Src24,Dst08":
    If (Rect2DDiagonalLength(SrcRect) > Rect2DDiagonalLength(DestRect)) Then
     For CurY = DestRect.Y1 To DestRect.Y2
      SrcY = (SrcRect.Y1 + (CurY - DestRect.Y1)): DstY = CurY
      For CurX = DestRect.X1 To DestRect.X2
       SrcX = (SrcRect.X1 + (CurX - DestRect.X1)): DstX = CurX
       Filtered.R = ((DestBitMap.Datas(0, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(BltFlag.ChannelFrom, SrcX, SrcY) * AlphaMax))
       If (Filtered.R <> BltFlag.TransColor.R) Then
        DestBitMap.Datas(0, DstX, DstY) = Filtered.R
       End If
      Next CurX
     Next CurY
    Else
     For CurY = SrcRect.Y1 To SrcRect.Y2
      SrcY = CurY: DstY = (DestRect.Y1 + (CurY - SrcRect.Y1))
      For CurX = SrcRect.X1 To SrcRect.X2
       SrcX = CurX: DstX = (DestRect.X1 + (CurX - SrcRect.X1))
       Filtered.R = ((DestBitMap.Datas(0, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(BltFlag.ChannelFrom, SrcX, SrcY) * AlphaMax))
       If (Filtered.R <> BltFlag.TransColor.R) Then
        DestBitMap.Datas(0, DstX, DstY) = Filtered.R
       End If
      Next CurX
     Next CurY
    End If
   Case "T1,S0,Src24,Dst24":
    If (Rect2DDiagonalLength(SrcRect) > Rect2DDiagonalLength(DestRect)) Then
     For CurY = DestRect.Y1 To DestRect.Y2
      SrcY = (SrcRect.Y1 + (CurY - DestRect.Y1)): DstY = CurY
      For CurX = DestRect.X1 To DestRect.X2
       SrcX = (SrcRect.X1 + (CurX - DestRect.X1)): DstX = CurX
       Filtered.R = ((DestBitMap.Datas(0, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(0, SrcX, SrcY) * AlphaMax))
       Filtered.G = ((DestBitMap.Datas(1, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(1, SrcX, SrcY) * AlphaMax))
       Filtered.B = ((DestBitMap.Datas(2, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(2, SrcX, SrcY) * AlphaMax))
       If (ColorCompare(Filtered, BltFlag.TransColor) = False) Then
        DestBitMap.Datas(0, DstX, DstY) = Filtered.R
        DestBitMap.Datas(1, DstX, DstY) = Filtered.G
        DestBitMap.Datas(2, DstX, DstY) = Filtered.B
       End If
      Next CurX
     Next CurY
    Else
     For CurY = SrcRect.Y1 To SrcRect.Y2
      SrcY = CurY: DstY = (DestRect.Y1 + (CurY - SrcRect.Y1))
      For CurX = SrcRect.X1 To SrcRect.X2
       SrcX = CurX: DstX = (DestRect.X1 + (CurX - SrcRect.X1))
       Filtered.R = ((DestBitMap.Datas(0, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(0, SrcX, SrcY) * AlphaMax))
       Filtered.G = ((DestBitMap.Datas(1, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(1, SrcX, SrcY) * AlphaMax))
       Filtered.B = ((DestBitMap.Datas(2, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(2, SrcX, SrcY) * AlphaMax))
       If (ColorCompare(Filtered, BltFlag.TransColor) = False) Then
        DestBitMap.Datas(0, DstX, DstY) = Filtered.R
        DestBitMap.Datas(1, DstX, DstY) = Filtered.G
        DestBitMap.Datas(2, DstX, DstY) = Filtered.B
       End If
      Next CurX
     Next CurY
    End If
   Case "T1,S1,Src08,Dst08":
    For CurY = DestRect.Y1 To DestRect.Y2
     ParamY = ((CurY - DestRect.Y1) / (DestRect.Y2 - DestRect.Y1))
     For CurX = DestRect.X1 To DestRect.X2
      ParamX = ((CurX - DestRect.X1) / (DestRect.X2 - DestRect.X1))
      CurU = (SrcRect.X1 + ((SrcRect.X2 - SrcRect.X1) * ParamX))
      CurV = (SrcRect.Y1 + ((SrcRect.Y2 - SrcRect.Y1) * ParamY))
      Filtered = DoTexelFiltering(BltFlag.PixelFilter, SrcBitMap, CurU, CurV, False)
      Filtered.R = ((DestBitMap.Datas(0, CurX, CurY) * AlphaMin) + (Filtered.R * AlphaMax))
      If (Filtered.R <> BltFlag.TransColor.R) Then DestBitMap.Datas(0, CurX, CurY) = Filtered.R
     Next CurX
    Next CurY
   Case "T1,S1,Src08,Dst24":
    For CurY = DestRect.Y1 To DestRect.Y2
     ParamY = ((CurY - DestRect.Y1) / (DestRect.Y2 - DestRect.Y1))
     For CurX = DestRect.X1 To DestRect.X2
      ParamX = ((CurX - DestRect.X1) / (DestRect.X2 - DestRect.X1))
      CurU = (SrcRect.X1 + ((SrcRect.X2 - SrcRect.X1) * ParamX))
      CurV = (SrcRect.Y1 + ((SrcRect.Y2 - SrcRect.Y1) * ParamY))
      Filtered = DoTexelFiltering(BltFlag.PixelFilter, SrcBitMap, CurU, CurV, False)
      Filtered.R = ((DestBitMap.Datas(BltFlag.ChannelTo, CurX, CurY) * AlphaMin) + (Filtered.R * AlphaMax))
      If (Filtered.R <> BltFlag.TransColor.R) Then DestBitMap.Datas(BltFlag.ChannelTo, CurX, CurY) = Filtered.R
     Next CurX
    Next CurY
   Case "T1,S1,Src24,Dst08":
    For CurY = DestRect.Y1 To DestRect.Y2
     ParamY = ((CurY - DestRect.Y1) / (DestRect.Y2 - DestRect.Y1))
     For CurX = DestRect.X1 To DestRect.X2
      ParamX = ((CurX - DestRect.X1) / (DestRect.X2 - DestRect.X1))
      CurU = (SrcRect.X1 + ((SrcRect.X2 - SrcRect.X1) * ParamX))
      CurV = (SrcRect.Y1 + ((SrcRect.Y2 - SrcRect.Y1) * ParamY))
      Filtered = DoTexelFiltering(BltFlag.PixelFilter, SrcBitMap, CurU, CurV, False)
      If (BltFlag.ChannelFrom = 0) Then
       Filtered.R = ((DestBitMap.Datas(0, CurX, CurY) * AlphaMin) + (Filtered.R * AlphaMax))
       If (Filtered.R <> BltFlag.TransColor.R) Then DestBitMap.Datas(0, CurX, CurY) = Filtered.R
      ElseIf (BltFlag.ChannelFrom = 1) Then
       Filtered.G = ((DestBitMap.Datas(1, CurX, CurY) * AlphaMin) + (Filtered.G * AlphaMax))
       If (Filtered.G <> BltFlag.TransColor.G) Then DestBitMap.Datas(1, CurX, CurY) = Filtered.G
      ElseIf (BltFlag.ChannelFrom = 2) Then
       Filtered.B = ((DestBitMap.Datas(2, CurX, CurY) * AlphaMin) + (Filtered.B * AlphaMax))
       If (Filtered.B <> BltFlag.TransColor.B) Then DestBitMap.Datas(2, CurX, CurY) = Filtered.B
      End If
     Next CurX
    Next CurY
   Case "T1,S1,Src24,Dst24":
    For CurY = DestRect.Y1 To DestRect.Y2
     ParamY = ((CurY - DestRect.Y1) / (DestRect.Y2 - DestRect.Y1))
     For CurX = DestRect.X1 To DestRect.X2
      ParamX = ((CurX - DestRect.X1) / (DestRect.X2 - DestRect.X1))
      CurU = (SrcRect.X1 + ((SrcRect.X2 - SrcRect.X1) * ParamX))
      CurV = (SrcRect.Y1 + ((SrcRect.Y2 - SrcRect.Y1) * ParamY))
      Filtered = DoTexelFiltering(BltFlag.PixelFilter, SrcBitMap, CurU, CurV, False)
      Filtered.R = ((DestBitMap.Datas(0, CurX, CurY) * AlphaMin) + (Filtered.R * AlphaMax))
      Filtered.G = ((DestBitMap.Datas(1, CurX, CurY) * AlphaMin) + (Filtered.G * AlphaMax))
      Filtered.B = ((DestBitMap.Datas(2, CurX, CurY) * AlphaMin) + (Filtered.B * AlphaMax))
      If (ColorCompare(Filtered, BltFlag.TransColor) = False) Then
       DestBitMap.Datas(0, CurX, CurY) = Filtered.R
       DestBitMap.Datas(1, CurX, CurY) = Filtered.G
       DestBitMap.Datas(2, CurX, CurY) = Filtered.B
      End If
     Next CurX
    Next CurY
  End Select

 ElseIf (BltFlag.AlphaFlag = 3) Then 'Alpha blend with a mapping (AlphaMask)
  Select Case OperationCode
   Case "T0,S0,Src08,Dst08":
    If (Rect2DDiagonalLength(SrcRect) > Rect2DDiagonalLength(DestRect)) Then
     For CurY = DestRect.Y1 To DestRect.Y2
      SrcY = (SrcRect.Y1 + (CurY - DestRect.Y1)): DstY = CurY
      AlphaY = (BltFlag.AlphaRect.Y1 + (CurY - DestRect.Y1))
      For CurX = DestRect.X1 To DestRect.X2
       SrcX = (SrcRect.X1 + (CurX - DestRect.X1)): DstX = CurX
       AlphaX = (BltFlag.AlphaRect.X1 + (CurX - DestRect.X1))
       AlphaCol.R = BltFlag.AlphaMask.Datas(BltFlag.ChannelAlphaFrom, AlphaX, AlphaY)
       AlphaMin = (1 / AlphaCol.R): AlphaMax = (1 - AlphaMin)
       DestBitMap.Datas(0, DstX, DstY) = ((DestBitMap.Datas(0, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(0, SrcX, SrcY) * AlphaMax))
      Next CurX
     Next CurY
    Else
     For CurY = SrcRect.Y1 To SrcRect.Y2
      SrcY = CurY: DstY = (DestRect.Y1 + (CurY - SrcRect.Y1))
      AlphaY = (BltFlag.AlphaRect.Y1 + (CurY - SrcRect.Y1))
      For CurX = SrcRect.X1 To SrcRect.X2
       SrcX = CurX: DstX = (DestRect.X1 + (CurX - SrcRect.X1))
       AlphaX = (BltFlag.AlphaRect.X1 + (CurX - SrcRect.X1))
       AlphaCol.R = BltFlag.AlphaMask.Datas(BltFlag.ChannelAlphaFrom, AlphaX, AlphaY)
       AlphaMin = (1 / AlphaCol.R): AlphaMax = (1 - AlphaMin)
       DestBitMap.Datas(0, DstX, DstY) = ((DestBitMap.Datas(0, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(0, SrcX, SrcY) * AlphaMax))
      Next CurX
     Next CurY
    End If
   Case "T0,S0,Src08,Dst24":
    If (Rect2DDiagonalLength(SrcRect) > Rect2DDiagonalLength(DestRect)) Then
     For CurY = DestRect.Y1 To DestRect.Y2
      SrcY = (SrcRect.Y1 + (CurY - DestRect.Y1)): DstY = CurY
      AlphaY = (BltFlag.AlphaRect.Y1 + (CurY - DestRect.Y1))
      For CurX = DestRect.X1 To DestRect.X2
       SrcX = (SrcRect.X1 + (CurX - DestRect.X1)): DstX = CurX
       AlphaX = (BltFlag.AlphaRect.X1 + (CurX - DestRect.X1))
       AlphaCol.R = BltFlag.AlphaMask.Datas(BltFlag.ChannelAlphaFrom, AlphaX, AlphaY)
       AlphaMin = (1 / AlphaCol.R): AlphaMax = (1 - AlphaMin)
       DestBitMap.Datas(BltFlag.ChannelTo, DstX, DstY) = ((DestBitMap.Datas(BltFlag.ChannelTo, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(0, SrcX, SrcY) * AlphaMax))
      Next CurX
     Next CurY
    Else
     For CurY = SrcRect.Y1 To SrcRect.Y2
      SrcY = CurY: DstY = (DestRect.Y1 + (CurY - SrcRect.Y1))
      AlphaY = (BltFlag.AlphaRect.Y1 + (CurY - SrcRect.Y1))
      For CurX = SrcRect.X1 To SrcRect.X2
       SrcX = CurX: DstX = (DestRect.X1 + (CurX - SrcRect.X1))
       AlphaX = (BltFlag.AlphaRect.X1 + (CurX - SrcRect.X1))
       AlphaCol.R = BltFlag.AlphaMask.Datas(BltFlag.ChannelAlphaFrom, AlphaX, AlphaY)
       AlphaMin = (1 / AlphaCol.R): AlphaMax = (1 - AlphaMin)
       DestBitMap.Datas(BltFlag.ChannelTo, DstX, DstY) = ((DestBitMap.Datas(BltFlag.ChannelTo, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(0, SrcX, SrcY) * AlphaMax))
      Next CurX
     Next CurY
    End If
   Case "T0,S0,Src24,Dst08":
    If (Rect2DDiagonalLength(SrcRect) > Rect2DDiagonalLength(DestRect)) Then
     For CurY = DestRect.Y1 To DestRect.Y2
      SrcY = (SrcRect.Y1 + (CurY - DestRect.Y1)): DstY = CurY
      AlphaY = (BltFlag.AlphaRect.Y1 + (CurY - DestRect.Y1))
      For CurX = DestRect.X1 To DestRect.X2
       SrcX = (SrcRect.X1 + (CurX - DestRect.X1)): DstX = CurX
       AlphaX = (BltFlag.AlphaRect.X1 + (CurX - DestRect.X1))
       AlphaCol.R = BltFlag.AlphaMask.Datas(BltFlag.ChannelAlphaFrom, AlphaX, AlphaY)
       AlphaMin = (1 / AlphaCol.R): AlphaMax = (1 - AlphaMin)
       DestBitMap.Datas(0, DstX, DstY) = ((DestBitMap.Datas(0, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(BltFlag.ChannelFrom, SrcX, SrcY) * AlphaMax))
      Next CurX
     Next CurY
    Else
     For CurY = SrcRect.Y1 To SrcRect.Y2
      SrcY = CurY: DstY = (DestRect.Y1 + (CurY - SrcRect.Y1))
      AlphaY = (BltFlag.AlphaRect.Y1 + (CurY - SrcRect.Y1))
      For CurX = SrcRect.X1 To SrcRect.X2
       SrcX = CurX: DstX = (DestRect.X1 + (CurX - SrcRect.X1))
       AlphaX = (BltFlag.AlphaRect.X1 + (CurX - SrcRect.X1))
       AlphaCol.R = BltFlag.AlphaMask.Datas(BltFlag.ChannelAlphaFrom, AlphaX, AlphaY)
       AlphaMin = (1 / AlphaCol.R): AlphaMax = (1 - AlphaMin)
       DestBitMap.Datas(0, DstX, DstY) = ((DestBitMap.Datas(0, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(BltFlag.ChannelFrom, SrcX, SrcY) * AlphaMax))
      Next CurX
     Next CurY
    End If
   Case "T0,S0,Src24,Dst24":
    If (Rect2DDiagonalLength(SrcRect) > Rect2DDiagonalLength(DestRect)) Then
     For CurY = DestRect.Y1 To DestRect.Y2
      SrcY = (SrcRect.Y1 + (CurY - DestRect.Y1)): DstY = CurY
      AlphaY = (BltFlag.AlphaRect.Y1 + (CurY - DestRect.Y1))
      For CurX = DestRect.X1 To DestRect.X2
       SrcX = (SrcRect.X1 + (CurX - DestRect.X1)): DstX = CurX
       AlphaX = (BltFlag.AlphaRect.X1 + (CurX - DestRect.X1))
       AlphaCol.R = BltFlag.AlphaMask.Datas(BltFlag.ChannelAlphaFrom, AlphaX, AlphaY)
       AlphaMin = (1 / AlphaCol.R): AlphaMax = (1 - AlphaMin)
       DestBitMap.Datas(0, DstX, DstY) = ((DestBitMap.Datas(0, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(0, SrcX, SrcY) * AlphaMax))
       DestBitMap.Datas(1, DstX, DstY) = ((DestBitMap.Datas(1, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(1, SrcX, SrcY) * AlphaMax))
       DestBitMap.Datas(2, DstX, DstY) = ((DestBitMap.Datas(2, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(2, SrcX, SrcY) * AlphaMax))
      Next CurX
     Next CurY
    Else
     For CurY = SrcRect.Y1 To SrcRect.Y2
      SrcY = CurY: DstY = (DestRect.Y1 + (CurY - SrcRect.Y1))
      AlphaY = (BltFlag.AlphaRect.Y1 + (CurY - SrcRect.Y1))
      For CurX = SrcRect.X1 To SrcRect.X2
       SrcX = CurX: DstX = (DestRect.X1 + (CurX - SrcRect.X1))
       AlphaX = (BltFlag.AlphaRect.X1 + (CurX - SrcRect.X1))
       AlphaCol.R = BltFlag.AlphaMask.Datas(BltFlag.ChannelAlphaFrom, AlphaX, AlphaY)
       AlphaMin = (1 / AlphaCol.R): AlphaMax = (1 - AlphaMin)
       DestBitMap.Datas(0, DstX, DstY) = ((DestBitMap.Datas(0, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(0, SrcX, SrcY) * AlphaMax))
       DestBitMap.Datas(1, DstX, DstY) = ((DestBitMap.Datas(1, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(1, SrcX, SrcY) * AlphaMax))
       DestBitMap.Datas(2, DstX, DstY) = ((DestBitMap.Datas(2, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(2, SrcX, SrcY) * AlphaMax))
      Next CurX
     Next CurY
    End If
   Case "T0,S1,Src08,Dst08":
    For CurY = DestRect.Y1 To DestRect.Y2
     ParamY = ((CurY - DestRect.Y1) / (DestRect.Y2 - DestRect.Y1))
     For CurX = DestRect.X1 To DestRect.X2
      ParamX = ((CurX - DestRect.X1) / (DestRect.X2 - DestRect.X1))
      CurU = (SrcRect.X1 + ((SrcRect.X2 - SrcRect.X1) * ParamX))
      CurV = (SrcRect.Y1 + ((SrcRect.Y2 - SrcRect.Y1) * ParamY))
      ACurU = (BltFlag.AlphaRect.X1 + ((BltFlag.AlphaRect.X2 - BltFlag.AlphaRect.X1) * ParamX))
      ACurV = (BltFlag.AlphaRect.Y1 + ((BltFlag.AlphaRect.Y2 - BltFlag.AlphaRect.Y1) * ParamY))
      AlphaCol = DoTexelFiltering(BltFlag.PixelFilter, BltFlag.AlphaMask, ACurU, ACurV, False)
      AlphaMin = (1 / AlphaCol.R): AlphaMax = (1 - AlphaMin)
      DestBitMap.Datas(0, CurX, CurY) = ((DestBitMap.Datas(0, CurX, CurY) * AlphaMin) + (DoTexelFiltering(BltFlag.PixelFilter, SrcBitMap, CurU, CurV, False).R * AlphaMax))
     Next CurX
    Next CurY
   Case "T0,S1,Src08,Dst24":
    For CurY = DestRect.Y1 To DestRect.Y2
     ParamY = ((CurY - DestRect.Y1) / (DestRect.Y2 - DestRect.Y1))
     For CurX = DestRect.X1 To DestRect.X2
      ParamX = ((CurX - DestRect.X1) / (DestRect.X2 - DestRect.X1))
      CurU = (SrcRect.X1 + ((SrcRect.X2 - SrcRect.X1) * ParamX))
      CurV = (SrcRect.Y1 + ((SrcRect.Y2 - SrcRect.Y1) * ParamY))
      ACurU = (BltFlag.AlphaRect.X1 + ((BltFlag.AlphaRect.X2 - BltFlag.AlphaRect.X1) * ParamX))
      ACurV = (BltFlag.AlphaRect.Y1 + ((BltFlag.AlphaRect.Y2 - BltFlag.AlphaRect.Y1) * ParamY))
      AlphaCol = DoTexelFiltering(BltFlag.PixelFilter, BltFlag.AlphaMask, ACurU, ACurV, False)
      AlphaMin = (1 / AlphaCol.R): AlphaMax = (1 - AlphaMin)
      DestBitMap.Datas(BltFlag.ChannelTo, CurX, CurY) = ((DestBitMap.Datas(BltFlag.ChannelTo, CurX, CurY) * AlphaMin) + (DoTexelFiltering(BltFlag.PixelFilter, SrcBitMap, CurU, CurV, False).R * AlphaMax))
     Next CurX
    Next CurY
   Case "T0,S1,Src24,Dst08":
    For CurY = DestRect.Y1 To DestRect.Y2
     ParamY = ((CurY - DestRect.Y1) / (DestRect.Y2 - DestRect.Y1))
     For CurX = DestRect.X1 To DestRect.X2
      ParamX = ((CurX - DestRect.X1) / (DestRect.X2 - DestRect.X1))
      CurU = (SrcRect.X1 + ((SrcRect.X2 - SrcRect.X1) * ParamX))
      CurV = (SrcRect.Y1 + ((SrcRect.Y2 - SrcRect.Y1) * ParamY))
      ACurU = (BltFlag.AlphaRect.X1 + ((BltFlag.AlphaRect.X2 - BltFlag.AlphaRect.X1) * ParamX))
      ACurV = (BltFlag.AlphaRect.Y1 + ((BltFlag.AlphaRect.Y2 - BltFlag.AlphaRect.Y1) * ParamY))
      AlphaCol = DoTexelFiltering(BltFlag.PixelFilter, BltFlag.AlphaMask, ACurU, ACurV, False)
      AlphaMin = (1 / AlphaCol.R): AlphaMax = (1 - AlphaMin)
      If (BltFlag.ChannelFrom = 0) Then
       DestBitMap.Datas(0, CurX, CurY) = ((DestBitMap.Datas(0, CurX, CurY) * AlphaMin) + (DoTexelFiltering(BltFlag.PixelFilter, SrcBitMap, CurU, CurV, False).R * AlphaMax))
      ElseIf (BltFlag.ChannelFrom = 1) Then
       DestBitMap.Datas(0, CurX, CurY) = ((DestBitMap.Datas(0, CurX, CurY) * AlphaMin) + (DoTexelFiltering(BltFlag.PixelFilter, SrcBitMap, CurU, CurV, False).G * AlphaMax))
      ElseIf (BltFlag.ChannelFrom = 2) Then
       DestBitMap.Datas(0, CurX, CurY) = ((DestBitMap.Datas(0, CurX, CurY) * AlphaMin) + (DoTexelFiltering(BltFlag.PixelFilter, SrcBitMap, CurU, CurV, False).B * AlphaMax))
      End If
     Next CurX
    Next CurY
   Case "T0,S1,Src24,Dst24":
    For CurY = DestRect.Y1 To DestRect.Y2
     ParamY = ((CurY - DestRect.Y1) / (DestRect.Y2 - DestRect.Y1))
     For CurX = DestRect.X1 To DestRect.X2
      ParamX = ((CurX - DestRect.X1) / (DestRect.X2 - DestRect.X1))
      CurU = (SrcRect.X1 + ((SrcRect.X2 - SrcRect.X1) * ParamX))
      CurV = (SrcRect.Y1 + ((SrcRect.Y2 - SrcRect.Y1) * ParamY))
      ACurU = (BltFlag.AlphaRect.X1 + ((BltFlag.AlphaRect.X2 - BltFlag.AlphaRect.X1) * ParamX))
      ACurV = (BltFlag.AlphaRect.Y1 + ((BltFlag.AlphaRect.Y2 - BltFlag.AlphaRect.Y1) * ParamY))
      AlphaCol = DoTexelFiltering(BltFlag.PixelFilter, BltFlag.AlphaMask, ACurU, ACurV, False)
      AlphaMin = (1 / AlphaCol.R): AlphaMax = (1 - AlphaMin)
      Filtered = DoTexelFiltering(BltFlag.PixelFilter, SrcBitMap, CurU, CurV, False)
      DestBitMap.Datas(0, CurX, CurY) = ((DestBitMap.Datas(0, CurX, CurY) * AlphaMin) + (Filtered.R * AlphaMax))
      DestBitMap.Datas(1, CurX, CurY) = ((DestBitMap.Datas(1, CurX, CurY) * AlphaMin) + (Filtered.G * AlphaMax))
      DestBitMap.Datas(2, CurX, CurY) = ((DestBitMap.Datas(2, CurX, CurY) * AlphaMin) + (Filtered.B * AlphaMax))
     Next CurX
    Next CurY
   Case "T1,S0,Src08,Dst08":
    If (Rect2DDiagonalLength(SrcRect) > Rect2DDiagonalLength(DestRect)) Then
     For CurY = DestRect.Y1 To DestRect.Y2
      SrcY = (SrcRect.Y1 + (CurY - DestRect.Y1)): DstY = CurY
      AlphaY = (BltFlag.AlphaRect.Y1 + (CurY - DestRect.Y1))
      For CurX = DestRect.X1 To DestRect.X2
       SrcX = (SrcRect.X1 + (CurX - DestRect.X1)): DstX = CurX
       AlphaX = (BltFlag.AlphaRect.X1 + (CurX - DestRect.X1))
       AlphaCol.R = BltFlag.AlphaMask.Datas(BltFlag.ChannelAlphaFrom, AlphaX, AlphaY)
       AlphaMin = (1 / AlphaCol.R): AlphaMax = (1 - AlphaMin)
       Filtered.R = ((DestBitMap.Datas(0, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(0, SrcX, SrcY) * AlphaMax))
       If (Filtered.R <> BltFlag.TransColor.R) Then
        DestBitMap.Datas(0, DstX, DstY) = Filtered.R
       End If
      Next CurX
     Next CurY
    Else
     For CurY = SrcRect.Y1 To SrcRect.Y2
      SrcY = CurY: DstY = (DestRect.Y1 + (CurY - SrcRect.Y1))
      AlphaY = (BltFlag.AlphaRect.Y1 + (CurY - SrcRect.Y1))
      For CurX = SrcRect.X1 To SrcRect.X2
       SrcX = CurX: DstX = (DestRect.X1 + (CurX - SrcRect.X1))
       AlphaX = (BltFlag.AlphaRect.X1 + (CurX - SrcRect.X1))
       AlphaCol.R = BltFlag.AlphaMask.Datas(BltFlag.ChannelAlphaFrom, AlphaX, AlphaY)
       AlphaMin = (1 / AlphaCol.R): AlphaMax = (1 - AlphaMin)
       Filtered.R = ((DestBitMap.Datas(0, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(0, SrcX, SrcY) * AlphaMax))
       If (Filtered.R <> BltFlag.TransColor.R) Then
        DestBitMap.Datas(0, DstX, DstY) = Filtered.R
       End If
      Next CurX
     Next CurY
    End If
   Case "T1,S0,Src08,Dst24":
    If (Rect2DDiagonalLength(SrcRect) > Rect2DDiagonalLength(DestRect)) Then
     For CurY = DestRect.Y1 To DestRect.Y2
      SrcY = (SrcRect.Y1 + (CurY - DestRect.Y1)): DstY = CurY
      AlphaY = (BltFlag.AlphaRect.Y1 + (CurY - DestRect.Y1))
      For CurX = DestRect.X1 To DestRect.X2
       SrcX = (SrcRect.X1 + (CurX - DestRect.X1)): DstX = CurX
       AlphaX = (BltFlag.AlphaRect.X1 + (CurX - DestRect.X1))
       AlphaCol.R = BltFlag.AlphaMask.Datas(BltFlag.ChannelAlphaFrom, AlphaX, AlphaY)
       AlphaMin = (1 / AlphaCol.R): AlphaMax = (1 - AlphaMin)
       Filtered.R = ((DestBitMap.Datas(BltFlag.ChannelTo, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(0, SrcX, SrcY) * AlphaMax))
       If (Filtered.R <> BltFlag.TransColor.R) Then
        DestBitMap.Datas(BltFlag.ChannelTo, DstX, DstY) = Filtered.R
       End If
      Next CurX
     Next CurY
    Else
     For CurY = SrcRect.Y1 To SrcRect.Y2
      SrcY = CurY: DstY = (DestRect.Y1 + (CurY - SrcRect.Y1))
      AlphaY = (BltFlag.AlphaRect.Y1 + (CurY - SrcRect.Y1))
      For CurX = SrcRect.X1 To SrcRect.X2
       SrcX = CurX: DstX = (DestRect.X1 + (CurX - SrcRect.X1))
       AlphaX = (BltFlag.AlphaRect.X1 + (CurX - SrcRect.X1))
       AlphaCol.R = BltFlag.AlphaMask.Datas(BltFlag.ChannelAlphaFrom, AlphaX, AlphaY)
       AlphaMin = (1 / AlphaCol.R): AlphaMax = (1 - AlphaMin)
       Filtered.R = ((DestBitMap.Datas(BltFlag.ChannelTo, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(0, SrcX, SrcY) * AlphaMax))
       If (Filtered.R <> BltFlag.TransColor.R) Then
        DestBitMap.Datas(BltFlag.ChannelTo, DstX, DstY) = Filtered.R
       End If
      Next CurX
     Next CurY
    End If
   Case "T1,S0,Src24,Dst08":
    If (Rect2DDiagonalLength(SrcRect) > Rect2DDiagonalLength(DestRect)) Then
     For CurY = DestRect.Y1 To DestRect.Y2
      SrcY = (SrcRect.Y1 + (CurY - DestRect.Y1)): DstY = CurY
      AlphaY = (BltFlag.AlphaRect.Y1 + (CurY - DestRect.Y1))
      For CurX = DestRect.X1 To DestRect.X2
       SrcX = (SrcRect.X1 + (CurX - DestRect.X1)): DstX = CurX
       AlphaX = (BltFlag.AlphaRect.X1 + (CurX - DestRect.X1))
       AlphaCol.R = BltFlag.AlphaMask.Datas(BltFlag.ChannelAlphaFrom, AlphaX, AlphaY)
       AlphaMin = (1 / AlphaCol.R): AlphaMax = (1 - AlphaMin)
       Filtered.R = ((DestBitMap.Datas(0, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(BltFlag.ChannelFrom, SrcX, SrcY) * AlphaMax))
       If (Filtered.R <> BltFlag.TransColor.R) Then
        DestBitMap.Datas(0, DstX, DstY) = Filtered.R
       End If
      Next CurX
     Next CurY
    Else
     For CurY = SrcRect.Y1 To SrcRect.Y2
      SrcY = CurY: DstY = (DestRect.Y1 + (CurY - SrcRect.Y1))
      AlphaY = (BltFlag.AlphaRect.Y1 + (CurY - SrcRect.Y1))
      For CurX = SrcRect.X1 To SrcRect.X2
       SrcX = CurX: DstX = (DestRect.X1 + (CurX - SrcRect.X1))
       AlphaX = (BltFlag.AlphaRect.X1 + (CurX - SrcRect.X1))
       AlphaCol.R = BltFlag.AlphaMask.Datas(BltFlag.ChannelAlphaFrom, AlphaX, AlphaY)
       AlphaMin = (1 / AlphaCol.R): AlphaMax = (1 - AlphaMin)
       Filtered.R = ((DestBitMap.Datas(0, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(BltFlag.ChannelFrom, SrcX, SrcY) * AlphaMax))
       If (Filtered.R <> BltFlag.TransColor.R) Then
        DestBitMap.Datas(0, DstX, DstY) = Filtered.R
       End If
      Next CurX
     Next CurY
    End If
   Case "T1,S0,Src24,Dst24":
    If (Rect2DDiagonalLength(SrcRect) > Rect2DDiagonalLength(DestRect)) Then
     For CurY = DestRect.Y1 To DestRect.Y2
      SrcY = (SrcRect.Y1 + (CurY - DestRect.Y1)): DstY = CurY
      AlphaY = (BltFlag.AlphaRect.Y1 + (CurY - DestRect.Y1))
      For CurX = DestRect.X1 To DestRect.X2
       SrcX = (SrcRect.X1 + (CurX - DestRect.X1)): DstX = CurX
       AlphaX = (BltFlag.AlphaRect.X1 + (CurX - DestRect.X1))
       AlphaCol.R = BltFlag.AlphaMask.Datas(BltFlag.ChannelAlphaFrom, AlphaX, AlphaY)
       AlphaMin = (1 / AlphaCol.R): AlphaMax = (1 - AlphaMin)
       Filtered.R = ((DestBitMap.Datas(0, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(0, SrcX, SrcY) * AlphaMax))
       Filtered.G = ((DestBitMap.Datas(1, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(1, SrcX, SrcY) * AlphaMax))
       Filtered.B = ((DestBitMap.Datas(2, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(2, SrcX, SrcY) * AlphaMax))
       If (ColorCompare(Filtered, BltFlag.TransColor) = False) Then
        DestBitMap.Datas(0, DstX, DstY) = Filtered.R
        DestBitMap.Datas(1, DstX, DstY) = Filtered.G
        DestBitMap.Datas(2, DstX, DstY) = Filtered.B
       End If
      Next CurX
     Next CurY
    Else
     For CurY = SrcRect.Y1 To SrcRect.Y2
      SrcY = CurY: DstY = (DestRect.Y1 + (CurY - SrcRect.Y1))
      AlphaY = (BltFlag.AlphaRect.Y1 + (CurY - SrcRect.Y1))
      For CurX = SrcRect.X1 To SrcRect.X2
       SrcX = CurX: DstX = (DestRect.X1 + (CurX - SrcRect.X1))
       AlphaX = (BltFlag.AlphaRect.X1 + (CurX - SrcRect.X1))
       AlphaCol.R = BltFlag.AlphaMask.Datas(BltFlag.ChannelAlphaFrom, AlphaX, AlphaY)
       AlphaMin = (1 / AlphaCol.R): AlphaMax = (1 - AlphaMin)
       Filtered.R = ((DestBitMap.Datas(0, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(0, SrcX, SrcY) * AlphaMax))
       Filtered.G = ((DestBitMap.Datas(1, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(1, SrcX, SrcY) * AlphaMax))
       Filtered.B = ((DestBitMap.Datas(2, DstX, DstY) * AlphaMin) + (SrcBitMap.Datas(2, SrcX, SrcY) * AlphaMax))
       If (ColorCompare(Filtered, BltFlag.TransColor) = False) Then
        DestBitMap.Datas(0, DstX, DstY) = Filtered.R
        DestBitMap.Datas(1, DstX, DstY) = Filtered.G
        DestBitMap.Datas(2, DstX, DstY) = Filtered.B
       End If
      Next CurX
     Next CurY
    End If
   Case "T1,S1,Src08,Dst08":
    For CurY = DestRect.Y1 To DestRect.Y2
     ParamY = ((CurY - DestRect.Y1) / (DestRect.Y2 - DestRect.Y1))
     For CurX = DestRect.X1 To DestRect.X2
      ParamX = ((CurX - DestRect.X1) / (DestRect.X2 - DestRect.X1))
      CurU = (SrcRect.X1 + ((SrcRect.X2 - SrcRect.X1) * ParamX))
      CurV = (SrcRect.Y1 + ((SrcRect.Y2 - SrcRect.Y1) * ParamY))
      ACurU = (BltFlag.AlphaRect.X1 + ((BltFlag.AlphaRect.X2 - BltFlag.AlphaRect.X1) * ParamX))
      ACurV = (BltFlag.AlphaRect.Y1 + ((BltFlag.AlphaRect.Y2 - BltFlag.AlphaRect.Y1) * ParamY))
      Filtered = DoTexelFiltering(BltFlag.PixelFilter, SrcBitMap, CurU, CurV, False)
      AlphaCol = DoTexelFiltering(BltFlag.PixelFilter, BltFlag.AlphaMask, ACurU, ACurV, False)
      AlphaMin = (1 / AlphaCol.R): AlphaMax = (1 - AlphaMin)
      Filtered.R = ((DestBitMap.Datas(0, CurX, CurY) * AlphaMin) + (Filtered.R * AlphaMax))
      If (Filtered.R <> BltFlag.TransColor.R) Then
       DestBitMap.Datas(0, CurX, CurY) = Filtered.R
      End If
     Next CurX
    Next CurY
   Case "T1,S1,Src08,Dst24":
    For CurY = DestRect.Y1 To DestRect.Y2
     ParamY = ((CurY - DestRect.Y1) / (DestRect.Y2 - DestRect.Y1))
     For CurX = DestRect.X1 To DestRect.X2
      ParamX = ((CurX - DestRect.X1) / (DestRect.X2 - DestRect.X1))
      CurU = (SrcRect.X1 + ((SrcRect.X2 - SrcRect.X1) * ParamX))
      CurV = (SrcRect.Y1 + ((SrcRect.Y2 - SrcRect.Y1) * ParamY))
      ACurU = (BltFlag.AlphaRect.X1 + ((BltFlag.AlphaRect.X2 - BltFlag.AlphaRect.X1) * ParamX))
      ACurV = (BltFlag.AlphaRect.Y1 + ((BltFlag.AlphaRect.Y2 - BltFlag.AlphaRect.Y1) * ParamY))
      Filtered = DoTexelFiltering(BltFlag.PixelFilter, SrcBitMap, CurU, CurV, False)
      AlphaCol = DoTexelFiltering(BltFlag.PixelFilter, BltFlag.AlphaMask, ACurU, ACurV, False)
      AlphaMin = (1 / AlphaCol.R): AlphaMax = (1 - AlphaMin)
      Filtered.R = ((DestBitMap.Datas(BltFlag.ChannelTo, CurX, CurY) * AlphaMin) + (Filtered.R * AlphaMax))
      If (Filtered.R <> BltFlag.TransColor.R) Then
       DestBitMap.Datas(BltFlag.ChannelTo, CurX, CurY) = Filtered.R
      End If
     Next CurX
    Next CurY
   Case "T1,S1,Src24,Dst08":
    For CurY = DestRect.Y1 To DestRect.Y2
     ParamY = ((CurY - DestRect.Y1) / (DestRect.Y2 - DestRect.Y1))
     For CurX = DestRect.X1 To DestRect.X2
      ParamX = ((CurX - DestRect.X1) / (DestRect.X2 - DestRect.X1))
      CurU = (SrcRect.X1 + ((SrcRect.X2 - SrcRect.X1) * ParamX))
      CurV = (SrcRect.Y1 + ((SrcRect.Y2 - SrcRect.Y1) * ParamY))
      ACurU = (BltFlag.AlphaRect.X1 + ((BltFlag.AlphaRect.X2 - BltFlag.AlphaRect.X1) * ParamX))
      ACurV = (BltFlag.AlphaRect.Y1 + ((BltFlag.AlphaRect.Y2 - BltFlag.AlphaRect.Y1) * ParamY))
      Filtered = DoTexelFiltering(BltFlag.PixelFilter, SrcBitMap, CurU, CurV, False)
      AlphaCol = DoTexelFiltering(BltFlag.PixelFilter, BltFlag.AlphaMask, ACurU, ACurV, False)
      AlphaMin = (1 / AlphaCol.R): AlphaMax = (1 - AlphaMin)
      If (BltFlag.ChannelFrom = 0) Then
       Filtered.R = ((DestBitMap.Datas(0, CurX, CurY) * AlphaMin) + (Filtered.R * AlphaMax))
       If (Filtered.R <> BltFlag.TransColor.R) Then DestBitMap.Datas(0, CurX, CurY) = Filtered.R
      End If
      If (BltFlag.ChannelFrom = 1) Then
       Filtered.G = ((DestBitMap.Datas(1, CurX, CurY) * AlphaMin) + (Filtered.G * AlphaMax))
       If (Filtered.G <> BltFlag.TransColor.G) Then DestBitMap.Datas(1, CurX, CurY) = Filtered.G
      End If
      If (BltFlag.ChannelFrom = 2) Then
       Filtered.B = ((DestBitMap.Datas(2, CurX, CurY) * AlphaMin) + (Filtered.B * AlphaMax))
       If (Filtered.B <> BltFlag.TransColor.B) Then DestBitMap.Datas(2, CurX, CurY) = Filtered.B
      End If
     Next CurX
    Next CurY
   Case "T1,S1,Src24,Dst24":
    For CurY = DestRect.Y1 To DestRect.Y2
     ParamY = ((CurY - DestRect.Y1) / (DestRect.Y2 - DestRect.Y1))
     For CurX = DestRect.X1 To DestRect.X2
      ParamX = ((CurX - DestRect.X1) / (DestRect.X2 - DestRect.X1))
      CurU = (SrcRect.X1 + ((SrcRect.X2 - SrcRect.X1) * ParamX))
      CurV = (SrcRect.Y1 + ((SrcRect.Y2 - SrcRect.Y1) * ParamY))
      ACurU = (BltFlag.AlphaRect.X1 + ((BltFlag.AlphaRect.X2 - BltFlag.AlphaRect.X1) * ParamX))
      ACurV = (BltFlag.AlphaRect.Y1 + ((BltFlag.AlphaRect.Y2 - BltFlag.AlphaRect.Y1) * ParamY))
      Filtered = DoTexelFiltering(BltFlag.PixelFilter, SrcBitMap, CurU, CurV, False)
      AlphaCol = DoTexelFiltering(BltFlag.PixelFilter, BltFlag.AlphaMask, ACurU, ACurV, False)
      AlphaMin = (1 / AlphaCol.R): AlphaMax = (1 - AlphaMin)
      Filtered.R = ((DestBitMap.Datas(0, CurX, CurY) * AlphaMin) + (Filtered.R * AlphaMax))
      Filtered.G = ((DestBitMap.Datas(1, CurX, CurY) * AlphaMin) + (Filtered.G * AlphaMax))
      Filtered.B = ((DestBitMap.Datas(2, CurX, CurY) * AlphaMin) + (Filtered.B * AlphaMax))
      If (ColorCompare(Filtered, BltFlag.TransColor) = False) Then
       DestBitMap.Datas(0, CurX, CurY) = Filtered.R
       DestBitMap.Datas(1, CurX, CurY) = Filtered.G
       DestBitMap.Datas(2, CurX, CurY) = Filtered.B
      End If
     Next CurX
    Next CurY
  End Select
 End If

End Function
Function BitMap2D_ColorFill(TheBitmap As BitMap2D, DestRect As Rect2D, Color As ColorRGB)

 '##################################################################################
 '##################################################################################
 '###
 '###  FUNCTION    : BiMap2D_ColorFill
 '###
 '###  DESCRIPTION : Fill the specified rectangular region with the given color.
 '###
 '#################################################################################
 '#################################################################################

 If (BitMap2D_IsValid(TheBitmap) = False) Then Exit Function
 If (Rect2DIsInsideRegion(DestRect, TheBitmap.Dimensions) = False) Then Exit Function
 If (ColorIsValid(Color) = False) Then Exit Function

 Dim CurX&, CurY&
 Select Case TheBitmap.BitsDepth
  Case 8:
   For CurY = DestRect.Y1 To DestRect.Y2
    For CurX = DestRect.X1 To DestRect.X2
     TheBitmap.Datas(0, CurX, CurY) = Color.R
    Next CurX
   Next CurY
  Case 24:
   For CurY = DestRect.Y1 To DestRect.Y2
    For CurX = DestRect.X1 To DestRect.X2
     TheBitmap.Datas(0, CurX, CurY) = Color.R
     TheBitmap.Datas(1, CurX, CurY) = Color.G
     TheBitmap.Datas(2, CurX, CurY) = Color.B
    Next CurX
   Next CurY
 End Select

End Function
Function BitMap2D_Create(DestBitMap As BitMap2D, Label As String, BitsDepth As Byte, Width%, Height%, BackGroundColor As ColorRGB)

 '##################################################################################
 '##################################################################################
 '###
 '###  FUNCTION    : BiMap2D_Create
 '###
 '###  DESCRIPTION : Create directly a bitmap buffer.
 '###
 '##################################################################################
 '##################################################################################

 With DestBitMap
  If ((BitsDepth <> 8) And (BitsDepth <> 24)) Then Exit Function
  If (Width < MinBitMapWidth) Then Exit Function
  If (Height < MinBitMapHeight) Then Exit Function
  If (Width > MaxBitMapWidth) Then Exit Function
  If (Height > MaxBitMapHeight) Then Exit Function
  .Label = Label: .BitsDepth = BitsDepth
  .Dimensions.X = Width: .Dimensions.Y = Height
  .BackGroundColor = BackGroundColor
  Select Case BitsDepth
   Case 8:  ReDim .Datas(0, Width, Height)
   Case 24: ReDim .Datas(2, Width, Height)
  End Select
 End With

 BitMap2D_ColorFill DestBitMap, Rect2DInput(0, 0, DestBitMap.Dimensions.X, DestBitMap.Dimensions.Y), BackGroundColor

End Function
Function BitMap2D_DisplayToDC1(SrcBitMap As BitMap2D, X&, Y&, DestDC&)

 '##################################################################################
 '##################################################################################
 '###
 '###  FUNCTION    : BiMap2D_DisplayToDC1
 '###
 '###  DESCRIPTION : 24 bits bitmaps only!
 '###                Display a bitmap's buffer, in a specified handle of a Device
 '###                Context(DC), using the 'SetPixel' API (GDI).
 '###
 '##################################################################################
 '##################################################################################

 If (BitMap2D_IsValid(SrcBitMap) = False) Then Exit Function
 If (SrcBitMap.BitsDepth <> 24) Then Exit Function

 Dim CurX&, CurY&
 For CurY = Y To (Y + SrcBitMap.Dimensions.Y)
  For CurX = X To (X + SrcBitMap.Dimensions.X)
   SetPixel DestDC, CurX, CurY, ColorRGBToLong(ColorInput(CInt(SrcBitMap.Datas(0, (CurX - X), (CurY - Y))), _
                                                          CInt(SrcBitMap.Datas(1, (CurX - X), (CurY - Y))), _
                                                          CInt(SrcBitMap.Datas(2, (CurX - X), (CurY - Y)))))
  Next CurX
 Next CurY

End Function
Function BitMap2D_DisplayToDC2(SrcBitMap As BitMap2D, DestDC&, DestImage&)

 '##################################################################################
 '##################################################################################
 '###
 '###  FUNCTION    : BiMap2D_DisplayToDC1
 '###
 '###  DESCRIPTION : 24 bits bitmaps only!
 '###                Display a bitmap's buffer, in a specified handle of a Device
 '###                Context(DC) and a bitmap object, using the 'SetDIBits' API (GDI).
 '###
 '##################################################################################
 '##################################################################################

 If (BitMap2D_IsValid(SrcBitMap) = False) Then Exit Function
 If (SrcBitMap.BitsDepth <> 24) Then Exit Function

 Dim BInfos As BITMAPINFO

 With BInfos
  .bmiHeader.biBitCount = 24
  .bmiHeader.biWidth = SrcBitMap.Dimensions.X
  .bmiHeader.biHeight = -SrcBitMap.Dimensions.Y
  .bmiHeader.biPlanes = 1
  .bmiHeader.biSize = LenB(.bmiHeader)
 End With

 SetDIBits DestDC, DestImage, 0, SrcBitMap.Dimensions.Y, SrcBitMap.Datas(0, 0, 0), BInfos, 0

End Function
Function BitMap2D_Dummy() As BitMap2D

 '##################################################################################
 '##################################################################################
 '###
 '###  FUNCTION    : BiMap2D_Dummy
 '###
 '###  DESCRIPTION : Just return an 'empty' bitmap, used to avoid the declaration
 '###                of a dummy bitmap.
 '###
 '##################################################################################
 '##################################################################################

End Function
Function BitMap2D_Erase(TheBitmap As BitMap2D)

 '##################################################################################
 '##################################################################################
 '###
 '###  FUNCTION    : BiMap2D_Erase
 '###
 '###  DESCRIPTION : Do erase the bitmap, an erase operation is just filling
 '###                all the bitmap, with the bitmap's backcolor (uses bltColorFill)
 '###
 '##################################################################################
 '##################################################################################

 If (BitMap2D_IsValid(TheBitmap) = False) Then Exit Function

 BitMap2D_ColorFill TheBitmap, Rect2DInput(0, 0, TheBitmap.Dimensions.X, TheBitmap.Dimensions.Y), TheBitmap.BackGroundColor

End Function
Function BitMap2D_Delete(TheBitmap As BitMap2D)

 '##################################################################################
 '##################################################################################
 '###
 '###  FUNCTION    : BiMap2D_Delete
 '###
 '###  DESCRIPTION : Destroy a given bitmap.
 '###
 '##################################################################################
 '##################################################################################

 With TheBitmap
  .Label = vbNullString
  .BitsDepth = 0
  .Dimensions.X = 0: .Dimensions.Y = 0
  .BackGroundColor = ColorBlack
  ReDim .Datas(0)
 End With

End Function
Function BitMap2D_Filter(TheBitmap As BitMap2D, DestRect As Rect2D, BitMapFilter As K3DE_BITMAP_FILTER_MODES, FilterVal As Single)

 '##################################################################################
 '##################################################################################
 '###
 '###  FUNCTION    : BiMap2D_Filter
 '###
 '###  DESCRIPTION : Applicate a chosen 1 from the cool collections of
 '###                2D bitmaps filters, this is includ :
 '###
 '###                Brightness, Contrast, Gamma-correction, Grey-Scaling
 '###                Monochromatic conversion, Colors-inversion (negatif),
 '###                and swapping between image's RGB channels.
 '###
 '##################################################################################
 '##################################################################################

 If (BitMap2D_IsValid(TheBitmap) = False) Then Exit Function
 If (Rect2DIsInsideRegion(DestRect, TheBitmap.Dimensions) = False) Then Exit Function

 Dim CurX&, CurY&, InColor As ColorRGB, OutColor As ColorRGB

 Select Case BitMapFilter
  Case K3DE_BFM_BRIGHTNESS:
   If ((FilterVal < -255) Or (FilterVal > 255)) Then Exit Function
   FilterVal = Fix(FilterVal)
   Select Case TheBitmap.BitsDepth
    Case 8:
     For CurY = DestRect.Y1 To DestRect.Y2
      For CurX = DestRect.X1 To DestRect.X2
       InColor.R = CInt(TheBitmap.Datas(0, CurX, CurY))
       TheBitmap.Datas(0, CurX, CurY) = ColorFilterBrightness(InColor, CInt(FilterVal)).R
      Next CurX
     Next CurY
    Case 24:
     For CurY = DestRect.Y1 To DestRect.Y2
      For CurX = DestRect.X1 To DestRect.X2
       InColor.R = CInt(TheBitmap.Datas(0, CurX, CurY))
       InColor.G = CInt(TheBitmap.Datas(1, CurX, CurY))
       InColor.B = CInt(TheBitmap.Datas(2, CurX, CurY))
       OutColor = ColorFilterBrightness(InColor, CInt(FilterVal))
       TheBitmap.Datas(0, CurX, CurY) = OutColor.R
       TheBitmap.Datas(1, CurX, CurY) = OutColor.G
       TheBitmap.Datas(2, CurX, CurY) = OutColor.B
      Next CurX
     Next CurY
   End Select
  Case K3DE_BFM_CONTRAST:
   If ((FilterVal < 0) Or (FilterVal > 10)) Then Exit Function
   Select Case TheBitmap.BitsDepth
    Case 8:
     For CurY = DestRect.Y1 To DestRect.Y2
      For CurX = DestRect.X1 To DestRect.X2
       InColor.R = CInt(TheBitmap.Datas(0, CurX, CurY))
       TheBitmap.Datas(0, CurX, CurY) = ColorFilterContrast(InColor, FilterVal).R
      Next CurX
     Next CurY
    Case 24:
     For CurY = DestRect.Y1 To DestRect.Y2
      For CurX = DestRect.X1 To DestRect.X2
       InColor.R = CInt(TheBitmap.Datas(0, CurX, CurY))
       InColor.G = CInt(TheBitmap.Datas(1, CurX, CurY))
       InColor.B = CInt(TheBitmap.Datas(2, CurX, CurY))
       OutColor = ColorFilterContrast(InColor, FilterVal)
       TheBitmap.Datas(0, CurX, CurY) = OutColor.R
       TheBitmap.Datas(1, CurX, CurY) = OutColor.G
       TheBitmap.Datas(2, CurX, CurY) = OutColor.B
      Next CurX
     Next CurY
   End Select
  Case K3DE_BFM_GAMMA:

   If ((FilterVal < ApproachVal) Or (FilterVal > 10)) Then Exit Function
   Select Case TheBitmap.BitsDepth
    Case 8:
     For CurY = DestRect.Y1 To DestRect.Y2
      For CurX = DestRect.X1 To DestRect.X2
       InColor.R = CInt(TheBitmap.Datas(0, CurX, CurY))
       TheBitmap.Datas(0, CurX, CurY) = ColorFilterGamma(InColor, FilterVal).R
      Next CurX
     Next CurY
    Case 24:
     For CurY = DestRect.Y1 To DestRect.Y2
      For CurX = DestRect.X1 To DestRect.X2
       InColor.R = CInt(TheBitmap.Datas(0, CurX, CurY))
       InColor.G = CInt(TheBitmap.Datas(1, CurX, CurY))
       InColor.B = CInt(TheBitmap.Datas(2, CurX, CurY))
       OutColor = ColorFilterGamma(InColor, FilterVal)
       TheBitmap.Datas(0, CurX, CurY) = OutColor.R
       TheBitmap.Datas(1, CurX, CurY) = OutColor.G
       TheBitmap.Datas(2, CurX, CurY) = OutColor.B
      Next CurX
     Next CurY
   End Select

  Case K3DE_BFM_GREYSCALE:
   If (TheBitmap.BitsDepth = 8) Then Exit Function
   For CurY = DestRect.Y1 To DestRect.Y2
    For CurX = DestRect.X1 To DestRect.X2
     InColor.R = CInt(TheBitmap.Datas(0, CurX, CurY))
     InColor.G = CInt(TheBitmap.Datas(1, CurX, CurY))
     InColor.B = CInt(TheBitmap.Datas(2, CurX, CurY))
     OutColor = ColorFilterGreyScale(InColor)
     TheBitmap.Datas(0, CurX, CurY) = OutColor.R
     TheBitmap.Datas(1, CurX, CurY) = OutColor.G
     TheBitmap.Datas(2, CurX, CurY) = OutColor.B
    Next CurX
   Next CurY
  Case K3DE_BFM_MONOCHROME:
   If (TheBitmap.BitsDepth = 8) Then Exit Function
   For CurY = DestRect.Y1 To DestRect.Y2
    For CurX = DestRect.X1 To DestRect.X2
     InColor.R = CInt(TheBitmap.Datas(0, CurX, CurY))
     InColor.G = CInt(TheBitmap.Datas(1, CurX, CurY))
     InColor.B = CInt(TheBitmap.Datas(2, CurX, CurY))
     OutColor = ColorFilterMono(InColor)
     TheBitmap.Datas(0, CurX, CurY) = OutColor.R
     TheBitmap.Datas(1, CurX, CurY) = OutColor.G
     TheBitmap.Datas(2, CurX, CurY) = OutColor.B
    Next CurX
   Next CurY
  Case K3DE_BFM_INVERT:
   Select Case TheBitmap.BitsDepth
    Case 8:
     For CurY = DestRect.Y1 To DestRect.Y2
      For CurX = DestRect.X1 To DestRect.X2
       InColor.R = CInt(TheBitmap.Datas(0, CurX, CurY))
       TheBitmap.Datas(0, CurX, CurY) = ColorInvert(InColor).R
      Next CurX
     Next CurY
    Case 24:
     For CurY = DestRect.Y1 To DestRect.Y2
      For CurX = DestRect.X1 To DestRect.X2
       InColor.R = CInt(TheBitmap.Datas(0, CurX, CurY))
       InColor.G = CInt(TheBitmap.Datas(1, CurX, CurY))
       InColor.B = CInt(TheBitmap.Datas(2, CurX, CurY))
       OutColor = ColorInvert(InColor)
       TheBitmap.Datas(0, CurX, CurY) = OutColor.R
       TheBitmap.Datas(1, CurX, CurY) = OutColor.G
       TheBitmap.Datas(2, CurX, CurY) = OutColor.B
      Next CurX
     Next CurY
   End Select
  Case K3DE_BFM_SWAPCHANNELS:
   If (TheBitmap.BitsDepth = 8) Then Exit Function
   If ((FilterVal < 1) Or (FilterVal > 5)) Then Exit Function
   FilterVal = Fix(FilterVal)
   Select Case FilterVal
    Case 1: 'rbg
     For CurY = DestRect.Y1 To DestRect.Y2
      For CurX = DestRect.X1 To DestRect.X2
       InColor.G = CInt(TheBitmap.Datas(1, CurX, CurY))
       TheBitmap.Datas(1, CurX, CurY) = TheBitmap.Datas(2, CurX, CurY)
       TheBitmap.Datas(2, CurX, CurY) = InColor.G
      Next CurX
     Next CurY
    Case 2: 'gbr
     For CurY = DestRect.Y1 To DestRect.Y2
      For CurX = DestRect.X1 To DestRect.X2
       InColor.R = CInt(TheBitmap.Datas(0, CurX, CurY))
       InColor.G = CInt(TheBitmap.Datas(1, CurX, CurY))
       InColor.B = CInt(TheBitmap.Datas(2, CurX, CurY))
       TheBitmap.Datas(0, CurX, CurY) = InColor.G
       TheBitmap.Datas(1, CurX, CurY) = InColor.B
       TheBitmap.Datas(2, CurX, CurY) = InColor.R
      Next CurX
     Next CurY
    Case 3: 'grb
     For CurY = DestRect.Y1 To DestRect.Y2
      For CurX = DestRect.X1 To DestRect.X2
       InColor.R = CInt(TheBitmap.Datas(0, CurX, CurY))
       InColor.G = CInt(TheBitmap.Datas(1, CurX, CurY))
       TheBitmap.Datas(0, CurX, CurY) = InColor.G
       TheBitmap.Datas(1, CurX, CurY) = InColor.R
      Next CurX
     Next CurY
    Case 4: 'bgr
     For CurY = DestRect.Y1 To DestRect.Y2
      For CurX = DestRect.X1 To DestRect.X2
       InColor.R = CInt(TheBitmap.Datas(0, CurX, CurY))
       InColor.B = CInt(TheBitmap.Datas(2, CurX, CurY))
       TheBitmap.Datas(0, CurX, CurY) = InColor.B
       TheBitmap.Datas(2, CurX, CurY) = InColor.R
      Next CurX
     Next CurY
    Case 5: 'brg
     For CurY = DestRect.Y1 To DestRect.Y2
      For CurX = DestRect.X1 To DestRect.X2
       InColor.R = CInt(TheBitmap.Datas(0, CurX, CurY))
       InColor.G = CInt(TheBitmap.Datas(1, CurX, CurY))
       InColor.B = CInt(TheBitmap.Datas(2, CurX, CurY))
       TheBitmap.Datas(0, CurX, CurY) = InColor.B
       TheBitmap.Datas(1, CurX, CurY) = InColor.R
       TheBitmap.Datas(2, CurX, CurY) = InColor.G
      Next CurX
     Next CurY
   End Select
 End Select

End Function
Function BitMap2D_GetPixel(TheBitmap As BitMap2D, X%, Y%) As ColorRGB

 '##################################################################################
 '##################################################################################
 '###
 '###  FUNCTION    : BiMap2D_GetPixel
 '###
 '###  DESCRIPTION : Recieve the pixel color at a given X,Y coordinates.
 '###
 '##################################################################################
 '##################################################################################

 BitMap2D_GetPixel = ColorNone

 If (BitMap2D_IsValid(TheBitmap) = False) Then Exit Function
 If (Point2DInRegion(TheBitmap.Dimensions, Point2DInput(X, Y)) = False) Then Exit Function

 Select Case TheBitmap.BitsDepth
  Case 8:  BitMap2D_GetPixel = ColorInput(CInt(TheBitmap.Datas(0, X, Y)), -1, -1)
  Case 24: BitMap2D_GetPixel = ColorInput(CInt(TheBitmap.Datas(0, X, Y)), CInt(TheBitmap.Datas(1, X, Y)), CInt(TheBitmap.Datas(2, X, Y)))
 End Select

End Function
Function BitMap2D_IsValid(TheBitmap As BitMap2D) As Boolean

 '##################################################################################
 '##################################################################################
 '###
 '###  FUNCTION    : BiMap2D_IsValid
 '###
 '###  DESCRIPTION : Do a validation check on a given bitmap.
 '###
 '##################################################################################
 '##################################################################################

 On Error GoTo Eror

 Select Case TheBitmap.BitsDepth
  Case 8:  If (UBound(TheBitmap.Datas(), 1) > 0) Then Exit Function
  Case 24: If (UBound(TheBitmap.Datas(), 1) > 2) Then Exit Function
  Case Else: Exit Function 'the engine support 8, 24 bits images only
 End Select

 If (UBound(TheBitmap.Datas(), 2) <> TheBitmap.Dimensions.X) Then Exit Function
 If (UBound(TheBitmap.Datas(), 3) <> TheBitmap.Dimensions.Y) Then Exit Function

 If (TheBitmap.Dimensions.X < MinBitMapWidth) Then Exit Function
 If (TheBitmap.Dimensions.Y < MinBitMapHeight) Then Exit Function
 If (TheBitmap.Dimensions.X > MaxBitMapWidth) Then Exit Function
 If (TheBitmap.Dimensions.Y > MaxBitMapHeight) Then Exit Function

 BitMap2D_IsValid = True

Eror:
 If (Err.Number = 9) Then Err.Number = 0: Exit Function

End Function
Function BitMap2D_SetPixel(TheBitmap As BitMap2D, X%, Y%, Color As ColorRGB)

 '##################################################################################
 '##################################################################################
 '###
 '###  FUNCTION    : BiMap2D_SetPixel
 '###
 '###  DESCRIPTION : Update the pixel color at a given XY position.
 '###
 '##################################################################################
 '##################################################################################

 If (BitMap2D_IsValid(TheBitmap) = False) Then Exit Function
 If (Point2DInRegion(TheBitmap.Dimensions, Point2DInput(X, Y)) = False) Then Exit Function
 If (ColorIsValid(Color) = False) Then Exit Function

 Select Case TheBitmap.BitsDepth
  Case 8:  If (ColorCheckChannels(Color) = 2) Then TheBitmap.Datas(0, X, Y) = Color.R
  Case 24:
           If (ColorCheckChannels(Color) = 8) Then
            TheBitmap.Datas(0, X, Y) = Color.R
            TheBitmap.Datas(1, X, Y) = Color.G
            TheBitmap.Datas(2, X, Y) = Color.B
           End If
 End Select

End Function
Function BitMap2D_Resize(TheBitmap As BitMap2D, NewWidth%, NewHeight%)

 '##################################################################################
 '##################################################################################
 '###
 '###  FUNCTION    : BiMap2D_Resize
 '###
 '###  DESCRIPTION : Resize a given bitmap, no strech is applied (resampling), but
 '###                the new regions are filled with the bitmap's background color.
 '###
 '##################################################################################
 '##################################################################################

 If (BitMap2D_IsValid(TheBitmap) = False) Then Exit Function
 If ((NewWidth < MinBitMapWidth) Or (NewWidth > MaxBitMapWidth)) Then Exit Function
 If ((NewHeight < MinBitMapHeight) Or (NewHeight > MaxBitMapHeight)) Then Exit Function

 Dim TmpBitMap As BitMap2D
 Dim PreservedWidth%, PreservedHeight%
 Dim BltFlag As BltFlags, BltRect As Rect2D

 PreservedWidth = TheBitmap.Dimensions.X
 PreservedHeight = TheBitmap.Dimensions.Y
 If (NewWidth < TheBitmap.Dimensions.X) Then PreservedWidth = NewWidth
 If (NewHeight < TheBitmap.Dimensions.Y) Then PreservedHeight = NewHeight

 With TmpBitMap
  .Label = "Temporary !"
  .Dimensions = Point2DInput(PreservedWidth, PreservedHeight)
  .BitsDepth = TheBitmap.BitsDepth
  .BackGroundColor = TheBitmap.BackGroundColor
  If (.BitsDepth = 8) Then
   ReDim .Datas(0, .Dimensions.X, .Dimensions.Y)
  ElseIf (.BitsDepth = 24) Then
   ReDim .Datas(2, .Dimensions.X, .Dimensions.Y)
  End If
 End With

 BltFlag.AlphaFlag = 1
 BltRect = Rect2DInput(0, 0, PreservedWidth, PreservedHeight)
 BitMap2D_Blt TheBitmap, TmpBitMap, BltRect, BltRect, BltFlag

 Select Case TheBitmap.BitsDepth
  Case 8:  ReDim TheBitmap.Datas(0, NewWidth, NewHeight)
  Case 24: ReDim TheBitmap.Datas(2, NewWidth, NewHeight)
 End Select
 TheBitmap.Dimensions.X = NewWidth
 TheBitmap.Dimensions.Y = NewHeight

 BitMap2D_Blt TmpBitMap, TheBitmap, BltRect, BltRect, BltFlag

 If (NewWidth > PreservedWidth) Then
  BltRect = Rect2DInput((PreservedWidth + 1), 0, NewWidth, NewHeight)
  BitMap2D_ColorFill TheBitmap, BltRect, TheBitmap.BackGroundColor
 End If

 If (NewHeight > PreservedHeight) Then
  BltRect = Rect2DInput(0, (PreservedHeight + 1), PreservedWidth, NewHeight)
  BitMap2D_ColorFill TheBitmap, BltRect, TheBitmap.BackGroundColor
 End If

 BitMap2D_Delete TmpBitMap

End Function
Function BitMap2D_Resample(TheBitmap As BitMap2D, NewWidth%, NewHeight%, PixelFilter As K3DE_TEXELS_FILTER_MODES)

 '##################################################################################
 '##################################################################################
 '###
 '###  FUNCTION    : BiMap2D_Resample
 '###
 '###  DESCRIPTION : Resample the given bitmap with a new width & height,
 '###                using a selected pixel filter (interpolation).
 '###
 '##################################################################################
 '##################################################################################

 If (BitMap2D_IsValid(TheBitmap) = False) Then Exit Function
 If ((NewWidth < MinBitMapWidth) Or (NewWidth > MaxBitMapWidth)) Then Exit Function
 If ((NewHeight < MinBitMapHeight) Or (NewHeight > MaxBitMapHeight)) Then Exit Function

 Dim TmpBitMap As BitMap2D, BltFlag As BltFlags, BltRect As Rect2D

 With TmpBitMap
  .Label = "Temporary !"
  .Dimensions = TheBitmap.Dimensions
  .BitsDepth = TheBitmap.BitsDepth
  .BackGroundColor = TheBitmap.BackGroundColor
  If (.BitsDepth = 8) Then
   ReDim .Datas(0, .Dimensions.X, .Dimensions.Y)
  ElseIf (.BitsDepth = 24) Then
   ReDim .Datas(2, .Dimensions.X, .Dimensions.Y)
  End If
 End With

 BltFlag.AlphaFlag = 1
 BltRect = Rect2DInput(0, 0, TheBitmap.Dimensions.X, TheBitmap.Dimensions.Y)
 BitMap2D_Blt TheBitmap, TmpBitMap, BltRect, BltRect, BltFlag

 Select Case TheBitmap.BitsDepth
  Case 8:  ReDim TheBitmap.Datas(0, NewWidth, NewHeight)
  Case 24: ReDim TheBitmap.Datas(2, NewWidth, NewHeight)
 End Select
 TheBitmap.Dimensions.X = NewWidth
 TheBitmap.Dimensions.Y = NewHeight

 BltFlag.Stretch = True: BltFlag.PixelFilter = PixelFilter
 BitMap2D_Blt TmpBitMap, TheBitmap, BltRect, Rect2DInput(0, 0, NewWidth, NewHeight), BltFlag

 BitMap2D_Delete TmpBitMap

End Function
