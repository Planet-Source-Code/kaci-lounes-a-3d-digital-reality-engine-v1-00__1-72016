Attribute VB_Name = "BAS_Mathematics"

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
'###  MODULE      : BAS_Mathematics.BAS
'###
'###  DESCRIPTION : Includ the mathematics operations on the following
'###                data-structures:
'###
'###                ColorRGB, Matrix4x4, Point2D, Rect2D, Vector3D
'###
'##################################################################################
'##################################################################################

Option Explicit
Function ColorCheckChannels(Col As ColorRGB) As Byte

 'Codes the color's channels on a single value (Byte data type)

 '0 = No valid color
 '1 = No channels (!)
 '2 = R
 '3 = G
 '4 = B
 '5 = RG
 '6 = GB
 '7 = RB
 '8 = RGB

 If (ColorIsValid(Col) = False) Then Exit Function

 Select Case Col.R
  Case -1:
   Select Case Col.G
    Case -1:
     Select Case Col.B
      Case -1:   ColorCheckChannels = 1
      Case Else: ColorCheckChannels = 4
     End Select
    Case Else:
     Select Case Col.B
      Case -1:   ColorCheckChannels = 3
      Case Else: ColorCheckChannels = 6
     End Select
   End Select
  Case Else:
   Select Case Col.G
    Case -1:
     Select Case Col.B
      Case -1:   ColorCheckChannels = 2
      Case Else: ColorCheckChannels = 7
     End Select
    Case Else:
     Select Case Col.B
      Case -1:   ColorCheckChannels = 5
      Case Else: ColorCheckChannels = 8
     End Select
   End Select
 End Select

End Function
Function ColorFilterContrast(Col As ColorRGB, ContrastFactor As Single) As ColorRGB

 'Gives a contrast effect.

 ColorFilterContrast = ColorLimit(ColorAdd(ColorScale(ColorSubtract(Col, ColorInput(128, 128, 128)), ContrastFactor), ColorInput(128, 128, 128)))

End Function
Function ColorFilterGamma(Col As ColorRGB, GammaFactor As Single) As ColorRGB

 'Gives a gamma-correction effect.

 ColorFilterGamma.R = ((Col.R * AlphaFactor) ^ (1 / GammaFactor)) * 255
 ColorFilterGamma.G = ((Col.G * AlphaFactor) ^ (1 / GammaFactor)) * 255
 ColorFilterGamma.B = ((Col.B * AlphaFactor) ^ (1 / GammaFactor)) * 255

End Function
Function ColorFilterGreyScale(Col As ColorRGB) As ColorRGB

 'Gives the grey-scaled form.

 ColorFilterGreyScale.R = (Col.R + Col.G + Col.B) * OneByThree
 ColorFilterGreyScale.G = ColorFilterGreyScale.R
 ColorFilterGreyScale.B = ColorFilterGreyScale.G

End Function
Function ColorFilterBrightness(Col As ColorRGB, Brightness As Integer) As ColorRGB

 'Do a brightness effect, a simple additive white color.

 ColorFilterBrightness = ColorLimit(ColorAdd(Col, ColorInput(Brightness, Brightness, Brightness)))

End Function
Function ColorFilterMono(Col As ColorRGB) As ColorRGB

 'With an input color, the function maps the monochromatic color (white or black).

 If (Col.R < 192) And (Col.G < 192) And (Col.B < 192) Then
  ColorFilterMono = ColorBlack
 Else
  ColorFilterMono = ColorWhite
 End If

End Function
Function ColorBlack() As ColorRGB

 'ColorBlack.R = 0
 'ColorBlack.G = 0
 'ColorBlack.B = 0

End Function
Function ColorCompare(ColA As ColorRGB, ColB As ColorRGB) As Boolean

 If (ColA.R = ColB.R) Then
  If (ColA.G = ColB.G) Then
   If (ColA.B = ColB.B) Then
    ColorCompare = True
   End If
  End If
 End If

End Function
Function ColorHSLToHex(Col As ColorHSL) As String

 ColorHSLToHex = ColorRGBToHex(ColorHSLToRGB(Col))

End Function
Function ColorHSLToLong(Col As ColorHSL) As Long

 ColorHSLToLong = ColorRGBToLong(ColorHSLToRGB(Col))

End Function
Function ColorHSLToRGB(Col As ColorHSL) As ColorRGB

 Dim A!, B!, C!, D!, E!

 If (Col.S = 0) Then
  If (Col.H = -1) Then ColorHSLToRGB.R = Col.L: ColorHSLToRGB.G = Col.L: ColorHSLToRGB.B = Col.L
 Else
  Col.H = (Col.H Mod 360) / 60: A = Int(Col.H): B = (Col.H - A)
  C = (1 - Col.S) * Col.L: D = (1 - (Col.S * B)) * Col.L: E = (1 - (Col.S * (1 - B))) * Col.L
  Select Case A
   Case 0:
    ColorHSLToRGB.R = Col.L: ColorHSLToRGB.G = E: ColorHSLToRGB.B = C
   Case 1:
    ColorHSLToRGB.R = D: ColorHSLToRGB.G = Col.L: ColorHSLToRGB.B = C
   Case 2:
    ColorHSLToRGB.R = C: ColorHSLToRGB.G = Col.L: ColorHSLToRGB.B = E
   Case 3:
    ColorHSLToRGB.R = C: ColorHSLToRGB.G = D: ColorHSLToRGB.B = Col.L
   Case 4:
    ColorHSLToRGB.R = E: ColorHSLToRGB.G = C: ColorHSLToRGB.B = Col.L
   Case 5:
    ColorHSLToRGB.R = Col.L: ColorHSLToRGB.G = C: ColorHSLToRGB.B = D
  End Select
 End If

End Function
Function ColorInput(Red As Integer, Green As Integer, Blue As Integer) As ColorRGB

 ColorInput.R = Red
 ColorInput.G = Green
 ColorInput.B = Blue

End Function
Function ColorInputHSL(Hue As Single, Saturation As Single, Lightness As Single) As ColorHSL

 ColorInputHSL.H = Hue
 ColorInputHSL.S = Saturation
 ColorInputHSL.L = Lightness

End Function
Function ColorInvert(Col As ColorRGB) As ColorRGB

 ColorInvert.R = (255 - Col.R)
 ColorInvert.G = (255 - Col.G)
 ColorInvert.B = (255 - Col.B)

End Function
Function ColorIsNone(Col As ColorRGB) As Boolean

 If (ColorCompare(Col, ColorNone) = True) Then ColorIsNone = True

End Function
Function ColorIsValid(Col As ColorRGB) As Boolean

 If (ColorCompare(Col, ColorLimit(Col)) = True) Then ColorIsValid = True

End Function
Function ColorLongToHSL(Col As Long) As ColorHSL

 ColorLongToHSL = ColorRGBToHSL(ColorLongToRGB(Col))

End Function
Function ColorNone() As ColorRGB

 ColorNone = ColorInput(-1, -1, -1)

End Function
Function ColorRandomPrimary() As ColorRGB

 Select Case Fix(Rnd * 7)
  Case 0: ColorRandomPrimary = ColorBlack
  Case 1: ColorRandomPrimary = ColorWhite
  Case 2: ColorRandomPrimary = ColorRed
  Case 3: ColorRandomPrimary = ColorGreen
  Case 4: ColorRandomPrimary = ColorBlue
  Case 5: ColorRandomPrimary = ColorYellow
  Case 6: ColorRandomPrimary = ColorMagenta
  Case 7: ColorRandomPrimary = ColorCyan
 End Select

End Function
Function ColorRandom() As ColorRGB

 ColorRandom.R = (Rnd * 255)
 ColorRandom.G = (Rnd * 255)
 ColorRandom.B = (Rnd * 255)

End Function
Function ColorRandomFrom(FromValue As Byte) As ColorRGB

 ColorRandomFrom.R = (FromValue + (Rnd * (255 - FromValue)))
 ColorRandomFrom.G = (FromValue + (Rnd * (255 - FromValue)))
 ColorRandomFrom.B = (FromValue + (Rnd * (255 - FromValue)))

End Function
Function ColorSwap(ColA As ColorRGB, ColB As ColorRGB)

 Dim ColC As ColorRGB

 ColC.R = ColA.R: ColC.G = ColA.G: ColC.B = ColA.B
 ColA.R = ColB.R: ColA.G = ColB.G: ColA.B = ColB.B
 ColB.R = ColC.R: ColB.G = ColC.G: ColB.B = ColC.B

End Function
Function ColorSwapChannels(ColA As ColorRGB, ColB As ColorRGB, SwapRed As Boolean, SwapGreen As Boolean, SwapBlue As Boolean)

 Dim ColC As ColorRGB

 If (SwapRed = True) Then ColC.R = ColA.R: ColA.R = ColB.R: ColB.R = ColC.R
 If (SwapGreen = True) Then ColC.G = ColA.G: ColA.G = ColB.G: ColB.G = ColC.G
 If (SwapBlue = True) Then ColC.B = ColA.B: ColA.B = ColB.B: ColB.B = ColC.B

End Function
Function ColorRGBToLong(Col As ColorRGB) As Long

 'ColorRGBToLong = Col.R + ((Col.G * 256) + (Col.B * 65536))
 ColorRGBToLong = RGB(Col.R, Col.G, Col.B)

End Function
Function ColorRGBToHex(Col As ColorRGB) As String

 ColorRGBToHex = Hex(ColorRGBToLong(Col))

End Function
Function ColorLongToRGB(Col As Long) As ColorRGB

 ColorLongToRGB.R = (Col And 255)
 ColorLongToRGB.G = (Col And 65280) / 256
 ColorLongToRGB.B = (Col And 16711680) / 65536

End Function
Function ColorLongToHex(Col As Long) As String

 ColorLongToHex = Hex(Col)

End Function
Function ColorHexToRGB(Col As String) As ColorRGB

 ColorHexToRGB.R = (CLng(Col) And 255)
 ColorHexToRGB.G = (CLng(Col) And 65280) / 256
 ColorHexToRGB.B = (CLng(Col) And 16711680) / 65536

End Function
Function ColorHexToHSL(Col As String) As ColorHSL

 ColorHexToHSL = ColorRGBToHSL(ColorHexToRGB(Col))

End Function
Function ColorHexToLong(Col As String) As Long

 ColorHexToLong = ColorRGBToLong(ColorHexToRGB(Col))

End Function
Function ColorRGBToHSL(Col As ColorRGB) As ColorHSL

 Dim RR!, GG!, BB!, MaxBright!, MinBright!, Tmp!

 RR = (255 / Col.R): GG = (255 / Col.G): BB = (255 / Col.B)

 If (RR > 0) Then MaxBright = RR
 If (GG > MaxBright) Then MaxBright = GG
 If (BB > MaxBright) Then MaxBright = BB
 ColorRGBToHSL.L = MaxBright

 MinBright = 1
 If (RR < 1) Then MinBright = RR
 If (GG < MinBright) Then MinBright = GG
 If (BB < MinBright) Then MinBright = BB
 If (MaxBright = 0) Then
  ColorRGBToHSL.S = 0
 Else
  ColorRGBToHSL.S = (MaxBright - MinBright) / MaxBright
 End If

 If (ColorRGBToHSL.S = 0) Then
  ColorRGBToHSL.H = -1
 Else
  Tmp = (MaxBright - MinBright)
  If (RR = MaxBright) Then
   ColorRGBToHSL.H = (GG - BB) / Tmp
  ElseIf (GG = MaxBright) Then
   ColorRGBToHSL.H = ((BB - RR) / Tmp) + 2
  ElseIf (BB = MaxBright) Then
   ColorRGBToHSL.H = ((RR - GG) / Tmp) + 4
  End If
  ColorRGBToHSL.H = (ColorRGBToHSL.H * 60)
  If (ColorRGBToHSL.H < 0) Then ColorRGBToHSL.H = (ColorRGBToHSL.H + 360)
 End If

End Function
Function ColorRGBToVector(Col As ColorRGB) As Vector3D

 ColorRGBToVector.X = Col.R
 ColorRGBToVector.Y = Col.G
 ColorRGBToVector.Z = Col.B

End Function
Function ColorVectorToRGB(Col As Vector3D) As ColorRGB

 'No limitations checks...

 ColorVectorToRGB.R = CInt(Col.X)
 ColorVectorToRGB.G = CInt(Col.Y)
 ColorVectorToRGB.B = CInt(Col.Z)

End Function
Function ColorWhite() As ColorRGB

 ColorWhite.R = 255
 ColorWhite.G = 255
 ColorWhite.B = 255

End Function
Function ColorRed() As ColorRGB

 ColorRed.R = 255
 'ColorRed.G = 0
 'ColorRed.B = 0

End Function
Function ColorGreen() As ColorRGB

 'ColorGreen.R = 0
 ColorGreen.G = 255
 'ColorGreen.B = 0

End Function
Function ColorBlue() As ColorRGB

 'ColorBlue.R = 0
 'ColorBlue.G = 0
 ColorBlue.B = 255

End Function
Function ColorYellow() As ColorRGB

 ColorYellow.R = 255
 ColorYellow.G = 255
 'ColorYellow.B = 0

End Function
Function ColorMagenta() As ColorRGB

 ColorMagenta.R = 255
 'ColorMagenta.G = 0
 ColorMagenta.B = 255

End Function
Function ColorCyan() As ColorRGB

 'ColorCyan.R
 ColorCyan.G = 255
 ColorCyan.B = 255

End Function
Function ColorAdd(ColA As ColorRGB, ColB As ColorRGB) As ColorRGB

 ColorAdd.R = (ColA.R + ColB.R)
 ColorAdd.G = (ColA.G + ColB.G)
 ColorAdd.B = (ColA.B + ColB.B)

End Function
Function ColorSubtract(ColA As ColorRGB, ColB As ColorRGB) As ColorRGB

 ColorSubtract.R = (ColA.R - ColB.R)
 ColorSubtract.G = (ColA.G - ColB.G)
 ColorSubtract.B = (ColA.B - ColB.B)

End Function
Function ColorScale(Col As ColorRGB, ScaleFactor As Single) As ColorRGB

 ColorScale.R = (Col.R * ScaleFactor)
 ColorScale.G = (Col.G * ScaleFactor)
 ColorScale.B = (Col.B * ScaleFactor)

End Function
Function ColorMultiply(ColA As ColorRGB, ColB As ColorRGB) As ColorRGB

 ColorMultiply.R = (ColA.R * ColB.R)
 ColorMultiply.G = (ColA.G * ColB.G)
 ColorMultiply.B = (ColA.B * ColB.B)

End Function
Function ColorPower(Col As ColorRGB, PowerFactor As Single) As ColorRGB

 ColorPower.R = (Col.R ^ PowerFactor)
 ColorPower.G = (Col.G ^ PowerFactor)
 ColorPower.B = (Col.B ^ PowerFactor)

End Function
Function ColorInterpolate(ColA As ColorRGB, ColB As ColorRGB, InterpolationFactor As Single) As ColorRGB

 ColorInterpolate.R = ColA.R + ((ColB.R - ColA.R) * InterpolationFactor)
 ColorInterpolate.G = ColA.G + ((ColB.G - ColA.G) * InterpolationFactor)
 ColorInterpolate.B = ColA.B + ((ColB.B - ColA.B) * InterpolationFactor)

End Function
Function ColorAbsorp(ColA As ColorRGB, ColB As ColorRGB) As ColorRGB

 'A important operation in 3D, for calculating the resulting light's
 'color on an intersection point after an absorption operation,
 'the light's energy is splitted into two parts: the resulting
 'visible light (as the function return), and the reste of the
 'energy is transformed to a heat form (of course we skip
 'that...just don't touch the screen !!).

 With ColorAbsorp
  .R = (ColA.R * (ColB.R * AlphaFactor))
  .G = (ColA.G * (ColB.G * AlphaFactor))
  .B = (ColA.B * (ColB.B * AlphaFactor))
 End With

End Function
Function ColorAbsorp2(ColA As Vector3D, ColB As Vector3D) As Vector3D

 '(Vectors version, for the floating precision)

 With ColorAbsorp2
  .X = (ColA.X * (ColB.X * AlphaFactor))
  .Y = (ColA.Y * (ColB.Y * AlphaFactor))
  .Z = (ColA.Z * (ColB.Z * AlphaFactor))
 End With

End Function
Function ColorLimit(Col As ColorRGB) As ColorRGB

 If (Col.R < 0) Then ColorLimit.R = 0 Else If (Col.R > 255) Then ColorLimit.R = 255 Else ColorLimit.R = Col.R
 If (Col.G < 0) Then ColorLimit.G = 0 Else If (Col.G > 255) Then ColorLimit.G = 255 Else ColorLimit.G = Col.G
 If (Col.B < 0) Then ColorLimit.B = 0 Else If (Col.B > 255) Then ColorLimit.B = 255 Else ColorLimit.B = Col.B

End Function
Function ColorIsValidHSL(Col As ColorHSL) As Boolean

 If ((Col.H >= -1) Or (Col.H <= 360)) Then
  If ((Col.S >= 0) Or (Col.S <= 1)) Then
   If ((Col.L >= 0) Or (Col.L <= 1)) Then
    ColorIsValidHSL = True
   End If
  End If
 End If

End Function
Function Point2DAdd(PointA As Point2D, PointB As Point2D) As Point2D

 Point2DAdd.X = (PointA.X + PointB.X)
 Point2DAdd.Y = (PointA.Y + PointB.Y)

End Function
Function Point2DCompare(PointA As Point2D, PointB As Point2D) As Boolean

 If ((PointA.X = PointB.X) And (PointA.Y = PointB.Y)) Then Point2DCompare = True

End Function
Function Point2DDistance(PointA As Point2D, PointB As Point2D) As Single

 Point2DDistance = Point2DMagnitude(Point2DSubtract(PointB, PointA))

End Function
Function Point2DInput(X%, Y%) As Point2D

 Point2DInput.X = X
 Point2DInput.Y = Y

End Function
Function Point2DInRegion(Region As Point2D, Point As Point2D) As Boolean

 If ((Point.X >= 0) And (Point.X <= Region.X)) Then
  If ((Point.Y >= 0) And (Point.X <= Region.Y)) Then
   Point2DInRegion = True
  End If
 End If

End Function
Function Point2DInterpolate(PointA As Point2D, PointB As Point2D, InterpolationFactor As Single) As Point2D

 Point2DInterpolate = Point2DAdd(PointA, Point2DScale(Point2DSubtract(PointB, PointA), InterpolationFactor))

End Function
Function Point2DInvert(Point As Point2D) As Point2D

 Point2DInvert.X = (1 / Point.X)
 Point2DInvert.Y = (1 / Point.Y)

End Function
Function Point2DIsNull(Point As Point2D) As Boolean

 If (Point2DCompare(Point, Point2DNull) = True) Then Point2DIsNull = True

End Function
Function Point2DMagnitude(Point As Point2D) As Single

 Point2DMagnitude = Sqr((Point.X ^ 2) + (Point.Y ^ 2))

End Function
Function Point2DMultiply(PointA As Point2D, PointB As Point2D) As Point2D

 Point2DMultiply.X = (PointA.X * PointB.X)
 Point2DMultiply.Y = (PointA.Y * PointB.Y)

End Function
Function Point2DNull() As Point2D

End Function
Function Point2DPower(Point As Point2D, PowerFactor As Point2D) As Point2D

 Point2DPower.X = (Point.X ^ PowerFactor)
 Point2DPower.Y = (Point.Y ^ PowerFactor)

End Function
Function Point2DRandom(MinPoint As Point2D, MaxPoint As Point2D) As Point2D

 Point2DRandom = Point2DAdd(MinPoint, Point2DScale(Point2DSubtract(MaxPoint, MinPoint), Rnd))

End Function
Function Point2DRotate(Point As Point2D, Angle As Single) As Point2D

 Dim Sinus As Single: Sinus = Sin(Angle)
 Dim Cosinus As Single: Cosinus = Cos(Angle)
 Point2DRotate.X = (Point.X * Cosinus) - (Point.Y * Sinus)
 Point2DRotate.Y = (Point.X * Sinus) + (Point.Y * Cosinus)

End Function
Function Point2DScale(Point As Point2D, ScaleFactor As Single) As Point2D

 Point2DScale.X = (Point.X * ScaleFactor)
 Point2DScale.Y = (Point.Y * ScaleFactor)

End Function
Function Point2DSubtract(PointA As Point2D, PointB As Point2D) As Point2D

 Point2DSubtract.X = (PointA.X - PointB.X)
 Point2DSubtract.Y = (PointA.Y - PointB.Y)

End Function
Function Point2DSwap(PointA As Point2D, PointB As Point2D)

 Dim PointC As Point2D
 PointC.X = PointA.X: PointA.X = PointB.X: PointB.X = PointC.X
 PointC.Y = PointA.Y: PointA.Y = PointB.Y: PointB.Y = PointC.Y

End Function
Function Rect2DCompare(RectA As Rect2D, RectB As Rect2D) As Boolean

 If (RectA.X1 = RectB.X1) Then
  If (RectA.Y1 = RectB.Y1) Then
   If (RectA.X2 = RectB.X2) Then
    If (RectA.Y2 = RectB.Y2) Then
     Rect2DCompare = True
    End If
   End If
  End If
 End If

End Function
Function Rect2DInput(X1%, Y1%, X2%, Y2%) As Rect2D

 Rect2DInput.X1 = X1: Rect2DInput.Y1 = Y1: Rect2DInput.X2 = X2: Rect2DInput.Y2 = Y2

End Function
Function Rect2DIsInsideRegion(Rect As Rect2D, Region As Point2D) As Boolean

 If (Rect2DIsEmpty(Rect) = False) Then
  If (Rect.X1 >= 0) Then
   If (Rect.Y1 >= 0) Then
    If (Rect.X2 <= Region.X) Then
     If (Rect.Y2 <= Region.Y) Then
      Rect2DIsInsideRegion = True
     End If
    End If
   End If
  End If
 End If

End Function
Function Rect2DIntersect(RectA As Rect2D, RectB As Rect2D) As Rect2D

 'test de validité, de null
 'test rect rect intersection

End Function
Function Rect2DDiagonalLength(Rect As Rect2D) As Single

 Dim Diag As Point2D: Diag = Rect2DDiagonal(Rect)
 Dim XX&, YY&: XX = Diag.X: YY = Diag.Y
 Rect2DDiagonalLength = Sqr((XX * XX) + (YY * YY))

End Function
Function Rect2DSwap(RectA As Rect2D, RectB As Rect2D)

 Dim RectC As Rect2D

 RectC.X1 = RectA.X1: RectC.X2 = RectA.X2: RectC.Y1 = RectA.Y1: RectC.Y2 = RectA.Y2
 RectA.X1 = RectB.X1: RectA.X2 = RectB.X2: RectA.Y1 = RectB.Y1: RectA.Y2 = RectB.Y2
 RectB.X1 = RectC.X1: RectB.X2 = RectC.X2: RectB.Y1 = RectC.Y1: RectB.Y2 = RectC.Y2

End Function
Function Rect2DIsEmpty(Rect As Rect2D) As Boolean

 Dim Diag As Point2D: Diag = Rect2DDiagonal(Rect)
 If ((Diag.X = 0) And (Diag.Y = 0)) Then Rect2DIsEmpty = True

End Function
Function Rect2DIsValid(Rect As Rect2D) As Boolean

 If (Rect.X2 >= Rect.X1) Then
  If (Rect.Y2 >= Rect.Y1) Then
   Rect2DIsValid = True
  End If
 End If

End Function
Function Rect2DDiagonal(Rect As Rect2D) As Point2D

 If (Rect2DIsValid(Rect) = False) Then Exit Function

 Rect2DDiagonal.X = (Rect.X2 - Rect.X1)
 Rect2DDiagonal.Y = (Rect.Y2 - Rect.Y1)

End Function
Function Rect2DMove(Rect As Rect2D, Distance As Point2D) As Rect2D

 If (Rect2DIsValid(Rect) = False) Then Exit Function

 Rect2DMove.X1 = (Rect.X1 + Distance.X): Rect2DMove.X2 = (Rect.X2 + Distance.X)
 Rect2DMove.Y1 = (Rect.Y1 + Distance.Y): Rect2DMove.Y2 = (Rect.Y2 + Distance.Y)

End Function
Function Rect2DPointTest(Rect As Rect2D, Point As Point2D) As Boolean

 If (Rect2DIsEmpty(Rect) = False) Then
  If (Point.X >= Rect.X1) Then
   If (Point.X <= Rect.X2) Then
    If (Point.Y >= Rect.Y1) Then
     If (Point.Y <= Rect.Y2) Then
      Rect2DPointTest = True
     End If
    End If
   End If
  End If
 End If

End Function
Function Rect2DSetOrigin(Rect As Rect2D, Position As Point2D) As Rect2D

 If (Rect2DIsValid(Rect) = False) Then Exit Function

 Rect2DSetOrigin.X1 = Position.X
 Rect2DSetOrigin.Y1 = Position.Y
 Rect2DSetOrigin.X2 = (Position.X + (Rect.X2 - Rect.X1))
 Rect2DSetOrigin.Y2 = (Position.Y + (Rect.Y2 - Rect.Y1))

End Function
Function MatrixAdd(MatA As Matrix4x4, MatB As Matrix4x4) As Matrix4x4

 With MatrixAdd
  .M11 = (MatA.M11 + MatB.M11): .M12 = (MatA.M12 + MatB.M12)
  .M13 = (MatA.M13 + MatB.M13): .M14 = (MatA.M14 + MatB.M14)
  .M21 = (MatA.M21 + MatB.M21): .M22 = (MatA.M22 + MatB.M22)
  .M23 = (MatA.M23 + MatB.M23): .M24 = (MatA.M24 + MatB.M24)
  .M31 = (MatA.M31 + MatB.M31): .M32 = (MatA.M32 + MatB.M32)
  .M33 = (MatA.M33 + MatB.M33): .M34 = (MatA.M34 + MatB.M34)
  .M41 = (MatA.M41 + MatB.M41): .M42 = (MatA.M42 + MatB.M42)
  .M43 = (MatA.M43 + MatB.M43): .M44 = (MatA.M44 + MatB.M44)
 End With

End Function
Function MatrixCompare(MatA As Matrix4x4, MatB As Matrix4x4) As Boolean

 If (MatA.M11 = MatB.M11) And (MatA.M12 = MatB.M12) And (MatA.M13 = MatB.M13) And (MatA.M14 = MatB.M14) And _
    (MatA.M21 = MatB.M21) And (MatA.M22 = MatB.M22) And (MatA.M23 = MatB.M23) And (MatA.M24 = MatB.M24) And _
    (MatA.M31 = MatB.M31) And (MatA.M32 = MatB.M32) And (MatA.M33 = MatB.M33) And (MatA.M34 = MatB.M34) And _
    (MatA.M41 = MatB.M41) And (MatA.M42 = MatB.M42) And (MatA.M43 = MatB.M43) And (MatA.M44 = MatB.M44) Then MatrixCompare = True

End Function
Function MatrixCopy(MatSrc As Matrix4x4, MatDest As Matrix4x4)

 With MatDest
  .M11 = MatSrc.M11: .M12 = MatSrc.M12: .M13 = MatSrc.M13: .M14 = MatSrc.M14
  .M21 = MatSrc.M21: .M22 = MatSrc.M22: .M23 = MatSrc.M23: .M24 = MatSrc.M24
  .M31 = MatSrc.M31: .M32 = MatSrc.M32: .M33 = MatSrc.M33: .M34 = MatSrc.M34
  .M41 = MatSrc.M41: .M42 = MatSrc.M42: .M43 = MatSrc.M43: .M44 = MatSrc.M44
 End With

End Function
Function MatrixDeterminant(Mat As Matrix4x4) As Single

 With Mat
  MatrixDeterminant = (.M14 * .M23 * .M32 * .M41) - (.M13 * .M24 * .M32 * .M41) - (.M14 * .M22 * .M33 * .M41) + (.M12 * .M24 * .M33 * .M41) + _
                      (.M13 * .M22 * .M34 * .M41) - (.M12 * .M23 * .M34 * .M41) - (.M14 * .M23 * .M31 * .M42) + (.M13 * .M24 * .M31 * .M42) + _
                      (.M14 * .M21 * .M33 * .M42) - (.M11 * .M24 * .M33 * .M42) - (.M13 * .M21 * .M34 * .M42) + (.M11 * .M23 * .M34 * .M42) + _
                      (.M14 * .M22 * .M31 * .M43) - (.M12 * .M24 * .M31 * .M43) - (.M14 * .M21 * .M32 * .M43) + (.M11 * .M24 * .M32 * .M43) + _
                      (.M12 * .M21 * .M34 * .M43) - (.M11 * .M22 * .M34 * .M43) - (.M13 * .M22 * .M31 * .M44) + (.M12 * .M23 * .M31 * .M44) + _
                      (.M13 * .M21 * .M32 * .M44) - (.M11 * .M23 * .M32 * .M44) - (.M12 * .M21 * .M33 * .M44) + (.M11 * .M22 * .M33 * .M44)
 End With

End Function
Function MatrixInvert(Mat As Matrix4x4) As Matrix4x4

 With Mat
  MatrixInvert.M11 = (.M23 * .M34 * .M42) - (.M24 * .M33 * .M42) + (.M24 * .M32 * .M43) - (.M22 * .M34 * .M43) - (.M23 * .M32 * .M44) + (.M22 * .M33 * .M44)
  MatrixInvert.M12 = (.M14 * .M33 * .M42) - (.M13 * .M34 * .M42) - (.M14 * .M32 * .M43) + (.M12 * .M34 * .M43) + (.M13 * .M32 * .M44) - (.M12 * .M33 * .M44)
  MatrixInvert.M13 = (.M13 * .M24 * .M42) - (.M14 * .M23 * .M42) + (.M14 * .M22 * .M43) - (.M12 * .M24 * .M43) - (.M13 * .M22 * .M44) + (.M12 * .M23 * .M44)
  MatrixInvert.M14 = (.M14 * .M23 * .M32) - (.M13 * .M24 * .M32) - (.M14 * .M22 * .M33) + (.M12 * .M24 * .M33) + (.M13 * .M22 * .M34) - (.M12 * .M23 * .M34)
  MatrixInvert.M21 = (.M24 * .M33 * .M41) - (.M23 * .M34 * .M41) - (.M24 * .M31 * .M43) + (.M21 * .M34 * .M43) + (.M23 * .M31 * .M44) - (.M21 * .M33 * .M44)
  MatrixInvert.M22 = (.M13 * .M34 * .M41) - (.M14 * .M33 * .M41) + (.M14 * .M31 * .M43) - (.M11 * .M34 * .M43) - (.M13 * .M31 * .M44) + (.M11 * .M33 * .M44)
  MatrixInvert.M23 = (.M14 * .M23 * .M41) - (.M13 * .M24 * .M41) - (.M14 * .M21 * .M43) + (.M11 * .M24 * .M43) + (.M13 * .M21 * .M44) - (.M11 * .M23 * .M44)
  MatrixInvert.M24 = (.M13 * .M24 * .M31) - (.M14 * .M23 * .M31) + (.M14 * .M21 * .M33) - (.M11 * .M24 * .M33) - (.M13 * .M21 * .M34) + (.M11 * .M23 * .M34)
  MatrixInvert.M31 = (.M22 * .M34 * .M41) - (.M24 * .M32 * .M41) + (.M24 * .M31 * .M42) - (.M21 * .M34 * .M42) - (.M22 * .M31 * .M44) + (.M21 * .M32 * .M44)
  MatrixInvert.M32 = (.M14 * .M32 * .M41) - (.M12 * .M34 * .M41) - (.M14 * .M31 * .M42) + (.M11 * .M34 * .M42) + (.M12 * .M31 * .M44) - (.M11 * .M32 * .M44)
  MatrixInvert.M33 = (.M12 * .M24 * .M41) - (.M14 * .M22 * .M41) + (.M14 * .M21 * .M42) - (.M11 * .M24 * .M42) - (.M12 * .M21 * .M44) + (.M11 * .M22 * .M44)
  MatrixInvert.M34 = (.M14 * .M22 * .M31) - (.M12 * .M24 * .M31) - (.M14 * .M21 * .M32) + (.M11 * .M24 * .M32) + (.M12 * .M21 * .M34) - (.M11 * .M22 * .M34)
  MatrixInvert.M41 = (.M23 * .M32 * .M41) - (.M22 * .M33 * .M41) - (.M23 * .M31 * .M42) + (.M21 * .M33 * .M42) + (.M22 * .M31 * .M43) - (.M21 * .M32 * .M43)
  MatrixInvert.M42 = (.M12 * .M33 * .M41) - (.M13 * .M32 * .M41) + (.M13 * .M31 * .M42) - (.M11 * .M33 * .M42) - (.M12 * .M31 * .M43) + (.M11 * .M32 * .M43)
  MatrixInvert.M43 = (.M13 * .M22 * .M41) - (.M12 * .M23 * .M41) - (.M13 * .M21 * .M42) + (.M11 * .M23 * .M42) + (.M12 * .M21 * .M43) - (.M11 * .M22 * .M43)
  MatrixInvert.M44 = (.M12 * .M23 * .M31) - (.M13 * .M22 * .M31) + (.M13 * .M21 * .M32) - (.M11 * .M23 * .M32) - (.M12 * .M21 * .M33) + (.M11 * .M22 * .M33)
 End With

 Dim MatScale As Matrix4x4, Det As Single

 Det = (1 / MatrixDeterminant(Mat))
 MatScale = MatrixScaling4(VectorInput(Det, Det, Det), Det)

 MatrixInvert = MatrixMultiply(MatrixInvert, MatScale)

End Function
Function MatrixIsIdentity(Mat As Matrix4x4) As Boolean

 If (MatrixCompare(Mat, MatrixIdentity) = True) Then MatrixIsIdentity = True

End Function
Function MatrixIsNull(Mat As Matrix4x4) As Boolean

 If (MatrixCompare(Mat, MatrixNull) = True) Then MatrixIsNull = True

End Function
Function MatrixMultiplyVector(Vec As Vector3D, Mat As Matrix4x4) As Vector3D

 MatrixMultiplyVector.X = (Mat.M11 * Vec.X) + (Mat.M12 * Vec.Y) + (Mat.M13 * Vec.Z) + (Mat.M14)
 MatrixMultiplyVector.Y = (Mat.M21 * Vec.X) + (Mat.M22 * Vec.Y) + (Mat.M23 * Vec.Z) + (Mat.M24)
 MatrixMultiplyVector.Z = (Mat.M31 * Vec.X) + (Mat.M32 * Vec.Y) + (Mat.M33 * Vec.Z) + (Mat.M34)

End Function
Function MatrixNull() As Matrix4x4

End Function
Function MatrixRotation(Axis As Byte, Angle As Single) As Matrix4x4

 ' Note well that the function use only one call of
 ' Sin/Cos functions (in every case of course), few
 ' calculations, few memory.

 With MatrixRotation
  Select Case Axis
   Case 0:
    .M11 = 1
    .M22 = Cos(Angle)
    .M23 = -Sin(Angle)
    .M32 = -.M23
    .M33 = .M22
    .M44 = 1
   Case 1:
    .M11 = Cos(Angle)
    .M13 = Sin(Angle)
    .M22 = 1
    .M31 = -.M13
    .M33 = .M11
    .M44 = 1
   Case 2:
    .M11 = Cos(Angle)
    .M12 = -Sin(Angle)
    .M21 = -.M12
    .M22 = .M11
    .M33 = 1
    .M44 = 1
  End Select
 End With

End Function
Function MatrixScaling(Factor As Vector3D) As Matrix4x4

 MatrixScaling.M11 = Factor.X
 MatrixScaling.M22 = Factor.Y
 MatrixScaling.M33 = Factor.Z
 MatrixScaling.M44 = 1

End Function
Function MatrixScaling4(Factor As Vector3D, W!) As Matrix4x4

 'Used for matrix-inversion

 MatrixScaling4.M11 = Factor.X
 MatrixScaling4.M22 = Factor.Y
 MatrixScaling4.M33 = Factor.Z
 MatrixScaling4.M44 = W

End Function
Function MatrixSubtract(MatA As Matrix4x4, MatB As Matrix4x4) As Matrix4x4

 With MatrixSubtract
  .M11 = (MatA.M11 - MatB.M11): .M12 = (MatA.M12 - MatB.M12)
  .M13 = (MatA.M13 - MatB.M13): .M14 = (MatA.M14 - MatB.M14)
  .M21 = (MatA.M21 - MatB.M21): .M22 = (MatA.M22 - MatB.M22)
  .M23 = (MatA.M23 - MatB.M23): .M24 = (MatA.M24 - MatB.M24)
  .M31 = (MatA.M31 - MatB.M31): .M32 = (MatA.M32 - MatB.M32)
  .M33 = (MatA.M33 - MatB.M33): .M34 = (MatA.M34 - MatB.M34)
  .M41 = (MatA.M41 - MatB.M41): .M42 = (MatA.M42 - MatB.M42)
  .M43 = (MatA.M43 - MatB.M43): .M44 = (MatA.M44 - MatB.M44)
 End With

End Function
Function MatrixSwap(MatA As Matrix4x4, MatB As Matrix4x4)

 Dim MatC As Matrix4x4

 MatrixCopy MatA, MatC
 MatrixCopy MatB, MatA
 MatrixCopy MatC, MatB

End Function
Function MatrixTranspose(Mat As Matrix4x4) As Matrix4x4

 'Swap a 4x4 matrix from rows mode, to colmuns mode (and vise-versa)

 With MatrixTranspose
  .M11 = Mat.M11: .M12 = Mat.M21: .M13 = Mat.M31: .M14 = Mat.M41
  .M21 = Mat.M12: .M22 = Mat.M22: .M23 = Mat.M32: .M24 = Mat.M42
  .M31 = Mat.M13: .M32 = Mat.M23: .M33 = Mat.M33: .M34 = Mat.M43
  .M41 = Mat.M14: .M42 = Mat.M24: .M43 = Mat.M34: .M44 = Mat.M44
 End With

End Function
Function VectorDistance(VecA As Vector3D, VecB As Vector3D) As Single

 VectorDistance = VectorLength(VectorSubtract(VecA, VecB))

End Function
Function VectorAngle(VecA As Vector3D, VecB As Vector3D) As Single

 If (VectorCompare(VecA, VectorNull) = False) Then
  If (VectorCompare(VecB, VectorNull) = False) Then
   VectorAngle = VectorDotProduct(VectorNormalize(VecA), VectorNormalize(VecB))
  End If
 End If

End Function
Function VectorCrossProduct(VecA As Vector3D, VecB As Vector3D) As Vector3D

 VectorCrossProduct.X = ((VecA.Y * VecB.Z) - (VecA.Z * VecB.Y))
 VectorCrossProduct.Y = ((VecA.Z * VecB.X) - (VecA.X * VecB.Z))
 VectorCrossProduct.Z = ((VecA.X * VecB.Y) - (VecA.Y * VecB.X))

End Function
Function VectorDotProduct(VecA As Vector3D, VecB As Vector3D) As Single

 VectorDotProduct = (((VecA.X * VecB.X) + (VecA.Y * VecB.Y)) + (VecA.Z * VecB.Z))

End Function
Function VectorGetNormal(VecA As Vector3D, VecB As Vector3D, VecC As Vector3D) As Vector3D

 VectorGetNormal = VectorNormalize(VectorCrossProduct(VectorSubtract(VecA, VecB), VectorSubtract(VecC, VecB)))

End Function
Function VectorGetNormal2(VecA As Vector3D, VecB As Vector3D, VecC As Vector3D) As Vector3D

 VectorGetNormal2 = VectorNormalize(VectorCrossProduct(VectorSubtract(VecA, VecB), VectorSubtract(VecA, VecC)))

End Function
Function VectorInput(X!, Y!, Z!) As Vector3D

 VectorInput.X = X
 VectorInput.Y = Y
 VectorInput.Z = Z

End Function
Function VectorInterpolate(VecA As Vector3D, VecB As Vector3D, InterpolationFactor As Single) As Vector3D

 VectorInterpolate = VectorAdd(VecA, VectorScale(VectorSubtract(VecB, VecA), InterpolationFactor))

End Function
Function VectorLength(Vec As Vector3D) As Single

 VectorLength = Sqr(((Vec.X * Vec.X) + (Vec.Y * Vec.Y)) + (Vec.Z * Vec.Z))

End Function
Function VectorNormalize(Vec As Vector3D) As Vector3D

 If (VectorCompare(Vec, VectorNull) = False) Then
  VectorNormalize = VectorScale(Vec, (1 / VectorLength(Vec)))
 End If

End Function
Function VectorCompare(VecA As Vector3D, VecB As Vector3D) As Boolean

 If (((VecA.X = VecB.X) And (VecA.Y = VecB.Y)) And (VecA.Z = VecB.Z)) Then VectorCompare = True

End Function
Function VectorGetCenter(VecA As Vector3D, VecB As Vector3D, VecC As Vector3D) As Vector3D

 VectorGetCenter.X = ((((VecA.X + VecB.X) + VecC.X)) * OneByThree)
 VectorGetCenter.Y = ((((VecA.Y + VecB.Y) + VecC.Y)) * OneByThree)
 VectorGetCenter.Z = ((((VecA.Z + VecB.Z) + VecC.Z)) * OneByThree)

End Function
Function VectorSwap(VecA As Vector3D, VecB As Vector3D)

 Dim VecC As Vector3D

 VecC.X = VecA.X: VecC.Y = VecA.Y: VecC.Z = VecA.Z
 VecA.X = VecB.X: VecA.Y = VecB.Y: VecA.Z = VecB.Z
 VecB.X = VecC.X: VecB.Y = VecC.Y: VecB.Z = VecC.Z

End Function
Function VectorReflect(TheNormal As Vector3D, TheInput As Vector3D, TheMagnitude!) As Vector3D

 Dim N As Vector3D: N = VectorNormalize(TheNormal)
 Dim I As Vector3D: I = VectorNormalize(TheInput)

 VectorReflect = VectorSubtract(VectorScale(N, (VectorDotProduct(N, I) * 2)), I)
 VectorReflect = VectorScale(VectorNormalize(VectorReflect), TheMagnitude)

End Function
Function VectorRefract(TheNormal As Vector3D, TheInput As Vector3D, OldN!, NewN!, TheMagnitude!) As Vector3D

 'Snell Descartes's law of refraction.

 Dim N As Vector3D: N = VectorNormalize(TheNormal)
 Dim I As Vector3D: I = VectorNormalize(TheInput)

 VectorRefract = VectorSubtract(VectorScale(VectorAdd(VectorInverse(I), N), (OldN / NewN)), N)
 VectorRefract = VectorScale(VectorNormalize(VectorRefract), TheMagnitude)

End Function
Function VectorNull() As Vector3D

End Function
Function VectorRotate(Vec As Vector3D, Axis As Byte, Angle As Single) As Vector3D

 'Basic rotations (without matrices)

 Dim Sinus!: Sinus = Sin(Angle)
 Dim Cosinus!: Cosinus = Cos(Angle)

 Select Case Axis
  Case 0: 'X rotation, a rotation around the YZ plane
   VectorRotate.X = Vec.X
   VectorRotate.Y = (Cosinus * Vec.Y) - (Sinus * Vec.Z)
   VectorRotate.Z = (Sinus * Vec.Y) + (Cosinus * Vec.Z)
  Case 1: 'Y rotation, a rotation around the XZ plane
   VectorRotate.X = (Cosinus * Vec.X) + (Sinus * Vec.Z)
   VectorRotate.Y = Vec.Y
   VectorRotate.Z = -(Sinus * Vec.X) + (Cosinus * Vec.Z)
  Case 2: 'Z rotation, a rotation around the XY plane
   VectorRotate.X = (Cosinus * Vec.X) - (Sinus * Vec.Y)
   VectorRotate.Y = (Sinus * Vec.X) + (Cosinus * Vec.Y)
   VectorRotate.Z = Vec.Z
 End Select

End Function
Function VectorAdd(VecA As Vector3D, VecB As Vector3D) As Vector3D

 VectorAdd.X = (VecA.X + VecB.X)
 VectorAdd.Y = (VecA.Y + VecB.Y)
 VectorAdd.Z = (VecA.Z + VecB.Z)

End Function
Function VectorScale(Vec As Vector3D, ScaleFactor As Single) As Vector3D

 VectorScale.X = (Vec.X * ScaleFactor)
 VectorScale.Y = (Vec.Y * ScaleFactor)
 VectorScale.Z = (Vec.Z * ScaleFactor)

End Function
Function VectorInverse(Vec As Vector3D) As Vector3D

 VectorInverse.X = -Vec.X
 VectorInverse.Y = -Vec.Y
 VectorInverse.Z = -Vec.Z

End Function
Function VectorSubtract(VecA As Vector3D, VecB As Vector3D) As Vector3D

 'VecB will be in the Position

 VectorSubtract.X = (VecA.X - VecB.X)
 VectorSubtract.Y = (VecA.Y - VecB.Y)
 VectorSubtract.Z = (VecA.Z - VecB.Z)

End Function
Function MatrixIdentity() As Matrix4x4

 ' This is the 'Default' matrix, because we get:
 '   (AMatrix * IdentityMatrix) = AMatrix

 With MatrixIdentity
  .M11 = 1: .M12 = 0: .M13 = 0: .M14 = 0
  .M21 = 0: .M22 = 1: .M23 = 0: .M24 = 0
  .M31 = 0: .M32 = 0: .M33 = 1: .M34 = 0
  .M41 = 0: .M42 = 0: .M43 = 0: .M44 = 1
 End With

End Function
Function MatrixMultiply(MatA As Matrix4x4, MatB As Matrix4x4) As Matrix4x4

 'If two matrices A & B, gives different effects (for example rotation and scale),
 'we use MatrixMultiply to give boths transformations in a single matrix, in other
 'words, matrix multiplication 'combine two matrices, in condition that A*B <> B*A

 With MatrixMultiply
  .M11 = (MatA.M11 * MatB.M11) + (MatA.M21 * MatB.M12) + (MatA.M31 * MatB.M13) + (MatA.M41 * MatB.M14)
  .M12 = (MatA.M12 * MatB.M11) + (MatA.M22 * MatB.M12) + (MatA.M32 * MatB.M13) + (MatA.M42 * MatB.M14)
  .M13 = (MatA.M13 * MatB.M11) + (MatA.M23 * MatB.M12) + (MatA.M33 * MatB.M13) + (MatA.M43 * MatB.M14)
  .M14 = (MatA.M14 * MatB.M11) + (MatA.M24 * MatB.M12) + (MatA.M34 * MatB.M13) + (MatA.M44 * MatB.M14)
  .M21 = (MatA.M11 * MatB.M21) + (MatA.M21 * MatB.M22) + (MatA.M31 * MatB.M23) + (MatA.M41 * MatB.M24)
  .M22 = (MatA.M12 * MatB.M21) + (MatA.M22 * MatB.M22) + (MatA.M32 * MatB.M23) + (MatA.M42 * MatB.M24)
  .M23 = (MatA.M13 * MatB.M21) + (MatA.M23 * MatB.M22) + (MatA.M33 * MatB.M23) + (MatA.M43 * MatB.M24)
  .M24 = (MatA.M14 * MatB.M21) + (MatA.M24 * MatB.M22) + (MatA.M34 * MatB.M23) + (MatA.M44 * MatB.M24)
  .M31 = (MatA.M11 * MatB.M31) + (MatA.M21 * MatB.M32) + (MatA.M31 * MatB.M33) + (MatA.M41 * MatB.M34)
  .M32 = (MatA.M12 * MatB.M31) + (MatA.M22 * MatB.M32) + (MatA.M32 * MatB.M33) + (MatA.M42 * MatB.M34)
  .M33 = (MatA.M13 * MatB.M31) + (MatA.M23 * MatB.M32) + (MatA.M33 * MatB.M33) + (MatA.M43 * MatB.M34)
  .M34 = (MatA.M14 * MatB.M31) + (MatA.M24 * MatB.M32) + (MatA.M34 * MatB.M33) + (MatA.M44 * MatB.M34)
  .M41 = (MatA.M11 * MatB.M41) + (MatA.M21 * MatB.M42) + (MatA.M31 * MatB.M43) + (MatA.M41 * MatB.M44)
  .M42 = (MatA.M12 * MatB.M41) + (MatA.M22 * MatB.M42) + (MatA.M32 * MatB.M43) + (MatA.M42 * MatB.M44)
  .M43 = (MatA.M13 * MatB.M41) + (MatA.M23 * MatB.M42) + (MatA.M33 * MatB.M43) + (MatA.M43 * MatB.M44)
  .M44 = (MatA.M14 * MatB.M41) + (MatA.M24 * MatB.M42) + (MatA.M34 * MatB.M43) + (MatA.M44 * MatB.M44)
 End With

End Function
Function MatrixTranslation(Distance As Vector3D) As Matrix4x4

 MatrixTranslation.M11 = 1
 MatrixTranslation.M14 = Distance.X
 MatrixTranslation.M22 = 1
 MatrixTranslation.M24 = Distance.Y
 MatrixTranslation.M33 = 1
 MatrixTranslation.M34 = Distance.Z
 MatrixTranslation.M44 = 1

End Function
Function MatrixView(VecFrom As Vector3D, VecLookAt As Vector3D, RollAngle As Single) As Matrix4x4

 ' We must use 'Virtual cameras' in 3D graphics, we can specify this in three parts:
 '
 ' - Translation      , Translation = -CameraTranslation
 ' - Orientation3D    , The function can map orientations by a 'LookAt' vector (or a view-ray)
 ' - RollAngle        , This is simply the rotation around the view-ray (or around the screen after the view transformation)

 Dim N As Vector3D, U As Vector3D, V As Vector3D

 N = VectorNormalize(VectorSubtract(VecLookAt, VecFrom))
 U = VectorNormalize(VectorCrossProduct(MatrixMultiplyVector(VectorInput(0, 1, 0), MatrixRotation(2, RollAngle)), N))
 V = VectorCrossProduct(N, U) 'The cross-product give a normalized vector,
                              ' because both input vectors are normalized,
                              '  then we d'ont need to normalize.

 With MatrixView
  .M11 = U.X: .M12 = U.Y: .M13 = U.Z
  .M21 = V.X: .M22 = V.Y: .M23 = V.Z
  .M31 = N.X: .M32 = N.Y: .M33 = N.Z
  .M44 = 1
 End With

 MatrixView = MatrixMultiply(MatrixTranslation(VectorInput(-VecFrom.X, -VecFrom.Y, -VecFrom.Z)), MatrixView)

End Function
Function MatrixRotationByVectors(VecFrom As Vector3D, VecTo As Vector3D) As Matrix4x4

 'Vectors-based matrix rotation, rotate from a free
 'point to another free point (arbitrary rotation or planar rotation).

 'This is very important function in this program, because we use it for:
 ' - Orienting the cameras
 ' - Making the Primitives (the most of theme)

 Dim N As Vector3D, U As Vector3D, V As Vector3D

 N = VectorNormalize(VecTo)
 U = VectorNormalize(VectorCrossProduct(VectorNormalize(VecFrom), N))
 V = VectorCrossProduct(N, U) 'The cross-product gives a normalized vector,
                              'because both input vectors are normalized,
                              'so we don't need to normalize.

 With MatrixRotationByVectors
  .M11 = U.X: .M12 = U.Y: .M13 = U.Z
  .M21 = N.X: .M22 = N.Y: .M23 = N.Z
  .M31 = V.X: .M32 = V.Y: .M33 = V.Z
  .M44 = 1
 End With

End Function
Function MatrixWorld(VecTranslate As Vector3D, VecScale As Vector3D, XPitch!, YYaw!, ZRoll!) As Matrix4x4

 ' The world matrix is a set of: a translation, a scale, and three rotations.
 '
 ' Note: We can use the MatrixRotationByVectors function to orient the 3D object
 '       to another one, but i prefer to use directly the orientation angles.

 Dim MatTrans As Matrix4x4, MatRotat As Matrix4x4, MatScale As Matrix4x4

 MatTrans = MatrixTranslation(VecTranslate)
 MatScale = MatrixScaling(VecScale)
 MatRotat = MatrixMultiply(MatrixMultiply(MatrixRotation(0, XPitch), MatrixRotation(1, YYaw)), MatrixRotation(2, ZRoll))

 MatrixWorld = MatrixMultiply(MatrixMultiply(MatScale, MatRotat), MatTrans)

End Function
