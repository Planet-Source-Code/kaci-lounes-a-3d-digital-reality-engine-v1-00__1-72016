Attribute VB_Name = "BAS_Filtering"

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
'###  MODULE      : BAS_Filtering.BAS
'###
'###  DESCRIPTION : Includ the anit-magnification & anti-minification filters.
'###
'##################################################################################
'##################################################################################

Option Explicit
Function DoMipFiltering(TheBitmap As BitMap2D, TheMipMaps As MipTextures, U!, V!, M!) As ColorRGB

 ' About Mip mapping
 ' =================
 '
 ' Introduced by Lance Williams, 1983
 ' MIP = "MultumIn Parvo" (many things in a small place)
 ' Pre-process input textures by prefiltering it at multiple resolutions.

 Dim MipCol1 As ColorRGB, MipCol2 As ColorRGB, MipLevel1!, MipLevel2!, MipFrac!

 If (M = 1) Then M = (1 - ApproachVal)
 MipLevel1 = ((MipMapsLevel + 2) * M)
 If (Fix(MipLevel1) = (MipMapsLevel + 1)) Then
  MipLevel2 = Fix(MipLevel1 - 1)
  MipFrac = 0
 Else
  MipLevel2 = Fix(MipLevel1 + 1)
  MipFrac = 1 - (MipLevel2 - MipLevel1)
 End If
 MipLevel1 = Fix(MipLevel1)

 Select Case TheTexturesFilter
  Case K3DE_TFM_NEAREST:
   DoMipFiltering = DoTexelFiltering(K3DE_XFM_NOFILTER, TheBitmap, U, V, True)

  Case K3DE_TFM_NEAREST_MIP_NEAREST:
   If (MipLevel1 = 0) Then
    DoMipFiltering = DoTexelFiltering(K3DE_XFM_NOFILTER, TheBitmap, U, V, True)
   Else
    DoMipFiltering = DoTexelFiltering(K3DE_XFM_NOFILTER, TheMipMaps.MipSequance((MipMapsLevel - MipLevel1) + 1), U, V, True)
   End If

  Case K3DE_TFM_NEAREST_MIP_LINEAR:
   If (MipLevel1 = 0) Then
    MipCol1 = DoTexelFiltering(K3DE_XFM_NOFILTER, TheBitmap, U, V, True)
    MipCol2 = DoTexelFiltering(K3DE_XFM_NOFILTER, TheMipMaps.MipSequance(MipMapsLevel), U, V, True)
   Else
    MipCol1 = DoTexelFiltering(K3DE_XFM_NOFILTER, TheMipMaps.MipSequance((MipMapsLevel - MipLevel1) + 1), U, V, True)
    MipCol2 = DoTexelFiltering(K3DE_XFM_NOFILTER, TheMipMaps.MipSequance((MipMapsLevel - MipLevel2) + 1), U, V, True)
   End If
   DoMipFiltering = ColorInterpolate(MipCol1, MipCol2, MipFrac)

  Case K3DE_TFM_FILTERED:
   If (TheTexelsFilter = K3DE_XFM_NOFILTER) Then
    DoMipFiltering = DoTexelFiltering(K3DE_XFM_BILINEAR, TheBitmap, U, V, True)
   Else
    DoMipFiltering = DoTexelFiltering(TheTexelsFilter, TheBitmap, U, V, True)
   End If

  Case K3DE_TFM_FILTERED_MIP_NEAREST:
   If (MipLevel1 = 0) Then
    If (TheTexelsFilter = K3DE_XFM_NOFILTER) Then
     DoMipFiltering = DoTexelFiltering(K3DE_XFM_BILINEAR, TheBitmap, U, V, True)
    Else
     DoMipFiltering = DoTexelFiltering(TheTexelsFilter, TheBitmap, U, V, True)
    End If
   Else
    If (TheTexelsFilter = K3DE_XFM_NOFILTER) Then
     DoMipFiltering = DoTexelFiltering(K3DE_XFM_BILINEAR, TheMipMaps.MipSequance((MipMapsLevel - MipLevel1) + 1), U, V, True)
    Else
     DoMipFiltering = DoTexelFiltering(TheTexelsFilter, TheMipMaps.MipSequance((MipMapsLevel - MipLevel1) + 1), U, V, True)
    End If
   End If

  Case K3DE_TFM_FILTERED_MIP_LINEAR:
   If (MipLevel1 = 0) Then
    If (TheTexelsFilter = K3DE_XFM_NOFILTER) Then
     MipCol1 = DoTexelFiltering(K3DE_XFM_BILINEAR, TheBitmap, U, V, True)
     MipCol2 = DoTexelFiltering(K3DE_XFM_BILINEAR, TheMipMaps.MipSequance(MipMapsLevel), U, V, True)
    Else
     MipCol1 = DoTexelFiltering(TheTexelsFilter, TheBitmap, U, V, True)
     MipCol2 = DoTexelFiltering(TheTexelsFilter, TheMipMaps.MipSequance(MipMapsLevel), U, V, True)
    End If
   Else
    If (TheTexelsFilter = K3DE_XFM_NOFILTER) Then
     MipCol1 = DoTexelFiltering(K3DE_XFM_BILINEAR, TheMipMaps.MipSequance((MipMapsLevel - MipLevel1) + 1), U, V, True)
     MipCol2 = DoTexelFiltering(K3DE_XFM_BILINEAR, TheMipMaps.MipSequance((MipMapsLevel - MipLevel2) + 1), U, V, True)
    Else
     MipCol1 = DoTexelFiltering(TheTexelsFilter, TheMipMaps.MipSequance((MipMapsLevel - MipLevel1) + 1), U, V, True)
     MipCol2 = DoTexelFiltering(TheTexelsFilter, TheMipMaps.MipSequance((MipMapsLevel - MipLevel2) + 1), U, V, True)
    End If
   End If
   DoMipFiltering = ColorInterpolate(MipCol1, MipCol2, MipFrac)

 End Select

End Function
Function DoTexelFiltering(FilterType As K3DE_TEXELS_FILTER_MODES, TheBitmap As BitMap2D, U!, V!, Parametric As Boolean) As ColorRGB

 On Error Resume Next 'des erreurs 9: 'Indice en dehors de la plage'

 If (TheBitmap.BitsDepth = 8) Then Exit Function

 Dim IntX%, IntY%, FracX!, FracY!, M&, N&, Fact1!, Fact2!, R1%, G1%, B1%

 If (Parametric = True) Then
  IntX = Fix(U * TheBitmap.Dimensions.X): FracX = ((U * TheBitmap.Dimensions.X) - IntX)
  IntY = Fix(V * TheBitmap.Dimensions.Y): FracY = ((V * TheBitmap.Dimensions.Y) - IntY)
 ElseIf (Parametric = False) Then
  IntX = Fix(U): FracX = (U - IntX)
  IntY = Fix(V): FracY = (V - IntY)
 End If

 Select Case FilterType

  Case K3DE_XFM_NOFILTER:

   DoTexelFiltering.R = TheBitmap.Datas(0, IntX, IntY)
   DoTexelFiltering.G = TheBitmap.Datas(1, IntX, IntY)
   DoTexelFiltering.B = TheBitmap.Datas(2, IntX, IntY)

  Case K3DE_XFM_BILINEAR:

   Dim R2%, R3%, R4%, IR1%, IR2%
   Dim G2%, G3%, G4%, IG1%, IG2%
   Dim B2%, B3%, B4%, IB1%, IB2%

   R1 = TheBitmap.Datas(0, IntX, IntY)
   G1 = TheBitmap.Datas(1, IntX, IntY)
   B1 = TheBitmap.Datas(2, IntX, IntY)

   R2 = TheBitmap.Datas(0, (IntX + 1), IntY)
   G2 = TheBitmap.Datas(1, (IntX + 1), IntY)
   B2 = TheBitmap.Datas(2, (IntX + 1), IntY)

   R3 = TheBitmap.Datas(0, IntX, (IntY + 1))
   G3 = TheBitmap.Datas(1, IntX, (IntY + 1))
   B3 = TheBitmap.Datas(2, IntX, (IntY + 1))

   R4 = TheBitmap.Datas(0, (IntX + 1), (IntY + 1))
   G4 = TheBitmap.Datas(1, (IntX + 1), (IntY + 1))
   B4 = TheBitmap.Datas(2, (IntX + 1), (IntY + 1))

   IR1 = (FracY * R3) + ((1 - FracY) * R1)
   IG1 = (FracY * G3) + ((1 - FracY) * G1)
   IB1 = (FracY * B3) + ((1 - FracY) * B1)

   IR2 = (FracY * R4) + ((1 - FracY) * R2)
   IG2 = (FracY * G4) + ((1 - FracY) * G2)
   IB2 = (FracY * B4) + ((1 - FracY) * B2)

   DoTexelFiltering.R = (FracX * IR2) + ((1 - FracX) * IR1)
   DoTexelFiltering.G = (FracX * IG2) + ((1 - FracX) * IG1)
   DoTexelFiltering.B = (FracX * IB2) + ((1 - FracX) * IB1)

  Case Else:

   For M = -1 To 2

    Select Case FilterType
     Case K3DE_XFM_BELL:
          Fact1 = FilterFunc_Bell(M - FracY)
     Case K3DE_XFM_GAUSSIAN:
          Fact1 = FilterFunc_Gaussian(M - FracY)
     Case K3DE_XFM_CUBIC_SPLINE_B:
          Fact1 = FilterFunc_SplineB(M - FracY)
     Case K3DE_XFM_CUBIC_SPLINE_BC:
          Fact1 = FilterFunc_SplineBC(M - FracY)
     Case K3DE_XFM_CUBIC_SPLINE_CARDINAL:
          Fact1 = FilterFunc_SplineCardinal(M - FracY)
    End Select

    For N = -1 To 2

     Select Case FilterType
      Case K3DE_XFM_BELL:
           Fact2 = FilterFunc_Bell(FracX - N)
      Case K3DE_XFM_GAUSSIAN:
           Fact2 = FilterFunc_Gaussian(FracX - N)
      Case K3DE_XFM_CUBIC_SPLINE_B:
           Fact2 = FilterFunc_SplineB(FracX - N)
      Case K3DE_XFM_CUBIC_SPLINE_BC:
           Fact2 = FilterFunc_SplineBC(FracX - N)
      Case K3DE_XFM_CUBIC_SPLINE_CARDINAL:
           Fact2 = FilterFunc_SplineCardinal(FracX - N)
     End Select

     R1 = TheBitmap.Datas(0, (IntX + N), (IntY + M))
     G1 = TheBitmap.Datas(1, (IntX + N), (IntY + M))
     B1 = TheBitmap.Datas(2, (IntX + N), (IntY + M))

     DoTexelFiltering.R = (DoTexelFiltering.R + ((R1 * Fact1) * Fact2))
     DoTexelFiltering.G = (DoTexelFiltering.G + ((G1 * Fact1) * Fact2))
     DoTexelFiltering.B = (DoTexelFiltering.B + ((B1 * Fact1) * Fact2))

    Next N
   Next M

   DoTexelFiltering = ColorLimit(DoTexelFiltering)

 End Select

End Function
Function DoTexelFiltering8(FilterType As K3DE_TEXELS_FILTER_MODES, TheBitmap As BitMap2D, U!, V!, Parametric As Boolean) As Byte

 On Error Resume Next 'des erreurs 9: 'Indice en dehors de la plage'

 If (TheBitmap.BitsDepth = 24) Then Exit Function

 Dim IntX%, IntY%, FracX!, FracY!, M&, N&, Fact1!, Fact2!, R1%

 If (Parametric = True) Then
  IntX = Fix(U * TheBitmap.Dimensions.X): FracX = ((U * TheBitmap.Dimensions.X) - IntX)
  IntY = Fix(V * TheBitmap.Dimensions.Y): FracY = ((V * TheBitmap.Dimensions.Y) - IntY)
 ElseIf (Parametric = False) Then
  IntX = Fix(U): FracX = (U - IntX)
  IntY = Fix(V): FracY = (V - IntY)
 End If

 Select Case FilterType

  Case K3DE_XFM_NOFILTER:

   DoTexelFiltering8 = TheBitmap.Datas(0, IntX, IntY)

  Case K3DE_XFM_BILINEAR:

   Dim R2%, R3%, R4%, IR1%, IR2%

   R1 = TheBitmap.Datas(0, IntX, IntY)
   R2 = TheBitmap.Datas(0, (IntX + 1), IntY)
   R3 = TheBitmap.Datas(0, IntX, (IntY + 1))
   R4 = TheBitmap.Datas(0, (IntX + 1), (IntY + 1))

   IR1 = (FracY * R3) + ((1 - FracY) * R1)
   IR2 = (FracY * R4) + ((1 - FracY) * R2)

   DoTexelFiltering8 = (FracX * IR2) + ((1 - FracX) * IR1)

  Case Else:

   For M = -1 To 2

    Select Case FilterType
     Case K3DE_XFM_BELL:
          Fact1 = FilterFunc_Bell(M - FracY)
     Case K3DE_XFM_GAUSSIAN:
          Fact1 = FilterFunc_Gaussian(M - FracY)
     Case K3DE_XFM_CUBIC_SPLINE_B:
          Fact1 = FilterFunc_SplineB(M - FracY)
     Case K3DE_XFM_CUBIC_SPLINE_BC:
          Fact1 = FilterFunc_SplineBC(M - FracY)
     Case K3DE_XFM_CUBIC_SPLINE_CARDINAL:
          Fact1 = FilterFunc_SplineCardinal(M - FracY)
    End Select

    For N = -1 To 2

     Select Case FilterType
      Case K3DE_XFM_BELL:
           Fact2 = FilterFunc_Bell(FracX - N)
      Case K3DE_XFM_GAUSSIAN:
           Fact2 = FilterFunc_Gaussian(FracX - N)
      Case K3DE_XFM_CUBIC_SPLINE_B:
           Fact2 = FilterFunc_SplineB(FracX - N)
      Case K3DE_XFM_CUBIC_SPLINE_BC:
           Fact2 = FilterFunc_SplineBC(FracX - N)
      Case K3DE_XFM_CUBIC_SPLINE_CARDINAL:
           Fact2 = FilterFunc_SplineCardinal(FracX - N)
     End Select

     R1 = TheBitmap.Datas(0, (IntX + N), (IntY + M))

     DoTexelFiltering8 = (DoTexelFiltering8 + ((R1 * Fact1) * Fact2))

    Next N
   Next M

   If (DoTexelFiltering8 < 0) Then
    DoTexelFiltering8 = 0
   ElseIf (DoTexelFiltering8 > 255) Then
    DoTexelFiltering8 = 255
   End If

 End Select

End Function
Private Function FilterFunc_Gaussian(Val!) As Single

 If (Abs(Val) < KernelSize) Then
  Dim O!: O = (KernelSize / 3.141593)
  FilterFunc_Gaussian = Exp((-Val * Val) / ((O * O) * 2)) * (0.3989423 / O)
 End If

End Function
Private Function FilterFunc_SplineCardinal(Val!) As Single

 Val = Abs(Val)

 If (Val < 1) Then
  FilterFunc_SplineCardinal = (((CubicA + 2) * (Val ^ 3)) - ((CubicA + 3) * (Val ^ 2))) + 1
 ElseIf (Val < 2) Then
  FilterFunc_SplineCardinal = (((CubicA * (Val ^ 3)) - ((5 * CubicA) * (Val ^ 2))) + ((8 * CubicA) * Val)) - (4 * CubicA)
 End If

End Function
Private Function FilterFunc_SplineBC(Val!) As Single

 Val = Abs(Val)

 If (Val < 1) Then
  FilterFunc_SplineBC = ((12 - (9 * CubicB)) - (6 * CubicC)) * (Val ^ 3)
  FilterFunc_SplineBC = FilterFunc_SplineBC + (((-18 + (12 * CubicB)) + (6 * CubicC)) * (Val ^ 2))
  FilterFunc_SplineBC = ((FilterFunc_SplineBC + 6) - (2 * CubicB)) * 0.1666666
 ElseIf (Val < 2) Then
  FilterFunc_SplineBC = (-CubicB - (6 * CubicC)) * (Val ^ 3)
  FilterFunc_SplineBC = (FilterFunc_SplineBC + ((6 * CubicB) + (30 * CubicC)) * (Val ^ 2))
  FilterFunc_SplineBC = (FilterFunc_SplineBC + ((-12 * CubicB) - (48 * CubicC)) * Val)
  FilterFunc_SplineBC = (FilterFunc_SplineBC + ((8 * CubicB) + (24 * CubicC))) * 0.1666666
 End If

End Function
Private Function FilterFunc_SplineB(Val!) As Single

 If (Val < 2) Then
  Dim A!, B!, C!, D!, Tmp!
  Tmp = (Val + 2): If (Tmp > 0) Then A = (Tmp ^ 3)
  Tmp = (Val + 1): If (Tmp > 0) Then B = ((Tmp ^ 3) * 4)
  If (Val > 0) Then C = ((Val ^ 3) * 6)
  Tmp = (Val - 1): If (Tmp > 0) Then D = ((Tmp ^ 3) * 4)
  FilterFunc_SplineB = (((A - B) + (C - D)) * 0.1666666)
 End If

End Function
Private Function FilterFunc_Bell(Val!) As Single

 Val = Abs(Val)

 If (Val < 0.5) Then
  FilterFunc_Bell = 0.75 - (Val ^ 2)
 ElseIf (Val < 1.5) Then
  FilterFunc_Bell = ((Val - 1.5) ^ 2) * 0.5
 End If

End Function
