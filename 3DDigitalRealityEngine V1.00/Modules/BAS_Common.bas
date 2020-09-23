Attribute VB_Name = "BAS_Common"

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
'###  MODULE      : BAS_Common.BAS
'###
'###  DESCRIPTION : Includ a collection of the commonly-used functions.
'###
'##################################################################################
'##################################################################################

Option Explicit

Public Declare Function ShellExecute Lib "SHELL32.DLL" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Global CamWindowMode As Boolean, CamWindowIndex As Long
Global OmniWindowMode As Boolean, OmniWindowIndex As Long
Global SpotWindowMode As Boolean, SpotWindowIndex As Long
Global MatWindowMode As Boolean, MatWindowIndex As Long
Global MaterialWindowIndex As Long

'(The map-browser window)
Global TheOutputMap As BitMap2D, Browsed As Boolean, BrowseType As Byte
Function GetAddressLast(TheAddress As Address) As Long

 GetAddressLast = ((TheAddress.Start - 1) + TheAddress.Length)

End Function
Function SetAddressLast(TheAddress As Address, TheLast As Long)

 TheAddress.Length = (TheLast - (TheAddress.Start + 1))

End Function
Function ClipRay(RX1!, RY1!, RZ1!, RX2!, RY2!, RZ2!, X1!, Y1!, Z1!, X2!, Y2!, Z2!, OutX1!, OutY1!, OutZ1!, OutX2!, OutY2!, OutZ2!) As Boolean

 'Liang-Barsky parametric line-clipping algorithm (1984) (3D version)

 Dim PX1!, PY1!, PZ1!, PX2!, PY2!, PZ2!
 Dim U1!, U2!, DX!, Dy!, Dz!, P!, Q!, R!, Temp!, CT As Byte

 OutX1 = 0.123456: OutY1 = OutX1: OutZ1 = OutX1
 OutX2 = OutX1: OutY2 = OutX1: OutZ2 = OutX1
 ClipRay = True

 If (RX1 > RX2) Then Temp = RX1: RX1 = RX2: RX2 = Temp
 If (RY1 > RY2) Then Temp = RY1: RY1 = RY2: RY2 = Temp
 If (RZ1 > RZ2) Then Temp = RZ1: RZ1 = RZ2: RZ2 = Temp

 U1 = 0: U2 = 1
 PX1 = X1: PY1 = Y1: PZ1 = Z1
 PX2 = X2: PY2 = Y2: PZ2 = Z2
 DX = (PX2 - PX1): Dy = (PY2 - PY1): Dz = (PZ2 - PZ1)

 P = -DX: Q = (PX1 - RX1)
 If (P < 0) Then
  R = (Q / P): If (R > U2) Then CT = 1 Else If (R > U1) Then U1 = R
 ElseIf (P > 0) Then
  R = (Q / P): If (R < U1) Then CT = 1 Else If (R < U2) Then U2 = R
 ElseIf (Q < 0) Then
  CT = 1
 End If
 If CT = 0 Then
  P = DX: Q = (RX2 - PX1)
  If (P < 0) Then
   R = (Q / P): If (R > U2) Then CT = 1 Else If (R > U1) Then U1 = R
  ElseIf (P > 0) Then
   R = (Q / P): If (R < U1) Then CT = 1 Else If (R < U2) Then U2 = R
  ElseIf (Q < 0) Then
   CT = 1
  End If
  If CT = 0 Then
   P = -Dy: Q = (PY1 - RY1)
   If (P < 0) Then
    R = (Q / P): If (R > U2) Then CT = 1 Else If (R > U1) Then U1 = R
   ElseIf (P > 0) Then
    R = (Q / P): If (R < U1) Then CT = 1 Else If (R < U2) Then U2 = R
   ElseIf (Q < 0) Then
    CT = 1
   End If
   If CT = 0 Then
    P = Dy: Q = (RY2 - PY1)
    If (P < 0) Then
     R = (Q / P): If (R > U2) Then CT = 1 Else If (R > U1) Then U1 = R
    ElseIf (P > 0) Then
     R = (Q / P): If (R < U1) Then CT = 1 Else If (R < U2) Then U2 = R
    ElseIf (Q < 0) Then
     CT = 1
    End If
    If CT = 0 Then
     P = -Dz: Q = (PZ1 - RZ1)
     If (P < 0) Then
      R = (Q / P): If (R > U2) Then CT = 1 Else If (R > U1) Then U1 = R
     ElseIf (P > 0) Then
      R = (Q / P): If (R < U1) Then CT = 1 Else If (R < U2) Then U2 = R
     ElseIf (Q < 0) Then
      CT = 1
     End If
     If CT = 0 Then
      P = Dz: Q = (RZ2 - PZ1)
      If (P < 0) Then
       R = (Q / P): If (R > U2) Then CT = 1 Else If (R > U1) Then U1 = R
      ElseIf (P > 0) Then
       R = (Q / P): If (R < U1) Then CT = 1 Else If (R < U2) Then U2 = R
      ElseIf (Q < 0) Then
       CT = 1
      End If
      If CT = 0 Then
       If (U2 < 1) Then PX2 = (PX1 + (U2 * DX)): PY2 = (PY1 + (U2 * Dy)): PZ2 = (PZ1 + (U2 * Dz))
       If (U1 > 0) Then PX1 = (PX1 + (U1 * DX)): PY1 = (PY1 + (U1 * Dy)): PZ1 = (PZ1 + (U1 * Dz))
       OutX1 = PX1: OutY1 = PY1: OutZ1 = PZ1
       OutX2 = PX2: OutY2 = PY2: OutZ2 = PZ2
      End If
     End If
    End If
   End If
  End If
 End If

 If ((OutX1 = 0.123456) And (OutY1 = 0.123456) And (OutZ1 = 0.123456)) Then
  If ((OutX2 = 0.123456) And (OutY2 = 0.123456) And (OutZ2 = 0.123456)) Then
   ClipRay = False
  End If
 End If

End Function
Function RayBoxIntersect(TheRay As Ray3D, TheBox As Ray3D) As Boolean

 Dim OX1!, OY1!, OZ1!, OX2!, OY2!, OZ2!
 If (ClipRay(TheBox.Position.X, TheBox.Position.Y, TheBox.Position.Z, TheBox.Direction.X, TheBox.Direction.Y, TheBox.Direction.Z, _
             TheRay.Position.X, TheRay.Position.Y, TheRay.Position.Z, TheRay.Direction.X, TheRay.Direction.Y, TheRay.Direction.Z, _
             OX1, OY1, OZ1, OX2, OY2, OZ2) = True) Then
  RayBoxIntersect = True
 End If

End Function
Function SimplyCrypt(TheString As String) As String

 'Translate in english!!!

 'Fait un simple cryptage grace à l'inversion du code ASCII:
 '(ASCII = American Standard Code for Informations Interchange)
 '(code américain standard pour l'échange d'informations).
 'C'est une simple façon pour crypter des données, pour q'un
 'user ne peut les consulter en les affichent directement l'ors
 'l'ouverture de fichier avec un éditeur comme le notepad ou le wordpad.
 '
 ' Algorithme : NewASCII = (255 - OldASCII)

 Dim CurChar As Long, Char As String, TmpStr As String

 For CurChar = 1 To Len(TheString)
  Char = Chr(255 - Asc(Mid(TheString, CurChar, 1)))
  TmpStr = (TmpStr & Char)
 Next CurChar

 SimplyCrypt = TmpStr

End Function
Function CheckAxe(TheOption1 As OptionButton, TheOption2 As OptionButton, TheOption3 As OptionButton) As Byte

 If (TheOption1.Value = True) Then
  CheckAxe = 0
 ElseIf (TheOption2.Value = True) Then
  CheckAxe = 1
 ElseIf (TheOption3.Value = True) Then
  CheckAxe = 2
 End If

End Function
Function CheckOut(TheCheckBox As CheckBox) As Boolean

 If (TheCheckBox.Value = vbChecked) Then CheckOut = True

End Function
Function ComputeElaspedTime(StartTime As String, EndTime As String) As String

 'I did NOT fix this yet..

 Dim SH&, SM&, SS&, EH&, EM&, ES&, RH&, RM&, RS&

 SH = CLng(Left(StartTime, 2)): SM = CLng(Mid(StartTime, 4, 2)): SS = CLng(Right(StartTime, 2))
 EH = CLng(Left(EndTime, 2)): EM = CLng(Mid(EndTime, 4, 2)): ES = CLng(Right(EndTime, 2))

 RH = (EH - SH): RM = (EM - SM): RS = (ES - SS)

 'If (RM < 0) Then RS = (60 - RS): RM = (RM + 1)
 'If (RH < 0) Then RM = (60 - RM): RH = (RH + 1)

 If ((RH < 0) Or (RM < 0) Or (RH < 0)) Then
  ComputeElaspedTime = "From : " & StartTime & "  To : " & EndTime: Exit Function
 Else
  ComputeElaspedTime = RH & ":" & RM & ":" & RS
 End If

End Function
Sub DisplayAMap(TheBitmap As BitMap2D)

 Load FRM_DisplayMap
 FRM_DisplayMap.MousePointer = 11

 Dim ImageW%, ImageH%, CurX&, CurY&

 ImageW = TheBitmap.Dimensions.X: ImageH = TheBitmap.Dimensions.Y
 FRM_DisplayMap.Picture2.Width = ImageW: FRM_DisplayMap.Picture2.Height = ImageH

 If ((ImageW = 0) And (ImageH = 0)) Then FRM_DisplayMap.Label1.Visible = True: GoTo Jump

 If (TheBitmap.BitsDepth = 24) Then
  For CurY = 0 To ImageH
   For CurX = 0 To ImageW
    FRM_DisplayMap.Picture2.PSet (CurX, CurY), RGB(TheBitmap.Datas(0, CurX, CurY), TheBitmap.Datas(1, CurX, CurY), TheBitmap.Datas(2, CurX, CurY))
   Next CurX
  Next CurY
 ElseIf (TheBitmap.BitsDepth = 8) Then
  For CurY = 0 To ImageH
   For CurX = 0 To ImageW
    FRM_DisplayMap.Picture2.PSet (CurX, CurY), RGB(TheBitmap.Datas(0, CurX, CurY), TheBitmap.Datas(0, CurX, CurY), TheBitmap.Datas(0, CurX, CurY))
   Next CurX
  Next CurY
 End If

 'Ajust scroll bars:
 If (FRM_DisplayMap.Picture2.ScaleWidth > FRM_DisplayMap.Picture3.ScaleWidth) Then
  FRM_DisplayMap.HScroll1.Enabled = True: FRM_DisplayMap.HScroll1.Max = (FRM_DisplayMap.Picture3.ScaleWidth - FRM_DisplayMap.Picture2.ScaleWidth)
 Else
  FRM_DisplayMap.HScroll1.Enabled = False
 End If
 If (FRM_DisplayMap.Picture2.ScaleHeight > FRM_DisplayMap.Picture3.ScaleHeight) Then
  FRM_DisplayMap.VScroll1.Enabled = True: FRM_DisplayMap.VScroll1.Max = (FRM_DisplayMap.Picture3.ScaleHeight - FRM_DisplayMap.Picture2.ScaleHeight)
 Else
  FRM_DisplayMap.VScroll1.Enabled = False
 End If

Jump:

 FRM_DisplayMap.MousePointer = 0
 FRM_DisplayMap.Show 1

End Sub
Function IntegerMax(Number1%, Number2%) As Integer

 If (Number1 > Number2) Then IntegerMax = Number1 Else IntegerMax = Number2

End Function
Function RadToDeg(TheAngle As Single) As Single

 RadToDeg = (TheAngle * (1 / Deg))

End Function
Function DegToRad(TheAngle As Single) As Single

 DegToRad = (TheAngle * Deg)

End Function
Function SingleMin(Number1!, Number2!) As Single

 If (Number1 < Number2) Then SingleMin = Number1 Else SingleMin = Number2

End Function
Function LongMin(Number1&, Number2&) As Long

 If (Number1 < Number2) Then LongMin = Number1 Else LongMin = Number2

End Function
Function IntegerMin(Number1%, Number2%) As Integer

 If (Number1 < Number2) Then IntegerMin = Number1 Else IntegerMin = Number2

End Function
Function LongMax(Number1&, Number2&) As Long

 If (Number1 > Number2) Then LongMax = Number1 Else LongMax = Number2

End Function
Function SingleMax(Number1!, Number2!) As Single

 If (Number1 > Number2) Then SingleMax = Number1 Else SingleMax = Number2

End Function
Function SignedRnd() As Single

 'Use the half propability (Rnd it-self) to sign Rnd as negative number.

 If (Rnd < 0.5) Then SignedRnd = -Rnd Else SignedRnd = Rnd

End Function
Function FileExist(TheFileName$) As Boolean

 On Error Resume Next

 Dim NF%: NF = FreeFile
 Open TheFileName For Input As NF
  FileExist = IIf(Err = 0, True, False)
 Close NF: Err = 0

End Function
Function ValidPath(TheFileName$) As Boolean

 On Error Resume Next

 Dim NF%: NF = FreeFile
 Open TheFileName For Binary As NF
  ValidPath = IIf(Err = 0, True, False)
 Close NF: Err = 0
 If (ValidPath = True) Then Kill TheFileName

End Function
Function ExpScale(Linear!, ExpFactor1!, ExpFactor2!) As Single

 'About
 '=====
 'An exponential scaling function, by giving a linear parametric value (0...1), and the
 'parametrics exponential factors (0...1), depent on this factors, the results is in both
 'direction pos or neg, by using the same value of factors, the result is linear.
 'I find this simply by applying a 2D projection over a 2D triangle (see Exponential.jpg
 'in the program folder), the same reason that a 3D perspective view give 'Not Equispaced'
 'projections of parallal lines, we can use this function in any program need an
 ' 'exponential scaling', I also use it for exponential fog-mapping, for example.

 Dim PointB As Vector3D, PointC As Vector3D
 Dim LinearPoint As Vector3D, ProjPoint As Vector3D
 Dim ExpPoint1 As Vector3D, ExpPoint2 As Vector3D
 Dim D!, U!, V!, T!

 PointB.X = 200: PointB.Y = 200
 PointC.X = -200: PointC.Y = 200

 ExpPoint1 = VectorScale(PointC, ExpFactor1)
 ExpPoint2 = VectorScale(PointB, ExpFactor2)
 LinearPoint = VectorInterpolate(ExpPoint1, ExpPoint2, Linear)

 D = (1 / VectorCrossProduct(PointB, PointC).Z): If (D = 0) Then D = ApproachVal
 U = VectorCrossProduct(LinearPoint, PointC).Z * D
 V = VectorCrossProduct(PointB, LinearPoint).Z * D

 T = (U / (U + V))
 ProjPoint = VectorInterpolate(PointC, PointB, T)
 ExpScale = (VectorDistance(PointC, ProjPoint) / VectorDistance(PointC, PointB))

End Function
Function IsBackFace(X1!, Y1!, X2!, Y2!, X3!, Y3!) As Boolean

 If (VectorCrossProduct(VectorInput((X2 - X1), (Y2 - Y1), 0), VectorInput((X3 - X1), (Y3 - Y1), 0)).Z > 0) Then IsBackFace = True

End Function
Function IsPointInTriangle(X1!, Y1!, X2!, Y2!, X3!, Y3!, PX!, PY!) As Boolean

 ' FUNCTION : IsPointInTriangle
 ' ===========================
 '
 ' RETURNED VALUE: Boolean
 '
 ' 2D point in triangle check (barycentrics).

 Dim CRZ1!, CRZ2!, CRZ3!

 CRZ1 = (((X2 - PX) * (Y3 - PY)) - ((Y2 - PY) * (X3 - PX)))
 CRZ2 = (((X1 - PX) * (Y2 - PY)) - ((Y1 - PY) * (X2 - PX)))
 CRZ3 = (((X3 - PX) * (Y1 - PY)) - ((Y3 - PY) * (X1 - PX)))

 'The point is inside the triangle if the vars (CRZ1, CRZ2 & CRZ3) haves the same sign:
 If ((CRZ1 >= 0) And (CRZ2 >= 0) And (CRZ3 >= 0)) Or _
    ((CRZ1 <= 0) And (CRZ2 <= 0) And (CRZ3 <= 0)) Then IsPointInTriangle = True

End Function
Function FocalDistance(TheAngle!, TheProjection!) As Single

 ' Focal Distance
 ' ==============
 '
 ' Compute the focal distance with a given
 ' FOV angle in radians (0 > FOV! > Pi),
 ' and the half-size of the projection plane
 '
 '                  °
 '                 /|
 '               /  |
 '             /    | V
 '           /      | P
 '         /        | 2
 '       / FOV/2    |
 '      °-----------°
 '      <--- F.D --->

 Dim A!: A = -Sin(TheAngle * 0.5)
 Dim B!: B = -Cos(TheAngle * 0.5)

 FocalDistance = (A / B) * (TheProjection + (B / A))

End Function
Function ClipLine(RX1%, RY1%, RX2%, RY2%, X1%, Y1%, X2%, Y2%, OutX1%, OutY1%, OutX2%, OutY2%) As Boolean

 'Liang-Barsky parametric line clipping algorithm

 Dim PX1!, PY1!, PX2!, PY2!, U1!, U2!, DX!, Dy!, P!, Q!, R!, Temp!, CT As Byte

 If (RX1 > RX2) Then Temp = RX1: RX1 = RX2: RX2 = Temp
 If (RY1 > RY2) Then Temp = RY1: RY1 = RY2: RY2 = Temp

 OutX1 = 0.123456: OutY1 = OutX1: OutX2 = OutX1: OutY2 = OutX1: ClipLine = True

 U1 = 0: U2 = 1: PX1 = X1: PY1 = Y1: PX2 = X2: PY2 = Y2
 DX = (PX2 - PX1): Dy = (PY2 - PY1)

 P = -DX: Q = (PX1 - RX1)
 If (P < 0) Then
  R = (Q / P): If (R > U2) Then CT = 1 Else If (R > U1) Then U1 = R
 ElseIf (P > 0) Then
  R = (Q / P): If (R < U1) Then CT = 1 Else If (R < U2) Then U2 = R
 ElseIf (Q < 0) Then
  CT = 1
 End If
 If (CT = 0) Then
  P = DX: Q = (RX2 - PX1)
  If (P < 0) Then
   R = (Q / P): If (R > U2) Then CT = 1 Else If (R > U1) Then U1 = R
  ElseIf (P > 0) Then
   R = (Q / P): If (R < U1) Then CT = 1 Else If (R < U2) Then U2 = R
  ElseIf (Q < 0) Then
   CT = 1
  End If
  If (CT = 0) Then
   P = -Dy: Q = (PY1 - RY1)
   If (P < 0) Then
    R = (Q / P): If (R > U2) Then CT = 1 Else If (R > U1) Then U1 = R
   ElseIf (P > 0) Then
    R = (Q / P): If (R < U1) Then CT = 1 Else If (R < U2) Then U2 = R
   ElseIf (Q < 0) Then
    CT = 1
   End If
   If (CT = 0) Then
    P = Dy: Q = (RY2 - PY1)
    If (P < 0) Then
     R = (Q / P): If (R > U2) Then CT = 1 Else If (R > U1) Then U1 = R
    ElseIf (P > 0) Then
     R = (Q / P): If (R < U1) Then CT = 1 Else If (R < U2) Then U2 = R
    ElseIf (Q < 0) Then
     CT = 1
    End If
    If (CT = 0) Then
     If (U2 < 1) Then PX2 = (PX1 + (U2 * DX)): PY2 = (PY1 + (U2 * Dy))
     If (U1 > 0) Then PX1 = (PX1 + (U1 * DX)): PY1 = (PY1 + (U1 * Dy))
     OutX1 = PX1: OutY1 = PY1: OutX2 = PX2: OutY2 = PY2
    End If
   End If
  End If
 End If

 If ((OutX1 = 0.123456) And (OutY1 = 0.123456)) Then
  If ((OutX2 = 0.123456) And (OutY2 = 0.123456)) Then
   ClipLine = False
  End If
 End If

End Function
Sub ExtractionSortInteger(TheArray() As Integer, Ascending As Boolean)

 'Extraction sort algorithm for Integer data type
 'Acsending: low to high

 Dim Elem1&, Elem2&, Elem3&, TmpInt%

 If (Ascending = True) Then
  For Elem1 = LBound(TheArray) To UBound(TheArray)
   Elem3 = Elem1
   For Elem2 = Elem1 To UBound(TheArray)
    If (TheArray(Elem3) > TheArray(Elem2)) Then Elem3 = Elem2
   Next Elem2
   TmpInt = TheArray(Elem3): TheArray(Elem3) = TheArray(Elem1): TheArray(Elem1) = TmpInt
  Next Elem1
 Else
  For Elem1 = LBound(TheArray) To UBound(TheArray)
   Elem3 = Elem1
   For Elem2 = Elem1 To UBound(TheArray)
    If (TheArray(Elem3) < TheArray(Elem2)) Then Elem3 = Elem2 'Change the way of sorting
   Next Elem2                                                 'by changing the sign ( > or < )
   TmpInt = TheArray(Elem3): TheArray(Elem3) = TheArray(Elem1): TheArray(Elem1) = TmpInt
  Next Elem1
 End If

End Sub
Sub ExtractionSortLong(TheArray() As Long, Ascending As Boolean)

 'Extraction sort algorithm for Long data type
 'Acsending: low to high

 Dim Elem1&, Elem2&, Elem3&, TmpLng&

 If (Ascending = True) Then
  For Elem1 = LBound(TheArray) To UBound(TheArray)
   Elem3 = Elem1
   For Elem2 = Elem1 To UBound(TheArray)
    If (TheArray(Elem3) > TheArray(Elem2)) Then Elem3 = Elem2
   Next Elem2
   TmpLng = TheArray(Elem3): TheArray(Elem3) = TheArray(Elem1): TheArray(Elem1) = TmpLng
  Next Elem1
 Else
  For Elem1 = LBound(TheArray) To UBound(TheArray)
   Elem3 = Elem1
   For Elem2 = Elem1 To UBound(TheArray)
    If (TheArray(Elem3) < TheArray(Elem2)) Then Elem3 = Elem2 'Change the way of sorting
   Next Elem2                                                 'by changing the sign ( > or < )
   TmpLng = TheArray(Elem3): TheArray(Elem3) = TheArray(Elem1): TheArray(Elem1) = TmpLng
  Next Elem1
 End If

End Sub
Sub ExtractionSortSingle(TheArray() As Single, Ascending As Boolean)

 'Extraction sort algorithm for Single data type
 'Acsending: low to high

 Dim Elem1&, Elem2&, Elem3&, TmpSng!

 If (Ascending = True) Then
  For Elem1 = LBound(TheArray) To UBound(TheArray)
   Elem3 = Elem1
   For Elem2 = Elem1 To UBound(TheArray)
    If (TheArray(Elem3) > TheArray(Elem2)) Then Elem3 = Elem2
   Next Elem2
   TmpSng = TheArray(Elem3): TheArray(Elem3) = TheArray(Elem1): TheArray(Elem1) = TmpSng
  Next Elem1
 Else
  For Elem1 = LBound(TheArray) To UBound(TheArray)
   Elem3 = Elem1
   For Elem2 = Elem1 To UBound(TheArray)
    If (TheArray(Elem3) < TheArray(Elem2)) Then Elem3 = Elem2 'Change the way of sorting
   Next Elem2                                                 'by changing the sign ( > or < )
   TmpSng = TheArray(Elem3): TheArray(Elem3) = TheArray(Elem1): TheArray(Elem1) = TmpSng
  Next Elem1
 End If

End Sub
Sub ExtractionSortSingleLong(TheSingleArray() As Single, TheLongArray() As Long, Ascending As Boolean)

 'Extraction sort algorithm for Single data type
 'Acsending: low to high

 Dim Elem1&, Elem2&, Elem3&, TmpSng!, TmpLng&

 If (Ascending = True) Then
  For Elem1 = LBound(TheSingleArray) To UBound(TheSingleArray)
   Elem3 = Elem1
   For Elem2 = Elem1 To UBound(TheSingleArray)
    If (TheSingleArray(Elem3) > TheSingleArray(Elem2)) Then Elem3 = Elem2
   Next Elem2
   TmpSng = TheSingleArray(Elem3): TheSingleArray(Elem3) = TheSingleArray(Elem1): TheSingleArray(Elem1) = TmpSng
   TmpLng = TheLongArray(Elem3): TheLongArray(Elem3) = TheLongArray(Elem1): TheLongArray(Elem1) = TmpLng
  Next Elem1
 Else
  For Elem1 = LBound(TheSingleArray) To UBound(TheSingleArray)
   Elem3 = Elem1
   For Elem2 = Elem1 To UBound(TheSingleArray)
    If (TheSingleArray(Elem3) < TheSingleArray(Elem2)) Then Elem3 = Elem2
   Next Elem2
   TmpSng = TheSingleArray(Elem3): TheSingleArray(Elem3) = TheSingleArray(Elem1): TheSingleArray(Elem1) = TmpSng
   TmpLng = TheLongArray(Elem3): TheLongArray(Elem3) = TheLongArray(Elem1): TheLongArray(Elem1) = TmpLng
  Next Elem1
 End If

End Sub
