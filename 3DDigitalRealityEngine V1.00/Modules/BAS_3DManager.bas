Attribute VB_Name = "BAS_3DManager"

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
'###  MODULE      : BAS_Manager.BAS
'###
'###  DESCRIPTION : Includ the management functions for the 3D data-base.
'###
'##################################################################################
'##################################################################################

Option Explicit

'3D DATA BASE:

Global MaxVertices As Long
Global MaxFaces As Long
Global MaxMeshs As Long
Global MaxLights As Long
Global MaxCameras As Long
Global MaxSplines As Long

'Geometry, lights and cameras:
Global TheVertices() As Vertex3D
Global TheFaces() As Face3D
Global TheMeshs() As Mesh3D
Global TheMaterials() As Material
Global TheSphereLights() As SphereLight3D
Global TheConeLights() As ConeLight3D
Global TheCameras() As Camera3D
Global TheSplines() As Spline3D

'Textures:
Global TheAlphaTextures() As BitMap2D, TheAlphaUsed() As Boolean, TheAlphaIndexs() As Long
Global TheColorTextures() As BitMap2D, TheColorMips() As MipTextures, TheColorUsed() As Boolean, TheColorIndexs() As Long
Global TheReflectionTextures() As BitMap2D, TheReflectionUsed() As Boolean, TheReflectionIndexs() As Long
Global TheRefractionTextures() As BitMap2D, TheRefractionUsed() As Boolean, TheRefractionIndexs() As Long
Global TheRefractionNTextures() As BitMap2D, TheRefractionNUsed() As Boolean, TheRefractionNIndexs() As Long

Global TheMeshsCount&, TheSphereLightsCount&, TheConeLightsCount&, TheCamerasCount&
Sub Scene3D_Clear()

 Mesh3D_Clear
 SphereLight3D_Clear
 ConeLight3D_Clear
 Camera3D_Clear

End Sub
Function TEX_Alpha_Add() As Long

On Error GoTo Eror

 Dim CurIndex As Long, TexIndex As Long

 For CurIndex = 0 To MaxMeshs
  If (TheAlphaUsed(CurIndex) = False) Then
   TexIndex = CurIndex: Exit For
  End If
 Next CurIndex

 TheAlphaTextures(TexIndex).Label = "AlphaTexture_" & (TheMeshsCount + 1)
 TEX_Alpha_Add = TexIndex: TheAlphaUsed(TexIndex) = True
 ReDim Preserve TheAlphaIndexs(UBound(TheAlphaIndexs) + 1)
 TheAlphaIndexs(UBound(TheAlphaIndexs)) = TexIndex
 ExtractionSortLong TheAlphaIndexs(), True

Eror:

 If (Err.Number = 9) Then
  ReDim Preserve TheAlphaIndexs(0)
  TheAlphaIndexs(0) = TexIndex
 End If

End Function
Function TEX_Alpha_Clear()

On Error Resume Next

 Dim CurPos As Long

 For CurPos = 0 To MaxMeshs
  BitMap2D_Delete TheAlphaTextures(CurPos)
  TheAlphaUsed(CurPos) = False
 Next CurPos

 Erase TheAlphaIndexs()

End Function
Function TEX_Alpha_Remove(TheIndex As Long)

 If ((TheIndex < 0) Or (TheIndex > MaxMeshs)) Then Exit Function
 If (TheAlphaUsed(TheIndex) = False) Then Exit Function

 BitMap2D_Delete TheAlphaTextures(TheIndex)
 TheAlphaUsed(TheIndex) = False

 Dim CurPos As Long
 For CurPos = 0 To UBound(TheAlphaIndexs)
  If (TheAlphaIndexs(CurPos) = TheIndex) Then
   TheAlphaIndexs(CurPos) = (MaxMeshs + 1): Exit For
  End If
 Next CurPos

 If (UBound(TheAlphaIndexs) = 0) Then
  Erase TheAlphaIndexs()
 Else
  ExtractionSortLong TheAlphaIndexs(), True
  ReDim Preserve TheAlphaIndexs(UBound(TheAlphaIndexs) - 1)
 End If

End Function
Function TEX_Alpha_Set(TheIndex As Long, TheBitmap As BitMap2D)

 If ((TheIndex < 0) Or (TheIndex > MaxMeshs)) Then Exit Function
 If (TheAlphaUsed(TheIndex) = False) Then Exit Function
 If (BitMap2D_IsValid(TheBitmap) = False) Then Exit Function
 If (TheBitmap.BitsDepth <> 8) Then Exit Function

 Dim OldLabel As String
 OldLabel = TheAlphaTextures(TheIndex).Label
 TheAlphaTextures(TheIndex) = TheBitmap
 TheAlphaTextures(TheIndex).Label = OldLabel

End Function
Function TEX_Color_Add() As Long

On Error GoTo Eror

 Dim CurIndex As Long, TexIndex As Long

 For CurIndex = 0 To MaxMeshs
  If (TheColorUsed(CurIndex) = False) Then
   TexIndex = CurIndex: Exit For
  End If
 Next CurIndex

 TheColorTextures(TexIndex).Label = "ColorTexture_" & (TheMeshsCount + 1)
 TEX_Color_Add = TexIndex: TheColorUsed(TexIndex) = True
 ReDim TheColorMips(TexIndex).MipSequance(MipMapsLevel)
 ReDim Preserve TheColorIndexs(UBound(TheColorIndexs) + 1)
 TheColorIndexs(UBound(TheColorIndexs)) = TexIndex
 ExtractionSortLong TheColorIndexs(), True

Eror:

 If (Err.Number = 9) Then
  ReDim Preserve TheColorIndexs(0)
  TheColorIndexs(0) = TexIndex
 End If

End Function
Function TEX_Color_Clear()

On Error Resume Next

 Dim CurPos As Long

 For CurPos = 0 To MaxMeshs
  BitMap2D_Delete TheColorTextures(CurPos)
  Erase TheColorMips(CurPos).MipSequance()
  TheColorUsed(CurPos) = False
 Next CurPos

 Erase TheColorIndexs()

End Function
Function TEX_Color_Remove(TheIndex As Long)

 If ((TheIndex < 0) Or (TheIndex > MaxMeshs)) Then Exit Function
 If (TheColorUsed(TheIndex) = False) Then Exit Function

 BitMap2D_Delete TheColorTextures(TheIndex)
 Erase TheColorMips(TheIndex).MipSequance()
 TheColorUsed(TheIndex) = False

 Dim CurPos As Long
 For CurPos = 0 To UBound(TheColorIndexs)
  If (TheColorIndexs(CurPos) = TheIndex) Then
   TheColorIndexs(CurPos) = (MaxMeshs + 1): Exit For
  End If
 Next CurPos

 If (UBound(TheColorIndexs) = 0) Then
  Erase TheColorIndexs()
 Else
  ExtractionSortLong TheColorIndexs(), True
  ReDim Preserve TheColorIndexs(UBound(TheColorIndexs) - 1)
 End If

End Function
Function TEX_Color_Set(TheIndex As Long, TheBitmap As BitMap2D)

 If ((TheIndex < 0) Or (TheIndex > MaxMeshs)) Then Exit Function
 If (TheColorUsed(TheIndex) = False) Then Exit Function
 If (BitMap2D_IsValid(TheBitmap) = False) Then Exit Function
 If (TheBitmap.BitsDepth <> 24) Then Exit Function

 Dim OldLabel As String
 OldLabel = TheColorTextures(TheIndex).Label
 TheColorTextures(TheIndex) = TheBitmap
 TheColorTextures(TheIndex).Label = OldLabel

 'CREATE THE MIP-MAPS:
 If ((TheTexturesFilter = K3DE_TFM_NEAREST_MIP_NEAREST) Or _
     (TheTexturesFilter = K3DE_TFM_NEAREST_MIP_LINEAR) Or _
     (TheTexturesFilter = K3DE_TFM_FILTERED_MIP_NEAREST) Or _
     (TheTexturesFilter = K3DE_TFM_FILTERED_MIP_LINEAR)) Then

  Dim CurMip&, MinW%, MinH%, StpW%, StpH%, NewW%, NewH%
  MinW = (TheBitmap.Dimensions.X * (MipMapsMinPurcent / 100))
  MinH = (TheBitmap.Dimensions.Y * (MipMapsMinPurcent / 100))
  If (MinW < MinBitMapWidth) Then MinW = MinBitMapWidth
  If (MinH < MinBitMapHeight) Then MinW = MinBitMapHeight
  StpW = ((TheBitmap.Dimensions.X - MinW) / (MipMapsLevel + 1))
  StpH = ((TheBitmap.Dimensions.Y - MinH) / (MipMapsLevel + 1))

  NewW = MinW: NewH = MinH
  For CurMip = 0 To MipMapsLevel
   TheColorMips(TheIndex).MipSequance(CurMip) = TheBitmap
   TheColorMips(TheIndex).MipSequance(CurMip).Label = OldLabel & "_MIP" & CurMip
   BitMap2D_Resample TheColorMips(TheIndex).MipSequance(CurMip), NewW, NewH, TheTexelsFilter
   NewW = (NewW + StpW): NewH = (NewH + StpH)
  Next CurMip

 End If

End Function
Function TEX_Reflection_Add() As Long

On Error GoTo Eror

 Dim CurIndex As Long, TexIndex As Long

 For CurIndex = 0 To MaxMeshs
  If (TheReflectionUsed(CurIndex) = False) Then
   TexIndex = CurIndex: Exit For
  End If
 Next CurIndex

 TheReflectionTextures(TexIndex).Label = "ReflectionTexture_" & (TheMeshsCount + 1)
 TEX_Reflection_Add = TexIndex: TheReflectionUsed(TexIndex) = True
 ReDim Preserve TheReflectionIndexs(UBound(TheReflectionIndexs) + 1)
 TheReflectionIndexs(UBound(TheReflectionIndexs)) = TexIndex
 ExtractionSortLong TheReflectionIndexs(), True

Eror:

 If (Err.Number = 9) Then
  ReDim Preserve TheReflectionIndexs(0)
  TheReflectionIndexs(0) = TexIndex
 End If

End Function
Function TEX_Reflection_Clear()

On Error Resume Next

 Dim CurPos As Long

 For CurPos = 0 To MaxMeshs
  BitMap2D_Delete TheReflectionTextures(CurPos)
  TheReflectionUsed(CurPos) = False
 Next CurPos

 Erase TheReflectionIndexs()

End Function
Function TEX_Reflection_Remove(TheIndex As Long)

 If ((TheIndex < 0) Or (TheIndex > MaxMeshs)) Then Exit Function
 If (TheReflectionUsed(TheIndex) = False) Then Exit Function

 BitMap2D_Delete TheReflectionTextures(TheIndex)
 TheReflectionUsed(TheIndex) = False

 Dim CurPos As Long
 For CurPos = 0 To UBound(TheReflectionIndexs)
  If (TheReflectionIndexs(CurPos) = TheIndex) Then
   TheReflectionIndexs(CurPos) = (MaxMeshs + 1): Exit For
  End If
 Next CurPos

 If (UBound(TheReflectionIndexs) = 0) Then
  Erase TheReflectionIndexs()
 Else
  ExtractionSortLong TheReflectionIndexs(), True
  ReDim Preserve TheReflectionIndexs(UBound(TheReflectionIndexs) - 1)
 End If

End Function
Function TEX_Reflection_Set(TheIndex As Long, TheBitmap As BitMap2D)

 If ((TheIndex < 0) Or (TheIndex > MaxMeshs)) Then Exit Function
 If (TheReflectionUsed(TheIndex) = False) Then Exit Function
 If (BitMap2D_IsValid(TheBitmap) = False) Then Exit Function
 If (TheBitmap.BitsDepth <> 8) Then Exit Function

 Dim OldLabel As String
 OldLabel = TheReflectionTextures(TheIndex).Label
 TheReflectionTextures(TheIndex) = TheBitmap
 TheReflectionTextures(TheIndex).Label = OldLabel

End Function
Function TEX_RefractionN_Add() As Long

On Error GoTo Eror

 Dim CurIndex As Long, TexIndex As Long

 For CurIndex = 0 To MaxMeshs
  If (TheRefractionNUsed(CurIndex) = False) Then
   TexIndex = CurIndex: Exit For
  End If
 Next CurIndex

 TheRefractionNTextures(TexIndex).Label = "RefractionNTexture_" & (TheMeshsCount + 1)
 TEX_RefractionN_Add = TexIndex: TheRefractionNUsed(TexIndex) = True
 ReDim Preserve TheRefractionNIndexs(UBound(TheRefractionNIndexs) + 1)
 TheRefractionNIndexs(UBound(TheRefractionNIndexs)) = TexIndex
 ExtractionSortLong TheRefractionNIndexs(), True

Eror:

 If (Err.Number = 9) Then
  ReDim Preserve TheRefractionNIndexs(0)
  TheRefractionNIndexs(0) = TexIndex
 End If

End Function
Function TEX_RefractionN_Clear()

On Error Resume Next

 Dim CurPos As Long

 For CurPos = 0 To MaxMeshs
  BitMap2D_Delete TheRefractionNTextures(CurPos)
  TheRefractionNUsed(CurPos) = False
 Next CurPos

 Erase TheRefractionNIndexs()

End Function
Function TEX_RefractionN_Remove(TheIndex As Long)

 If ((TheIndex < 0) Or (TheIndex > MaxMeshs)) Then Exit Function
 If (TheRefractionNUsed(TheIndex) = False) Then Exit Function

 BitMap2D_Delete TheRefractionNTextures(TheIndex)
 TheRefractionNUsed(TheIndex) = False

 Dim CurPos As Long
 For CurPos = 0 To UBound(TheRefractionNIndexs)
  If (TheRefractionNIndexs(CurPos) = TheIndex) Then
   TheRefractionNIndexs(CurPos) = (MaxMeshs + 1): Exit For
  End If
 Next CurPos

 If (UBound(TheRefractionNIndexs) = 0) Then
  Erase TheRefractionNIndexs()
 Else
  ExtractionSortLong TheRefractionNIndexs(), True
  ReDim Preserve TheRefractionNIndexs(UBound(TheRefractionNIndexs) - 1)
 End If

End Function
Function TEX_RefractionN_Set(TheIndex As Long, TheBitmap As BitMap2D)

 If ((TheIndex < 0) Or (TheIndex > MaxMeshs)) Then Exit Function
 If (TheRefractionNUsed(TheIndex) = False) Then Exit Function
 If (BitMap2D_IsValid(TheBitmap) = False) Then Exit Function
 If (TheBitmap.BitsDepth <> 8) Then Exit Function

 Dim OldLabel As String
 OldLabel = TheRefractionNTextures(TheIndex).Label
 TheRefractionNTextures(TheIndex) = TheBitmap
 TheRefractionNTextures(TheIndex).Label = OldLabel

End Function
Function Textures_Add(TheMaterial As Material)

 With TheMaterial
 .ColorTextureID = TEX_Color_Add
 .AlphaTextureID = TEX_Alpha_Add
 .ReflectionTextureID = TEX_Reflection_Add
 .RefractionTextureID = TEX_Refraction_Add
 .RefractionNTextureID = TEX_RefractionN_Add
 End With

End Function
Function Textures_Clear()

 TEX_Color_Clear
 TEX_Alpha_Clear
 TEX_Reflection_Clear
 TEX_Refraction_Clear
 TEX_RefractionN_Clear

End Function
Function TEX_Refraction_Add() As Long

On Error GoTo Eror

 Dim CurIndex As Long, TexIndex As Long

 For CurIndex = 0 To MaxMeshs
  If (TheRefractionUsed(CurIndex) = False) Then
   TexIndex = CurIndex: Exit For
  End If
 Next CurIndex

 TheRefractionTextures(TexIndex).Label = "RefractionTexture_" & (TheMeshsCount + 1)
 TEX_Refraction_Add = TexIndex: TheRefractionUsed(TexIndex) = True
 ReDim Preserve TheRefractionIndexs(UBound(TheRefractionIndexs) + 1)
 TheRefractionIndexs(UBound(TheRefractionIndexs)) = TexIndex
 ExtractionSortLong TheRefractionIndexs(), True

Eror:

 If (Err.Number = 9) Then
  ReDim Preserve TheRefractionIndexs(0)
  TheRefractionIndexs(0) = TexIndex
 End If

End Function
Function TEX_Refraction_Clear()

On Error Resume Next

 Dim CurPos As Long

 For CurPos = 0 To MaxMeshs
  BitMap2D_Delete TheRefractionTextures(CurPos)
  TheRefractionUsed(CurPos) = False
 Next CurPos

 Erase TheRefractionIndexs()

End Function
Function TEX_Refraction_Remove(TheIndex As Long)

 If ((TheIndex < 0) Or (TheIndex > MaxMeshs)) Then Exit Function
 If (TheRefractionUsed(TheIndex) = False) Then Exit Function

 BitMap2D_Delete TheRefractionTextures(TheIndex)
 TheRefractionUsed(TheIndex) = False

 Dim CurPos As Long
 For CurPos = 0 To UBound(TheRefractionIndexs)
  If (TheRefractionIndexs(CurPos) = TheIndex) Then
   TheRefractionIndexs(CurPos) = (MaxMeshs + 1): Exit For
  End If
 Next CurPos

 If (UBound(TheRefractionIndexs) = 0) Then
  Erase TheRefractionIndexs()
 Else
  ExtractionSortLong TheRefractionIndexs(), True
  ReDim Preserve TheRefractionIndexs(UBound(TheRefractionIndexs) - 1)
 End If

End Function
Function TEX_Refraction_Set(TheIndex As Long, TheBitmap As BitMap2D)

 If ((TheIndex < 0) Or (TheIndex > MaxMeshs)) Then Exit Function
 If (TheRefractionUsed(TheIndex) = False) Then Exit Function
 If (BitMap2D_IsValid(TheBitmap) = False) Then Exit Function
 If (TheBitmap.BitsDepth <> 8) Then Exit Function

 Dim OldLabel As String
 OldLabel = TheRefractionTextures(TheIndex).Label
 TheRefractionTextures(TheIndex) = TheBitmap
 TheRefractionTextures(TheIndex).Label = OldLabel

End Function
Function Camera3D_Add() As Long

 If ((TheCamerasCount + 1) <= MaxCameras) Then
  If ((TheCamerasCount + 1) <> -1) Then
   Camera3D_Add = (TheCamerasCount + 1)
   TheCameras(TheCamerasCount + 1) = Camera3D_Default
   TheCamerasCount = (TheCamerasCount + 1)
  End If
 End If

End Function
Function Camera3D_Null() As Camera3D

End Function
Function Camera3D_Remove(CameraIndex As Long)

 If (TheCamerasCount = 0) Then Camera3D_Clear: Exit Function

 Dim CurCamera&
 For CurCamera = CameraIndex To (TheCamerasCount - 1)
  TheCameras(CurCamera) = TheCameras(CurCamera + 1)
 Next CurCamera
 TheCameras(TheCamerasCount) = Camera3D_Null

 TheCamerasCount = (TheCamerasCount - 1)

End Function
Function Camera3D_Copy(CameraIndex As Long) As Long

 If (TheCamerasCount = -1) Then Exit Function

 Camera3D_Copy = Camera3D_Add
 TheCameras(Camera3D_Copy) = TheCameras(CameraIndex)
 TheCameras(Camera3D_Copy).Label = "CopyOf_" & TheCameras(CameraIndex).Label

End Function
Function Camera3D_Find(TheLabel As String) As Long

 Dim CurCamera&

 For CurCamera = 0 To TheCamerasCount
  If (TheCameras(CurCamera).Label = TheLabel) Then Camera3D_Find = CurCamera: Exit For
 Next CurCamera

End Function
Function Camera3D_Clear()

 Dim CurCamera&

 For CurCamera = 0 To TheCamerasCount
  TheCameras(CurCamera) = Camera3D_Null
 Next CurCamera

 TheCamerasCount = -1

End Function
Function Material_Default() As Material

 Randomize
 With Material_Default
  .Label = "Material_" & (TheMeshsCount + 1)
  .Color = ColorRandomFrom(127)
  .Reflection = 0
  .Refraction = 0
  .RefractionN = 1
  .SpecularPowerK = 1
  .SpecularPowerN = 5
 End With

End Function
Function Material_Null() As Material

End Function
Function Mesh3D_AllocateVertices(Count As Long) As Long

 Dim CurMesh1&, CurMesh2&, GapStart&, GapEnd&

 If (Count <> 0) Then
  For CurMesh1 = 0 To TheMeshsCount
   GapStart = 0: GapEnd = (TheMeshs(CurMesh1).Vertices.Start - 1)
   For CurMesh2 = 0 To TheMeshsCount
    If ((TheMeshs(CurMesh2).Vertices.Start < TheMeshs(CurMesh1).Vertices.Start) And _
        (GetAddressLast(TheMeshs(CurMesh2).Vertices) > GapStart)) Then GapStart = GetAddressLast(TheMeshs(CurMesh2).Vertices)
   Next CurMesh2
   GapStart = (GapStart + 1)
   If ((GapEnd - (GapStart + 1)) >= Count) Then: Mesh3D_AllocateVertices = GapStart: Exit Function
  Next CurMesh1
 End If

 GapStart = 0
 For CurMesh1 = 0 To TheMeshsCount
  If (GetAddressLast(TheMeshs(CurMesh1).Vertices) > GapStart) Then
   GapStart = GetAddressLast(TheMeshs(CurMesh1).Vertices)
  End If
 Next CurMesh1

 GapStart = (GapStart + 1)
 If (((GapStart - 1) + Count) <= MaxVertices) Then Mesh3D_AllocateVertices = GapStart

End Function
Function Camera3D_Default() As Camera3D

 With Camera3D_Default
  .Label = "Camera_" & (TheCamerasCount + 1)
  .Position = VectorInput(-100, -100, 100)
  .Direction = VectorNull
  .RollAngle = ApproachVal 'Just for ignoring some errors because the zero
  .FOVAngle = (Deg * 90)
  .ClearDistance = 450
  .Dispersion = 0
  .BackFaceCulling = False
  .MakeMatrix = True
  .ViewMatrix = MatrixIdentity
 End With

End Function
Function Face3D_Null() As Face3D

End Function
Function Mesh3D_Default() As Mesh3D

 With Mesh3D_Default
  .Label = "Mesh_" & (TheMeshsCount + 1)
  .Position = VectorNull
  .Scales = VectorInput(1, 1, 1)
  .Angles = VectorNull
  .Visible = True
  .MakeMatrix = True
  .WorldMatrix = MatrixIdentity
 End With

End Function
Function Mesh3D_Null() As Mesh3D

End Function
Function Face3D_Remove(FaceIndex As Long)

 Dim CurMesh&, CurFace&

 For CurMesh = 0 To TheMeshsCount
  If (Face3D_IsInMesh(FaceIndex, CurMesh) = True) Then
   For CurFace = FaceIndex To (GetAddressLast(TheMeshs(CurMesh).Faces) - 1)
    TheFaces(CurFace) = TheFaces(CurFace + 1)
   Next CurFace
   TheFaces(GetAddressLast(TheMeshs(CurMesh).Faces)) = Face3D_Null
   TheMeshs(CurMesh).Faces.Length = (TheMeshs(CurMesh).Faces.Length - 1)
  End If
 Next CurMesh

End Function
Function Face3D_IsInMesh(FaceIndex As Long, MeshIndex As Long) As Boolean

 If ((FaceIndex >= TheMeshs(MeshIndex).Faces.Start) And _
     (FaceIndex <= GetAddressLast(TheMeshs(MeshIndex).Faces))) Then
  Face3D_IsInMesh = True
 End If

End Function
Function Face3D_FindByVertex(VertexIndex As Long) As Long

 'Find the first one

 Dim CurMesh&, CurFace&

 For CurMesh = 0 To TheMeshsCount
  For CurFace = TheMeshs(CurMesh).Faces.Start To GetAddressLast(TheMeshs(CurMesh).Faces)
   If (Vertex3D_IsInFace(VertexIndex, CurFace) = True) Then Face3D_FindByVertex = CurFace: Exit For
  Next CurFace
 Next CurMesh

End Function
Function Mesh3D_Add(VerticesCount As Long, FacesCount As Long) As Long

 Mesh3D_Add = -1

 If ((TheMeshsCount + 1) <= MaxMeshs) Then
  If ((TheMeshsCount + 1) <> -1) Then
   Dim VerticesStart&, FacesStart&
   VerticesStart = Mesh3D_AllocateVertices(VerticesCount)
   FacesStart = Mesh3D_AllocateFaces(FacesCount)
   If ((VerticesStart <> 0) And (FacesStart <> 0)) Then
    Mesh3D_Add = (TheMeshsCount + 1)
    TheMeshs(Mesh3D_Add) = Mesh3D_Default
    TheMeshs(Mesh3D_Add).Vertices.Start = VerticesStart
    TheMeshs(Mesh3D_Add).Vertices.Length = VerticesCount
    TheMeshs(Mesh3D_Add).Faces.Start = FacesStart
    TheMeshs(Mesh3D_Add).Faces.Length = FacesCount
    TheMaterials(Mesh3D_Add) = Material_Default
    Textures_Add TheMaterials(Mesh3D_Add)
    TheMeshsCount = (TheMeshsCount + 1)
   End If
  End If
 End If

 If (Mesh3D_Add <> -1) Then
  Dim CurFace&
  For CurFace = TheMeshs(Mesh3D_Add).Faces.Start To GetAddressLast(TheMeshs(Mesh3D_Add).Faces)
   TheFaces(CurFace).Visible = True
  Next CurFace
 Else
  MsgBox "The allocated memory is not enough, can't continue.", vbCritical, "Memory"
 End If

End Function
Function Mesh3D_Remove(MeshIndex As Long)

 If (TheMeshsCount = 0) Then Mesh3D_Clear: Exit Function

 TEX_Alpha_Remove TheMaterials(MeshIndex).AlphaTextureID
 TEX_Color_Remove TheMaterials(MeshIndex).ColorTextureID
 TEX_Reflection_Remove TheMaterials(MeshIndex).ReflectionTextureID
 TEX_Refraction_Remove TheMaterials(MeshIndex).RefractionTextureID
 TEX_RefractionN_Remove TheMaterials(MeshIndex).RefractionNTextureID

 Dim CurMesh&
 For CurMesh = MeshIndex To (TheMeshsCount - 1)
  TheMeshs(CurMesh) = TheMeshs(CurMesh + 1)
  TheMaterials(CurMesh) = TheMaterials(CurMesh + 1)
 Next CurMesh
 TheMeshs(TheMeshsCount) = Mesh3D_Null
 TheMaterials(TheMeshsCount) = Material_Null

 TheMeshsCount = (TheMeshsCount - 1)

End Function
Function Mesh3D_Clear()

 Dim CurMesh&

 For CurMesh = 0 To TheMeshsCount
  TheMeshs(CurMesh) = Mesh3D_Null
  TheMaterials(CurMesh) = Material_Null
 Next CurMesh
 Textures_Clear

 TheMeshsCount = -1

End Function
Function Mesh3D_Copy(MeshIndex As Long) As Long

 If (TheMeshsCount = -1) Then Exit Function

 Dim CurVertex&, CurFace&, VertexOffset&, FaceOffset&
 Dim TmpAlphaMap As BitMap2D, TmpColorMap As BitMap2D, TmpReflectionMap As BitMap2D

 Mesh3D_Copy = Mesh3D_Add(TheMeshs(MeshIndex).Vertices.Length, TheMeshs(MeshIndex).Faces.Length)
 If (Mesh3D_Copy = -1) Then Exit Function

 'Copy vertices:
 VertexOffset = (TheMeshs(Mesh3D_Copy).Vertices.Start - TheMeshs(MeshIndex).Vertices.Start)
 For CurVertex = TheMeshs(Mesh3D_Copy).Vertices.Start To GetAddressLast(TheMeshs(Mesh3D_Copy).Vertices)
  TheVertices(CurVertex) = TheVertices(CurVertex - VertexOffset)
 Next CurVertex

 'Copy faces:
 FaceOffset = (TheMeshs(Mesh3D_Copy).Faces.Start - TheMeshs(MeshIndex).Faces.Start)
 For CurFace = TheMeshs(Mesh3D_Copy).Faces.Start To GetAddressLast(TheMeshs(Mesh3D_Copy).Faces)
  TheFaces(CurFace) = TheFaces(CurFace - FaceOffset)
  TheFaces(CurFace).A = (TheFaces(CurFace).A + VertexOffset)
  TheFaces(CurFace).B = (TheFaces(CurFace).B + VertexOffset)
  TheFaces(CurFace).C = (TheFaces(CurFace).C + VertexOffset)
 Next CurFace

 'Mesh's properties:
 TheMeshs(Mesh3D_Copy).Label = "CopyOf_" & TheMeshs(MeshIndex).Label
 TheMeshs(Mesh3D_Copy).Position = TheMeshs(MeshIndex).Position
 TheMeshs(Mesh3D_Copy).Scales = TheMeshs(MeshIndex).Scales
 TheMeshs(Mesh3D_Copy).Angles = TheMeshs(MeshIndex).Angles
 TheMeshs(Mesh3D_Copy).Vertices.Start = (TheMeshs(MeshIndex).Vertices.Start + VertexOffset)
 TheMeshs(Mesh3D_Copy).Faces.Start = (TheMeshs(MeshIndex).Faces.Start + FaceOffset)
 TheMeshs(Mesh3D_Copy).MakeMatrix = TheMeshs(MeshIndex).MakeMatrix
 If (TheMeshs(Mesh3D_Copy).MakeMatrix = False) Then
  TheMeshs(Mesh3D_Copy).WorldMatrix = TheMeshs(MeshIndex).WorldMatrix
 End If
 TheMeshs(Mesh3D_Copy).Visible = TheMeshs(MeshIndex).Visible

 'Material:
 TheMaterials(Mesh3D_Copy).Label = "CopyOf_" & TheMaterials(MeshIndex).Label
 TheMaterials(Mesh3D_Copy).Color = TheMaterials(MeshIndex).Color
 TheMaterials(Mesh3D_Copy).Reflection = TheMaterials(MeshIndex).Reflection
 TheMaterials(Mesh3D_Copy).Refraction = TheMaterials(MeshIndex).Refraction
 TheMaterials(Mesh3D_Copy).RefractionN = TheMaterials(MeshIndex).RefractionN
 TheMaterials(Mesh3D_Copy).SpecularPowerK = TheMaterials(MeshIndex).SpecularPowerK
 TheMaterials(Mesh3D_Copy).SpecularPowerN = TheMaterials(MeshIndex).SpecularPowerN

 'Copy textures
 If (TheMaterials(MeshIndex).UseAlphaTexture = True) Then
  TEX_Alpha_Set Mesh3D_Copy, TheAlphaTextures(MeshIndex)
  TheMaterials(Mesh3D_Copy).UseAlphaTexture = True
 End If
 If (TheMaterials(MeshIndex).UseColorTexture = True) Then
  TEX_Color_Set Mesh3D_Copy, TheColorTextures(MeshIndex)
  TheMaterials(Mesh3D_Copy).UseColorTexture = True
 End If
 If (TheMaterials(MeshIndex).UseReflectionTexture = True) Then
  TEX_Reflection_Set Mesh3D_Copy, TheReflectionTextures(MeshIndex)
  TheMaterials(Mesh3D_Copy).UseReflectionTexture = True
 End If
 If (TheMaterials(MeshIndex).UseRefractionTexture = True) Then
  TEX_Refraction_Set Mesh3D_Copy, TheRefractionTextures(MeshIndex)
  TheMaterials(Mesh3D_Copy).UseRefractionTexture = True
 End If
 If (TheMaterials(MeshIndex).UseRefractionNTexture = True) Then
  TEX_RefractionN_Set Mesh3D_Copy, TheRefractionNTextures(MeshIndex)
  TheMaterials(Mesh3D_Copy).UseRefractionNTexture = True
 End If

End Function
Function Mesh3D_Attach(MeshIndexA As Long, MeshIndexB As Long, KeepTransformedVertices As Boolean) As Long

 Dim MeshIndexC&, CurVertex&, CurFace&, VertexOffset&, FaceOffset&

 MeshIndexC = Mesh3D_Add((TheMeshs(MeshIndexA).Vertices.Length + TheMeshs(MeshIndexB).Vertices.Length), _
                         (TheMeshs(MeshIndexA).Faces.Length + TheMeshs(MeshIndexB).Faces.Length))
 If (MeshIndexC = -1) Then Exit Function

 VertexOffset = (TheMeshs(MeshIndexC).Vertices.Start - TheMeshs(MeshIndexA).Vertices.Start)
 If (KeepTransformedVertices = True) Then
  For CurVertex = TheMeshs(MeshIndexC).Vertices.Start To GetAddressLast(TheMeshs(MeshIndexC).Vertices)
   'TheVertices(CurVertex).Position = TheVertices(CurVertex - VertexOffset).TmpPos
   TheVertices(CurVertex).Position = MatrixMultiplyVector(TheVertices(CurVertex - VertexOffset).Position, TheMeshs(MeshIndexA).WorldMatrix)
  Next CurVertex
 Else
  For CurVertex = TheMeshs(MeshIndexC).Vertices.Start To GetAddressLast(TheMeshs(MeshIndexC).Vertices)
   TheVertices(CurVertex) = TheVertices(CurVertex - VertexOffset)
  Next CurVertex
 End If
 FaceOffset = (TheMeshs(MeshIndexC).Faces.Start - TheMeshs(MeshIndexA).Faces.Start)
 For CurFace = TheMeshs(MeshIndexC).Faces.Start To GetAddressLast(TheMeshs(MeshIndexC).Faces)
  TheFaces(CurFace) = TheFaces(CurFace - FaceOffset)
  TheFaces(CurFace).A = (TheFaces(CurFace).A + VertexOffset)
  TheFaces(CurFace).B = (TheFaces(CurFace).B + VertexOffset)
  TheFaces(CurFace).C = (TheFaces(CurFace).C + VertexOffset)
 Next CurFace

 VertexOffset = ((TheMeshs(MeshIndexC).Vertices.Start + TheMeshs(MeshIndexA).Vertices.Length) - TheMeshs(MeshIndexB).Vertices.Start)
 If (KeepTransformedVertices = True) Then
  For CurVertex = (TheMeshs(MeshIndexC).Vertices.Start + TheMeshs(MeshIndexA).Vertices.Length) To GetAddressLast(TheMeshs(MeshIndexC).Vertices)
   'TheVertices(CurVertex).Position = TheVertices(CurVertex - VertexOffset).TmpPos
   TheVertices(CurVertex).Position = MatrixMultiplyVector(TheVertices(CurVertex - VertexOffset).Position, TheMeshs(MeshIndexB).WorldMatrix)
  Next CurVertex
 Else
  For CurVertex = (TheMeshs(MeshIndexC).Vertices.Start + TheMeshs(MeshIndexA).Vertices.Length) To GetAddressLast(TheMeshs(MeshIndexC).Vertices)
   TheVertices(CurVertex) = TheVertices(CurVertex - VertexOffset)
  Next CurVertex
 End If
 FaceOffset = ((TheMeshs(MeshIndexC).Faces.Start + TheMeshs(MeshIndexA).Faces.Length) - TheMeshs(MeshIndexB).Faces.Start)
 For CurFace = (TheMeshs(MeshIndexC).Faces.Start + TheMeshs(MeshIndexA).Faces.Length) To GetAddressLast(TheMeshs(MeshIndexC).Faces)
  TheFaces(CurFace) = TheFaces(CurFace - FaceOffset)
  TheFaces(CurFace).A = (TheFaces(CurFace).A + VertexOffset)
  TheFaces(CurFace).B = (TheFaces(CurFace).B + VertexOffset)
  TheFaces(CurFace).C = (TheFaces(CurFace).C + VertexOffset)
 Next CurFace

 If (MeshIndexA = MeshIndexB) Then
  Mesh3D_Remove MeshIndexA
 ElseIf (MeshIndexA < MeshIndexB) Then
  Mesh3D_Remove MeshIndexA: Mesh3D_Remove (MeshIndexB - 1)
 ElseIf (MeshIndexA > MeshIndexB) Then
  Mesh3D_Remove MeshIndexB: Mesh3D_Remove (MeshIndexA - 1)
 End If

 Mesh3D_Attach = (MeshIndexC - 2)

End Function
Function Mesh3D_Trim(MeshIndex As Long)

 Dim A&, B&, Reference As Boolean

 'Remove any doubled vertex:
 For A = GetAddressLast(TheMeshs(MeshIndex).Vertices) To TheMeshs(MeshIndex).Vertices.Start Step -1
  For B = TheMeshs(MeshIndex).Vertices.Start To GetAddressLast(TheMeshs(MeshIndex).Vertices)
   If (B <> A) Then
    If (VectorDistance(TheVertices(B).Position, TheVertices(A).Position) < ApproachVal) Then Vertex3D_Remove B, A: Exit For
   End If
  Next B
 Next A

 'Remove any vertex not used by the faces:
 For A = GetAddressLast(TheMeshs(MeshIndex).Vertices) To TheMeshs(MeshIndex).Vertices.Start Step -1
  Reference = False
  For B = TheMeshs(MeshIndex).Faces.Start To GetAddressLast(TheMeshs(MeshIndex).Faces)
   If ((TheFaces(B).A = A) Or (TheFaces(B).B = A) Or (TheFaces(B).C = A)) Then Reference = True
  Next B
  If (Reference = False) Then Vertex3D_Remove A, 0
 Next A

 'Remove the face that one or more or theire vertices appart to another face:
 For A = GetAddressLast(TheMeshs(MeshIndex).Faces) To TheMeshs(MeshIndex).Faces.Start Step -1
  If ((Vertex3D_IsInMesh(TheFaces(A).A, MeshIndex) = False) Or _
      (Vertex3D_IsInMesh(TheFaces(A).B, MeshIndex) = False) Or _
      (Vertex3D_IsInMesh(TheFaces(A).C, MeshIndex) = False)) Then Face3D_Remove A
 Next A

 'Remove wrong faces : A=B, B=C, A=C
 For A = GetAddressLast(TheMeshs(MeshIndex).Faces) To TheMeshs(MeshIndex).Faces.Start Step -1
  If ((TheFaces(A).A = TheFaces(A).B) Or _
      (TheFaces(A).B = TheFaces(A).C) Or _
      (TheFaces(A).C = TheFaces(A).A)) Then Face3D_Remove A
 Next A

 'Remove a face that one or more of theire corners indicate's vertex 0:
 For A = GetAddressLast(TheMeshs(MeshIndex).Faces) To TheMeshs(MeshIndex).Faces.Start Step -1
  If ((TheFaces(A).A = 0) Or (TheFaces(A).B = 0) Or (TheFaces(A).C = 0)) Then Face3D_Remove A
 Next A

End Function
Function Mesh3D_GetCenter(MeshIndex As Long, OriginalVectors As Boolean) As Vector3D

 Mesh3D_GetCenter = VectorInterpolate(Mesh3D_MaxDims(MeshIndex, OriginalVectors), Mesh3D_MinDims(MeshIndex, OriginalVectors), 0.5)

End Function
Function Mesh3D_MaxDims(MeshIndex As Long, OriginalVectors As Boolean) As Vector3D

 Dim CurVertex&

 Mesh3D_MaxDims = VectorInput(-MaxSingleFloat, -MaxSingleFloat, -MaxSingleFloat)

 If (OriginalVectors = True) Then
  For CurVertex = TheMeshs(MeshIndex).Vertices.Start To GetAddressLast(TheMeshs(MeshIndex).Vertices)
   If (TheVertices(CurVertex).Position.X > Mesh3D_MaxDims.X) Then Mesh3D_MaxDims.X = TheVertices(CurVertex).Position.X
   If (TheVertices(CurVertex).Position.Y > Mesh3D_MaxDims.Y) Then Mesh3D_MaxDims.Y = TheVertices(CurVertex).Position.Y
   If (TheVertices(CurVertex).Position.Z > Mesh3D_MaxDims.Z) Then Mesh3D_MaxDims.Z = TheVertices(CurVertex).Position.Z
  Next CurVertex
 Else
  For CurVertex = TheMeshs(MeshIndex).Vertices.Start To GetAddressLast(TheMeshs(MeshIndex).Vertices)
   If (TheVertices(CurVertex).TmpPos.X > Mesh3D_MaxDims.X) Then Mesh3D_MaxDims.X = TheVertices(CurVertex).TmpPos.X
   If (TheVertices(CurVertex).TmpPos.Y > Mesh3D_MaxDims.Y) Then Mesh3D_MaxDims.Y = TheVertices(CurVertex).TmpPos.Y
   If (TheVertices(CurVertex).TmpPos.Z > Mesh3D_MaxDims.Z) Then Mesh3D_MaxDims.Z = TheVertices(CurVertex).TmpPos.Z
  Next CurVertex
 End If

End Function
Function Mesh3D_ComputeBoundingBox(MeshIndex As Long, OriginalVectors As Boolean) As Ray3D

 Mesh3D_ComputeBoundingBox.Position = Mesh3D_MinDims(MeshIndex, OriginalVectors)
 Mesh3D_ComputeBoundingBox.Direction = Mesh3D_MaxDims(MeshIndex, OriginalVectors)

End Function
Function Mesh3D_MinDims(MeshIndex As Long, OriginalVectors As Boolean) As Vector3D

 Dim CurVertex&

 Mesh3D_MinDims = VectorInput(MaxSingleFloat, MaxSingleFloat, MaxSingleFloat)

 If (OriginalVectors = True) Then
  For CurVertex = TheMeshs(MeshIndex).Vertices.Start To GetAddressLast(TheMeshs(MeshIndex).Vertices)
   If (TheVertices(CurVertex).Position.X < Mesh3D_MinDims.X) Then Mesh3D_MinDims.X = TheVertices(CurVertex).Position.X
   If (TheVertices(CurVertex).Position.Y < Mesh3D_MinDims.Y) Then Mesh3D_MinDims.Y = TheVertices(CurVertex).Position.Y
   If (TheVertices(CurVertex).Position.Z < Mesh3D_MinDims.Z) Then Mesh3D_MinDims.Z = TheVertices(CurVertex).Position.Z
  Next CurVertex
 Else
  For CurVertex = TheMeshs(MeshIndex).Vertices.Start To GetAddressLast(TheMeshs(MeshIndex).Vertices)
   If (TheVertices(CurVertex).TmpPos.X < Mesh3D_MinDims.X) Then Mesh3D_MinDims.X = TheVertices(CurVertex).TmpPos.X
   If (TheVertices(CurVertex).TmpPos.Y < Mesh3D_MinDims.Y) Then Mesh3D_MinDims.Y = TheVertices(CurVertex).TmpPos.Y
   If (TheVertices(CurVertex).TmpPos.Z < Mesh3D_MinDims.Z) Then Mesh3D_MinDims.Z = TheVertices(CurVertex).TmpPos.Z
  Next CurVertex
 End If

End Function
Function Mesh3D_FindByFace(FaceIndex As Long) As Long

 Dim CurMesh&

 For CurMesh = 0 To TheMeshsCount
  If (Face3D_IsInMesh(FaceIndex, CurMesh) = True) Then Mesh3D_FindByFace = CurMesh: Exit For
 Next CurMesh

End Function
Function Mesh3D_FindByVertex(VertexIndex As Long) As Long

 Dim CurMesh&

 For CurMesh = 0 To TheMeshsCount
  If (Vertex3D_IsInMesh(VertexIndex, CurMesh) = True) Then Mesh3D_FindByVertex = CurMesh: Exit For
 Next CurMesh

End Function
Function Mesh3D_FindByLabel(TheLabel As String) As Long

 Dim CurMesh&

 For CurMesh = 0 To TheMeshsCount
  If (TheMeshs(CurMesh).Label = TheLabel) Then Mesh3D_FindByLabel = CurMesh: Exit For
 Next CurMesh

End Function
Function Vertex3D_IsInMesh(VertexIndex As Long, MeshIndex As Long) As Boolean

 If ((VertexIndex >= TheMeshs(MeshIndex).Vertices.Start) And _
     (VertexIndex <= GetAddressLast(TheMeshs(MeshIndex).Vertices))) Then Vertex3D_IsInMesh = True

End Function
Function Vertex3D_IsInFace(VertexIndex As Long, FaceIndex As Long) As Boolean

 If ((VertexIndex = TheFaces(FaceIndex).A) Or _
     (VertexIndex = TheFaces(FaceIndex).B) Or _
     (VertexIndex = TheFaces(FaceIndex).C)) Then Vertex3D_IsInFace = True

End Function
Function Vertex3D_Remove(VertexIndex As Long, Substitute As Long)

 Dim CurMesh&, CurVertex&, CurFace&

 For CurMesh = 0 To TheMeshsCount
  If (Vertex3D_IsInMesh(VertexIndex, CurMesh) = True) Then
   For CurVertex = VertexIndex To (GetAddressLast(TheMeshs(CurMesh).Vertices) - 1)
    TheVertices(CurVertex) = TheVertices(CurVertex + 1)
   Next CurVertex
   TheVertices(GetAddressLast(TheMeshs(CurMesh).Vertices)) = Vertex3D_Null
   TheMeshs(CurMesh).Vertices.Length = (TheMeshs(CurMesh).Vertices.Length - 1)
   For CurFace = GetAddressLast(TheMeshs(CurMesh).Faces) To TheMeshs(CurMesh).Faces.Start Step -1
    If (Vertex3D_IsInFace(VertexIndex, CurFace) = True) Then
     If (Substitute = 0) Then
      Face3D_Remove CurFace
     Else
      If (TheFaces(CurFace).A = VertexIndex) Then TheFaces(CurFace).A = Substitute
      If (TheFaces(CurFace).B = VertexIndex) Then TheFaces(CurFace).B = Substitute
      If (TheFaces(CurFace).C = VertexIndex) Then TheFaces(CurFace).C = Substitute
     End If
    End If
    If (TheFaces(CurFace).A > VertexIndex) Then TheFaces(CurFace).A = (TheFaces(CurFace).A - 1)
    If (TheFaces(CurFace).B > VertexIndex) Then TheFaces(CurFace).B = (TheFaces(CurFace).B - 1)
    If (TheFaces(CurFace).C > VertexIndex) Then TheFaces(CurFace).C = (TheFaces(CurFace).C - 1)
   Next CurFace
  End If
 Next CurMesh

End Function
Function Mesh3D_AllocateFaces(Count As Long) As Long

 Dim CurMesh1&, CurMesh2&, GapStart&, GapEnd&

 If (Count <> 0) Then
  For CurMesh1 = 0 To TheMeshsCount
   GapStart = 0: GapEnd = (TheMeshs(CurMesh1).Faces.Start - 1)
   For CurMesh2 = 0 To TheMeshsCount
    If ((TheMeshs(CurMesh2).Faces.Start < TheMeshs(CurMesh1).Faces.Start) And _
        (GetAddressLast(TheMeshs(CurMesh2).Faces) > GapStart)) Then GapStart = GetAddressLast(TheMeshs(CurMesh2).Faces)
   Next CurMesh2
   GapStart = (GapStart + 1)
   If ((GapEnd - (GapStart + 1)) >= Count) Then Mesh3D_AllocateFaces = GapStart: Exit Function
  Next CurMesh1
 End If

 GapStart = 0
 For CurMesh1 = 0 To TheMeshsCount
  If (GetAddressLast(TheMeshs(CurMesh1).Faces) > GapStart) Then
   GapStart = GetAddressLast(TheMeshs(CurMesh1).Faces)
  End If
 Next CurMesh1

 GapStart = (GapStart + 1)
 If (((GapStart - 1) + Count) <= MaxFaces) Then Mesh3D_AllocateFaces = GapStart

End Function
Function Vertex3D_Null() As Vertex3D

End Function
Function SphereLight3D_Add() As Long

 If ((TheSphereLightsCount + 1) <= MaxLights) Then
  If ((TheSphereLightsCount + 1) <> -1) Then
   SphereLight3D_Add = (TheSphereLightsCount + 1)
   TheSphereLights(TheSphereLightsCount + 1) = SphereLight3D_Default
   TheSphereLightsCount = (TheSphereLightsCount + 1)
  End If
 End If

End Function
Function SphereLight3D_Null() As SphereLight3D

End Function
Function SphereLight3D_Remove(SphereLightIndex As Long)

 If (TheSphereLightsCount = 0) Then SphereLight3D_Clear: Exit Function

 Dim CurSphereLight&
 For CurSphereLight = SphereLightIndex To (TheSphereLightsCount - 1)
  TheSphereLights(CurSphereLight) = TheSphereLights(CurSphereLight + 1)
 Next CurSphereLight
 TheSphereLights(TheSphereLightsCount) = SphereLight3D_Null

 TheSphereLightsCount = (TheSphereLightsCount - 1)

End Function
Function SphereLight3D_Copy(SphereLightIndex As Long) As Long

 If (TheSphereLightsCount = -1) Then Exit Function

 SphereLight3D_Copy = SphereLight3D_Add
 TheSphereLights(SphereLight3D_Copy) = TheSphereLights(SphereLightIndex)
 TheSphereLights(SphereLight3D_Copy).Label = "CopyOf_" & TheSphereLights(SphereLightIndex).Label

End Function
Function SphereLight3D_Find(TheLabel As String) As Long

 Dim CurSphereLight&

 For CurSphereLight = 0 To TheSphereLightsCount
  If (TheSphereLights(CurSphereLight).Label = TheLabel) Then SphereLight3D_Find = CurSphereLight: Exit For
 Next CurSphereLight

End Function
Function SphereLight3D_Clear()

 Dim CurSphereLight&

 For CurSphereLight = 0 To TheSphereLightsCount
  TheSphereLights(CurSphereLight) = SphereLight3D_Null
 Next CurSphereLight

 TheSphereLightsCount = -1

End Function
Function SphereLight3D_Default() As SphereLight3D

 With SphereLight3D_Default
  .Label = "SphereLight_" & (TheSphereLightsCount + 1)
  .Color = ColorWhite
  .Position = VectorInput(0, 0, 0)
  .Range = 500
  .Enable = True
 End With

End Function
Function ConeLight3D_Add() As Long

 If ((TheConeLightsCount + 1) <= MaxLights) Then
  If ((TheConeLightsCount + 1) <> -1) Then
   ConeLight3D_Add = (TheConeLightsCount + 1)
   TheConeLights(TheConeLightsCount + 1) = ConeLight3D_Default
   TheConeLightsCount = (TheConeLightsCount + 1)
  End If
 End If

End Function
Function ConeLight3D_Null() As ConeLight3D

End Function
Function ConeLight3D_Remove(ConeLightIndex As Long)

 If (TheConeLightsCount = 0) Then ConeLight3D_Clear: Exit Function

 Dim CurConeLight&
 For CurConeLight = ConeLightIndex To (TheConeLightsCount - 1)
  TheConeLights(CurConeLight) = TheConeLights(CurConeLight + 1)
 Next CurConeLight
 TheConeLights(TheConeLightsCount) = ConeLight3D_Null

 TheConeLightsCount = (TheConeLightsCount - 1)

End Function
Function ConeLight3D_Copy(ConeLightIndex As Long) As Long

 If (TheConeLightsCount = -1) Then Exit Function

 ConeLight3D_Copy = ConeLight3D_Add
 TheConeLights(ConeLight3D_Copy) = TheConeLights(ConeLightIndex)
 TheConeLights(ConeLight3D_Copy).Label = "CopyOf_" & TheConeLights(ConeLightIndex).Label

End Function
Function ConeLight3D_Find(TheLabel As String) As Long

 Dim CurConeLight&

 For CurConeLight = 0 To TheConeLightsCount
  If (TheConeLights(CurConeLight).Label = TheLabel) Then ConeLight3D_Find = CurConeLight: Exit For
 Next CurConeLight

End Function
Function ConeLight3D_Clear()

 Dim CurConeLight&

 For CurConeLight = 0 To TheConeLightsCount
  TheConeLights(CurConeLight) = ConeLight3D_Null
 Next CurConeLight

 TheConeLightsCount = -1

End Function
Function ConeLight3D_Default() As ConeLight3D

 With ConeLight3D_Default
  .Label = "ConeLight_" & (TheConeLightsCount + 1)
  .Color = ColorWhite
  .Position = VectorInput(50, 50, 50)
  .Direction = VectorNull
  .Falloff = (Deg * 45)
  .Hotspot = (Deg * 30)
  .Range = 100
  .Enable = True
 End With

End Function
