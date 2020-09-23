Attribute VB_Name = "BAS_3DDREngine"

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
'###  MODULE      : BAS_3DDREngine.BAS
'###
'###  DESCRIPTION : 3D Digital Reality Engine source-code functions & algorithms.
'###
'##################################################################################
'##################################################################################

Option Explicit

'Global settings
'===============

'Stochastic pathtracer options:
Global ViewPathsPerPixel As Integer   'Paths count per a single pixel
Global SamplesPerViewPath As Integer  'Bounces count (view path)

'Shadows options:
Global EnableAreaShadows As Boolean
Global ShadowRaysCount As Integer
Global ShadowsApproxRadius As Single
Global TheShadowRays() As Vector3D

'Photon mapping options:
Global EnablePhotonMapping As Boolean
Global SamplesPerPhotonPath As Integer    'Bounces count (photon path)
Global MaximumAllocatedPhotons As Long
Global ThePhotonMap() As Photon           'Global photonmap
Global PhotonsSentCount As Long
Global PhotonsSearchRadius As Single
Global BleedingDistance As Single
Global EstimateMultiplier As Single
Global TheAmbiantLight As ColorRGB    '(Default added light if disabled photonmapping)

'Texels & textures sampling filters:
Global TheTexelsFilter As K3DE_TEXELS_FILTER_MODES
Global TheTexturesFilter As K3DE_TEXTURE_FILTER_MODES
Global CubicA!, CubicB!, CubicC!, KernelSize%, MipMapsLevel%, MipMapsMinPurcent%

'Bounding Boxes (used to accelerate the raytrace routine)
Global SceneBoundingBox As Ray3D
Global MeshsBoundingBoxes() As Ray3D

'Fog settings:
Global FogEnable As Boolean
Global FogRange As Single
Global FogMode As K3DE_FOG_MODES
Global FogColor As ColorRGB
Global FogExpFactor1 As Single, FogExpFactor2 As Single

'Background:
Global UseBackGround As Boolean, TheBackGroundColor As ColorRGB
Global TheBackGround As BitMap2D, OriginalBackGround As BitMap2D
Global MipRange As Single, TheCurrentCamera As Long

'Wireframe preview parameters:
Global Wire_PerspectiveDistorsion!, Wire_AddedDepth!, Wire_ScaleNormalTo!
Global Wire_CameraTo!, Wire_PhotonTo!, Wire_ParallalScale!
Global Wire_ParallalMoveToX!, Wire_ParallalMoveToY!
Global Wire_DefaultPerspectiveDistorsion!

'Others !
Global Started As Boolean, TheViewMatrix As Matrix4x4, TheTotalMatrix As Matrix4x4
Global UserAllocation As Boolean, OutputWidth%, OutputHeight%, StopRender As Boolean
Global PreviewMode As Boolean, DisplayMode As Integer, Previewed As Boolean
Global EstimateFromPhotonmap As Boolean, InvMipRange!
Function CalculateDirectLightContribution(TheIntersection As Intersection3D, FromPoint As Vector3D, IntersectionPoint As Vector3D, NormalVector As Vector3D, TheMatInfos As MaterialInfos) As ColorRGB

 If (Started = False) Then Exit Function

 ' Return the direct contribution of light (or direct lighting) in a
 ' given intersection point, passing by every enabled source of light in the
 ' scene. we ignore currently the indirect illumination.
 '
 '  Direct lighting factors
 '  -----------------------
 '
 ' - The shape of the light-source, the shape factor is used to
 '   determine the illumination area, but also give the desirable
 '   attenuation effect, i currenty include only the point-defined
 '   light sources, not volumetric (or participating medias) lighting
 '   so far (may be in the futur)...
 '   So i use two types:
 '   1. The sphere light sources: from the source point to all directions
 '   2. The cones lights sources: also called spotlight, just a portion from
 '                                a sphere light source, but also have an
 '                                included angular attenuation.
 '
 ' - The shadow factor option, used to render realistic shadows attenuation
 '   (or penumbra zones, also known by the name of 'Area shadows') or just
 '   the traditional sharp shadows, for this case, the engine trace only one
 '   shadow-ray (emmited ray from the light source to the intersection point,
 '   if any occulsion by objects, the area is so in shadow) for the intersection
 '   point, or the returned value is simply a boolean (shadowed/not shadowed).
 '   in the case of Area-shadows, the engine shots many shadow-rays (random samples)
 '   over an imaginary sphere around the light source, like that, we can
 '   approximate the density of the shadow just by doing the statistics,
 '   an example : we use 50 random shadow-rays, 25 are in shadow, where
 '   the other 25 are not, we get then a approximation of 0,5 (or 50%)
 '   for the density of the illumination (stochastic sampling).
 '
 ' - The diffusion factor (or the Lambert angle), is calculated just by
 '   computing the angle between the light source and the surface's normal vector.
 '
 ' - The specularity factor (or specular hightlight), the specular approach is
 '   just the lambert angle(!) between the view point and the reflected caming
 '   ray form the light source over the surface's normal vector, the specular
 '   form is just a special type of reflection (or light/object reflection,
 '   not object/object reflection), in fact, we must view the light sources as
 '   well the way we can view the reflected objects, because light sources are
 '   objects too !!. Since we have a point-from light sources (we can't just
 '   view a point...) instead, we just coloring the specular area by the light's color.
 '   The specular contribution of light depend's largely on the reflectivity of the
 '   intersected surface (a non-shiny surface can't emmit any specular reflection).

 Dim CurrentColor As ColorRGB, EndColor As ColorRGB
 Dim CurLight&, ValShape!, ValShadow!, ValLambert!, ValSpecular!
 Dim ShadowRay As Ray3D, ShadowResult As Intersection3D

 '////////////////////////////// SPHERE LIGHTS ILLUMINATION ////////////////////////////////

 If (TheSphereLightsCount = -1) Then GoTo Jump1
 For CurLight = 0 To TheSphereLightsCount
  If (TheSphereLights(CurLight).Enable = True) Then
   'Calculate the geometrical attenuation of light, a sphere in this case:
   ValShape = Shader_GetAttenuation(TheSphereLights(CurLight).TmpPos, TheSphereLights(CurLight).Range, IntersectionPoint)
   If (ValShape <> 0) Then 'Intersected light volume
    'Compute shadows :
    If (EnableAreaShadows = True) Then MakeShadowRaysAsSphere
    ValShadow = Shader_GetShadow(TheSphereLights(CurLight).TmpPos, TheIntersection, IntersectionPoint)
    If (ValShadow <> 0) Then
     'Calculate the diffusion factor (lambert angle):
     ValLambert = Shader_GetLambertAngle(FromPoint, TheSphereLights(CurLight).TmpPos, NormalVector, IntersectionPoint)
     If (ValLambert <> 0) Then
      'Calculate the output color with the given illumination factors:
      CurrentColor = ColorScale(ColorAbsorp(TheMatInfos.Color, TheSphereLights(CurLight).Color), ValLambert)
      'Calculate the light's specularity factor:
      If (TheMatInfos.Reflection > 0) Then
       ValSpecular = (Shader_GetSpecularity(FromPoint, TheSphereLights(CurLight).TmpPos, NormalVector, IntersectionPoint) * (TheMatInfos.Reflection * AlphaFactor))
       CurrentColor = ColorAdd(CurrentColor, ColorScale(TheSphereLights(CurLight).Color, (TheMatInfos.SpecularPowerK * (ValSpecular ^ TheMatInfos.SpecularPowerN))))
      End If
      CurrentColor = ColorScale(CurrentColor, (ValShape * ValShadow))
      EndColor = ColorLimit(ColorAdd(EndColor, CurrentColor))
      If (ColorCompare(EndColor, ColorWhite) = True) Then 'The maximum light generated by the monitor!
       CalculateDirectLightContribution = ColorWhite: Exit Function
      End If
     End If
    End If
   End If
  End If
 Next CurLight

 '////////////////////////////// CONE LIGHTS ILLUMINATION ////////////////////////////////

Jump1:
 If (TheConeLightsCount = -1) Then CalculateDirectLightContribution = EndColor: Exit Function
 For CurLight = 0 To TheConeLightsCount
  If (TheConeLights(CurLight).Enable = True) Then
   'Calculate the geometrical attenuation of light, a cone in this case:
   ValShape = Shader_GetShapeCone(TheConeLights(CurLight).TmpPos, TheConeLights(CurLight).TmpDir, TheConeLights(CurLight).Falloff, TheConeLights(CurLight).Hotspot, TheConeLights(CurLight).Range, IntersectionPoint)
   If (ValShape <> 0) Then 'Intersected light volume
    'Compute shadows :
    If (EnableAreaShadows = True) Then MakeShadowRaysAsCone TheConeLights(CurLight)
    ValShadow = Shader_GetShadow(TheConeLights(CurLight).TmpPos, TheIntersection, IntersectionPoint)
    If (ValShadow <> 0) Then
     'Calculate the diffusion factor (lambert angle):
     ValLambert = Shader_GetLambertAngle(FromPoint, TheConeLights(CurLight).TmpPos, NormalVector, IntersectionPoint)
     If (ValLambert <> 0) Then
      'Calculate the output color with the given illumination factors:
      CurrentColor = ColorScale(ColorAbsorp(TheMatInfos.Color, TheConeLights(CurLight).Color), ValLambert)
      'Calculate the light's specularity factor:
      If (TheMatInfos.Reflection > 0) Then
       ValSpecular = (Shader_GetSpecularity(FromPoint, TheConeLights(CurLight).TmpPos, NormalVector, IntersectionPoint) * (TheMatInfos.Reflection * AlphaFactor))
       CurrentColor = ColorAdd(CurrentColor, ColorScale(TheConeLights(CurLight).Color, (TheMatInfos.SpecularPowerK * (ValSpecular ^ TheMatInfos.SpecularPowerN))))
      End If
      CurrentColor = ColorScale(CurrentColor, (ValShape * ValShadow))
      EndColor = ColorLimit(ColorAdd(EndColor, CurrentColor))
      If (ColorCompare(EndColor, ColorWhite) = True) Then 'The maximum light generated by the monitor!
       CalculateDirectLightContribution = ColorWhite: Exit Function
      End If
     End If
    End If
   End If
  End If
 Next CurLight

 CalculateDirectLightContribution = EndColor

End Function
Sub Engine_AllocateMemory()

'///////////////////////////////////////////////////////////////////
'///////////////// Memory-allocation limitations ///////////////////
'///////////////////////////////////////////////////////////////////

 If (UserAllocation = True) Then GoTo Jump

 'Set the default memory limitations:
 MaxVertices = 5000
 MaxFaces = 5000
 MaxMeshs = 100
 MaxLights = 100
 MaxCameras = 100
 MaxSplines = 500

Jump:

 'Geometry, lights and cameras:
 ReDim TheVertices(MaxVertices) As Vertex3D
 ReDim TheFaces(MaxFaces) As Face3D
 ReDim TheMeshs(MaxMeshs) As Mesh3D
 ReDim TheMaterials(MaxMeshs) As Material
 ReDim TheSphereLights(MaxLights) As SphereLight3D
 ReDim TheConeLights(MaxLights) As ConeLight3D
 ReDim TheCameras(MaxCameras) As Camera3D
 ReDim TheSplines(MaxSplines) As Spline3D

 'Textures
 ReDim TheAlphaTextures(MaxMeshs) As BitMap2D
 ReDim TheAlphaUsed(MaxMeshs) As Boolean
 ReDim TheColorTextures(MaxMeshs) As BitMap2D
 ReDim TheColorMips(MaxMeshs) As MipTextures
 ReDim TheColorUsed(MaxMeshs) As Boolean
 ReDim TheReflectionTextures(MaxMeshs) As BitMap2D
 ReDim TheReflectionUsed(MaxMeshs) As Boolean
 ReDim TheRefractionTextures(MaxMeshs) As BitMap2D
 ReDim TheRefractionUsed(MaxMeshs) As Boolean
 ReDim TheRefractionNTextures(MaxMeshs) As BitMap2D
 ReDim TheRefractionNUsed(MaxMeshs) As Boolean

End Sub
Sub Engine_ComputeBoundingBoxes()

 If ((Started = False) Or (TheMeshsCount = -1)) Then Exit Sub

 'Compute the bounding boxes, one of for every mesh,
 'another one is used as global bounding box.

 Dim CurMesh&, CurVertex&

 'Compute meshs's bounding boxes:
 ReDim MeshsBoundingBoxes(TheMeshsCount)
 For CurMesh = 0 To TheMeshsCount
  If (TheMeshs(CurMesh).Visible = True) Then
   MeshsBoundingBoxes(CurMesh) = Mesh3D_ComputeBoundingBox(CurMesh, False)
  End If
 Next CurMesh

 If (TheMeshsCount = 0) Then Exit Sub

 'Compute scene's bounding Boxes:
 SceneBoundingBox.Position = VectorInput(MaxSingleFloat, MaxSingleFloat, MaxSingleFloat)
 SceneBoundingBox.Direction = VectorInput(-MaxSingleFloat, -MaxSingleFloat, -MaxSingleFloat)
 For CurMesh = 0 To TheMeshsCount
  If (TheMeshs(CurMesh).Visible = True) Then
   For CurVertex = TheMeshs(CurMesh).Vertices.Start To GetAddressLast(TheMeshs(CurMesh).Vertices)
    If (TheVertices(CurVertex).TmpPos.X < SceneBoundingBox.Position.X) Then SceneBoundingBox.Position.X = TheVertices(CurVertex).TmpPos.X
    If (TheVertices(CurVertex).TmpPos.Y < SceneBoundingBox.Position.Y) Then SceneBoundingBox.Position.Y = TheVertices(CurVertex).TmpPos.Y
    If (TheVertices(CurVertex).TmpPos.Z < SceneBoundingBox.Position.Z) Then SceneBoundingBox.Position.Z = TheVertices(CurVertex).TmpPos.Z
    If (TheVertices(CurVertex).TmpPos.X > SceneBoundingBox.Direction.X) Then SceneBoundingBox.Direction.X = TheVertices(CurVertex).TmpPos.X
    If (TheVertices(CurVertex).TmpPos.Y > SceneBoundingBox.Direction.Y) Then SceneBoundingBox.Direction.Y = TheVertices(CurVertex).TmpPos.Y
    If (TheVertices(CurVertex).TmpPos.Z > SceneBoundingBox.Direction.Z) Then SceneBoundingBox.Direction.Z = TheVertices(CurVertex).TmpPos.Z
   Next CurVertex
  End If
 Next CurMesh

End Sub
Function Engine_ComputeUsedMemory(MemoryType As Byte) As Long

 If (Started = False) Then Exit Function

 'Return the length of the used memory in bytes

 Const BytesPerVertex As Integer = 24
 Const BytesPerFace As Integer = 146
 Const BytesPerMesh As Integer = 120
 Const BytesPerMaterial As Integer = 50
 Const BytesPerBitmap As Integer = 11
 Const BytesPerSphereLight As Integer = 36
 Const BytesPerConeLight As Integer = 68
 Const BytesPerCamera As Integer = 108
 Const BytesPerPhoton As Integer = 36

 Dim VerticesBytes&, FacesBytes&, MeshsBytes&, LightsBytes&, CamerasBytes&
 Dim TexturesBytes&, PhotonMapBytes&, MapBytes&, BackGroundBytes&
 Dim CurMesh&, CurVertex&, CurFace&, CurLight&, CurCamera&, CurMip&

 '///////////////////////////// GEOMETRY ///////////////////////////////

 If (TheMeshsCount = -1) Then GoTo Jump1
 For CurMesh = 0 To TheMeshsCount
  VerticesBytes = (VerticesBytes + (TheMeshs(CurMesh).Vertices.Length * BytesPerVertex))
  FacesBytes = (FacesBytes + (TheMeshs(CurMesh).Faces.Length * BytesPerFace))
  MeshsBytes = ((MeshsBytes + BytesPerMesh) + Len(TheMeshs(CurMesh).Label))
  MeshsBytes = ((MeshsBytes + BytesPerMaterial) + Len(TheMaterials(CurMesh).Label))
  With TheMaterials(CurMesh)
   If ((.UseColorTexture = True) And BitMap2D_IsValid(TheColorTextures(CurMesh)) = True) Then
    MapBytes = (CLng(TheColorTextures(CurMesh).Dimensions.X) * TheColorTextures(CurMesh).Dimensions.Y): MapBytes = (CLng(MapBytes) * 3)
    TexturesBytes = ((TexturesBytes + BytesPerBitmap) + Len(TheColorTextures(CurMesh).Label))
    TexturesBytes = (TexturesBytes + MapBytes)
    For CurMip = 0 To MipMapsLevel
     MapBytes = (CLng(TheColorMips(CurMesh).MipSequance(CurMip).Dimensions.X) * TheColorMips(CurMesh).MipSequance(CurMip).Dimensions.Y): MapBytes = (CLng(MapBytes) * 3)
     TexturesBytes = ((TexturesBytes + BytesPerBitmap) + Len(TheColorMips(CurMesh).MipSequance(CurMip).Label))
     TexturesBytes = (TexturesBytes + MapBytes)
    Next CurMip
   End If
   If ((.UseAlphaTexture = True) And BitMap2D_IsValid(TheAlphaTextures(CurMesh)) = True) Then
    MapBytes = (CLng(TheAlphaTextures(CurMesh).Dimensions.X) * TheAlphaTextures(CurMesh).Dimensions.Y)
    TexturesBytes = ((TexturesBytes + BytesPerBitmap) + Len(TheAlphaTextures(CurMesh).Label))
    TexturesBytes = (TexturesBytes + MapBytes)
   End If
   If ((.UseReflectionTexture = True) And BitMap2D_IsValid(TheReflectionTextures(CurMesh)) = True) Then
    MapBytes = (CLng(TheReflectionTextures(CurMesh).Dimensions.X) * TheReflectionTextures(CurMesh).Dimensions.Y)
    TexturesBytes = ((TexturesBytes + BytesPerBitmap) + Len(TheReflectionTextures(CurMesh).Label))
    TexturesBytes = (TexturesBytes + MapBytes)
   End If
   If ((.UseRefractionTexture = True) And BitMap2D_IsValid(TheRefractionTextures(CurMesh)) = True) Then
    MapBytes = (CLng(TheRefractionTextures(CurMesh).Dimensions.X) * TheRefractionTextures(CurMesh).Dimensions.Y)
    TexturesBytes = ((TexturesBytes + BytesPerBitmap) + Len(TheRefractionTextures(CurMesh).Label))
    TexturesBytes = (TexturesBytes + MapBytes)
   End If
   If ((.UseRefractionNTexture = True) And BitMap2D_IsValid(TheRefractionNTextures(CurMesh)) = True) Then
    MapBytes = (CLng(TheRefractionNTextures(CurMesh).Dimensions.X) * TheRefractionNTextures(CurMesh).Dimensions.Y)
    TexturesBytes = ((TexturesBytes + BytesPerBitmap) + Len(TheRefractionNTextures(CurMesh).Label))
    TexturesBytes = (TexturesBytes + MapBytes)
   End If
  End With
 Next CurMesh

Jump1:

 '///////////////////////////// LIGHTS ///////////////////////////////

 If (TheSphereLightsCount = -1) Then GoTo Jump2
 For CurLight = 0 To TheSphereLightsCount
  LightsBytes = ((LightsBytes + BytesPerSphereLight) + Len(TheSphereLights(CurLight).Label))
 Next CurLight

Jump2:

 If (TheConeLightsCount = -1) Then GoTo Jump3
 For CurLight = 0 To TheConeLightsCount
  LightsBytes = ((LightsBytes + BytesPerConeLight) + Len(TheConeLights(CurLight).Label))
 Next CurLight

Jump3:

 '///////////////////////////// CAMERAS //////////////////////////////

 If (TheCamerasCount = -1) Then GoTo Jump4
 For CurCamera = 0 To TheCamerasCount
  CamerasBytes = ((CamerasBytes + BytesPerCamera) + Len(TheCameras(CurCamera).Label))
 Next CurCamera

Jump4:

 '/////////////////////////// PHOTON MAP/////////////////////////////

 If (EnablePhotonMapping = False) Then GoTo Jump5
 PhotonMapBytes = (MaximumAllocatedPhotons * BytesPerPhoton)

Jump5:

 '/////////////////////////// BACK GROUND /////////////////////////////

 If (UseBackGround = False) Then GoTo Jump6
 BackGroundBytes = (CLng(OriginalBackGround.Dimensions.X) * OriginalBackGround.Dimensions.Y): BackGroundBytes = (CLng(BackGroundBytes) * 3)

 '//////////////////////////////////////////////////////////////////////

Jump6:

 Select Case MemoryType
  Case 0: Engine_ComputeUsedMemory = (VerticesBytes + FacesBytes + MeshsBytes + TexturesBytes + LightsBytes + CamerasBytes + PhotonMapBytes + BackGroundBytes)
  Case 1: Engine_ComputeUsedMemory = VerticesBytes
  Case 2: Engine_ComputeUsedMemory = FacesBytes
  Case 3: Engine_ComputeUsedMemory = MeshsBytes
  Case 4: Engine_ComputeUsedMemory = TexturesBytes
  Case 5: Engine_ComputeUsedMemory = LightsBytes
  Case 6: Engine_ComputeUsedMemory = CamerasBytes
  Case 7: Engine_ComputeUsedMemory = PhotonMapBytes
  Case 8: Engine_ComputeUsedMemory = BackGroundBytes
 End Select

End Function
Function Engine_LoadScene(TheFileName As String) As Boolean

 If (Started = False) Then Exit Function

 'Load scene from the source file.

 If (FileExist(TheFileName) = False) Then
  MsgBox "No file was selected !", vbCritical, "Load scene": Exit Function
 End If

 If (Right(TheFileName, 4) <> SceneFileExtension) Then
  MsgBox "Invalid file !", vbCritical, "Load scene": Exit Function
 End If

 Dim CurMesh&, CurVertex&, CurFace&, CurLight&, CurCamera&
 Dim MeshsMax&, VerticesMax&, FacesMax&, OLightsMax&, SLightsMax&, CamerasMax&
 Dim TheIndex&, TheByte As Byte, TheInteger%, TheString$, TmpBitMap As BitMap2D

 '========================================================================

 Open TheFileName For Binary Access Read Lock Read Write As 1

  'APP/FILE PASSWORD:

  'Test for the valid file format, a simple password in
  'application/file level will do the thing, the password is saved
  'with a simple encryption (just an inversion of the ASCII code)
  'we use the encryption for any reader can't observe the
  'app/file password, when visualizing the file with notepad.exe, for example.
  TheString = String(Len(AppFilePassword), " ")
  Get 1, , TheString: TheString = SimplyCrypt(TheString)
  If (TheString <> AppFilePassword) Then
   MsgBox "Invalid file !", vbCritical, "Load scene": Close 1: Exit Function
  End If

  'USER/FILE PASSWORD:
  Get 1, , TheByte
  If (TheByte = 1) Then
   Get 1, , TheInteger: TheString = String(TheInteger, " "): Get 1, , TheString
   TheString = SimplyCrypt(TheString) 'Decode the password
   If (InputBox("Type the password:", "Protected file") <> TheString) Then
    MsgBox "Invalid password, can't load the file !", vbCritical, "Wrong password": Close 1: Exit Function
   End If
  End If

  FRM_Main.DoReset False
  FRM_Progress.Show
  FRM_Progress.DisplayLoadProgress 1, 0, 0

  'ALLOCATIONS:
  Get 1, , MaxVertices: Get 1, , MaxFaces
  Get 1, , MaxMeshs: Get 1, , MaxLights: Get 1, , MaxCameras
  UserAllocation = True: Engine_Reset: Camera3D_Remove 0

  'OUTPUT SIZE:
  Get 1, , OutputWidth: Get 1, , OutputHeight

  'PATH-TRACER:
  Get 1, , ViewPathsPerPixel: Get 1, , SamplesPerViewPath

  'SHADOWS:
  Get 1, , TheByte
  If (TheByte = 1) Then
   EnableAreaShadows = False
  ElseIf (TheByte = 2) Then
   EnableAreaShadows = True: Get 1, , ShadowRaysCount: Get 1, , ShadowsApproxRadius
  End If

  'PHOTON MAPPING:
  Get 1, , TheAmbiantLight: Get 1, , TheByte
  If (TheByte = 1) Then
   EnablePhotonMapping = False
  ElseIf (TheByte = 2) Then
   EnablePhotonMapping = True
   Get 1, , MaximumAllocatedPhotons: ReDim ThePhotonMap(MaximumAllocatedPhotons)
   Get 1, , SamplesPerPhotonPath: Get 1, , PhotonsSearchRadius
   Get 1, , BleedingDistance: Get 1, , EstimateMultiplier
  End If

  'FOG:
  Get 1, , TheByte
  If (TheByte = 1) Then
   FogEnable = False
  ElseIf (TheByte = 2) Then
   FogEnable = True: Get 1, , FogRange: Get 1, , FogColor: Get 1, , TheByte
   If (TheByte = 1) Then
    FogMode = K3DE_FM_LINEAR
   ElseIf (TheByte = 2) Then
    FogMode = K3DE_FM_EXP: Get 1, , FogExpFactor1: Get 1, , FogExpFactor2
   End If
  End If

  'FILTERING:
  Get 1, , TheByte
  Select Case TheByte
   Case 1:
    TheTexturesFilter = K3DE_TFM_NEAREST
   Case 2:
    TheTexturesFilter = K3DE_TFM_NEAREST_MIP_NEAREST
    Get 1, , MipMapsLevel: Get 1, , MipMapsMinPurcent
   Case 3:
    TheTexturesFilter = K3DE_TFM_NEAREST_MIP_LINEAR
    Get 1, , MipMapsLevel: Get 1, , MipMapsMinPurcent
   Case 4:
    TheTexturesFilter = K3DE_TFM_FILTERED
    Get 1, , TheByte
    Select Case TheByte
     Case 1: TheTexelsFilter = K3DE_XFM_BILINEAR
     Case 2: TheTexelsFilter = K3DE_XFM_BELL
     Case 3: TheTexelsFilter = K3DE_XFM_GAUSSIAN:              Get 1, , KernelSize
     Case 4: TheTexelsFilter = K3DE_XFM_CUBIC_SPLINE_B
     Case 5: TheTexelsFilter = K3DE_XFM_CUBIC_SPLINE_BC:       Get 1, , CubicB: Get 1, , CubicC
     Case 6: TheTexelsFilter = K3DE_XFM_CUBIC_SPLINE_CARDINAL: Get 1, , CubicA
    End Select
   Case 5:
    TheTexturesFilter = K3DE_TFM_FILTERED_MIP_NEAREST
    Get 1, , TheByte
    Select Case TheByte
     Case 1: TheTexelsFilter = K3DE_XFM_BILINEAR
     Case 2: TheTexelsFilter = K3DE_XFM_BELL
     Case 3: TheTexelsFilter = K3DE_XFM_GAUSSIAN:              Get 1, , KernelSize
     Case 4: TheTexelsFilter = K3DE_XFM_CUBIC_SPLINE_B
     Case 5: TheTexelsFilter = K3DE_XFM_CUBIC_SPLINE_BC:       Get 1, , CubicB: Get 1, , CubicC
     Case 6: TheTexelsFilter = K3DE_XFM_CUBIC_SPLINE_CARDINAL: Get 1, , CubicA
    End Select
    Get 1, , MipMapsLevel: Get 1, , MipMapsMinPurcent
   Case 6:
    TheTexturesFilter = K3DE_TFM_FILTERED_MIP_LINEAR
    Get 1, , TheByte
    Select Case TheByte
     Case 1: TheTexelsFilter = K3DE_XFM_BILINEAR
     Case 2: TheTexelsFilter = K3DE_XFM_BELL
     Case 3: TheTexelsFilter = K3DE_XFM_GAUSSIAN:              Get 1, , KernelSize
     Case 4: TheTexelsFilter = K3DE_XFM_CUBIC_SPLINE_B
     Case 5: TheTexelsFilter = K3DE_XFM_CUBIC_SPLINE_BC:       Get 1, , CubicB: Get 1, , CubicC
     Case 6: TheTexelsFilter = K3DE_XFM_CUBIC_SPLINE_CARDINAL: Get 1, , CubicA
    End Select
    Get 1, , MipMapsLevel: Get 1, , MipMapsMinPurcent
  End Select

  FRM_Progress.DisplayLoadProgress 1, (1 / 6), 1
  FRM_Progress.DisplayLoadProgress 2, (1 / 6), 0

  'BACKGROUND:
  Get 1, , TheByte
  If (TheByte = 1) Then
   UseBackGround = False
  ElseIf (TheByte = 2) Then
   UseBackGround = True
   Get 1, , TheInteger
   OriginalBackGround.Label = String(TheInteger, " "): Get 1, , OriginalBackGround.Label
   Get 1, , OriginalBackGround.BitsDepth
   Get 1, , OriginalBackGround.Dimensions
   Get 1, , OriginalBackGround.BackGroundColor
   If (OriginalBackGround.BitsDepth = 8) Then
    ReDim OriginalBackGround.Datas(0, OriginalBackGround.Dimensions.X, OriginalBackGround.Dimensions.Y)
   ElseIf (OriginalBackGround.BitsDepth = 24) Then
    ReDim OriginalBackGround.Datas(2, OriginalBackGround.Dimensions.X, OriginalBackGround.Dimensions.Y)
   End If
   Get 1, , OriginalBackGround.Datas()
  End If
  Get 1, , TheBackGroundColor

  FRM_Progress.DisplayLoadProgress 2, (2 / 6), 1
  FRM_Progress.DisplayLoadProgress 3, (2 / 6), 0

  'MESHS:
  Get 1, , MeshsMax
  If (MeshsMax = -1) Then GoTo Jump1:
  For CurMesh = 0 To MeshsMax

   'VERTICES & FACES COUNT:
   Get 1, , VerticesMax: Get 1, , FacesMax
   TheIndex = Mesh3D_Add(VerticesMax, FacesMax)

   'VERTICES:
   For CurVertex = TheMeshs(TheIndex).Vertices.Start To GetAddressLast(TheMeshs(TheIndex).Vertices)
    Get 1, , TheVertices(CurVertex).Position
   Next CurVertex

   'FACES:
   For CurFace = TheMeshs(TheIndex).Faces.Start To GetAddressLast(TheMeshs(TheIndex).Faces)
    Get 1, , TheByte: If (TheByte = 1) Then TheFaces(CurFace).Visible = False Else If (TheByte = 2) Then TheFaces(CurFace).Visible = True
    Get 1, , TheFaces(CurFace).A: Get 1, , TheFaces(CurFace).B: Get 1, , TheFaces(CurFace).C
    Get 1, , TheFaces(CurFace).AlphaVectors
    Get 1, , TheFaces(CurFace).ColorVectors
    Get 1, , TheFaces(CurFace).ReflectionVectors
    Get 1, , TheFaces(CurFace).RefractionVectors
    Get 1, , TheFaces(CurFace).RefractionNVectors
   Next CurFace

   'MESH'S PROPERTIES:
   Get 1, , TheInteger
   TheMeshs(TheIndex).Label = String(TheInteger, " "): Get 1, , TheMeshs(TheIndex).Label
   Get 1, , TheMeshs(TheIndex).Position
   Get 1, , TheMeshs(TheIndex).Scales
   Get 1, , TheMeshs(TheIndex).Angles
   Get 1, , TheByte
   If (TheByte = 1) Then
    TheMeshs(TheIndex).MakeMatrix = False
    Get 1, , TheMeshs(TheIndex).WorldMatrix
   ElseIf (TheByte = 2) Then
    TheMeshs(TheIndex).MakeMatrix = True
   End If
   Get 1, , TheByte
   If (TheByte = 1) Then
    TheMeshs(TheIndex).Visible = False
   ElseIf (TheByte = 2) Then
    TheMeshs(TheIndex).Visible = True
   End If

   'MATERIAL:
   Get 1, , TheInteger
   TheMaterials(TheIndex).Label = String(TheInteger, " "): Get 1, , TheMaterials(TheIndex).Label
   Get 1, , TheMaterials(TheIndex).Color
   Get 1, , TheMaterials(TheIndex).Reflection
   Get 1, , TheMaterials(TheIndex).Refraction
   Get 1, , TheMaterials(TheIndex).RefractionN
   Get 1, , TheMaterials(TheIndex).SpecularPowerK
   Get 1, , TheMaterials(TheIndex).SpecularPowerN

   'ALPHA MAP:
   Get 1, , TheByte
   If (TheByte = 1) Then
    TheMaterials(TheIndex).UseAlphaTexture = False
   ElseIf (TheByte = 2) Then
    TheMaterials(TheIndex).UseAlphaTexture = True
    BitMap2D_Delete TmpBitMap
    Get 1, , TheInteger
    TmpBitMap.Label = String(TheInteger, " "): Get 1, , TmpBitMap.Label
    Get 1, , TmpBitMap.BitsDepth
    Get 1, , TmpBitMap.Dimensions
    Get 1, , TmpBitMap.BackGroundColor
    If (TmpBitMap.BitsDepth = 8) Then
     ReDim TmpBitMap.Datas(0, TmpBitMap.Dimensions.X, TmpBitMap.Dimensions.Y)
    ElseIf (TmpBitMap.BitsDepth = 24) Then
     ReDim TmpBitMap.Datas(2, TmpBitMap.Dimensions.X, TmpBitMap.Dimensions.Y)
    End If
    Get 1, , TmpBitMap.Datas()
    TEX_Alpha_Set TheIndex, TmpBitMap
   End If

   'COLOR MAP:
   Get 1, , TheByte
   If (TheByte = 1) Then
    TheMaterials(TheIndex).UseColorTexture = False
   ElseIf (TheByte = 2) Then
    TheMaterials(TheIndex).UseColorTexture = True
    BitMap2D_Delete TmpBitMap
    Get 1, , TheInteger
    TmpBitMap.Label = String(TheInteger, " "): Get 1, , TmpBitMap.Label
    Get 1, , TmpBitMap.BitsDepth
    Get 1, , TmpBitMap.Dimensions
    Get 1, , TmpBitMap.BackGroundColor
    If (TmpBitMap.BitsDepth = 8) Then
     ReDim TmpBitMap.Datas(0, TmpBitMap.Dimensions.X, TmpBitMap.Dimensions.Y)
    ElseIf (TmpBitMap.BitsDepth = 24) Then
     ReDim TmpBitMap.Datas(2, TmpBitMap.Dimensions.X, TmpBitMap.Dimensions.Y)
    End If
    Get 1, , TmpBitMap.Datas()
    TEX_Color_Set TheIndex, TmpBitMap
   End If

   'REFLECTION MAP:
   Get 1, , TheByte
   If (TheByte = 1) Then
    TheMaterials(TheIndex).UseReflectionTexture = False
   ElseIf (TheByte = 2) Then
    TheMaterials(TheIndex).UseReflectionTexture = True
    BitMap2D_Delete TmpBitMap
    Get 1, , TheInteger
    TmpBitMap.Label = String(TheInteger, " "): Get 1, , TmpBitMap.Label
    Get 1, , TmpBitMap.BitsDepth
    Get 1, , TmpBitMap.Dimensions
    Get 1, , TmpBitMap.BackGroundColor
    If (TmpBitMap.BitsDepth = 8) Then
     ReDim TmpBitMap.Datas(0, TmpBitMap.Dimensions.X, TmpBitMap.Dimensions.Y)
    ElseIf (TmpBitMap.BitsDepth = 24) Then
     ReDim TmpBitMap.Datas(2, TmpBitMap.Dimensions.X, TmpBitMap.Dimensions.Y)
    End If
    Get 1, , TmpBitMap.Datas()
    TEX_Reflection_Set TheIndex, TmpBitMap
   End If

   'REFRACTION MAP:
   Get 1, , TheByte
   If (TheByte = 1) Then
    TheMaterials(TheIndex).UseRefractionTexture = False
   ElseIf (TheByte = 2) Then
    TheMaterials(TheIndex).UseRefractionTexture = True
    BitMap2D_Delete TmpBitMap
    Get 1, , TheInteger
    TmpBitMap.Label = String(TheInteger, " "): Get 1, , TmpBitMap.Label
    Get 1, , TmpBitMap.BitsDepth
    Get 1, , TmpBitMap.Dimensions
    Get 1, , TmpBitMap.BackGroundColor
    If (TmpBitMap.BitsDepth = 8) Then
     ReDim TmpBitMap.Datas(0, TmpBitMap.Dimensions.X, TmpBitMap.Dimensions.Y)
    ElseIf (TmpBitMap.BitsDepth = 24) Then
     ReDim TmpBitMap.Datas(2, TmpBitMap.Dimensions.X, TmpBitMap.Dimensions.Y)
    End If
    Get 1, , TmpBitMap.Datas()
    TEX_Refraction_Set TheIndex, TmpBitMap
   End If

   'REFRACTION-N MAP:
   Get 1, , TheByte
   If (TheByte = 1) Then
    TheMaterials(TheIndex).UseRefractionNTexture = False
   ElseIf (TheByte = 2) Then
    TheMaterials(TheIndex).UseRefractionNTexture = True
    BitMap2D_Delete TmpBitMap
    Get 1, , TheInteger
    TmpBitMap.Label = String(TheInteger, " "): Get 1, , TmpBitMap.Label
    Get 1, , TmpBitMap.BitsDepth
    Get 1, , TmpBitMap.Dimensions
    Get 1, , TmpBitMap.BackGroundColor
    If (TmpBitMap.BitsDepth = 8) Then
     ReDim TmpBitMap.Datas(0, TmpBitMap.Dimensions.X, TmpBitMap.Dimensions.Y)
    ElseIf (TmpBitMap.BitsDepth = 24) Then
     ReDim TmpBitMap.Datas(2, TmpBitMap.Dimensions.X, TmpBitMap.Dimensions.Y)
    End If
    Get 1, , TmpBitMap.Datas()
    TEX_RefractionN_Set TheIndex, TmpBitMap
   End If
   If ((MeshsMax > 0) And (CurMesh <> MeshsMax)) Then
    FRM_Progress.DisplayLoadProgress 3, (2 / 6), CSng(CurMesh / MeshsMax)
   End If
  Next CurMesh

  FRM_Progress.DisplayLoadProgress 3, (3 / 6), 1
  FRM_Progress.DisplayLoadProgress 4, (3 / 6), 0

Jump1:

  'OMNI LIGHTS:
  Get 1, , OLightsMax
  If (OLightsMax = -1) Then GoTo Jump2:
  For CurLight = 0 To OLightsMax
   TheIndex = SphereLight3D_Add
   Get 1, , TheInteger: TheSphereLights(TheIndex).Label = String(TheInteger, " ")
   Get 1, , TheSphereLights(TheIndex).Label
   Get 1, , TheSphereLights(TheIndex).Color
   Get 1, , TheSphereLights(TheIndex).Position
   Get 1, , TheSphereLights(TheIndex).Range
   Get 1, , TheByte
   If (TheByte = 1) Then
    TheSphereLights(TheIndex).Enable = False
   ElseIf (TheByte = 2) Then
    TheSphereLights(TheIndex).Enable = True
   End If
   If ((OLightsMax > 0) And (CurLight <> OLightsMax)) Then
    FRM_Progress.DisplayLoadProgress 4, (3 / 6), CSng(CurLight / OLightsMax)
   End If
  Next CurLight

  FRM_Progress.DisplayLoadProgress 4, (4 / 6), 1
  FRM_Progress.DisplayLoadProgress 5, (4 / 6), 0

Jump2:

  'SPOT LIGHTS:
  Get 1, , SLightsMax
  If (SLightsMax = -1) Then GoTo Jump3:
  For CurLight = 0 To SLightsMax
   TheIndex = ConeLight3D_Add
   Get 1, , TheInteger: TheConeLights(TheIndex).Label = String(TheInteger, " ")
   Get 1, , TheConeLights(TheIndex).Label
   Get 1, , TheConeLights(TheIndex).Color
   Get 1, , TheConeLights(TheIndex).Position
   Get 1, , TheConeLights(TheIndex).Direction
   Get 1, , TheConeLights(TheIndex).Falloff
   Get 1, , TheConeLights(TheIndex).Hotspot
   Get 1, , TheConeLights(TheIndex).Range
   Get 1, , TheByte
   If (TheByte = 1) Then
    TheConeLights(TheIndex).Enable = False
   ElseIf (TheByte = 2) Then
    TheConeLights(TheIndex).Enable = True
   End If
   If ((SLightsMax > 0) And (CurLight <> SLightsMax)) Then
    FRM_Progress.DisplayLoadProgress 5, (4 / 6), CSng(CurLight / SLightsMax)
   End If
  Next CurLight

  FRM_Progress.DisplayLoadProgress 5, (5 / 6), 1
  FRM_Progress.DisplayLoadProgress 6, (5 / 6), 0

Jump3:

  'CAMERAS:
  Get 1, , CamerasMax
  For CurCamera = 0 To CamerasMax
   TheIndex = Camera3D_Add
   Get 1, , TheInteger: TheCameras(TheIndex).Label = String(TheInteger, " ")
   Get 1, , TheCameras(TheIndex).Label
   Get 1, , TheCameras(TheIndex).Position
   Get 1, , TheCameras(TheIndex).Direction
   Get 1, , TheCameras(TheIndex).RollAngle
   Get 1, , TheCameras(TheIndex).FOVAngle
   Get 1, , TheCameras(TheIndex).ClearDistance
   Get 1, , TheCameras(TheIndex).Dispersion
   Get 1, , TheByte
   If (TheByte = 1) Then
    TheCameras(CurCamera).BackFaceCulling = False
   ElseIf (TheByte = 2) Then
    TheCameras(CurCamera).BackFaceCulling = True
   End If
   Get 1, , TheByte
   If (TheByte = 1) Then
    TheCameras(CurCamera).MakeMatrix = False
    Get 1, , TheCameras(CurCamera).ViewMatrix
   ElseIf (TheByte = 2) Then
    TheCameras(CurCamera).MakeMatrix = True
   End If
   If ((CamerasMax > 0) And (CurCamera <> CamerasMax)) Then
    FRM_Progress.DisplayLoadProgress 6, (5 / 6), CSng(CurCamera / CamerasMax)
   End If
  Next CurCamera

  'VIEWPORTS:
  Get 1, , TheCurrentCamera
  Get 1, , DisplayMode
  Get 1, , TheByte: If (TheByte = 1) Then FRM_Main.Check1.Value = vbUnchecked Else If (TheByte = 2) Then FRM_Main.Check1.Value = vbChecked
  Get 1, , TheByte: If (TheByte = 1) Then FRM_Main.Check2.Value = vbUnchecked Else If (TheByte = 2) Then FRM_Main.Check2.Value = vbChecked
  Get 1, , TheByte: If (TheByte = 1) Then FRM_Main.Check3.Value = vbUnchecked Else If (TheByte = 2) Then FRM_Main.Check3.Value = vbChecked
  Get 1, , TheByte: If (TheByte = 1) Then FRM_Main.Check4.Value = vbUnchecked Else If (TheByte = 2) Then FRM_Main.Check4.Value = vbChecked
  Get 1, , TheByte: If (TheByte = 1) Then FRM_Main.Check5.Value = vbUnchecked Else If (TheByte = 2) Then FRM_Main.Check5.Value = vbChecked
  Get 1, , TheByte: If (TheByte = 1) Then FRM_Main.Check6.Value = vbUnchecked Else If (TheByte = 2) Then FRM_Main.Check6.Value = vbChecked

  FRM_Progress.DisplayLoadProgress 6, 1, 1

 Close 1

 FRM_Main.ChooseDisplay DisplayMode
 Unload FRM_Progress
 Engine_LoadScene = True
 FRM_Main.RefreshViews

End Function
Function Engine_LoadMesh(TheFileName As String) As Boolean

 If (Started = False) Then Exit Function

 'Load mesh from the source file.

 If (FileExist(TheFileName) = False) Then
  MsgBox "No file was selected !", vbCritical, "Load object": Exit Function
 End If

 If (Right(TheFileName, 4) <> ObjectFileExtension) Then
  MsgBox "Invalid file !", vbCritical, "Load object": Exit Function
 End If

 Dim CurMesh&, CurVertex&, CurFace&, CurLight&, CurCamera&
 Dim MeshsMax&, VerticesMax&, FacesMax&, OLightsMax&, SLightsMax&, CamerasMax&
 Dim TheIndex&, TheByte As Byte, TheInteger%, TheLong&, TheString$
 Dim TmpBitMap As BitMap2D

 '========================================================================

 Open TheFileName For Binary Access Read Lock Read Write As 1

  'APP/FILE PASSWORD:
  'Test for the valid file format, a simple password in
  'application/file level will do the thing, the password is saved
  'with a simple encryption (just an inversion of the ASCII code)
  'we use the encryption for any reader can't observe the
  'app/file password, when visualizing the file with notepad.exe, for example.
  TheString = String(Len(AppFilePassword), " ")
  Get 1, , TheString: TheString = SimplyCrypt(TheString)
  If (TheString <> AppFilePassword) Then
   MsgBox "Invalid file !", vbCritical, "Load object": Close 1: Exit Function
  End If

  'USER/FILE PASSWORD:
  Get 1, , TheByte
  If (TheByte = 1) Then
   Get 1, , TheInteger: TheString = String(TheInteger, " "): Get 1, , TheString
   TheString = SimplyCrypt(TheString) 'Decode the password
   If (InputBox("Type the password:", "Protected file") <> TheString) Then
    MsgBox "Invalid password, can't load the file !", vbCritical, "Wrong password": Close 1: Exit Function
   End If
  End If

  'VERTICES & FACES COUNT:
  Get 1, , VerticesMax: Get 1, , FacesMax
  TheIndex = Mesh3D_Add(VerticesMax, FacesMax)
  If (TheIndex = -1) Then Close 1: Exit Function

  'VERTICES:
  For CurVertex = TheMeshs(TheIndex).Vertices.Start To GetAddressLast(TheMeshs(TheIndex).Vertices)
   Get 1, , TheVertices(CurVertex).Position
  Next CurVertex

  'FACES:
  For CurFace = TheMeshs(TheIndex).Faces.Start To GetAddressLast(TheMeshs(TheIndex).Faces)
   Get 1, , TheByte: If (TheByte = 1) Then TheFaces(CurFace).Visible = False Else If (TheByte = 2) Then TheFaces(CurFace).Visible = True
   Get 1, , TheLong: TheFaces(CurFace).A = TheLong + (TheMeshs(TheIndex).Vertices.Start - 1)
   Get 1, , TheLong: TheFaces(CurFace).B = TheLong + (TheMeshs(TheIndex).Vertices.Start - 1)
   Get 1, , TheLong: TheFaces(CurFace).C = TheLong + (TheMeshs(TheIndex).Vertices.Start - 1)
   Get 1, , TheFaces(CurFace).AlphaVectors
   Get 1, , TheFaces(CurFace).ColorVectors
   Get 1, , TheFaces(CurFace).ReflectionVectors
   Get 1, , TheFaces(CurFace).RefractionVectors
   Get 1, , TheFaces(CurFace).RefractionNVectors
  Next CurFace

  'MESH'S PROPERTIES:
  Get 1, , TheInteger
  TheMeshs(TheIndex).Label = String(TheInteger, " "): Get 1, , TheMeshs(TheIndex).Label
  Get 1, , TheMeshs(TheIndex).Position
  Get 1, , TheMeshs(TheIndex).Scales
  Get 1, , TheMeshs(TheIndex).Angles
  Get 1, , TheByte
  If (TheByte = 1) Then
   TheMeshs(TheIndex).MakeMatrix = False
   Get 1, , TheMeshs(TheIndex).WorldMatrix
  ElseIf (TheByte = 2) Then
   TheMeshs(TheIndex).MakeMatrix = True
  End If
  Get 1, , TheByte
  If (TheByte = 1) Then
   TheMeshs(TheIndex).Visible = False
  ElseIf (TheByte = 2) Then
   TheMeshs(TheIndex).Visible = True
  End If

  'MATERIAL:
  Get 1, , TheInteger
  TheMaterials(TheIndex).Label = String(TheInteger, " "): Get 1, , TheMaterials(TheIndex).Label
  Get 1, , TheMaterials(TheIndex).Color
  Get 1, , TheMaterials(TheIndex).Reflection
  Get 1, , TheMaterials(TheIndex).Refraction
  Get 1, , TheMaterials(TheIndex).RefractionN
  Get 1, , TheMaterials(TheIndex).SpecularPowerK
  Get 1, , TheMaterials(TheIndex).SpecularPowerN

  'ALPHA MAP:
  Get 1, , TheByte
  If (TheByte = 1) Then
   TheMaterials(TheIndex).UseAlphaTexture = False
  ElseIf (TheByte = 2) Then
   TheMaterials(TheIndex).UseAlphaTexture = True
   BitMap2D_Delete TmpBitMap
   Get 1, , TheInteger
   TmpBitMap.Label = String(TheInteger, " "): Get 1, , TmpBitMap.Label
   Get 1, , TmpBitMap.BitsDepth
   Get 1, , TmpBitMap.Dimensions
   Get 1, , TmpBitMap.BackGroundColor
   If (TmpBitMap.BitsDepth = 8) Then
    ReDim TmpBitMap.Datas(0, TmpBitMap.Dimensions.X, TmpBitMap.Dimensions.Y)
   ElseIf (TmpBitMap.BitsDepth = 24) Then
    ReDim TmpBitMap.Datas(2, TmpBitMap.Dimensions.X, TmpBitMap.Dimensions.Y)
   End If
   Get 1, , TmpBitMap.Datas()
   TEX_Alpha_Set TheIndex, TmpBitMap
  End If

  'COLOR MAP:
  Get 1, , TheByte
  If (TheByte = 1) Then
   TheMaterials(TheIndex).UseColorTexture = False
  ElseIf (TheByte = 2) Then
   TheMaterials(TheIndex).UseColorTexture = True
   BitMap2D_Delete TmpBitMap
   Get 1, , TheInteger
   TmpBitMap.Label = String(TheInteger, " "): Get 1, , TmpBitMap.Label
   Get 1, , TmpBitMap.BitsDepth
   Get 1, , TmpBitMap.Dimensions
   Get 1, , TmpBitMap.BackGroundColor
   If (TmpBitMap.BitsDepth = 8) Then
    ReDim TmpBitMap.Datas(0, TmpBitMap.Dimensions.X, TmpBitMap.Dimensions.Y)
   ElseIf (TmpBitMap.BitsDepth = 24) Then
    ReDim TmpBitMap.Datas(2, TmpBitMap.Dimensions.X, TmpBitMap.Dimensions.Y)
   End If
   Get 1, , TmpBitMap.Datas()
   TEX_Color_Set TheIndex, TmpBitMap
  End If

  'REFLECTION MAP:
  Get 1, , TheByte
  If (TheByte = 1) Then
   TheMaterials(TheIndex).UseReflectionTexture = False
  ElseIf (TheByte = 2) Then
   TheMaterials(TheIndex).UseReflectionTexture = True
   BitMap2D_Delete TmpBitMap
   Get 1, , TheInteger
   TmpBitMap.Label = String(TheInteger, " "): Get 1, , TmpBitMap.Label
   Get 1, , TmpBitMap.BitsDepth
   Get 1, , TmpBitMap.Dimensions
   Get 1, , TmpBitMap.BackGroundColor
   If (TmpBitMap.BitsDepth = 8) Then
    ReDim TmpBitMap.Datas(0, TmpBitMap.Dimensions.X, TmpBitMap.Dimensions.Y)
   ElseIf (TmpBitMap.BitsDepth = 24) Then
    ReDim TmpBitMap.Datas(2, TmpBitMap.Dimensions.X, TmpBitMap.Dimensions.Y)
   End If
   Get 1, , TmpBitMap.Datas()
   TEX_Reflection_Set TheIndex, TmpBitMap
  End If

  'REFRACTION MAP:
  Get 1, , TheByte
  If (TheByte = 1) Then
   TheMaterials(TheIndex).UseRefractionTexture = False
  ElseIf (TheByte = 2) Then
   TheMaterials(TheIndex).UseRefractionTexture = True
   BitMap2D_Delete TmpBitMap
   Get 1, , TheInteger
   TmpBitMap.Label = String(TheInteger, " "): Get 1, , TmpBitMap.Label
   Get 1, , TmpBitMap.BitsDepth
   Get 1, , TmpBitMap.Dimensions
   Get 1, , TmpBitMap.BackGroundColor
   If (TmpBitMap.BitsDepth = 8) Then
    ReDim TmpBitMap.Datas(0, TmpBitMap.Dimensions.X, TmpBitMap.Dimensions.Y)
   ElseIf (TmpBitMap.BitsDepth = 24) Then
    ReDim TmpBitMap.Datas(2, TmpBitMap.Dimensions.X, TmpBitMap.Dimensions.Y)
   End If
   Get 1, , TmpBitMap.Datas()
   TEX_Refraction_Set TheIndex, TmpBitMap
  End If

  'REFRACTION-N MAP:
  Get 1, , TheByte
  If (TheByte = 1) Then
   TheMaterials(TheIndex).UseRefractionNTexture = False
  ElseIf (TheByte = 2) Then
   TheMaterials(TheIndex).UseRefractionNTexture = True
   BitMap2D_Delete TmpBitMap
   Get 1, , TheInteger
   TmpBitMap.Label = String(TheInteger, " "): Get 1, , TmpBitMap.Label
   Get 1, , TmpBitMap.BitsDepth
   Get 1, , TmpBitMap.Dimensions
   Get 1, , TmpBitMap.BackGroundColor
   If (TmpBitMap.BitsDepth = 8) Then
    ReDim TmpBitMap.Datas(0, TmpBitMap.Dimensions.X, TmpBitMap.Dimensions.Y)
   ElseIf (TmpBitMap.BitsDepth = 24) Then
    ReDim TmpBitMap.Datas(2, TmpBitMap.Dimensions.X, TmpBitMap.Dimensions.Y)
   End If
   Get 1, , TmpBitMap.Datas()
   TEX_RefractionN_Set TheIndex, TmpBitMap
  End If

 Close 1

 Engine_LoadMesh = True

End Function
Sub Engine_SaveScene(TheFileName As String)

 If (Started = False) Then Exit Sub

 'Save the current scene to the destination file.

 If (Right(TheFileName, 4) <> SceneFileExtension) Then
  MsgBox "Invalid file !", vbCritical, "Save scene": Exit Sub
 End If

 If (FileExist(TheFileName) = True) Then
  If (MsgBox("File already exist, overwrite ?", (vbQuestion + vbYesNo), "Save scene") = vbYes) Then
   Kill TheFileName
  Else
   MsgBox "Aborted saving operation.", vbCritical, "Abort": Exit Sub
  End If
 End If

 '========================================================================

 Dim CurMesh&, CurVertex&, CurFace&, CurLight&, CurCamera&, UserPassword$

 Open TheFileName For Binary Access Write Lock Read Write As 1

  'APP/FILE PASSWORD:
  Put 1, , SimplyCrypt(AppFilePassword)

  'USER/FILE PASSWORD:
  If (MsgBox("Secure with a password ?", (vbQuestion + vbYesNo), "Password") = vbYes) Then
   Put 1, , CByte(1)
   Do: UserPassword = InputBox("Type a password :", "User password")
   Loop Until (Trim(UserPassword) <> vbNullString)
   UserPassword = SimplyCrypt(UserPassword)
   Put 1, , CInt(Len(UserPassword)): Put 1, , UserPassword
  Else
   Put 1, , CByte(2)
  End If

  FRM_Progress.Show
  FRM_Progress.DisplaySaveProgress 1, 0, 0

  'ALLOCATIONS:
  Put 1, , MaxVertices: Put 1, , MaxFaces
  Put 1, , MaxMeshs: Put 1, , MaxLights: Put 1, , MaxCameras

  'OUTPUT SIZE:
  Put 1, , OutputWidth: Put 1, , OutputHeight

  'PATH-TRACER:
  Put 1, , ViewPathsPerPixel: Put 1, , SamplesPerViewPath

  'SHADOWS:
  If (EnableAreaShadows = False) Then
   Put 1, , CByte(1)
  Else
   Put 1, , CByte(2): Put 1, , ShadowRaysCount: Put 1, , ShadowsApproxRadius
  End If

  'PHOTON MAPPING:
  Put 1, , TheAmbiantLight
  If (EnablePhotonMapping = False) Then
   Put 1, , CByte(1)
  Else
   Put 1, , CByte(2): Put 1, , MaximumAllocatedPhotons: Put 1, , SamplesPerPhotonPath
   Put 1, , PhotonsSearchRadius: Put 1, , BleedingDistance: Put 1, , EstimateMultiplier
  End If

  'FOG:
  If (FogEnable = False) Then
   Put 1, , CByte(1)
  Else
   Put 1, , CByte(2): Put 1, , FogRange: Put 1, , FogColor
   If (FogMode = K3DE_FM_LINEAR) Then
    Put 1, , CByte(1)
   ElseIf (FogMode = K3DE_FM_EXP) Then
    Put 1, , CByte(2): Put 1, , FogExpFactor1: Put 1, , FogExpFactor2
   End If
  End If

  'FILTERING:
  Select Case TheTexturesFilter
   Case K3DE_TFM_NEAREST:
    Put 1, , CByte(1)
   Case K3DE_TFM_NEAREST_MIP_NEAREST:
    Put 1, , CByte(2)
    Put 1, , MipMapsLevel: Put 1, , MipMapsMinPurcent
   Case K3DE_TFM_NEAREST_MIP_LINEAR:
    Put 1, , CByte(3)
    Put 1, , MipMapsLevel: Put 1, , MipMapsMinPurcent
   Case K3DE_TFM_FILTERED:
    Put 1, , CByte(4)
    Select Case TheTexelsFilter
     Case K3DE_XFM_BILINEAR:
      Put 1, , CByte(1)
     Case K3DE_XFM_BELL:
      Put 1, , CByte(2)
     Case K3DE_XFM_GAUSSIAN:
      Put 1, , CByte(3): Put 1, , KernelSize
     Case K3DE_XFM_CUBIC_SPLINE_B:
      Put 1, , CByte(4)
     Case K3DE_XFM_CUBIC_SPLINE_BC:
      Put 1, , CByte(5): Put 1, , CubicB: Put 1, , CubicC
     Case K3DE_XFM_CUBIC_SPLINE_CARDINAL:
      Put 1, , CByte(6): Put 1, , CubicA
    End Select
   Case K3DE_TFM_FILTERED_MIP_NEAREST:
    Put 1, , CByte(5)
    Select Case TheTexelsFilter
     Case K3DE_XFM_BILINEAR:
      Put 1, , CByte(1)
     Case K3DE_XFM_BELL:
      Put 1, , CByte(2)
     Case K3DE_XFM_GAUSSIAN:
      Put 1, , CByte(3): Put 1, , KernelSize
     Case K3DE_XFM_CUBIC_SPLINE_B:
      Put 1, , CByte(4)
     Case K3DE_XFM_CUBIC_SPLINE_BC:
      Put 1, , CByte(5): Put 1, , CubicB: Put 1, , CubicC
     Case K3DE_XFM_CUBIC_SPLINE_CARDINAL:
      Put 1, , CByte(6): Put 1, , CubicA
    End Select
    Put 1, , MipMapsLevel: Put 1, , MipMapsMinPurcent
   Case K3DE_TFM_FILTERED_MIP_LINEAR:
    Put 1, , CByte(6)
    Select Case TheTexelsFilter
     Case K3DE_XFM_BILINEAR:
      Put 1, , CByte(1)
     Case K3DE_XFM_BELL:
      Put 1, , CByte(2)
     Case K3DE_XFM_GAUSSIAN:
      Put 1, , CByte(3): Put 1, , KernelSize
     Case K3DE_XFM_CUBIC_SPLINE_B:
      Put 1, , CByte(4)
     Case K3DE_XFM_CUBIC_SPLINE_BC:
      Put 1, , CByte(5): Put 1, , CubicB: Put 1, , CubicC
     Case K3DE_XFM_CUBIC_SPLINE_CARDINAL:
      Put 1, , CByte(6): Put 1, , CubicA
    End Select
    Put 1, , MipMapsLevel: Put 1, , MipMapsMinPurcent
  End Select

  FRM_Progress.DisplaySaveProgress 1, (1 / 6), 1
  FRM_Progress.DisplaySaveProgress 2, (1 / 6), 0

  'BACKGROUND:
  If (UseBackGround = False) Then
   Put 1, , CByte(1)
  Else
   Put 1, , CByte(2)
   Put 1, , CInt(Len(OriginalBackGround.Label))
   Put 1, , OriginalBackGround.Label
   Put 1, , OriginalBackGround.BitsDepth
   Put 1, , OriginalBackGround.Dimensions
   Put 1, , OriginalBackGround.BackGroundColor
   Put 1, , OriginalBackGround.Datas()
  End If
  Put 1, , TheBackGroundColor

  FRM_Progress.DisplaySaveProgress 2, (2 / 6), 1
  FRM_Progress.DisplaySaveProgress 3, (2 / 6), 0

  'MESHS:
  Put 1, , TheMeshsCount
  If (TheMeshsCount = -1) Then GoTo Jump1:
  For CurMesh = 0 To TheMeshsCount

   'VERTICES & FACES COUNT:
   Put 1, , TheMeshs(CurMesh).Vertices.Length
   Put 1, , TheMeshs(CurMesh).Faces.Length

   'VERTICES:
   For CurVertex = TheMeshs(CurMesh).Vertices.Start To GetAddressLast(TheMeshs(CurMesh).Vertices)
    Put 1, , TheVertices(CurVertex).Position
   Next CurVertex

   'FACES:
   For CurFace = TheMeshs(CurMesh).Faces.Start To GetAddressLast(TheMeshs(CurMesh).Faces)
    If (TheFaces(CurFace).Visible = False) Then Put 1, , CByte(1) Else Put 1, , CByte(2)
    Put 1, , TheFaces(CurFace).A: Put 1, , TheFaces(CurFace).B: Put 1, , TheFaces(CurFace).C
    Put 1, , TheFaces(CurFace).AlphaVectors
    Put 1, , TheFaces(CurFace).ColorVectors
    Put 1, , TheFaces(CurFace).ReflectionVectors
    Put 1, , TheFaces(CurFace).RefractionVectors
    Put 1, , TheFaces(CurFace).RefractionNVectors
   Next CurFace

   'MESH'S PROPERTIES:
   Put 1, , CInt(Len(TheMeshs(CurMesh).Label))
   Put 1, , TheMeshs(CurMesh).Label
   Put 1, , TheMeshs(CurMesh).Position
   Put 1, , TheMeshs(CurMesh).Scales
   Put 1, , TheMeshs(CurMesh).Angles
   If (TheMeshs(CurMesh).MakeMatrix = False) Then Put 1, , CByte(1): Put 1, , TheMeshs(CurMesh).WorldMatrix Else Put 1, , CByte(2)
   If (TheMeshs(CurMesh).Visible = False) Then Put 1, , CByte(1) Else Put 1, , CByte(2)

   'MATERIAL:
   Put 1, , CInt(Len(TheMaterials(CurMesh).Label))
   Put 1, , TheMaterials(CurMesh).Label
   Put 1, , TheMaterials(CurMesh).Color
   Put 1, , TheMaterials(CurMesh).Reflection
   Put 1, , TheMaterials(CurMesh).Refraction
   Put 1, , TheMaterials(CurMesh).RefractionN
   Put 1, , TheMaterials(CurMesh).SpecularPowerK
   Put 1, , TheMaterials(CurMesh).SpecularPowerN

   'TEXTURES:
   If (TheMaterials(CurMesh).UseAlphaTexture = False) Then
    Put 1, , CByte(1)
   Else
    Put 1, , CByte(2)
    Put 1, , CInt(Len(TheAlphaTextures(TheMaterials(CurMesh).AlphaTextureID).Label))
    Put 1, , TheAlphaTextures(TheMaterials(CurMesh).AlphaTextureID).Label
    Put 1, , TheAlphaTextures(TheMaterials(CurMesh).AlphaTextureID).BitsDepth
    Put 1, , TheAlphaTextures(TheMaterials(CurMesh).AlphaTextureID).Dimensions
    Put 1, , TheAlphaTextures(TheMaterials(CurMesh).AlphaTextureID).BackGroundColor
    Put 1, , TheAlphaTextures(TheMaterials(CurMesh).AlphaTextureID).Datas()
   End If
   If (TheMaterials(CurMesh).UseColorTexture = False) Then
    Put 1, , CByte(1)
   Else
    Put 1, , CByte(2)
    Put 1, , CInt(Len(TheColorTextures(TheMaterials(CurMesh).ColorTextureID).Label))
    Put 1, , TheColorTextures(TheMaterials(CurMesh).ColorTextureID).Label
    Put 1, , TheColorTextures(TheMaterials(CurMesh).ColorTextureID).BitsDepth
    Put 1, , TheColorTextures(TheMaterials(CurMesh).ColorTextureID).Dimensions
    Put 1, , TheColorTextures(TheMaterials(CurMesh).ColorTextureID).BackGroundColor
    Put 1, , TheColorTextures(TheMaterials(CurMesh).ColorTextureID).Datas()
   End If
   If (TheMaterials(CurMesh).UseReflectionTexture = False) Then
    Put 1, , CByte(1)
   Else
    Put 1, , CByte(2)
    Put 1, , CInt(Len(TheReflectionTextures(TheMaterials(CurMesh).ReflectionTextureID).Label))
    Put 1, , TheReflectionTextures(TheMaterials(CurMesh).ReflectionTextureID).Label
    Put 1, , TheReflectionTextures(TheMaterials(CurMesh).ReflectionTextureID).BitsDepth
    Put 1, , TheReflectionTextures(TheMaterials(CurMesh).ReflectionTextureID).Dimensions
    Put 1, , TheReflectionTextures(TheMaterials(CurMesh).ReflectionTextureID).BackGroundColor
    Put 1, , TheReflectionTextures(TheMaterials(CurMesh).ReflectionTextureID).Datas()
   End If
   If (TheMaterials(CurMesh).UseRefractionTexture = False) Then
    Put 1, , CByte(1)
   Else
    Put 1, , CByte(2)
    Put 1, , CInt(Len(TheRefractionTextures(TheMaterials(CurMesh).RefractionTextureID).Label))
    Put 1, , TheRefractionTextures(TheMaterials(CurMesh).RefractionTextureID).Label
    Put 1, , TheRefractionTextures(TheMaterials(CurMesh).RefractionTextureID).BitsDepth
    Put 1, , TheRefractionTextures(TheMaterials(CurMesh).RefractionTextureID).Dimensions
    Put 1, , TheRefractionTextures(TheMaterials(CurMesh).RefractionTextureID).BackGroundColor
    Put 1, , TheRefractionTextures(TheMaterials(CurMesh).RefractionTextureID).Datas()
   End If
   If (TheMaterials(CurMesh).UseRefractionNTexture = False) Then
    Put 1, , CByte(1)
   Else
    Put 1, , CByte(2)
    Put 1, , CInt(Len(TheRefractionNTextures(TheMaterials(CurMesh).RefractionNTextureID).Label))
    Put 1, , TheRefractionNTextures(TheMaterials(CurMesh).RefractionNTextureID).Label
    Put 1, , TheRefractionNTextures(TheMaterials(CurMesh).RefractionNTextureID).BitsDepth
    Put 1, , TheRefractionNTextures(TheMaterials(CurMesh).RefractionNTextureID).Dimensions
    Put 1, , TheRefractionNTextures(TheMaterials(CurMesh).RefractionNTextureID).BackGroundColor
    Put 1, , TheRefractionNTextures(TheMaterials(CurMesh).RefractionNTextureID).Datas()
   End If
   If ((TheMeshsCount > 0) And (CurMesh <> TheMeshsCount)) Then
    FRM_Progress.DisplaySaveProgress 3, (2 / 6), CSng(CurMesh / TheMeshsCount)
   End If
  Next CurMesh

  FRM_Progress.DisplaySaveProgress 3, (3 / 6), 1
  FRM_Progress.DisplaySaveProgress 4, (3 / 6), 0

Jump1:

  'OMNI LIGHTS:
  Put 1, , TheSphereLightsCount
  If (TheSphereLightsCount = -1) Then GoTo Jump2:
  For CurLight = 0 To TheSphereLightsCount
   Put 1, , CInt(Len(TheSphereLights(CurLight).Label))
   Put 1, , TheSphereLights(CurLight).Label
   Put 1, , TheSphereLights(CurLight).Color
   Put 1, , TheSphereLights(CurLight).Position
   Put 1, , TheSphereLights(CurLight).Range
   If (TheSphereLights(CurLight).Enable = False) Then Put 1, , CByte(1) Else Put 1, , CByte(2)
   If ((TheSphereLightsCount > 0) And (CurLight <> TheSphereLightsCount)) Then
    FRM_Progress.DisplaySaveProgress 4, (3 / 6), CSng(CurLight / TheSphereLightsCount)
   End If
  Next CurLight

  FRM_Progress.DisplaySaveProgress 4, (4 / 6), 1
  FRM_Progress.DisplaySaveProgress 5, (4 / 6), 0

Jump2:

  'SPOT LIGHTS:
  Put 1, , TheConeLightsCount
  If (TheConeLightsCount = -1) Then GoTo Jump3:
  For CurLight = 0 To TheConeLightsCount
   Put 1, , CInt(Len(TheConeLights(CurLight).Label))
   Put 1, , TheConeLights(CurLight).Label
   Put 1, , TheConeLights(CurLight).Color
   Put 1, , TheConeLights(CurLight).Position
   Put 1, , TheConeLights(CurLight).Direction
   Put 1, , TheConeLights(CurLight).Falloff
   Put 1, , TheConeLights(CurLight).Hotspot
   Put 1, , TheConeLights(CurLight).Range
   If (TheConeLights(CurLight).Enable = False) Then Put 1, , CByte(1) Else Put 1, , CByte(2)
   If ((TheConeLightsCount > 0) And (CurLight <> TheConeLightsCount)) Then
    FRM_Progress.DisplaySaveProgress 5, (4 / 6), CSng(CurLight / TheConeLightsCount)
   End If
  Next CurLight

  FRM_Progress.DisplaySaveProgress 5, (5 / 6), 1
  FRM_Progress.DisplaySaveProgress 6, (5 / 6), 0

Jump3:

  'CAMERAS:
  Put 1, , TheCamerasCount
  For CurCamera = 0 To TheCamerasCount
   Put 1, , CInt(Len(TheCameras(CurCamera).Label))
   Put 1, , TheCameras(CurCamera).Label
   Put 1, , TheCameras(CurCamera).Position
   Put 1, , TheCameras(CurCamera).Direction
   Put 1, , TheCameras(CurCamera).RollAngle
   Put 1, , TheCameras(CurCamera).FOVAngle
   Put 1, , TheCameras(CurCamera).ClearDistance
   Put 1, , TheCameras(CurCamera).Dispersion
   If (TheCameras(CurCamera).BackFaceCulling = False) Then Put 1, , CByte(1) Else Put 1, , CByte(1)
   If (TheCameras(CurCamera).MakeMatrix = False) Then Put 1, , CByte(1): Put 1, , TheCameras(CurCamera).ViewMatrix Else Put 1, , CByte(2)
   If ((TheCamerasCount > 0) And (CurCamera <> TheCamerasCount)) Then
    FRM_Progress.DisplaySaveProgress 6, (5 / 6), CSng(CurCamera / TheCamerasCount)
   End If
  Next CurCamera

  'VIEWPORTS:
  Put 1, , TheCurrentCamera
  Put 1, , DisplayMode
  If (CheckOut(FRM_Main.Check1) = False) Then Put 1, , CByte(1) Else Put 1, , CByte(2)
  If (CheckOut(FRM_Main.Check2) = False) Then Put 1, , CByte(1) Else Put 1, , CByte(2)
  If (CheckOut(FRM_Main.Check3) = False) Then Put 1, , CByte(1) Else Put 1, , CByte(2)
  If (CheckOut(FRM_Main.Check4) = False) Then Put 1, , CByte(1) Else Put 1, , CByte(2)
  If (CheckOut(FRM_Main.Check5) = False) Then Put 1, , CByte(1) Else Put 1, , CByte(2)
  If (CheckOut(FRM_Main.Check6) = False) Then Put 1, , CByte(1) Else Put 1, , CByte(2)

  FRM_Progress.DisplaySaveProgress 6, 1, 1

 Close 1

 Unload FRM_Progress
 MsgBox "File saved.", vbInformation, "Save scene"

End Sub
Sub Engine_SaveMesh(TheMeshIndex&, TheFileName As String)

 If (Started = False) Then Exit Sub

 'Save the mesh to the destination file.

 If (Right(TheFileName, 4) <> ObjectFileExtension) Then
  MsgBox "Invalid file !", vbCritical, "Save object": Exit Sub
 End If

 If (FileExist(TheFileName) = True) Then
  If (MsgBox("File already exist, overwrite ?", (vbQuestion + vbYesNo), "Save object") = vbYes) Then
   Kill TheFileName
  Else
   MsgBox "Aborted saving operation.", vbCritical, "Abort": Exit Sub
  End If
 End If

 Dim CurVertex&, CurFace&, UserPassword$

 Open TheFileName For Binary Access Write Lock Read Write As 1

  'APP/FILE PASSWORD:
  Put 1, , SimplyCrypt(AppFilePassword)

  'USER/FILE PASSWORD:
  If (MsgBox("Secure with a password ?", (vbQuestion + vbYesNo), "Password") = vbYes) Then
   Put 1, , CByte(1)
   Do: UserPassword = InputBox("Type a password :", "User password")
   Loop Until (Trim(UserPassword) <> vbNullString)
   UserPassword = SimplyCrypt(UserPassword)
   Put 1, , CInt(Len(UserPassword)): Put 1, , UserPassword
  Else
   Put 1, , CByte(2)
  End If

  'VERTICES & FACES COUNT:
  Put 1, , TheMeshs(TheMeshIndex).Vertices.Length
  Put 1, , TheMeshs(TheMeshIndex).Faces.Length

  'VERTICES:
  For CurVertex = TheMeshs(TheMeshIndex).Vertices.Start To GetAddressLast(TheMeshs(TheMeshIndex).Vertices)
   Put 1, , TheVertices(CurVertex).Position
  Next CurVertex

  'FACES:
  For CurFace = TheMeshs(TheMeshIndex).Faces.Start To GetAddressLast(TheMeshs(TheMeshIndex).Faces)
   If (TheFaces(CurFace).Visible = False) Then Put 1, , CByte(1) Else Put 1, , CByte(2)
   Put 1, , CLng(TheFaces(CurFace).A - (TheMeshs(TheMeshIndex).Vertices.Start - 1))
   Put 1, , CLng(TheFaces(CurFace).B - (TheMeshs(TheMeshIndex).Vertices.Start - 1))
   Put 1, , CLng(TheFaces(CurFace).C - (TheMeshs(TheMeshIndex).Vertices.Start - 1))
   Put 1, , TheFaces(CurFace).AlphaVectors
   Put 1, , TheFaces(CurFace).ColorVectors
   Put 1, , TheFaces(CurFace).ReflectionVectors
   Put 1, , TheFaces(CurFace).RefractionVectors
   Put 1, , TheFaces(CurFace).RefractionNVectors
  Next CurFace

  'MESH'S PROPERTIES:
  Put 1, , CInt(Len(TheMeshs(TheMeshIndex).Label))
  Put 1, , TheMeshs(TheMeshIndex).Label
  Put 1, , TheMeshs(TheMeshIndex).Position
  Put 1, , TheMeshs(TheMeshIndex).Scales
  Put 1, , TheMeshs(TheMeshIndex).Angles
  If (TheMeshs(TheMeshIndex).MakeMatrix = False) Then Put 1, , CByte(1): Put 1, , TheMeshs(TheMeshIndex).WorldMatrix Else Put 1, , CByte(2)
  If (TheMeshs(TheMeshIndex).Visible = False) Then Put 1, , CByte(1) Else Put 1, , CByte(2)

  'MATERIAL:
  Put 1, , CInt(Len(TheMaterials(TheMeshIndex).Label))
  Put 1, , TheMaterials(TheMeshIndex).Label
  Put 1, , TheMaterials(TheMeshIndex).Color
  Put 1, , TheMaterials(TheMeshIndex).Reflection
  Put 1, , TheMaterials(TheMeshIndex).Refraction
  Put 1, , TheMaterials(TheMeshIndex).RefractionN
  Put 1, , TheMaterials(TheMeshIndex).SpecularPowerK
  Put 1, , TheMaterials(TheMeshIndex).SpecularPowerN

  'TEXTURES:
  If (TheMaterials(TheMeshIndex).UseAlphaTexture = False) Then
   Put 1, , CByte(1)
  Else
   Put 1, , CByte(2)
   Put 1, , CInt(Len(TheAlphaTextures(TheMaterials(TheMeshIndex).AlphaTextureID).Label))
   Put 1, , TheAlphaTextures(TheMaterials(TheMeshIndex).AlphaTextureID).Label
   Put 1, , TheAlphaTextures(TheMaterials(TheMeshIndex).AlphaTextureID).BitsDepth
   Put 1, , TheAlphaTextures(TheMaterials(TheMeshIndex).AlphaTextureID).Dimensions
   Put 1, , TheAlphaTextures(TheMaterials(TheMeshIndex).AlphaTextureID).BackGroundColor
   Put 1, , TheAlphaTextures(TheMaterials(TheMeshIndex).AlphaTextureID).Datas()
  End If

  If (TheMaterials(TheMeshIndex).UseColorTexture = False) Then
   Put 1, , CByte(1)
  Else
   Put 1, , CByte(2)
   Put 1, , CInt(Len(TheColorTextures(TheMaterials(TheMeshIndex).ColorTextureID).Label))
   Put 1, , TheColorTextures(TheMaterials(TheMeshIndex).ColorTextureID).Label
   Put 1, , TheColorTextures(TheMaterials(TheMeshIndex).ColorTextureID).BitsDepth
   Put 1, , TheColorTextures(TheMaterials(TheMeshIndex).ColorTextureID).Dimensions
   Put 1, , TheColorTextures(TheMaterials(TheMeshIndex).ColorTextureID).BackGroundColor
   Put 1, , TheColorTextures(TheMaterials(TheMeshIndex).ColorTextureID).Datas()
  End If

  If (TheMaterials(TheMeshIndex).UseReflectionTexture = False) Then
   Put 1, , CByte(1)
  Else
   Put 1, , CByte(2)
   Put 1, , CInt(Len(TheReflectionTextures(TheMaterials(TheMeshIndex).ReflectionTextureID).Label))
   Put 1, , TheReflectionTextures(TheMaterials(TheMeshIndex).ReflectionTextureID).Label
   Put 1, , TheReflectionTextures(TheMaterials(TheMeshIndex).ReflectionTextureID).BitsDepth
   Put 1, , TheReflectionTextures(TheMaterials(TheMeshIndex).ReflectionTextureID).Dimensions
   Put 1, , TheReflectionTextures(TheMaterials(TheMeshIndex).ReflectionTextureID).BackGroundColor
   Put 1, , TheReflectionTextures(TheMaterials(TheMeshIndex).ReflectionTextureID).Datas()
  End If

  If (TheMaterials(TheMeshIndex).UseRefractionTexture = False) Then
   Put 1, , CByte(1)
  Else
   Put 1, , CByte(2)
   Put 1, , CInt(Len(TheRefractionTextures(TheMaterials(TheMeshIndex).RefractionTextureID).Label))
   Put 1, , TheRefractionTextures(TheMaterials(TheMeshIndex).RefractionTextureID).Label
   Put 1, , TheRefractionTextures(TheMaterials(TheMeshIndex).RefractionTextureID).BitsDepth
   Put 1, , TheRefractionTextures(TheMaterials(TheMeshIndex).RefractionTextureID).Dimensions
   Put 1, , TheRefractionTextures(TheMaterials(TheMeshIndex).RefractionTextureID).BackGroundColor
   Put 1, , TheRefractionTextures(TheMaterials(TheMeshIndex).RefractionTextureID).Datas()
  End If

  If (TheMaterials(TheMeshIndex).UseRefractionNTexture = False) Then
   Put 1, , CByte(1)
  Else
   Put 1, , CByte(2)
   Put 1, , CInt(Len(TheRefractionNTextures(TheMaterials(TheMeshIndex).RefractionNTextureID).Label))
   Put 1, , TheRefractionNTextures(TheMaterials(TheMeshIndex).RefractionNTextureID).Label
   Put 1, , TheRefractionNTextures(TheMaterials(TheMeshIndex).RefractionNTextureID).BitsDepth
   Put 1, , TheRefractionNTextures(TheMaterials(TheMeshIndex).RefractionNTextureID).Dimensions
   Put 1, , TheRefractionNTextures(TheMaterials(TheMeshIndex).RefractionNTextureID).BackGroundColor
   Put 1, , TheRefractionNTextures(TheMaterials(TheMeshIndex).RefractionNTextureID).Datas()
  End If

 Close 1

 MsgBox "File saved.", vbInformation, "Save object"

End Sub
Sub ReCreateMipMaps()

 'Recreate the mip maps, this function is called when changing the mip-mapping level.

 If ((Started = False) Or (TheMeshsCount = -1)) Then Exit Sub

 Dim CurMesh&, CurMip&, MinW%, MinH%, StpW%, StpH%, NewW%, NewH%

 For CurMesh = 0 To TheMeshsCount
  If (TheMaterials(CurMesh).UseColorTexture = True) Then
   If (BitMap2D_IsValid(TheColorTextures(CurMesh)) = True) Then
    ReDim TheColorMips(CurMesh).MipSequance(MipMapsLevel)
    MinW = (TheColorTextures(CurMesh).Dimensions.X * (MipMapsMinPurcent / 100))
    MinH = (TheColorTextures(CurMesh).Dimensions.Y * (MipMapsMinPurcent / 100))
    If (MinW < MinBitMapWidth) Then MinW = MinBitMapWidth
    If (MinH < MinBitMapHeight) Then MinW = MinBitMapHeight
    StpW = ((TheColorTextures(CurMesh).Dimensions.X - MinW) / (MipMapsLevel + 1))
    StpH = ((TheColorTextures(CurMesh).Dimensions.Y - MinH) / (MipMapsLevel + 1))
    NewW = MinW: NewH = MinH
    For CurMip = 0 To MipMapsLevel
     TheColorMips(CurMesh).MipSequance(CurMip) = TheColorTextures(CurMesh)
     TheColorMips(CurMesh).MipSequance(CurMip).Label = TheMeshs(CurMesh).Label & "_MIP" & CurMip
     BitMap2D_Resample TheColorMips(CurMesh).MipSequance(CurMip), NewW, NewH, TheTexelsFilter
     NewW = (NewW + StpW): NewH = (NewH + StpH)
    Next CurMip
   End If
  End If
 Next CurMesh

End Sub
Sub SetDefaultParameters()

 UseBackGround = False
 TheBackGroundColor = ColorLongToRGB(16374999)
 TheAmbiantLight = ColorInput(0, 0, 0)
 TheTexturesFilter = K3DE_TFM_NEAREST
 TheTexelsFilter = K3DE_XFM_BILINEAR
 OutputWidth = 640
 OutputHeight = 480
 MipRange = 750: InvMipRange = (1 / MipRange)

 ViewPathsPerPixel = 1
 SamplesPerViewPath = 10         'Bounces per view path

 EnableAreaShadows = False
 ShadowRaysCount = 20
 ShadowsApproxRadius = 5

 CubicA = -0.5
 CubicB = 0.5
 CubicC = 0.5
 KernelSize = 2
 MipMapsLevel = 3
 MipMapsMinPurcent = 30          'The last mip is 30% size of original texture

 EnablePhotonMapping = False
 ReDim ThePhotonMap(0)
 MaximumAllocatedPhotons = 0
 SamplesPerPhotonPath = 10       'Bounces count (photon path)
 PhotonsSearchRadius = 30
 BleedingDistance = 250
 EstimateMultiplier = 50

 FogEnable = False
 FogRange = 1000
 FogColor = ColorWhite
 FogExpFactor1 = 0.1: FogExpFactor2 = 0.2
 FogMode = K3DE_FM_EXP

 Wire_DefaultPerspectiveDistorsion = 400
 Wire_PerspectiveDistorsion! = Wire_DefaultPerspectiveDistorsion
 Wire_AddedDepth! = 450
 Wire_ScaleNormalTo! = 20
 Wire_CameraTo! = 100
 Wire_PhotonTo! = 5
 Wire_ParallalScale! = 1

End Sub
Function Shader_GetFog(TheDistance!) As Single

 'Calculate the Fog factor

 If (TheDistance > FogRange) Then Exit Function

 Dim InvFog!: InvFog = (TheDistance / FogRange)

 Select Case FogMode
  Case K3DE_FM_LINEAR: Shader_GetFog = (1 - InvFog) 'Linear fog
  Case K3DE_FM_EXP:    Shader_GetFog = ExpScale((1 - InvFog), FogExpFactor1, FogExpFactor2) 'Exponential fog
 End Select

End Function
Function Shader_GetShadow(TheLightOrigin As Vector3D, TheIntersection As Intersection3D, TheIntersectionPoint As Vector3D) As Single

 Dim TmpShadowRay As Ray3D, ShadowResult As Intersection3D, CurShadowRay&, ShadowRaysFound&

 TmpShadowRay.Direction = TheIntersectionPoint
 If (EnableAreaShadows = False) Then 'Use 1 shadow-ray
  'Trace a ray from the light's Position to the intersection point
  TmpShadowRay.Position = TheLightOrigin: ShadowResult = TraceRayFirst(TmpShadowRay, True, True)
  If ((ShadowResult.MeshNumber = TheIntersection.MeshNumber) And (ShadowResult.FaceNumber = TheIntersection.FaceNumber)) Then
   Shader_GetShadow = 1
  End If
 Else
  'Approximate the penumbra regions by sampling with shadow-rays (area).
  For CurShadowRay = 0 To (ShadowRaysCount - 1)
   TmpShadowRay.Position = VectorAdd(TheLightOrigin, TheShadowRays(CurShadowRay))
   ShadowResult = TraceRayFirst(TmpShadowRay, True, True)
   If ((ShadowResult.MeshNumber = TheIntersection.MeshNumber) And (ShadowResult.FaceNumber = TheIntersection.FaceNumber)) Then
    ShadowRaysFound = (ShadowRaysFound + 1)
   End If
  Next CurShadowRay
  If (ShadowRaysFound <> 0) Then Shader_GetShadow = (ShadowRaysFound / ShadowRaysCount)
 End If

End Function
Function Shader_GetAttenuation(VecFrom As Vector3D, Range!, VecInput As Vector3D) As Single

 'Use this for a simple radial (spherical) linear attenuation.

 Dim Distance!: Distance = VectorDistance(VecFrom, VecInput)
 If (Distance < Range) Then Shader_GetAttenuation = (1 - (Distance / Range))

End Function
Function Shader_GetShapeCone(VecOrigin As Vector3D, VecDirection As Vector3D, FallOffAngle!, HotSpotAngle!, Range!, VecInput As Vector3D) As Single

 'Return the spot-light filter factor

 Shader_GetShapeCone = VectorAngle(VectorSubtract(VecDirection, VecOrigin), VectorSubtract(VecInput, VecOrigin))

 If (Shader_GetShapeCone < 0) Then
  Shader_GetShapeCone = 0
 Else
  Dim FallOffCos As Single, HotSpotCos As Single
  FallOffCos = Cos(FallOffAngle): HotSpotCos = Cos(HotSpotAngle)
  'Angular attenuation:
  Shader_GetShapeCone = ((Shader_GetShapeCone - HotSpotCos) / (FallOffCos - HotSpotCos))
  If (Shader_GetShapeCone < 0) Then Shader_GetShapeCone = 0
  If (Shader_GetShapeCone > 1) Then Shader_GetShapeCone = 1
  Shader_GetShapeCone = (1 - Shader_GetShapeCone)
  'Distance attenuation:
  Shader_GetShapeCone = (Shader_GetShapeCone * Shader_GetAttenuation(VecOrigin, Range, VecInput))
 End If

End Function
Function Shader_GetLambertAngle(VecView As Vector3D, VecLight As Vector3D, VecNormal As Vector3D, VecInput As Vector3D) As Single

 'Calculate the lambert angle, A condition that the light and eye
 'points must are in the same direction from the plane, the lambert angle
 'is so returned, 0 othewise.

 Dim ViewAngle!, LightAngle!

 LightAngle = VectorAngle(VectorSubtract(VecLight, VecInput), VecNormal)
 ViewAngle = VectorAngle(VectorSubtract(VecView, VecInput), VecNormal)

 If ((LightAngle > 0) And (ViewAngle > 0)) Then
  Shader_GetLambertAngle = LightAngle
 ElseIf ((LightAngle < 0) And (ViewAngle < 0)) Then
  Shader_GetLambertAngle = -LightAngle
 End If

End Function
Function Shader_GetSpecularity(VecView As Vector3D, VecLight As Vector3D, VecNormal As Vector3D, VecInput As Vector3D) As Single

 'Return the specularity-contribution factor

 Dim Reflection As Vector3D, TheNormal As Vector3D
 Dim ViewDir As Vector3D, ViewAngle!, LightDir As Vector3D, LightAngle!

 ViewDir = VectorNormalize(VectorSubtract(VecView, VecInput))
 LightDir = VectorNormalize(VectorSubtract(VecLight, VecInput))
 TheNormal = VectorNormalize(VecNormal)
 LightAngle = VectorDotProduct(LightDir, TheNormal)
 Reflection = VectorSubtract(VectorScale(TheNormal, (2 * LightAngle)), LightDir)
 ViewAngle = VectorDotProduct(Reflection, ViewDir)

 If (ViewAngle > 0) Then Shader_GetSpecularity = ViewAngle

End Function
Function GetIntersectionPoint(TheIntersection As Intersection3D) As Vector3D

 If (Started = False) Then Exit Function

 'Recieve the intersected point on the triangle (just an interpolation by the barycentrics coordinates).

 GetIntersectionPoint.X = ((TheIntersection.U * TheVertices(TheFaces(TheIntersection.FaceNumber).A).TmpPos.X) + _
                           (TheIntersection.V * TheVertices(TheFaces(TheIntersection.FaceNumber).B).TmpPos.X) + _
                           (TheIntersection.W * TheVertices(TheFaces(TheIntersection.FaceNumber).C).TmpPos.X))

 GetIntersectionPoint.Y = ((TheIntersection.U * TheVertices(TheFaces(TheIntersection.FaceNumber).A).TmpPos.Y) + _
                           (TheIntersection.V * TheVertices(TheFaces(TheIntersection.FaceNumber).B).TmpPos.Y) + _
                           (TheIntersection.W * TheVertices(TheFaces(TheIntersection.FaceNumber).C).TmpPos.Y))

 GetIntersectionPoint.Z = ((TheIntersection.U * TheVertices(TheFaces(TheIntersection.FaceNumber).A).TmpPos.Z) + _
                           (TheIntersection.V * TheVertices(TheFaces(TheIntersection.FaceNumber).B).TmpPos.Z) + _
                           (TheIntersection.W * TheVertices(TheFaces(TheIntersection.FaceNumber).C).TmpPos.Z))

End Function
Function GetMaterialColor(TheIntersection As Intersection3D) As ColorRGB

 If (Started = False) Then Exit Function

 'Calculate the Color:

 Dim CurU!, CurV!, LinearZ!

 If (TheMaterials(TheIntersection.MeshNumber).UseColorTexture = False) Then
  GetMaterialColor = TheMaterials(TheIntersection.MeshNumber).Color
 Else
  CurU = ((TheIntersection.U * TheFaces(TheIntersection.FaceNumber).ColorVectors.U1) + _
          (TheIntersection.V * TheFaces(TheIntersection.FaceNumber).ColorVectors.U2) + _
          (TheIntersection.W * TheFaces(TheIntersection.FaceNumber).ColorVectors.U3))
  CurV = ((TheIntersection.U * TheFaces(TheIntersection.FaceNumber).ColorVectors.V1) + _
          (TheIntersection.V * TheFaces(TheIntersection.FaceNumber).ColorVectors.V2) + _
          (TheIntersection.W * TheFaces(TheIntersection.FaceNumber).ColorVectors.V3))
  LinearZ = (((TheIntersection.U * TheVertices(TheFaces(TheIntersection.FaceNumber).A).TmpPos.Z) + _
              (TheIntersection.V * TheVertices(TheFaces(TheIntersection.FaceNumber).B).TmpPos.Z) + _
              (TheIntersection.W * TheVertices(TheFaces(TheIntersection.FaceNumber).C).TmpPos.Z)) * InvMipRange)
  If (LinearZ > 1) Then LinearZ = 1
  GetMaterialColor = DoMipFiltering(TheColorTextures(TheIntersection.MeshNumber), TheColorMips(TheIntersection.MeshNumber), CurU, CurV, LinearZ)
 End If

End Function
Function GetMaterialReflectivity(TheIntersection As Intersection3D) As Byte

 If (Started = False) Then Exit Function

 'Calculate Reflection % (just an interpolation by the barycentrics coordinates):

 Dim CurU!, CurV!

 If (TheMaterials(TheIntersection.MeshNumber).UseReflectionTexture = False) Then
  GetMaterialReflectivity = TheMaterials(TheIntersection.MeshNumber).Reflection
 Else
  CurU = ((TheIntersection.U * TheFaces(TheIntersection.FaceNumber).ReflectionVectors.U1) + _
          (TheIntersection.V * TheFaces(TheIntersection.FaceNumber).ReflectionVectors.U2) + _
          (TheIntersection.W * TheFaces(TheIntersection.FaceNumber).ReflectionVectors.U3))
  CurV = ((TheIntersection.U * TheFaces(TheIntersection.FaceNumber).ReflectionVectors.V1) + _
          (TheIntersection.V * TheFaces(TheIntersection.FaceNumber).ReflectionVectors.V2) + _
          (TheIntersection.W * TheFaces(TheIntersection.FaceNumber).ReflectionVectors.V3))
  GetMaterialReflectivity = DoTexelFiltering8(TheTexelsFilter, TheReflectionTextures(TheIntersection.MeshNumber), CurU, CurV, True)
 End If

End Function
Function GetMaterialRefractionN(TheIntersection As Intersection3D) As Single

 If (Started = False) Then Exit Function

 'Calculate refraction index N (just an interpolation by the barycentrics coordinates) :

 Dim CurU!, CurV!

 If (TheMaterials(TheIntersection.MeshNumber).UseRefractionNTexture = False) Then
  GetMaterialRefractionN = TheMaterials(TheIntersection.MeshNumber).RefractionN
 Else
  CurU = ((TheIntersection.U * TheFaces(TheIntersection.FaceNumber).RefractionNVectors.U1) + _
          (TheIntersection.V * TheFaces(TheIntersection.FaceNumber).RefractionNVectors.U2) + _
          (TheIntersection.W * TheFaces(TheIntersection.FaceNumber).RefractionNVectors.U3))
  CurV = ((TheIntersection.U * TheFaces(TheIntersection.FaceNumber).RefractionNVectors.V1) + _
          (TheIntersection.V * TheFaces(TheIntersection.FaceNumber).RefractionNVectors.V2) + _
          (TheIntersection.W * TheFaces(TheIntersection.FaceNumber).RefractionNVectors.V3))
  GetMaterialRefractionN = DoTexelFiltering8(TheTexelsFilter, TheRefractionNTextures(TheIntersection.MeshNumber), CurU, CurV, True)
  GetMaterialRefractionN = (GetMaterialRefractionN * AlphaFactor * MaximumRefractionNFactor)
 End If

End Function
Function GetMaterialRefractivity(TheIntersection As Intersection3D) As Byte

 If (Started = False) Then Exit Function

 'Calculate Refraction % (just an interpolation by the barycentrics coordinates):

 Dim CurU!, CurV!

 If (TheMaterials(TheIntersection.MeshNumber).UseRefractionTexture = False) Then
  GetMaterialRefractivity = TheMaterials(TheIntersection.MeshNumber).Refraction
 Else
  CurU = ((TheIntersection.U * TheFaces(TheIntersection.FaceNumber).RefractionVectors.U1) + _
          (TheIntersection.V * TheFaces(TheIntersection.FaceNumber).RefractionVectors.U2) + _
          (TheIntersection.W * TheFaces(TheIntersection.FaceNumber).RefractionVectors.U3))
  CurV = ((TheIntersection.U * TheFaces(TheIntersection.FaceNumber).RefractionVectors.V1) + _
          (TheIntersection.V * TheFaces(TheIntersection.FaceNumber).RefractionVectors.V2) + _
          (TheIntersection.W * TheFaces(TheIntersection.FaceNumber).RefractionVectors.V3))
  GetMaterialRefractivity = DoTexelFiltering8(TheTexelsFilter, TheRefractionTextures(TheIntersection.MeshNumber), CurU, CurV, True)
 End If

End Function
Function GetMaterialVisibility(TheIntersection As Intersection3D) As Boolean

 If (Started = False) Then Exit Function

 'Calculate the visibility information in the intersection point by
 'reading the alpha map (just an interpolation by the barycentrics coordinates) :

 GetMaterialVisibility = True 'Default

 If (TheMaterials(TheIntersection.MeshNumber).UseAlphaTexture = True) Then

  Dim CurU!, CurV!

  CurU = ((TheIntersection.U * TheFaces(TheIntersection.FaceNumber).AlphaVectors.U1) + _
          (TheIntersection.V * TheFaces(TheIntersection.FaceNumber).AlphaVectors.U2) + _
          (TheIntersection.W * TheFaces(TheIntersection.FaceNumber).AlphaVectors.U3))
  CurV = ((TheIntersection.U * TheFaces(TheIntersection.FaceNumber).AlphaVectors.V1) + _
          (TheIntersection.V * TheFaces(TheIntersection.FaceNumber).AlphaVectors.V2) + _
          (TheIntersection.W * TheFaces(TheIntersection.FaceNumber).AlphaVectors.V3))

  If (DoTexelFiltering8(TheTexelsFilter, TheAlphaTextures(TheIntersection.MeshNumber), CurU, CurV, True) = 0) Then
   GetMaterialVisibility = False
  End If

 End If

End Function
Function GetMaterialInfos(TheIntersection As Intersection3D) As MaterialInfos

 If (Started = False) Then Exit Function

 'Gives the material informations, from the intersection point, the values
 'are computed form the textures in the case of a texture-mapped surface.

 With GetMaterialInfos
  .Color = GetMaterialColor(TheIntersection)
  .Reflection = GetMaterialReflectivity(TheIntersection)
  .Refraction = GetMaterialRefractivity(TheIntersection)
  .RefractionN = GetMaterialRefractionN(TheIntersection)
  .SpecularPowerK = TheMaterials(TheIntersection.MeshNumber).SpecularPowerK
  .SpecularPowerN = TheMaterials(TheIntersection.MeshNumber).SpecularPowerN
 End With

End Function
Function TraceViewPath(TheStartRay As Ray3D) As ColorRGB

 'FUTURE ADDITION: Hyrarchical tracing (recursive) instead of Russian Roulette

 If (Started = False) Then Exit Function

 Dim CurSample%, CurrentPoint As Vector3D, TheNewRay As Ray3D
 Dim TotalWeight%, RussianRoulette%, DirectionVec As Vector3D
 Dim CurIntersection As Intersection3D, CurMatInfos As MaterialInfos
 Dim ResultColors() As ColorRGB, ResultDiffusions() As Single
 Dim ResultColor As ColorRGB, DiffusionRate%, CurrentResult%, FogFactor!
 Dim NormalVector As Vector3D, AnyOne As Boolean, TmpRay As Ray3D

 ReDim ResultColors(0): ReDim ResultDiffusions(0)

 TheNewRay = TheStartRay
 For CurSample = 1 To SamplesPerViewPath
  CurIntersection = TraceRayFirstBackFacesCull(TheNewRay, True, True)
  If (CurIntersection.MeshNumber = -1) Then
   If (CurSample = 1) Then TraceViewPath = ColorInput(-1, -1, -1): Exit Function
  Else
   CurMatInfos = GetMaterialInfos(CurIntersection)
   CurrentPoint = GetIntersectionPoint(CurIntersection)
   NormalVector = TheFaces(CurIntersection.FaceNumber).Normal
   If (VectorAngle(VectorInverse(CurrentPoint), NormalVector) < 0) Then
    NormalVector = VectorInverse(NormalVector) 'Reorient the normal vector
   End If
   TotalWeight = (CInt(CurMatInfos.Reflection) + CInt(CurMatInfos.Refraction))
   DiffusionRate = (500 - TotalWeight)
   If (DiffusionRate > 0) Then
    AnyOne = True

    '1- Calculate the direct-light contribution on this point:
    ResultColors(UBound(ResultColors())) = CalculateDirectLightContribution(CurIntersection, TheNewRay.Position, CurrentPoint, NormalVector, CurMatInfos)

    '2- Add the bounced radiance (diffuse interreflexions and caustics) from the photonmap:
    If ((EnablePhotonMapping = True) And (EstimateFromPhotonmap = True)) Then
     TmpRay.Position = CurrentPoint: TmpRay.Direction = NormalVector
     'Add the indirect lighting:
     ResultColors(UBound(ResultColors())) = ColorLimit(ColorAdd(ResultColors(UBound(ResultColors())), EstimateRadianceFromPhotonmap(TheNewRay.Position, TmpRay, CurMatInfos)))
    Else
     'Add the flat ambiant light (approximative value):
     ResultColors(UBound(ResultColors())) = ColorLimit(ColorAdd(ResultColors(UBound(ResultColors())), TheAmbiantLight))
    End If

    '3- Apply the fogging effect:
    If (FogEnable = True) Then
     '[Only add the fog color when an intersection with a any type of light]
     FogFactor = Shader_GetFog(VectorDistance(TheNewRay.Position, CurrentPoint))
     ResultColors(UBound(ResultColors())) = ColorInterpolate(FogColor, ResultColors(UBound(ResultColors())), FogFactor)
    End If

    ReDim Preserve ResultColors(UBound(ResultColors()) + 1)
    ResultDiffusions(UBound(ResultDiffusions())) = ((DiffusionRate * 0.5) * AlphaFactor)
    ReDim Preserve ResultDiffusions(UBound(ResultDiffusions()) + 1)
    If (DiffusionRate = 500) Then Exit For 'Totaly opaque point, we can't continue the bouncing.
   End If
   DirectionVec = VectorSubtract(TheNewRay.Position, CurrentPoint)
   TheNewRay.Position = CurrentPoint
   RussianRoulette = (Rnd * TotalWeight) 'Propably deciding with the roulette
   If (RussianRoulette <= CurMatInfos.Reflection) Then
    'Reflect the ray
    TheNewRay.Direction = VectorAdd(TheNewRay.Position, VectorReflect(TheFaces(CurIntersection.FaceNumber).Normal, DirectionVec, 1))
   Else
    'Refract the ray
    TheNewRay.Direction = VectorAdd(TheNewRay.Position, VectorRefract(TheFaces(CurIntersection.FaceNumber).Normal, DirectionVec, CurMatInfos.RefractionN, 1, 1))
   End If
   'Avoid self-intersecting :
   TheNewRay.Position = VectorInterpolate(TheNewRay.Position, TheNewRay.Direction, 0.0001)
  End If
 Next CurSample

 If (AnyOne = False) Then
  Exit Function
 Else
  ReDim Preserve ResultColors(UBound(ResultColors()) - 1)
  ReDim Preserve ResultDiffusions(UBound(ResultDiffusions()) - 1)
  If (UBound(ResultDiffusions()) = 0) Then
   ResultColor = ResultColors(0)
  Else
   ResultColor = ResultColors(UBound(ResultColors()))
   For CurrentResult = (UBound(ResultDiffusions()) - 1) To 0 Step -1
    ResultColor = ColorInterpolate(ResultColor, ResultColors(CurrentResult), ResultDiffusions(CurrentResult))
   Next CurrentResult
  End If
 End If

 TraceViewPath = ResultColor

End Function
Function TraceRayHeap(TheRay As Ray3D, AlphaCheck As Boolean) As TraceResult

 If (Started = False) Then Exit Function
 If (TheMeshsCount = -1) Then TraceRayHeap.IntersectCount = -1: Exit Function

 '///////////////////////////////////////////////////////////////
 '
 ' Trace a ray over the 3D scenry, with/out alpha-masks,
 ' and return the resulting intersections in the output array.
 '
 ' The function return the index of the intersected face(s) (also
 ' the index of the mesh where the face came from), the barycentrics
 ' coordinates in the projective space of the triangle(s) for computing
 ' interpolations, a flag to determine the orientation from the ray
 ' (backface, or counter/clockwise directions), and the parametric
 ' distance of the intersection point on the ray.
 '
 ' The algorithm orient (transform) the input triangles by a view matrix,
 ' using the given input ray as a view-ray, in the way that TheRay.Position is the
 ' view-point, and TheRay.Direction is the look-at point, after that, this is a
 ' simple 2D-case check (point in triangle 2D test), a similar algorithm
 ' is the Moller's algorithm [Moller97], that's use orientations for tracing.
 ' There are many algorithms for computing ray-triangle intersections, like
 ' [Snyder87], [Badouel90], [Segura98], you can see the PDF file in the
 ' program folder in '\Extras\Documentations' section.
 '
 ' To accelerate the tracing routine, we use the 'bounding volume'
 ' mechanism (box in this program), before checking the ray with all
 ' the mesh's triangles, we firstly test for any possible intersection
 ' by the bounding volume, if not, we simply skip all the internal
 ' checks (1 test for saving many tests).
 '
 '///////////////////////////////////////////////////////////////

 Dim CurMesh&, CurFace&, CurIntersection&, RayViewMatrix As Matrix4x4
 Dim TmpVec1 As Vector3D, TmpVec2 As Vector3D, TmpVec3 As Vector3D
 Dim IsABackFace As Boolean, OutCol As Byte, OldCount As Long
 Dim D!, U!, V!, W!, InterpoledZ!, Dist!, InvDist!, AnyHit As Boolean
 Dim TmpI As Intersection3D, DoTracing() As Boolean, MinDist!, MaxDist!

 'Check with the scene's bounding box:
 If ((RayBoxIntersect(TheRay, SceneBoundingBox) = False) And (TheMeshsCount > 0)) Then
  TraceRayHeap.IntersectCount = -1: Exit Function
 End If

 'Deciding the tracing for every mesh:
 ReDim DoTracing(TheMeshsCount)
 For CurMesh = 0 To TheMeshsCount
  If (RayBoxIntersect(TheRay, MeshsBoundingBoxes(CurMesh)) = False) Then
   DoTracing(CurMesh) = False
  Else
   DoTracing(CurMesh) = True
  End If
 Next CurMesh

 'Make a 'ray-level' view matrix:
 RayViewMatrix = MatrixView(TheRay.Position, TheRay.Direction, 0)
 Dist = VectorDistance(TheRay.Position, TheRay.Direction)
 ReDim TraceRayHeap.Intersections(0)
 For CurMesh = 0 To TheMeshsCount
  If ((TheMeshs(CurMesh).Visible = True) And (DoTracing(CurMesh) = True)) Then
   For CurFace = TheMeshs(CurMesh).Faces.Start To GetAddressLast(TheMeshs(CurMesh).Faces)
    If (TheFaces(CurFace).Visible = True) Then
     'Transform the triangle's vertices with the ray matrix:
     TmpVec1 = MatrixMultiplyVector(TheVertices(TheFaces(CurFace).A).TmpPos, RayViewMatrix)
     TmpVec2 = MatrixMultiplyVector(TheVertices(TheFaces(CurFace).B).TmpPos, RayViewMatrix)
     TmpVec3 = MatrixMultiplyVector(TheVertices(TheFaces(CurFace).C).TmpPos, RayViewMatrix)
     '2D inside check:
     If (IsPointInTriangle(TmpVec1.X, TmpVec1.Y, TmpVec2.X, TmpVec2.Y, TmpVec3.X, TmpVec3.Y, 0, 0) = True) Then
      D = (((TmpVec2.X - TmpVec1.X) * (TmpVec3.Y - TmpVec1.Y)) - ((TmpVec2.Y - TmpVec1.Y) * (TmpVec3.X - TmpVec1.X)))
      If (D <> 0) Then
       D = (1 / D)
       'Calculate the barycentrics coordinates:
       U = (((TmpVec2.X * TmpVec3.Y) - (TmpVec2.Y * TmpVec3.X)) * D)
       V = (((TmpVec3.X * TmpVec1.Y) - (TmpVec3.Y * TmpVec1.X)) * D)
       W = (1 - (U + V))
       'Recieve the distance of the intersection point on the triangle,
       'by interpolating the vertices distances(Zs) using the barycentrics coordinates:
       InterpoledZ = (((U * TmpVec1.Z) + (V * TmpVec2.Z)) + (W * TmpVec3.Z))
       If ((InterpoledZ >= 0) And (InterpoledZ <= Dist)) Then 'Distance limits on the ray
        InvDist = (1 / Dist)
        If (AlphaCheck = True) Then
         With TmpI
          .MeshNumber = CurMesh: .FaceNumber = CurFace
          .IsBackFace = IsBackFace(TmpVec1.X, TmpVec1.Y, TmpVec2.X, TmpVec2.Y, TmpVec3.X, TmpVec3.Y)
          .U = U: .V = V: .W = W: .Zt = (InterpoledZ * InvDist)
         End With
         If (GetMaterialVisibility(TmpI) = True) Then
          With TraceRayHeap.Intersections(UBound(TraceRayHeap.Intersections()))
           .MeshNumber = TmpI.MeshNumber: .FaceNumber = TmpI.FaceNumber
           .U = TmpI.U: .V = TmpI.V: .W = TmpI.W: .Zt = TmpI.Zt
           .IsBackFace = TmpI.IsBackFace
          End With
          ReDim Preserve TraceRayHeap.Intersections(UBound(TraceRayHeap.Intersections()) + 1)
          TraceRayHeap.IntersectCount = (TraceRayHeap.IntersectCount + 1): AnyHit = True
         End If
        Else
         With TraceRayHeap.Intersections(UBound(TraceRayHeap.Intersections()))
          .MeshNumber = CurMesh: .FaceNumber = CurFace
          .IsBackFace = IsBackFace(TmpVec1.X, TmpVec1.Y, TmpVec2.X, TmpVec2.Y, TmpVec3.X, TmpVec3.Y)
          .U = U: .V = V: .W = W: .Zt = (InterpoledZ * InvDist)
         End With
         ReDim Preserve TraceRayHeap.Intersections(UBound(TraceRayHeap.Intersections()) + 1)
         TraceRayHeap.IntersectCount = (TraceRayHeap.IntersectCount + 1): AnyHit = True
        End If
       End If
      End If
     End If
    End If
   Next CurFace
  End If
 Next CurMesh

 If (AnyHit = False) Then
  TraceRayHeap.IntersectCount = -1: Erase TraceRayHeap.Intersections(): Exit Function
 Else
  ReDim Preserve TraceRayHeap.Intersections(UBound(TraceRayHeap.Intersections()) - 1)
  TraceRayHeap.IntersectCount = (TraceRayHeap.IntersectCount - 1)
  If (TraceRayHeap.IntersectCount <> 0) Then SortIntersections TraceRayHeap, True
 End If

End Function
Function RemoveBackFacesIntersections(TheTraceResults As TraceResult)

 If (TheTraceResults.IntersectCount = -1) Then Exit Function

 Dim CurIntersection&

ReCheck:
 For CurIntersection = 0 To TheTraceResults.IntersectCount
  If (TheTraceResults.Intersections(CurIntersection).IsBackFace = True) Then
   RemoveIntersection TheTraceResults, CurIntersection
   GoTo ReCheck
  End If
 Next CurIntersection

 SortIntersections TheTraceResults, True

End Function
Function TraceRayFirst(TheRay As Ray3D, AlphaCheck As Boolean, IgnoreLength As Boolean) As Intersection3D

 If (Started = False) Then Exit Function

 'Return the first point of intersection, use IgnoreLength flag
 'to say that the distance of the ray isn't important, and so the
 'ray is traced to the infinity.

 Dim TmpRay As Ray3D

 TmpRay.Position = TheRay.Position
 TmpRay.Direction = TheRay.Direction

 If (IgnoreLength = True) Then 'Scale the ray to the virtual 'infinity'
  TmpRay.Direction = VectorAdd(TmpRay.Position, VectorScale(VectorSubtract(TmpRay.Direction, TmpRay.Position), 49999))
 End If

 Dim TheTraceResult As TraceResult
 TheTraceResult = TraceRayHeap(TmpRay, AlphaCheck)

 If (TheTraceResult.IntersectCount = -1) Then
  With TraceRayFirst
   .MeshNumber = -1: .FaceNumber = -1
  End With
 Else
  TraceRayFirst = TheTraceResult.Intersections(TheTraceResult.IntersectCount)
 End If

End Function
Function TraceRayFirstBackFacesCull(TheRay As Ray3D, AlphaCheck As Boolean, IgnoreLength As Boolean) As Intersection3D

 If (Started = False) Then Exit Function

 'Return the first point of intersection, use IgnoreLength flag
 'to say that the distance of the ray isn't important, and so the
 'ray is traced to the infinity.

 Dim TmpRay As Ray3D

 TmpRay.Position = TheRay.Position
 TmpRay.Direction = TheRay.Direction

 If (IgnoreLength = True) Then 'Scale the ray to the virtual 'infinity'
  TmpRay.Direction = VectorAdd(TmpRay.Position, VectorScale(VectorSubtract(TmpRay.Direction, TmpRay.Position), 49999))
 End If

 Dim TheTraceResult As TraceResult
 TheTraceResult = TraceRayHeap(TmpRay, AlphaCheck)

 If (TheTraceResult.IntersectCount = -1) Then
  With TraceRayFirstBackFacesCull
   .MeshNumber = -1: .FaceNumber = -1
  End With
 Else
  If (TheCameras(TheCurrentCamera).BackFaceCulling = True) Then RemoveBackFacesIntersections TheTraceResult
  If (TheTraceResult.IntersectCount = -1) Then
   With TraceRayFirstBackFacesCull
    .MeshNumber = -1: .FaceNumber = -1
   End With
  Else
   TraceRayFirstBackFacesCull = TheTraceResult.Intersections(TheTraceResult.IntersectCount)
  End If
 End If

End Function
Function RemoveIntersection(TheTraceResults As TraceResult, TheIntersectionNumber As Long)

 'Remove an intersection from the intersections heap

 If (TheTraceResults.IntersectCount = 0) Then
  Erase TheTraceResults.Intersections()
  TheTraceResults.IntersectCount = -1
  Exit Function
 End If

 Dim CurIntersection As Long
 For CurIntersection = TheIntersectionNumber To (TheTraceResults.IntersectCount - 1)
  TheTraceResults.Intersections(CurIntersection) = TheTraceResults.Intersections(CurIntersection + 1)
 Next CurIntersection

 ReDim Preserve TheTraceResults.Intersections(UBound(TheTraceResults.Intersections()) - 1)
 TheTraceResults.IntersectCount = (TheTraceResults.IntersectCount - 1)

End Function
Function SortIntersections(TheTraceResults As TraceResult, Ascending As Boolean)

 'Uses extraction-sort algorithm to sort the intersections
 '(on the ray), this is like the Z-buffer in some way...

 Dim Elem1&, Elem2&, Elem3&, TmpIntersection As Intersection3D

 If (Ascending = True) Then
  For Elem1 = 0 To TheTraceResults.IntersectCount
   Elem3 = Elem1
   For Elem2 = Elem1 To TheTraceResults.IntersectCount
    If (TheTraceResults.Intersections(Elem3).Zt < TheTraceResults.Intersections(Elem2).Zt) Then Elem3 = Elem2
   Next Elem2
   TmpIntersection = TheTraceResults.Intersections(Elem3)
   TheTraceResults.Intersections(Elem3) = TheTraceResults.Intersections(Elem1)
   TheTraceResults.Intersections(Elem1) = TmpIntersection
  Next Elem1
 ElseIf (Ascending = False) Then
  For Elem1 = 0 To TheTraceResults.IntersectCount
   Elem3 = Elem1
   For Elem2 = Elem1 To TheTraceResults.IntersectCount
    If (TheTraceResults.Intersections(Elem3).Zt > TheTraceResults.Intersections(Elem2).Zt) Then Elem3 = Elem2
   Next Elem2
   TmpIntersection = TheTraceResults.Intersections(Elem3)
   TheTraceResults.Intersections(Elem3) = TheTraceResults.Intersections(Elem1)
   TheTraceResults.Intersections(Elem1) = TmpIntersection
  Next Elem1
 End If

End Function
Function EmittPhotonsFromSphereLight(TheSphereLight As SphereLight3D, MaxAllocatedPhotons&) As Integer

 If (Started = False) Then Exit Function

 'Emmit a photons heap from an omni light.

 'After this limited number of tries (tracing rays)
 'and no one intersection at least is found, the function
 'return a value that means that the light hits nothing.
 Const MaxTries% = 1000

 Dim CurPhoton As Photon, PhotonPathsCount&, CurPhotonPath&
 Dim PhotonsHeapSent&, PhotonTraceResult&, LastSamples&, TriesCount&
 Dim PhotonPower As Vector3D, EmissionPurcent!

 PhotonPower = VectorScale(ColorRGBToVector(TheSphereLight.Color), (1 / MaxAllocatedPhotons))
 PhotonPathsCount = (MaxAllocatedPhotons / SamplesPerPhotonPath)

 Do
  For CurPhotonPath = 1 To PhotonPathsCount
   LastSamples = (MaxAllocatedPhotons - PhotonsHeapSent): TriesCount = 0
   CurPhoton.Direction = TheSphereLight.TmpPos: CurPhoton.Power = PhotonPower
   Do
    CurPhoton.Position = VectorAdd(VectorInput(SignedRnd, SignedRnd, SignedRnd), CurPhoton.Direction)
    If (LastSamples < SamplesPerPhotonPath) Then
     PhotonTraceResult = TracePhotonPath(CurPhoton, LastSamples, TheSphereLight, ConeLight3D_Null, True)
    Else
     PhotonTraceResult = TracePhotonPath(CurPhoton, -1, TheSphereLight, ConeLight3D_Null, True)
    End If
    If (PhotonTraceResult = -1) Then 'The photonmap is full
     EmittPhotonsFromSphereLight = -1: Exit Function
    ElseIf (PhotonTraceResult = -2) Then 'Light hits nothing !
     If ((TriesCount = MaxTries) And (PhotonsHeapSent = 0)) Then
      EmittPhotonsFromSphereLight = -2: Exit Function
     Else
      TriesCount = (TriesCount + 1)
     End If
    Else
     PhotonsHeapSent = (PhotonsHeapSent + PhotonTraceResult)
     EmissionPurcent = (PhotonsSentCount / MaximumAllocatedPhotons)
     FRM_Render.Label3.Width = (EmissionPurcent * FRM_Render.Label2.Width)
     FRM_Render.Label4.Caption = "Emitting... " & Fix(EmissionPurcent * 100) & " %"
     DoEvents
     Exit Do
    End If
   Loop
   If (LastSamples < SamplesPerPhotonPath) Then Exit For
  Next CurPhotonPath
 Loop Until (PhotonsHeapSent = MaxAllocatedPhotons)

End Function
Function EmittPhotonsFromConeLight(TheConeLight As ConeLight3D, MaxAllocatedPhotons&) As Integer

 If (Started = False) Then Exit Function

 'Emmit a photons heap from a spot light.

 'After this limited number of tries (tracing rays)
 'and no one intersection at least is found, the function
 'return a value that means that the light hits nothing.
 Const MaxTries% = 1000

 Dim CurPhoton As Photon, PhotonPathsCount&, CurPhotonPath&
 Dim PhotonsHeapSent&, PhotonTraceResult&, LastSamples&, TriesCount&
 Dim PhotonPower As Vector3D, EmissionPurcent!

 PhotonPower = VectorScale(ColorRGBToVector(TheConeLight.Color), (1 / MaxAllocatedPhotons))
 PhotonPathsCount = (MaxAllocatedPhotons / SamplesPerPhotonPath)

 Do
  For CurPhotonPath = 1 To PhotonPathsCount
   LastSamples = (MaxAllocatedPhotons - PhotonsHeapSent): TriesCount = 0
   CurPhoton.Direction = TheConeLight.TmpPos: CurPhoton.Power = PhotonPower
   Do
    CurPhoton.Position = VectorAdd(VectorInput(SignedRnd, SignedRnd, SignedRnd), CurPhoton.Direction)
    If (LastSamples < SamplesPerPhotonPath) Then
     PhotonTraceResult = TracePhotonPath(CurPhoton, LastSamples, SphereLight3D_Null, TheConeLight, False)
    Else
     PhotonTraceResult = TracePhotonPath(CurPhoton, -1, SphereLight3D_Null, TheConeLight, False)
    End If
    If (PhotonTraceResult = -1) Then 'The photonmap is full
     EmittPhotonsFromConeLight = -1: Exit Function
    ElseIf (PhotonTraceResult = -2) Then 'Light hits nothing !
     If ((TriesCount = MaxTries) And (PhotonsHeapSent = 0)) Then
      EmittPhotonsFromConeLight = -2: Exit Function
     Else
      TriesCount = (TriesCount + 1)
     End If
    Else
     PhotonsHeapSent = (PhotonsHeapSent + PhotonTraceResult)
     EmissionPurcent = (PhotonsSentCount / MaximumAllocatedPhotons)
     FRM_Render.Label3.Width = (EmissionPurcent * FRM_Render.Label2.Width)
     FRM_Render.Label4.Caption = "Emitting... " & Fix(EmissionPurcent * 100) & " %"
     DoEvents
     Exit Do
    End If
   Loop
   If (LastSamples < SamplesPerPhotonPath) Then Exit For
  Next CurPhotonPath
 Loop Until (PhotonsHeapSent = MaxAllocatedPhotons)

End Function
Function EstimateRadianceFromPhotonmap(FromPoint As Vector3D, TheRay As Ray3D, TheMatInfos As MaterialInfos) As ColorRGB

 If (Started = False) Then Exit Function

 'Locate the nearest photons from the photon map, (linear search,
 'not a kD-tree for now), and add thier contribution to estimate
 'the indirect-illumination.

 Dim PhotonColor As Vector3D, Output As Vector3D
 Dim CurPhoton&, CurDistance!, CurAngle!, InvEstimateDistance!, TheScalar!

 InvEstimateDistance = (1 / PhotonsSearchRadius)

 For CurPhoton = 0 To PhotonsSentCount
  'Distance check:
  CurDistance = VectorDistance(ThePhotonMap(CurPhoton).Position, TheRay.Position)
  If (CurDistance < PhotonsSearchRadius) Then
   'Angle check:
   CurAngle = Shader_GetLambertAngle(VectorSubtract(FromPoint, ThePhotonMap(CurPhoton).Position), VectorSubtract(ThePhotonMap(CurPhoton).Direction, ThePhotonMap(CurPhoton).Position), TheRay.Direction, VectorSubtract(TheRay.Position, ThePhotonMap(CurPhoton).Position))
   If (CurAngle > 0) Then
    'Filter the estimate to remove the aliasing effect on the
    'photons egdes, but also can gives sharpe edges in the caustics
    TheScalar = (CurDistance * InvEstimateDistance)
    If (TheScalar < EstimateFilterSize) Then
     TheScalar = 1
    Else
     TheScalar = (1 - (TheScalar - EstimateFilterSize) / (1 - EstimateFilterSize))
    End If
    TheScalar = ((TheScalar * CurAngle) * EstimateMultiplier)
    PhotonColor = VectorScale(ThePhotonMap(CurPhoton).Power, TheScalar)
    'Calculate the specular contribution of the photon:
    If (TheMatInfos.Reflection > 0) Then
     TheScalar = Shader_GetSpecularity(VectorSubtract(FromPoint, ThePhotonMap(CurPhoton).Position), VectorSubtract(ThePhotonMap(CurPhoton).Direction, ThePhotonMap(CurPhoton).Position), TheRay.Direction, VectorSubtract(TheRay.Position, ThePhotonMap(CurPhoton).Position))
     TheScalar = (TheScalar * (TheMatInfos.Reflection * AlphaFactor))
     PhotonColor = VectorAdd(PhotonColor, VectorScale(ThePhotonMap(CurPhoton).Power, (TheMatInfos.SpecularPowerK * (TheScalar ^ TheMatInfos.SpecularPowerN))))
    End If
    Output = VectorAdd(Output, PhotonColor) 'Add the photon's contribution
    'Set limits:
    If (Output.X > 255) Then Output.X = 255
    If (Output.Y > 255) Then Output.Y = 255
    If (Output.Z > 255) Then Output.Z = 255
    If (VectorCompare(Output, VectorInput(255, 255, 255)) = True) Then Exit For
   End If
  End If
 Next CurPhoton

 EstimateRadianceFromPhotonmap = ColorVectorToRGB(Output)

End Function
Sub MakeShadowRaysAsSphere()

 ReDim TheShadowRays(ShadowRaysCount - 1)

 Dim CurShadowRay&
 For CurShadowRay = 0 To (ShadowRaysCount - 1)
  TheShadowRays(CurShadowRay) = VectorScale(VectorNormalize(VectorInput(SignedRnd, SignedRnd, SignedRnd)), ShadowsApproxRadius)
 Next CurShadowRay

End Sub
Sub MakeShadowRaysAsCone(TheConeLight As ConeLight3D)

 ReDim TheShadowRays(ShadowRaysCount - 1)

 Dim CurShadowRay&
 For CurShadowRay = 0 To (ShadowRaysCount - 1)
  Do
   TheShadowRays(CurShadowRay) = VectorInput(SignedRnd, SignedRnd, SignedRnd)
  Loop Until Shader_GetShapeCone(TheConeLight.TmpPos, TheConeLight.TmpDir, TheConeLight.Falloff, (TheConeLight.Falloff - ApproachVal), 99999, TheShadowRays(CurShadowRay)) <> 0
  TheShadowRays(CurShadowRay) = VectorScale(VectorNormalize(TheShadowRays(CurShadowRay)), ShadowsApproxRadius)
 Next CurShadowRay

End Sub
Sub Render(RenderToViewPort As Boolean, TheViewPort As BitMap2D)

 If (Started = False) Then Exit Sub
 If (RenderToViewPort = True) Then
  If ((BitMap2D_IsValid(TheViewPort) = False) Or (TheViewPort.BitsDepth = 8)) Then Exit Sub
 End If

 Dim ProjectorW!, ProjectorH!, ImageAspectRatio!, ParamX!, ParamY!, ParamZ!
 Dim RenderPurcent!, StartTime$, EndTime$, CurViewPath&, PathsRed&, PathsGreen&, PathsBlue&
 Dim CurX%, CurY%, CurrentViewRay As Ray3D, EndColor As ColorRGB, ScanedLine() As Long

 FRM_Render.MousePointer = 11

 ProjectorW = 1
 ImageAspectRatio = (OutputWidth / OutputHeight)
 ProjectorH = (ProjectorW * ImageAspectRatio)
 ReDim ScanedLine(OutputWidth)

 'Resize the background image to fit the viewport size:
 TheBackGround = OriginalBackGround
 FRM_Render.Label1.Caption = "Resizing the background..."
 If ((TheTexturesFilter = K3DE_TFM_NEAREST) Or _
     (TheTexturesFilter = K3DE_TFM_NEAREST_MIP_NEAREST) Or _
     (TheTexturesFilter = K3DE_TFM_NEAREST_MIP_LINEAR)) Then
  BitMap2D_Resample TheBackGround, OutputWidth, OutputHeight, K3DE_XFM_NOFILTER
 Else
  BitMap2D_Resample TheBackGround, OutputWidth, OutputHeight, TheTexelsFilter
 End If

 FRM_Render.Label1.Visible = True
 FRM_Render.Label2.Visible = True
 FRM_Render.Label3.Visible = True
 FRM_Render.Label4.Visible = True
 FRM_Render.Command1.Enabled = False
 If (PreviewMode = True) Then
  FRM_Render.Label1.Caption = "Previewing output from the camera : '" & TheCameras(TheCurrentCamera).Label & "'"
 Else
  FRM_Render.Label1.Caption = " Output size: " & OutputWidth & "x" & OutputHeight & ", from the camera : '" & TheCameras(TheCurrentCamera).Label & "'"
 End If
 FRM_Render.Picture3.Width = (OutputWidth * 15)
 FRM_Render.Picture3.Height = (OutputHeight * 15)
 FRM_Render.Picture3.Cls
 DoEvents

 StartTime = Time

 '////////////////////////////////////////////////////////////////////////

 For CurY = 0 To OutputHeight
  ParamY = (CurY / OutputHeight)
  ParamY = (-(ProjectorW * 0.5) + ((ProjectorW * 0.5) - -(ProjectorW * 0.5)) * ParamY)
  For CurX = 0 To OutputWidth
   If (StopRender = True) Then
    If (MsgBox("Abort the rendering proccess ?", (vbQuestion + vbYesNo), "Abort") = vbYes) Then
     Exit For
    Else
     StopRender = False
    End If
   End If
   ParamX = (CurX / OutputWidth)
   ParamX = (-(ProjectorH * 0.5) + ((ProjectorH * 0.5) - -(ProjectorH * 0.5)) * ParamX)
   ParamZ = FocalDistance((Pi - TheCameras(TheCurrentCamera).FOVAngle), (ProjectorW * 0.5))
   CurrentViewRay.Direction = VectorInput(ParamX, ParamY, ParamZ)

   '***************************************************************
   EndColor = TraceViewPath(CurrentViewRay)
   If (ColorCompare(EndColor, ColorInput(-1, -1, -1)) = True) Then GoTo NoIntersection
   If (ViewPathsPerPixel > 1) Then
    PathsRed = EndColor.R: PathsGreen = EndColor.G: PathsBlue = EndColor.B
    For CurViewPath = 1 To (ViewPathsPerPixel - 1)
     EndColor = TraceViewPath(CurrentViewRay)
     PathsRed = (PathsRed + EndColor.R)
     PathsGreen = (PathsGreen + EndColor.G)
     PathsBlue = (PathsBlue + EndColor.B)
    Next CurViewPath
    PathsRed = (PathsRed / ViewPathsPerPixel)
    PathsGreen = (PathsGreen / ViewPathsPerPixel)
    PathsBlue = (PathsBlue / ViewPathsPerPixel)
    EndColor.R = PathsRed
    EndColor.G = PathsGreen
    EndColor.B = PathsBlue
   End If
NoIntersection:
   If (ColorCompare(EndColor, ColorInput(-1, -1, -1)) = True) Then
    If (UseBackGround = True) Then
     EndColor.R = TheBackGround.Datas(0, CurX, CurY)
     EndColor.G = TheBackGround.Datas(1, CurX, CurY)
     EndColor.B = TheBackGround.Datas(2, CurX, CurY)
    Else
     EndColor = TheBackGroundColor
    End If
   End If
   '***************************************************************

   If (RenderToViewPort = True) Then
    TheViewPort.Datas(0, CurX, CurY) = EndColor.R
    TheViewPort.Datas(1, CurX, CurY) = EndColor.G
    TheViewPort.Datas(2, CurX, CurY) = EndColor.B
   End If
   ScanedLine(CurX) = ColorRGBToLong(EndColor)
  Next CurX
  If (StopRender = True) Then Exit For
  For CurX = 0 To OutputWidth
   FRM_Render.Picture3.PSet (CurX, CurY), ScanedLine(CurX)
  Next CurX
  FRM_Render.Line1.Y1 = CurY: FRM_Render.Line1.Y2 = CurY
  FRM_Render.Line1.X1 = 0: FRM_Render.Line1.X2 = OutputWidth
  RenderPurcent = (CurY / OutputHeight)
  FRM_Render.Label3.Width = (RenderPurcent * FRM_Render.Label2.Width)
  If (PreviewMode = True) Then
   FRM_Render.Label4.Caption = "Previewing... " & Fix(RenderPurcent * 100) & " %"
  Else
   FRM_Render.Label4.Caption = "Rendering... " & Fix(RenderPurcent * 100) & " %"
  End If
  DoEvents
 Next CurY

 '////////////////////////////////////////////////////////////////////////

 If (StopRender = True) Then
  StopRender = False: Unload FRM_Render: FRM_Main.Show: Exit Sub
 End If

 EndTime = Time

 'Ajust scroll bars:
 If (FRM_Render.Picture3.Width > FRM_Render.Picture1.Width) Then
  FRM_Render.HScroll1.Enabled = True: FRM_Render.HScroll1.Max = (FRM_Render.Picture1.Width - FRM_Render.Picture3.Width)
 Else
  FRM_Render.HScroll1.Enabled = False
 End If
 If (FRM_Render.Picture3.Height > FRM_Render.Picture1.Height) Then
  FRM_Render.VScroll1.Enabled = True: FRM_Render.VScroll1.Max = (FRM_Render.Picture1.Height - FRM_Render.Picture3.Height)
 Else
  FRM_Render.VScroll1.Enabled = False
 End If

 FRM_Render.Caption = "Render scene (" & ComputeElaspedTime(StartTime, EndTime) & " rendering time)"

 If (PreviewMode = True) Then FRM_Render.Label1.Visible = False
 FRM_Render.Label2.Visible = False
 FRM_Render.Label3.Visible = False
 FRM_Render.Label4.Visible = False
 FRM_Render.Command1.Enabled = True
 FRM_Main.Show

 FRM_Render.Show
 FRM_Render.MousePointer = 0
 If (PreviewMode = False) Then Beep

End Sub
Sub Engine_EmittPhotons()

 If ((Started = False) Or (TheMeshsCount = -1)) Then Exit Sub

 'Do propagate the photons into the scene, from the sources
 'of lights, with a versatile and robuste emmision technique.
 'Photons power are scaled in the way of the attenuations
 'of the sources of light.

 Const InvPi2 As Single = (1 / Pi2)

 Dim AllocatedPhotonsPerLight&, LightsCount&, CurLight&, LightShapeFactor!
 Dim ExitLoop As Boolean, CurLightResult&

 PhotonsSentCount = 0: LightsCount = -1
 If (TheSphereLightsCount <> -1) Then LightsCount = TheSphereLightsCount
 If (TheConeLightsCount <> -1) Then
  If (LightsCount = -1) Then
   LightsCount = TheConeLightsCount
  Else
   LightsCount = (LightsCount + TheConeLightsCount)
  End If
 End If
 If (LightsCount = -1) Then Exit Sub 'No light sources
 LightsCount = (LightsCount + 1)
 AllocatedPhotonsPerLight = Fix(MaximumAllocatedPhotons / LightsCount): Randomize

 FRM_Render.Label2.Visible = True
 FRM_Render.Label3.Visible = True
 FRM_Render.Label4.Visible = True

 ExitLoop = True
 Do
  'Emitt from the sphere light-sources:
  If (TheSphereLightsCount <> -1) Then
   For CurLight = 0 To TheSphereLightsCount
    CurLightResult = EmittPhotonsFromSphereLight(TheSphereLights(CurLight), AllocatedPhotonsPerLight)
    If (CurLightResult <> -2) Then ExitLoop = False
   Next CurLight
  End If
  'Emitt from the cone light-sources:
  If (TheConeLightsCount <> -1) Then
   For CurLight = 0 To TheConeLightsCount
    LightShapeFactor = (TheConeLights(CurLight).Falloff * InvPi2)
    CurLightResult = EmittPhotonsFromConeLight(TheConeLights(CurLight), (AllocatedPhotonsPerLight * LightShapeFactor))
    If (CurLightResult <> -2) Then ExitLoop = False
   Next CurLight
  End If
  If (ExitLoop = True) Then EstimateFromPhotonmap = False: Exit Sub
 Loop Until (PhotonsSentCount = MaximumAllocatedPhotons) 'Send photons until the container is full

 If (PhotonsSentCount > 0) Then PhotonsSentCount = (PhotonsSentCount - 1)
 EstimateFromPhotonmap = True

End Sub
Sub Engine_Reset()

 If (Started = False) Then Exit Sub

 'Do a master reset

 Scene3D_Clear
 Engine_Start

End Sub
Sub Engine_Start()

 'Do a master start

 Engine_AllocateMemory

 TheMeshsCount = -1
 TheSphereLightsCount = -1
 TheConeLightsCount = -1
 TheCamerasCount = -1

 'Since there is no way to visualize the scene
 'without a camera..the engine create a new 'default' camera.
 TheCurrentCamera = Camera3D_Add
 TheCameras(TheCurrentCamera).Label = "Default camera"

 SetDefaultParameters

 Started = True

End Sub
Sub Engine_ComputeNormals()

 If (Started = False) Then Exit Sub

 'Compute the "normal" vectors of the faces, a normal vector
 'is the perpendicular vector to the plane defined by the face's vertices.

 Dim CurMesh&, CurFace&

 For CurMesh = 0 To TheMeshsCount
  If (TheMeshs(CurMesh).Visible = True) Then
   For CurFace = TheMeshs(CurMesh).Faces.Start To GetAddressLast(TheMeshs(CurMesh).Faces)
    If (TheFaces(CurFace).Visible = True) Then
     TheFaces(CurFace).Normal = VectorGetNormal(TheVertices(TheFaces(CurFace).A).TmpPos, _
                                                TheVertices(TheFaces(CurFace).B).TmpPos, _
                                                TheVertices(TheFaces(CurFace).C).TmpPos)
    End If
   Next CurFace
  End If
 Next CurMesh

End Sub
Sub Engine_Render(TheViewPort As BitMap2D)

 If (Started = False) Then Exit Sub

 'Make the ViewMatrix of the current selected camera:
 If (TheCameras(TheCurrentCamera).MakeMatrix = False) Then
  TheViewMatrix = TheCameras(TheCurrentCamera).ViewMatrix
 Else
  TheCameras(TheCurrentCamera).ViewMatrix = MatrixView(TheCameras(TheCurrentCamera).Position, TheCameras(TheCurrentCamera).Direction, TheCameras(TheCurrentCamera).RollAngle)
  TheViewMatrix = TheCameras(TheCurrentCamera).ViewMatrix
 End If

 'World/View transforms:
 Engine_Transform False

 'Normal vectors computation:
 Engine_ComputeNormals

 'Compute the bounding Boxes:
 Engine_ComputeBoundingBoxes

 FRM_Manager.Hide: FRM_Main.Hide

 'Photons emission:
 If (EnablePhotonMapping = False) Then
  ReDim ThePhotonMap(0): PhotonsSentCount = 0
 Else
  If (UBound(ThePhotonMap()) <> MaximumAllocatedPhotons) Then ReDim ThePhotonMap(MaximumAllocatedPhotons): PhotonsSentCount = 0
  If ((PreviewMode = True) Or ((PreviewMode = False) And (Previewed = False))) Then Engine_EmittPhotons
 End If

 'Render the 3D
 Render False, TheViewPort

End Sub
Sub Engine_Transform(TransformLightsOnly As Boolean)

 If (Started = False) Then Exit Sub

 '///////////////////////////////////////////////////////
 '//////////////// World/View transforms ////////////////
 '///////////////////////////////////////////////////////

 'Just for the Wirefame previewing option, We don't need to
 'tranform the geometry if we do not display the geometry !, logic no?!
 If (TransformLightsOnly = True) Then GoTo Jump

 Dim CurMesh&, CurVertex&, CurLight&, TheWorldMatrix As Matrix4x4

 'Transfrom the geometry
 '======================

 For CurMesh = 0 To TheMeshsCount
  If (TheMeshs(CurMesh).MakeMatrix = False) Then
   TheWorldMatrix = TheMeshs(CurMesh).WorldMatrix
  Else
   TheMeshs(CurMesh).WorldMatrix = MatrixWorld(TheMeshs(CurMesh).Position, TheMeshs(CurMesh).Scales, TheMeshs(CurMesh).Angles.X, TheMeshs(CurMesh).Angles.Y, TheMeshs(CurMesh).Angles.Z)
   TheWorldMatrix = TheMeshs(CurMesh).WorldMatrix
  End If
  TheTotalMatrix = MatrixMultiply(TheWorldMatrix, TheViewMatrix)
  For CurVertex = TheMeshs(CurMesh).Vertices.Start To GetAddressLast(TheMeshs(CurMesh).Vertices)
   TheVertices(CurVertex).TmpPos = MatrixMultiplyVector(TheVertices(CurVertex).Position, TheTotalMatrix)
  Next CurVertex
 Next CurMesh

Jump:

 'Transfrom the lights
 '====================

 'Sphere lights:
 For CurLight = 0 To TheSphereLightsCount
  TheSphereLights(CurLight).TmpPos = MatrixMultiplyVector(TheSphereLights(CurLight).Position, TheViewMatrix)
 Next CurLight

 'Cone lights:
 For CurLight = 0 To TheConeLightsCount
  TheConeLights(CurLight).TmpPos = MatrixMultiplyVector(TheConeLights(CurLight).Position, TheViewMatrix)
  TheConeLights(CurLight).TmpDir = MatrixMultiplyVector(TheConeLights(CurLight).Direction, TheViewMatrix)
 Next CurLight

End Sub
Function TracePhotonPath(TheStartPhotonPath As Photon, BouncesCount&, OLight As SphereLight3D, SLight As ConeLight3D, IsAnOmni As Boolean) As Long

 If (Started = False) Then Exit Function

 Dim CurSample&, CurPhoton As Photon, TmpRay As Ray3D
 Dim CurIntersection As Intersection3D, CurMatInfos As MaterialInfos
 Dim CurrentPoint As Vector3D, DirectionVec As Vector3D, RussianRoulette%, BounceTimes&
 Dim DiffusionRate%, Prop1 As Boolean, Prop2 As Boolean, Prop3 As Boolean
 Dim Dist!, LightShapeFactor!, InitialPhotonsCount&, Updated As Boolean

 If (BouncesCount = -1) Then BounceTimes = SamplesPerPhotonPath Else BounceTimes = BouncesCount

 CurPhoton = TheStartPhotonPath
 InitialPhotonsCount = PhotonsSentCount
 For CurSample = 1 To BounceTimes
  'Limits check:
  If (PhotonsSentCount = (MaximumAllocatedPhotons + 1)) Then
   PhotonsSentCount = MaximumAllocatedPhotons: TracePhotonPath = -1: Exit Function 'Container is full
  End If
  TmpRay.Position = CurPhoton.Direction: TmpRay.Direction = CurPhoton.Position
  CurIntersection = TraceRayFirst(TmpRay, True, True)
  If (CurIntersection.MeshNumber = -1) Then 'No intersection
   If (Updated = False) Then
    TracePhotonPath = -2
   Else
    TracePhotonPath = (PhotonsSentCount - InitialPhotonsCount)
   End If
   Exit Function
  Else
   CurMatInfos = GetMaterialInfos(CurIntersection) 'Recieve the material's informations at the point
   CurrentPoint = GetIntersectionPoint(CurIntersection)
   DiffusionRate = (500 - (CInt(CurMatInfos.Reflection) + CInt(CurMatInfos.Refraction)))
   If (CurSample = 1) Then 'First bounce of the light ray
    'Scale the photon's power in the way of the attenuation of the
    'light source (shape), we this little truck to avoid to define
    'photometric light sources (light units are used both in the
    'in/direct illumination).
    If (IsAnOmni = True) Then
     LightShapeFactor = Shader_GetAttenuation(OLight.TmpPos, (OLight.Range * 0.5), CurrentPoint)
    Else
     LightShapeFactor = Shader_GetShapeCone(SLight.TmpPos, SLight.TmpDir, SLight.Falloff, SLight.Hotspot, (SLight.Range * 0.5), CurrentPoint)
    End If
    If (LightShapeFactor = 0) Then
     TracePhotonPath = -2: Exit Function
    Else
     'Update the photon radiance by the intersection color:
     CurPhoton.Power = VectorScale(CurPhoton.Power, LightShapeFactor)
     CurPhoton.Power = ColorAbsorp2(ColorRGBToVector(CurMatInfos.Color), CurPhoton.Power)
     If (VectorCompare(CurPhoton.Power, VectorNull) = True) Then Exit Function 'Fully absorped
    End If
   Else
    'Conditions for recording the photon in the photonmap:
    '1. Record only from the second bounce, or the result will be
    '   simply the direct illumination based on photons!
    If (DiffusionRate > 0) Then
     'Update the photon radiance by the intersection color:
     CurPhoton.Power = ColorAbsorp2(ColorRGBToVector(CurMatInfos.Color), CurPhoton.Power)
     If (VectorCompare(CurPhoton.Power, VectorNull) = True) Then Exit Function 'Fully absorped
     '2. Diffuse or partially diffuse surface
     Dist = VectorDistance(CurPhoton.Position, CurrentPoint)
     If (Dist < BleedingDistance) Then
      '3. Distance is less than color-bledding distance, or
      '   we can see the color-bledding effect from 100 km away..?!
      CurPhoton.Power = VectorScale(CurPhoton.Power, (1 - (Dist / BleedingDistance)))
      '4. Absoption part: my choice, is to record the less photons in the dark
      '   area and vise-versa (roulette), instead of a naive absoption-map.
      If (ColorCompare(CurMatInfos.Color, ColorBlack) = False) Then
       RussianRoulette = (Rnd * 255) '33% chance of recording the photon
       If ((CurMatInfos.Color.R <> 0) And (RussianRoulette <= CurMatInfos.Color.R)) Then Prop1 = True Else Prop1 = False
       RussianRoulette = (Rnd * 255) '+ 33% chance
       If ((CurMatInfos.Color.G <> 0) And (RussianRoulette <= CurMatInfos.Color.G)) Then Prop2 = True Else Prop2 = False
       RussianRoulette = (Rnd * 255) '+ 34% chance
       If ((CurMatInfos.Color.B <> 0) And (RussianRoulette <= CurMatInfos.Color.B)) Then Prop3 = True Else Prop3 = False
       If (((Prop1 = True) And (Prop2 = True)) And (Prop3 = True)) Then
        'Record the photon:
        ThePhotonMap(PhotonsSentCount).Power = CurPhoton.Power
        ThePhotonMap(PhotonsSentCount).Position = CurrentPoint
        ThePhotonMap(PhotonsSentCount).Direction = CurPhoton.Direction
        PhotonsSentCount = (PhotonsSentCount + 1): Updated = True
       End If
      End If
     End If
    End If
   End If

   DirectionVec = VectorSubtract(CurPhoton.Direction, CurrentPoint)
   RussianRoulette = (Rnd * (500 - DiffusionRate))

   ' So what the hell is Russian Roulette ?
   ' --------------------------------------
   '
   ' Russian Roulette is a variant of the Monte-Carlo techniques, the main
   ' goal is to randomly distributing samples in an specific domain, where is
   ' infeasible or impossible to do this in other way, still respecting the
   ' weight of the propability (statistical sampling, or random sampling).
   '
   ' A simple example, if we takes a random number from 0 to 1000, then
   ' the purcent of a propability p that this number is <= 300,
   ' will be 30%, 70% otherwise. by generating more random numbers, we
   ' get more precision of the 30 purcent value with more exactitude.
   '
   ' We use Russian Roulette in this case, to propably decide how many
   ' rays should be reflected or refracted, depends on the % of the surface
   ' material properties, if by example a surface reflect 0.7, and refract
   ' 0.3, the most 70% of rays are reflected, the other 30% of rays are refracted.
   '
   ' We Russian Roulette as a distribution technique, because we can obtain
   ' the correct result (or just an approximation) staying controling the
   ' memory costs (rays count).

   If (RussianRoulette <= CurMatInfos.Reflection) Then
    'Reflect the ray, both in diffuse or specular.
    '
    ' *Why?.. why not just a random walk for diffusion like monte-carlo methods ??*
    '
    ' Because:
    '  1. The reflected ray, is one from an infinite rays in the hemisphere
    '     directed by the normal vector from the intersection point.
    '  2. I use the reflected rays, instead of random rays, to globaly
    '     populate the photons over the scene, even with a low number of photons,
    '     because the reflection has like 'an order', that a random walk, has not.
    CurPhoton.Direction = CurrentPoint
    CurPhoton.Position = VectorAdd(CurPhoton.Direction, VectorReflect(TheFaces(CurIntersection.FaceNumber).Normal, DirectionVec, 1))
   Else
    'Refract the ray (transmission):
    CurPhoton.Direction = CurrentPoint
    CurPhoton.Position = VectorAdd(CurPhoton.Direction, VectorRefract(TheFaces(CurIntersection.FaceNumber).Normal, DirectionVec, CurMatInfos.RefractionN, 1, 1))
   End If
   'Avoid self-intersecting:
   CurPhoton.Direction = VectorInterpolate(CurPhoton.Direction, CurPhoton.Position, 0.0001)
  End If
 Next CurSample

End Function
Sub Engine_WireframePreview(TheCanavas As PictureBox, DisplayGeometry As Boolean, DisplayNormals As Boolean, DisplaySLights As Boolean, DisplayCLights As Boolean, DisplayCameras As Boolean, DisplayPhotonmap As Boolean, PerspectiveProjection As Boolean)

 If (Started = False) Then Exit Sub

 'Do a fast wirfeframe pre-visualization.

 On Error Resume Next

 If (TheMeshsCount = -1) Then
  If (TheSphereLightsCount = -1) Then
   If (TheConeLightsCount = -1) Then
    If (TheCamerasCount = -1) Then
     If (EnablePhotonMapping = False) Then
      Exit Sub 'Nothing to display !
     End If
    End If
   End If
  End If
 End If

 'Make the ViewMatrix:
 If (TheCameras(TheCurrentCamera).MakeMatrix = False) Then
  TheViewMatrix = TheCameras(TheCurrentCamera).ViewMatrix
 Else
  TheCameras(TheCurrentCamera).ViewMatrix = MatrixView(TheCameras(TheCurrentCamera).Position, TheCameras(TheCurrentCamera).Direction, TheCameras(TheCurrentCamera).RollAngle)
  TheViewMatrix = TheCameras(TheCurrentCamera).ViewMatrix
 End If

 Engine_Transform True

 '////////////////////////////////////////////////////////////////////////////////

 Dim CurMesh&, CurFace&, CurLight&, CurCamera&, CurPhoton&
 Dim FaceCenter As Vector3D, FaceCenterTo As Vector3D
 Dim Tmp1 As Vector3D, Tmp2 As Vector3D, Tmp3 As Vector3D
 Dim CenterX&, CenterY&, MeshColor&, LightColor&, TmpRange!

 TheCanavas.ScaleMode = vbPixels
 CenterX = (TheCanavas.ScaleWidth * 0.5)
 CenterY = (TheCanavas.ScaleHeight * 0.5)

 If (PerspectiveProjection = False) Then GoTo ParallalView

 '//////////////////////////// GEOMETRY ////////////////////////////////

 If ((DisplayGeometry = False) Or (TheMeshsCount = -1)) Then GoTo Jump1

 'World/View transforms:
 Engine_Transform False
 'Normal vectors computation:
 Engine_ComputeNormals

 For CurMesh = 0 To TheMeshsCount
  If (TheMeshs(CurMesh).Visible = True) Then
   MeshColor = ColorRGBToLong(TheMaterials(CurMesh).Color)
   For CurFace = TheMeshs(CurMesh).Faces.Start To GetAddressLast(TheMeshs(CurMesh).Faces)
    If (TheFaces(CurFace).Visible = True) Then
     'Draw triangle edges:
     Tmp1.X = ((TheVertices(TheFaces(CurFace).A).TmpPos.X / (TheVertices(TheFaces(CurFace).A).TmpPos.Z + Wire_AddedDepth)) * Wire_PerspectiveDistorsion)
     Tmp1.Y = ((TheVertices(TheFaces(CurFace).A).TmpPos.Y / (TheVertices(TheFaces(CurFace).A).TmpPos.Z + Wire_AddedDepth)) * Wire_PerspectiveDistorsion)
     Tmp1.Z = TheVertices(TheFaces(CurFace).A).TmpPos.Z
     Tmp2.X = ((TheVertices(TheFaces(CurFace).B).TmpPos.X / (TheVertices(TheFaces(CurFace).B).TmpPos.Z + Wire_AddedDepth)) * Wire_PerspectiveDistorsion)
     Tmp2.Y = ((TheVertices(TheFaces(CurFace).B).TmpPos.Y / (TheVertices(TheFaces(CurFace).B).TmpPos.Z + Wire_AddedDepth)) * Wire_PerspectiveDistorsion)
     Tmp2.Z = TheVertices(TheFaces(CurFace).B).TmpPos.Z
     Tmp3.X = ((TheVertices(TheFaces(CurFace).C).TmpPos.X / (TheVertices(TheFaces(CurFace).C).TmpPos.Z + Wire_AddedDepth)) * Wire_PerspectiveDistorsion)
     Tmp3.Y = ((TheVertices(TheFaces(CurFace).C).TmpPos.Y / (TheVertices(TheFaces(CurFace).C).TmpPos.Z + Wire_AddedDepth)) * Wire_PerspectiveDistorsion)
     Tmp3.Z = TheVertices(TheFaces(CurFace).C).TmpPos.Z
     If ((Tmp1.Z > 0) And (Tmp2.Z > 0)) Then TheCanavas.Line (CenterX + Tmp1.X, CenterY + Tmp1.Y)-(CenterX + Tmp2.X, CenterY + Tmp2.Y), MeshColor
     If ((Tmp2.Z > 0) And (Tmp3.Z > 0)) Then TheCanavas.Line (CenterX + Tmp2.X, CenterY + Tmp2.Y)-(CenterX + Tmp3.X, CenterY + Tmp3.Y), MeshColor
     If ((Tmp3.Z > 0) And (Tmp1.Z > 0)) Then TheCanavas.Line (CenterX + Tmp3.X, CenterY + Tmp3.Y)-(CenterX + Tmp1.X, CenterY + Tmp1.Y), MeshColor
    End If
   Next CurFace
  End If
 Next CurMesh

 If (DisplayNormals = False) Then GoTo Jump1
 For CurMesh = 0 To TheMeshsCount
  If (TheMeshs(CurMesh).Visible = True) Then
   MeshColor = ColorRGBToLong(TheMaterials(CurMesh).Color)
   For CurFace = TheMeshs(CurMesh).Faces.Start To GetAddressLast(TheMeshs(CurMesh).Faces)
    If (TheFaces(CurFace).Visible = True) Then
     'Draw face's normal:
     FaceCenter = VectorGetCenter(TheVertices(TheFaces(CurFace).A).TmpPos, _
                                  TheVertices(TheFaces(CurFace).B).TmpPos, _
                                  TheVertices(TheFaces(CurFace).C).TmpPos)
     FaceCenterTo = VectorScale(VectorInverse(TheFaces(CurFace).Normal), Wire_ScaleNormalTo)
     FaceCenterTo = VectorAdd(FaceCenter, FaceCenterTo)
     FaceCenter.X = ((FaceCenter.X / (FaceCenter.Z + Wire_AddedDepth)) * Wire_PerspectiveDistorsion)
     FaceCenter.Y = ((FaceCenter.Y / (FaceCenter.Z + Wire_AddedDepth)) * Wire_PerspectiveDistorsion)
     FaceCenterTo.X = ((FaceCenterTo.X / (FaceCenterTo.Z + Wire_AddedDepth)) * Wire_PerspectiveDistorsion)
     FaceCenterTo.Y = ((FaceCenterTo.Y / (FaceCenterTo.Z + Wire_AddedDepth)) * Wire_PerspectiveDistorsion)
     If ((FaceCenter.Z > 0) And (FaceCenterTo.Z > 0)) Then
      TheCanavas.Circle (CenterX + FaceCenter.X, CenterY + FaceCenter.Y), 1, vbCyan
      TheCanavas.Line (CenterX + FaceCenter.X, CenterY + FaceCenter.Y)-(CenterX + FaceCenterTo.X, CenterY + FaceCenterTo.Y), MeshColor
     End If
    End If
   Next CurFace
  End If
 Next CurMesh

 '//////////////////////////// SPHERE LIGHTS ////////////////////////////////

Jump1:
 If ((DisplaySLights = False) Or (TheSphereLightsCount = -1)) Then GoTo Jump2
 For CurLight = 0 To TheSphereLightsCount
  If (TheSphereLights(CurLight).Enable = True) Then
   Tmp1 = TheSphereLights(CurLight).TmpPos
   If (Tmp1.Z > 0) Then
    LightColor = ColorRGBToLong(TheSphereLights(CurLight).Color)
    Tmp1.X = ((Tmp1.X / (Tmp1.Z + Wire_AddedDepth)) * Wire_PerspectiveDistorsion)
    Tmp1.Y = ((Tmp1.Y / (Tmp1.Z + Wire_AddedDepth)) * Wire_PerspectiveDistorsion)
    TmpRange = ((TheSphereLights(CurLight).Range / (Tmp1.Z + Wire_AddedDepth)) * Wire_PerspectiveDistorsion)
    TheCanavas.Circle (CenterX + Tmp1.X, CenterY + Tmp1.Y), 5, LightColor
    TheCanavas.Circle (CenterX + Tmp1.X, CenterY + Tmp1.Y), 2, LightColor
    TheCanavas.DrawStyle = 2
    TheCanavas.Circle (CenterX + Tmp1.X, CenterY + Tmp1.Y), (TmpRange * 0.5), LightColor
    TheCanavas.DrawStyle = 0
   End If
  End If
 Next CurLight

 '//////////////////////////// CONE LIGHTS ////////////////////////////////

Jump2:
 If ((DisplayCLights = False) Or (TheConeLightsCount = -1)) Then GoTo Jump3
 For CurLight = 0 To TheConeLightsCount
  If (TheConeLights(CurLight).Enable = True) Then
   Tmp1 = TheConeLights(CurLight).TmpPos: Tmp2 = TheConeLights(CurLight).TmpDir
   If ((Tmp1.Z > 0) And (Tmp2.Z > 0)) Then
    LightColor = ColorRGBToLong(TheConeLights(CurLight).Color)
    Tmp1.X = ((Tmp1.X / (Tmp1.Z + Wire_AddedDepth)) * Wire_PerspectiveDistorsion)
    Tmp1.Y = ((Tmp1.Y / (Tmp1.Z + Wire_AddedDepth)) * Wire_PerspectiveDistorsion)
    Tmp2 = VectorAdd(Tmp1, VectorScale(VectorNormalize(VectorSubtract(Tmp2, Tmp1)), TheConeLights(CurLight).Range))
    Tmp2.X = ((Tmp2.X / (Tmp2.Z + Wire_AddedDepth)) * Wire_PerspectiveDistorsion)
    Tmp2.Y = ((Tmp2.Y / (Tmp2.Z + Wire_AddedDepth)) * Wire_PerspectiveDistorsion)
    TheCanavas.Circle (CenterX + Tmp1.X, CenterY + Tmp1.Y), 5, LightColor
    TheCanavas.Circle (CenterX + Tmp1.X, CenterY + Tmp1.Y), 2, LightColor
    TheCanavas.Circle (CenterX + Tmp2.X, CenterY + Tmp2.Y), 2, LightColor
    TheCanavas.DrawStyle = 2
    TheCanavas.Line (CenterX + Tmp1.X, CenterY + Tmp1.Y)-(CenterX + Tmp2.X, CenterY + Tmp2.Y), LightColor
    TheCanavas.DrawStyle = 0
   End If
  End If
 Next CurLight

 '//////////////////////////// CAMERAS ////////////////////////////////

Jump3:
 If ((DisplayCameras = False) Or (TheCamerasCount = -1)) Then GoTo Jump4
 For CurCamera = 0 To TheCamerasCount
  If (CurCamera <> TheCurrentCamera) Then 'D'ont display the current camera
   'Transfrom the camera vectors to view coordinate system, in a temporarly storage:
   Tmp1 = MatrixMultiplyVector(TheCameras(CurCamera).Position, TheViewMatrix)
   Tmp2 = MatrixMultiplyVector(TheCameras(CurCamera).Direction, TheViewMatrix)
   Tmp2 = VectorAdd(Tmp1, VectorScale(VectorNormalize(VectorSubtract(Tmp2, Tmp1)), Wire_CameraTo))
   If ((Tmp1.Z > 0) And (Tmp2.Z > 0)) Then
    Tmp1.X = ((Tmp1.X / (Tmp1.Z + Wire_AddedDepth)) * Wire_PerspectiveDistorsion)
    Tmp1.Y = ((Tmp1.Y / (Tmp1.Z + Wire_AddedDepth)) * Wire_PerspectiveDistorsion)
    Tmp2.X = ((Tmp2.X / (Tmp2.Z + Wire_AddedDepth)) * Wire_PerspectiveDistorsion)
    Tmp2.Y = ((Tmp2.Y / (Tmp2.Z + Wire_AddedDepth)) * Wire_PerspectiveDistorsion)
    TheCanavas.Circle (CenterX + Tmp1.X, CenterY + Tmp1.Y), 7, vbCyan
    TheCanavas.Circle (CenterX + Tmp1.X, CenterY + Tmp1.Y), 5, vbCyan
    TheCanavas.Circle (CenterX + Tmp1.X, CenterY + Tmp1.Y), 3, vbCyan
    TheCanavas.Circle (CenterX + Tmp2.X, CenterY + Tmp2.Y), 3, vbCyan
    TheCanavas.Circle (CenterX + Tmp2.X, CenterY + Tmp2.Y), 1, vbCyan
    TheCanavas.DrawStyle = 1
    TheCanavas.Line (CenterX + Tmp1.X, CenterY + Tmp1.Y)-(CenterX + Tmp2.X, CenterY + Tmp2.Y), vbCyan
    TheCanavas.DrawStyle = 0
   End If
  End If
 Next CurCamera

 '//////////////////////////// PHOTON MAP ////////////////////////////////

Jump4:
 If ((DisplayPhotonmap = False) Or (EnablePhotonMapping = False)) Then Exit Sub
 If (UBound(ThePhotonMap()) <> MaximumAllocatedPhotons) Then Exit Sub
 For CurPhoton = 0 To PhotonsSentCount
  Tmp1 = ThePhotonMap(CurPhoton).Position: Tmp2 = ThePhotonMap(CurPhoton).Direction
  Tmp2 = VectorAdd(Tmp1, VectorScale(VectorNormalize(VectorSubtract(Tmp2, Tmp1)), Wire_PhotonTo))
  If ((Tmp1.Z > 0) And (Tmp2.Z > 0)) Then
   Tmp1.X = ((Tmp1.X / (Tmp1.Z + Wire_AddedDepth)) * Wire_PerspectiveDistorsion)
   Tmp1.Y = ((Tmp1.Y / (Tmp1.Z + Wire_AddedDepth)) * Wire_PerspectiveDistorsion)
   Tmp2.X = ((Tmp2.X / (Tmp2.Z + Wire_AddedDepth)) * Wire_PerspectiveDistorsion)
   Tmp2.Y = ((Tmp2.Y / (Tmp2.Z + Wire_AddedDepth)) * Wire_PerspectiveDistorsion)
   TheCanavas.Line (CenterX + Tmp1.X, CenterY + Tmp1.Y)-(CenterX + Tmp2.X, CenterY + Tmp2.Y), vbGreen
  End If
 Next CurPhoton

 Exit Sub '-----------------------------------------------------------------------

ParallalView:

 CenterX = (CenterX + Wire_ParallalMoveToX)
 CenterY = (CenterY + Wire_ParallalMoveToY)

 '//////////////////////////// GEOMETRY ////////////////////////////////

 If ((DisplayGeometry = False) Or (TheMeshsCount = -1)) Then GoTo Jump21

 'World/View transforms:
 Engine_Transform False
 'Normal vectors computation:
 Engine_ComputeNormals

 For CurMesh = 0 To TheMeshsCount
  If (TheMeshs(CurMesh).Visible = True) Then
   MeshColor = ColorRGBToLong(TheMaterials(CurMesh).Color)
   For CurFace = TheMeshs(CurMesh).Faces.Start To GetAddressLast(TheMeshs(CurMesh).Faces)
    If (TheFaces(CurFace).Visible = True) Then
     'Draw triangle edges:
     Tmp1.X = (TheVertices(TheFaces(CurFace).A).TmpPos.X * Wire_ParallalScale)
     Tmp1.Y = (TheVertices(TheFaces(CurFace).A).TmpPos.Y * Wire_ParallalScale)
     Tmp1.Z = TheVertices(TheFaces(CurFace).A).TmpPos.Z
     Tmp2.X = (TheVertices(TheFaces(CurFace).B).TmpPos.X * Wire_ParallalScale)
     Tmp2.Y = (TheVertices(TheFaces(CurFace).B).TmpPos.Y * Wire_ParallalScale)
     Tmp2.Z = TheVertices(TheFaces(CurFace).B).TmpPos.Z
     Tmp3.X = (TheVertices(TheFaces(CurFace).C).TmpPos.X * Wire_ParallalScale)
     Tmp3.Y = (TheVertices(TheFaces(CurFace).C).TmpPos.Y * Wire_ParallalScale)
     Tmp3.Z = TheVertices(TheFaces(CurFace).C).TmpPos.Z
     If ((Tmp1.Z > 0) And (Tmp2.Z > 0)) Then TheCanavas.Line (CenterX + Tmp1.X, CenterY + Tmp1.Y)-(CenterX + Tmp2.X, CenterY + Tmp2.Y), MeshColor
     If ((Tmp2.Z > 0) And (Tmp3.Z > 0)) Then TheCanavas.Line (CenterX + Tmp2.X, CenterY + Tmp2.Y)-(CenterX + Tmp3.X, CenterY + Tmp3.Y), MeshColor
     If ((Tmp3.Z > 0) And (Tmp1.Z > 0)) Then TheCanavas.Line (CenterX + Tmp3.X, CenterY + Tmp3.Y)-(CenterX + Tmp1.X, CenterY + Tmp1.Y), MeshColor
    End If
   Next CurFace
  End If
 Next CurMesh

 If (DisplayNormals = False) Then GoTo Jump21
 For CurMesh = 0 To TheMeshsCount
  If (TheMeshs(CurMesh).Visible = True) Then
   MeshColor = ColorRGBToLong(TheMaterials(CurMesh).Color)
   For CurFace = TheMeshs(CurMesh).Faces.Start To GetAddressLast(TheMeshs(CurMesh).Faces)
    If (TheFaces(CurFace).Visible = True) Then
     'Draw face's normal:
     FaceCenter = VectorGetCenter(TheVertices(TheFaces(CurFace).A).TmpPos, _
                                  TheVertices(TheFaces(CurFace).B).TmpPos, _
                                  TheVertices(TheFaces(CurFace).C).TmpPos)
     FaceCenterTo = VectorScale(VectorInverse(TheFaces(CurFace).Normal), Wire_ScaleNormalTo)
     FaceCenterTo = VectorAdd(FaceCenter, FaceCenterTo)
     FaceCenter.X = (FaceCenter.X * Wire_ParallalScale): FaceCenter.Y = (FaceCenter.Y * Wire_ParallalScale)
     FaceCenterTo.X = (FaceCenterTo.X * Wire_ParallalScale): FaceCenterTo.Y = (FaceCenterTo.Y * Wire_ParallalScale)
     If ((FaceCenter.Z > 0) And (FaceCenterTo.Z > 0)) Then
      TheCanavas.Circle (CenterX + FaceCenter.X, CenterY + FaceCenter.Y), 1, vbCyan
      TheCanavas.Line (CenterX + FaceCenter.X, CenterY + FaceCenter.Y)-(CenterX + FaceCenterTo.X, CenterY + FaceCenterTo.Y), MeshColor
     End If
    End If
   Next CurFace
  End If
 Next CurMesh

 '//////////////////////////// SPHERE LIGHTS ////////////////////////////////

Jump21:
 If ((DisplaySLights = False) Or (TheSphereLightsCount = -1)) Then GoTo Jump22
 For CurLight = 0 To TheSphereLightsCount
  If (TheSphereLights(CurLight).Enable = True) Then
   Tmp1 = TheSphereLights(CurLight).TmpPos
   If (Tmp1.Z > 0) Then
    LightColor = ColorRGBToLong(TheSphereLights(CurLight).Color)
    Tmp1.X = (Tmp1.X * Wire_ParallalScale): Tmp1.Y = (Tmp1.Y * Wire_ParallalScale)
    TmpRange = (TheSphereLights(CurLight).Range * Wire_ParallalScale)
    TheCanavas.Circle (CenterX + Tmp1.X, CenterY + Tmp1.Y), 5, LightColor
    TheCanavas.Circle (CenterX + Tmp1.X, CenterY + Tmp1.Y), 2, LightColor
    TheCanavas.DrawStyle = 2
    TheCanavas.Circle (CenterX + Tmp1.X, CenterY + Tmp1.Y), (TmpRange * 0.5), LightColor
    TheCanavas.DrawStyle = 0
   End If
  End If
 Next CurLight

 '//////////////////////////// CONE LIGHTS ////////////////////////////////

Jump22:
 If ((DisplayCLights = False) Or (TheConeLightsCount = -1)) Then GoTo Jump23
 For CurLight = 0 To TheConeLightsCount
  If (TheConeLights(CurLight).Enable = True) Then
   Tmp1 = TheConeLights(CurLight).TmpPos: Tmp2 = TheConeLights(CurLight).TmpDir
   If ((Tmp1.Z > 0) And (Tmp2.Z > 0)) Then
    LightColor = ColorRGBToLong(TheConeLights(CurLight).Color)
    Tmp1.X = (Tmp1.X * Wire_ParallalScale): Tmp1.Y = (Tmp1.Y * Wire_ParallalScale)
    Tmp2 = VectorAdd(Tmp1, VectorScale(VectorNormalize(VectorSubtract(Tmp2, Tmp1)), TheConeLights(CurLight).Range))
    Tmp2.X = (Tmp2.X * Wire_ParallalScale): Tmp2.Y = (Tmp2.Y * Wire_ParallalScale)
    TheCanavas.Circle (CenterX + Tmp1.X, CenterY + Tmp1.Y), 5, LightColor
    TheCanavas.Circle (CenterX + Tmp1.X, CenterY + Tmp1.Y), 2, LightColor
    TheCanavas.Circle (CenterX + Tmp2.X, CenterY + Tmp2.Y), 2, LightColor
    TheCanavas.DrawStyle = 2
    TheCanavas.Line (CenterX + Tmp1.X, CenterY + Tmp1.Y)-(CenterX + Tmp2.X, CenterY + Tmp2.Y), LightColor
    TheCanavas.DrawStyle = 0
   End If
  End If
 Next CurLight

 '//////////////////////////// CAMERAS ////////////////////////////////

Jump23:
 If ((DisplayCameras = False) Or (TheCamerasCount = -1)) Then GoTo Jump24
 For CurCamera = 0 To TheCamerasCount
  If (CurCamera <> TheCurrentCamera) Then 'D'ont display the current camera
   'Transfrom the camera vectors to view coordinate system, in a temporarly storage:
   Tmp1 = MatrixMultiplyVector(TheCameras(CurCamera).Position, TheViewMatrix)
   Tmp2 = MatrixMultiplyVector(TheCameras(CurCamera).Direction, TheViewMatrix)
   Tmp2 = VectorAdd(Tmp1, VectorScale(VectorNormalize(VectorSubtract(Tmp2, Tmp1)), Wire_CameraTo))
   If ((Tmp1.Z > 0) And (Tmp2.Z > 0)) Then
    Tmp1.X = (Tmp1.X * Wire_ParallalScale): Tmp1.Y = (Tmp1.Y * Wire_ParallalScale)
    Tmp2.X = (Tmp2.X * Wire_ParallalScale): Tmp2.Y = (Tmp2.Y * Wire_ParallalScale)
    TheCanavas.Circle (CenterX + Tmp1.X, CenterY + Tmp1.Y), 7, vbCyan
    TheCanavas.Circle (CenterX + Tmp1.X, CenterY + Tmp1.Y), 5, vbCyan
    TheCanavas.Circle (CenterX + Tmp1.X, CenterY + Tmp1.Y), 3, vbCyan
    TheCanavas.Circle (CenterX + Tmp2.X, CenterY + Tmp2.Y), 3, vbCyan
    TheCanavas.Circle (CenterX + Tmp2.X, CenterY + Tmp2.Y), 1, vbCyan
    TheCanavas.DrawStyle = 1
    TheCanavas.Line (CenterX + Tmp1.X, CenterY + Tmp1.Y)-(CenterX + Tmp2.X, CenterY + Tmp2.Y), vbCyan
    TheCanavas.DrawStyle = 0
   End If
  End If
 Next CurCamera

 '//////////////////////////// PHOTON MAP ////////////////////////////////

Jump24:
 If ((DisplayPhotonmap = False) Or (EnablePhotonMapping = False)) Then Exit Sub
 If (UBound(ThePhotonMap()) <> MaximumAllocatedPhotons) Then Exit Sub
 For CurPhoton = 0 To PhotonsSentCount
  Tmp1 = ThePhotonMap(CurPhoton).Position: Tmp2 = ThePhotonMap(CurPhoton).Direction
  Tmp2 = VectorAdd(Tmp1, VectorScale(VectorNormalize(VectorSubtract(Tmp2, Tmp1)), Wire_PhotonTo))
  If ((Tmp1.Z > 0) And (Tmp2.Z > 0)) Then
   Tmp1.X = (Tmp1.X * Wire_ParallalScale): Tmp1.Y = (Tmp1.Y * Wire_ParallalScale)
   Tmp2.X = (Tmp2.X * Wire_ParallalScale): Tmp2.Y = (Tmp2.Y * Wire_ParallalScale)
   TheCanavas.Line (CenterX + Tmp1.X, CenterY + Tmp1.Y)-(CenterX + Tmp2.X, CenterY + Tmp2.Y), vbGreen
  End If
 Next CurPhoton

End Sub
