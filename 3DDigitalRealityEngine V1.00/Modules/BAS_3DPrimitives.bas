Attribute VB_Name = "BAS_3DPrimitives"

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
'###  MODULE      : BAS_Primitives.BAS
'###
'###  DESCRIPTION : Includ the Primitives creation functions, also the .X
'###                files importation function.
'###
'##################################################################################
'##################################################################################

'BUG: MAPPING V TEXTURE-COORDINATE IN Primitive MAKE-FROM-SPLINES function

Option Explicit
Function ImportMeshFromXFile(TheFileName As String, UniformScale!) As Long

 ' Importation from Microsoft(R) DirectX(R) files
 ' ==============================================
 '
 ' Uses DirectX7(R) library file (DX7VB.DLL) to load the geometry.
 ' Works only with triangular meshs, and the textures are external.
 ' This is the only thing that we do with this library, instead of writing
 ' a naive parser for loading the X files.

 ImportMeshFromXFile = -1

 If ((FileExist(TheFileName) = False) Or (UniformScale <= 0)) Then Exit Function

 'DirectX(R) objects:
 Dim TheDirectX As New DirectX7
 Dim TheD3D As Direct3DRM3
 Dim TheMeshBuilder As Direct3DRMMeshBuilder3
 Dim TheDXFace As Direct3DRMFace2
 Dim VerticesCount&, FacesCount&, CurVertex&, CurFace&, Bad As Boolean
 Dim CurU!, CurV!, VerticesStart&, FacesStart&, TmpVec As D3DVECTOR
 Dim CurU1!, CurV1!, CurU2!, CurV2!, CurU3!, CurV3!
 Dim LoadWithoutTextureCoords As Boolean

 'Create DirectX'(R) objects:
 Set TheD3D = TheDirectX.Direct3DRMCreate
 Set TheMeshBuilder = TheD3D.CreateMeshBuilder()

 On Error Resume Next

 'Load the .X file:
 TheMeshBuilder.LoadFromFile TheFileName, 0, D3DRMLOAD_FROMFILE, Nothing, Nothing

 If (Err.Number <> 0) Then MsgBox "Invalid DirectX(R) file !", vbCritical, "Bad file": Exit Function

 VerticesCount = TheMeshBuilder.GetVertexCount
 FacesCount = TheMeshBuilder.GetFaceCount

 'A validation check: vertices & faces counts, triangular mesh & valid texture coordinates
 If ((VerticesCount < 3) Or (FacesCount < 1)) Then MsgBox "Invalid mesh !", vbCritical: GoTo DoUnload
 For CurFace = 0 To (FacesCount - 1)
  Set TheDXFace = TheMeshBuilder.GetFace(CurFace)
  If (TheDXFace.GetVertexCount <> 3) Then
   MsgBox "No triangular geometry, c'ant load file !", vbCritical
   Bad = True: Exit For
  End If
  For CurVertex = 0 To (TheDXFace.GetVertexCount - 1)
   TheDXFace.GetTextureCoordinates CurVertex, CurU, CurV
   If ((CurU < 0) Or (CurU > 1) Or (CurV < 0) Or (CurV > 1)) Then
    If (MsgBox("Invalid texture coordinates, continue the loading any way ?", (vbInformation + vbYesNo), "Bad texture corrdinates") = vbYes) Then
     LoadWithoutTextureCoords = True: Exit For
    Else
     Bad = True: Exit For
    End If
   End If
  Next CurVertex
  If (Bad = True) Then Exit For
  If (LoadWithoutTextureCoords = True) Then Exit For
 Next CurFace
 If (Bad = True) Then Bad = False: GoTo DoUnload

 ImportMeshFromXFile = Mesh3D_Add(VerticesCount, FacesCount)
 If (ImportMeshFromXFile = -1) Then GoTo DoUnload

 TheMeshs(ImportMeshFromXFile).Label = "XFile_" & CStr(TheMeshsCount)
 VerticesStart = TheMeshs(ImportMeshFromXFile).Vertices.Start
 FacesStart = TheMeshs(ImportMeshFromXFile).Faces.Start

 'Update vertices:
 For CurVertex = 0 To (VerticesCount - 1)
  TheMeshBuilder.GetVertex CurVertex, TmpVec
  TmpVec.X = (TmpVec.X * UniformScale)
  TmpVec.Y = (TmpVec.Y * UniformScale)
  TmpVec.Z = (TmpVec.Z * UniformScale)
  TheVertices(VerticesStart + CurVertex).Position.X = TmpVec.X
  TheVertices(VerticesStart + CurVertex).Position.Y = TmpVec.Y
  TheVertices(VerticesStart + CurVertex).Position.Z = TmpVec.Z
 Next CurVertex

 'Update faces:
 For CurFace = 0 To (FacesCount - 1)
  Set TheDXFace = TheMeshBuilder.GetFace(CurFace)
  TheFaces(FacesStart + CurFace).A = TheDXFace.GetVertexIndex(0)
  TheFaces(FacesStart + CurFace).B = TheDXFace.GetVertexIndex(1)
  TheFaces(FacesStart + CurFace).C = TheDXFace.GetVertexIndex(2)
  If (LoadWithoutTextureCoords = False) Then
   TheDXFace.GetTextureCoordinates 0, CurU1, CurV1
   TheDXFace.GetTextureCoordinates 1, CurU2, CurV2
   TheDXFace.GetTextureCoordinates 2, CurU3, CurV3
   TheFaces(FacesStart + CurFace).AlphaVectors.U1 = CurU1
   TheFaces(FacesStart + CurFace).AlphaVectors.V1 = CurV1
   TheFaces(FacesStart + CurFace).AlphaVectors.U2 = CurU2
   TheFaces(FacesStart + CurFace).AlphaVectors.V2 = CurV2
   TheFaces(FacesStart + CurFace).AlphaVectors.U3 = CurU3
   TheFaces(FacesStart + CurFace).AlphaVectors.V3 = CurV3
   TheFaces(FacesStart + CurFace).ColorVectors.U1 = CurU1
   TheFaces(FacesStart + CurFace).ColorVectors.V1 = CurV1
   TheFaces(FacesStart + CurFace).ColorVectors.U2 = CurU2
   TheFaces(FacesStart + CurFace).ColorVectors.V2 = CurV2
   TheFaces(FacesStart + CurFace).ColorVectors.U3 = CurU3
   TheFaces(FacesStart + CurFace).ColorVectors.V3 = CurV3
   TheFaces(FacesStart + CurFace).ReflectionVectors.U1 = CurU1
   TheFaces(FacesStart + CurFace).ReflectionVectors.V1 = CurV1
   TheFaces(FacesStart + CurFace).ReflectionVectors.U2 = CurU2
   TheFaces(FacesStart + CurFace).ReflectionVectors.V2 = CurV2
   TheFaces(FacesStart + CurFace).ReflectionVectors.U3 = CurU3
   TheFaces(FacesStart + CurFace).ReflectionVectors.V3 = CurV3
   TheFaces(FacesStart + CurFace).RefractionVectors.U1 = CurU1
   TheFaces(FacesStart + CurFace).RefractionVectors.V1 = CurV1
   TheFaces(FacesStart + CurFace).RefractionVectors.U2 = CurU2
   TheFaces(FacesStart + CurFace).RefractionVectors.V2 = CurV2
   TheFaces(FacesStart + CurFace).RefractionVectors.U3 = CurU3
   TheFaces(FacesStart + CurFace).RefractionVectors.V3 = CurV3
   TheFaces(FacesStart + CurFace).RefractionNVectors.U1 = CurU1
   TheFaces(FacesStart + CurFace).RefractionNVectors.V1 = CurV1
   TheFaces(FacesStart + CurFace).RefractionNVectors.U2 = CurU2
   TheFaces(FacesStart + CurFace).RefractionNVectors.V2 = CurV2
   TheFaces(FacesStart + CurFace).RefractionNVectors.U3 = CurU3
   TheFaces(FacesStart + CurFace).RefractionNVectors.V3 = CurV3
  End If
 Next CurFace

DoUnload:
 'Unload DirectX(R) objects:
 Set TheMeshBuilder = Nothing
 Set TheD3D = Nothing
 Set TheDirectX = Nothing

End Function
Function Primitive_Box(CreateAsDefault As Boolean, Dimensions As Vector3D, DoubleSided As Boolean) As Long

 Primitive_Box = -1

 If (CreateAsDefault = False) Then
  If (VectorLength(Dimensions) = 0) Then Exit Function
 ElseIf (CreateAsDefault = True) Then
  Dimensions = VectorInput(50, 50, 50)
  DoubleSided = False
 End If

 Dim CurVertex&, CurFace&, CurF&

 If (DoubleSided = True) Then
  Primitive_Box = Mesh3D_Add(8, 24)
 Else
  Primitive_Box = Mesh3D_Add(8, 12)
 End If

 If (Primitive_Box = -1) Then Exit Function

 TheMeshs(Primitive_Box).Label = "Box_" & CStr(TheMeshsCount)
 CurVertex = TheMeshs(Primitive_Box).Vertices.Start
 CurFace = TheMeshs(Primitive_Box).Faces.Start

 'Add vertices
 '============
 TheVertices(CurVertex).Position = VectorInput(-Dimensions.X, -Dimensions.Y, -Dimensions.Z)
 TheVertices(CurVertex + 1).Position = VectorInput(Dimensions.X, -Dimensions.Y, -Dimensions.Z)
 TheVertices(CurVertex + 2).Position = VectorInput(Dimensions.X, Dimensions.Y, -Dimensions.Z)
 TheVertices(CurVertex + 3).Position = VectorInput(-Dimensions.X, Dimensions.Y, -Dimensions.Z)
 TheVertices(CurVertex + 4).Position = VectorInput(-Dimensions.X, -Dimensions.Y, Dimensions.Z)
 TheVertices(CurVertex + 5).Position = VectorInput(Dimensions.X, -Dimensions.Y, Dimensions.Z)
 TheVertices(CurVertex + 6).Position = VectorInput(Dimensions.X, Dimensions.Y, Dimensions.Z)
 TheVertices(CurVertex + 7).Position = VectorInput(-Dimensions.X, Dimensions.Y, Dimensions.Z)

 'Add faces
 '=========
 TheFaces(CurFace).A = (CurVertex + 0): TheFaces(CurFace).B = (CurVertex + 1): TheFaces(CurFace).C = (CurVertex + 3)
 TheFaces(CurFace + 1).A = (CurVertex + 1): TheFaces(CurFace + 1).B = (CurVertex + 2): TheFaces(CurFace + 1).C = (CurVertex + 3)
 TheFaces(CurFace + 2).A = (CurVertex + 1): TheFaces(CurFace + 2).B = (CurVertex + 5): TheFaces(CurFace + 2).C = (CurVertex + 2)
 TheFaces(CurFace + 3).A = (CurVertex + 5): TheFaces(CurFace + 3).B = (CurVertex + 6): TheFaces(CurFace + 3).C = (CurVertex + 2)
 TheFaces(CurFace + 4).A = (CurVertex + 5): TheFaces(CurFace + 4).B = (CurVertex + 4): TheFaces(CurFace + 4).C = (CurVertex + 6)
 TheFaces(CurFace + 5).A = (CurVertex + 4): TheFaces(CurFace + 5).B = (CurVertex + 7): TheFaces(CurFace + 5).C = (CurVertex + 6)
 TheFaces(CurFace + 6).A = (CurVertex + 4): TheFaces(CurFace + 6).B = (CurVertex + 0): TheFaces(CurFace + 6).C = (CurVertex + 7)
 TheFaces(CurFace + 7).A = (CurVertex + 0): TheFaces(CurFace + 7).B = (CurVertex + 3): TheFaces(CurFace + 7).C = (CurVertex + 7)
 TheFaces(CurFace + 8).A = (CurVertex + 4): TheFaces(CurFace + 8).B = (CurVertex + 5): TheFaces(CurFace + 8).C = (CurVertex + 0)
 TheFaces(CurFace + 9).A = (CurVertex + 5): TheFaces(CurFace + 9).B = (CurVertex + 1): TheFaces(CurFace + 9).C = (CurVertex + 0)
 TheFaces(CurFace + 10).A = (CurVertex + 3): TheFaces(CurFace + 10).B = (CurVertex + 2): TheFaces(CurFace + 10).C = (CurVertex + 7)
 TheFaces(CurFace + 11).A = (CurVertex + 2): TheFaces(CurFace + 11).B = (CurVertex + 6): TheFaces(CurFace + 11).C = (CurVertex + 7)

 'Setup texture coordinates
 '=========================
 For CurF = 0 To 11 Step 2
  'Alpha texture-coordinates:
  TheFaces(CurFace + CurF).AlphaVectors.U1 = 0: TheFaces(CurFace + CurF).AlphaVectors.V1 = 0
  TheFaces(CurFace + CurF).AlphaVectors.U2 = 1: TheFaces(CurFace + CurF).AlphaVectors.V2 = 0
  TheFaces(CurFace + CurF).AlphaVectors.U3 = 0: TheFaces(CurFace + CurF).AlphaVectors.V3 = 1
  TheFaces(CurFace + (CurF + 1)).AlphaVectors.U1 = 1: TheFaces(CurFace + (CurF + 1)).AlphaVectors.V1 = 0
  TheFaces(CurFace + (CurF + 1)).AlphaVectors.U2 = 1: TheFaces(CurFace + (CurF + 1)).AlphaVectors.V2 = 1
  TheFaces(CurFace + (CurF + 1)).AlphaVectors.U3 = 0: TheFaces(CurFace + (CurF + 1)).AlphaVectors.V3 = 1
  'Color texture-coordinates:
  TheFaces(CurFace + CurF).ColorVectors.U1 = 0: TheFaces(CurFace + CurF).ColorVectors.V1 = 0
  TheFaces(CurFace + CurF).ColorVectors.U2 = 1: TheFaces(CurFace + CurF).ColorVectors.V2 = 0
  TheFaces(CurFace + CurF).ColorVectors.U3 = 0: TheFaces(CurFace + CurF).ColorVectors.V3 = 1
  TheFaces(CurFace + (CurF + 1)).ColorVectors.U1 = 1: TheFaces(CurFace + (CurF + 1)).ColorVectors.V1 = 0
  TheFaces(CurFace + (CurF + 1)).ColorVectors.U2 = 1: TheFaces(CurFace + (CurF + 1)).ColorVectors.V2 = 1
  TheFaces(CurFace + (CurF + 1)).ColorVectors.U3 = 0: TheFaces(CurFace + (CurF + 1)).ColorVectors.V3 = 1
  'Reflection texture-coordinates:
  TheFaces(CurFace + CurF).ReflectionVectors.U1 = 0: TheFaces(CurFace + CurF).ReflectionVectors.V1 = 0
  TheFaces(CurFace + CurF).ReflectionVectors.U2 = 1: TheFaces(CurFace + CurF).ReflectionVectors.V2 = 0
  TheFaces(CurFace + CurF).ReflectionVectors.U3 = 0: TheFaces(CurFace + CurF).ReflectionVectors.V3 = 1
  TheFaces(CurFace + (CurF + 1)).ReflectionVectors.U1 = 1: TheFaces(CurFace + (CurF + 1)).ReflectionVectors.V1 = 0
  TheFaces(CurFace + (CurF + 1)).ReflectionVectors.U2 = 1: TheFaces(CurFace + (CurF + 1)).ReflectionVectors.V2 = 1
  TheFaces(CurFace + (CurF + 1)).ReflectionVectors.U3 = 0: TheFaces(CurFace + (CurF + 1)).ReflectionVectors.V3 = 1
  'Refraction texture-coordinates:
  TheFaces(CurFace + CurF).RefractionVectors.U1 = 0: TheFaces(CurFace + CurF).RefractionVectors.V1 = 0
  TheFaces(CurFace + CurF).RefractionVectors.U2 = 1: TheFaces(CurFace + CurF).RefractionVectors.V2 = 0
  TheFaces(CurFace + CurF).RefractionVectors.U3 = 0: TheFaces(CurFace + CurF).RefractionVectors.V3 = 1
  TheFaces(CurFace + (CurF + 1)).RefractionVectors.U1 = 1: TheFaces(CurFace + (CurF + 1)).RefractionVectors.V1 = 0
  TheFaces(CurFace + (CurF + 1)).RefractionVectors.U2 = 1: TheFaces(CurFace + (CurF + 1)).RefractionVectors.V2 = 1
  TheFaces(CurFace + (CurF + 1)).RefractionVectors.U3 = 0: TheFaces(CurFace + (CurF + 1)).RefractionVectors.V3 = 1
  'RefractionN texture-coordinates:
  TheFaces(CurFace + CurF).RefractionNVectors.U1 = 0: TheFaces(CurFace + CurF).RefractionNVectors.V1 = 0
  TheFaces(CurFace + CurF).RefractionNVectors.U2 = 1: TheFaces(CurFace + CurF).RefractionNVectors.V2 = 0
  TheFaces(CurFace + CurF).RefractionNVectors.U3 = 0: TheFaces(CurFace + CurF).RefractionNVectors.V3 = 1
  TheFaces(CurFace + (CurF + 1)).RefractionNVectors.U1 = 1: TheFaces(CurFace + (CurF + 1)).RefractionNVectors.V1 = 0
  TheFaces(CurFace + (CurF + 1)).RefractionNVectors.U2 = 1: TheFaces(CurFace + (CurF + 1)).RefractionNVectors.V2 = 1
  TheFaces(CurFace + (CurF + 1)).RefractionNVectors.U3 = 0: TheFaces(CurFace + (CurF + 1)).RefractionNVectors.V3 = 1
 Next CurF

 If (DoubleSided = True) Then
  TheFaces(CurFace + 12).A = (CurVertex + 0): TheFaces(CurFace + 12).C = (CurVertex + 1): TheFaces(CurFace + 12).B = (CurVertex + 3)
  TheFaces(CurFace + 13).A = (CurVertex + 1): TheFaces(CurFace + 13).C = (CurVertex + 2): TheFaces(CurFace + 13).B = (CurVertex + 3)
  TheFaces(CurFace + 14).A = (CurVertex + 1): TheFaces(CurFace + 14).C = (CurVertex + 5): TheFaces(CurFace + 14).B = (CurVertex + 2)
  TheFaces(CurFace + 15).A = (CurVertex + 5): TheFaces(CurFace + 15).C = (CurVertex + 6): TheFaces(CurFace + 15).B = (CurVertex + 2)
  TheFaces(CurFace + 16).A = (CurVertex + 5): TheFaces(CurFace + 16).C = (CurVertex + 4): TheFaces(CurFace + 16).B = (CurVertex + 6)
  TheFaces(CurFace + 17).A = (CurVertex + 4): TheFaces(CurFace + 17).C = (CurVertex + 7): TheFaces(CurFace + 17).B = (CurVertex + 6)
  TheFaces(CurFace + 18).A = (CurVertex + 4): TheFaces(CurFace + 18).C = (CurVertex + 0): TheFaces(CurFace + 18).B = (CurVertex + 7)
  TheFaces(CurFace + 19).A = (CurVertex + 0): TheFaces(CurFace + 19).C = (CurVertex + 3): TheFaces(CurFace + 19).B = (CurVertex + 7)
  TheFaces(CurFace + 20).A = (CurVertex + 4): TheFaces(CurFace + 20).C = (CurVertex + 5): TheFaces(CurFace + 20).B = (CurVertex + 0)
  TheFaces(CurFace + 21).A = (CurVertex + 5): TheFaces(CurFace + 21).C = (CurVertex + 1): TheFaces(CurFace + 21).B = (CurVertex + 0)
  TheFaces(CurFace + 22).A = (CurVertex + 3): TheFaces(CurFace + 22).C = (CurVertex + 2): TheFaces(CurFace + 22).B = (CurVertex + 7)
  TheFaces(CurFace + 23).A = (CurVertex + 2): TheFaces(CurFace + 23).C = (CurVertex + 6): TheFaces(CurFace + 23).B = (CurVertex + 7)
  'Setup texture coordinates:
  For CurF = 12 To 23 Step 2
   'Alpha texture-coordinates:
   TheFaces(CurFace + CurF).AlphaVectors.U1 = 0: TheFaces(CurFace + CurF).AlphaVectors.V1 = 0
   TheFaces(CurFace + CurF).AlphaVectors.U2 = 1: TheFaces(CurFace + CurF).AlphaVectors.V2 = 0
   TheFaces(CurFace + CurF).AlphaVectors.U3 = 0: TheFaces(CurFace + CurF).AlphaVectors.V3 = 1
   TheFaces(CurFace + (CurF + 1)).AlphaVectors.U1 = 1: TheFaces(CurFace + (CurF + 1)).AlphaVectors.V1 = 0
   TheFaces(CurFace + (CurF + 1)).AlphaVectors.U2 = 1: TheFaces(CurFace + (CurF + 1)).AlphaVectors.V2 = 1
   TheFaces(CurFace + (CurF + 1)).AlphaVectors.U3 = 0: TheFaces(CurFace + (CurF + 1)).AlphaVectors.V3 = 1
   'Color texture-coordinates:
   TheFaces(CurFace + CurF).ColorVectors.U1 = 0: TheFaces(CurFace + CurF).ColorVectors.V1 = 0
   TheFaces(CurFace + CurF).ColorVectors.U2 = 1: TheFaces(CurFace + CurF).ColorVectors.V2 = 0
   TheFaces(CurFace + CurF).ColorVectors.U3 = 0: TheFaces(CurFace + CurF).ColorVectors.V3 = 1
   TheFaces(CurFace + (CurF + 1)).ColorVectors.U1 = 1: TheFaces(CurFace + (CurF + 1)).ColorVectors.V1 = 0
   TheFaces(CurFace + (CurF + 1)).ColorVectors.U2 = 1: TheFaces(CurFace + (CurF + 1)).ColorVectors.V2 = 1
   TheFaces(CurFace + (CurF + 1)).ColorVectors.U3 = 0: TheFaces(CurFace + (CurF + 1)).ColorVectors.V3 = 1
   'Reflection texture-coordinates:
   TheFaces(CurFace + CurF).ReflectionVectors.U1 = 0: TheFaces(CurFace + CurF).ReflectionVectors.V1 = 0
   TheFaces(CurFace + CurF).ReflectionVectors.U2 = 1: TheFaces(CurFace + CurF).ReflectionVectors.V2 = 0
   TheFaces(CurFace + CurF).ReflectionVectors.U3 = 0: TheFaces(CurFace + CurF).ReflectionVectors.V3 = 1
   TheFaces(CurFace + (CurF + 1)).ReflectionVectors.U1 = 1: TheFaces(CurFace + (CurF + 1)).ReflectionVectors.V1 = 0
   TheFaces(CurFace + (CurF + 1)).ReflectionVectors.U2 = 1: TheFaces(CurFace + (CurF + 1)).ReflectionVectors.V2 = 1
   TheFaces(CurFace + (CurF + 1)).ReflectionVectors.U3 = 0: TheFaces(CurFace + (CurF + 1)).ReflectionVectors.V3 = 1
   'Refraction texture-coordinates:
   TheFaces(CurFace + CurF).RefractionVectors.U1 = 0: TheFaces(CurFace + CurF).RefractionVectors.V1 = 0
   TheFaces(CurFace + CurF).RefractionVectors.U2 = 1: TheFaces(CurFace + CurF).RefractionVectors.V2 = 0
   TheFaces(CurFace + CurF).RefractionVectors.U3 = 0: TheFaces(CurFace + CurF).RefractionVectors.V3 = 1
   TheFaces(CurFace + (CurF + 1)).RefractionVectors.U1 = 1: TheFaces(CurFace + (CurF + 1)).RefractionVectors.V1 = 0
   TheFaces(CurFace + (CurF + 1)).RefractionVectors.U2 = 1: TheFaces(CurFace + (CurF + 1)).RefractionVectors.V2 = 1
   TheFaces(CurFace + (CurF + 1)).RefractionVectors.U3 = 0: TheFaces(CurFace + (CurF + 1)).RefractionVectors.V3 = 1
   'RefractionN texture-coordinates:
   TheFaces(CurFace + CurF).RefractionNVectors.U1 = 0: TheFaces(CurFace + CurF).RefractionNVectors.V1 = 0
   TheFaces(CurFace + CurF).RefractionNVectors.U2 = 1: TheFaces(CurFace + CurF).RefractionNVectors.V2 = 0
   TheFaces(CurFace + CurF).RefractionNVectors.U3 = 0: TheFaces(CurFace + CurF).RefractionNVectors.V3 = 1
   TheFaces(CurFace + (CurF + 1)).RefractionNVectors.U1 = 1: TheFaces(CurFace + (CurF + 1)).RefractionNVectors.V1 = 0
   TheFaces(CurFace + (CurF + 1)).RefractionNVectors.U2 = 1: TheFaces(CurFace + (CurF + 1)).RefractionNVectors.V2 = 1
   TheFaces(CurFace + (CurF + 1)).RefractionNVectors.U3 = 0: TheFaces(CurFace + (CurF + 1)).RefractionNVectors.V3 = 1
  Next CurF
 End If

End Function
Function Primitive_Cone(CreateAsDefault As Boolean, Base!, Radius!, Steps&, Axis As Byte, DoubleSided As Boolean) As Long

 Primitive_Cone = -1

 If (CreateAsDefault = False) Then
  If ((Base <= 0) Or (Radius <= 0)) Then Exit Function
  If ((Steps < 2) Or (Axis < 0) Or (Axis > 2)) Then Exit Function
 ElseIf (CreateAsDefault = True) Then
  Base = 100: Radius = 50: Steps = 10:  Axis = 1: DoubleSided = False
 End If

 Dim AShapeSpline(2) As Vector3D, AMoveSpline() As Vector3D, CurVertex&, Tmp!

 AShapeSpline(0).Z = -(Base * 0.5)
 AShapeSpline(1).X = Radius: AShapeSpline(1).Z = -AShapeSpline(0).Z
 AShapeSpline(2).Z = AShapeSpline(1).Z

 Spline_Circle AMoveSpline(), ApproachVal, 0, 0, 2, Steps
 ReDim Preserve AMoveSpline(UBound(AMoveSpline) + 1)
 AMoveSpline(UBound(AMoveSpline)) = AMoveSpline(1)

 Primitive_Cone = Primitive_MakeFromSplines(AShapeSpline(), AMoveSpline(), 0.5, 0, 0, 1, 1, 1, 1, DoubleSided, True)
 If (Primitive_Cone = -1) Then Exit Function

 TheMeshs(Primitive_Cone).Label = "Cone_" & CStr(TheMeshsCount)
 Select Case Axis
  Case 0:
   For CurVertex = TheMeshs(Primitive_Cone).Vertices.Start To GetAddressLast(TheMeshs(Primitive_Cone).Vertices)
    'Swap X to Z:
    Tmp = TheVertices(CurVertex).Position.X
    TheVertices(CurVertex).Position.X = TheVertices(CurVertex).Position.Z
    TheVertices(CurVertex).Position.Z = Tmp
    'Swap Z to Y:
    Tmp = TheVertices(CurVertex).Position.Z
    TheVertices(CurVertex).Position.Z = TheVertices(CurVertex).Position.Y
    TheVertices(CurVertex).Position.Y = Tmp
   Next CurVertex
  Case 1:
   For CurVertex = TheMeshs(Primitive_Cone).Vertices.Start To GetAddressLast(TheMeshs(Primitive_Cone).Vertices)
    'Swap Z to Y:
    Tmp = TheVertices(CurVertex).Position.Z
    TheVertices(CurVertex).Position.Z = TheVertices(CurVertex).Position.Y
    TheVertices(CurVertex).Position.Y = Tmp
   Next CurVertex
 End Select

End Function
Function Primitive_CornellBox(CreateAsDefault As Boolean, Dimensions As Vector3D, Subdivisions As Long, DoubleSided As Boolean)

 If (CreateAsDefault = False) Then
  If ((VectorLength(Dimensions) = 0) Or (Subdivisions < 2)) Then Exit Function
 ElseIf (CreateAsDefault = True) Then
  Dimensions = VectorInput(50, 50, 50)
  Subdivisions = 2
  DoubleSided = False
 End If

 Dim PlaneIndex As Long

 'Front plane:
 PlaneIndex = Primitive_Grid(False, (Dimensions.X * 2), (Dimensions.Y * 2), Subdivisions, 2, DoubleSided)
 If (PlaneIndex = -1) Then Exit Function
 TheMeshs(PlaneIndex).Position.Z = -Dimensions.Z
 TheMaterials(PlaneIndex).Color = ColorWhite

 'Left plane:
 PlaneIndex = Primitive_Grid(False, (Dimensions.Z * 2), (Dimensions.Y * 2), Subdivisions, 0, DoubleSided)
 If (PlaneIndex = -1) Then Exit Function
 TheMeshs(PlaneIndex).Position.X = -Dimensions.X
 TheMaterials(PlaneIndex).Color = ColorRed

 'Right plane:
 PlaneIndex = Primitive_Grid(False, (Dimensions.Z * 2), (Dimensions.Y * 2), Subdivisions, 0, DoubleSided)
 If (PlaneIndex = -1) Then Exit Function
 TheMeshs(PlaneIndex).Position.X = Dimensions.X
 TheMaterials(PlaneIndex).Color = ColorBlue

 'Top plane:
 PlaneIndex = Primitive_Grid(False, (Dimensions.X * 2), (Dimensions.Z * 2), Subdivisions, 1, DoubleSided)
 If (PlaneIndex = -1) Then Exit Function
 TheMeshs(PlaneIndex).Position.Y = -Dimensions.Y
 TheMaterials(PlaneIndex).Color = ColorWhite

 'Bottom plane:
 PlaneIndex = Primitive_Grid(False, (Dimensions.X * 2), (Dimensions.Z * 2), Subdivisions, 1, DoubleSided)
 If (PlaneIndex = -1) Then Exit Function
 TheMeshs(PlaneIndex).Position.Y = Dimensions.Y
 TheMaterials(PlaneIndex).Color = ColorWhite

End Function
Function Primitive_Pyramid(CreateAsDefault As Boolean, Base!, Radius!, Axis As Byte, DoubleSided As Boolean) As Long

 Primitive_Pyramid = -1

 If (CreateAsDefault = False) Then
  If ((Base <= 0) Or (Radius <= 0) Or (Axis < 0) Or (Axis > 2)) Then Exit Function
 ElseIf (CreateAsDefault = True) Then
  Base = 100: Radius = 100: Axis = 1: DoubleSided = False
 End If

 Primitive_Pyramid = Primitive_Cone(False, Base, Radius, 4, Axis, DoubleSided)
 If (Primitive_Pyramid <> -1) Then TheMeshs(Primitive_Pyramid).Label = "Pyramid_" & CStr(TheMeshsCount)

End Function
Function Primitive_Cylinder(CreateAsDefault As Boolean, Base!, Radius!, Steps&, Axis As Byte, DoubleSided As Boolean) As Long

 Primitive_Cylinder = -1

 If (CreateAsDefault = False) Then
  If ((Base <= 0) Or (Radius <= 0)) Then Exit Function
  If ((Steps < 2) Or (Axis < 0) Or (Axis > 2)) Then Exit Function
 ElseIf (CreateAsDefault = True) Then
  Base = 100: Radius = 50: Steps = 10:  Axis = 1: DoubleSided = False
 End If

 Dim AShapeSpline(3) As Vector3D, AMoveSpline() As Vector3D, CurVertex&, Tmp!

 AShapeSpline(0).Z = -(Base * 0.5)
 AShapeSpline(1).X = Radius: AShapeSpline(1).Z = AShapeSpline(0).Z
 AShapeSpline(2).X = Radius: AShapeSpline(2).Z = -AShapeSpline(1).Z
 AShapeSpline(3).Z = AShapeSpline(2).Z

 Spline_Circle AMoveSpline(), ApproachVal, 0, 0, 2, Steps
 ReDim Preserve AMoveSpline(UBound(AMoveSpline) + 1)
 AMoveSpline(UBound(AMoveSpline)) = AMoveSpline(1)

 Primitive_Cylinder = Primitive_MakeFromSplines(AShapeSpline(), AMoveSpline(), 0.5, 0, 0, 1, 1, 1, 1, DoubleSided, True)
 If (Primitive_Cylinder = -1) Then Exit Function

 TheMeshs(Primitive_Cylinder).Label = "Cylinder_" & CStr(TheMeshsCount)
 Select Case Axis
  Case 0:
   For CurVertex = TheMeshs(Primitive_Cylinder).Vertices.Start To GetAddressLast(TheMeshs(Primitive_Cylinder).Vertices)
    'Swap X to Z:
    Tmp = TheVertices(CurVertex).Position.X
    TheVertices(CurVertex).Position.X = TheVertices(CurVertex).Position.Z
    TheVertices(CurVertex).Position.Z = Tmp
    'Swap Z to Y:
    Tmp = TheVertices(CurVertex).Position.Z
    TheVertices(CurVertex).Position.Z = TheVertices(CurVertex).Position.Y
    TheVertices(CurVertex).Position.Y = Tmp
   Next CurVertex
  Case 1:
   For CurVertex = TheMeshs(Primitive_Cylinder).Vertices.Start To GetAddressLast(TheMeshs(Primitive_Cylinder).Vertices)
    'Swap Z to Y:
    Tmp = TheVertices(CurVertex).Position.Z
    TheVertices(CurVertex).Position.Z = TheVertices(CurVertex).Position.Y
    TheVertices(CurVertex).Position.Y = Tmp
   Next CurVertex
 End Select

End Function
Function Primitive_Disk(CreateAsDefault As Boolean, Radius1!, Radius2!, Steps&, Axis As Byte, DoubleSided As Boolean) As Long

 Primitive_Disk = -1

 If (CreateAsDefault = False) Then
  If ((Radius1 < 0) Or (Radius1 > Radius2)) Then Exit Function
  If ((Steps < 2) Or (Axis < 0) Or (Axis > 2)) Then Exit Function
 ElseIf (CreateAsDefault = True) Then
  Radius1 = 50: Radius2 = 100: Steps = 10:  Axis = 1: DoubleSided = False
 End If

 Dim AShapeSpline(1) As Vector3D, AMoveSpline() As Vector3D, CurVertex&, Tmp!

 AShapeSpline(0).X = Radius1: AShapeSpline(1).X = Radius2

 Spline_Circle AMoveSpline(), ApproachVal, 0, 0, 2, Steps
 ReDim Preserve AMoveSpline(UBound(AMoveSpline) + 1)
 AMoveSpline(UBound(AMoveSpline)) = AMoveSpline(1)

 Primitive_Disk = Primitive_MakeFromSplines(AShapeSpline(), AMoveSpline(), 0.5, 0, 0, 1, 1, 1, 1, DoubleSided, True)
 If (Primitive_Disk = -1) Then Exit Function

 TheMeshs(Primitive_Disk).Label = "Disk_" & CStr(TheMeshsCount)
 Select Case Axis
  Case 0:
   For CurVertex = TheMeshs(Primitive_Disk).Vertices.Start To GetAddressLast(TheMeshs(Primitive_Disk).Vertices)
    'Swap X to Z:
    Tmp = TheVertices(CurVertex).Position.X
    TheVertices(CurVertex).Position.X = TheVertices(CurVertex).Position.Z
    TheVertices(CurVertex).Position.Z = Tmp
    'Swap Z to Y:
    Tmp = TheVertices(CurVertex).Position.Z
    TheVertices(CurVertex).Position.Z = TheVertices(CurVertex).Position.Y
    TheVertices(CurVertex).Position.Y = Tmp
   Next CurVertex
  Case 1:
   For CurVertex = TheMeshs(Primitive_Disk).Vertices.Start To GetAddressLast(TheMeshs(Primitive_Disk).Vertices)
    'Swap Z to Y:
    Tmp = TheVertices(CurVertex).Position.Z
    TheVertices(CurVertex).Position.Z = TheVertices(CurVertex).Position.Y
    TheVertices(CurVertex).Position.Y = Tmp
   Next CurVertex
 End Select

End Function
Function Primitive_Grid(CreateAsDefault As Boolean, Width!, Height!, Steps&, Axis As Byte, DoubleSided As Boolean) As Long

 Primitive_Grid = -1

 If (CreateAsDefault = False) Then
  If ((Width <= 0) Or (Height <= 0)) Then Exit Function
  If ((Steps < 2) Or (Axis < 0) Or (Axis > 2)) Then Exit Function
 ElseIf (CreateAsDefault = True) Then
  Width = 100: Height = 100: Steps = 10: Axis = 2: DoubleSided = False
 End If

 Dim AShapeSpline() As Vector3D, AMoveSpline() As Vector3D, CurVertex&, Tmp!

 ReDim AShapeSpline(Steps): ReDim AMoveSpline(Steps + 1)

 For CurVertex = 0 To Steps
  AShapeSpline(CurVertex) = VectorInterpolate(VectorInput(-(Height * 0.5), 0, 0), VectorInput((Height * 0.5), 0, 0), (CurVertex / Steps))
  AMoveSpline(CurVertex) = VectorInterpolate(VectorInput(0, -(Width * 0.5), 0), VectorInput(0, (Width * 0.5), 0), (CurVertex / Steps))
 Next CurVertex
 AMoveSpline(UBound(AMoveSpline())).Y = (AMoveSpline(UBound(AMoveSpline()) - 1).Y - (AMoveSpline(UBound(AMoveSpline()) - 2).Y - AMoveSpline(UBound(AMoveSpline()) - 1).Y))

 Primitive_Grid = Primitive_MakeFromSplines(AShapeSpline(), AMoveSpline(), 1, 0, 0, 1, 1, 1, 1, DoubleSided, False)
 If (Primitive_Grid = -1) Then Exit Function

 TheMeshs(Primitive_Grid).Label = "Grid_" & CStr(TheMeshsCount)
 Select Case Axis
  Case 0:
   For CurVertex = TheMeshs(Primitive_Grid).Vertices.Start To GetAddressLast(TheMeshs(Primitive_Grid).Vertices)
    'Swap X to Z:
    Tmp = TheVertices(CurVertex).Position.X
    TheVertices(CurVertex).Position.X = TheVertices(CurVertex).Position.Z
    TheVertices(CurVertex).Position.Z = Tmp
    'Swap Z to Y:
    Tmp = TheVertices(CurVertex).Position.Z
    TheVertices(CurVertex).Position.Z = TheVertices(CurVertex).Position.Y
    TheVertices(CurVertex).Position.Y = Tmp
   Next CurVertex
  Case 1:
   For CurVertex = TheMeshs(Primitive_Grid).Vertices.Start To GetAddressLast(TheMeshs(Primitive_Grid).Vertices)
    'Swap Z to Y:
    Tmp = TheVertices(CurVertex).Position.Z
    TheVertices(CurVertex).Position.Z = TheVertices(CurVertex).Position.Y
    TheVertices(CurVertex).Position.Y = Tmp
   Next CurVertex
 End Select

End Function
Function Primitive_Rectangle(CreateAsDefault As Boolean, Width!, Height!, Axis As Byte, DoubleSided As Boolean) As Long

 Primitive_Rectangle = -1

 If (CreateAsDefault = False) Then
  If ((Width <= 0) Or (Height <= 0)) Then Exit Function
  If ((Axis < 0) Or (Axis > 2)) Then Exit Function
 ElseIf (CreateAsDefault = True) Then
  Width = 100: Height = 100: Axis = 2: DoubleSided = False
 End If

 Dim CurVertex&, CurFace&, CurF&

 If (DoubleSided = True) Then
  Primitive_Rectangle = Mesh3D_Add(4, 4)
 Else
  Primitive_Rectangle = Mesh3D_Add(4, 2)
 End If

 If (Primitive_Rectangle = -1) Then Exit Function

 TheMeshs(Primitive_Rectangle).Label = "Rectangle_" & CStr(TheMeshsCount)
 CurVertex = TheMeshs(Primitive_Rectangle).Vertices.Start
 CurFace = TheMeshs(Primitive_Rectangle).Faces.Start

 Select Case Axis
  Case 0:
   TheVertices(CurVertex).Position = VectorInput(0, -Width, -Height)
   TheVertices(CurVertex + 1).Position = VectorInput(0, Width, -Height)
   TheVertices(CurVertex + 2).Position = VectorInput(0, Width, Height)
   TheVertices(CurVertex + 3).Position = VectorInput(0, -Width, Height)
  Case 1:
   TheVertices(CurVertex).Position = VectorInput(-Width, 0, -Height)
   TheVertices(CurVertex + 1).Position = VectorInput(Width, 0, -Height)
   TheVertices(CurVertex + 2).Position = VectorInput(Width, 0, Height)
   TheVertices(CurVertex + 3).Position = VectorInput(-Width, 0, Height)
  Case 2:
   TheVertices(CurVertex).Position = VectorInput(-Width, -Height, 0)
   TheVertices(CurVertex + 1).Position = VectorInput(Width, -Height, 0)
   TheVertices(CurVertex + 2).Position = VectorInput(Width, Height, 0)
   TheVertices(CurVertex + 3).Position = VectorInput(-Width, Height, 0)
 End Select

 TheFaces(CurFace).A = (CurVertex + 0): TheFaces(CurFace).B = (CurVertex + 1): TheFaces(CurFace).C = (CurVertex + 3)
 TheFaces(CurFace + 1).A = (CurVertex + 1): TheFaces(CurFace + 1).B = (CurVertex + 2): TheFaces(CurFace + 1).C = (CurVertex + 3)
 'Alpha texture-coordinates:
 TheFaces(CurFace).AlphaVectors.U1 = 0: TheFaces(CurFace).AlphaVectors.V1 = 0
 TheFaces(CurFace).AlphaVectors.U2 = 1: TheFaces(CurFace).AlphaVectors.V2 = 0
 TheFaces(CurFace).AlphaVectors.U3 = 0: TheFaces(CurFace).AlphaVectors.V3 = 1
 TheFaces(CurFace + 1).AlphaVectors.U1 = 1: TheFaces(CurFace + 1).AlphaVectors.V1 = 0
 TheFaces(CurFace + 1).AlphaVectors.U2 = 1: TheFaces(CurFace + 1).AlphaVectors.V2 = 1
 TheFaces(CurFace + 1).AlphaVectors.U3 = 0: TheFaces(CurFace + 1).AlphaVectors.V3 = 1
 'Color texture-coordinates:
 TheFaces(CurFace).ColorVectors.U1 = 0: TheFaces(CurFace).ColorVectors.V1 = 0
 TheFaces(CurFace).ColorVectors.U2 = 1: TheFaces(CurFace).ColorVectors.V2 = 0
 TheFaces(CurFace).ColorVectors.U3 = 0: TheFaces(CurFace).ColorVectors.V3 = 1
 TheFaces(CurFace + 1).ColorVectors.U1 = 1: TheFaces(CurFace + 1).ColorVectors.V1 = 0
 TheFaces(CurFace + 1).ColorVectors.U2 = 1: TheFaces(CurFace + 1).ColorVectors.V2 = 1
 TheFaces(CurFace + 1).ColorVectors.U3 = 0: TheFaces(CurFace + 1).ColorVectors.V3 = 1
 'Reflection texture-coordinates:
 TheFaces(CurFace).ReflectionVectors.U1 = 0: TheFaces(CurFace).ReflectionVectors.V1 = 0
 TheFaces(CurFace).ReflectionVectors.U2 = 1: TheFaces(CurFace).ReflectionVectors.V2 = 0
 TheFaces(CurFace).ReflectionVectors.U3 = 0: TheFaces(CurFace).ReflectionVectors.V3 = 1
 TheFaces(CurFace + 1).ReflectionVectors.U1 = 1: TheFaces(CurFace + 1).ReflectionVectors.V1 = 0
 TheFaces(CurFace + 1).ReflectionVectors.U2 = 1: TheFaces(CurFace + 1).ReflectionVectors.V2 = 1
 TheFaces(CurFace + 1).ReflectionVectors.U3 = 0: TheFaces(CurFace + 1).ReflectionVectors.V3 = 1
 'Refraction texture-coordinates:
 TheFaces(CurFace).RefractionVectors.U1 = 0: TheFaces(CurFace).RefractionVectors.V1 = 0
 TheFaces(CurFace).RefractionVectors.U2 = 1: TheFaces(CurFace).RefractionVectors.V2 = 0
 TheFaces(CurFace).RefractionVectors.U3 = 0: TheFaces(CurFace).RefractionVectors.V3 = 1
 TheFaces(CurFace + 1).RefractionVectors.U1 = 1: TheFaces(CurFace + 1).RefractionVectors.V1 = 0
 TheFaces(CurFace + 1).RefractionVectors.U2 = 1: TheFaces(CurFace + 1).RefractionVectors.V2 = 1
 TheFaces(CurFace + 1).RefractionVectors.U3 = 0: TheFaces(CurFace + 1).RefractionVectors.V3 = 1
 'RefractionN texture-coordinates:
 TheFaces(CurFace).RefractionNVectors.U1 = 0: TheFaces(CurFace).RefractionNVectors.V1 = 0
 TheFaces(CurFace).RefractionNVectors.U2 = 1: TheFaces(CurFace).RefractionNVectors.V2 = 0
 TheFaces(CurFace).RefractionNVectors.U3 = 0: TheFaces(CurFace).RefractionNVectors.V3 = 1
 TheFaces(CurFace + 1).RefractionNVectors.U1 = 1: TheFaces(CurFace + 1).RefractionNVectors.V1 = 0
 TheFaces(CurFace + 1).RefractionNVectors.U2 = 1: TheFaces(CurFace + 1).RefractionNVectors.V2 = 1
 TheFaces(CurFace + 1).RefractionNVectors.U3 = 0: TheFaces(CurFace + 1).RefractionNVectors.V3 = 1

 If (DoubleSided = True) Then
  TheFaces(CurFace + 2).A = (CurVertex + 0): TheFaces(CurFace + 12).C = (CurVertex + 1): TheFaces(CurFace + 12).B = (CurVertex + 3)
  TheFaces(CurFace + 3).A = (CurVertex + 1): TheFaces(CurFace + 13).C = (CurVertex + 2): TheFaces(CurFace + 13).B = (CurVertex + 3)
  'Alpha texture-coordinates:
  TheFaces(CurFace + 2).AlphaVectors.U1 = 0: TheFaces(CurFace + 2).AlphaVectors.V1 = 0
  TheFaces(CurFace + 2).AlphaVectors.U2 = 1: TheFaces(CurFace + 2).AlphaVectors.V2 = 0
  TheFaces(CurFace + 2).AlphaVectors.U3 = 0: TheFaces(CurFace + 2).AlphaVectors.V3 = 1
  TheFaces(CurFace + 3).AlphaVectors.U1 = 1: TheFaces(CurFace + 3).AlphaVectors.V1 = 0
  TheFaces(CurFace + 3).AlphaVectors.U2 = 1: TheFaces(CurFace + 3).AlphaVectors.V2 = 1
  TheFaces(CurFace + 3).AlphaVectors.U3 = 0: TheFaces(CurFace + 3).AlphaVectors.V3 = 1
  'Color texture-coordinates:
  TheFaces(CurFace + 2).ColorVectors.U1 = 0: TheFaces(CurFace + 2).ColorVectors.V1 = 0
  TheFaces(CurFace + 2).ColorVectors.U2 = 1: TheFaces(CurFace + 2).ColorVectors.V2 = 0
  TheFaces(CurFace + 2).ColorVectors.U3 = 0: TheFaces(CurFace + 2).ColorVectors.V3 = 1
  TheFaces(CurFace + 3).ColorVectors.U1 = 1: TheFaces(CurFace + 3).ColorVectors.V1 = 0
  TheFaces(CurFace + 3).ColorVectors.U2 = 1: TheFaces(CurFace + 3).ColorVectors.V2 = 1
  TheFaces(CurFace + 3).ColorVectors.U3 = 0: TheFaces(CurFace + 3).ColorVectors.V3 = 1
  'Reflection texture-coordinates:
  TheFaces(CurFace + 2).ReflectionVectors.U1 = 0: TheFaces(CurFace + 2).ReflectionVectors.V1 = 0
  TheFaces(CurFace + 2).ReflectionVectors.U2 = 1: TheFaces(CurFace + 2).ReflectionVectors.V2 = 0
  TheFaces(CurFace + 2).ReflectionVectors.U3 = 0: TheFaces(CurFace + 2).ReflectionVectors.V3 = 1
  TheFaces(CurFace + 3).ReflectionVectors.U1 = 1: TheFaces(CurFace + 3).ReflectionVectors.V1 = 0
  TheFaces(CurFace + 3).ReflectionVectors.U2 = 1: TheFaces(CurFace + 3).ReflectionVectors.V2 = 1
  TheFaces(CurFace + 3).ReflectionVectors.U3 = 0: TheFaces(CurFace + 3).ReflectionVectors.V3 = 1
  'Refraction texture-coordinates:
  TheFaces(CurFace + 2).RefractionVectors.U1 = 0: TheFaces(CurFace + 2).RefractionVectors.V1 = 0
  TheFaces(CurFace + 2).RefractionVectors.U2 = 1: TheFaces(CurFace + 2).RefractionVectors.V2 = 0
  TheFaces(CurFace + 2).RefractionVectors.U3 = 0: TheFaces(CurFace + 2).RefractionVectors.V3 = 1
  TheFaces(CurFace + 3).RefractionVectors.U1 = 1: TheFaces(CurFace + 3).RefractionVectors.V1 = 0
  TheFaces(CurFace + 3).RefractionVectors.U2 = 1: TheFaces(CurFace + 3).RefractionVectors.V2 = 1
  TheFaces(CurFace + 3).RefractionVectors.U3 = 0: TheFaces(CurFace + 3).RefractionVectors.V3 = 1
  'RefractionN texture-coordinates:
  TheFaces(CurFace + 2).RefractionNVectors.U1 = 0: TheFaces(CurFace + 2).RefractionNVectors.V1 = 0
  TheFaces(CurFace + 2).RefractionNVectors.U2 = 1: TheFaces(CurFace + 2).RefractionNVectors.V2 = 0
  TheFaces(CurFace + 2).RefractionNVectors.U3 = 0: TheFaces(CurFace + 2).RefractionNVectors.V3 = 1
  TheFaces(CurFace + 3).RefractionNVectors.U1 = 1: TheFaces(CurFace + 3).RefractionNVectors.V1 = 0
  TheFaces(CurFace + 3).RefractionNVectors.U2 = 1: TheFaces(CurFace + 3).RefractionNVectors.V2 = 1
  TheFaces(CurFace + 3).RefractionNVectors.U3 = 0: TheFaces(CurFace + 3).RefractionNVectors.V3 = 1
 End If

End Function
Function Primitive_HemiSphere(CreateAsDefault As Boolean, Radius!, Steps1&, Steps2&, Axis As Byte, Opened As Boolean, DoubleSided As Boolean) As Long

 Primitive_HemiSphere = -1

 If (CreateAsDefault = False) Then
  If ((Radius <= 0) Or (Steps1 < 2) Or (Steps2 < 2) Or (Axis < 0) Or (Axis > 2)) Then Exit Function
 ElseIf (CreateAsDefault = True) Then
  Radius = 100: Steps1 = 10: Steps2 = 10: Axis = 1: DoubleSided = False
 End If

 Dim AShapeSpline() As Vector3D, AMoveSpline() As Vector3D, CurVertex&, Tmp!

 Spline_Arc AShapeSpline(), 0.25, Radius, 0, 0, 1, Steps1
 If (Opened = False) Then ReDim Preserve AShapeSpline(UBound(AShapeSpline) + 1)

 Spline_Circle AMoveSpline(), ApproachVal, 0, 0, 2, Steps2
 ReDim Preserve AMoveSpline(UBound(AMoveSpline) + 1)
 AMoveSpline(UBound(AMoveSpline)) = AMoveSpline(1)

 Primitive_HemiSphere = Primitive_MakeFromSplines(AShapeSpline(), AMoveSpline(), 0.5, 0, 0, 1, 1, 1, 1, DoubleSided, True)
 If (Primitive_HemiSphere = -1) Then Exit Function

 TheMeshs(Primitive_HemiSphere).Label = "HemiSphere_" & CStr(TheMeshsCount)
 Select Case Axis
  Case 0:
   For CurVertex = TheMeshs(Primitive_HemiSphere).Vertices.Start To GetAddressLast(TheMeshs(Primitive_HemiSphere).Vertices)
    'Swap X to Z:
    Tmp = TheVertices(CurVertex).Position.X
    TheVertices(CurVertex).Position.X = TheVertices(CurVertex).Position.Z
    TheVertices(CurVertex).Position.Z = Tmp
    'Swap Z to Y:
    Tmp = TheVertices(CurVertex).Position.Z
    TheVertices(CurVertex).Position.Z = TheVertices(CurVertex).Position.Y
    TheVertices(CurVertex).Position.Y = Tmp
   Next CurVertex
  Case 1:
   For CurVertex = TheMeshs(Primitive_HemiSphere).Vertices.Start To GetAddressLast(TheMeshs(Primitive_HemiSphere).Vertices)
    'Swap Z to Y:
    Tmp = TheVertices(CurVertex).Position.Z
    TheVertices(CurVertex).Position.Z = TheVertices(CurVertex).Position.Y
    TheVertices(CurVertex).Position.Y = Tmp
   Next CurVertex
 End Select

End Function
Function Primitive_Landscape(TheBitmap As BitMap2D, ATexelsFilter As K3DE_TEXELS_FILTER_MODES, Width!, Height!, Depth!, Steps&, Axis As Byte, DoubleSided As Boolean) As Long

 'Apply a depth variation effect (displacement mapping)
 'on a grid using an 8-bits channel bitmap.

 Primitive_Landscape = -1

 If (BitMap2D_IsValid(TheBitmap) = False) Then Exit Function
 If (TheBitmap.BitsDepth = 24) Then Exit Function
 If ((Width <= 0) Or (Height <= 0) Or (Depth < 0)) Then Exit Function
 If ((Steps < 2) Or (Axis < 0) Or (Axis > 2)) Then Exit Function

 Dim CurX&, CurY&, CurU!, CurV!, CurVertex&, TheDepth!, TheColor As Byte

 Primitive_Landscape = Primitive_Grid(False, Width, Height, Steps, Axis, DoubleSided)

 If (Primitive_Landscape = -1) Then Exit Function
 TheMeshs(Primitive_Landscape).Label = "Landscape_" & CStr(TheMeshsCount)
 If (ATexelsFilter = K3DE_XFM_NOFILTER) Then ATexelsFilter = K3DE_XFM_BILINEAR 'Set default bilinear

 For CurY = 0 To Steps: CurV = (CurY / (Steps + 1))
  For CurX = 0 To Steps: CurU = (CurX / (Steps + 1))
   TheColor = DoTexelFiltering8(ATexelsFilter, TheBitmap, CurU, CurV, True)
   TheDepth = ((TheColor * AlphaFactor) * Depth)
   CurVertex = ((TheMeshs(Primitive_Landscape).Vertices.Start - 1) + ((CurY * (Steps + 1)) + (CurX + 1)))
   Select Case Axis
    Case 0: TheVertices(CurVertex).Position.X = ((Depth * 0.5) - TheDepth)
    Case 1: TheVertices(CurVertex).Position.Y = ((Depth * 0.5) - TheDepth)
    Case 2: TheVertices(CurVertex).Position.Z = ((Depth * 0.5) - TheDepth)
   End Select
  Next CurX
 Next CurY

End Function
Function Primitive_MakeFromSplines(ShapeSpline() As Vector3D, MoveSpline() As Vector3D, ThePosOnRay!, StartAngle!, EndAngle!, StartScale1!, EndScale1!, StartScale2!, EndScale2!, DoubleSided As Boolean, TrimIt As Boolean) As Long

 Primitive_MakeFromSplines = -1

 If ((UBound(ShapeSpline) < 1) Or (UBound(MoveSpline) < 2)) Then Exit Function
 If ((ThePosOnRay < 0) Or (ThePosOnRay > 1)) Then Exit Function
 If ((StartAngle < 0)) Then Exit Function
 If ((EndAngle < StartAngle)) Then Exit Function
 If ((StartScale1 <= 0) Or (EndScale1 <= 0)) Then Exit Function
 If ((StartScale2 <= 0) Or (EndScale2 <= 0)) Then Exit Function

 Dim MovesCount&, SamplesCount&, CurVertex&, CurFace&
 Dim CurMove&, CurSample&, VertexPos&, FacePos&
 Dim CurAngle!, CurScale1!, CurScale2!, CurU1!, CurV1!, CurU2!, CurV2!
 Dim Orientation As Vector3D, NotScale As Boolean

 SamplesCount = UBound(ShapeSpline): MovesCount = UBound(MoveSpline)

 If (DoubleSided = False) Then
  Primitive_MakeFromSplines = Mesh3D_Add(((SamplesCount + 1) * MovesCount), ((SamplesCount * (MovesCount - 1)) * 2))
 Else
  Primitive_MakeFromSplines = Mesh3D_Add((((SamplesCount + 1) * MovesCount) * 2), ((SamplesCount * (MovesCount - 1)) * 4))
 End If

 If (Primitive_MakeFromSplines = -1) Then Exit Function

 CurVertex = TheMeshs(Primitive_MakeFromSplines).Vertices.Start
 CurFace = TheMeshs(Primitive_MakeFromSplines).Faces.Start
 Orientation.Z = 1
 If ((StartScale1 = 1) And (EndScale1 = 1)) Then
  If ((StartScale2 = 1) And (EndScale2 = 1)) Then
   NotScale = True
  End If
 End If

 'Add vertices
 '============

 For CurMove = 1 To MovesCount
  Orientation = VectorSubtract(MoveSpline(CurMove), MoveSpline(CurMove - 1))
  For CurSample = 0 To SamplesCount
   CurAngle = (StartAngle + ((EndAngle - StartAngle) * (CurMove / MovesCount)))
   TheVertices(CurVertex + VertexPos).Position = MatrixMultiplyVector(ShapeSpline(CurSample), MatrixRotation(2, CurAngle))
   If (NotScale = False) Then
    CurScale1 = (StartScale1 + ((EndScale1 - StartScale1) * (CurMove / MovesCount)))
    CurScale2 = (StartScale2 + ((EndScale2 - StartScale2) * (CurMove / MovesCount)))
    TheVertices(CurVertex + VertexPos).Position = MatrixMultiplyVector(TheVertices(CurVertex + VertexPos).Position, MatrixScaling(VectorInput(CurScale1, CurScale2, 1)))
   End If
   TheVertices(CurVertex + VertexPos).Position = MatrixMultiplyVector(TheVertices(CurVertex + VertexPos).Position, MatrixRotationByVectors(VectorInput(0, 0, 1), Orientation))
   TheVertices(CurVertex + VertexPos).Position = VectorAdd(TheVertices(CurVertex + VertexPos).Position, VectorInterpolate(MoveSpline(CurMove), MoveSpline(CurMove - 1), ThePosOnRay))
   VertexPos = (VertexPos + 1)
  Next CurSample
 Next CurMove

 'Add faces
 '=========

 For CurMove = 0 To (MovesCount - 2)
  CurV1 = (CurMove / MovesCount): CurV2 = ((CurMove + 1) / MovesCount)
  For CurSample = 0 To (SamplesCount - 1)
   CurU1 = (CurSample / SamplesCount): CurU2 = ((CurSample + 1) / SamplesCount)
   TheFaces(CurFace + FacePos).A = (CurVertex + (((SamplesCount + 1) * (CurMove + 1)) + CurSample))
   TheFaces(CurFace + FacePos).B = (CurVertex + (((SamplesCount + 1) * CurMove) + (CurSample + 1)))
   TheFaces(CurFace + FacePos).C = (CurVertex + (((SamplesCount + 1) * CurMove) + CurSample))
   'Setup texture coordinates
   '=========================
   TheFaces(CurFace + FacePos).AlphaVectors.U1 = CurU1        'Alpha texture-coordinates
   TheFaces(CurFace + FacePos).AlphaVectors.V1 = CurV2
   TheFaces(CurFace + FacePos).AlphaVectors.U2 = CurU2
   TheFaces(CurFace + FacePos).AlphaVectors.V2 = CurV1
   TheFaces(CurFace + FacePos).AlphaVectors.U3 = CurU1
   TheFaces(CurFace + FacePos).AlphaVectors.V3 = CurV1
   TheFaces(CurFace + FacePos).ColorVectors.U1 = CurU1        'Color texture-coordinates
   TheFaces(CurFace + FacePos).ColorVectors.V1 = CurV2
   TheFaces(CurFace + FacePos).ColorVectors.U2 = CurU2
   TheFaces(CurFace + FacePos).ColorVectors.V2 = CurV1
   TheFaces(CurFace + FacePos).ColorVectors.U3 = CurU1
   TheFaces(CurFace + FacePos).ColorVectors.V3 = CurV1
   TheFaces(CurFace + FacePos).ReflectionVectors.U1 = CurU1      'Reflection texture-coordinates
   TheFaces(CurFace + FacePos).ReflectionVectors.V1 = CurV2
   TheFaces(CurFace + FacePos).ReflectionVectors.U2 = CurU2
   TheFaces(CurFace + FacePos).ReflectionVectors.V2 = CurV1
   TheFaces(CurFace + FacePos).ReflectionVectors.U3 = CurU1
   TheFaces(CurFace + FacePos).ReflectionVectors.V3 = CurV1
   TheFaces(CurFace + FacePos).RefractionVectors.U1 = CurU1    'Refraction texture-coordinates
   TheFaces(CurFace + FacePos).RefractionVectors.V1 = CurV2
   TheFaces(CurFace + FacePos).RefractionVectors.U2 = CurU2
   TheFaces(CurFace + FacePos).RefractionVectors.V2 = CurV1
   TheFaces(CurFace + FacePos).RefractionVectors.U3 = CurU1
   TheFaces(CurFace + FacePos).RefractionVectors.V3 = CurV1
   TheFaces(CurFace + FacePos).RefractionNVectors.U1 = CurU1    'RefractionN texture-coordinates
   TheFaces(CurFace + FacePos).RefractionNVectors.V1 = CurV2
   TheFaces(CurFace + FacePos).RefractionNVectors.U2 = CurU2
   TheFaces(CurFace + FacePos).RefractionNVectors.V2 = CurV1
   TheFaces(CurFace + FacePos).RefractionNVectors.U3 = CurU1
   TheFaces(CurFace + FacePos).RefractionNVectors.V3 = CurV1
   FacePos = (FacePos + 1)

   TheFaces(CurFace + FacePos).A = (CurVertex + (((SamplesCount + 1) * (CurMove + 1)) + CurSample))
   TheFaces(CurFace + FacePos).B = (CurVertex + (((SamplesCount + 1) * (CurMove + 1)) + (CurSample + 1)))
   TheFaces(CurFace + FacePos).C = (CurVertex + (((SamplesCount + 1) * CurMove) + (CurSample + 1)))
   'Setup texture coordinates
   '=========================
   TheFaces(CurFace + FacePos).AlphaVectors.U1 = CurU1        'Alpha texture-coordinates
   TheFaces(CurFace + FacePos).AlphaVectors.V1 = CurV2
   TheFaces(CurFace + FacePos).AlphaVectors.U2 = CurU2
   TheFaces(CurFace + FacePos).AlphaVectors.V2 = CurV2
   TheFaces(CurFace + FacePos).AlphaVectors.U3 = CurU2
   TheFaces(CurFace + FacePos).AlphaVectors.V3 = CurV1
   TheFaces(CurFace + FacePos).ColorVectors.U1 = CurU1        'Color texture-coordinates
   TheFaces(CurFace + FacePos).ColorVectors.V1 = CurV2
   TheFaces(CurFace + FacePos).ColorVectors.U2 = CurU2
   TheFaces(CurFace + FacePos).ColorVectors.V2 = CurV2
   TheFaces(CurFace + FacePos).ColorVectors.U3 = CurU2
   TheFaces(CurFace + FacePos).ColorVectors.V3 = CurV1
   TheFaces(CurFace + FacePos).ReflectionVectors.U1 = CurU1      'Reflection texture-coordinates
   TheFaces(CurFace + FacePos).ReflectionVectors.V1 = CurV2
   TheFaces(CurFace + FacePos).ReflectionVectors.U2 = CurU2
   TheFaces(CurFace + FacePos).ReflectionVectors.V2 = CurV2
   TheFaces(CurFace + FacePos).ReflectionVectors.U3 = CurU2
   TheFaces(CurFace + FacePos).ReflectionVectors.V3 = CurV1
   TheFaces(CurFace + FacePos).RefractionVectors.U1 = CurU1    'Refraction texture-coordinates
   TheFaces(CurFace + FacePos).RefractionVectors.V1 = CurV2
   TheFaces(CurFace + FacePos).RefractionVectors.U2 = CurU2
   TheFaces(CurFace + FacePos).RefractionVectors.V2 = CurV2
   TheFaces(CurFace + FacePos).RefractionVectors.U3 = CurU2
   TheFaces(CurFace + FacePos).RefractionVectors.V3 = CurV1
   TheFaces(CurFace + FacePos).RefractionNVectors.U1 = CurU1    'RefractionN texture-coordinates
   TheFaces(CurFace + FacePos).RefractionNVectors.V1 = CurV2
   TheFaces(CurFace + FacePos).RefractionNVectors.U2 = CurU2
   TheFaces(CurFace + FacePos).RefractionNVectors.V2 = CurV2
   TheFaces(CurFace + FacePos).RefractionNVectors.U3 = CurU2
   TheFaces(CurFace + FacePos).RefractionNVectors.V3 = CurV1
   FacePos = (FacePos + 1)

   If (DoubleSided = True) Then
    TheFaces(CurFace + FacePos).C = (CurVertex + (((SamplesCount + 1) * (CurMove + 1)) + CurSample))
    TheFaces(CurFace + FacePos).B = (CurVertex + (((SamplesCount + 1) * CurMove) + (CurSample + 1)))
    TheFaces(CurFace + FacePos).A = (CurVertex + (((SamplesCount + 1) * CurMove) + CurSample))
    'Setup texture coordinates
    '=========================
    TheFaces(CurFace + FacePos).AlphaVectors.U3 = CurU1        'Alpha texture-coordinates
    TheFaces(CurFace + FacePos).AlphaVectors.V3 = CurV2
    TheFaces(CurFace + FacePos).AlphaVectors.U2 = CurU2
    TheFaces(CurFace + FacePos).AlphaVectors.V2 = CurV1
    TheFaces(CurFace + FacePos).AlphaVectors.U1 = CurU1
    TheFaces(CurFace + FacePos).AlphaVectors.V1 = CurV1
    TheFaces(CurFace + FacePos).ColorVectors.U3 = CurU1        'Color texture-coordinates
    TheFaces(CurFace + FacePos).ColorVectors.V3 = CurV2
    TheFaces(CurFace + FacePos).ColorVectors.U2 = CurU2
    TheFaces(CurFace + FacePos).ColorVectors.V2 = CurV1
    TheFaces(CurFace + FacePos).ColorVectors.U1 = CurU1
    TheFaces(CurFace + FacePos).ColorVectors.V1 = CurV1
    TheFaces(CurFace + FacePos).ReflectionVectors.U3 = CurU1      'Reflection texture-coordinates
    TheFaces(CurFace + FacePos).ReflectionVectors.V3 = CurV2
    TheFaces(CurFace + FacePos).ReflectionVectors.U2 = CurU2
    TheFaces(CurFace + FacePos).ReflectionVectors.V2 = CurV1
    TheFaces(CurFace + FacePos).ReflectionVectors.U1 = CurU1
    TheFaces(CurFace + FacePos).ReflectionVectors.V1 = CurV1
    TheFaces(CurFace + FacePos).RefractionVectors.U3 = CurU1    'Refraction texture-coordinates
    TheFaces(CurFace + FacePos).RefractionVectors.V3 = CurV2
    TheFaces(CurFace + FacePos).RefractionVectors.U2 = CurU2
    TheFaces(CurFace + FacePos).RefractionVectors.V2 = CurV1
    TheFaces(CurFace + FacePos).RefractionVectors.U1 = CurU1
    TheFaces(CurFace + FacePos).RefractionVectors.V1 = CurV1
    TheFaces(CurFace + FacePos).RefractionNVectors.U3 = CurU1    'RefractionN texture-coordinates
    TheFaces(CurFace + FacePos).RefractionNVectors.V3 = CurV2
    TheFaces(CurFace + FacePos).RefractionNVectors.U2 = CurU2
    TheFaces(CurFace + FacePos).RefractionNVectors.V2 = CurV1
    TheFaces(CurFace + FacePos).RefractionNVectors.U1 = CurU1
    TheFaces(CurFace + FacePos).RefractionNVectors.V1 = CurV1
    FacePos = (FacePos + 1)

    TheFaces(CurFace + FacePos).C = (CurVertex + (((SamplesCount + 1) * (CurMove + 1)) + CurSample))
    TheFaces(CurFace + FacePos).B = (CurVertex + (((SamplesCount + 1) * (CurMove + 1)) + (CurSample + 1)))
    TheFaces(CurFace + FacePos).A = (CurVertex + (((SamplesCount + 1) * CurMove) + (CurSample + 1)))
    'Setup texture coordinates
    '=========================
    TheFaces(CurFace + FacePos).AlphaVectors.U3 = CurU1        'Alpha texture-coordinates
    TheFaces(CurFace + FacePos).AlphaVectors.V3 = CurV2
    TheFaces(CurFace + FacePos).AlphaVectors.U2 = CurU2
    TheFaces(CurFace + FacePos).AlphaVectors.V2 = CurV2
    TheFaces(CurFace + FacePos).AlphaVectors.U1 = CurU2
    TheFaces(CurFace + FacePos).AlphaVectors.V1 = CurV1
    TheFaces(CurFace + FacePos).ColorVectors.U3 = CurU1        'Color texture-coordinates
    TheFaces(CurFace + FacePos).ColorVectors.V3 = CurV2
    TheFaces(CurFace + FacePos).ColorVectors.U2 = CurU2
    TheFaces(CurFace + FacePos).ColorVectors.V2 = CurV2
    TheFaces(CurFace + FacePos).ColorVectors.U1 = CurU2
    TheFaces(CurFace + FacePos).ColorVectors.V1 = CurV1
    TheFaces(CurFace + FacePos).ReflectionVectors.U3 = CurU1      'Reflection texture-coordinates
    TheFaces(CurFace + FacePos).ReflectionVectors.V3 = CurV2
    TheFaces(CurFace + FacePos).ReflectionVectors.U2 = CurU2
    TheFaces(CurFace + FacePos).ReflectionVectors.V2 = CurV2
    TheFaces(CurFace + FacePos).ReflectionVectors.U1 = CurU2
    TheFaces(CurFace + FacePos).ReflectionVectors.V1 = CurV1
    TheFaces(CurFace + FacePos).RefractionVectors.U3 = CurU1    'Refraction texture-coordinates
    TheFaces(CurFace + FacePos).RefractionVectors.V3 = CurV2
    TheFaces(CurFace + FacePos).RefractionVectors.U2 = CurU2
    TheFaces(CurFace + FacePos).RefractionVectors.V2 = CurV2
    TheFaces(CurFace + FacePos).RefractionVectors.U1 = CurU2
    TheFaces(CurFace + FacePos).RefractionVectors.V1 = CurV1
    TheFaces(CurFace + FacePos).RefractionNVectors.U3 = CurU1    'RefractionN texture-coordinates
    TheFaces(CurFace + FacePos).RefractionNVectors.V3 = CurV2
    TheFaces(CurFace + FacePos).RefractionNVectors.U2 = CurU2
    TheFaces(CurFace + FacePos).RefractionNVectors.V2 = CurV2
    TheFaces(CurFace + FacePos).RefractionNVectors.U1 = CurU2
    TheFaces(CurFace + FacePos).RefractionNVectors.V1 = CurV1
    FacePos = (FacePos + 1)
   End If
  Next CurSample
 Next CurMove

 If (TrimIt = True) Then Mesh3D_Trim Primitive_MakeFromSplines

End Function
Function Primitive_Octahedron(CreateAsDefault As Boolean, Base!, Radius!, DoubleSided As Boolean) As Long

 Primitive_Octahedron = -1

 If (CreateAsDefault = False) Then
  If ((Base <= 0) Or (Radius <= 0)) Then Exit Function
 ElseIf (CreateAsDefault = True) Then
  Base = 100: Radius = 100: DoubleSided = False
 End If

 Dim CurVertex&, CurFace&

 If (DoubleSided = False) Then
  Primitive_Octahedron = Mesh3D_Add(6, 8)
 Else
  Primitive_Octahedron = Mesh3D_Add(6, 16)
 End If

 If (Primitive_Octahedron = -1) Then Exit Function

 TheMeshs(Primitive_Octahedron).Label = "Octahedron_" & CStr(TheMeshsCount)
 CurVertex = TheMeshs(Primitive_Octahedron).Vertices.Start
 CurFace = TheMeshs(Primitive_Octahedron).Faces.Start

 'Add vertices
 '============
 TheVertices(CurVertex).Position = VectorInput(0, -Base, 0)
 TheVertices(CurVertex + 1).Position = VectorInput(Radius, 0, 0)
 TheVertices(CurVertex + 2).Position = VectorInput(0, 0, Radius)
 TheVertices(CurVertex + 3).Position = VectorInput(-Radius, 0, 0)
 TheVertices(CurVertex + 4).Position = VectorInput(0, 0, -Radius)
 TheVertices(CurVertex + 5).Position = VectorInput(0, Base, 0)

 'Add faces
 '=========
 TheFaces(CurFace).A = (CurVertex + 0): TheFaces(CurFace).B = (CurVertex + 1): TheFaces(CurFace).C = (CurVertex + 2)
 TheFaces(CurFace + 1).A = (CurVertex + 0): TheFaces(CurFace + 1).B = (CurVertex + 2): TheFaces(CurFace + 1).C = (CurVertex + 3)
 TheFaces(CurFace + 2).A = (CurVertex + 0): TheFaces(CurFace + 2).B = (CurVertex + 3): TheFaces(CurFace + 2).C = (CurVertex + 4)
 TheFaces(CurFace + 3).A = (CurVertex + 0): TheFaces(CurFace + 3).B = (CurVertex + 4): TheFaces(CurFace + 3).C = (CurVertex + 1)
 TheFaces(CurFace + 4).A = (CurVertex + 1): TheFaces(CurFace + 4).B = (CurVertex + 5): TheFaces(CurFace + 4).C = (CurVertex + 2)
 TheFaces(CurFace + 5).A = (CurVertex + 2): TheFaces(CurFace + 5).B = (CurVertex + 5): TheFaces(CurFace + 5).C = (CurVertex + 3)
 TheFaces(CurFace + 6).A = (CurVertex + 3): TheFaces(CurFace + 6).B = (CurVertex + 5): TheFaces(CurFace + 6).C = (CurVertex + 4)
 TheFaces(CurFace + 7).A = (CurVertex + 4): TheFaces(CurFace + 7).B = (CurVertex + 5): TheFaces(CurFace + 7).C = (CurVertex + 1)

 'Setup texture coordinates
 '=========================
 'Alpha texture-coordinates:
 TheFaces(CurFace).AlphaVectors.U1 = 0.5: TheFaces(CurFace).AlphaVectors.V1 = 0
 TheFaces(CurFace).AlphaVectors.U2 = 0.25: TheFaces(CurFace).AlphaVectors.V2 = 0.5
 TheFaces(CurFace).AlphaVectors.U3 = 0: TheFaces(CurFace).AlphaVectors.V3 = 0.5
 TheFaces(CurFace + 1).AlphaVectors.U1 = 0.5: TheFaces(CurFace + 1).AlphaVectors.V1 = 0
 TheFaces(CurFace + 1).AlphaVectors.U2 = 0.5: TheFaces(CurFace + 1).AlphaVectors.V2 = 0.5
 TheFaces(CurFace + 1).AlphaVectors.U3 = 0.25: TheFaces(CurFace + 1).AlphaVectors.V3 = 0.5
 TheFaces(CurFace + 2).AlphaVectors.U1 = 0.5: TheFaces(CurFace + 2).AlphaVectors.V1 = 0
 TheFaces(CurFace + 2).AlphaVectors.U2 = 0.75: TheFaces(CurFace + 2).AlphaVectors.V2 = 0.5
 TheFaces(CurFace + 2).AlphaVectors.U3 = 0.5: TheFaces(CurFace + 2).AlphaVectors.V3 = 0.5
 TheFaces(CurFace + 3).AlphaVectors.U1 = 0.5: TheFaces(CurFace + 3).AlphaVectors.V1 = 0
 TheFaces(CurFace + 3).AlphaVectors.U2 = 1: TheFaces(CurFace + 3).AlphaVectors.V2 = 0.5
 TheFaces(CurFace + 3).AlphaVectors.U3 = 0.75: TheFaces(CurFace + 3).AlphaVectors.V3 = 0.5
 TheFaces(CurFace + 4).AlphaVectors.U1 = 0.25: TheFaces(CurFace + 4).AlphaVectors.V1 = 0.5
 TheFaces(CurFace + 4).AlphaVectors.U2 = 0.5: TheFaces(CurFace + 4).AlphaVectors.V2 = 1
 TheFaces(CurFace + 4).AlphaVectors.U3 = 0: TheFaces(CurFace + 4).AlphaVectors.V3 = 0.5
 TheFaces(CurFace + 5).AlphaVectors.U1 = 0.5: TheFaces(CurFace + 5).AlphaVectors.V1 = 0.5
 TheFaces(CurFace + 5).AlphaVectors.U2 = 0.5: TheFaces(CurFace + 5).AlphaVectors.V2 = 1
 TheFaces(CurFace + 5).AlphaVectors.U3 = 0.25: TheFaces(CurFace + 5).AlphaVectors.V3 = 0.5
 TheFaces(CurFace + 6).AlphaVectors.U1 = 0.75: TheFaces(CurFace + 6).AlphaVectors.V1 = 0.5
 TheFaces(CurFace + 6).AlphaVectors.U2 = 0.5: TheFaces(CurFace + 6).AlphaVectors.V2 = 1
 TheFaces(CurFace + 6).AlphaVectors.U3 = 0.5: TheFaces(CurFace + 6).AlphaVectors.V3 = 0.5
 TheFaces(CurFace + 7).AlphaVectors.U1 = 1: TheFaces(CurFace + 7).AlphaVectors.V1 = 0.5
 TheFaces(CurFace + 7).AlphaVectors.U2 = 0.5: TheFaces(CurFace + 7).AlphaVectors.V2 = 1
 TheFaces(CurFace + 7).AlphaVectors.U3 = 0.75: TheFaces(CurFace + 7).AlphaVectors.V3 = 0.5
 'Color texture-coordinates:
 TheFaces(CurFace).ColorVectors.U1 = 0.5: TheFaces(CurFace).ColorVectors.V1 = 0
 TheFaces(CurFace).ColorVectors.U2 = 0.25: TheFaces(CurFace).ColorVectors.V2 = 0.5
 TheFaces(CurFace).ColorVectors.U3 = 0: TheFaces(CurFace).ColorVectors.V3 = 0.5
 TheFaces(CurFace + 1).ColorVectors.U1 = 0.5: TheFaces(CurFace + 1).ColorVectors.V1 = 0
 TheFaces(CurFace + 1).ColorVectors.U2 = 0.5: TheFaces(CurFace + 1).ColorVectors.V2 = 0.5
 TheFaces(CurFace + 1).ColorVectors.U3 = 0.25: TheFaces(CurFace + 1).ColorVectors.V3 = 0.5
 TheFaces(CurFace + 2).ColorVectors.U1 = 0.5: TheFaces(CurFace + 2).ColorVectors.V1 = 0
 TheFaces(CurFace + 2).ColorVectors.U2 = 0.75: TheFaces(CurFace + 2).ColorVectors.V2 = 0.5
 TheFaces(CurFace + 2).ColorVectors.U3 = 0.5: TheFaces(CurFace + 2).ColorVectors.V3 = 0.5
 TheFaces(CurFace + 3).ColorVectors.U1 = 0.5: TheFaces(CurFace + 3).ColorVectors.V1 = 0
 TheFaces(CurFace + 3).ColorVectors.U2 = 1: TheFaces(CurFace + 3).ColorVectors.V2 = 0.5
 TheFaces(CurFace + 3).ColorVectors.U3 = 0.75: TheFaces(CurFace + 3).ColorVectors.V3 = 0.5
 TheFaces(CurFace + 4).ColorVectors.U1 = 0.25: TheFaces(CurFace + 4).ColorVectors.V1 = 0.5
 TheFaces(CurFace + 4).ColorVectors.U2 = 0.5: TheFaces(CurFace + 4).ColorVectors.V2 = 1
 TheFaces(CurFace + 4).ColorVectors.U3 = 0: TheFaces(CurFace + 4).ColorVectors.V3 = 0.5
 TheFaces(CurFace + 5).ColorVectors.U1 = 0.5: TheFaces(CurFace + 5).ColorVectors.V1 = 0.5
 TheFaces(CurFace + 5).ColorVectors.U2 = 0.5: TheFaces(CurFace + 5).ColorVectors.V2 = 1
 TheFaces(CurFace + 5).ColorVectors.U3 = 0.25: TheFaces(CurFace + 5).ColorVectors.V3 = 0.5
 TheFaces(CurFace + 6).ColorVectors.U1 = 0.75: TheFaces(CurFace + 6).ColorVectors.V1 = 0.5
 TheFaces(CurFace + 6).ColorVectors.U2 = 0.5: TheFaces(CurFace + 6).ColorVectors.V2 = 1
 TheFaces(CurFace + 6).ColorVectors.U3 = 0.5: TheFaces(CurFace + 6).ColorVectors.V3 = 0.5
 TheFaces(CurFace + 7).ColorVectors.U1 = 1: TheFaces(CurFace + 7).ColorVectors.V1 = 0.5
 TheFaces(CurFace + 7).ColorVectors.U2 = 0.5: TheFaces(CurFace + 7).ColorVectors.V2 = 1
 TheFaces(CurFace + 7).ColorVectors.U3 = 0.75: TheFaces(CurFace + 7).ColorVectors.V3 = 0.5
 'Reflection texture-coordinates:
 TheFaces(CurFace).ReflectionVectors.U1 = 0.5: TheFaces(CurFace).ReflectionVectors.V1 = 0
 TheFaces(CurFace).ReflectionVectors.U2 = 0.25: TheFaces(CurFace).ReflectionVectors.V2 = 0.5
 TheFaces(CurFace).ReflectionVectors.U3 = 0: TheFaces(CurFace).ReflectionVectors.V3 = 0.5
 TheFaces(CurFace + 1).ReflectionVectors.U1 = 0.5: TheFaces(CurFace + 1).ReflectionVectors.V1 = 0
 TheFaces(CurFace + 1).ReflectionVectors.U2 = 0.5: TheFaces(CurFace + 1).ReflectionVectors.V2 = 0.5
 TheFaces(CurFace + 1).ReflectionVectors.U3 = 0.25: TheFaces(CurFace + 1).ReflectionVectors.V3 = 0.5
 TheFaces(CurFace + 2).ReflectionVectors.U1 = 0.5: TheFaces(CurFace + 2).ReflectionVectors.V1 = 0
 TheFaces(CurFace + 2).ReflectionVectors.U2 = 0.75: TheFaces(CurFace + 2).ReflectionVectors.V2 = 0.5
 TheFaces(CurFace + 2).ReflectionVectors.U3 = 0.5: TheFaces(CurFace + 2).ReflectionVectors.V3 = 0.5
 TheFaces(CurFace + 3).ReflectionVectors.U1 = 0.5: TheFaces(CurFace + 3).ReflectionVectors.V1 = 0
 TheFaces(CurFace + 3).ReflectionVectors.U2 = 1: TheFaces(CurFace + 3).ReflectionVectors.V2 = 0.5
 TheFaces(CurFace + 3).ReflectionVectors.U3 = 0.75: TheFaces(CurFace + 3).ReflectionVectors.V3 = 0.5
 TheFaces(CurFace + 4).ReflectionVectors.U1 = 0.25: TheFaces(CurFace + 4).ReflectionVectors.V1 = 0.5
 TheFaces(CurFace + 4).ReflectionVectors.U2 = 0.5: TheFaces(CurFace + 4).ReflectionVectors.V2 = 1
 TheFaces(CurFace + 4).ReflectionVectors.U3 = 0: TheFaces(CurFace + 4).ReflectionVectors.V3 = 0.5
 TheFaces(CurFace + 5).ReflectionVectors.U1 = 0.5: TheFaces(CurFace + 5).ReflectionVectors.V1 = 0.5
 TheFaces(CurFace + 5).ReflectionVectors.U2 = 0.5: TheFaces(CurFace + 5).ReflectionVectors.V2 = 1
 TheFaces(CurFace + 5).ReflectionVectors.U3 = 0.25: TheFaces(CurFace + 5).ReflectionVectors.V3 = 0.5
 TheFaces(CurFace + 6).ReflectionVectors.U1 = 0.75: TheFaces(CurFace + 6).ReflectionVectors.V1 = 0.5
 TheFaces(CurFace + 6).ReflectionVectors.U2 = 0.5: TheFaces(CurFace + 6).ReflectionVectors.V2 = 1
 TheFaces(CurFace + 6).ReflectionVectors.U3 = 0.5: TheFaces(CurFace + 6).ReflectionVectors.V3 = 0.5
 TheFaces(CurFace + 7).ReflectionVectors.U1 = 1: TheFaces(CurFace + 7).ReflectionVectors.V1 = 0.5
 TheFaces(CurFace + 7).ReflectionVectors.U2 = 0.5: TheFaces(CurFace + 7).ReflectionVectors.V2 = 1
 TheFaces(CurFace + 7).ReflectionVectors.U3 = 0.75: TheFaces(CurFace + 7).ReflectionVectors.V3 = 0.5
 'Refraction texture-coordinates:
 TheFaces(CurFace).RefractionVectors.U1 = 0.5: TheFaces(CurFace).RefractionVectors.V1 = 0
 TheFaces(CurFace).RefractionVectors.U2 = 0.25: TheFaces(CurFace).RefractionVectors.V2 = 0.5
 TheFaces(CurFace).RefractionVectors.U3 = 0: TheFaces(CurFace).RefractionVectors.V3 = 0.5
 TheFaces(CurFace + 1).RefractionVectors.U1 = 0.5: TheFaces(CurFace + 1).RefractionVectors.V1 = 0
 TheFaces(CurFace + 1).RefractionVectors.U2 = 0.5: TheFaces(CurFace + 1).RefractionVectors.V2 = 0.5
 TheFaces(CurFace + 1).RefractionVectors.U3 = 0.25: TheFaces(CurFace + 1).RefractionVectors.V3 = 0.5
 TheFaces(CurFace + 2).RefractionVectors.U1 = 0.5: TheFaces(CurFace + 2).RefractionVectors.V1 = 0
 TheFaces(CurFace + 2).RefractionVectors.U2 = 0.75: TheFaces(CurFace + 2).RefractionVectors.V2 = 0.5
 TheFaces(CurFace + 2).RefractionVectors.U3 = 0.5: TheFaces(CurFace + 2).RefractionVectors.V3 = 0.5
 TheFaces(CurFace + 3).RefractionVectors.U1 = 0.5: TheFaces(CurFace + 3).RefractionVectors.V1 = 0
 TheFaces(CurFace + 3).RefractionVectors.U2 = 1: TheFaces(CurFace + 3).RefractionVectors.V2 = 0.5
 TheFaces(CurFace + 3).RefractionVectors.U3 = 0.75: TheFaces(CurFace + 3).RefractionVectors.V3 = 0.5
 TheFaces(CurFace + 4).RefractionVectors.U1 = 0.25: TheFaces(CurFace + 4).RefractionVectors.V1 = 0.5
 TheFaces(CurFace + 4).RefractionVectors.U2 = 0.5: TheFaces(CurFace + 4).RefractionVectors.V2 = 1
 TheFaces(CurFace + 4).RefractionVectors.U3 = 0: TheFaces(CurFace + 4).RefractionVectors.V3 = 0.5
 TheFaces(CurFace + 5).RefractionVectors.U1 = 0.5: TheFaces(CurFace + 5).RefractionVectors.V1 = 0.5
 TheFaces(CurFace + 5).RefractionVectors.U2 = 0.5: TheFaces(CurFace + 5).RefractionVectors.V2 = 1
 TheFaces(CurFace + 5).RefractionVectors.U3 = 0.25: TheFaces(CurFace + 5).RefractionVectors.V3 = 0.5
 TheFaces(CurFace + 6).RefractionVectors.U1 = 0.75: TheFaces(CurFace + 6).RefractionVectors.V1 = 0.5
 TheFaces(CurFace + 6).RefractionVectors.U2 = 0.5: TheFaces(CurFace + 6).RefractionVectors.V2 = 1
 TheFaces(CurFace + 6).RefractionVectors.U3 = 0.5: TheFaces(CurFace + 6).RefractionVectors.V3 = 0.5
 TheFaces(CurFace + 7).RefractionVectors.U1 = 1: TheFaces(CurFace + 7).RefractionVectors.V1 = 0.5
 TheFaces(CurFace + 7).RefractionVectors.U2 = 0.5: TheFaces(CurFace + 7).RefractionVectors.V2 = 1
 TheFaces(CurFace + 7).RefractionVectors.U3 = 0.75: TheFaces(CurFace + 7).RefractionVectors.V3 = 0.5
 'RefractionN texture-coordinates:
 TheFaces(CurFace).RefractionNVectors.U1 = 0.5: TheFaces(CurFace).RefractionNVectors.V1 = 0
 TheFaces(CurFace).RefractionNVectors.U2 = 0.25: TheFaces(CurFace).RefractionNVectors.V2 = 0.5
 TheFaces(CurFace).RefractionNVectors.U3 = 0: TheFaces(CurFace).RefractionNVectors.V3 = 0.5
 TheFaces(CurFace + 1).RefractionNVectors.U1 = 0.5: TheFaces(CurFace + 1).RefractionNVectors.V1 = 0
 TheFaces(CurFace + 1).RefractionNVectors.U2 = 0.5: TheFaces(CurFace + 1).RefractionNVectors.V2 = 0.5
 TheFaces(CurFace + 1).RefractionNVectors.U3 = 0.25: TheFaces(CurFace + 1).RefractionNVectors.V3 = 0.5
 TheFaces(CurFace + 2).RefractionNVectors.U1 = 0.5: TheFaces(CurFace + 2).RefractionNVectors.V1 = 0
 TheFaces(CurFace + 2).RefractionNVectors.U2 = 0.75: TheFaces(CurFace + 2).RefractionNVectors.V2 = 0.5
 TheFaces(CurFace + 2).RefractionNVectors.U3 = 0.5: TheFaces(CurFace + 2).RefractionNVectors.V3 = 0.5
 TheFaces(CurFace + 3).RefractionNVectors.U1 = 0.5: TheFaces(CurFace + 3).RefractionNVectors.V1 = 0
 TheFaces(CurFace + 3).RefractionNVectors.U2 = 1: TheFaces(CurFace + 3).RefractionNVectors.V2 = 0.5
 TheFaces(CurFace + 3).RefractionNVectors.U3 = 0.75: TheFaces(CurFace + 3).RefractionNVectors.V3 = 0.5
 TheFaces(CurFace + 4).RefractionNVectors.U1 = 0.25: TheFaces(CurFace + 4).RefractionNVectors.V1 = 0.5
 TheFaces(CurFace + 4).RefractionNVectors.U2 = 0.5: TheFaces(CurFace + 4).RefractionNVectors.V2 = 1
 TheFaces(CurFace + 4).RefractionNVectors.U3 = 0: TheFaces(CurFace + 4).RefractionNVectors.V3 = 0.5
 TheFaces(CurFace + 5).RefractionNVectors.U1 = 0.5: TheFaces(CurFace + 5).RefractionNVectors.V1 = 0.5
 TheFaces(CurFace + 5).RefractionNVectors.U2 = 0.5: TheFaces(CurFace + 5).RefractionNVectors.V2 = 1
 TheFaces(CurFace + 5).RefractionNVectors.U3 = 0.25: TheFaces(CurFace + 5).RefractionNVectors.V3 = 0.5
 TheFaces(CurFace + 6).RefractionNVectors.U1 = 0.75: TheFaces(CurFace + 6).RefractionNVectors.V1 = 0.5
 TheFaces(CurFace + 6).RefractionNVectors.U2 = 0.5: TheFaces(CurFace + 6).RefractionNVectors.V2 = 1
 TheFaces(CurFace + 6).RefractionNVectors.U3 = 0.5: TheFaces(CurFace + 6).RefractionNVectors.V3 = 0.5
 TheFaces(CurFace + 7).RefractionNVectors.U1 = 1: TheFaces(CurFace + 7).RefractionNVectors.V1 = 0.5
 TheFaces(CurFace + 7).RefractionNVectors.U2 = 0.5: TheFaces(CurFace + 7).RefractionNVectors.V2 = 1
 TheFaces(CurFace + 7).RefractionNVectors.U3 = 0.75: TheFaces(CurFace + 7).RefractionNVectors.V3 = 0.5

 If (DoubleSided = True) Then
  'Add faces
  '=========
  TheFaces(CurFace + 8).C = (CurVertex + 0): TheFaces(CurFace + 8).B = (CurVertex + 1): TheFaces(CurFace + 8).A = (CurVertex + 2)
  TheFaces(CurFace + 9).C = (CurVertex + 0): TheFaces(CurFace + 9).B = (CurVertex + 2): TheFaces(CurFace + 9).A = (CurVertex + 3)
  TheFaces(CurFace + 10).C = (CurVertex + 0): TheFaces(CurFace + 10).B = (CurVertex + 3): TheFaces(CurFace + 10).A = (CurVertex + 4)
  TheFaces(CurFace + 11).C = (CurVertex + 0): TheFaces(CurFace + 11).B = (CurVertex + 4): TheFaces(CurFace + 11).A = (CurVertex + 1)
  TheFaces(CurFace + 12).C = (CurVertex + 1): TheFaces(CurFace + 12).B = (CurVertex + 5): TheFaces(CurFace + 12).A = (CurVertex + 2)
  TheFaces(CurFace + 13).C = (CurVertex + 2): TheFaces(CurFace + 13).B = (CurVertex + 5): TheFaces(CurFace + 13).A = (CurVertex + 3)
  TheFaces(CurFace + 14).C = (CurVertex + 3): TheFaces(CurFace + 14).B = (CurVertex + 5): TheFaces(CurFace + 14).A = (CurVertex + 4)
  TheFaces(CurFace + 15).C = (CurVertex + 4): TheFaces(CurFace + 15).B = (CurVertex + 5): TheFaces(CurFace + 15).A = (CurVertex + 1)
  'Setup texture coordinates
  '=========================
  'Alpha texture-coordinates:
  TheFaces(CurFace).AlphaVectors.U3 = 0.5: TheFaces(CurFace).AlphaVectors.V3 = 0
  TheFaces(CurFace).AlphaVectors.U2 = 0.25: TheFaces(CurFace).AlphaVectors.V2 = 0.5
  TheFaces(CurFace).AlphaVectors.U1 = 0: TheFaces(CurFace).AlphaVectors.V1 = 0.5
  TheFaces(CurFace + 1).AlphaVectors.U3 = 0.5: TheFaces(CurFace + 1).AlphaVectors.V3 = 0
  TheFaces(CurFace + 1).AlphaVectors.U2 = 0.5: TheFaces(CurFace + 1).AlphaVectors.V2 = 0.5
  TheFaces(CurFace + 1).AlphaVectors.U1 = 0.25: TheFaces(CurFace + 1).AlphaVectors.V1 = 0.5
  TheFaces(CurFace + 2).AlphaVectors.U3 = 0.5: TheFaces(CurFace + 2).AlphaVectors.V3 = 0
  TheFaces(CurFace + 2).AlphaVectors.U2 = 0.75: TheFaces(CurFace + 2).AlphaVectors.V2 = 0.5
  TheFaces(CurFace + 2).AlphaVectors.U1 = 0.5: TheFaces(CurFace + 2).AlphaVectors.V1 = 0.5
  TheFaces(CurFace + 3).AlphaVectors.U3 = 0.5: TheFaces(CurFace + 3).AlphaVectors.V3 = 0
  TheFaces(CurFace + 3).AlphaVectors.U2 = 1: TheFaces(CurFace + 3).AlphaVectors.V2 = 0.5
  TheFaces(CurFace + 3).AlphaVectors.U1 = 0.75: TheFaces(CurFace + 3).AlphaVectors.V1 = 0.5
  TheFaces(CurFace + 4).AlphaVectors.U3 = 0.25: TheFaces(CurFace + 4).AlphaVectors.V3 = 0.5
  TheFaces(CurFace + 4).AlphaVectors.U2 = 0.5: TheFaces(CurFace + 4).AlphaVectors.V2 = 1
  TheFaces(CurFace + 4).AlphaVectors.U1 = 0: TheFaces(CurFace + 4).AlphaVectors.V1 = 0.5
  TheFaces(CurFace + 5).AlphaVectors.U3 = 0.5: TheFaces(CurFace + 5).AlphaVectors.V3 = 0.5
  TheFaces(CurFace + 5).AlphaVectors.U2 = 0.5: TheFaces(CurFace + 5).AlphaVectors.V2 = 1
  TheFaces(CurFace + 5).AlphaVectors.U1 = 0.25: TheFaces(CurFace + 5).AlphaVectors.V1 = 0.5
  TheFaces(CurFace + 6).AlphaVectors.U3 = 0.75: TheFaces(CurFace + 6).AlphaVectors.V3 = 0.5
  TheFaces(CurFace + 6).AlphaVectors.U2 = 0.5: TheFaces(CurFace + 6).AlphaVectors.V2 = 1
  TheFaces(CurFace + 6).AlphaVectors.U1 = 0.5: TheFaces(CurFace + 6).AlphaVectors.V1 = 0.5
  TheFaces(CurFace + 7).AlphaVectors.U3 = 1: TheFaces(CurFace + 7).AlphaVectors.V3 = 0.5
  TheFaces(CurFace + 7).AlphaVectors.U2 = 0.5: TheFaces(CurFace + 7).AlphaVectors.V2 = 1
  TheFaces(CurFace + 7).AlphaVectors.U1 = 0.75: TheFaces(CurFace + 7).AlphaVectors.V1 = 0.5
  'Color texture-coordinates:
  TheFaces(CurFace).ColorVectors.U3 = 0.5: TheFaces(CurFace).ColorVectors.V3 = 0
  TheFaces(CurFace).ColorVectors.U2 = 0.25: TheFaces(CurFace).ColorVectors.V2 = 0.5
  TheFaces(CurFace).ColorVectors.U1 = 0: TheFaces(CurFace).ColorVectors.V1 = 0.5
  TheFaces(CurFace + 1).ColorVectors.U3 = 0.5: TheFaces(CurFace + 1).ColorVectors.V3 = 0
  TheFaces(CurFace + 1).ColorVectors.U2 = 0.5: TheFaces(CurFace + 1).ColorVectors.V2 = 0.5
  TheFaces(CurFace + 1).ColorVectors.U1 = 0.25: TheFaces(CurFace + 1).ColorVectors.V1 = 0.5
  TheFaces(CurFace + 2).ColorVectors.U3 = 0.5: TheFaces(CurFace + 2).ColorVectors.V3 = 0
  TheFaces(CurFace + 2).ColorVectors.U2 = 0.75: TheFaces(CurFace + 2).ColorVectors.V2 = 0.5
  TheFaces(CurFace + 2).ColorVectors.U1 = 0.5: TheFaces(CurFace + 2).ColorVectors.V1 = 0.5
  TheFaces(CurFace + 3).ColorVectors.U3 = 0.5: TheFaces(CurFace + 3).ColorVectors.V3 = 0
  TheFaces(CurFace + 3).ColorVectors.U2 = 1: TheFaces(CurFace + 3).ColorVectors.V2 = 0.5
  TheFaces(CurFace + 3).ColorVectors.U1 = 0.75: TheFaces(CurFace + 3).ColorVectors.V1 = 0.5
  TheFaces(CurFace + 4).ColorVectors.U3 = 0.25: TheFaces(CurFace + 4).ColorVectors.V3 = 0.5
  TheFaces(CurFace + 4).ColorVectors.U2 = 0.5: TheFaces(CurFace + 4).ColorVectors.V2 = 1
  TheFaces(CurFace + 4).ColorVectors.U1 = 0: TheFaces(CurFace + 4).ColorVectors.V1 = 0.5
  TheFaces(CurFace + 5).ColorVectors.U3 = 0.5: TheFaces(CurFace + 5).ColorVectors.V3 = 0.5
  TheFaces(CurFace + 5).ColorVectors.U2 = 0.5: TheFaces(CurFace + 5).ColorVectors.V2 = 1
  TheFaces(CurFace + 5).ColorVectors.U1 = 0.25: TheFaces(CurFace + 5).ColorVectors.V1 = 0.5
  TheFaces(CurFace + 6).ColorVectors.U3 = 0.75: TheFaces(CurFace + 6).ColorVectors.V3 = 0.5
  TheFaces(CurFace + 6).ColorVectors.U2 = 0.5: TheFaces(CurFace + 6).ColorVectors.V2 = 1
  TheFaces(CurFace + 6).ColorVectors.U1 = 0.5: TheFaces(CurFace + 6).ColorVectors.V1 = 0.5
  TheFaces(CurFace + 7).ColorVectors.U3 = 1: TheFaces(CurFace + 7).ColorVectors.V3 = 0.5
  TheFaces(CurFace + 7).ColorVectors.U2 = 0.5: TheFaces(CurFace + 7).ColorVectors.V2 = 1
  TheFaces(CurFace + 7).ColorVectors.U1 = 0.75: TheFaces(CurFace + 7).ColorVectors.V1 = 0.5
  'Reflection texture-coordinates:
  TheFaces(CurFace).ReflectionVectors.U3 = 0.5: TheFaces(CurFace).ReflectionVectors.V3 = 0
  TheFaces(CurFace).ReflectionVectors.U2 = 0.25: TheFaces(CurFace).ReflectionVectors.V2 = 0.5
  TheFaces(CurFace).ReflectionVectors.U1 = 0: TheFaces(CurFace).ReflectionVectors.V1 = 0.5
  TheFaces(CurFace + 1).ReflectionVectors.U3 = 0.5: TheFaces(CurFace + 1).ReflectionVectors.V3 = 0
  TheFaces(CurFace + 1).ReflectionVectors.U2 = 0.5: TheFaces(CurFace + 1).ReflectionVectors.V2 = 0.5
  TheFaces(CurFace + 1).ReflectionVectors.U1 = 0.25: TheFaces(CurFace + 1).ReflectionVectors.V1 = 0.5
  TheFaces(CurFace + 2).ReflectionVectors.U3 = 0.5: TheFaces(CurFace + 2).ReflectionVectors.V3 = 0
  TheFaces(CurFace + 2).ReflectionVectors.U2 = 0.75: TheFaces(CurFace + 2).ReflectionVectors.V2 = 0.5
  TheFaces(CurFace + 2).ReflectionVectors.U1 = 0.5: TheFaces(CurFace + 2).ReflectionVectors.V1 = 0.5
  TheFaces(CurFace + 3).ReflectionVectors.U3 = 0.5: TheFaces(CurFace + 3).ReflectionVectors.V3 = 0
  TheFaces(CurFace + 3).ReflectionVectors.U2 = 1: TheFaces(CurFace + 3).ReflectionVectors.V2 = 0.5
  TheFaces(CurFace + 3).ReflectionVectors.U1 = 0.75: TheFaces(CurFace + 3).ReflectionVectors.V1 = 0.5
  TheFaces(CurFace + 4).ReflectionVectors.U3 = 0.25: TheFaces(CurFace + 4).ReflectionVectors.V3 = 0.5
  TheFaces(CurFace + 4).ReflectionVectors.U2 = 0.5: TheFaces(CurFace + 4).ReflectionVectors.V2 = 1
  TheFaces(CurFace + 4).ReflectionVectors.U1 = 0: TheFaces(CurFace + 4).ReflectionVectors.V1 = 0.5
  TheFaces(CurFace + 5).ReflectionVectors.U3 = 0.5: TheFaces(CurFace + 5).ReflectionVectors.V3 = 0.5
  TheFaces(CurFace + 5).ReflectionVectors.U2 = 0.5: TheFaces(CurFace + 5).ReflectionVectors.V2 = 1
  TheFaces(CurFace + 5).ReflectionVectors.U1 = 0.25: TheFaces(CurFace + 5).ReflectionVectors.V1 = 0.5
  TheFaces(CurFace + 6).ReflectionVectors.U3 = 0.75: TheFaces(CurFace + 6).ReflectionVectors.V3 = 0.5
  TheFaces(CurFace + 6).ReflectionVectors.U2 = 0.5: TheFaces(CurFace + 6).ReflectionVectors.V2 = 1
  TheFaces(CurFace + 6).ReflectionVectors.U1 = 0.5: TheFaces(CurFace + 6).ReflectionVectors.V1 = 0.5
  TheFaces(CurFace + 7).ReflectionVectors.U3 = 1: TheFaces(CurFace + 7).ReflectionVectors.V3 = 0.5
  TheFaces(CurFace + 7).ReflectionVectors.U2 = 0.5: TheFaces(CurFace + 7).ReflectionVectors.V2 = 1
  TheFaces(CurFace + 7).ReflectionVectors.U1 = 0.75: TheFaces(CurFace + 7).ReflectionVectors.V1 = 0.5
  'Refraction texture-coordinates:
  TheFaces(CurFace).RefractionVectors.U3 = 0.5: TheFaces(CurFace).RefractionVectors.V3 = 0
  TheFaces(CurFace).RefractionVectors.U2 = 0.25: TheFaces(CurFace).RefractionVectors.V2 = 0.5
  TheFaces(CurFace).RefractionVectors.U1 = 0: TheFaces(CurFace).RefractionVectors.V1 = 0.5
  TheFaces(CurFace + 1).RefractionVectors.U3 = 0.5: TheFaces(CurFace + 1).RefractionVectors.V3 = 0
  TheFaces(CurFace + 1).RefractionVectors.U2 = 0.5: TheFaces(CurFace + 1).RefractionVectors.V2 = 0.5
  TheFaces(CurFace + 1).RefractionVectors.U1 = 0.25: TheFaces(CurFace + 1).RefractionVectors.V1 = 0.5
  TheFaces(CurFace + 2).RefractionVectors.U3 = 0.5: TheFaces(CurFace + 2).RefractionVectors.V3 = 0
  TheFaces(CurFace + 2).RefractionVectors.U2 = 0.75: TheFaces(CurFace + 2).RefractionVectors.V2 = 0.5
  TheFaces(CurFace + 2).RefractionVectors.U1 = 0.5: TheFaces(CurFace + 2).RefractionVectors.V1 = 0.5
  TheFaces(CurFace + 3).RefractionVectors.U3 = 0.5: TheFaces(CurFace + 3).RefractionVectors.V3 = 0
  TheFaces(CurFace + 3).RefractionVectors.U2 = 1: TheFaces(CurFace + 3).RefractionVectors.V2 = 0.5
  TheFaces(CurFace + 3).RefractionVectors.U1 = 0.75: TheFaces(CurFace + 3).RefractionVectors.V1 = 0.5
  TheFaces(CurFace + 4).RefractionVectors.U3 = 0.25: TheFaces(CurFace + 4).RefractionVectors.V3 = 0.5
  TheFaces(CurFace + 4).RefractionVectors.U2 = 0.5: TheFaces(CurFace + 4).RefractionVectors.V2 = 1
  TheFaces(CurFace + 4).RefractionVectors.U1 = 0: TheFaces(CurFace + 4).RefractionVectors.V1 = 0.5
  TheFaces(CurFace + 5).RefractionVectors.U3 = 0.5: TheFaces(CurFace + 5).RefractionVectors.V3 = 0.5
  TheFaces(CurFace + 5).RefractionVectors.U2 = 0.5: TheFaces(CurFace + 5).RefractionVectors.V2 = 1
  TheFaces(CurFace + 5).RefractionVectors.U1 = 0.25: TheFaces(CurFace + 5).RefractionVectors.V1 = 0.5
  TheFaces(CurFace + 6).RefractionVectors.U3 = 0.75: TheFaces(CurFace + 6).RefractionVectors.V3 = 0.5
  TheFaces(CurFace + 6).RefractionVectors.U2 = 0.5: TheFaces(CurFace + 6).RefractionVectors.V2 = 1
  TheFaces(CurFace + 6).RefractionVectors.U1 = 0.5: TheFaces(CurFace + 6).RefractionVectors.V1 = 0.5
  TheFaces(CurFace + 7).RefractionVectors.U3 = 1: TheFaces(CurFace + 7).RefractionVectors.V3 = 0.5
  TheFaces(CurFace + 7).RefractionVectors.U2 = 0.5: TheFaces(CurFace + 7).RefractionVectors.V2 = 1
  TheFaces(CurFace + 7).RefractionVectors.U1 = 0.75: TheFaces(CurFace + 7).RefractionVectors.V1 = 0.5
  'RefractionN texture-coordinates:
  TheFaces(CurFace).RefractionNVectors.U3 = 0.5: TheFaces(CurFace).RefractionNVectors.V3 = 0
  TheFaces(CurFace).RefractionNVectors.U2 = 0.25: TheFaces(CurFace).RefractionNVectors.V2 = 0.5
  TheFaces(CurFace).RefractionNVectors.U1 = 0: TheFaces(CurFace).RefractionNVectors.V1 = 0.5
  TheFaces(CurFace + 1).RefractionNVectors.U3 = 0.5: TheFaces(CurFace + 1).RefractionNVectors.V3 = 0
  TheFaces(CurFace + 1).RefractionNVectors.U2 = 0.5: TheFaces(CurFace + 1).RefractionNVectors.V2 = 0.5
  TheFaces(CurFace + 1).RefractionNVectors.U1 = 0.25: TheFaces(CurFace + 1).RefractionNVectors.V1 = 0.5
  TheFaces(CurFace + 2).RefractionNVectors.U3 = 0.5: TheFaces(CurFace + 2).RefractionNVectors.V3 = 0
  TheFaces(CurFace + 2).RefractionNVectors.U2 = 0.75: TheFaces(CurFace + 2).RefractionNVectors.V2 = 0.5
  TheFaces(CurFace + 2).RefractionNVectors.U1 = 0.5: TheFaces(CurFace + 2).RefractionNVectors.V1 = 0.5
  TheFaces(CurFace + 3).RefractionNVectors.U3 = 0.5: TheFaces(CurFace + 3).RefractionNVectors.V3 = 0
  TheFaces(CurFace + 3).RefractionNVectors.U2 = 1: TheFaces(CurFace + 3).RefractionNVectors.V2 = 0.5
  TheFaces(CurFace + 3).RefractionNVectors.U1 = 0.75: TheFaces(CurFace + 3).RefractionNVectors.V1 = 0.5
  TheFaces(CurFace + 4).RefractionNVectors.U3 = 0.25: TheFaces(CurFace + 4).RefractionNVectors.V3 = 0.5
  TheFaces(CurFace + 4).RefractionNVectors.U2 = 0.5: TheFaces(CurFace + 4).RefractionNVectors.V2 = 1
  TheFaces(CurFace + 4).RefractionNVectors.U1 = 0: TheFaces(CurFace + 4).RefractionNVectors.V1 = 0.5
  TheFaces(CurFace + 5).RefractionNVectors.U3 = 0.5: TheFaces(CurFace + 5).RefractionNVectors.V3 = 0.5
  TheFaces(CurFace + 5).RefractionNVectors.U2 = 0.5: TheFaces(CurFace + 5).RefractionNVectors.V2 = 1
  TheFaces(CurFace + 5).RefractionNVectors.U1 = 0.25: TheFaces(CurFace + 5).RefractionNVectors.V1 = 0.5
  TheFaces(CurFace + 6).RefractionNVectors.U3 = 0.75: TheFaces(CurFace + 6).RefractionNVectors.V3 = 0.5
  TheFaces(CurFace + 6).RefractionNVectors.U2 = 0.5: TheFaces(CurFace + 6).RefractionNVectors.V2 = 1
  TheFaces(CurFace + 6).RefractionNVectors.U1 = 0.5: TheFaces(CurFace + 6).RefractionNVectors.V1 = 0.5
  TheFaces(CurFace + 7).RefractionNVectors.U3 = 1: TheFaces(CurFace + 7).RefractionNVectors.V3 = 0.5
  TheFaces(CurFace + 7).RefractionNVectors.U2 = 0.5: TheFaces(CurFace + 7).RefractionNVectors.V2 = 1
  TheFaces(CurFace + 7).RefractionNVectors.U1 = 0.75: TheFaces(CurFace + 7).RefractionNVectors.V1 = 0.5
 End If

End Function
Function Primitive_Sphere(CreateAsDefault As Boolean, Radius!, Steps1&, Steps2&, DoubleSided As Boolean) As Long

 Primitive_Sphere = -1

 If (CreateAsDefault = False) Then
  If ((Radius <= 0) Or (Steps2 < 2) Or (Steps1 < 2)) Then Exit Function
 ElseIf (CreateAsDefault = True) Then
  Radius = 100: Steps1 = 10: Steps2 = 10: DoubleSided = False
 End If

 Dim AShapeSpline() As Vector3D, AMoveSpline() As Vector3D, CurVertex&, Tmp!

 Spline_Arc AShapeSpline(), 0.5, Radius, 0, 0, 1, Steps1

 Spline_Circle AMoveSpline(), ApproachVal, 0, 0, 2, Steps2
 ReDim Preserve AMoveSpline(UBound(AMoveSpline) + 1)
 AMoveSpline(UBound(AMoveSpline)) = AMoveSpline(1)

 Primitive_Sphere = Primitive_MakeFromSplines(AShapeSpline(), AMoveSpline(), 0.5, 0, 0, 1, 1, 1, 1, DoubleSided, True)
 If (Primitive_Sphere <> -1) Then TheMeshs(Primitive_Sphere).Label = "Sphere_" & CStr(TheMeshsCount)

End Function
Function Primitive_Tetrahedron(CreateAsDefault As Boolean, Base!, Radius!, DoubleSided As Boolean) As Long

 Primitive_Tetrahedron = -1

 If (CreateAsDefault = False) Then
  If ((Base <= 0) Or (Radius <= 0)) Then Exit Function
 ElseIf (CreateAsDefault = True) Then
  Base = 100: Radius = 100: DoubleSided = False
 End If

 Dim CurVertex&, CurFace&

 If (DoubleSided = False) Then
  Primitive_Tetrahedron = Mesh3D_Add(4, 4)
 Else
  Primitive_Tetrahedron = Mesh3D_Add(4, 8)
 End If

 If (Primitive_Tetrahedron = -1) Then Exit Function

 TheMeshs(Primitive_Tetrahedron).Label = "Tetrahedron_" & CStr(TheMeshsCount)
 CurVertex = TheMeshs(Primitive_Tetrahedron).Vertices.Start
 CurFace = TheMeshs(Primitive_Tetrahedron).Faces.Start

 'Add vertices
 '============
 TheVertices(CurVertex).Position = VectorInput(0, -Base, 0)
 TheVertices(CurVertex + 1).Position = VectorInput(0, 0, -Radius)
 TheVertices(CurVertex + 2).Position = VectorRotate(VectorInput(0, 0, -Radius), 1, (Pi2 * OneByThree))
 TheVertices(CurVertex + 3).Position = VectorRotate(VectorInput(0, 0, -Radius), 1, ((Pi2 * OneByThree) * 2))

 'Add faces
 '=========
 TheFaces(CurFace).A = (CurVertex + 0): TheFaces(CurFace).B = (CurVertex + 1): TheFaces(CurFace).C = (CurVertex + 2)
 TheFaces(CurFace + 1).A = (CurVertex + 0): TheFaces(CurFace + 1).B = (CurVertex + 2): TheFaces(CurFace + 1).C = (CurVertex + 3)
 TheFaces(CurFace + 2).A = (CurVertex + 0): TheFaces(CurFace + 2).B = (CurVertex + 3): TheFaces(CurFace + 2).C = (CurVertex + 1)
 TheFaces(CurFace + 3).A = (CurVertex + 1): TheFaces(CurFace + 3).B = (CurVertex + 2): TheFaces(CurFace + 3).C = (CurVertex + 3)

 'Setup texture coordinates
 '=========================
 'Alpha texture-coordinates:
 TheFaces(CurFace).AlphaVectors.U1 = 0.5: TheFaces(CurFace).AlphaVectors.V1 = 0
 TheFaces(CurFace).AlphaVectors.U2 = OneByThree: TheFaces(CurFace).AlphaVectors.V2 = 0.5
 TheFaces(CurFace).AlphaVectors.U3 = 0: TheFaces(CurFace).AlphaVectors.V3 = 0.5
 TheFaces(CurFace + 1).AlphaVectors.U1 = 0.5: TheFaces(CurFace + 1).AlphaVectors.V1 = 0
 TheFaces(CurFace + 1).AlphaVectors.U2 = (OneByThree * 2): TheFaces(CurFace + 1).AlphaVectors.V2 = 0.5
 TheFaces(CurFace + 1).AlphaVectors.U3 = OneByThree: TheFaces(CurFace + 1).AlphaVectors.V3 = 0.5
 TheFaces(CurFace + 2).AlphaVectors.U1 = 0.5: TheFaces(CurFace + 2).AlphaVectors.V1 = 0
 TheFaces(CurFace + 2).AlphaVectors.U2 = 1: TheFaces(CurFace + 2).AlphaVectors.V2 = 0.5
 TheFaces(CurFace + 2).AlphaVectors.U3 = (OneByThree * 2): TheFaces(CurFace + 2).AlphaVectors.V3 = 0.5
 TheFaces(CurFace + 3).AlphaVectors.U1 = 0: TheFaces(CurFace + 3).AlphaVectors.V1 = 0.5
 TheFaces(CurFace + 3).AlphaVectors.U2 = 1: TheFaces(CurFace + 3).AlphaVectors.V2 = 0.5
 TheFaces(CurFace + 3).AlphaVectors.U3 = 0.5: TheFaces(CurFace + 3).AlphaVectors.V3 = 1
 'Color texture-coordinates:
 TheFaces(CurFace).ColorVectors.U1 = 0.5: TheFaces(CurFace).ColorVectors.V1 = 0
 TheFaces(CurFace).ColorVectors.U2 = OneByThree: TheFaces(CurFace).ColorVectors.V2 = 0.5
 TheFaces(CurFace).ColorVectors.U3 = 0: TheFaces(CurFace).ColorVectors.V3 = 0.5
 TheFaces(CurFace + 1).ColorVectors.U1 = 0.5: TheFaces(CurFace + 1).ColorVectors.V1 = 0
 TheFaces(CurFace + 1).ColorVectors.U2 = (OneByThree * 2): TheFaces(CurFace + 1).ColorVectors.V2 = 0.5
 TheFaces(CurFace + 1).ColorVectors.U3 = OneByThree: TheFaces(CurFace + 1).ColorVectors.V3 = 0.5
 TheFaces(CurFace + 2).ColorVectors.U1 = 0.5: TheFaces(CurFace + 2).ColorVectors.V1 = 0
 TheFaces(CurFace + 2).ColorVectors.U2 = 1: TheFaces(CurFace + 2).ColorVectors.V2 = 0.5
 TheFaces(CurFace + 2).ColorVectors.U3 = (OneByThree * 2): TheFaces(CurFace + 2).ColorVectors.V3 = 0.5
 TheFaces(CurFace + 3).ColorVectors.U1 = 0: TheFaces(CurFace + 3).ColorVectors.V1 = 0.5
 TheFaces(CurFace + 3).ColorVectors.U2 = 1: TheFaces(CurFace + 3).ColorVectors.V2 = 0.5
 TheFaces(CurFace + 3).ColorVectors.U3 = 0.5: TheFaces(CurFace + 3).ColorVectors.V3 = 1
 'Reflection texture-coordinates:
 TheFaces(CurFace).ReflectionVectors.U1 = 0.5: TheFaces(CurFace).ReflectionVectors.V1 = 0
 TheFaces(CurFace).ReflectionVectors.U2 = OneByThree: TheFaces(CurFace).ReflectionVectors.V2 = 0.5
 TheFaces(CurFace).ReflectionVectors.U3 = 0: TheFaces(CurFace).ReflectionVectors.V3 = 0.5
 TheFaces(CurFace + 1).ReflectionVectors.U1 = 0.5: TheFaces(CurFace + 1).ReflectionVectors.V1 = 0
 TheFaces(CurFace + 1).ReflectionVectors.U2 = (OneByThree * 2): TheFaces(CurFace + 1).ReflectionVectors.V2 = 0.5
 TheFaces(CurFace + 1).ReflectionVectors.U3 = OneByThree: TheFaces(CurFace + 1).ReflectionVectors.V3 = 0.5
 TheFaces(CurFace + 2).ReflectionVectors.U1 = 0.5: TheFaces(CurFace + 2).ReflectionVectors.V1 = 0
 TheFaces(CurFace + 2).ReflectionVectors.U2 = 1: TheFaces(CurFace + 2).ReflectionVectors.V2 = 0.5
 TheFaces(CurFace + 2).ReflectionVectors.U3 = (OneByThree * 2): TheFaces(CurFace + 2).ReflectionVectors.V3 = 0.5
 TheFaces(CurFace + 3).ReflectionVectors.U1 = 0: TheFaces(CurFace + 3).ReflectionVectors.V1 = 0.5
 TheFaces(CurFace + 3).ReflectionVectors.U2 = 1: TheFaces(CurFace + 3).ReflectionVectors.V2 = 0.5
 TheFaces(CurFace + 3).ReflectionVectors.U3 = 0.5: TheFaces(CurFace + 3).ReflectionVectors.V3 = 1
 'Refraction texture-coordinates:
 TheFaces(CurFace).RefractionVectors.U1 = 0.5: TheFaces(CurFace).RefractionVectors.V1 = 0
 TheFaces(CurFace).RefractionVectors.U2 = OneByThree: TheFaces(CurFace).RefractionVectors.V2 = 0.5
 TheFaces(CurFace).RefractionVectors.U3 = 0: TheFaces(CurFace).RefractionVectors.V3 = 0.5
 TheFaces(CurFace + 1).RefractionVectors.U1 = 0.5: TheFaces(CurFace + 1).RefractionVectors.V1 = 0
 TheFaces(CurFace + 1).RefractionVectors.U2 = (OneByThree * 2): TheFaces(CurFace + 1).RefractionVectors.V2 = 0.5
 TheFaces(CurFace + 1).RefractionVectors.U3 = OneByThree: TheFaces(CurFace + 1).RefractionVectors.V3 = 0.5
 TheFaces(CurFace + 2).RefractionVectors.U1 = 0.5: TheFaces(CurFace + 2).RefractionVectors.V1 = 0
 TheFaces(CurFace + 2).RefractionVectors.U2 = 1: TheFaces(CurFace + 2).RefractionVectors.V2 = 0.5
 TheFaces(CurFace + 2).RefractionVectors.U3 = (OneByThree * 2): TheFaces(CurFace + 2).RefractionVectors.V3 = 0.5
 TheFaces(CurFace + 3).RefractionVectors.U1 = 0: TheFaces(CurFace + 3).RefractionVectors.V1 = 0.5
 TheFaces(CurFace + 3).RefractionVectors.U2 = 1: TheFaces(CurFace + 3).RefractionVectors.V2 = 0.5
 TheFaces(CurFace + 3).RefractionVectors.U3 = 0.5: TheFaces(CurFace + 3).RefractionVectors.V3 = 1
 'RefractionN texture-coordinates:
 TheFaces(CurFace).RefractionNVectors.U1 = 0.5: TheFaces(CurFace).RefractionNVectors.V1 = 0
 TheFaces(CurFace).RefractionNVectors.U2 = OneByThree: TheFaces(CurFace).RefractionNVectors.V2 = 0.5
 TheFaces(CurFace).RefractionNVectors.U3 = 0: TheFaces(CurFace).RefractionNVectors.V3 = 0.5
 TheFaces(CurFace + 1).RefractionNVectors.U1 = 0.5: TheFaces(CurFace + 1).RefractionNVectors.V1 = 0
 TheFaces(CurFace + 1).RefractionNVectors.U2 = (OneByThree * 2): TheFaces(CurFace + 1).RefractionNVectors.V2 = 0.5
 TheFaces(CurFace + 1).RefractionNVectors.U3 = OneByThree: TheFaces(CurFace + 1).RefractionNVectors.V3 = 0.5
 TheFaces(CurFace + 2).RefractionNVectors.U1 = 0.5: TheFaces(CurFace + 2).RefractionNVectors.V1 = 0
 TheFaces(CurFace + 2).RefractionNVectors.U2 = 1: TheFaces(CurFace + 2).RefractionNVectors.V2 = 0.5
 TheFaces(CurFace + 2).RefractionNVectors.U3 = (OneByThree * 2): TheFaces(CurFace + 2).RefractionNVectors.V3 = 0.5
 TheFaces(CurFace + 3).RefractionNVectors.U1 = 0: TheFaces(CurFace + 3).RefractionNVectors.V1 = 0.5
 TheFaces(CurFace + 3).RefractionNVectors.U2 = 1: TheFaces(CurFace + 3).RefractionNVectors.V2 = 0.5
 TheFaces(CurFace + 3).RefractionNVectors.U3 = 0.5: TheFaces(CurFace + 3).RefractionNVectors.V3 = 1

 If (DoubleSided = True) Then
  'Add faces
  '=========
  TheFaces(CurFace + 4).C = (CurVertex + 0): TheFaces(CurFace + 4).B = (CurVertex + 1): TheFaces(CurFace + 4).A = (CurVertex + 2)
  TheFaces(CurFace + 5).C = (CurVertex + 0): TheFaces(CurFace + 5).B = (CurVertex + 2): TheFaces(CurFace + 5).A = (CurVertex + 3)
  TheFaces(CurFace + 6).C = (CurVertex + 0): TheFaces(CurFace + 6).B = (CurVertex + 3): TheFaces(CurFace + 6).A = (CurVertex + 1)
  TheFaces(CurFace + 7).C = (CurVertex + 1): TheFaces(CurFace + 7).B = (CurVertex + 2): TheFaces(CurFace + 7).A = (CurVertex + 3)
  'Setup texture coordinates
  '=========================
  'Alpha texture-coordinates:
  TheFaces(CurFace + 4).AlphaVectors.U3 = 0.5: TheFaces(CurFace + 4).AlphaVectors.V3 = 0
  TheFaces(CurFace + 4).AlphaVectors.U2 = OneByThree: TheFaces(CurFace + 4).AlphaVectors.V2 = 0.5
  TheFaces(CurFace + 4).AlphaVectors.U1 = 0: TheFaces(CurFace + 4).AlphaVectors.V1 = 0.5
  TheFaces(CurFace + 5).AlphaVectors.U3 = 0.5: TheFaces(CurFace + 5).AlphaVectors.V3 = 0
  TheFaces(CurFace + 5).AlphaVectors.U2 = (OneByThree * 2): TheFaces(CurFace + 5).AlphaVectors.V2 = 0.5
  TheFaces(CurFace + 5).AlphaVectors.U1 = OneByThree: TheFaces(CurFace + 5).AlphaVectors.V1 = 0.5
  TheFaces(CurFace + 6).AlphaVectors.U3 = 0.5: TheFaces(CurFace + 6).AlphaVectors.V3 = 0
  TheFaces(CurFace + 6).AlphaVectors.U2 = 1: TheFaces(CurFace + 6).AlphaVectors.V2 = 0.5
  TheFaces(CurFace + 6).AlphaVectors.U1 = (OneByThree * 2): TheFaces(CurFace + 6).AlphaVectors.V1 = 0.5
  TheFaces(CurFace + 7).AlphaVectors.U3 = 0: TheFaces(CurFace + 7).AlphaVectors.V3 = 0.5
  TheFaces(CurFace + 7).AlphaVectors.U2 = 1: TheFaces(CurFace + 7).AlphaVectors.V2 = 0.5
  TheFaces(CurFace + 7).AlphaVectors.U1 = 0.5: TheFaces(CurFace + 7).AlphaVectors.V1 = 1
  'Color texture-coordinates:
  TheFaces(CurFace + 4).ColorVectors.U3 = 0.5: TheFaces(CurFace + 4).ColorVectors.V3 = 0
  TheFaces(CurFace + 4).ColorVectors.U2 = OneByThree: TheFaces(CurFace + 4).ColorVectors.V2 = 0.5
  TheFaces(CurFace + 4).ColorVectors.U1 = 0: TheFaces(CurFace + 4).ColorVectors.V1 = 0.5
  TheFaces(CurFace + 5).ColorVectors.U3 = 0.5: TheFaces(CurFace + 5).ColorVectors.V3 = 0
  TheFaces(CurFace + 5).ColorVectors.U2 = (OneByThree * 2): TheFaces(CurFace + 5).ColorVectors.V2 = 0.5
  TheFaces(CurFace + 5).ColorVectors.U1 = OneByThree: TheFaces(CurFace + 5).ColorVectors.V1 = 0.5
  TheFaces(CurFace + 6).ColorVectors.U3 = 0.5: TheFaces(CurFace + 6).ColorVectors.V3 = 0
  TheFaces(CurFace + 6).ColorVectors.U2 = 1: TheFaces(CurFace + 6).ColorVectors.V2 = 0.5
  TheFaces(CurFace + 6).ColorVectors.U1 = (OneByThree * 2): TheFaces(CurFace + 6).ColorVectors.V1 = 0.5
  TheFaces(CurFace + 7).ColorVectors.U3 = 0: TheFaces(CurFace + 7).ColorVectors.V3 = 0.5
  TheFaces(CurFace + 7).ColorVectors.U2 = 1: TheFaces(CurFace + 7).ColorVectors.V2 = 0.5
  TheFaces(CurFace + 7).ColorVectors.U1 = 0.5: TheFaces(CurFace + 7).ColorVectors.V1 = 1
  'Reflection texture-coordinates:
  TheFaces(CurFace + 4).ReflectionVectors.U3 = 0.5: TheFaces(CurFace + 4).ReflectionVectors.V3 = 0
  TheFaces(CurFace + 4).ReflectionVectors.U2 = OneByThree: TheFaces(CurFace + 4).ReflectionVectors.V2 = 0.5
  TheFaces(CurFace + 4).ReflectionVectors.U1 = 0: TheFaces(CurFace + 4).ReflectionVectors.V1 = 0.5
  TheFaces(CurFace + 5).ReflectionVectors.U3 = 0.5: TheFaces(CurFace + 5).ReflectionVectors.V3 = 0
  TheFaces(CurFace + 5).ReflectionVectors.U2 = (OneByThree * 2): TheFaces(CurFace + 5).ReflectionVectors.V2 = 0.5
  TheFaces(CurFace + 5).ReflectionVectors.U1 = OneByThree: TheFaces(CurFace + 5).ReflectionVectors.V1 = 0.5
  TheFaces(CurFace + 6).ReflectionVectors.U3 = 0.5: TheFaces(CurFace + 6).ReflectionVectors.V3 = 0
  TheFaces(CurFace + 6).ReflectionVectors.U2 = 1: TheFaces(CurFace + 6).ReflectionVectors.V2 = 0.5
  TheFaces(CurFace + 6).ReflectionVectors.U1 = (OneByThree * 2): TheFaces(CurFace + 6).ReflectionVectors.V1 = 0.5
  TheFaces(CurFace + 7).ReflectionVectors.U3 = 0: TheFaces(CurFace + 7).ReflectionVectors.V3 = 0.5
  TheFaces(CurFace + 7).ReflectionVectors.U2 = 1: TheFaces(CurFace + 7).ReflectionVectors.V2 = 0.5
  TheFaces(CurFace + 7).ReflectionVectors.U1 = 0.5: TheFaces(CurFace + 7).ReflectionVectors.V1 = 1
  'Refraction texture-coordinates:
  TheFaces(CurFace + 4).RefractionVectors.U3 = 0.5: TheFaces(CurFace + 4).RefractionVectors.V3 = 0
  TheFaces(CurFace + 4).RefractionVectors.U2 = OneByThree: TheFaces(CurFace + 4).RefractionVectors.V2 = 0.5
  TheFaces(CurFace + 4).RefractionVectors.U1 = 0: TheFaces(CurFace + 4).RefractionVectors.V1 = 0.5
  TheFaces(CurFace + 5).RefractionVectors.U3 = 0.5: TheFaces(CurFace + 5).RefractionVectors.V3 = 0
  TheFaces(CurFace + 5).RefractionVectors.U2 = (OneByThree * 2): TheFaces(CurFace + 5).RefractionVectors.V2 = 0.5
  TheFaces(CurFace + 5).RefractionVectors.U1 = OneByThree: TheFaces(CurFace + 5).RefractionVectors.V1 = 0.5
  TheFaces(CurFace + 6).RefractionVectors.U3 = 0.5: TheFaces(CurFace + 6).RefractionVectors.V3 = 0
  TheFaces(CurFace + 6).RefractionVectors.U2 = 1: TheFaces(CurFace + 6).RefractionVectors.V2 = 0.5
  TheFaces(CurFace + 6).RefractionVectors.U1 = (OneByThree * 2): TheFaces(CurFace + 6).RefractionVectors.V1 = 0.5
  TheFaces(CurFace + 7).RefractionVectors.U3 = 0: TheFaces(CurFace + 7).RefractionVectors.V3 = 0.5
  TheFaces(CurFace + 7).RefractionVectors.U2 = 1: TheFaces(CurFace + 7).RefractionVectors.V2 = 0.5
  TheFaces(CurFace + 7).RefractionVectors.U1 = 0.5: TheFaces(CurFace + 7).RefractionVectors.V1 = 1
  'RefractionN texture-coordinates:
  TheFaces(CurFace + 4).RefractionNVectors.U3 = 0.5: TheFaces(CurFace + 4).RefractionNVectors.V3 = 0
  TheFaces(CurFace + 4).RefractionNVectors.U2 = OneByThree: TheFaces(CurFace + 4).RefractionNVectors.V2 = 0.5
  TheFaces(CurFace + 4).RefractionNVectors.U1 = 0: TheFaces(CurFace + 4).RefractionNVectors.V1 = 0.5
  TheFaces(CurFace + 5).RefractionNVectors.U3 = 0.5: TheFaces(CurFace + 5).RefractionNVectors.V3 = 0
  TheFaces(CurFace + 5).RefractionNVectors.U2 = (OneByThree * 2): TheFaces(CurFace + 5).RefractionNVectors.V2 = 0.5
  TheFaces(CurFace + 5).RefractionNVectors.U1 = OneByThree: TheFaces(CurFace + 5).RefractionNVectors.V1 = 0.5
  TheFaces(CurFace + 6).RefractionNVectors.U3 = 0.5: TheFaces(CurFace + 6).RefractionNVectors.V3 = 0
  TheFaces(CurFace + 6).RefractionNVectors.U2 = 1: TheFaces(CurFace + 6).RefractionNVectors.V2 = 0.5
  TheFaces(CurFace + 6).RefractionNVectors.U1 = (OneByThree * 2): TheFaces(CurFace + 6).RefractionNVectors.V1 = 0.5
  TheFaces(CurFace + 7).RefractionNVectors.U3 = 0: TheFaces(CurFace + 7).RefractionNVectors.V3 = 0.5
  TheFaces(CurFace + 7).RefractionNVectors.U2 = 1: TheFaces(CurFace + 7).RefractionNVectors.V2 = 0.5
  TheFaces(CurFace + 7).RefractionNVectors.U1 = 0.5: TheFaces(CurFace + 7).RefractionNVectors.V1 = 1
 End If

End Function
Function Primitive_Torus(CreateAsDefault As Boolean, Radius1!, Radius2!, Steps1&, Steps2&, Axis As Byte, DoubleSided As Boolean) As Long

 Primitive_Torus = -1

 If (CreateAsDefault = False) Then
  If ((Radius1 <= 0) Or (Radius1 <= Radius2)) Then Exit Function
  If ((Steps1 < 2) Or (Steps2 < 2) Or (Axis < 0) Or (Axis > 2)) Then Exit Function
 ElseIf (CreateAsDefault = True) Then
  Radius1 = 100: Radius2 = 40: Steps1 = 10: Steps2 = 10: Axis = 1: DoubleSided = False
 End If

 Dim AShapeSpline() As Vector3D, AMoveSpline() As Vector3D, CurVertex&, Tmp!

 Spline_Circle AShapeSpline(), Radius2, Radius1, 0, 1, Steps1

 Spline_Circle AMoveSpline(), ApproachVal, 0, 0, 2, Steps2
 ReDim Preserve AMoveSpline(UBound(AMoveSpline) + 1)
 AMoveSpline(UBound(AMoveSpline)) = AMoveSpline(1)

 Primitive_Torus = Primitive_MakeFromSplines(AShapeSpline(), AMoveSpline(), 0.5, 0, 0, 1, 1, 1, 1, DoubleSided, True)
 If (Primitive_Torus = -1) Then Exit Function

 TheMeshs(Primitive_Torus).Label = "Torus_" & CStr(TheMeshsCount)
 Select Case Axis
  Case 0:
   For CurVertex = TheMeshs(Primitive_Torus).Vertices.Start To GetAddressLast(TheMeshs(Primitive_Torus).Vertices)
    'Swap X to Z:
    Tmp = TheVertices(CurVertex).Position.X
    TheVertices(CurVertex).Position.X = TheVertices(CurVertex).Position.Z
    TheVertices(CurVertex).Position.Z = Tmp
    'Swap Z to Y:
    Tmp = TheVertices(CurVertex).Position.Z
    TheVertices(CurVertex).Position.Z = TheVertices(CurVertex).Position.Y
    TheVertices(CurVertex).Position.Y = Tmp
   Next CurVertex
  Case 1:
   For CurVertex = TheMeshs(Primitive_Torus).Vertices.Start To GetAddressLast(TheMeshs(Primitive_Torus).Vertices)
    'Swap Z to Y:
    Tmp = TheVertices(CurVertex).Position.Z
    TheVertices(CurVertex).Position.Z = TheVertices(CurVertex).Position.Y
    TheVertices(CurVertex).Position.Y = Tmp
   Next CurVertex
 End Select

End Function
Function Primitive_TubeFlat(CreateAsDefault As Boolean, Base!, Radius!, Steps&, Axis As Byte, DoubleSided As Boolean) As Long

 Primitive_TubeFlat = -1

 If (CreateAsDefault = False) Then
  If ((Base <= 0) Or (Radius <= 0)) Then Exit Function
  If ((Steps < 2) Or (Axis < 0) Or (Axis > 2)) Then Exit Function
 ElseIf (CreateAsDefault = True) Then
  Base = 100: Radius = 50: Steps = 10: Axis = 1: DoubleSided = False
 End If

 Dim AShapeSpline() As Vector3D, AMoveSpline() As Vector3D, CurVertex&, Tmp!

 ReDim AShapeSpline(1)
 AShapeSpline(0) = VectorInput(Radius, 0, -(Base * 0.5))
 AShapeSpline(1) = VectorInput(Radius, 0, (Base * 0.5))

 Spline_Circle AMoveSpline(), ApproachVal, 0, 0, 2, Steps
 ReDim Preserve AMoveSpline(UBound(AMoveSpline) + 1)
 AMoveSpline(UBound(AMoveSpline)) = AMoveSpline(1)

 Primitive_TubeFlat = Primitive_MakeFromSplines(AShapeSpline(), AMoveSpline(), 0.5, 0, 0, 1, 1, 1, 1, DoubleSided, True)
 If (Primitive_TubeFlat = -1) Then Exit Function

 TheMeshs(Primitive_TubeFlat).Label = "FlatTube_" & CStr(TheMeshsCount)
 Select Case Axis
  Case 0:
   For CurVertex = TheMeshs(Primitive_TubeFlat).Vertices.Start To GetAddressLast(TheMeshs(Primitive_TubeFlat).Vertices)
    'Swap X to Z:
    Tmp = TheVertices(CurVertex).Position.X
    TheVertices(CurVertex).Position.X = TheVertices(CurVertex).Position.Z
    TheVertices(CurVertex).Position.Z = Tmp
    'Swap Z to Y:
    Tmp = TheVertices(CurVertex).Position.Z
    TheVertices(CurVertex).Position.Z = TheVertices(CurVertex).Position.Y
    TheVertices(CurVertex).Position.Y = Tmp
   Next CurVertex
  Case 1:
   For CurVertex = TheMeshs(Primitive_TubeFlat).Vertices.Start To GetAddressLast(TheMeshs(Primitive_TubeFlat).Vertices)
    'Swap Z to Y:
    Tmp = TheVertices(CurVertex).Position.Z
    TheVertices(CurVertex).Position.Z = TheVertices(CurVertex).Position.Y
    TheVertices(CurVertex).Position.Y = Tmp
   Next CurVertex
 End Select

End Function
Function Primitive_Capsule(CreateAsDefault As Boolean, Base!, Radius!, Steps1&, Steps2&, Axis As Byte, DoubleSided As Boolean) As Long

 Primitive_Capsule = -1

 If (CreateAsDefault = False) Then
  If ((Base <= 0) Or (Radius <= 0)) Then Exit Function
  If ((Steps1 < 2) Or (Steps2 < 2) Or (Axis < 0) Or (Axis > 2)) Then Exit Function
 ElseIf (CreateAsDefault = True) Then
  Base = 50: Radius = 25: Steps1 = 10: Steps2 = 10: Axis = 1: DoubleSided = False
 End If

 Dim AShapeSpline() As Vector3D, TmpShapeSpline() As Vector3D
 Dim AMoveSpline() As Vector3D, CurVertex&, Tmp!

 Spline_Arc AShapeSpline(), 0.25, Radius, 0, 0, 1, (Steps2 * 0.5)
 For CurVertex = 0 To UBound(AShapeSpline())
  AShapeSpline(CurVertex).Z = (-AShapeSpline(CurVertex).Z - (Base * 0.5))
 Next CurVertex
 ReDim Preserve AShapeSpline(UBound(AShapeSpline()) + 1)
 AShapeSpline(UBound(AShapeSpline())) = VectorInput(Radius, 0, (Base * 0.5))
 Tmp = (UBound(AShapeSpline()) + 1)
 Spline_Arc TmpShapeSpline(), 0.25, Radius, 0, 0, 1, (Steps2 * 0.5)
 For CurVertex = 0 To UBound(TmpShapeSpline())
  TmpShapeSpline(CurVertex).Z = (TmpShapeSpline(CurVertex).Z + (Base * 0.5))
 Next CurVertex
 ReDim Preserve AShapeSpline(UBound(AShapeSpline()) + (UBound(TmpShapeSpline()) + 1))
 For CurVertex = 0 To UBound(TmpShapeSpline())
  AShapeSpline(Tmp + CurVertex) = TmpShapeSpline(CurVertex)
 Next CurVertex

 Spline_Circle AMoveSpline(), ApproachVal, 0, 0, 2, Steps1
 ReDim Preserve AMoveSpline(UBound(AMoveSpline) + 1)
 AMoveSpline(UBound(AMoveSpline)) = AMoveSpline(1)

 Primitive_Capsule = Primitive_MakeFromSplines(AShapeSpline(), AMoveSpline(), 0.5, 0, 0, 1, 1, 1, 1, DoubleSided, True)
 If (Primitive_Capsule = -1) Then Exit Function

 TheMeshs(Primitive_Capsule).Label = "Capsule_" & CStr(TheMeshsCount)
 Select Case Axis
  Case 0:
   For CurVertex = TheMeshs(Primitive_Capsule).Vertices.Start To GetAddressLast(TheMeshs(Primitive_Capsule).Vertices)
    'Swap X to Z:
    Tmp = TheVertices(CurVertex).Position.X
    TheVertices(CurVertex).Position.X = TheVertices(CurVertex).Position.Z
    TheVertices(CurVertex).Position.Z = Tmp
    'Swap Z to Y:
    Tmp = TheVertices(CurVertex).Position.Z
    TheVertices(CurVertex).Position.Z = TheVertices(CurVertex).Position.Y
    TheVertices(CurVertex).Position.Y = Tmp
   Next CurVertex
  Case 1:
   For CurVertex = TheMeshs(Primitive_Capsule).Vertices.Start To GetAddressLast(TheMeshs(Primitive_Capsule).Vertices)
    'Swap Z to Y:
    Tmp = TheVertices(CurVertex).Position.Z
    TheVertices(CurVertex).Position.Z = TheVertices(CurVertex).Position.Y
    TheVertices(CurVertex).Position.Y = Tmp
   Next CurVertex
 End Select

End Function
Function Primitive_Tube(CreateAsDefault As Boolean, Base!, Radius1!, Radius2!, Steps&, Axis As Byte, DoubleSided As Boolean) As Long

 Primitive_Tube = -1

 If (CreateAsDefault = False) Then
  If ((Base <= 0) Or (Radius1 <= 0) Or (Radius1 <= Radius2)) Then Exit Function
  If ((Steps < 2) Or (Axis < 0) Or (Axis > 2)) Then Exit Function
 ElseIf (CreateAsDefault = True) Then
  Base = 100: Radius1 = 50: Radius2 = 25: Steps = 10: Axis = 1: DoubleSided = False
 End If

 Dim AShapeSpline() As Vector3D, AMoveSpline() As Vector3D, CurVertex&, Tmp!

 ReDim AShapeSpline(3)
 AShapeSpline(0) = VectorInput((Radius1 - (Radius2 * 0.5)), 0, -(Base * 0.5))
 AShapeSpline(1) = VectorInput((Radius1 + (Radius2 * 0.5)), 0, -(Base * 0.5))
 AShapeSpline(2) = VectorInput((Radius1 + (Radius2 * 0.5)), 0, (Base * 0.5))
 AShapeSpline(3) = VectorInput((Radius1 - (Radius2 * 0.5)), 0, (Base * 0.5))

 Spline_Circle AMoveSpline(), ApproachVal, 0, 0, 2, Steps
 ReDim Preserve AMoveSpline(UBound(AMoveSpline) + 1)
 AMoveSpline(UBound(AMoveSpline)) = AMoveSpline(1)

 Primitive_Tube = Primitive_MakeFromSplines(AShapeSpline(), AMoveSpline(), 0.5, 0, 0, 1, 1, 1, 1, DoubleSided, True)
 If (Primitive_Tube = -1) Then Exit Function

 TheMeshs(Primitive_Tube).Label = "Tube_" & CStr(TheMeshsCount)
 Select Case Axis
  Case 0:
   For CurVertex = TheMeshs(Primitive_Tube).Vertices.Start To GetAddressLast(TheMeshs(Primitive_Tube).Vertices)
    'Swap X to Z:
    Tmp = TheVertices(CurVertex).Position.X
    TheVertices(CurVertex).Position.X = TheVertices(CurVertex).Position.Z
    TheVertices(CurVertex).Position.Z = Tmp
    'Swap Z to Y:
    Tmp = TheVertices(CurVertex).Position.Z
    TheVertices(CurVertex).Position.Z = TheVertices(CurVertex).Position.Y
    TheVertices(CurVertex).Position.Y = Tmp
   Next CurVertex
  Case 1:
   For CurVertex = TheMeshs(Primitive_Tube).Vertices.Start To GetAddressLast(TheMeshs(Primitive_Tube).Vertices)
    'Swap Z to Y:
    Tmp = TheVertices(CurVertex).Position.Z
    TheVertices(CurVertex).Position.Z = TheVertices(CurVertex).Position.Y
    TheVertices(CurVertex).Position.Y = Tmp
   Next CurVertex
 End Select

End Function
Function Spline_Bezier2D(TheSamples() As Vector3D, SX!, SY!, EX!, EY!, C1X!, C1Y!, C2X!, C2Y!, Angle!, Steps&) As Boolean

 If ((Steps < 4) Or (Angle < 0)) Then Exit Function

 Dim A!, TheTimer!, Density!, Counter&

 ReDim TheSamples(Steps): Density = (1 / (Steps + 1))

 For TheTimer = 0 To 1 Step Density
  A = (1 - TheTimer)
  'Cubic interpolations:
  TheSamples(Counter).X = ((SX * (A ^ 3)) + (EX * (TheTimer ^ 3)) + (((C1X * 3) * (A ^ 2)) * TheTimer)) + (((C2X * 3) * (TheTimer ^ 2)) * A)
  TheSamples(Counter).Y = ((SY * (A ^ 3)) + (EY * (TheTimer ^ 3)) + (((C1Y * 3) * (A ^ 2)) * TheTimer)) + (((C2Y * 3) * (TheTimer ^ 2)) * A)
  If (Angle <> 0) Then TheSamples(Counter) = VectorRotate(TheSamples(Counter), 2, Angle)
  Counter = (Counter + 1)
 Next TheTimer

 Spline_Bezier2D = True

End Function
Function Spline_Bezier3D(TheSamples() As Vector3D, SX!, SY!, SZ!, EX!, EY!, EZ!, C1X!, C1Y!, C1Z!, C2X!, C2Y!, C2Z!, Steps&) As Boolean

 If (Steps < 4) Then Exit Function

 Dim A!, TheTimer!, Density!, Counter&

 ReDim TheSamples(Steps): Density = (1 / (Steps + 1))

 For TheTimer = 0 To 1 Step Density
  A = (1 - TheTimer)
  'Cubic interpolations:
  TheSamples(Counter).X = ((SX * (A ^ 3)) + (EX * (TheTimer ^ 3)) + (((C1X * 3) * (A ^ 2)) * TheTimer)) + (((C2X * 3) * (TheTimer ^ 2)) * A)
  TheSamples(Counter).Y = ((SY * (A ^ 3)) + (EY * (TheTimer ^ 3)) + (((C1Y * 3) * (A ^ 2)) * TheTimer)) + (((C2Y * 3) * (TheTimer ^ 2)) * A)
  TheSamples(Counter).Z = ((SZ * (A ^ 3)) + (EZ * (TheTimer ^ 3)) + (((C1Z * 3) * (A ^ 2)) * TheTimer)) + (((C2Z * 3) * (TheTimer ^ 2)) * A)
  Counter = (Counter + 1)
 Next TheTimer

 Spline_Bezier3D = True

End Function
Function Spline_Circle(TheSamples() As Vector3D, Radius!, Pos1!, Pos2!, Axis As Byte, Steps&) As Boolean

 If ((Steps < 2) Or (Axis < 0) Or (Axis > 2) Or (Radius <= 0)) Then Exit Function

 Dim CurAngle!, StepAngle!, CurSample&

 ReDim TheSamples(Steps): StepAngle = (Pi2 / Steps)

 Select Case Axis
  Case 0:
   For CurSample = 0 To Steps
    TheSamples(CurSample).Y = (Pos1 + (Radius * Sin(CurAngle)))
    TheSamples(CurSample).Z = (Pos2 + (Radius * Cos(CurAngle)))
    CurAngle = (CurAngle + StepAngle)
   Next CurSample
  Case 1:
   For CurSample = 0 To Steps
    TheSamples(CurSample).X = (Pos1 + (Radius * Sin(CurAngle)))
    TheSamples(CurSample).Z = (Pos2 + (Radius * Cos(CurAngle)))
    CurAngle = (CurAngle + StepAngle)
   Next CurSample
  Case 2:
   For CurSample = 0 To Steps
    TheSamples(CurSample).X = (Pos1 + (Radius * Sin(CurAngle)))
    TheSamples(CurSample).Y = (Pos2 + (Radius * Cos(CurAngle)))
    CurAngle = (CurAngle + StepAngle)
   Next CurSample
 End Select

 Spline_Circle = True

End Function
Function Spline_Arc(TheSamples() As Vector3D, ArcEnd!, Radius!, Pos1!, Pos2!, Axis As Byte, Steps&) As Boolean

 If ((Steps < 2) Or (Axis < 0) Or (Axis > 2) Or (Radius <= 0)) Then Exit Function
 If ((ArcEnd <= 0) Or (ArcEnd > 1)) Then Exit Function

 Dim CurAngle!, StepAngle!, CurSample&

 ReDim TheSamples(Steps): StepAngle = ((Pi2 * ArcEnd) / Steps)

 Select Case Axis
  Case 0:
   For CurSample = 0 To Steps
    TheSamples(CurSample).Y = (Pos1 + (Radius * Sin(CurAngle)))
    TheSamples(CurSample).Z = (Pos2 + (Radius * Cos(CurAngle)))
    CurAngle = (CurAngle + StepAngle)
   Next CurSample
  Case 1:
   For CurSample = 0 To Steps
    TheSamples(CurSample).X = (Pos1 + (Radius * Sin(CurAngle)))
    TheSamples(CurSample).Z = (Pos2 + (Radius * Cos(CurAngle)))
    CurAngle = (CurAngle + StepAngle)
   Next CurSample
  Case 2:
   For CurSample = 0 To Steps
    TheSamples(CurSample).X = (Pos1 + (Radius * Sin(CurAngle)))
    TheSamples(CurSample).Y = (Pos2 + (Radius * Cos(CurAngle)))
    CurAngle = (CurAngle + StepAngle)
   Next CurSample
 End Select

 Spline_Arc = True

End Function
Function Spline_Ellipse(TheSamples() As Vector3D, Min1!, Max1!, Min2!, Max2!, Angle!, Axis As Byte, Steps&) As Boolean

 If ((Steps < 2) Or (Axis < 0) Or (Axis > 2) Or (Angle < 0)) Then Exit Function
 If (((Min2 - Min1) < 0) Or ((Max2 - Max1) < 0)) Then Exit Function

 Dim CurAngle!, StepAngle!, CurSample&, Diff1!, Diff2!

 ReDim TheSamples(Steps): StepAngle = (Pi2 / Steps)
 Diff1 = (Min2 - Min1): Diff2 = (Max2 - Max1)

 Select Case Axis
  Case 0:
   For CurSample = 0 To Steps
    TheSamples(CurSample).Y = ((Diff1 * 0.5) * Sin(CurAngle))
    TheSamples(CurSample).Z = ((Diff2 * 0.5) * Cos(CurAngle))
    If (Angle <> 0) Then TheSamples(CurSample) = VectorRotate(TheSamples(CurSample), Axis, Angle)
    TheSamples(CurSample).Y = (TheSamples(CurSample).Y + (Min1 + (Diff1 * 0.5)))
    TheSamples(CurSample).Z = (TheSamples(CurSample).Z + (Max1 + (Diff2 * 0.5)))
    CurAngle = (CurAngle + StepAngle)
   Next CurSample
  Case 1:
   For CurSample = 0 To Steps
    TheSamples(CurSample).X = ((Diff1 * 0.5) * Sin(CurAngle))
    TheSamples(CurSample).Z = ((Diff2 * 0.5) * Cos(CurAngle))
    If (Angle <> 0) Then TheSamples(CurSample) = VectorRotate(TheSamples(CurSample), Axis, Angle)
    TheSamples(CurSample).X = (TheSamples(CurSample).X + (Min1 + (Diff1 * 0.5)))
    TheSamples(CurSample).Z = (TheSamples(CurSample).Z + (Max1 + (Diff2 * 0.5)))
    CurAngle = (CurAngle + StepAngle)
   Next CurSample
  Case 2:
   For CurSample = 0 To Steps
    TheSamples(CurSample).X = ((Diff1 * 0.5) * Sin(CurAngle))
    TheSamples(CurSample).Y = ((Diff2 * 0.5) * Cos(CurAngle))
    If (Angle <> 0) Then TheSamples(CurSample) = VectorRotate(TheSamples(CurSample), Axis, Angle)
    TheSamples(CurSample).X = (TheSamples(CurSample).X + (Min1 + (Diff1 * 0.5)))
    TheSamples(CurSample).Y = (TheSamples(CurSample).Y + (Max1 + (Diff2 * 0.5)))
    CurAngle = (CurAngle + StepAngle)
   Next CurSample
 End Select

 Spline_Ellipse = True

End Function
Function Spline_Star(TheSamples() As Vector3D, Pos1!, Pos2!, MinRadius!, MaxRadius!, Axis As Byte, Steps&) As Boolean

 If ((Steps < 5) Or (Axis < 0) Or (Axis > 2)) Then Exit Function
 If (MinRadius <= 0) Then Exit Function
 If (MinRadius >= MaxRadius) Then Exit Function

 Dim CurAngle!, StepAngle!, CurSample&, What As Boolean

 ReDim TheSamples(Steps): StepAngle = (Pi2 / Steps)

 Select Case Axis
  Case 0:
   For CurSample = 0 To Steps
    If (What = False) Then
     TheSamples(CurSample).Y = (Pos1 + (MaxRadius * Sin(CurAngle)))
     TheSamples(CurSample).Z = (Pos2 + (MaxRadius * Cos(CurAngle)))
    ElseIf (What = True) Then
     TheSamples(CurSample).Y = (Pos1 + (MinRadius * Sin(CurAngle)))
     TheSamples(CurSample).Z = (Pos2 + (MinRadius * Cos(CurAngle)))
    End If
    CurAngle = (CurAngle + StepAngle)
    What = (Not What)
   Next CurSample
  Case 1:
   For CurSample = 0 To Steps
    If (What = False) Then
     TheSamples(CurSample).X = (Pos1 + (MaxRadius * Sin(CurAngle)))
     TheSamples(CurSample).Z = (Pos2 + (MaxRadius * Cos(CurAngle)))
    ElseIf (What = True) Then
     TheSamples(CurSample).X = (Pos1 + (MinRadius * Sin(CurAngle)))
     TheSamples(CurSample).Z = (Pos2 + (MinRadius * Cos(CurAngle)))
    End If
    CurAngle = (CurAngle + StepAngle)
    What = (Not What)
   Next CurSample
  Case 2:
   For CurSample = 0 To Steps
    If (What = False) Then
     TheSamples(CurSample).X = (Pos1 + (MaxRadius * Sin(CurAngle)))
     TheSamples(CurSample).Y = (Pos2 + (MaxRadius * Cos(CurAngle)))
    ElseIf (What = True) Then
     TheSamples(CurSample).X = (Pos1 + (MinRadius * Sin(CurAngle)))
     TheSamples(CurSample).Y = (Pos2 + (MinRadius * Cos(CurAngle)))
    End If
    CurAngle = (CurAngle + StepAngle)
    What = (Not What)
   Next CurSample
 End Select

 Spline_Star = True

End Function
Function Spline_SinStar(TheSamples() As Vector3D, Pos1!, Pos2!, Radius!, Amplitude!, Frequency&, Flower As Boolean, Axis As Byte, Steps&) As Boolean

 If ((Steps < 19) Or (Axis < 0) Or (Axis > 2) Or (Radius <= 0)) Then Exit Function
 If ((Amplitude > Radius) Or (Frequency < 2)) Then Exit Function

 Dim CurAngle!, StepAngle!, CurSin!, CurSample&

 ReDim TheSamples(Steps): StepAngle = (Pi2 / Steps)

 Select Case Axis
  Case 0:
   For CurSample = 0 To Steps
    If (Flower = False) Then
     CurSin = Sin(((CurSample / Steps) * Pi2) * (Frequency + 1))
    ElseIf (Flower = True) Then
     CurSin = Abs(Sin(((CurSample / Steps) * Pi2) * (Frequency + 1)))
    End If
    TheSamples(CurSample).Y = (Pos1 + ((Radius + (Amplitude * CurSin)) * Sin(CurAngle)))
    TheSamples(CurSample).Z = (Pos2 + ((Radius + (Amplitude * CurSin)) * Cos(CurAngle)))
    CurAngle = (CurAngle + StepAngle)
   Next CurSample
  Case 1:
   For CurSample = 0 To Steps
    If (Flower = False) Then
     CurSin = Sin(((CurSample / Steps) * Pi2) * (Frequency + 1))
    ElseIf (Flower = True) Then
     CurSin = Abs(Sin(((CurSample / Steps) * Pi2) * (Frequency + 1)))
    End If
    TheSamples(CurSample).X = (Pos1 + ((Radius + (Amplitude * CurSin)) * Sin(CurAngle)))
    TheSamples(CurSample).Z = (Pos2 + ((Radius + (Amplitude * CurSin)) * Cos(CurAngle)))
    CurAngle = (CurAngle + StepAngle)
   Next CurSample
  Case 2:
   For CurSample = 0 To Steps
    If (Flower = False) Then
     CurSin = Sin(((CurSample / Steps) * Pi2) * (Frequency + 1))
    ElseIf (Flower = True) Then
     CurSin = Abs(Sin(((CurSample / Steps) * Pi2) * (Frequency + 1)))
    End If
    TheSamples(CurSample).X = (Pos1 + ((Radius + (Amplitude * CurSin)) * Sin(CurAngle)))
    TheSamples(CurSample).Y = (Pos2 + ((Radius + (Amplitude * CurSin)) * Cos(CurAngle)))
    CurAngle = (CurAngle + StepAngle)
   Next CurSample
 End Select

 Spline_SinStar = True

End Function
