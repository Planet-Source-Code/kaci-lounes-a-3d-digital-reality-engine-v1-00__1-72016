Attribute VB_Name = "BAS_Declarations"

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
'###  MODULE      : BAS_Declarations.BAS
'###
'###  DESCRIPTION : Constants, enumrations & data-structures definitions.
'###
'##################################################################################
'##################################################################################

Option Explicit

' =====================================================================
' ========================== GLOBAL CONSTANTS =========================
' =====================================================================

Global Const Pi As Single = 3.141593
Global Const Pi2 As Single = 6.283185
Global Const Deg As Single = 0.0174532
Global Const OneByThree As Single = 0.3333333
Global Const AlphaFactor As Single = 0.0039215  '(1/255)
Global Const ApproachVal As Single = 0.0000001
Global Const MaxSingleFloat As Single = 3.402823E+38
Global Const SceneFileExtension As String = ".3DR"
Global Const ObjectFileExtension As String = ".3DO"
Global Const AppFilePassword As String = "3D Digital Reality Engine, by KACI Lounes"
Global Const EstimateFilterSize! = 0.8       'Estimate filter size for the photons
'Global Const NaturalLogBase As Single = 2.718282
Global Const MaximumRefractionNFactor As Single = 10 'We need a maximum value because
                                                     'a mapping gives only a Byte (0-255)
                                                     'and this is a Single value that we
                                                     'are talking about.

'=====================================================================
'============================ ENUMRATIONS ============================
'=====================================================================

Public Enum K3DE_BITMAP_FILTER_MODES
 K3DE_BFM_BRIGHTNESS = 61             'Additive white color
 K3DE_BFM_CONTRAST = 62               'Contrast effect
 K3DE_BFM_GAMMA = 63                  'Gamma correction
 K3DE_BFM_GREYSCALE = 64              'GreyScaled channel
 K3DE_BFM_INVERT = 65                 'Colors-inverstion
 K3DE_BFM_MONOCHROME = 66             'Monochromatic (white & black)
 K3DE_BFM_SWAPCHANNELS = 67           'Swaps between RGB channels
End Enum

'Texels filtering modes
'======================
Public Enum K3DE_TEXELS_FILTER_MODES
 K3DE_XFM_NOFILTER = 41               'No Filter
 K3DE_XFM_BILINEAR = 42               'Bi-linear interpolation
 K3DE_XFM_BELL = 43
 K3DE_XFM_GAUSSIAN = 44               'Parametric (KernelSize)
 K3DE_XFM_CUBIC_SPLINE_B = 45
 K3DE_XFM_CUBIC_SPLINE_BC = 46        'Parametric (CubicB, CubicC)
 K3DE_XFM_CUBIC_SPLINE_CARDINAL = 47  'Parametric (CubicA)
End Enum

'Texture filtering modes
'=======================
Public Enum K3DE_TEXTURE_FILTER_MODES
 K3DE_TFM_NEAREST = 31
 K3DE_TFM_NEAREST_MIP_NEAREST = 32
 K3DE_TFM_NEAREST_MIP_LINEAR = 33
 K3DE_TFM_FILTERED = 34
 K3DE_TFM_FILTERED_MIP_NEAREST = 35
 K3DE_TFM_FILTERED_MIP_LINEAR = 36     '(Trilinear interpolation)
End Enum

'Fog modes
'=========
Public Enum K3DE_FOG_MODES
 K3DE_FM_LINEAR = 61
 K3DE_FM_EXP = 62
End Enum

'=====================================================================
'=========================== DATA TYPES ==============================
'=====================================================================

Public Type Matrix4x4
 M11 As Single: M12 As Single: M13 As Single: M14 As Single
 M21 As Single: M22 As Single: M23 As Single: M24 As Single
 M31 As Single: M32 As Single: M33 As Single: M34 As Single
 M41 As Single: M42 As Single: M43 As Single: M44 As Single
End Type

Public Type Vector3D
 X As Single
 Y As Single
 Z As Single
End Type

Public Type Ray3D
 Position As Vector3D
 Direction As Vector3D
End Type

Public Type TextureVectors
 U1 As Single: V1 As Single
 U2 As Single: V2 As Single
 U3 As Single: V3 As Single
End Type

Public Type ColorRGB
 R As Integer
 G As Integer
 B As Integer
End Type

Public Type ColorHSL    'A model of colors, (Hue, Saturation & Lightness)
 H As Single            'it's also used by the standard 'National Television
 S As Single            'System Committee' (NTSC) to code colors on an analog
 L As Single            'video signal that can be displayed on a TV system (RCA cables),
End Type                '(like PlayStation(R) I or PSTwo(R) ), The H & S datas are coded
                        'into amplitude and phase of the signal respectively, for the
                        'luminance(L), the user define this value generaly by a button
                        'on the TV or just the remote-control. like that, the advantage
                        'of this color-coding is the use of just 2 bands (H&S) instead of
                        'the standard 3 bands (RGB) color model.

Public Type Point2D
 X As Integer
 Y As Integer
End Type

Public Type Rect2D
 X1 As Integer
 Y1 As Integer
 X2 As Integer
 Y2 As Integer
End Type

'//////////////////////////////////////////

'(36 Bytes, instead of 20 Bytes defined by Dr.Jensen, just for a little speed)
Public Type Photon
 Power As Vector3D
 Position As Vector3D
 Direction As Vector3D
End Type

Public Type Address
 Start As Long
 Length As Long
End Type

Public Type Intersection3D
 FaceNumber As Long
 MeshNumber As Long
 IsBackFace As Boolean
 Zt As Single 'A parametric form of the intersection point on the ray.
 U As Single  'The barycentrics coordinates given by the the intersection
 V As Single  'on the projectif space of the triangle, we can map (interpolate)
 W As Single  'any values using these simple parametric coords (u+v+w=1)
End Type      'like the textures coordinates...

Public Type TraceResult 'Resulting intersections (array) of one trace
 IntersectCount As Long
 Intersections() As Intersection3D
End Type

Public Type MaterialInfos 'The material description at an intersection point
 Color As ColorRGB
 Reflection As Byte
 Refraction As Byte
 RefractionN As Single
 SpecularPowerK As Single
 SpecularPowerN As Single
End Type

'///////////////////////////////////////
'//// Primitives 2D data-structures ////
'///////////////////////////////////////

Public Type Circle2D
 Center As Point2D
 Radius As Single
End Type

Public Type Ellipse2D
 MinPoint As Point2D
 MaxPoint As Point2D
 Angle As Single
End Type

Public Type Rectangle2D
 MinPoint As Point2D
 MaxPoint As Point2D
 Angle As Single
End Type

Public Type Bezier2D
 SPoint As Point2D
 EPoint As Point2D
 CPoint1 As Point2D
 CPoint2 As Point2D
End Type

Public Type PolyLine2D
 Points() As Point2D
End Type

Public Type PolyBezier2D
 Points() As Point2D
 CPoints() As Point2D
End Type

'///////////////////////////////////////
'//// 2D SURFACES DATA STRUCTURES //////
'///////////////////////////////////////

Public Type BitMap2D
 Label As String
 BitsDepth As Byte           '8 or 24 bits
 Dimensions As Point2D
 BackGroundColor As ColorRGB
 Datas() As Byte
End Type

Public Type MipTextures
 MipSequance() As BitMap2D
End Type

Public Type BltFlags
 Stretch As Boolean
 PixelFilter As K3DE_TEXELS_FILTER_MODES
 Transparent As Boolean
 TransColor As ColorRGB
 AlphaFlag As Byte
 AlphaValueRed As Byte
 AlphaValueGreen As Byte
 AlphaValueBlue As Byte
 AlphaMask As BitMap2D
 AlphaRect As Rect2D
 ChannelAlphaFrom As Byte
 ChannelFrom As Byte
 ChannelTo As Byte
End Type

'/////////////////////////////////////////
'////////// 3D DATA-STRUCTURES ///////////
'/////////////////////////////////////////

'(24 Bytes)
Public Type Vertex3D
 Position As Vector3D
 TmpPos As Vector3D    'A temporary storage for transformations's calculation
End Type

'(146 Bytes)
Public Type Face3D                     'Triangular face
 Normal As Vector3D                    'Face's perpendicular vector
 Visible As Boolean                    'The face MUST be visible or not (not in alpha-mapping).
 A As Long: B As Long: C As Long       'Indexs of three vertices
 AlphaVectors As TextureVectors        'Parametric texture coordinates (0...1)
 ColorVectors As TextureVectors        '
 ReflectionVectors As TextureVectors   '
 RefractionVectors As TextureVectors   '
 RefractionNVectors As TextureVectors  '
End Type

'(120 Bytes)
Public Type Mesh3D
 Label As String
 Position As Vector3D
 Scales As Vector3D
 Angles As Vector3D
 Vertices As Address
 Faces As Address
 MakeMatrix As Boolean    'Define if the program should automaticly create the matrix
                          'using the meshs's properties, the user define the matrix instead.
 WorldMatrix As Matrix4x4
 Visible As Boolean
End Type

'(50 Bytes)
Public Type Material
 Label As String
 'No-mapping options:
 Color As ColorRGB: Reflection As Byte
 Refraction As Byte: RefractionN As Single
 SpecularPowerK As Single: SpecularPowerN As Single
 'Texture-mapped material:
 UseAlphaTexture As Boolean: AlphaTextureID As Long
 UseColorTexture As Boolean: ColorTextureID As Long
 UseReflectionTexture As Boolean: ReflectionTextureID As Long
 UseRefractionTexture As Boolean: RefractionTextureID As Long
 UseRefractionNTexture As Boolean: RefractionNTextureID As Long
End Type

'The following light data-types are only different
'in the way of attenuation, spheric, a cone, and the
'way of distributing photons in photon mapping.

'(36 Bytes)
Public Type SphereLight3D 'Attenuation begin's from the Position to the range (radius)
 Label As String
 Color As ColorRGB
 Position As Vector3D: TmpPos As Vector3D
 Range As Single
 Enable As Boolean
End Type

'(68 Bytes)
Public Type ConeLight3D 'Attenuation begin's from the Position to the range,
 Label As String        'and from the HotSpot angle to the FallOff angle.
 Color As ColorRGB
 Position As Vector3D:  TmpPos As Vector3D
 Direction As Vector3D: TmpDir As Vector3D
 Falloff As Single      'The cone angle
 Hotspot As Single      'The penumbra angle
 Range As Single
 Enable As Boolean
End Type

'A 3D virtual camera, we use cameras to view the scenry everywhere.
'(108 Bytes)
Public Type Camera3D
 Label As String
 Position As Vector3D
 Direction As Vector3D
 RollAngle As Single        'The Z rotation over the screen
 FOVAngle As Single         'Field Of View angle
 ClearDistance As Single    'The focal distance, or the Depth Of Field distance (future versions)
 Dispersion As Single       'Control the blur level added by the Depth Of Field simulation (future versions)
 BackFaceCulling As Boolean
 MakeMatrix As Boolean      'Define if the program should automaticly create the matrix
                            'using the camera's properties, the user define the matrix instead.
 ViewMatrix As Matrix4x4    'The camera-space matrix, for orienting objects in the way of the camera.
End Type

Public Type Spline3D        '(used to animate the camera, future versions)
 Points() As Vector3D
End Type
