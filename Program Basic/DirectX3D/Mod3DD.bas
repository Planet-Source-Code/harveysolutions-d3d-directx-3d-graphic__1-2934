Attribute VB_Name = "Module3DD"
Option Explicit

Public Const VIEWPORT_WIDTH = 640
Public Const VIEWPORT_HEIGHT = 480
Public Const PI = 3.14159265358979

Public D3DRM As IDirect3DRM

'***********************************
'*  Function to create a Sphere **
'***********************************

Public Sub BuildSphere(objMeshBuilder As IDirect3DRMMeshBuilder)
    Dim aVertices(1 To 1000) As D3DVECTOR
    Dim aNormals(0) As D3DVECTOR
    Dim aFaces(1 To 10000) As Long
    Dim intVertices As Long
    Const STEPA = 10
    Const STEPB = 10
    Dim axeZ As D3DVECTOR, origine As D3DVECTOR, AxeY As D3DVECTOR
    origine.x = 0:   origine.y = 1:   origine.z = 0
    axeZ.x = 0:      axeZ.y = 0:      axeZ.z = 1
    AxeY.x = 0:      AxeY.y = 1:      AxeY.z = 0
    intVertices = 1
    Dim i As Integer, j As Integer
    Dim tmp As D3DVECTOR
    For i = STEPA To 180 - STEPA Step STEPA
    For j = 0 To 360 - STEPB Step STEPB
            D3DRMVectorRotate tmp, origine, axeZ, i * PI / 180
            D3DRMVectorRotate aVertices(intVertices), tmp, AxeY, j * PI / 180
            intVertices = intVertices + 1
       Next
    Next
    intVertices = intVertices - 1
    Dim Index As Integer
    Index = 1
    For i = STEPA To 180 - 2 * STEPA Step STEPA
        Dim FirstIndex As Long
        FirstIndex = Index
        For j = 0 To 360 - STEPB Step STEPB
            aFaces(Index) = 4
            aFaces(Index + 1) = (Index \ 5) + 1
            aFaces(Index + 2) = (Index \ 5)
            aFaces(Index + 3) = ((Index \ 5) + (360 \ STEPB))
            aFaces(Index + 4) = (Index \ 5) + 1 + (360 \ STEPB)
            If j = 360 - STEPB Then
                aFaces(Index + 1) = FirstIndex \ 5  '+ 1
                aFaces(Index + 4) = FirstIndex \ 5 + (360 \ STEPB)
            End If
            Index = Index + 5
        Next
    Next
    aFaces(Index) = (360 / STEPB) - 1
    Index = Index + 1
    For i = 1 To (360 / STEPB) - 1
        aFaces(Index) = i
        Index = Index + 1
    Next
    aFaces(Index) = 360 / STEPB
    Index = Index + 1
    For i = 0 To (360 / STEPB) - 1
        aFaces(Index) = intVertices - i - 1
        Index = Index + 1
    Next
    aFaces(Index) = 0
    objMeshBuilder.AddFaces intVertices, aVertices(1), 0, aNormals(0), aFaces(1), Nothing
  End Sub
Public Sub PutSphereTexture(D3DRM As IDirect3DRM, MeshBuilder As IDirect3DRMMeshBuilder, ByVal strTextureFileName As String)
    Dim Box As D3DRMBOX
    Dim MaxY As Single, MinY As Single
    Dim Height As Single
    Dim Wrap As IDirect3DRMWrap
    Dim Texture As IDirect3DRMTexture
    ' Bounding box
    MeshBuilder.GetBox Box
    MaxY = Box.Max.y
    MinY = Box.Min.y
    Height = MaxY - MinY
    D3DRM.CreateWrap D3DRMWRAP_CYLINDER, Nothing, 0, 0, 0, 0, 1, 0, 0, 0, 1, 0, MinY / Height, 1, -1 / Height, Wrap
    Wrap.Apply MeshBuilder
    D3DRM.LoadTexture strTextureFileName, Texture
    MeshBuilder.SetTexture Texture
End Sub
Sub BuildCube(MeshBuilder As IDirect3DRMMeshBuilder, CHeight, CWidth, CDepth)
    Dim aVertices(0 To 8) As D3DVECTOR
    Dim aNormals(0) As D3DVECTOR
    Dim aFaces(1 To 31) As Long
    Dim FaceArray As IDirect3DRMFaceArray
    Dim Face As IDirect3DRMFace
    Dim Texture As IDirect3DRMTexture
    ' Floor vertices
    aVertices(0).x = -(CWidth / 2)
    aVertices(0).y = 0
    aVertices(0).z = -(CDepth / 2)
    
    aVertices(1).x = -(CWidth / 2)
    aVertices(1).y = 0
    aVertices(1).z = (CDepth / 2)
    
    aVertices(2).x = (CWidth / 2)
    aVertices(2).y = 0
    aVertices(2).z = (CDepth / 2)
    
    aVertices(3).x = (CWidth / 2)
    aVertices(3).y = 0
    aVertices(3).z = -(CDepth / 2)
    ' Ceiling vertices
    Dim i As Long
    For i = 0 To 3
        aVertices(4 + i) = aVertices(i)
        aVertices(4 + i).y = CHeight
    Next
    ' Floor
    aFaces(1) = 4: aFaces(2) = 3: aFaces(3) = 2: aFaces(4) = 1: aFaces(5) = 0
    ' Ceiling
    aFaces(6) = 4: aFaces(7) = 4: aFaces(8) = 5: aFaces(9) = 6:  aFaces(10) = 7
    ' Front wall
    aFaces(11) = 4: aFaces(12) = 2: aFaces(13) = 6: aFaces(14) = 5:  aFaces(15) = 1
    ' Left wall
    aFaces(16) = 4: aFaces(17) = 1: aFaces(18) = 5: aFaces(19) = 4:  aFaces(20) = 0
    ' Right wall
    aFaces(21) = 4: aFaces(22) = 3: aFaces(23) = 7: aFaces(24) = 6:  aFaces(25) = 2
    ' Back wall
    aFaces(26) = 4: aFaces(27) = 0: aFaces(28) = 4: aFaces(29) = 7:  aFaces(30) = 3
    ' Terminator
    aFaces(31) = 0
    
    'D3DRM.CreateMeshBuilder MeshBuilder
    MeshBuilder.AddFaces 8, aVertices(0), 0, aNormals(0), aFaces(1), Nothing
    MeshBuilder.SetPerspective 1
    MeshBuilder.GetFaces FaceArray
    D3DRM.LoadTexture App.Path & "\Bricks.bmp", Texture
    'Slection of the Faces
    FaceArray.GetElement 0, Face:      Face.SetColorRGB 1, 0, 1
    FaceArray.GetElement 1, Face:  Face.SetColorRGB 0, 0, 1
    FaceArray.GetElement 2, Face:  Face.SetColorRGB 1, 1, 1
    FaceArray.GetElement 3, Face:  Face.SetColorRGB 1, 0, 1
    FaceArray.GetElement 4, Face:  Face.SetColorRGB 1, 0, 0
    FaceArray.GetElement 5, Face ':  Face.SetColorRGB 0, 1, 0
    'Put Texture on the Front Face
    Face.SetTextureCoordinates 0, 1, 0
    Face.SetTextureCoordinates 1, 1, 1
    Face.SetTextureCoordinates 2, 0, 1
    Face.SetTextureCoordinates 3, 0, 0
    Face.SetTexture Texture
    
    
End Sub
