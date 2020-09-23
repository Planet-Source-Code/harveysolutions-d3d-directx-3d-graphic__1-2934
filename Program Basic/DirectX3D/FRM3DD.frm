VERSION 5.00
Begin VB.Form frm3DD 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Directx 6.0 "
   ClientHeight    =   7200
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   6000
   End
End
Attribute VB_Name = "frm3DD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************
'**                                                            *
'**   CREDIT CREDIT TO:                                        *
'**   90% to Carl Harvey Alias Carlos                          *
'**   10% Based on the DirectX SDK retained mode tutorial...   *
'**                                                            *
'***************************************************************
Option Explicit
Dim ox%, oy%, dwn%
Dim Scene As IDirect3DRMFrame
Dim Camera As IDirect3DRMFrame
Dim Clipper As IDirectDrawClipper
Dim Device As IDirect3DRMDevice
Dim Viewport As IDirect3DRMViewPort
Dim WorldFrame As IDirect3DRMFrame
Dim CubeFrame As IDirect3DRMFrame
Dim Material As IDirect3DRMMaterial
Dim visuala As IDirect3DRMVisualArray
Dim visual As IDirect3DRMVisual
Dim SphereTextureFile As String
Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim temp1
    Select Case KeyAscii
        Case 43:  Camera.AddTranslation D3DRMCOMBINE_BEFORE, 0, 0, 10 '* 10
        Case 45:  Camera.AddTranslation D3DRMCOMBINE_BEFORE, 0, 0, -10  '* 10
        Case 50:  Camera.AddTranslation D3DRMCOMBINE_BEFORE, 0, 10, 0  '* 10
        Case 52:  Camera.AddTranslation D3DRMCOMBINE_BEFORE, 10, 0, 0
        Case 54:  Camera.AddTranslation D3DRMCOMBINE_BEFORE, -10, 0, 0
        Case 56:  Camera.AddTranslation D3DRMCOMBINE_BEFORE, 0, -10, 0
    End Select
    On Error Resume Next
    'Camera.SetPosition Scene, DistanceX, DistanceY, DistanceZ
    update_screen
End Sub
Private Sub Init_ViewPort()
Dim FormWidth As Long, FormHeight As Long
Dim Light As IDirect3DRMLight
Dim LightFrame As IDirect3DRMFrame
    Direct3DRMCreate D3DRM
    D3DRM.CreateFrame Nothing, Scene  ' Create the scene
    D3DRM.CreateFrame Scene, Camera   ' Create the camera
    Camera.SetPosition Scene, 0, 0, 0
    Camera.SetOrientation Scene, 0, 0, 1, 0, 1, 0
    DirectDrawCreateClipper 0, Clipper, Nothing  'create the engine for display in the windows mode
    Clipper.SetHWnd 0, hWnd  'Set output to the form handle
    D3DRM.CreateDeviceFromClipper Clipper, ByVal 0, VIEWPORT_WIDTH, VIEWPORT_HEIGHT, Device
    D3DRM.CreateViewport Device, Camera, 0, 0, VIEWPORT_WIDTH, VIEWPORT_HEIGHT, Viewport
    Viewport.SetBack 100000     'set the depth of the view
    Device.SetQuality D3DRMLIGHT_ON Or D3DRMFILL_SOLID Or D3DRMSHADE_GOURAUD
    ' Create the light frame
    D3DRM.CreateFrame Scene, LightFrame
    D3DRM.CreateLightRGB D3DRMLIGHT_AMBIENT, 0.9, 0.9, 0.9, Light
    Scene.AddLight Light
    Set Light = Nothing
End Sub

Private Sub Form_Load()
    Dim Globe1 As IDirect3DRMMeshBuilder
    Dim Cube1 As IDirect3DRMMeshBuilder
    SphereTextureFile = App.Path & "\globe.bmp"
    Init_ViewPort
    '**********************************************
    'Create the Globe and position and set animation
    '----------------------------------------------
    D3DRM.CreateFrame Scene, WorldFrame ' Create the globe frame
    WorldFrame.SetPosition Scene, -50, 0, 300
    WorldFrame.SetOrientation Scene, 0, 0, 1, 0, 1, 0
    WorldFrame.SetRotation Scene, 1, 1, 1, 0.17 '0.05
    CreateSphere Globe1, 40, 40, 40, SphereTextureFile
    WorldFrame.AddVisual Globe1
    
    '**********************************************
    'Create the Cube and position and set animation
    '----------------------------------------------
    D3DRM.CreateFrame Scene, CubeFrame
    CubeFrame.SetPosition Scene, 60, 0, 300
    CubeFrame.SetRotation Scene, 0, 20, 0, 0.15
    Create_Cube Cube1, 50, 50, 50
    CubeFrame.AddVisual Cube1
   
End Sub

Private Sub CreateSphere(Globe As IDirect3DRMMeshBuilder, ByVal SCX, ByVal SCY, ByVal SCZ, ByVal GlobeTextureFile)
D3DRM.CreateMeshBuilder Globe
BuildSphere Globe
  'Control optional parameters parameters
Globe.[Scale] SCX, SCY, SCZ
PutSphereTexture D3DRM, Globe, GlobeTextureFile
End Sub

Private Sub Create_Cube(Cube As IDirect3DRMMeshBuilder, CHeight, CWidth, CDepth)
D3DRM.CreateMeshBuilder Cube
BuildCube Cube, CHeight, CWidth, CDepth
End Sub
Private Sub Form_Paint()
update_screen
End Sub
Private Sub update_screen()
Viewport.Clear
Viewport.Render Scene
Device.Update
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set WorldFrame = Nothing
    Set Scene = Nothing
    Set Camera = Nothing
    Set Viewport = Nothing
    Set Device = Nothing
    Set D3DRM = Nothing
    Set Clipper = Nothing
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
dwn% = True
ox% = x: oy% = y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If dwn% Then
    If Button = 1 Then
        'move the camera as the user drags the mouse
        Camera.AddTranslation D3DRMCOMBINE_BEFORE, (x - ox%), 0, (y - oy%)     '* 10
        update_screen
        
    ElseIf Button = 2 Then
        'rotate the viewpoint
        Camera.AddRotation D3DRMCOMBINE_AFTER, 0, -1, 0, ((x - ox%) / 100)
        update_screen
    End If
    ox% = x: oy% = y
   
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
dwn% = False
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
    Scene.Move 1
    update_screen
End Sub
