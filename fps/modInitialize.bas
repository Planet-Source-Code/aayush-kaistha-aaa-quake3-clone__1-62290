Attribute VB_Name = "modInitialize"
Option Explicit

Public Function InitD3D() As Boolean

On Error GoTo Hell:

Dim D3DWindow As D3DPRESENT_PARAMETERS
Dim DispMode As D3DDISPLAYMODE

'initialize and allocate memory 4 directX variables
Set Dx = New DirectX8
Set D3D = Dx.Direct3DCreate
Set D3DX = New D3DX8

DispMode.Format = CheckDisplayMode(640, 480, 32)
If DispMode.Format > D3DFMT_UNKNOWN Then
    Debug.Print "Using 32-Bit format"
Else
    DispMode.Format = CheckDisplayMode(640, 480, 16)
    If DispMode.Format > D3DFMT_UNKNOWN Then
        Debug.Print "32-Bit format not supported. Using 16-Bit format"
    Else
        MsgBox "Neither 16-Bit nor 32-Bit Display Mode Supported", vbInformation, "ERROR"
        Unload frmMain
        End
    End If
End If

With D3DWindow
    .BackBufferCount = 1
    .BackBufferFormat = DispMode.Format
    .BackBufferWidth = 640
    .BackBufferHeight = 480
    .hDeviceWindow = frmMain.hWnd
    .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
End With

If D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DispMode.Format, D3DUSAGE_DEPTHSTENCIL, D3DRTYPE_SURFACE, D3DFMT_D32) = D3D_OK Then
    D3DWindow.AutoDepthStencilFormat = D3DFMT_D32
    D3DWindow.EnableAutoDepthStencil = 1
Else
    If D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DispMode.Format, D3DUSAGE_DEPTHSTENCIL, D3DRTYPE_SURFACE, D3DFMT_D24X8) = D3D_OK Then
        D3DWindow.AutoDepthStencilFormat = D3DFMT_D24X8
        D3DWindow.EnableAutoDepthStencil = 1
    Else
        If D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DispMode.Format, D3DUSAGE_DEPTHSTENCIL, D3DRTYPE_SURFACE, D3DFMT_D16) = D3D_OK Then
            D3DWindow.AutoDepthStencilFormat = D3DFMT_D16
            D3DWindow.EnableAutoDepthStencil = 1
        Else
            D3DWindow.EnableAutoDepthStencil = 0
            MsgBox "Depth buffer could not be enabled", vbInformation, "Depth buffer not supported"
        End If
    End If
End If

Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow)

With D3DDevice
    .SetVertexShader FVF_VERTEX
    .SetRenderState D3DRS_LIGHTING, 1
    .SetRenderState D3DRS_AMBIENT, D3DColorXRGB(150, 150, 150)
    .SetRenderState D3DRS_ZENABLE, 1
    .SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
End With

Set FlTex = D3DX.CreateTextureFromFile(D3DDevice, TexturePath + "flash.bmp")
Set hitTex = D3DX.CreateTextureFromFile(D3DDevice, TexturePath + "hit.jpg")

D3DDevice.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
D3DDevice.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
D3DDevice.SetTextureStageState 0, D3DTSS_COLORARG2, D3DTA_DIFFUSE
D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_DISABLE
    
D3DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
D3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
D3DDevice.SetTextureStageState 0, D3DTSS_MIPFILTER, D3DTEXF_LINEAR

D3DDevice.SetRenderState D3DRS_SRCBLEND, 13
D3DDevice.SetRenderState D3DRS_DESTBLEND, 13

'initialize our matrices
D3DXMatrixIdentity matWorld
D3DDevice.SetTransform D3DTS_WORLD, matWorld

D3DXMatrixLookAtLH matView, MakeVector(-2, 2, -2), MakeVector(0, 0, 0), MakeVector(0, 1, 0)
D3DDevice.SetTransform D3DTS_VIEW, matView

D3DXMatrixPerspectiveFovLH matProj, PI / 3, 1, 0.1, 1500
D3DDevice.SetTransform D3DTS_PROJECTION, matProj

'font settings
SetFont "Verdana", 12, True

InitD3D = True

Exit Function
Hell:
MsgBox "ERROR initializing D3D ", vbCritical, "ERROR"
InitD3D = False

End Function

Public Sub InitDInput()

Set DI = Dx.DirectInputCreate
Set DIDevice = DI.CreateDevice("guid_SysMouse")
Call DIDevice.SetCommonDataFormat(DIFORMAT_MOUSE)
Call DIDevice.SetCooperativeLevel(frmMain.hWnd, DISCL_FOREGROUND Or DISCL_EXCLUSIVE)

Dim diProp As DIPROPLONG
diProp.lHow = DIPH_DEVICE
diProp.lObj = 0
diProp.lData = 10
    
Call DIDevice.SetProperty("DIPROP_BUFFERSIZE", diProp)
hEvent = Dx.CreateEvent(frmMain)
DIDevice.SetEventNotification hEvent

DIDevice.Acquire

End Sub

Public Sub InitDSound()
On Error GoTo Hell:

Dim DSBDesc As DSBUFFERDESC

Set DSEnum = Dx.GetDSEnum
Set DS = Dx.DirectSoundCreate(DSEnum.GetGuid(1))
DS.SetCooperativeLevel frmMain.hWnd, DSSCL_NORMAL

DSBDesc.lFlags = DSBCAPS_CTRLVOLUME
Set sndShoot = DS.CreateSoundBufferFromFile(SoundPath + "shoot.wav", DSBDesc)
If sndShoot Is Nothing Then GoTo Hell:
sndShoot.SetVolume -1000

Set sndReload = DS.CreateSoundBufferFromFile(SoundPath + "reload.wav", DSBDesc)
If sndReload Is Nothing Then GoTo Hell:

Set sndEneShoot = DS.CreateSoundBufferFromFile(SoundPath + "shoot.wav", DSBDesc)
If sndEneShoot Is Nothing Then GoTo Hell:
sndEneShoot.SetVolume -1000

Set sndEmpty = DS.CreateSoundBufferFromFile(SoundPath + "empty.wav", DSBDesc)
If sndEmpty Is Nothing Then GoTo Hell:

Set sndEneHit = DS.CreateSoundBufferFromFile(SoundPath + "ene_hit.wav", DSBDesc)
If sndEneHit Is Nothing Then GoTo Hell:
sndEneHit.SetVolume -500

Set sndEneXplo = DS.CreateSoundBufferFromFile(SoundPath + "explosion.wav", DSBDesc)
If sndEneXplo Is Nothing Then GoTo Hell:

Set sndWarning = DS.CreateSoundBufferFromFile(SoundPath + "warning.wav", DSBDesc)
If sndWarning Is Nothing Then GoTo Hell:
sndWarning.SetVolume -1000

Set sndWin = DS.CreateSoundBufferFromFile(SoundPath + "win.wav", DSBDesc)
If sndWin Is Nothing Then GoTo Hell:

Set sndPDie = DS.CreateSoundBufferFromFile(SoundPath + "pl_die.wav", DSBDesc)
If sndPDie Is Nothing Then GoTo Hell:

Set sndPHeal = DS.CreateSoundBufferFromFile(SoundPath + "heal.wav", DSBDesc)
If sndPHeal Is Nothing Then GoTo Hell:

Exit Sub
Hell:
    MsgBox "Could not initialize Direct Sound", vbInformation, "ERROR"
End Sub

Public Sub LoadModels()

LoadFromX Floor, "floor.x"
LoadFromX Wall, "wall.x"
LoadFromX Wall1, "wall1.x"
LoadFromX WallBig, "wallbig.x"
LoadFromX WallBig1, "wallbig1.x"
LoadFromX Gun, "gun.x"
LoadFromX Ammo, "ammo.x"
LoadFromX Health, "health.x"
LoadFromX Ene_Body, "ene_body.x"
LoadFromX Ene_Head, "ene_head.x"
LoadFromX Ene_Gun, "ene_gun.x"
LoadFromX Part, "part.x"

End Sub

Public Sub LoadFromX(ByRef tmpObj As Object3D, XFile As String)

On Error GoTo Out:

Dim mtrlBuffer As D3DXBuffer
Dim i As Long, j As Integer

Set tmpObj.Mesh = D3DX.LoadMeshFromX(ModelPath + XFile, D3DXMESH_MANAGED, D3DDevice, Nothing, mtrlBuffer, tmpObj.nMaterials)
    
ReDim tmpObj.Materials(tmpObj.nMaterials) As D3DMATERIAL8
ReDim tmpObj.Textures(tmpObj.nMaterials) As Direct3DTexture8

For i = 0 To tmpObj.nMaterials - 1
    D3DX.BufferGetMaterial mtrlBuffer, i, tmpObj.Materials(i)
    tmpObj.Materials(i).Ambient = tmpObj.Materials(i).diffuse
    tmpObj.TextureFile = D3DX.BufferGetTextureName(mtrlBuffer, i)
    If tmpObj.TextureFile <> "" Then
        Set tmpObj.Textures(i) = D3DX.CreateTextureFromFile(D3DDevice, TexturePath + tmpObj.TextureFile)
    End If
Next

Exit Sub
Out:
    MsgBox "Error loading models", vbCritical, "ERROR"
End Sub

Public Sub SetFont(Name As String, Size As Integer, Bold As Boolean)

fnt.Name = Name
fnt.Size = Size
fnt.Bold = Bold
Set MainFontDesc = fnt
Set MainFont = D3DX.CreateFont(D3DDevice, MainFontDesc.hFont)

End Sub

Public Sub Initialize()

frmMain.Show

ModelPath = App.Path + "\models\"
TexturePath = App.Path + "\textures\"
SoundPath = App.Path + "\sound\"

bRunning = InitD3D
InitDInput
InitDSound
LoadModels
CalModelDimensions
LoadMap
SetLights
InitFlash
InitPHit
InitCH
InitStatItems
ResetVariables
GameOver = True
ShowCursor 0

End Sub

Public Sub ResetVariables()

With Player
    .Pos = MakeVector(460, 5, 460)
    .Rotation = PI
    .Health = 3
    .Ammo = 20
    .Dead = False
    .Hit = False
    .Score = 0
End With

With Enemy
    .Pos = MakeVector(-200, -20, -200)
    .Rotation = PI / 4
    .Health = 3
    .Dead = False
    .Hit = False
    .Score = 0
End With

SetWayPoints
FlashPos = 15
AmmoPos = WayPt(2)
HealthPos = WayPt(6)
DispHealth = False
CamPitch = 0
EndState = 0

End Sub

Public Function CheckDisplayMode(Width As Long, Height As Long, Depth As Long) As CONST_D3DFORMAT
Dim i As Long
Dim DispMode As D3DDISPLAYMODE
    
For i = 0 To D3D.GetAdapterModeCount(0) - 1
    D3D.EnumAdapterModes 0, i, DispMode
    If DispMode.Width = Width Then
        If DispMode.Height = Height Then
            If (DispMode.Format = D3DFMT_R5G6B5) Or (DispMode.Format = D3DFMT_X1R5G5B5) Or (DispMode.Format = D3DFMT_X4R4G4B4) Then
                '16 bit mode
                If Depth = 16 Then
                    CheckDisplayMode = DispMode.Format
                    Exit Function
                End If
            ElseIf (DispMode.Format = D3DFMT_R8G8B8) Or (DispMode.Format = D3DFMT_X8R8G8B8) Then
                '32bit mode
                If Depth = 32 Then
                    CheckDisplayMode = DispMode.Format
                    Exit Function
                End If
            End If
        End If
    End If
Next i
CheckDisplayMode = D3DFMT_UNKNOWN
End Function

Public Sub InitFlash()

flash(0) = CreateLitVertex(1.5, 0, 0, &HFFFFFFFF, 0, 0, 0)
flash(1) = CreateLitVertex(0, 0, 0, &HFFFFFF, 0, 0, 0)
flash(2) = CreateLitVertex(0, 3, 0, &HFFFFFFFF, 0, 0, 0)

flash(3) = CreateLitVertex(-1.5, 0, 0, &HFFFFFFFF, 0, 0, 0)
flash(4) = CreateLitVertex(0, 0, 0, &HFFFFFF, 0, 0, 0)
flash(5) = CreateLitVertex(0, 3, 0, &HFFFFFFFF, 0, 0, 0)

flash(6) = CreateLitVertex(1.5, 0, 0, &HFFFFFFFF, 0, 0, 0)
flash(7) = CreateLitVertex(0, 0, 0, &HFFFFFF, 0, 0, 0)
flash(8) = CreateLitVertex(0, -3, 0, &HFFFFFFFF, 0, 0, 0)

flash(9) = CreateLitVertex(-1.5, 0, 0, &HFFFFFFFF, 0, 0, 0)
flash(10) = CreateLitVertex(0, 0, 0, &HFFFFFF, 0, 0, 0)
flash(11) = CreateLitVertex(0, -3, 0, &HFFFFFFFF, 0, 0, 0)

End Sub

Public Sub InitPHit()

PHit(0) = CreateLitVertex(10, 10, 0, &HFFFFFF, 0, 0, 0)
PHit(1) = CreateLitVertex(0, 0, 0, &HAAAAAAAA, 0, 0, 0)
PHit(2) = CreateLitVertex(10, -10, 0, &HFFFFFF, 0, 0, 0)

PHit(3) = CreateLitVertex(-10, 10, 0, &HFFFFFF, 0, 0, 0)
PHit(4) = CreateLitVertex(0, 0, 0, &HAAAAAAAA, 0, 0, 0)
PHit(5) = CreateLitVertex(10, 10, 0, &HFFFFFF, 0, 0, 0)

PHit(6) = CreateLitVertex(-10, 10, 0, &HFFFFFF, 0, 0, 0)
PHit(7) = CreateLitVertex(0, 0, 0, &HAAAAAAAA, 0, 0, 0)
PHit(8) = CreateLitVertex(-10, -10, 0, &HFFFFFF, 0, 0, 0)

PHit(9) = CreateLitVertex(-10, -10, 0, &HFFFFFF, 0, 0, 0)
PHit(10) = CreateLitVertex(0, 0, 0, &HAAAAAAAA, 0, 0, 0)
PHit(11) = CreateLitVertex(10, -10, 0, &HFFFFFF, 0, 0, 0)

End Sub

Public Sub InitCH()

CH(0) = CreateTLVertex(300, 240, 0, 1, &HFF00&, 0, 0, 0)
CH(1) = CreateTLVertex(315, 240, 0, 1, &HFF00&, 0, 0, 0)
CH(2) = CreateTLVertex(325, 240, 0, 1, &HFF00&, 0, 0, 0)
CH(3) = CreateTLVertex(340, 240, 0, 1, &HFF00&, 0, 0, 0)

CH(4) = CreateTLVertex(320, 220, 0, 1, &HFF00&, 0, 0, 0)
CH(5) = CreateTLVertex(320, 235, 0, 1, &HFF00&, 0, 0, 0)
CH(6) = CreateTLVertex(320, 245, 0, 1, &HFF00&, 0, 0, 0)
CH(7) = CreateTLVertex(320, 260, 0, 1, &HFF00&, 0, 0, 0)

End Sub

Public Sub InitStatItems()

Set StatTex(0) = D3DX.CreateTextureFromFile(D3DDevice, TexturePath + "stat.bmp")
Set StatTex(1) = D3DX.CreateTextureFromFile(D3DDevice, TexturePath + "stat1.bmp")
Set StatTex(2) = D3DX.CreateTextureFromFile(D3DDevice, TexturePath + "stat2.bmp")
Set StatTex(3) = D3DX.CreateTextureFromFile(D3DDevice, TexturePath + "stat3.bmp")

CreateRectangle 0, 10, 150, 70, &HAAAAAA, Stats.Fps
CreateRectangle 160, 10, 310, 70, &HAAAAAA, Stats.Ammo
CreateRectangle 320, 10, 470, 70, &HAAAAAA, Stats.Health
CreateRectangle 480, 10, 630, 70, &HAAAAAA, Stats.Score

End Sub

Public Sub CreateRectangle(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, Color As Long, ByRef tmp() As TLVERTEX)

tmp(0) = CreateTLVertex(X1, Y1, 0, 1, Color, 0, 0, 0)
tmp(1) = CreateTLVertex(X2, Y1, 0, 1, Color, 0, 1, 0)
tmp(2) = CreateTLVertex(X1, Y2, 0, 1, Color, 0, 0, 1)
tmp(3) = CreateTLVertex(X2, Y2, 0, 1, Color, 0, 1, 1)

End Sub

Public Sub CalModelDimensions()

CalMeshDimen Wall
CalMeshDimen Wall1
CalMeshDimen WallBig
CalMeshDimen WallBig1
CalMeshDimen Ene_Body
CalMeshDimen Ene_Head

End Sub

Public Sub CalMeshDimen(ByRef tmpObj As Object3D)
Dim min As D3DVECTOR, max As D3DVECTOR

D3DX.ComputeBoundingBoxFromMesh tmpObj.Mesh, min, max
With tmpObj.Radius
    .X = (max.X - min.X) / 2
    .Y = (max.Y - min.Y) / 2
    .Z = (max.Z - min.Z) / 2
End With

End Sub

Public Sub SetWayPoints()

WayPt(0) = MakeVector(-100, 0, -100)
WayPt(1) = MakeVector(-350, 0, -100)
WayPt(2) = MakeVector(-200, 0, -250)
WayPt(3) = MakeVector(-450, 0, -450)
WayPt(4) = MakeVector(-450, 0, 100)
WayPt(5) = MakeVector(-150, 0, 250)
WayPt(6) = MakeVector(-300, 0, 100)
WayPt(7) = MakeVector(-100, 0, 450)
WayPt(8) = MakeVector(100, 0, 450)
WayPt(9) = MakeVector(400, 0, 300)
WayPt(10) = MakeVector(200, 0, 100)
WayPt(11) = MakeVector(450, 0, 200)
WayPt(12) = MakeVector(450, 0, -200)
WayPt(13) = MakeVector(100, 0, -100)
WayPt(14) = MakeVector(300, 0, -250)
WayPt(15) = MakeVector(200, 0, -450)
WayPt(16) = MakeVector(-200, 0, -450)

End Sub
