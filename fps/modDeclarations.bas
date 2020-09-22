Attribute VB_Name = "modDeclarations"
Option Explicit

Public Dx As DirectX8
Public D3D As Direct3D8
Public D3DX As D3DX8
Public D3DDevice As Direct3DDevice8
Public DI As DirectInput8, DIDevice As DirectInputDevice8
Public hEvent As Long 'a handle for an event...
Public DS As DirectSound8
Public DSEnum As DirectSoundEnum8
Public sndShoot As DirectSoundSecondaryBuffer8
Public sndReload As DirectSoundSecondaryBuffer8
Public sndEneShoot As DirectSoundSecondaryBuffer8
Public sndEmpty As DirectSoundSecondaryBuffer8
Public sndPDie As DirectSoundSecondaryBuffer8
Public sndPHeal As DirectSoundSecondaryBuffer8
Public sndEneHit As DirectSoundSecondaryBuffer8
Public sndEneXplo As DirectSoundSecondaryBuffer8
Public sndWarning As DirectSoundSecondaryBuffer8
Public sndWin As DirectSoundSecondaryBuffer8

Public ModelPath As String, TexturePath As String
Public SoundPath As String
Public BulHit As Long, HitDist As Single, HitPos As D3DVECTOR
Public bRunning As Boolean, GameOver As Boolean
Public UpKey As Boolean, DownKey As Boolean
Public Restart As Boolean, EndState As Integer
Public LeftKey As Boolean, RightKey As Boolean
Public WKey As Boolean, SKey As Boolean
Public fire As Boolean, FireTimer As Long
Public EnemySaw As Long, CanHit As Boolean
Public EneDir As D3DVECTOR
Public WarningMsg As String

Public Const FVF_VERTEX = (D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_TEX1)
Public Const FVF_LVERTEX = (D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR Or D3DFVF_TEX1)
Public Const FVF_TLVERTEX = (D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR)

Public Type LVERTEX
    X As Single
    Y As Single
    Z As Single
    Color As Long
    Specular As Long
    tu As Single
    tv As Single
End Type
Public flash(11) As LVERTEX
Public FlTex As Direct3DTexture8

Public PHit(11) As LVERTEX
Public hitTex As Direct3DTexture8
Public PHitTimer As Long

Public Shot(1) As LVERTEX
Public ShotStart As D3DVECTOR, ShotEnd As D3DVECTOR

Public Type VERTEX
    X As Single
    Y As Single
    Z As Single
    nx As Single
    ny As Single
    nz As Single
    tu As Single
    tv As Single
End Type

Public Type TLVERTEX
    X As Single
    Y As Single
    Z As Single
    rhw As Single
    Color As Long
    Specular As Long
    tu As Single
    tv As Single
End Type
Public CH(7) As TLVERTEX

Public Light As D3DLIGHT8

'all these variables r required to load 3d objects in directX
Public Type Object3D
    nMaterials As Long
    Materials() As D3DMATERIAL8
    Textures() As Direct3DTexture8
    TextureFile As String
    Mesh As D3DXMesh
    Radius As D3DVECTOR
End Type

Public Floor As Object3D
Public matFloor(24) As D3DMATRIX

Public Wall As Object3D, Wall1 As Object3D
Public matWall(19) As D3DMATRIX

Public WallBig As Object3D, WallBig1 As Object3D
Public matWallBig(7) As D3DMATRIX

Public Hut As Object3D
Public matHut As D3DMATRIX

Public Gun As Object3D
Public matGun As D3DMATRIX

Public Ammo As Object3D
Public matAmmo As D3DMATRIX
Public AmmoPos As D3DVECTOR

Public Health As Object3D
Public matHealth As D3DMATRIX
Public HealthPos As D3DVECTOR, DispHealth As Boolean

Public Ene_Body As Object3D
Public matEneBody As D3DMATRIX
Public Ene_Head As Object3D
Public matEneHead As D3DMATRIX
Public Ene_Gun As Object3D
Public matEneGun As D3DMATRIX

Public Part As Object3D
Public matPart(5) As D3DMATRIX
Public PartPos(5) As D3DVECTOR

Public FlashPos As Integer

Public Type Plyr_Data
    Pos As D3DVECTOR
    Rotation As Single
    MoveSpeed As Single
    Health As Integer
    Hit As Boolean
    Dead As Boolean
    DieTime As Long
    Score As Integer
    Ammo As Integer
End Type
Public Enemy As Plyr_Data

Public Type StatItems
    Fps(3) As TLVERTEX
    Health(3) As TLVERTEX
    Score(3) As TLVERTEX
    Ammo(3) As TLVERTEX
End Type
Public Stats As StatItems, StatTex(3) As Direct3DTexture8

Public Const nWayPts = 16
Public WayPt(nWayPts) As D3DVECTOR, Curpos As Integer

Public EyeLookDir As D3DVECTOR, EyeLookAt As D3DVECTOR
Public matCam As D3DMATRIX

'this only holds data req to calculate frames per second
Public Type FPS_data
    Count As Long
    Value As Long
    Last As Long
End Type
Public Fps As FPS_data

Public MainFont As D3DXFont
Public MainFontDesc As IFont
Public fnt As New StdFont

Public Const nColl_Obj = 54
Public Type Mesh_Dimen
    Center As D3DVECTOR
    Radius As D3DVECTOR
End Type
Public Coll_Obj(nColl_Obj) As Mesh_Dimen

Public Zoom As Boolean, CamPitch As Single
Public Player As Plyr_Data

Public matFlash As D3DMATRIX

Public matProj As D3DMATRIX 'this holds the camera settings
Public matView As D3DMATRIX 'this tells where the camera is n where it is looking at
Public matWorld As D3DMATRIX 'this holds the reference coordinates of entire 3d world

Public Const PI = 3.14159
Public Const Rad = PI / 180
Public Const DEG = 180 / PI

Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function CreateLitVertex(X As Single, Y As Single, Z As Single, Colour As Long, Specular As Long, tu As Single, tv As Single) As LVERTEX
    CreateLitVertex.X = X
    CreateLitVertex.Y = Y
    CreateLitVertex.Z = Z
    CreateLitVertex.Color = Colour
    CreateLitVertex.Specular = Specular
    CreateLitVertex.tu = tu
    CreateLitVertex.tv = tv
End Function

Public Function CreateVertex(X As Single, Y As Single, Z As Single, nx As Single, ny As Single, nz As Single, tu As Single, tv As Single) As VERTEX
    CreateVertex.X = X
    CreateVertex.Y = Y
    CreateVertex.Z = Z
    CreateVertex.nx = nx
    CreateVertex.ny = ny
    CreateVertex.nz = nz
    CreateVertex.tu = tu
    CreateVertex.tv = tv
End Function

Public Function CreateTLVertex(X As Single, Y As Single, Z As Single, rhw As Single, _
                                                Color As Long, Specular As Long, tu As Single, _
                                                tv As Single) As TLVERTEX
    CreateTLVertex.X = X
    CreateTLVertex.Y = Y
    CreateTLVertex.Z = Z
    CreateTLVertex.rhw = rhw
    CreateTLVertex.Color = Color
    CreateTLVertex.Specular = Specular
    CreateTLVertex.tu = tu
    CreateTLVertex.tv = tv
End Function

Public Function MakeVector(X As Single, Y As Single, Z As Single) As D3DVECTOR
    MakeVector.X = X: MakeVector.Y = Y: MakeVector.Z = Z
End Function

Public Function MakeRect(Left As Single, Right As Single, Top As Single, Bottom As Single) As RECT
    MakeRect.Left = Left
    MakeRect.Right = Right
    MakeRect.Top = Top
    MakeRect.Bottom = Bottom
End Function

