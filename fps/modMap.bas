Attribute VB_Name = "modMap"
Option Explicit

Public Sub LoadMap()

Dim Fx As Integer, Fy As Integer, j As Integer, i As Integer
Dim matTemp As D3DMATRIX, col As Integer

'the floor
Fx = -400: Fy = -400: j = 0
col = 0
While Fy <= 400
    While Fx <= 400
        D3DXMatrixIdentity matFloor(j)
        D3DXMatrixIdentity matTemp
        D3DXMatrixTranslation matTemp, Fx, -20, Fy
        D3DXMatrixMultiply matFloor(j), matFloor(j), matTemp
        
        Coll_Obj(col).Center.X = Fx
        Coll_Obj(col).Center.Y = -20
        Coll_Obj(col).Center.Z = Fy
        col = col + 1
        
        Fx = Fx + 200: j = j + 1
    Wend
    Fy = Fy + 200: Fx = -400
Wend
'=============================================

'the boundry walls
Fx = -400: Fy = -500: j = 0
For i = 0 To 1
    While Fx <= 400
        D3DXMatrixIdentity matWall(j)
        D3DXMatrixIdentity matTemp
        D3DXMatrixTranslation matTemp, Fx, -20, Fy
        D3DXMatrixMultiply matWall(j), matWall(j), matTemp
        
        Coll_Obj(col).Radius = Wall.Radius
        Coll_Obj(col).Center.X = Fx
        Coll_Obj(col).Center.Y = Coll_Obj(col).Radius.Y - 20
        Coll_Obj(col).Center.Z = Fy
        col = col + 1
        
        Fx = Fx + 200: j = j + 1
    Wend
    Fy = 500: Fx = -400
Next

Fy = -400: Fx = -500
For i = 0 To 1
    While Fy <= 400
        D3DXMatrixIdentity matWall(j)
        D3DXMatrixIdentity matTemp
        D3DXMatrixTranslation matTemp, Fx, -20, Fy
        D3DXMatrixMultiply matWall(j), matWall(j), matTemp
        
        With Coll_Obj(col).Radius
            .X = Wall1.Radius.X: .Y = Wall1.Radius.Y: .Z = Wall1.Radius.Z
        End With
        Coll_Obj(col).Center.X = Fx
        Coll_Obj(col).Center.Y = Coll_Obj(col).Radius.Y - 20
        Coll_Obj(col).Center.Z = Fy
        col = col + 1
        
        Fy = Fy + 200: j = j + 1
    Wend
    Fx = 500: Fy = -400
Next
'==================================================

'the Big walls
Fx = -300: Fy = 0: j = 0
While Fx <= 300
    D3DXMatrixIdentity matWallBig(j)
    D3DXMatrixIdentity matTemp
    D3DXMatrixTranslation matTemp, Fx, -20, Fy
    D3DXMatrixMultiply matWallBig(j), matWallBig(j), matTemp
    
    Coll_Obj(col).Radius = WallBig.Radius
    Coll_Obj(col).Center.X = Fx
    Coll_Obj(col).Center.Y = Coll_Obj(col).Radius.Y - 20
    Coll_Obj(col).Center.Z = Fy
    col = col + 1
        
    Fx = Fx + 200: j = j + 1
Wend

Fy = -300: Fx = 0
While Fy <= 300
    D3DXMatrixIdentity matWallBig(j)
    D3DXMatrixIdentity matTemp
    D3DXMatrixTranslation matTemp, Fx, -20, Fy
    D3DXMatrixMultiply matWallBig(j), matWallBig(j), matTemp
    
    With Coll_Obj(col).Radius
        .X = WallBig1.Radius.X: .Y = WallBig1.Radius.Y: .Z = WallBig1.Radius.Z
    End With
    Coll_Obj(col).Center.X = Fx
    Coll_Obj(col).Center.Y = Coll_Obj(col).Radius.Y - 20
    Coll_Obj(col).Center.Z = Fy
    col = col + 1
        
    Fy = Fy + 200: j = j + 1
Wend
'==================================================

End Sub

Public Sub SetLights()

Light.Type = D3DLIGHT_DIRECTIONAL
Light.Position = MakeVector(0, 1000, 0)
Light.Direction = MakeVector(0, -1, 0)
Light.Range = 1
With Light.diffuse
    .a = 1: .r = 0: .g = 1: .b = 1
End With

D3DDevice.SetLight 0, Light
D3DDevice.LightEnable 0, 1

End Sub
