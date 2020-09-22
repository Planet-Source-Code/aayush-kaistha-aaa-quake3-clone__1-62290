Attribute VB_Name = "modMain"
Option Explicit

Public Sub Main()
'This is the starting point of the program

Dim LastUpdated As Long, matTemp As D3DMATRIX

Initialize 'initialize the game

Fps.Last = GetTickCount
LastUpdated = GetTickCount

Do While bRunning
    LastUpdated = GetTickCount
    
    If GameOver Then
        MsgScreen
        DoEvents
    Else
        If Not Player.Dead Then
            CheckKeys

            'rotate the camera. yaw means rotating horizontally
            'pitch means rotating vertically. roll means about
            'z-axis like an aeroplane rotates
            'the yaw... function here sets our camera matrix according
            'yaw and pitch. then we transform this matrix into a vector
            'so that we can use it in our view settings
            D3DXMatrixIdentity matCam
            D3DXMatrixRotationYawPitchRoll matCam, Player.Rotation, CamPitch, 0
            D3DXVec3TransformCoord EyeLookDir, MakeVector(0, 0, 1), matCam
    
            'eyelookdir vector only gives the direction. we need to
            'add player pos to get to our current pos.
            D3DXVec3Add EyeLookAt, Player.Pos, EyeLookDir
            D3DXMatrixLookAtLH matView, Player.Pos, EyeLookAt, MakeVector(0, 1, 0)
            D3DDevice.SetTransform D3DTS_VIEW, matView
        End If
    
        D3DXMatrixScaling matWorld, 1, 1, 1
        D3DDevice.SetTransform D3DTS_WORLD, matWorld
    
        If Not Enemy.Dead Then
            If Not Player.Dead Then
                MoveEnemy
                CheckEnemyAttack
            End If
        End If
    
        CheckPowers 'check if our player has picked up ammo etc
        Render
    
        Fps.Count = Fps.Count + 1
        If ((GetTickCount - Fps.Last) >= 1000) Then
            Fps.Value = Fps.Count
            Fps.Count = 0
            Fps.Last = GetTickCount
        End If
    
        DoEvents
        'all movements in the game should be time dependent. the game
        'will run with different speed on different computers. this
        'prevents our game to run at varying speeds.
        Player.MoveSpeed = ((GetTickCount - LastUpdated) / 1000) * 150
        Enemy.MoveSpeed = ((GetTickCount - LastUpdated) / 1000) * 150
    End If
Loop

DestroyApp

End Sub

Public Sub Render()
Dim i As Long, j As Integer, temp As D3DMATRIX
Dim tmpAng As Single

D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, &H33, 1#, 0

D3DDevice.BeginScene
    
    For j = 0 To 24
        D3DDevice.SetTransform D3DTS_WORLD, matFloor(j)
        RenderXFile Floor
    Next
    
    For j = 0 To 9
        D3DDevice.SetTransform D3DTS_WORLD, matWall(j)
        RenderXFile Wall
    Next
    For j = 10 To 19
        D3DDevice.SetTransform D3DTS_WORLD, matWall(j)
        RenderXFile Wall1
    Next
    
    For j = 0 To 3
        D3DDevice.SetTransform D3DTS_WORLD, matWallBig(j)
        RenderXFile WallBig
    Next
    For j = 4 To 7
        D3DDevice.SetTransform D3DTS_WORLD, matWallBig(j)
        RenderXFile WallBig1
    Next
    
    If Enemy.Dead Then
        EnemyDie
    Else
        RenderEnemy
    End If
    
    If Player.Dead Then
        DrawPHit 'this is the transparent red color all over the screen
        D3DX.DrawText MainFont, &HFFFFFFAA, "Your Score : " + Str(Player.Score), MakeRect(150, 640, 220, 240), DT_TOP Or DT_LEFT
        D3DX.DrawText MainFont, &HFF8888FF, "Robo's Score : " + Str(Enemy.Score), MakeRect(150, 640, 260, 280), DT_TOP Or DT_LEFT

        If ((GetTickCount - Player.DieTime) >= 3000) Then
            Player.Dead = False
            Player.Health = 3
            Player.Ammo = 10
            'dont make our player alive in front of the enemy.
            'otherwise the enemy would kill u as soon as u r alive.
            If Enemy.Pos.X < 0 Then
                Player.Pos = WayPt(8)
            Else
                Player.Pos = WayPt(1)
            End If
            Player.Pos.Y = 5
        End If
    End If
    
    If Zoom = False Then 'display gun only if not zoomed
        tmpAng = Player.Rotation - (PI / 2)
    
        D3DXMatrixIdentity temp
        D3DXMatrixIdentity matGun
        D3DXMatrixTranslation temp, Player.Pos.X + Sin(Player.Rotation) - (Sin(tmpAng) * 2), 0, Player.Pos.Z + Cos(Player.Rotation) - (Cos(tmpAng) * 2)
        D3DXMatrixMultiply matGun, matCam, temp

        D3DDevice.SetTransform D3DTS_WORLD, matGun
        RenderXFile Gun
    End If
    
    RenderPowers 'draw ammo-box etc
    DrawCH 'draw cross hair
    DrawStats 'draw ammo health etc on top of the screen
    
    If fire Then
        If ((GetTickCount - FireTimer) <= 250) Then
            If Zoom = False Then DrawFlash
            DrawShot 'this is the line showing the path of the bullet
        Else
            fire = False
        End If
    End If
    
    If Player.Hit Then
        If ((GetTickCount - PHitTimer) <= 250) Then
            DrawPHit 'transparent red screen
        Else
            Player.Hit = False
        End If
    End If
    
    If Player.Ammo <= 5 Then _
        D3DX.DrawText MainFont, &HFFFFFFFF, WarningMsg, MakeRect(10, 300, 400, 420), DT_TOP Or DT_CENTER

D3DDevice.EndScene

D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0

End Sub

Public Sub RenderXFile(ByRef tmpObj As Object3D)
'draws objects loaded from x files
Dim i As Long
    For i = 0 To tmpObj.nMaterials - 1
        D3DDevice.SetTexture 0, tmpObj.Textures(i)
        D3DDevice.SetMaterial tmpObj.Materials(i)
        tmpObj.Mesh.DrawSubset i
    Next
End Sub

Public Sub MsgScreen()

D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, &HAA77AA, 1#, 0

D3DDevice.BeginScene
    DrawMsg
D3DDevice.EndScene

D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0

If Restart Then
    GameOver = False
    Restart = False
End If

End Sub

Public Sub DrawMsg()

Select Case EndState
    Case 0:
        D3DX.DrawText MainFont, &HFF00FF00, "A mad scientist has developed a Super Intelligent Robot. He claims that", MakeRect(10, 630, 30, 50), DT_TOP Or DT_CENTER
        D3DX.DrawText MainFont, &HFF00FF00, "the robot is unbeatable and no mortal on earth can defeat it.", MakeRect(10, 630, 60, 80), DT_TOP Or DT_CENTER
        D3DX.DrawText MainFont, &HFFFFFF00, "You being the toughest and bravest soldier have taken up the", MakeRect(10, 630, 90, 110), DT_TOP Or DT_CENTER
        D3DX.DrawText MainFont, &HFFFFFF00, "challenge to beat the robot and teach the scientist a lesson.", MakeRect(10, 630, 120, 140), DT_TOP Or DT_CENTER
        D3DX.DrawText MainFont, &HFF0000FF, "You must defeat the robot 10 times to win.", MakeRect(10, 630, 150, 170), DT_TOP Or DT_CENTER
        D3DX.DrawText MainFont, &HFFDD0000, "But Beware : if you fail, the scientist will set its robot to go out and", MakeRect(10, 630, 180, 200), DT_TOP Or DT_CENTER
        D3DX.DrawText MainFont, &HFFDD0000, "destroy the world.", MakeRect(10, 630, 210, 230), DT_TOP Or DT_CENTER
        D3DX.DrawText MainFont, &HFF00FFFF, "GO GO GO! The fate of mankind lies in your hands.", MakeRect(10, 630, 250, 270), DT_TOP Or DT_CENTER
    Case 1:
        D3DX.DrawText MainFont, &HFFFF0000, "CONGRATULATIONS!", MakeRect(10, 630, 150, 170), DT_TOP Or DT_CENTER
        D3DX.DrawText MainFont, &HFFFFFF00, "You are undoubtably the bravest man on earth.", MakeRect(10, 630, 200, 220), DT_TOP Or DT_CENTER
    Case 2:
        D3DX.DrawText MainFont, &HFFFF0000, "You are no match for the Super Robo", MakeRect(10, 630, 150, 170), DT_TOP Or DT_CENTER
        D3DX.DrawText MainFont, &HFFFFFF00, "Go and play with your toys somewhere else... LOSER", MakeRect(10, 630, 200, 220), DT_TOP Or DT_CENTER
    Case 3:
        D3DX.DrawText MainFont, &HFFFFFF00, "You have terminated the game.", MakeRect(10, 630, 200, 220), DT_TOP Or DT_CENTER
        D3DX.DrawText MainFont, &HFF00FFFF, "You ran away when the world needed you the most... LOSER", MakeRect(10, 630, 230, 250), DT_TOP Or DT_CENTER
End Select

D3DX.DrawText MainFont, &HFF000000, "================================================", MakeRect(10, 630, 270, 290), DT_TOP Or DT_CENTER
D3DX.DrawText MainFont, &HFF000000, "================================================", MakeRect(10, 630, 370, 390), DT_TOP Or DT_CENTER
If EndState = 0 Then
    D3DX.DrawText MainFont, &HFF00DD00, "Press Enter to Play.", MakeRect(10, 630, 300, 320), DT_TOP Or DT_CENTER
Else
    D3DX.DrawText MainFont, &HFF00DD00, "Press Enter to Play Again.", MakeRect(10, 630, 300, 320), DT_TOP Or DT_CENTER
End If
D3DX.DrawText MainFont, &HFFDD0000, "Press Escape to Quit.", MakeRect(10, 630, 330, 350), DT_TOP Or DT_CENTER
D3DX.DrawText MainFont, &HFFAA00AA, "Game Designed and programmed by : Aayush Kaistha.", MakeRect(10, 630, 400, 420), DT_TOP Or DT_CENTER
D3DX.DrawText MainFont, &HFFAA00AA, "UIET, Panjab University, Chandigarh. - aayushk_007@yahoo.com", MakeRect(10, 630, 430, 450), DT_TOP Or DT_CENTER
End Sub

Public Sub RenderEnemy()
Dim matTemp As D3DMATRIX

EneDir.X = Int(Enemy.Pos.X - Player.Pos.X)
EneDir.Y = 0
EneDir.Z = Int(Enemy.Pos.Z - Player.Pos.Z)

If EneDir.X = 0 Then
    Enemy.Rotation = 90 * Rad
Else
    Enemy.Rotation = (PI / 2) - Atn(EneDir.Z / EneDir.X)
End If
If EneDir.X > 0 Then Enemy.Rotation = Enemy.Rotation + PI

D3DXMatrixIdentity matEneBody
D3DXMatrixIdentity matTemp
D3DXMatrixRotationY matTemp, Enemy.Rotation
D3DXMatrixMultiply matEneBody, matEneBody, matTemp
D3DXMatrixIdentity matTemp
D3DXMatrixTranslation matTemp, Enemy.Pos.X, Enemy.Pos.Y, Enemy.Pos.Z
D3DXMatrixMultiply matEneBody, matEneBody, matTemp

D3DXMatrixIdentity matEneHead
D3DXMatrixIdentity matTemp
D3DXMatrixRotationY matTemp, Enemy.Rotation
D3DXMatrixMultiply matEneHead, matEneHead, matTemp
D3DXMatrixIdentity matTemp
D3DXMatrixTranslation matTemp, Enemy.Pos.X, Enemy.Pos.Y + 25, Enemy.Pos.Z
D3DXMatrixMultiply matEneHead, matEneHead, matTemp

D3DXMatrixIdentity matEneGun
D3DXMatrixIdentity matTemp
D3DXMatrixRotationY matTemp, Enemy.Rotation
D3DXMatrixMultiply matEneGun, matEneGun, matTemp
D3DXMatrixIdentity matTemp
D3DXMatrixTranslation matTemp, Enemy.Pos.X, Enemy.Pos.Y + 17, Enemy.Pos.Z
D3DXMatrixMultiply matEneGun, matEneGun, matTemp

Coll_Obj(53).Radius = Ene_Head.Radius
With Coll_Obj(53).Center
    .X = Enemy.Pos.X: .Y = Enemy.Pos.Y + 25: .Z = Enemy.Pos.Z
End With
Coll_Obj(54).Radius = Ene_Body.Radius
With Coll_Obj(54).Center
    .X = Enemy.Pos.X: .Y = Enemy.Pos.Y + Ene_Body.Radius.Y: .Z = Enemy.Pos.Z
End With

D3DDevice.SetTransform D3DTS_WORLD, matEneBody
RenderXFile Ene_Body
D3DDevice.SetTransform D3DTS_WORLD, matEneHead
RenderXFile Ene_Head
D3DDevice.SetTransform D3DTS_WORLD, matEneGun
RenderXFile Ene_Gun

End Sub

Public Sub RenderPowers()
Dim matTemp As D3DMATRIX

D3DXMatrixIdentity matAmmo
D3DXMatrixIdentity matTemp
D3DXMatrixTranslation matTemp, AmmoPos.X, -15, AmmoPos.Z
D3DXMatrixMultiply matAmmo, matAmmo, matTemp

D3DDevice.SetTransform D3DTS_WORLD, matAmmo
RenderXFile Ammo

If DispHealth Then
    D3DXMatrixIdentity matHealth
    D3DXMatrixIdentity matTemp
    D3DXMatrixTranslation matTemp, HealthPos.X, -15, HealthPos.Z
    D3DXMatrixMultiply matHealth, matHealth, matTemp

    D3DDevice.SetTransform D3DTS_WORLD, matHealth
    RenderXFile Health
End If

End Sub

Public Sub CheckPowers()

If Abs(Player.Pos.X - AmmoPos.X) <= 10 Then
    If Abs(Player.Pos.Z - AmmoPos.Z) <= 10 Then
        If Player.Pos.X < 0 Then
            AmmoPos = WayPt(10)
        Else
            AmmoPos = WayPt(5)
        End If
        Player.Ammo = Player.Ammo + 10
        If Player.Ammo > 20 Then Player.Ammo = 20
        sndReload.Play DSBPLAY_DEFAULT
    End If
End If

If DispHealth Then
    If Abs(Player.Pos.X - HealthPos.X) <= 10 Then
        If Abs(Player.Pos.Z - HealthPos.Z) <= 10 Then
            If Player.Pos.X < 0 Then
                HealthPos = WayPt(15)
            Else
                HealthPos = WayPt(0)
            End If
            Player.Health = Player.Health + 1
            If Player.Health > 3 Then Player.Health = 3
            sndPHeal.Play DSBPLAY_DEFAULT
        End If
    End If
End If

End Sub

Public Sub CheckKeys()
Dim tmpAng As Integer

If UpKey Then
    If Colliding(Player.Pos.X + (Sin(Player.Rotation) * 10), Player.Pos.Z + (Cos(Player.Rotation) * Player.MoveSpeed)) = False Then _
        Player.Pos.X = Player.Pos.X + (Sin(Player.Rotation) * Player.MoveSpeed)
    If Colliding(Player.Pos.X + (Sin(Player.Rotation) * Player.MoveSpeed), Player.Pos.Z + (Cos(Player.Rotation) * 10)) = False Then _
        Player.Pos.Z = Player.Pos.Z + (Cos(Player.Rotation) * Player.MoveSpeed)
End If
If DownKey Then
    If Colliding(Player.Pos.X - (Sin(Player.Rotation) * 10), Player.Pos.Z - (Cos(Player.Rotation) * Player.MoveSpeed)) = False Then _
        Player.Pos.X = Player.Pos.X - (Sin(Player.Rotation) * Player.MoveSpeed)
    If Colliding(Player.Pos.X - (Sin(Player.Rotation) * Player.MoveSpeed), Player.Pos.Z - (Cos(Player.Rotation) * 10)) = False Then _
        Player.Pos.Z = Player.Pos.Z - (Cos(Player.Rotation) * Player.MoveSpeed)
End If
If LeftKey Then
    tmpAng = Player.Rotation + (PI / 2)
    If Colliding(Player.Pos.X - (Sin(tmpAng) * 10), Player.Pos.Z - (Cos(tmpAng) * Player.MoveSpeed)) = False Then _
        Player.Pos.X = Player.Pos.X - (Sin(tmpAng) * Player.MoveSpeed)
    If Colliding(Player.Pos.X - (Sin(tmpAng) * Player.MoveSpeed), Player.Pos.Z - (Cos(tmpAng) * 10)) = False Then _
        Player.Pos.Z = Player.Pos.Z - (Cos(tmpAng) * Player.MoveSpeed)
End If
If RightKey Then
    tmpAng = Player.Rotation - (PI / 2)
    If Colliding(Player.Pos.X - (Sin(tmpAng) * 10), Player.Pos.Z - (Cos(tmpAng) * Player.MoveSpeed)) = False Then _
        Player.Pos.X = Player.Pos.X - (Sin(tmpAng) * Player.MoveSpeed)
    If Colliding(Player.Pos.X - (Sin(tmpAng) * Player.MoveSpeed), Player.Pos.Z - (Cos(tmpAng) * 10)) = False Then _
        Player.Pos.Z = Player.Pos.Z - (Cos(tmpAng) * Player.MoveSpeed)
End If

If SKey Then D3DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
If WKey Then D3DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_WIREFRAME

End Sub

Public Function Colliding(X As Single, Y As Single) As Boolean
Dim j As Integer

For j = 0 To nColl_Obj - 2
    If (Abs(X - Coll_Obj(j).Center.X) <= Coll_Obj(j).Radius.X) Then
        If (Abs(Y - Coll_Obj(j).Center.Z) <= Coll_Obj(j).Radius.Z) Then
            Colliding = True
            Exit Function
        End If
    End If
Next

Colliding = False

End Function

Public Sub CheckBulletColl(P As Boolean)
Dim StartPos As D3DVECTOR, U As Single, V As Single
Dim FI As Long, nHits As Long, i As Integer, Dist As Single
Dim Check As Boolean, tmpLookDir As D3DVECTOR

If P Then
    tmpLookDir = EyeLookDir
Else
    tmpLookDir = EneDir
End If

Check = False
BulHit = 0: HitDist = 1000

For i = 0 To 24
    StartPos = MakeVector(Player.Pos.X - Coll_Obj(i).Center.X, 25, Player.Pos.Z - Coll_Obj(i).Center.Z)
    D3DX.Intersect Floor.Mesh, StartPos, tmpLookDir, BulHit, FI, U, V, Dist, nHits
    If BulHit Then
        If Dist < HitDist Then
            HitDist = Dist
            Check = True
        End If
    End If
Next
For i = 25 To 34
    StartPos = MakeVector(Player.Pos.X - Coll_Obj(i).Center.X, 25, Player.Pos.Z - Coll_Obj(i).Center.Z)
    D3DX.Intersect Wall.Mesh, StartPos, tmpLookDir, BulHit, FI, U, V, Dist, nHits
    If BulHit Then
        If Dist < HitDist Then
            HitDist = Dist
            Check = True
        End If
    End If
Next
For i = 35 To 44
    StartPos = MakeVector(Player.Pos.X - Coll_Obj(i).Center.X, 25, Player.Pos.Z - Coll_Obj(i).Center.Z)
    D3DX.Intersect Wall1.Mesh, StartPos, tmpLookDir, BulHit, FI, U, V, Dist, nHits
    If BulHit Then
        If Dist < HitDist Then
            HitDist = Dist
            Check = True
        End If
    End If
Next

CanHit = False
For i = 45 To 48
    StartPos = MakeVector(Player.Pos.X - Coll_Obj(i).Center.X, 25, Player.Pos.Z - Coll_Obj(i).Center.Z)
    D3DX.Intersect WallBig.Mesh, StartPos, tmpLookDir, BulHit, FI, U, V, Dist, nHits
    If BulHit Then
        If Dist < HitDist Then
            HitDist = Dist
            Check = True
        End If
    End If
Next
For i = 49 To 52
    StartPos = MakeVector(Player.Pos.X - Coll_Obj(i).Center.X, 25, Player.Pos.Z - Coll_Obj(i).Center.Z)
    D3DX.Intersect WallBig1.Mesh, StartPos, tmpLookDir, BulHit, FI, U, V, Dist, nHits
    If BulHit Then
        If Dist < HitDist Then
            HitDist = Dist
            Check = True
        End If
    End If
Next

'Enemy Head
StartPos = MakeVector(Player.Pos.X - Coll_Obj(53).Center.X, 0, Player.Pos.Z - Coll_Obj(53).Center.Z)
D3DX.Intersect Ene_Head.Mesh, StartPos, tmpLookDir, BulHit, FI, U, V, Dist, nHits
If BulHit Then
    If Dist < HitDist Then
        HitDist = Dist
        Check = True
        If (P = False And HitDist <= 300) Then
            CanHit = True
        Else
            EnemyHit
        End If
    End If
End If

'Enemy Body
StartPos = MakeVector(Player.Pos.X - Coll_Obj(54).Center.X, 25, Player.Pos.Z - Coll_Obj(54).Center.Z)
D3DX.Intersect Ene_Body.Mesh, StartPos, tmpLookDir, BulHit, FI, U, V, Dist, nHits
If BulHit Then
    If Dist < HitDist Then
        HitDist = Dist
        Check = True
        EnemyHit
    End If
End If
    
If Not Check Then HitDist = 500

HitPos.X = Player.Pos.X + (EyeLookDir.X * HitDist)
HitPos.Y = Player.Pos.Y + (EyeLookDir.Y * HitDist)
HitPos.Z = Player.Pos.Z + (EyeLookDir.Z * HitDist)

End Sub

Public Sub FireBullet()

fire = True
Player.Ammo = Player.Ammo - 1
If Player.Ammo = 5 Then sndWarning.Play DSBPLAY_DEFAULT
If Player.Ammo <= 5 Then WarningMsg = "LOW AMMO WARNING"
If Player.Ammo = 0 Then WarningMsg = "OUT OF AMMO"
FireTimer = GetTickCount
sndShoot.Stop
sndShoot.SetCurrentPosition 0
sndShoot.Play DSBPLAY_DEFAULT
If Not Enemy.Dead Then CheckBulletColl True

Dim tmpAng As Single
tmpAng = Player.Rotation - (PI / 2)
ShotStart = MakeVector(Player.Pos.X + (Sin(Player.Rotation) * FlashPos) - (Sin(tmpAng) * 2), -FlashPos * Sin(CamPitch), Player.Pos.Z + (Cos(Player.Rotation) * FlashPos) - (Cos(tmpAng) * 2))
ShotEnd = HitPos

End Sub

Public Sub DrawFlash()
Dim matTemp As D3DMATRIX, tmpAng As Single

D3DDevice.SetRenderState D3DRS_LIGHTING, 0
D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
D3DDevice.SetVertexShader FVF_LVERTEX
        
tmpAng = Player.Rotation - (PI / 2)

D3DXMatrixIdentity matTemp
D3DXMatrixIdentity matFlash
D3DXMatrixTranslation matTemp, Player.Pos.X + (Sin(Player.Rotation) * FlashPos) - (Sin(tmpAng) * 2), -FlashPos * Sin(CamPitch), Player.Pos.Z + (Cos(Player.Rotation) * FlashPos) - (Cos(tmpAng) * 2)
D3DXMatrixMultiply matFlash, matCam, matTemp

D3DDevice.SetTransform D3DTS_WORLD, matFlash
D3DDevice.SetTexture 0, FlTex
D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, 4, flash(0), Len(flash(0))

D3DDevice.SetTexture 0, Nothing
D3DDevice.SetVertexShader FVF_VERTEX
D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 0
D3DDevice.SetRenderState D3DRS_LIGHTING, 1

End Sub

Public Sub DrawPHit()
Dim matTemp As D3DMATRIX, tmpAng As Single
Dim matHit As D3DMATRIX

D3DDevice.SetRenderState D3DRS_LIGHTING, 0
D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
D3DDevice.SetVertexShader FVF_LVERTEX

D3DXMatrixIdentity matTemp
D3DXMatrixIdentity matHit
D3DXMatrixTranslation matTemp, Player.Pos.X + (Sin(Player.Rotation) * 5), 5 + (-Sin(CamPitch) * 10), (Player.Pos.Z + Cos(Player.Rotation) * 5)
D3DXMatrixMultiply matHit, matCam, matTemp

D3DDevice.SetTransform D3DTS_WORLD, matHit
D3DDevice.SetTexture 0, hitTex
D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, 4, PHit(0), Len(PHit(0))

D3DDevice.SetTexture 0, Nothing
D3DDevice.SetVertexShader FVF_VERTEX
D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 0
D3DDevice.SetRenderState D3DRS_LIGHTING, 1

End Sub

Public Sub DrawShot()
On Error Resume Next

Shot(0) = CreateLitVertex(ShotStart.X, ShotStart.Y, ShotStart.Z, &H999999, 0, 0, 0)
Shot(1) = CreateLitVertex(ShotEnd.X, ShotEnd.Y, ShotEnd.Z, &H999999, 0, 0, 0)

D3DDevice.SetRenderState D3DRS_LIGHTING, 0
D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
D3DDevice.SetVertexShader FVF_LVERTEX

D3DDevice.SetTransform D3DTS_WORLD, matWorld
D3DDevice.SetTexture 0, Nothing
D3DDevice.DrawPrimitiveUP D3DPT_LINELIST, 1, Shot(0), Len(Shot(0))

D3DDevice.SetTexture 0, Nothing
D3DDevice.SetVertexShader FVF_VERTEX
D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 0
D3DDevice.SetRenderState D3DRS_LIGHTING, 1

End Sub

Public Sub DrawCH()

D3DDevice.SetRenderState D3DRS_LIGHTING, 0
D3DDevice.SetVertexShader FVF_TLVERTEX
D3DDevice.SetTexture 0, Nothing

D3DDevice.DrawPrimitiveUP D3DPT_LINELIST, 4, CH(0), Len(CH(0))

D3DDevice.SetRenderState D3DRS_LIGHTING, 1
D3DDevice.SetVertexShader FVF_VERTEX

End Sub

Public Sub DrawStats()

D3DDevice.SetRenderState D3DRS_LIGHTING, 0
D3DDevice.SetVertexShader FVF_TLVERTEX

D3DDevice.SetTexture 0, StatTex(0)
D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Stats.Fps(0), Len(Stats.Fps(0))
D3DX.DrawText MainFont, &HFFFF0000, "FPS", MakeRect(0, 150, 20, 40), DT_TOP Or DT_CENTER
D3DX.DrawText MainFont, &HFFFF0000, Str(Fps.Value), MakeRect(0, 150, 40, 60), DT_TOP Or DT_CENTER

D3DDevice.SetTexture 0, StatTex(1)
D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Stats.Ammo(0), Len(Stats.Ammo(0))
D3DX.DrawText MainFont, &HFF0000FF, "HEALTH", MakeRect(160, 310, 20, 40), DT_TOP Or DT_CENTER
D3DX.DrawText MainFont, &HFF0000FF, Str(Player.Health), MakeRect(160, 310, 40, 60), DT_TOP Or DT_CENTER

D3DDevice.SetTexture 0, StatTex(2)
D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Stats.Health(0), Len(Stats.Health(0))
D3DX.DrawText MainFont, &HFF00FF00, "AMMO", MakeRect(320, 470, 20, 40), DT_TOP Or DT_CENTER
D3DX.DrawText MainFont, &HFF00FF00, Str(Player.Ammo), MakeRect(320, 470, 40, 60), DT_TOP Or DT_CENTER

D3DDevice.SetTexture 0, StatTex(3)
D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Stats.Score(0), Len(Stats.Score(0))
D3DX.DrawText MainFont, &HFFFFFF00, "FRAGS", MakeRect(480, 630, 20, 40), DT_TOP Or DT_CENTER
D3DX.DrawText MainFont, &HFFFFFF00, Str(Player.Score) + " (" + Trim$(Str(Enemy.Score)) + ")", MakeRect(480, 630, 40, 60), DT_TOP Or DT_CENTER

D3DDevice.SetTexture 0, Nothing
D3DDevice.SetRenderState D3DRS_LIGHTING, 1
D3DDevice.SetVertexShader FVF_VERTEX

End Sub

Public Sub MoveEnemy()
Dim Dir As D3DVECTOR, tmp As Double

If Int(Abs(Enemy.Pos.X - WayPt(Curpos).X)) <= 1 Then
    If Int(Abs(Enemy.Pos.Z - WayPt(Curpos).Z)) <= 1 Then
        Curpos = Curpos + 1
        If Curpos > nWayPts Then Curpos = 0
    End If
End If

Dir.X = WayPt(Curpos).X - Enemy.Pos.X
Dir.Z = WayPt(Curpos).Z - Enemy.Pos.Z
tmp = Sqr((Dir.X * Dir.X) + (Dir.Z * Dir.Z))
Dir.X = Dir.X / tmp
Dir.Z = Dir.Z / tmp

If Colliding(Enemy.Pos.X + (Dir.X * 10), Enemy.Pos.Z + (Dir.Z * 10)) = True Then
    Curpos = Curpos - 1
    If Curpos < 0 Then Curpos = 2
Else
    Enemy.Pos.X = Enemy.Pos.X + (Dir.X * Enemy.MoveSpeed)
    Enemy.Pos.Z = Enemy.Pos.Z + (Dir.Z * Enemy.MoveSpeed)
End If

End Sub

Public Sub CheckEnemyAttack()

CheckBulletColl False
If CanHit Then
    If Player.Hit = False Then
        If ((GetTickCount - EnemySaw) >= 3000) Then
            EnemySaw = GetTickCount
        ElseIf ((GetTickCount - EnemySaw) >= 1500) Then
            sndEneShoot.Play DSBPLAY_DEFAULT
            Player.Hit = True
            Player.Health = Player.Health - 1
            
            If ((Player.Health Mod 2) = 0) Then
                DispHealth = True
            Else
                DispHealth = False
            End If
            
            If Player.Health < 0 Then
                Player.Dead = True
                sndPDie.Play DSBPLAY_DEFAULT
                Player.DieTime = GetTickCount
                Enemy.Score = Enemy.Score + 1
                If Enemy.Score = 10 Then
                    EndState = 2
                    GameOver = True
                End If
            End If
            
            EnemySaw = GetTickCount
            PHitTimer = GetTickCount
        End If
    End If
End If

End Sub

Public Sub EnemyHit()
Dim i As Integer

For i = 0 To 12
    If Colliding(Enemy.Pos.X + (EyeLookDir.X * 5), Enemy.Pos.Z) = False Then _
        Enemy.Pos.X = Enemy.Pos.X + (EyeLookDir.X * 5)
    If Colliding(Enemy.Pos.X, Enemy.Pos.Z + (EyeLookDir.X * 5)) = False Then _
        Enemy.Pos.Z = Enemy.Pos.Z + (EyeLookDir.Z * 5)
Next

Enemy.Health = Enemy.Health - 1
If Enemy.Health < 0 Then
    Enemy.Dead = True
    Player.Score = Player.Score + 1
    If Player.Score = 10 Then
        EndState = 1
        GameOver = True
    End If
    Enemy.DieTime = GetTickCount
    For i = 0 To 5
        PartPos(i) = Enemy.Pos
        PartPos(i).Y = 0
    Next
    If Player.Score = 10 Then
        sndWin.Play DSBPLAY_DEFAULT
    Else
        sndEneXplo.Play DSBPLAY_DEFAULT
    End If
    Exit Sub
End If

sndEneHit.Play DSBPLAY_DEFAULT
Curpos = Curpos + 1
If Curpos > nWayPts Then Curpos = nWayPts - 2

End Sub

Public Sub EnemyDie()
Dim matTemp As D3DMATRIX, tmpAng As Single
Dim i As Integer

If ((GetTickCount - Enemy.DieTime) >= 3000) Then
    Enemy.Dead = False
    Enemy.Health = 3
    If Player.Pos.X < 0 Then
        Enemy.Pos = WayPt(9)
        Curpos = 8
    Else
        Enemy.Pos = WayPt(1)
        Curpos = 0
    End If
    EnemySaw = GetTickCount
    Enemy.Pos.Y = -20
    Exit Sub
End If

For i = 0 To 5
    D3DXMatrixIdentity matPart(i)
    D3DXMatrixIdentity matTemp
    D3DXMatrixTranslation matTemp, PartPos(i).X, PartPos(i).Y, PartPos(i).Z
    D3DXMatrixMultiply matPart(i), matPart(i), matTemp

    D3DDevice.SetTransform D3DTS_WORLD, matPart(i)
    RenderXFile Part
Next

tmpAng = Enemy.Rotation + (PI / 2)
PartPos(0).X = PartPos(0).X + (Sin(tmpAng) * 0.5)
PartPos(0).Y = PartPos(0).Y + 0.5
PartPos(0).Z = PartPos(0).Z + (Cos(tmpAng) * 0.5)
PartPos(1).X = PartPos(1).X + (Sin(tmpAng) * 0.5)
PartPos(1).Z = PartPos(1).Z + (Cos(tmpAng) * 0.5)
PartPos(2).X = PartPos(2).X + (Sin(tmpAng) * 0.5)
PartPos(2).Y = PartPos(2).Y - 0.5
PartPos(2).Z = PartPos(2).Z + (Cos(tmpAng) * 0.5)

tmpAng = Enemy.Rotation - (PI / 2)
PartPos(3).X = PartPos(3).X + (Sin(tmpAng) * 0.5)
PartPos(3).Y = PartPos(3).Y + 0.5
PartPos(3).Z = PartPos(3).Z + (Cos(tmpAng) * 0.5)
PartPos(4).X = PartPos(4).X + (Sin(tmpAng) * 0.5)
PartPos(4).Z = PartPos(4).Z + (Cos(tmpAng) * 0.5)
PartPos(5).X = PartPos(5).X + (Sin(tmpAng) * 0.5)
PartPos(5).Y = PartPos(5).Y - 0.5
PartPos(5).Z = PartPos(5).Z + (Cos(tmpAng) * 0.5)

End Sub

Public Sub DestroyApp()
On Error Resume Next

If hEvent <> 0 Then Dx.DestroyEvent hEvent
Set DIDevice = Nothing
Set DI = Nothing

Set D3DX = Nothing
Set D3DDevice = Nothing
Set D3D = Nothing
Set Dx = Nothing

ShowCursor 1

End

End Sub
