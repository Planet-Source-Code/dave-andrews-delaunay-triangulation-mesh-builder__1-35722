VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDX3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public mdownX As Single
Public mDownY As Single
Public mMouseDown As Boolean

Public mDX7 As DirectX7
Public mDrm As Direct3DRM3
Public mDD4 As DirectDraw4
Public mDD7 As DirectDraw7
Public mFrs As Direct3DRMFrame3
Public mFrC As Direct3DRMFrame3
Public mFrO As Direct3DRMFrame3
Public mFrL As Direct3DRMFrame3
Public mDev As Direct3DRMDevice3
Public mVpt As Direct3DRMViewport2
Public DxL1 As Direct3DRMLight
Public DxL2 As Direct3DRMLight
Public DXClipper As DirectDrawClipper
Public DxMeshB As Direct3DRMMeshBuilder3
Dim mFrontBuffer As DirectDrawSurface4
Dim mBackBuffer As DirectDrawSurface4
Dim IsInit As Boolean

Public Function CreateBoxMesh(X As Single, Y As Single, z As Single) As Direct3DRMMeshBuilder3
Dim f As Direct3DRMFace2
Set CreateBoxMesh = mDrm.CreateMeshBuilder
Set f = mDrm.CreateFace
f.AddVertex -(X / 2), -(Y / 2), -(z / 2)
f.AddVertex -(X / 2), (Y / 2), -(z / 2)
f.AddVertex (X / 2), (Y / 2), -(z / 2)
f.AddVertex (X / 2), -(Y / 2), -(z / 2)
CreateBoxMesh.AddFace f
Set f = mDrm.CreateFace
f.AddVertex -(X / 2), -(Y / 2), (z / 2)
f.AddVertex (X / 2), -(Y / 2), (z / 2)
f.AddVertex (X / 2), (Y / 2), (z / 2)
f.AddVertex -(X / 2), (Y / 2), (z / 2)
CreateBoxMesh.AddFace f
Set f = mDrm.CreateFace
f.AddVertex (X / 2), -(Y / 2), -(z / 2)
f.AddVertex (X / 2), (Y / 2), -(z / 2)
f.AddVertex (X / 2), (Y / 2), (z / 2)
f.AddVertex (X / 2), -(Y / 2), (z / 2)
CreateBoxMesh.AddFace f
Set f = mDrm.CreateFace
f.AddVertex -(X / 2), -(Y / 2), -(z / 2)
f.AddVertex -(X / 2), -(Y / 2), (z / 2)
f.AddVertex -(X / 2), (Y / 2), (z / 2)
f.AddVertex -(X / 2), (Y / 2), -(z / 2)
CreateBoxMesh.AddFace f
Set f = mDrm.CreateFace
f.AddVertex -(X / 2), (Y / 2), -(z / 2)
f.AddVertex -(X / 2), (Y / 2), (z / 2)
f.AddVertex (X / 2), (Y / 2), (z / 2)
f.AddVertex (X / 2), (Y / 2), -(z / 2)
CreateBoxMesh.AddFace f
Set f = mDrm.CreateFace
f.AddVertex -(X / 2), -(Y / 2), -(z / 2)
f.AddVertex (X / 2), -(Y / 2), -(z / 2)
f.AddVertex (X / 2), -(Y / 2), (z / 2)
f.AddVertex -(X / 2), -(Y / 2), (z / 2)
CreateBoxMesh.AddFace f
Set f = Nothing
End Function

Public Sub InitDx(Canvas As Object)
'''Dim ddsd As DDSURFACEDESC2
'''Dim ddsdB As DDSURFACEDESC2
Set mDX7 = New DirectX7
Set mDrm = mDX7.Direct3DRMCreate
Set mDD7 = mDX7.DirectDrawCreate("")
'''Set mDD4 = mDX7.DirectDraw4Create("")
'''mDD4.SetCooperativeLevel Canvas.hWnd, DDSCL_NORMAL
'''ddsd.lFlags = DDSD_CAPS
'''ddsd.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
'''Set mFrontBuffer = mDD4.CreateSurface(ddsd)
'-Set Up Camera Frames And Lights-
Set mFrs = mDrm.CreateFrame(Nothing)
Set mFrC = mDrm.CreateFrame(mFrs)
Set mFrO = mDrm.CreateFrame(mFrs)
Set mFrL = mDrm.CreateFrame(mFrs)
Set DxL1 = mDrm.CreateLightRGB(D3DRMLIGHT_DIRECTIONAL, 0.8, 0.8, 0.8)
Set DxL2 = mDrm.CreateLightRGB(D3DRMLIGHT_AMBIENT, 0.8, 0.8, 0.8)
Canvas.ScaleMode = vbPixels
mFrL.AddLight DxL1
mFrL.AddLight DxL2
'-------Create Display--------
Set mVpt = Nothing
Set mDev = Nothing
'''Set DXClipper = mDD4.CreateClipper(0)
Set DXClipper = mDD7.CreateClipper(0)
Canvas.ScaleMode = vbPixels
DXClipper.SetHWnd Canvas.hWnd
'''mFrontBuffer.SetClipper DXClipper
'''ddsdB.lWidth = Canvas.ScaleWidth
'''ddsdB.lHeight = Canvas.ScaleHeight
'''ddsdB.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
'''ddsdB.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_3DDEVICE Or DDSCAPS_SYSTEMMEMORY
'''Set mBackBuffer = mDD4.CreateSurface(ddsdB)
'''Set mDev = mDrm.CreateDeviceFromSurface("IID_IDirect3DRGBDevice", mDD4, mBackBuffer, 0)
Set mDev = mDrm.CreateDeviceFromClipper(DXClipper, "", Canvas.ScaleWidth, Canvas.ScaleHeight)
Set mVpt = mDrm.CreateViewport(mDev, mFrC, 0, 0, Canvas.ScaleWidth, Canvas.ScaleHeight)
IsInit = True
End Sub
Public Function CreateUpdateableTexture(w As Long, h As Long, sfile As String) As Direct3DRMTexture3
    On Local Error GoTo errOut
    Dim sLoadFile As String
    Dim ddsd As DDSURFACEDESC2
    Dim SurfaceObject As DirectDrawSurface4
    Dim out As Direct3DRMTexture3
    Dim Init As Boolean
    
    ddsd.lFlags = DDSD_CAPS
    ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    If (w <> 0) Then
        ddsd.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
        ddsd.lWidth = w
        ddsd.lHeight = h
    End If
    
    If sfile = "" Then
        Set SurfaceObject = mDD4.CreateSurface(ddsd)
    Else
        Set SurfaceObject = mDD4.CreateSurfaceFromFile(sfile, ddsd)
    End If
    Set out = mDrm.CreateTextureFromSurface(SurfaceObject)
    Set CreateUpdateableTexture = out
    
errOut:
    'MsgBox Err.Description
    Set SurfaceObject = Nothing
End Function
Sub AddLine3D(dMesh As Direct3DRMMeshBuilder3, X1, Y1, Z1, X2, Y2, Z2)
Dim f As Direct3DRMFace2
Set f = mDrm.CreateFace
f.AddVertex X1, Y1, Z1
f.AddVertex (X1 + X2) / 2, (Y1 + Y2) / 2, (Z1 + Z2) / 2
f.AddVertex X2, Y2, Z2
dMesh.AddFace f
Set f = mDrm.CreateFace
f.AddVertex X1, Y1, Z1
f.AddVertex X2, Y2, Z2
f.AddVertex (X1 + X2) / 2, (Y1 + Y2) / 2, (Z1 + Z2) / 2
dMesh.AddFace f
Set f = Nothing
End Sub

Public Sub MouseDown(X As Single, Y As Single)
mMouseDown = True
mdownX = X
mDownY = Y
End Sub


Public Sub MouseMove(X As Single, Y As Single)
If Not mMouseDown Then Exit Sub
Rotate X, Y
End Sub

Public Sub MouseUp()
mMouseDown = False
StopRotate
End Sub


Public Function Pick(X As Single, Y As Single) As String
On Error Resume Next
Dim PickArray As Direct3DRMPickArray
Dim Desc As D3DRMPICKDESC
Set PickArray = mVpt.Pick(CLng(X), CLng(Y))
Pick = PickArray.GetPickVisual(0, Desc).GetName
End Function


Public Sub PointToMouse(ByRef dX As Single, ByRef dY As Single, ByRef Distance As Single, X As Single, Y As Single)
Dim sX As Single
Dim sY As Single
dX = mdownX - X
dY = mDownY - Y
sX = (dX * dX)
sY = (dY * dY)
Distance = Sqr(sX + sY)
End Sub

Sub Resize(Canvas As Object)
If Not IsInit Then Exit Sub
Set mDev = mDrm.CreateDeviceFromClipper(DXClipper, "", Canvas.ScaleWidth, Canvas.ScaleHeight)
'Set mDev = mDrm.CreateDeviceFromSurface("IID_IDirect3DRGBDevice", mDD4, mBackBuffer, 0)
Set mVpt = mDrm.CreateViewport(mDev, mFrC, 0, 0, Canvas.ScaleWidth, Canvas.ScaleHeight)
End Sub

Public Sub Rotate(X As Single, Y As Single)
Dim dX As Single
Dim dY As Single
Dim Distance As Single
Dim Theta As Single
PointToMouse dX, dY, Distance, X, Y
Theta = Distance / 1000
mFrO.SetRotation Nothing, dY, dX, 0, Theta
'mFrO.AddRotation D3DRMCOMBINE_BEFORE, dY, dX, 0, Theta
End Sub


Public Sub StopRotate()
mFrO.SetRotation Nothing, 0, 0, 0, 0
End Sub

Public Sub Update()
mFrs.Move 1
mVpt.Clear D3DRMCLEAR_ALL
mVpt.Render mFrs
mDev.Update
DoEvents
End Sub


