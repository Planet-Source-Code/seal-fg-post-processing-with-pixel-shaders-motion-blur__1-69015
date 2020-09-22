Attribute VB_Name = "mdlMain"

Option Explicit


Public Const Pi As Single = 3.14159265358979

Public confDevice As D3DPRESENT_PARAMETERS

Private objDX As DirectX8
Private objD3D As Direct3D8

Public objD3DDev As Direct3DDevice8
Public objD3Dhlp As D3DX8

Private txHelp As Direct3DTexture8

Private mhWalls As clsMesh
Private txWalls As Direct3DTexture8

Private mhStatue As clsMesh
Private txStatue As Direct3DTexture8

Public camAlpha As Single
Public camBeta As Single
Public camDistance As Single
Public camShift As Single

Public psTextureFade As Long

Private rtOriginalImage As clsRenderTarget
Private rtAccumulator As clsRenderTarget
Private rtTemp As clsRenderTarget

Public ppScreenQuad As clsPostProcessing
Private ppHelpRect As clsPostProcessing

Public effFilter As Long
Public effMotion As Long
Public effUV As Long

Public shwHelp As Long
Public shwWalls As Long

Public Sub Initialize()

  On Error Resume Next

  Set objDX = New DirectX8
  Set objD3D = objDX.Direct3DCreate
  Set objD3Dhlp = New D3DX8
  
  Static confDisplay As D3DDISPLAYMODE
  objD3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, confDisplay
  
  With confDevice
    .AutoDepthStencilFormat = D3DFMT_D24S8
    .BackBufferCount = 1
    .BackBufferFormat = confDisplay.Format
    .BackBufferHeight = wndRender.ScaleHeight
    .BackBufferWidth = wndRender.ScaleWidth
    .EnableAutoDepthStencil = 1
    .flags = 0
    .FullScreen_PresentationInterval = 0
    .FullScreen_RefreshRateInHz = 0
    .hDeviceWindow = wndRender.hWnd
    .MultiSampleType = D3DMULTISAMPLE_NONE
    .SwapEffect = D3DSWAPEFFECT_DISCARD
    .Windowed = 1
  End With

  Set objD3DDev = objD3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, confDevice.hDeviceWindow, D3DCREATE_HARDWARE_VERTEXPROCESSING, confDevice)
  If Not Err.Number = 0 Then
    Err.Clear
    MsgBox "Failed to create Direct3DDevice8. Application will now quit.", vbCritical Or vbOKOnly, "Error"
    Shutdown
  End If

  
  camDistance = 150
  camAlpha = 35 * Pi / 180
  camBeta = 15 * Pi / 180
  camShift = 50
  
  
  effFilter = 1
  effMotion = 1
  effUV = 1
  
  shwHelp = 1
  shwWalls = 1
  
  
  psTextureFade = shCompile(App.Path & "\psh_TextureFade_ps.1.1.txt")


  Set txHelp = txLoad(App.Path & "\texHelp.png")
  Set txWalls = txLoad(App.Path & "\texWalls.png")
  Set txStatue = txLoad(App.Path & "\texStatue.png")


  Set mhWalls = New clsMesh
  If Not mhWalls.objLoad(App.Path & "\objWalls.obj") Then
    mhWalls.memClear
    MsgBox "Failed to load mesh file: '" & App.Path & "\objWalls.obj" & "'.", vbCritical Or vbOKOnly, "Error"
  End If

  Set mhStatue = New clsMesh
  If Not mhStatue.objLoad(App.Path & "\objStatue.obj") Then
    mhStatue.memClear
    MsgBox "Failed to load mesh file: '" & App.Path & "\objStatue.obj" & "'.", vbCritical Or vbOKOnly, "Error"
  End If


  Set rtOriginalImage = New clsRenderTarget
  rtOriginalImage.rtAquire
  rtOriginalImage.rtCreate confDevice.BackBufferWidth, confDevice.BackBufferHeight
  
  Set rtAccumulator = New clsRenderTarget
  rtAccumulator.rtAquire
  rtAccumulator.rtCreate confDevice.BackBufferWidth, confDevice.BackBufferWidth

  Set rtTemp = New clsRenderTarget
  rtTemp.rtAquire
  rtTemp.rtCreate confDevice.BackBufferWidth, confDevice.BackBufferWidth


  Set ppHelpRect = New clsPostProcessing
  
  Set ppScreenQuad = New clsPostProcessing
  ppScreenQuad.objCreate5Tap confDevice.BackBufferWidth, confDevice.BackBufferWidth

End Sub


Public Sub Render()

  On Error Resume Next

  With objD3DDev
    
    
    If effMotion = 1 Then
      
      'pass0: render scene into RT texture
      rtOriginalImage.rtEnable True
    
    End If
    
    
    .Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, &HFF3F3F3F, 1, 0
    .BeginScene
    
    
    Static camX As Single
    Static camY As Single
    Static camZ As Single
    camX = Sin(camAlpha) * Cos(camBeta) * camDistance
    camY = Sin(camBeta) * camDistance
    camZ = Cos(camAlpha) * Cos(camBeta) * camDistance
    
    
    Static matView As D3DMATRIX
    D3DXMatrixLookAtLH matView, mkVec3f(camX, camY + camShift, camZ), mkVec3f(0, 0 + camShift, 0), mkVec3f(0, 1, 0)
    .SetTransform D3DTS_VIEW, matView
    
    Static matProjection As D3DMATRIX
    D3DXMatrixPerspectiveFovLH matProjection, 1, confDevice.BackBufferHeight / confDevice.BackBufferWidth, 1, 100000
    .SetTransform D3DTS_PROJECTION, matProjection
  
    
    
    Static iMap As Long
    For iMap = 0 To 1 Step 1
      .SetTextureStageState iMap, D3DTSS_TEXCOORDINDEX, iMap
      If effFilter = 1 Then
        .SetTextureStageState iMap, D3DTSS_MAXANISOTROPY, 16
        .SetTextureStageState iMap, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
        .SetTextureStageState iMap, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC
        .SetTextureStageState iMap, D3DTSS_MIPFILTER, D3DTEXF_ANISOTROPIC
      Else
        .SetTextureStageState iMap, D3DTSS_MAGFILTER, D3DTEXF_POINT
        .SetTextureStageState iMap, D3DTSS_MINFILTER, D3DTEXF_POINT
        .SetTextureStageState iMap, D3DTSS_MIPFILTER, D3DTEXF_POINT
      End If
      .SetTextureStageState iMap, D3DTSS_ADDRESSU, D3DTADDRESS_CLAMP
      .SetTextureStageState iMap, D3DTSS_ADDRESSV, D3DTADDRESS_CLAMP
    Next iMap
    
    
    .SetRenderState D3DRS_LIGHTING, 0
    .SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
    
    .SetPixelShader 0
  
  
    If shwWalls = 1 Then
      .SetTexture 0, txWalls
      If Not mhWalls.objRender Then
        MsgBox "Error rendering 'Walls' mesh. Application will now quit.", vbCritical Or vbOKOnly, "Error"
        Shutdown
      End If
    End If
  
    .SetTexture 0, txStatue
    If Not mhStatue.objRender Then
      MsgBox "Error rendering 'Statue' mesh. Application will now quit.", vbCritical Or vbOKOnly, "Error"
      Shutdown
    End If
  
    
      
    If shwHelp = 1 Then
      ppHelpRect.memClear
      ppHelpRect.objCreateUser 0.3, -0.5, 1, -1
      .SetPixelShader 0
      .SetTexture 0, txHelp
      .SetRenderState D3DRS_ALPHABLENDENABLE, 1
      .SetRenderState D3DRS_SRCBLEND, 1
      .SetRenderState D3DRS_DESTBLEND, 3
      ppHelpRect.objRender
      .SetRenderState D3DRS_ALPHABLENDENABLE, 0
    End If
  
    .EndScene
    
    
    
    If effMotion = 1 Then
    
      rtOriginalImage.rtEnable False
      
      'pass1: add rendered scene to accumulator using pixel shader (render to temp RT texture)
      rtTemp.rtEnable True
        .SetTexture 0, rtOriginalImage.objTexture
        .SetTexture 1, rtAccumulator.objTexture
        .SetPixelShader psTextureFade
        ppScreenQuad.objRender
        .SetTexture 1, Nothing
      rtTemp.rtEnable False
      
      rtAccumulator.rtEnable True
        .SetTexture 0, rtTemp.objTexture
        .SetPixelShader 0
        ppScreenQuad.objRender
      rtAccumulator.rtEnable False
      
      
      'pass2: show the accumulator
      .SetTexture 0, rtAccumulator.objTexture
      ppScreenQuad.objRender
    
    End If
    
    
    
    .SetTexture 0, Nothing
    
    
    If Not .TestCooperativeLevel = 0 Then
      MsgBox "Cooperative level lost. Application will now quit.", vbCritical Or vbOKOnly, "Error"
      Shutdown
    Else
      .Present ByVal 0, ByVal 0, 0, ByVal 0
    End If
  End With


  If Not Err.Number = 0 Then
    Err.Clear
    MsgBox "Unexpected error occured in rendering pipeline. Application will now quit.", vbCritical Or vbOKOnly, "Error"
    Shutdown
  End If


End Sub


Public Sub Shutdown()

  On Error Resume Next

  Set txHelp = Nothing

  ppScreenQuad.memClear
  Set ppScreenQuad = Nothing

  ppHelpRect.memClear
  Set ppHelpRect = Nothing

  rtOriginalImage.rtDestroy True
  rtTemp.rtDestroy True
  rtAccumulator.rtDestroy True
  
  Set rtTemp = Nothing
  Set rtAccumulator = Nothing
  Set rtOriginalImage = Nothing

  Set txWalls = Nothing

  mhWalls.memClear
  Set mhWalls = Nothing

  Set txStatue = Nothing

  mhStatue.memClear
  Set mhStatue = Nothing
  
  objD3DDev.DeletePixelShader psTextureFade

  Set objD3DDev = Nothing
  Set objD3D = Nothing
  Set objDX = Nothing
  
  If Not Err.Number = 0 Then Err.Clear
  End

End Sub

