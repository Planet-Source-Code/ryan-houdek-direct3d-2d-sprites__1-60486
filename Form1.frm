VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   3135
      Left            =   0
      ScaleHeight     =   15
      ScaleMode       =   0  'User
      ScaleWidth      =   20
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   2400
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'general d3d stuffs
Dim dx      As DirectX8
Dim d3d     As Direct3D8
Dim dev     As Direct3DDevice8
Dim d3dx     As D3DX8
Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR
Dim texSprite  As Direct3DTexture8   'our sprite
Dim Dir As Long
Dim SpriteX   As Integer    'our sprite coordinates
Dim SpriteY   As Integer
Dim fps As Long
Dim Init As Boolean 'Has main loop been started?
Dim Done As Boolean 'Is program exiting?


Private Sub Form_Activate()

   'After form has been displayed, start main loop (but only do so once)
   If Not Init Then
       Init = True
       Call MainLoop
   End If

End Sub

Private Sub Form_Resize()

Picture1.Width = Me.Width
Picture1.Height = Me.Height
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)

 If KeyCode = vbKeyEscape Then Call Unload(Me)

 If KeyCode = vbKeyLeft Then SpriteX = SpriteX - 5: Dir = 3
 If KeyCode = vbKeyRight Then SpriteX = SpriteX + 5: Dir = 1
 If KeyCode = vbKeyDown Then SpriteY = SpriteY + 5: Dir = 2
 If KeyCode = vbKeyUp Then SpriteY = SpriteY - 5: Dir = 0

End Sub

Private Sub Form_Load()

 Dim picFile     As String
 Dim spriteFile    As String

Dir = 0
 'setup the device, and get us in fullscreen
 InitDX dx, d3d, dev, Picture1.hWnd

 'setup our d3dx library
 Set d3dx = New D3DX8
 'load our surface, and our texture
 spriteFile = App.Path + "\sprite.bmp"

 'should be set to the size and format of the backbuffer
Set texSprite = d3dx.CreateTextureFromFileEx(dev, spriteFile, D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_A1R5G5B5, D3DPOOL_MANAGED, D3DX_FILTER_NONE, D3DX_FILTER_NONE, RGB(255, 0, 255), ByVal 0, ByVal 0)
End Sub


Public Sub InitDX(pDX As DirectX8, pD3D As Direct3D8, pDev As Direct3DDevice8, pHwnd As Long, Optional DevType As CONST_D3DDEVTYPE = D3DDEVTYPE_HAL)

 Dim d3dpp      As D3DPRESENT_PARAMETERS

 Set pDX = New DirectX8
 Set pD3D = dx.Direct3DCreate
Dim Displaymode As D3DDISPLAYMODE
Call pD3D.GetAdapterDisplayMode(D3DADAPTER_DEFAULT, Displaymode)
 With d3dpp
   .Windowed = True    'Graphics appear in window (not full screen)
   .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC  'Synchronize the copying of the backbuffer to the framebuffer with the monitor's vertical refresh rate
   .BackBufferWidth = 800
   .BackBufferHeight = 600
   'Here we set the backbuffer bit depth - can be changed
   'be careful, tho, you must keep this the same as the format
   'of the background image.
   .BackBufferFormat = Displaymode.Format  '16-bit D3DFMT_R5G6B5
 End With

 Set pDev = pD3D.CreateDevice(D3DADAPTER_DEFAULT, DevType, pHwnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, d3dpp)

 With pDev
   .SetVertexShader FVF
   .SetRenderState D3DRS_LIGHTING, 0
   .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
   .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
   .SetRenderState D3DRS_ALPHABLENDENABLE, 1
 End With

End Sub

Public Sub MainLoop()

 Dim spriteVerts(3)    As D3DTLVERTEX
'v1: x=0, y=0____________v3: x=128, y=0
'tu=0, tv=0              tu=1, tv=0
'   |  \                     |
'   |   \                    |
'   |     \                  |
'   |                    \   |
'   |                      \ |
'v0: x=0, y=128__________v2: x=128, y=128
'tu=0, tv=1              tu=1, tv=1


 'get our backbuffer handle!
 Set surfBack = dev.GetRenderTarget

Dim Color As Long
Color = RGB(255, 255, 255)
 'start the loop!
 Do While Not Done
 DoEvents
   'set up our sprite!
 With spriteVerts(0)
   .sx = 0: .sy = 32
   .tu = 0: .tv = 0 + (0.25 * (Dir + 1))
   .rhw = 1
   .Color = Color
 End With
 With spriteVerts(1)
   .sx = 0: .sy = 0
   .tu = 0: .tv = 0 + (0.25 * (Dir))
   .rhw = 1
   .Color = Color
 End With
 With spriteVerts(2)
   .sx = 24: .sy = 32
   .tu = ((24 / 128)): .tv = 0 + (0.25 * (Dir + 1))
   .rhw = 1
   .Color = Color
 End With
 With spriteVerts(3)
   .sx = 24: .sy = 0
   .tu = ((24 / 128)): .tv = 0 + (0.25 * (Dir))
   .rhw = 1
   .Color = Color
 End With
  'Don't try to update screen when window is minimized
  If WindowState = vbNormal Then
  
        'update the sprite position
        spriteVerts(0).sx = SpriteX: spriteVerts(0).sy = SpriteY + 128
        spriteVerts(1).sx = SpriteX: spriteVerts(1).sy = SpriteY
        spriteVerts(2).sx = SpriteX + 128: spriteVerts(2).sy = SpriteY + 128
        spriteVerts(3).sx = SpriteX + 128: spriteVerts(3).sy = SpriteY
       
        'clear the backbuffer!
        dev.Clear 0, ByVal 0&, D3DCLEAR_TARGET, &HFF, 0, 0
       
        'wonder what this does?;)
        dev.BeginScene
       
        'set the sprite texture
        dev.SetTexture 0, texSprite
       
        'draw the sprite!
        dev.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, spriteVerts(0), Len(spriteVerts(0))
       
        'self-explanatory...
        dev.EndScene
       
        'just like good old fashioned Flip
        dev.Present ByVal 0&, ByVal 0&, 0, ByVal 0&
        fps = fps + 1
          DoEvents
   End If
   'can't forget this!
   DoEvents

 Loop
 
End Sub

Private Sub Form_Unload(Cancel As Integer)

   'Stop main loop
   Done = True

End Sub

Private Sub Timer1_Timer()
Me.Caption = fps
fps = 0
End Sub
