Attribute VB_Name = "modDXRender"
Option Explicit

Private Const D3D8T_VERTEX = D3DFVF_XYZ Or D3DFVF_TEX1

Private Type d3dMyVertex
    X     As Single
    Y     As Single
    z     As Single
    tu    As Single
    tv    As Single
End Type

Dim d3T As Direct3DVertexBuffer8
Dim d3Flo As Direct3DVertexBuffer8

Dim dxTextre As Direct3DTexture8
Dim dxflr As Direct3DTexture8
Dim dxWall1 As Direct3DTexture8
Dim dxwall2 As Direct3DTexture8
Dim dxwall3 As Direct3DTexture8
Dim dxwall4 As Direct3DTexture8
Const PI = 3.14152678

Dim V(23) As d3dMyVertex


Dim D3DX As New D3DX8

Public ex As Single, ey As Single, ez As Single
Public el As Single, em As Single, en As Single

    Dim oViewMatrix As D3DMATRIX
    Dim oMatProj As D3DMATRIX
    Dim oEyeVector As D3DVECTOR
    Dim oLookAtVector As D3DVECTOR
    Dim oUpVector As D3DVECTOR
'render loooooop
Public Function Render()
    Dim mxWorld As D3DMATRIX
    Dim mxRot As D3DMATRIX
    Dim mxTran As D3DMATRIX
    
    Static ma As Single
    
    InitMatrices ex, ey, ez, el, em, en
    With d3dDevice
        .Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, _
                D3DColorXRGB(0, 0, 0), 1!, 0
        
        .BeginScene
        .SetVertexShader D3D8T_VERTEX
        .SetStreamSource 0, d3T, Len(V(0))
        
        .SetTexture 0, dxflr
        .DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
        
        .SetTexture 0, dxWall1
        .DrawPrimitive D3DPT_TRIANGLESTRIP, 4, 2
        
        .SetTexture 0, dxwall2
        .DrawPrimitive D3DPT_TRIANGLESTRIP, 8, 2
        
        .SetTexture 0, dxwall3
        .DrawPrimitive D3DPT_TRIANGLESTRIP, 12, 2
        
        .SetTexture 0, dxwall4
        .DrawPrimitive D3DPT_TRIANGLESTRIP, 16, 2
        
        
        .SetTexture 0, dxTextre
        .DrawPrimitive D3DPT_TRIANGLESTRIP, 20, 2
         
         
        .EndScene
        .SetTexture 0, Nothing
        .Present ByVal 0, ByVal 0, 0, ByVal 0
    End With
End Function

'setting the vertices......
Public Function InitVertex()

    With V(0)
        .X = -8!: .Y = -1.01!: .z = -8!: .tu = 1!: .tv = 1!
    End With
    
    With V(1)
        .X = -8!: .Y = -1.01!: .z = 8!: .tu = 0!: .tv = 1!
    End With
    
    With V(2)
        .X = 8!: .Y = -1.01!: .z = -8!: .tu = 1!: .tv = 0!
    End With
    
    With V(3)
        .X = 8!: .Y = -1.01!: .z = 8!: .tu = 0!: .tv = 0!
    End With
    
    With V(4)
        .X = -8: .Y = -1: .z = 8: .tu = 0: .tv = 1
    End With
    
    With V(5)
        .X = -8: .Y = 10: .z = 8: .tu = 0: .tv = 0
    End With
    
    With V(6)
        .X = 8: .Y = -1: .z = 8: .tu = 1: .tv = 1
    End With
    
    With V(7)
        .X = 8: .Y = 10: .z = 8: .tu = 1: .tv = 0
    End With
    
    With V(8)
        .X = 8: .Y = -1: .z = -8: .tu = 1: .tv = 1
    End With
    
    With V(9)
        .X = 8: .Y = 10: .z = -8: .tu = 1: .tv = 0
    End With
    
    With V(10)
        .X = 8: .Y = -1: .z = 8: .tu = 0: .tv = 1
    End With
    
    With V(11)
        .X = 8: .Y = 10: .z = 8: .tu = 0: .tv = 0
    End With
    
    With V(12)
        .X = -8: .Y = -1: .z = -8: .tu = 1: .tv = 1
    End With
    
    With V(13)
        .X = -8: .Y = 10: .z = -8: .tu = 1: .tv = 0
    End With
    
    With V(14)
        .X = -8: .Y = -1: .z = 8: .tu = 0: .tv = 1
    End With
    
    With V(15)
        .X = -8: .Y = 10: .z = 8: .tu = 0: .tv = 0
    End With
    
    With V(16)
        .X = -8: .Y = -1: .z = -8: .tu = 1: .tv = 1
    End With
    
    With V(17)
        .X = -8: .Y = 10: .z = -8: .tu = 1: .tv = 0
    End With
    
    With V(18)
        .X = 8: .Y = -1: .z = -8: .tu = 0: .tv = 1
    End With
    
    With V(19)
        .X = 8: .Y = 10: .z = -8: .tu = 0: .tv = 0
    End With
    
    With V(20)
        .X = -8: .Y = 10: .z = 8: .tu = 0: .tv = 1
    End With
    
    With V(21)
        .X = -8: .Y = 10: .z = -8: .tu = 0: .tv = 0
    End With
    
    With V(22)
        .X = 8: .Y = 10: .z = 8: .tu = 1: .tv = 1
    End With
    
    With V(23)
        .X = 8: .Y = 10: .z = -8: .tu = 1: .tv = 0
    End With
    
End Function

Public Function InitScene()
    d3dDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
    d3dDevice.SetRenderState D3DRS_LIGHTING, False
    d3dDevice.SetRenderState D3DRS_ZENABLE, D3DZB_TRUE

    d3dDevice.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_SELECTARG1
    d3dDevice.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
    d3dDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
    d3dDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
    d3dDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
    
    Set dxTextre = D3DX.CreateTextureFromFile(d3dDevice, Form1.Text1(0).Text)
    Set dxflr = D3DX.CreateTextureFromFile(d3dDevice, Form1.Text1(1).Text)
    Set dxWall1 = D3DX.CreateTextureFromFile(d3dDevice, Form1.Text1(2).Text)
    Set dxwall2 = D3DX.CreateTextureFromFile(d3dDevice, Form1.Text1(3).Text)
    Set dxwall3 = D3DX.CreateTextureFromFile(d3dDevice, Form1.Text1(4).Text)
    Set dxwall4 = D3DX.CreateTextureFromFile(d3dDevice, Form1.Text1(5).Text)
    
    
        
    Set d3T = d3dDevice.CreateVertexBuffer(24 * Len(V(0)), D3DUSAGE_WRITEONLY, D3D8T_VERTEX, D3DPOOL_MANAGED)
    Debug.Print D3DVertexBuffer8SetData(d3T, 0, Len(V(0)) * 24, 0, V(0))
    
    InitMatrices
End Function

Public Function DestroyVertex()
    
    If Not d3T Is Nothing Then
        Set d3T = Nothing
    End If
    
End Function

Public Sub InitMatrices(Optional X As Single = 0, Optional Y As Single = 0, Optional z As Single = -8, Optional alp As Single = 0, Optional bet As Single = 0, Optional gam As Single = 0)
    
    oEyeVector.X = X!
    oEyeVector.Y = Y!
    oEyeVector.z = z!
    
    'We are looking towards the origin
    oLookAtVector.X = X!
    oLookAtVector.Y = Y!
    oLookAtVector.z = z + 8!
    
    
    'The "up" direction is the positive direction on the Y-axis
    oUpVector.X = 0!
    oUpVector.Y = 1!
    oUpVector.z = 0!
    
    
    D3DXMatrixLookAtLH oViewMatrix, oEyeVector, oLookAtVector, oUpVector

    Dim xrot As D3DMATRIX
    Dim yrot As D3DMATRIX
    Dim zrot As D3DMATRIX
    
    D3DXMatrixRotationX xrot, alp
    D3DXMatrixRotationY yrot, bet
    D3DXMatrixRotationZ zrot, gam

    D3DXMatrixMultiply oViewMatrix, oViewMatrix, xrot
    D3DXMatrixMultiply oViewMatrix, oViewMatrix, yrot
    D3DXMatrixMultiply oViewMatrix, oViewMatrix, zrot
    
    d3dDevice.SetTransform D3DTS_VIEW, oViewMatrix
    
    D3DXMatrixPerspectiveFovLH oMatProj, 3.14152768 / 4, _
            CSng(768) / CSng(1024), 1!, 200!

    
    d3dDevice.SetTransform D3DTS_PROJECTION, oMatProj

End Sub

'''''''''''''''''''''''
'cmd arguments
'/p : preview
'/c settings
''''''''''''''''''''''''''
Sub Main()
    If Not GetBMPs Then
        Select Case Left(Command, 2)
            Case "/p":
                isWin = True
                frmMain.Top = (Screen.Height - frmMain.Height) / 2
                frmMain.Left = (Screen.Width - frmMain.Width) / 2
                frmMain.Show
            Case "/c":
                Form1.Show
            Case Else
                frmMain.Show
        End Select
    End If
End Sub
