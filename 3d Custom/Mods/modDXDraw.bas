Attribute VB_Name = "modDXDraw"
'sorry, i could not have time for detailed explantions ........
'its my first, ya i'm a beginner too.
'please feel free to contact me at arun_pbk@rediffmail.com
'press F8 to find the execution order
'

Option Explicit
'Direct X Variables
Public dx8         As New DirectX8 'main
Public dx3D        As Direct3D8

Public d3dDevice   As Direct3DDevice8 '3d
Public isWin As Boolean

Public blnSD As Boolean

Public Function InitDirectX(hFocWind As Long) As Boolean
    On Error GoTo InitFails
    InitDirectX = True
    
    Set dx3D = dx8.Direct3DCreate()
    Dim dx3dpp As D3DPRESENT_PARAMETERS
    Dim omode As D3DDISPLAYMODE
    
    dx3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, omode
    
    'dx3dpp.AutoDepthStencilFormat = D3DFMT_R8G8B8
    dx3dpp.BackBufferCount = 1
    dx3dpp.BackBufferWidth = omode.Width
    dx3dpp.BackBufferHeight = omode.Height
    dx3dpp.BackBufferFormat = omode.Format  'D3DFMT_R5G6B5
    
    'dx3dpp.flags = 0
    'dx3dpp.FullScreen_PresentationInterval = D3DPRESENT_INTERVAL_DEFAULT
    'dx3dpp.FullScreen_RefreshRateInHz = D3DPRESENT_RATE_DEFAULT
    dx3dpp.hDeviceWindow = hFocWind
    'dx3dpp.MultiSampleType = D3DMULTISAMPLE_NONE
    dx3dpp.SwapEffect = D3DSWAPEFFECT_DISCARD
    dx3dpp.Windowed = isWin
    
    dx3dpp.EnableAutoDepthStencil = True
    dx3dpp.AutoDepthStencilFormat = D3DFMT_D16
    
    Set d3dDevice = dx3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hFocWind, D3DCREATE_SOFTWARE_VERTEXPROCESSING, dx3dpp)
    '''
    
    Exit Function
InitFails:
    InitDirectX = False
    MsgBox Err.Number & " " & Err.Description
End Function



Public Function DestroyDirectX()
    
    If Not d3dDevice Is Nothing Then
        Set d3dDevice = Nothing
    End If
    
    If Not dx3D Is Nothing Then
        Set dx3D = Nothing
    End If

End Function


