VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3585
   LinkTopic       =   "Form1"
   ScaleHeight     =   2310
   ScaleWidth      =   3585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rndX As Single
Dim rndY As Single
Dim rndZ As Single
Dim rndL As Single
Dim rndM As Single
Dim rndN As Single

Private Sub Form_Click()
DestroyVertex
DestroyDirectX
End
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
DestroyVertex
DestroyDirectX
End
End Sub

Private Sub Form_Load()
    InitDirectX (Me.hWnd)
    InitVertex
    InitScene
   ' InitMatrices
   
   Randomize 1
    ex = 0
    ey = 0
    ez = -8
    rndX = Rnd / 10
    rndY = Rnd / 10
    rndZ = Rnd / 10
    
    
    Me.Show
    DoEvents


    
    Timer1.Enabled = True
    DoEvents
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If isWin Then
    End
End If
End Sub

Private Sub Timer1_Timer()
    Render
    
    ex = ex + rndX
    ey = ey + rndY
    ez = ez + rndZ
    'el = el + rndL
    em = em + rndM
    'en = en + rndN
    
    Debug.Print ez
    If Abs(ex) > 7 Then
        rndX = -rndX
        rndM = (0.5 - Rnd) / 60
    End If
    
    If ey > 9 Or ey < -0.5 Then
        rndY = -rndY
        rndM = (0.5 - Rnd) / 60
    End If
    If Abs(ez) > 8 Then
        rndZ = -rndZ
        rndM = (0.5 - Rnd) / 60
    End If
End Sub


