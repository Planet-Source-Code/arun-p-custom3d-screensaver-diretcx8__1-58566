VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Settings"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7590
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cmndlg1 
      Left            =   4200
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "OK"
      Height          =   855
      Left            =   6360
      TabIndex        =   22
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   855
      Left            =   5160
      TabIndex        =   21
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse"
      Height          =   255
      Index           =   5
      Left            =   6360
      TabIndex        =   14
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse"
      Height          =   255
      Index           =   4
      Left            =   6360
      TabIndex        =   13
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse"
      Height          =   255
      Index           =   3
      Left            =   6360
      TabIndex        =   12
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse"
      Height          =   255
      Index           =   2
      Left            =   6360
      TabIndex        =   11
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse"
      Height          =   255
      Index           =   1
      Left            =   6360
      TabIndex        =   10
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse"
      Height          =   255
      Index           =   0
      Left            =   6360
      TabIndex        =   9
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   5
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   2160
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   4
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1800
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1440
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1080
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   720
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   5535
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Wall4"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   20
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Wall3"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   19
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Wall2"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   18
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Wall1"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   17
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Floor"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Roof"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Free Ware"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "arun_pbk@rediffmail.com"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Screen Saver By Arun P"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   2520
      Width           =   5175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click(Index As Integer)
    On Error GoTo la
    cmndlg1.CancelError = True
    cmndlg1.Filter = "Bitmaps (*.bmp)|*.bmp"
    cmndlg1.DialogTitle = "Get BMP"
    cmndlg1.ShowOpen
    Text1(Index).Text = cmndlg1.FileName
la:
End Sub

Private Sub Command3_Click()
If blnregcre Then
    End
Else
    frmMain.Show
    Form1.Hide
End If
End Sub

Private Sub Command4_Click()
    For Each Control In Me
        If TypeOf Control Is TextBox Then
            If Control.Text = "" Then
                MsgBox "Incomplete Info", vbInformation, "Custom 3d room"
                Exit Sub
            End If
        End If
    Next
    
    If blnregcre Then
        Registry.CreateNewKey "software\Aruns\cstscr", HKEY_LOCAL_MACHINE
        Registry.SetKeyValue HKEY_LOCAL_MACHINE, "software\Aruns\cstscr", "isready", "yaa", REG_SZ
    End If
    Registry.SetKeyValue HKEY_LOCAL_MACHINE, "software\Aruns\cstscr", "roof", Text1(0).Text, REG_SZ
    Registry.SetKeyValue HKEY_LOCAL_MACHINE, "software\Aruns\cstscr", "floor", Text1(1).Text, REG_SZ
    Registry.SetKeyValue HKEY_LOCAL_MACHINE, "software\Aruns\cstscr", "wall1", Text1(2).Text, REG_SZ
    Registry.SetKeyValue HKEY_LOCAL_MACHINE, "software\Aruns\cstscr", "wall2", Text1(3).Text, REG_SZ
    Registry.SetKeyValue HKEY_LOCAL_MACHINE, "software\Aruns\cstscr", "wall3", Text1(4).Text, REG_SZ
    Registry.SetKeyValue HKEY_LOCAL_MACHINE, "software\Aruns\cstscr", "wall4", Text1(5).Text, REG_SZ
    frmMain.Show
    Form1.Hide
End Sub


