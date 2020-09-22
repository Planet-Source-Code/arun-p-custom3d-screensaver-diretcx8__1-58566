Attribute VB_Name = "modgens"
'sorry, i could not have time for detailed explantions ........
'its my first, ya i'm a beginner too.
'please feel free to contact me at arun_pbk@rediffmail.com
'press F8 to find the execution order
'

Public blnregcre As Boolean
''''''''''''''

Public Sub SetOnTop(hWnd As Long, SetOnTop As Boolean)
    If SetOnTop Then
        lFlag = HWND_TOPMOST
    Else
        lFlag = HWND_NOTOPMOST
    End If
    SetWindowPos hWnd, lFlag, 0, 0, 0, 0, 3
End Sub


Public Function GetBMPs() As Boolean
    Dim strbuf As String
    strbuf = Registry.QueryValue(HKEY_LOCAL_MACHINE, "software\Aruns\CstScr", "isReady")
    blnregcre = False
    If strbuf = "" Then
        blnregcre = True
        Form1.Show
        GetBMPs = True
        Exit Function
    End If
    'registry fns
    strbuf = Registry.QueryValue(HKEY_LOCAL_MACHINE, "software\Aruns\CstScr", "Roof")
    Form1.Text1(0).Text = strbuf
    strbuf = Registry.QueryValue(HKEY_LOCAL_MACHINE, "software\Aruns\CstScr", "Floor")
    Form1.Text1(1).Text = strbuf
    strbuf = Registry.QueryValue(HKEY_LOCAL_MACHINE, "software\Aruns\CstScr", "wall1")
    Form1.Text1(2).Text = strbuf
    strbuf = Registry.QueryValue(HKEY_LOCAL_MACHINE, "software\Aruns\CstScr", "wall2")
    Form1.Text1(3).Text = strbuf
    strbuf = Registry.QueryValue(HKEY_LOCAL_MACHINE, "software\Aruns\CstScr", "wall3")
    Form1.Text1(4).Text = strbuf
    strbuf = Registry.QueryValue(HKEY_LOCAL_MACHINE, "software\Aruns\CstScr", "wall4")
    Form1.Text1(5).Text = strbuf
    
End Function
