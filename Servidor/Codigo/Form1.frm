VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administración de Usuarios"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   3720
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      BackColor       =   &H00C0FFFF&
      Height          =   3960
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   4800
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Recargar"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Menu cmdOpc 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu cmdBan 
         Caption         =   "Banear"
      End
      Begin VB.Menu cmdEchar 
         Caption         =   "Echar"
      End
      Begin VB.Menu cmdStop 
         Caption         =   "Stopear"
      End
      Begin VB.Menu cmdHome 
         Caption         =   "Mandar a Tanaris"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

List1.Clear

Dim i As Long
For i = 1 To LastUser
List1.AddItem "" & i & ". " & UserList(i).Name & ""
Next i
End Sub
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Form_Load()

List1.Clear

Dim i As Long
For i = 1 To LastUser
List1.AddItem "" & i & ". " & UserList(i).Name & ""
Next i

End Sub
Private Sub cmdBan_Click()
    Call WriteVar(CharPath & UserList(List1.ListIndex + 1).Name & ".chr", "FLAGS", "Ban", "1")
    Call CloseSocket(List1.ListIndex + 1)
End Sub
Private Sub cmdEchar_Click()
    Call CloseSocket(List1.ListIndex + 1)
End Sub
Private Sub cmdHome_Click()
    Call WarpUserChar(List1.ListIndex + 1, 1, 54, 36, True)
End Sub
Private Sub cmdStop_Click()
    If UserList(List1.ListIndex + 1).flags.Stopped = 1 Then
        UserList(List1.ListIndex + 1).flags.Stopped = 0
        MsgBox "Usuario REMOVIDO"
    Else
        UserList(List1.ListIndex + 1).flags.Stopped = 1
        MsgBox "Usuario STOP"
    End If
End Sub
Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
PopupMenu cmdOpc
End If

End Sub

