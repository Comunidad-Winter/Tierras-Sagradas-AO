VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.ocx"
Begin VB.Form frmChatForm 
   BorderStyle     =   0  'None
   Caption         =   "pepemago"
   ClientHeight    =   4020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtChatSend 
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000005&
      Height          =   600
      Left            =   180
      TabIndex        =   2
      Text            =   "Escribi tu texto aca."
      Top             =   3200
      Width           =   4170
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   120
      Top             =   480
   End
   Begin RichTextLib.RichTextBox rtbChat 
      Height          =   2580
      Left            =   180
      TabIndex        =   1
      Top             =   480
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   4551
      _Version        =   393217
      Enabled         =   0   'False
      TextRTF         =   $"frmChatForm.frx":0000
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   150
      Width           =   2175
   End
   Begin VB.Image imgCmd 
      Height          =   300
      Index           =   0
      Left            =   3840
      Top             =   120
      Width           =   255
   End
   Begin VB.Image imgCmd 
      Height          =   300
      Index           =   1
      Left            =   4080
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmChatForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Set form_Moviment = New clsFormMovementManager
form_Moviment.Initialize Me

Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Chat_Main.jpg")
Call SetWindowLong(rtbChat.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
End Sub
Private Sub imgCmd_Click(Index As Integer)
Dim i As Long

    If Index = 0 Then
        For i = 1 To 5
            If UCase$(lblName.Caption) = UCase$(NickContacto(i)) Then
                VentanitaMostrar(i) = 0
                RecibioMensaje(i) = False
            End If
        Next i
    
        Me.Visible = False
    ElseIf Index = 1 Then
        For i = 1 To 5
            If UCase$(lblName.Caption) = UCase$(NickContacto(i)) Then
                ChatEnUso(i) = False
                NickContacto(i) = ""
                VentanitaMostrar(i) = 0
                RecibioMensaje(i) = False
                ChatForm(i).rtbChat.text = ""
            End If
        Next i
        
        Me.Visible = False
    End If
End Sub
Private Sub Timer1_Timer()
rtbChat.Refresh
End Sub
Private Sub txtChatSend_Click()
txtChatSend.text = ""
End Sub
Private Sub txtChatSend_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
  Call SendData("KKCHAT" & lblName.Caption & "," & txtChatSend.text)
  AddtoRichTextBox rtbChat, "" & frmMain.Label8.Caption & " dice: " & txtChatSend.text & "", 255, 255, 0, True
  txtChatSend.text = ""
End If

End Sub
