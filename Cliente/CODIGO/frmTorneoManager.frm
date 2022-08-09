VERSION 5.00
Begin VB.Form frmTorneoManager 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Lista de Usuarios"
   ClientHeight    =   4695
   ClientLeft      =   150
   ClientTop       =   315
   ClientWidth     =   3495
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Ass 
      Caption         =   "SUM"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "B. TODO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "BORRAR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ECHAR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "DV"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "IR A"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4350
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Menu menU_usuario 
      Caption         =   "Usuario"
      Visible         =   0   'False
      Begin VB.Menu mnuIR 
         Caption         =   "Ir donde esta el usuario"
      End
      Begin VB.Menu mnutraer 
         Caption         =   "Traer usuario"
      End
      Begin VB.Menu mnuBorrar 
         Caption         =   "Borrar mensaje"
      End
      Begin VB.Menu mnudeleteall 
         Caption         =   "Borrar Todos"
      End
   End
End
Attribute VB_Name = "frmTorneoManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub PonerListaTorneo(ByVal Rdata As String)

Dim j As Integer, k As Integer
For j = 0 To List1.ListCount - 1
    Me.List1.RemoveItem 0
Next j
k = CInt(ReadField(1, Rdata, 44))

For j = 1 To k
    List1.AddItem ReadField(1 + j, Rdata, 44)
Next j

Me.Show , frmMain

End Sub
Private Sub Ass_Click()
Call SendData("/SUM " & List1.Text & "")
End Sub
Private Sub Command2_Click()
Call SendData("/IRA " & List1.Text & "")
End Sub
Private Sub Command3_Click()
Call SendData("/DV " & List1.Text & "")
End Sub
Private Sub Command4_Click()
Call SendData("/ECHAR " & List1.Text & "")
End Sub
Private Sub Command5_Click()
If List1.Text = "" Then Exit Sub
If List1.listIndex = 0 Then Exit Sub
List1.RemoveItem List1.Text
End Sub
Private Sub Command6_Click()
List1.Clear
End Sub
