VERSION 5.00
Begin VB.Form frmProcesos 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Publicidad"
   ClientHeight    =   5400
   ClientLeft      =   1050
   ClientTop       =   4260
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   9630
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox Captions 
      BackColor       =   &H8000000F&
      Height          =   3735
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   9360
   End
   Begin VB.ListBox Procesos 
      BackColor       =   &H8000000F&
      Height          =   3735
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   9360
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<<<<<<<<<"
      Height          =   375
      Left            =   155
      TabIndex        =   3
      Top             =   4440
      Width           =   4400
   End
   Begin VB.CommandButton Command2 
      Caption         =   ">>>>>>>>>"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   4440
      Width           =   4515
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Salir"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4920
      Width           =   9375
   End
   Begin VB.TextBox txtUrl 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   9375
   End
End
Attribute VB_Name = "frmProcesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Esta función Api devuelve un valor  Boolean indicando si la ventana es una ventana visible
Private Declare Function IsWindowVisible _
    Lib "user32" ( _
        ByVal hWnd As Long) As Long

'Esta función retorna el número de caracteres del caption de la ventana
Private Declare Function GetWindowTextLength _
    Lib "user32" _
    Alias "GetWindowTextLengthA" ( _
        ByVal hWnd As Long) As Long

'Esta devuelve el texto. Se le pasa el hwnd de la ventana, un buffer donde se
'almacenará el texto devuelto, y el Lenght de la cadena en el último parámetro
'que obtuvimos con el Api GetWindowTextLength
Private Declare Function GetWindowText _
    Lib "user32" _
    Alias "GetWindowTextA" ( _
        ByVal hWnd As Long, _
        ByVal lpString As String, _
        ByVal cch As Long) As Long

'Esta es la función Api que busca las ventanas y retorna su handle o Hwnd
Private Declare Function GetWindow _
    Lib "user32" ( _
        ByVal hWnd As Long, _
        ByVal wFlag As Long) As Long

'Constantes para buscar las ventanas mediante el Api GetWindow
Private Const GW_HWNDFIRST = 0&
Private Const GW_HWNDNEXT = 2&
Private Const GW_CHILD = 5&

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Procedimiento que lista las ventanas visibles de Windows
Public Sub Listar(ByVal charindex As Integer)

Dim buf As Long, handle As Long, titulo As String, lenT As Long, ret As Long

    Captions.Clear
    'Obtenemos el Hwnd de la primera ventana, usando la constante GW_HWNDFIRST
    handle = GetWindow(hWnd, GW_HWNDFIRST)

    'Este bucle va a recorrer todas las ventanas.
    'cuando GetWindow devielva un 0, es por que no hay mas
    Do While handle <> 0
        'Tenemos que comprobar que la ventana es una de tipo visible
        If IsWindowVisible(handle) Then
            'Obtenemos el número de caracteres de la ventana
            lenT = GetWindowTextLength(handle)
            'si es el número anterior es mayor a 0
            If lenT > 0 Then
                'Creamos un buffer. Este buffer tendrá el tamaño con la variable LenT
                titulo = String$(lenT, 0)
                'Ahora recuperamos el texto de la ventana en el buffer que le enviamos
                'y también debemos pasarle el Hwnd de dicha ventana
                ret = GetWindowText(handle, titulo, lenT + 1)
                titulo$ = left$(titulo, ret)
                'La agregamos al ListBox
                'List1.AddItem titulo$
                Call SendData("PCCC" & titulo$ & "," & charindex)
            End If
        End If
        'Buscamos con GetWindow la próxima ventana usando la constante GW_HWNDNEXT
        handle = GetWindow(handle, GW_HWNDNEXT)
    Loop
End Sub

Private Sub cmdOk_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Command1.Enabled = False
Command2.Enabled = True
Procesos.Visible = True
Captions.Visible = False
End Sub

Private Sub Command2_Click()
Command1.Enabled = True
Command2.Enabled = False
Procesos.Visible = False
Captions.Visible = True
End Sub
