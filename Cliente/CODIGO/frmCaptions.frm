VERSION 5.00
Begin VB.Form frmCaptions 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Tierras Perdidas AO"
   ClientHeight    =   3660
   ClientLeft      =   165
   ClientTop       =   525
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   ScaleHeight     =   3660
   ScaleWidth      =   8280
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Matar Proceso"
      Height          =   495
      Left            =   4320
      TabIndex        =   6
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   ">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>"
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   2640
      Width           =   3975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   3975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Foto"
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   3120
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3120
      Width           =   3975
   End
   Begin VB.ListBox Captions 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   2370
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   8055
   End
   Begin VB.ListBox Procesos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   2370
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   8055
   End
End
Attribute VB_Name = "frmCaptions"
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
Private Sub Command1_Click()
Unload Me
End Sub
Private Sub Command3_Click()
Procesos.Visible = True
Captions.Visible = False
Command4.Enabled = True
Command3.Enabled = False
End Sub
Private Sub Command4_Click()
Procesos.Visible = False
Captions.Visible = True
Command4.Enabled = False
Command3.Enabled = True
End Sub
