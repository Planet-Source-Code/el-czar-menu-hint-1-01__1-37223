VERSION 5.00
Object = "*\ACZarSoftMenuHint.vbp"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIPrueba 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin CSarSoftMenuHint.CSMenuHint CSMenuHint1 
      Left            =   900
      Top             =   945
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   2865
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Enabled         =   0   'False
            Object.Width           =   5186
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "13:22"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu menua 
      Caption         =   "&Archivo"
      Index           =   1
      Begin VB.Menu menuarchivo 
         Caption         =   "&Nuevo"
         Index           =   1
      End
      Begin VB.Menu menuarchivo 
         Caption         =   "&Abrir"
         Index           =   2
      End
      Begin VB.Menu menuarchivo 
         Caption         =   "&Guardar"
         Index           =   3
      End
      Begin VB.Menu menuarchivo 
         Caption         =   "Guardar Como"
         Index           =   4
      End
      Begin VB.Menu menuarchivo 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu menuarchivo 
         Caption         =   "&Imprimir"
         Index           =   6
      End
      Begin VB.Menu menuarchivo 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu menuarchivo 
         Caption         =   "&Salir"
         Index           =   8
      End
   End
   Begin VB.Menu menua 
      Caption         =   "&Edicion"
      Index           =   2
      Begin VB.Menu menuedicion 
         Caption         =   "Copiar"
         Index           =   1
      End
      Begin VB.Menu menuedicion 
         Caption         =   "Cortar"
         Index           =   2
      End
      Begin VB.Menu menuedicion 
         Caption         =   "Pegar"
         Index           =   3
      End
   End
   Begin VB.Menu menua 
      Caption         =   "Acerca de"
      Index           =   3
   End
End
Attribute VB_Name = "MDIPrueba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CSMenuHint1_MostrarMensaje(Mensaje As String)
Dim MensajeSTR      As String
Select Case Mensaje
    Case "&Archivo"
        MensajeSTR = "Menu Archivo"
    Case "&Nuevo"
        MensajeSTR = "Nuevo Documento"
    Case "&Abrir"
        MensajeSTR = "Abrir Documento"
    Case "&Guardar"
        MensajeSTR = "Guardar Documento"
    Case "Guardar Como"
        MensajeSTR = "Guardar Documento como"
    Case "&Imprimir"
        MensajeSTR = "Imprimir Documento"
    Case "&Salir"
        MensajeSTR = "Salir de CSAR SOFT MenuHint 1.0"
    Case "&Edicion"
        MensajeSTR = "Menu Edicion"
    Case "Copiar"
        MensajeSTR = "Copiar Texto"
    Case "Cortar"
        MensajeSTR = "Cortar Texto"
    Case "Pegar"
        MensajeSTR = "Pegar Contenido del Portapapeles"
    Case "Acerca de"
        MensajeSTR = "Acerca de Csar Soft Menu Hint"
 
End Select
Me.StatusBar1.Panels(1).Text = MensajeSTR
End Sub



Private Sub menuarchivo_Click(Index As Integer)
MsgBox Index
End Sub

Private Sub MDIForm_Load()
Me.StatusBar1.Panels(1).Style = sbrText
Me.CSMenuHint1.hWnd = Me.hWnd
Me.CSMenuHint1.subClass
End Sub



Private Sub MDIForm_Unload(Cancel As Integer)

Me.CSMenuHint1.UnSubClass
End Sub



