VERSION 5.00
Begin VB.UserControl CSMenuHint 
   ClientHeight    =   570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   555
   InvisibleAtRuntime=   -1  'True
   Picture         =   "CSMenuHint.ctx":0000
   PropertyPages   =   "CSMenuHint.ctx":0C42
   ScaleHeight     =   570
   ScaleWidth      =   555
End
Attribute VB_Name = "CSMenuHint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'---------------------------------------------------------
'CZAR SOFT MENU Hint 1.0
'Control para visualizar mensajes de los menues
'cesargazzo@yahoo.com.ar
'*Los Grandes Artistas Copiamos, Los Pesimos ROBAN*
'---------------------------------------------------------
Option Explicit

Private WithEvents moSubClass As cSubClass
Attribute moSubClass.VB_VarHelpID = -1
'
'Event Declarations:
Event MostrarMensaje(Mensaje As String)
'Default Property Values:
Const m_def_hWnd = 0
'Property Variables:
Dim m_hWnd As Long


 



 

Private Sub moSubClass_MenuCaption(Caption As String)
RaiseEvent MostrarMensaje(Caption)

End Sub
 
 

Private Sub UserControl_Resize()
On Error Resume Next
UserControl.Width = 32 * Screen.TwipsPerPixelX
UserControl.Height = 32 * Screen.TwipsPerPixelY
End Sub

 
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=14
Public Function subClass()
    If Ambient.UserMode = False Then
        Exit Function
    End If
    '## Engage subclassing
    Set moSubClass = New cSubClass
    moSubClass.hWnd = Me.hWnd
    mSubClass.Hook moSubClass
End Function

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=14
Public Function UnSubClass()
  
    If Ambient.UserMode = False Then
        Exit Function
    End If
    mSubClass.UnHook
    Set moSubClass = Nothing
End Function

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=8,0,0,0
Public Property Get hWnd() As Long
Attribute hWnd.VB_ProcData.VB_Invoke_Property = "Menues"
    hWnd = m_hWnd
End Property

Public Property Let hWnd(ByVal New_hWnd As Long)
    m_hWnd = New_hWnd
    PropertyChanged "hWnd"
End Property

'Inicializar propiedades para control de usuario
Private Sub UserControl_InitProperties()
    m_hWnd = m_def_hWnd
End Sub

'Cargar valores de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_hWnd = PropBag.ReadProperty("hWnd", m_def_hWnd)
End Sub

'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("hWnd", m_hWnd, m_def_hWnd)
End Sub

