VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSubClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------
'CZAR SOFT MENU Hint 1.0
'Control para visualizar mensajes de los menues
'cesargazzo@yahoo.com.ar
'*Los Grandes Artistas Copiamos, Los Pesimos ROBAN*
'---------------------------------------------------------
Option Explicit

Private mhWnd       As Long
'
Event MenuCaption(Caption As String)

 
Friend Sub fRaiseEvent(MenuCaptionSTR As String)
 
If MenuCaptionSTR <> "Nada" Then RaiseEvent MenuCaption(MenuCaptionSTR)
   
End Sub

Friend Property Get hWnd() As Long
    hWnd = mhWnd
End Property

Friend Property Let hWnd(ByVal vNewValue As Long)
    mhWnd = vNewValue
End Property

 
