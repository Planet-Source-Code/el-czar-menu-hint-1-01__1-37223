Attribute VB_Name = "mSubClass"
'---------------------------------------------------------
'CZAR SOFT MENU Hint 1.0
'Control para visualizar mensajes de los menues
'cesargazzo@yahoo.com.ar
'*Los Grandes Artistas Copiamos, Los Pesimos ROBAN*
'---------------------------------------------------------
Option Explicit

Private moSubClass          As cSubClass
Public defWindowProc        As Long
Public Const GWL_WNDPROC    As Long = (-4)
Public Const MF_BYCOMMAND = &H0&
Public Const MF_BYPOSITION = &H400&
Public Const MF_STRING = &H0&
Public Const MF_GRAYED = &H1&
Public Const MF_DISABLED = &H2&
Public Const MF_BITMAP = &H4&
Public Const MF_CHECKED = &H8&
Public Const MF_POPUP = &H10&
Public Const MF_HILITE = &H80&
Public Const MF_OWNERDRAW = &H100&
Public Const MF_SEPARATOR = &H800&
Public Const MF_SYSMENU = &H2000&
Public Const MF_MOUSESELECT = &H8000&
Public Const WM_MENUSELECT = &H11F
Public Const WM_MENUCHAR = &H120

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long

Public Sub Hook(oSubClass As cSubClass)

    On Error Resume Next
    Set moSubClass = oSubClass
    defWindowProc = SetWindowLong(moSubClass.hWnd, GWL_WNDPROC, AddressOf WindowProc)
   
End Sub


Public Sub UnHook()
  
    '
    If defWindowProc Then
        SetWindowLong moSubClass.hWnd, GWL_WNDPROC, defWindowProc
        defWindowProc = 0
    End If
    Set moSubClass = Nothing

End Sub

Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error Resume Next

Dim lhMenu As Long
Dim lMenuItem As Long
Dim lFlags As Long

If (hWnd = moSubClass.hWnd) And (uMsg = WM_MENUSELECT Or uMsg = WM_MENUCHAR) Then
    lMenuItem = LoWord(wParam)
    lFlags = HiWord(wParam)
    lhMenu = lParam
    moSubClass.fRaiseEvent GetMenuHint(lhMenu, lMenuItem, lFlags)
    WindowProc = 0
Else
    moSubClass.fRaiseEvent "Nada"
    WindowProc = CallWindowProc(defWindowProc, hWnd, uMsg, wParam, lParam)
    
End If
  
  
End Function

Function GetMenuHint(ByVal lhMenu As Long, ByVal lMenuItem As Long, ByVal lFlags As Long) As String

  Dim sMenuString As String
  Dim lResult As Long
  Dim lcmdFlag As Long
  
  GetMenuHint = ""
  
' Flags which indicates, that the item is not a valid selected menu-entry.
  If (lFlags And MF_SEPARATOR) = MF_SEPARATOR Then Exit Function
  If (lFlags And MF_HILITE) = 0 Then Exit Function
  
  lcmdFlag = MF_BYCOMMAND
  If (lFlags And MF_POPUP) = MF_POPUP Then lcmdFlag = MF_BYPOSITION

' Get Item-Caption
  sMenuString = Space(100)
  lResult = GetMenuString(lhMenu, lMenuItem, sMenuString, 100, lcmdFlag)
  If lResult > 0 Then
    GetMenuHint = Trim(Left(sMenuString, lResult))
  Else
    Exit Function
  End If

 

End Function

Function HiWord(ByVal lDWord As Long) As Long

  Dim i As Long
  Dim dblTemp As Double

' Generate unsigned 32-bit value, if param is negative
' To prevent getting the VB "Overflow"-Error, dont add more than &H7FFFFFFF at a time.
  dblTemp = lDWord
  If dblTemp < 0 Then
    dblTemp = &H7FFFFFFF
    dblTemp = dblTemp + &H7FFFFFFF
    dblTemp = (dblTemp + 2) - Abs(lDWord)
  End If
  
' No "Shift"-operator in VB. Must be divided by two, 16 times.
  For i = 0 To 15
    dblTemp = Fix(dblTemp / 2)
  Next i

  lDWord = dblTemp
  HiWord = lDWord

End Function

Function LoWord(ByVal lDWord As Long) As Long

  Dim dblTemp As Double
  
' Generate unsigned 32-bit value, if param is negative
' To prevent getting the VB "Overflow"-Error, dont add more than &H7FFFFFFF at a time.
  dblTemp = lDWord
  If dblTemp < 0 Then
    dblTemp = &H7FFFFFFF
    dblTemp = dblTemp + &H7FFFFFFF
    dblTemp = (dblTemp + 2) - Abs(lDWord)
  End If
  
' To prevent getting the VB "Overflow"-Error with the "AND"-operation, delete the signed bit first.
  If dblTemp > &H7FFFFFFF Then
    dblTemp = dblTemp - &H7FFFFFFF
    dblTemp = dblTemp - 1
  End If
  
  lDWord = dblTemp
  lDWord = lDWord And 65535

  LoWord = lDWord

End Function

