VERSION 5.00
Begin VB.UserControl TrayControl 
   AutoRedraw      =   -1  'True
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1065
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   33
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   71
   ToolboxBitmap   =   "TrayControl.ctx":0000
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   480
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   0
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      Picture         =   "TrayControl.ctx":0312
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "TrayControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'              LitePad------------------
' 15-8-2004
'(C)Tiger Ds, Nguyen Thanh Phat
'              http://ng-phat.hit.as
'              tigerproand@yahoo.com
'neu ban nao co y kien hay phat hien loi xin cu mai ve hop mail cua toi, Xin cam on !!!!!!
'ban co the su dung phan mem va ma nguon nay tu do
'co the sao chep voi nhieu hinh thuc nhung:
'KHONG DUOC SU DUNG NO VAO MUC DICH THUONG MAI. CaM oN bAn Da Su DuNg ChUoNg Nay
Option Explicit

Private m_TipText As String
Private Const def_TipText = ""


Public frm As Form
Public IconObject As Object
Public lngPrevWndProc As Long
Public lngWndID As Long
Public lngHwnd As Long
Private Notify As NOTIFYICONDATA
Private BarData As APPBARDATA

Private Const GW_CHILD = 5
Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDLAST = 1
Private Const GW_HWNDNEXT = 2
Private Const GW_HWNDPREV = 3
Private Const GWL_WNDPROC = (-4)
Private Const IDANI_OPEN = &H1
Private Const IDANI_CLOSE = &H2
Private Const IDANI_CAPTION = &H3
Private Const NIF_TIP = &H4
Private Const NIM_ADD = 0&
Private Const NIM_DELETE = 2&
Private Const NIM_MODIFY = 1&
Private Const NIF_ICON = 2&
Private Const NIF_MESSAGE = 1&
Private Const ABM_GETTASKBARPOS = &H5&
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_MOUSEMOVE = &H200
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_USER = &H400

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, _
    ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Private Declare Function DrawAnimatedRects Lib "user32" (ByVal hwnd As Long, ByVal idAni As Long, lprcFrom As RECT, _
    lprcTo As RECT) As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
    ByVal hwnd As Long, ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, _
    ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Private Declare Function SHAppBarMessage Lib "shell32.dll" (ByVal dwMessage As Long, pData As APPBARDATA) As Long

Public Enum ZoomTypes
    ZOOM_FROM_TRAY
    ZOOM_TO_TRAY
End Enum

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

Private Type APPBARDATA
        cbSize As Long
        hwnd As Long
        uCallbackMessage As Long
        uEdge As Long
        rc As RECT
        lParam As Long
End Type

Public Function SendToTray()
    Dim lngRetVal As Long

    ZoomForm ZOOM_TO_TRAY, frm.hwnd
    frm.Visible = False
    Picture2.Picture = frm.Icon
        Set IconObject = frm.Icon
        AddIcon frm, IconObject.Handle, IconObject, m_TipText
    
End Function

Public Function RestoreFromTray()


    delIcon IconObject.Handle
    frm.Icon = Picture2.Picture
    ZoomForm ZOOM_FROM_TRAY, frm.hwnd
    frm.Visible = True
    
End Function

Public Property Get TipText() As String

    TipText = m_TipText

End Property

Public Property Let TipText(ByVal New_TipText As String)
    
    m_TipText = New_TipText
    PropertyChanged "TipText"

End Property

Private Sub UserControl_InitProperties()
       m_TipText = def_TipText

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    

    Set frm = Parent
    m_TipText = PropBag.ReadProperty("TipText", def_TipText)
    
End Sub

Private Sub UserControl_Resize()
    
    With UserControl
        .Height = 450
        .Width = 450
    End With

    Line (0, 0)-(ScaleWidth, 0), vb3DHighlight
    Line (2, 2)-(ScaleWidth - 2, 2), vb3DDKShadow
    Line (0, 0)-(0, ScaleHeight), vb3DHighlight
    Line (2, 2)-(2, ScaleHeight - 2), vb3DDKShadow
    Line (ScaleWidth - 1, 1)-(ScaleWidth - 1, ScaleHeight), vb3DDKShadow
    Line (ScaleWidth - 3, 3)-(ScaleWidth - 3, ScaleHeight - 3), vb3DHighlight
    Line (ScaleWidth - 1, ScaleHeight - 1)-(0, ScaleHeight - 1), vb3DDKShadow
    Line (ScaleWidth - 3, ScaleHeight - 3)-(1, ScaleHeight - 3), vb3DHighlight
    Refresh

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("TipText", m_TipText, def_TipText)
    
End Sub

Public Function ZoomForm(zoomToWhere As ZoomTypes, hwnd As Long) As Boolean
    
    Dim rctFrom As RECT
    Dim rctTo As RECT
    Dim lngTrayHand As Long
    Dim lngStartMenuHand As Long
    Dim lngChildHand As Long
    Dim strClass As String * 255
    Dim lngClassNameLen As Long
    Dim lngRetVal As Long

  
    Select Case zoomToWhere

        Case ZOOM_FROM_TRAY
      
            lngStartMenuHand = FindWindow("Shell_TrayWnd", vbNullString)


            lngChildHand = GetWindow(lngStartMenuHand, GW_CHILD)


            Do
                
                lngClassNameLen = GetClassName(lngChildHand, strClass, Len(strClass))


                If InStr(1, strClass, "TrayNotifyWnd") Then
                    lngTrayHand = lngChildHand
                    Exit Do
                End If

                lngChildHand = GetWindow(lngChildHand, GW_HWNDNEXT)
            
            Loop


            lngRetVal = GetWindowRect(hwnd, rctFrom)

    
            lngRetVal = GetWindowRect(lngTrayHand, rctTo)

    
            lngRetVal = DrawAnimatedRects(frm.hwnd, IDANI_CLOSE Or IDANI_CAPTION, rctTo, rctFrom)

        Case ZOOM_TO_TRAY

           
            lngStartMenuHand = FindWindow("Shell_TrayWnd", vbNullString)


            lngChildHand = GetWindow(lngStartMenuHand, GW_CHILD)


            Do
                
                lngClassNameLen = GetClassName(lngChildHand, strClass, Len(strClass))

                If InStr(1, strClass, "TrayNotifyWnd") Then
                    lngTrayHand = lngChildHand
                    Exit Do
                End If

                lngChildHand = GetWindow(lngChildHand, GW_HWNDNEXT)
            
            Loop

            lngRetVal = GetWindowRect(hwnd, rctFrom)

            lngRetVal = GetWindowRect(lngTrayHand, rctTo)

            lngRetVal = DrawAnimatedRects(frm.hwnd, IDANI_OPEN Or IDANI_CAPTION, rctFrom, rctTo)
    
    End Select

End Function

Public Sub modIcon(Form1 As Form, IconID As Long, Icon As Object, ToolTip As String)

    Dim Result As Long
    Notify.cbSize = 88&
    Notify.hwnd = Form1.hwnd
    Notify.uID = IconID
    Notify.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    Notify.uCallbackMessage = WM_MOUSEMOVE
    Notify.hIcon = Icon
    Notify.szTip = ToolTip & Chr$(0)
    Result = Shell_NotifyIcon(NIM_MODIFY, Notify)

End Sub

Public Sub AddIcon(Form1 As Form, IconID As Long, Icon As Object, ToolTip As String)
    

    Dim Result As Long
    BarData.cbSize = 36&
    Result = SHAppBarMessage(ABM_GETTASKBARPOS, BarData)
    Notify.cbSize = 88&
    Notify.hwnd = Form1.hwnd
    Notify.uID = IconID
    Notify.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    Notify.uCallbackMessage = WM_MOUSEMOVE
    Notify.hIcon = Icon
    Notify.szTip = ToolTip & Chr$(0)
    Result = Shell_NotifyIcon(NIM_ADD, Notify)

End Sub

Public Sub delIcon(IconID As Long)
    

    Dim Result As Long
    Notify.uID = IconID
    Result = Shell_NotifyIcon(NIM_DELETE, Notify)

End Sub
