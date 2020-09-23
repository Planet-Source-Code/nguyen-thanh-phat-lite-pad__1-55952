Attribute VB_Name = "ontop"
'              LitePad------------------
' 15-8-2004
'(C)Tiger Ds, Nguyen Thanh Phat
'              http://ng-phat.hit.as
'              tigerproand@yahoo.com
'neu ban nao co y kien hay phat hien loi xin cu mai ve hop mail cua toi, Xin cam on !!!!!!
'ban co the su dung phan mem va ma nguon nay tu do
'co the sao chep voi nhieu hinh thuc nhung:
'KHONG DUOC SU DUNG NO VAO MUC DICH THUONG MAI. CaM oN bAn Da Su DuNg ChUoNg Nay
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Boolean
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hWndCallback As Long) As Long
  Public Const HWND_TOPMOST = -1
  Public Const HWND_NOTOPMOST = -2
  Public Const SWP_NOACTIVATE = &H10
  Public Const SWP_SHOWWINDOW = &H40
Public Sub AlwaysOnTop(formname As Form, SetOnTop As Boolean)

    If SetOnTop Then
        lFlag = HWND_TOPMOST
    Else
        lFlag = HWND_NOTOPMOST
    End If
    SetWindowPos formname.hWnd, lFlag, _
    formname.Left / Screen.TwipsPerPixelX, _
    formname.Top / Screen.TwipsPerPixelY, _
    formname.Width / Screen.TwipsPerPixelX, _
    formname.Height / Screen.TwipsPerPixelY, _
    SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub
Function vbmciSendString(ByVal Command As String, ByVal hWnd As Long) As String
    Dim Buffer As String
    Dim dwRet As Long
    Buffer = Space$(100)
    dwRet = mciSendString(Command, ByVal Buffer, Len(Buffer), hWnd)
    vbmciSendString = Buffer
End Function
Public Sub CreateRegion()
Dim hRgn1, hRgn2, hRgn3 As Long
Dim ret As Long
   hRgn1 = CreateRectRgn(0, 0, 1, 2)
   hRgn2 = CreateEllipticRgn(0, 0, 110, 100)
   hRgn3 = CreateEllipticRgn(150, 0, 500, 200)
   ret = CombineRgn(hRgn1, hRgn1, hRgn2, 2)
   ret = CombineRgn(hRgn1, hRgn1, hRgn3, 2)
   ret = SetWindowRgn(frmabout.hWnd, hRgn1, True)
End Sub
