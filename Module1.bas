Attribute VB_Name = "document"
'              LitePad------------------
' 15-8-2004
'(C)Tiger Ds, Nguyen Thanh Phat
'              http://ng-phat.hit.as
'              tigerproand@yahoo.com
'neu ban nao co y kien hay phat hien loi xin cu mai ve hop mail cua toi, Xin cam on !!!!!!
'ban co the su dung phan mem va ma nguon nay tu do
'co the sao chep voi nhieu hinh thuc nhung:
'KHONG DUOC SU DUNG NO VAO MUC DICH THUONG MAI. CaM oN bAn Da Su DuNg ChUoNg Nay
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wp As Long, lp As Any) As Long
Public Const WM_USER = &H400
Public Const EM_UNDO = WM_USER + 23
Public Const rtbfilter = "Text File (*.txt)|*.txt|RichText Files (*.rtb)|*.rtb|Log Files (*.log)|*.log|Batch Files (*.bat)|*.bat|INI Files (*.ini)|*.ini|All Files (*.*)|*.*|"
Public Const rtbsave = "Text File (*.txt)|*.txt|RichText Files (*.rtb)|*.rtb|Log Files (*.log)|*.log|Batch Files (*.bat)|*.bat|INI Files (*.ini)|*.ini|"
Public Const WM_CUT = &H300
Public Const WM_COPY = &H301
Public Const WM_PASTE = &H302
Public Const EM_CANUNDO = &HC6
'--------------------Read VietNamese--------------------
Declare Function VietTTS Lib "VNSPEECH.DLL" (ByVal test As String) As Integer
