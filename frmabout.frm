VERSION 5.00
Begin VB.Form frmabout 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmabout.frx":0000
   ScaleHeight     =   2505
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail: tigerproand@yahoo.com"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   960
      Picture         =   "frmabout.frx":95D4
      Top             =   1680
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   960
      Picture         =   "frmabout.frx":9B5E
      Top             =   1320
      Width           =   240
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "http://ng-phat.hit.as"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cam On  Nguyen Hong Minh Da Cung Cap VnSpeech"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   2280
      Width           =   3855
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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
Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub Form_Click()
Unload Me
End Sub
Private Sub Form_Load()
AlwaysOnTop Me, True
'Call CreateRegion
End Sub
Private Sub Form_Unload(Cancel As Integer)
frmmain.Show
End Sub
Private Sub Label2_Click()
Dim ret&
ret& = ShellExecute(Me.hwnd, "Open", "http://ng-phat.hit.as", "", App.Path, 1)
End Sub

Private Sub Label3_Click()
Dim ret&
ret& = ShellExecute(Me.hwnd, "open", "mailto:tigerproand@yahoo.com", "", App.Path, 1)

End Sub
