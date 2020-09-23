VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmfont 
   Caption         =   "Font"
   ClientHeight    =   3885
   ClientLeft      =   5610
   ClientTop       =   2190
   ClientWidth     =   4515
   Icon            =   "frmfont.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   4515
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   14
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   600
      TabIndex        =   13
      Top             =   3480
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5880
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Color:"
      Height          =   615
      Left            =   2040
      TabIndex        =   10
      Top             =   1920
      Width           =   1335
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   720
         TabIndex        =   15
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000006&
         ForeColor       =   &H008080FF&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Font Style:"
      Height          =   1695
      Left            =   2040
      TabIndex        =   5
      Top             =   120
      Width           =   1335
      Begin VB.CheckBox Check4 
         Caption         =   "Strikethru"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Underline"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   975
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Italic"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Bold"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.ListBox List2 
      Height          =   1815
      Left            =   3600
      TabIndex        =   4
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3600
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   360
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LitePad"
      Height          =   735
      Left            =   1920
      TabIndex        =   12
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Font Size:"
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Font:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmfont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'              LitePad------------------
'               15-8-2004
'(C)Tiger Ds, Nguyen Thanh Phat
'              http://ng-phat.hit.as
'              tigerproand@yahoo.com
'neu ban nao co y kien hay phat hien loi xin cu mai ve hop mail cua toi, Xin cam on !!!!!!
'ban co the su dung phan mem va ma nguon nay tu do
'co the sao chep voi nhieu hinh thuc nhung:
'KHONG DUOC SU DUNG NO VAO MUC DICH THUONG MAI. CaM oN bAn Da Su DuNg ChUoNg Nay

Dim ret As Integer
Private Sub Check1_Click()
If Check1.Value = 1 Then
  Label4.FontBold = True
Else
  Label4.FontBold = False
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
  Label4.FontItalic = True
Else
  Label4.FontItalic = False
End If
End Sub

Private Sub Check3_Click()

If Check3.Value = 1 Then
  Label4.FontUnderline = True
Else

  Label4.FontUnderline = False
End If

End Sub

Private Sub Check4_Click()
If Check4.Value = 1 Then

  Label4.FontStrikeThru = True
Else

  Label4.FontStrikeThru = False
End If

End Sub

Private Sub Command1_Click()

With frmmain.rtb
.Font.Name = Label4.FontName
.Font.Size = Label4.FontSize
.Font.Bold = Label4.FontBold
.Font.Italic = Label4.FontItalic
.Font.Underline = Label4.FontUnderline
.Font.Strikethrough = Label4.FontStrikeThru
.SelColor = Label4.ForeColor
.BackColor = Label5.BackColor
End With

Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()

Debug.Print frmmain.rtb
For x = 1 To Screen.FontCount
List1.AddItem Screen.Fonts(x)
Next
For x = 5 To 72: List2.AddItem Str$(x): Next 'thiet lap kich thuoc font chu tu 5 toi 72

For x = 0 To List1.ListCount - 1
 If FontName = List1.List(x) Then
  List1.ListIndex = x
  Label4.FontName = List1.List(x)
  Exit For
 End If
Next

For x = 0 To List2.ListCount - 1
 If Int(Val(frmmain.rtb.Font.Size)) = Val(List2.List(x)) Then
  List2.ListIndex = x
  Label4.FontSize = Val(List2.List(x))
  Text1.Text = List2.List(x)
  Exit For
 End If
Next

If frmmain.rtb.Font.Bold = True Then
 Label4.FontBold = True
 Check1.Value = 1
End If

If frmmain.rtb.Font.Italic = True Then
 Label4.FontItalic = True
 Check2.Value = 1
End If

If frmmain.rtb.Font.Underline = True Then
 Label4.FontUnderline = True
 Check3.Value = 1
End If

If frmmain.rtb.Font.Strikethrough = True Then
 Label4.FontStrikeThru = True
 Check4.Value = 1
End If

Label3.ForeColor = frmmain.rtb.BackColor
Label5.ForeColor = frmmain.ForeColor
End Sub

Private Sub Label3_Click()

CommonDialog1.ShowColor
Label3.BackColor = CommonDialog1.Color
Label4.ForeColor = CommonDialog1.Color

End Sub

Private Sub Label5_Click()
CommonDialog1.ShowColor
Label4.BackColor = CommonDialog1.Color
Label5.BackColor = CommonDialog1.Color
End Sub

Private Sub List1_Click()

On Error Resume Next
Label4.FontName = List1.List(List1.ListIndex)
End Sub
Private Sub List2_Click()
Text1.Text = List2.List(List2.ListIndex)
Label4.FontSize = Val(Text1.Text)
End Sub
Private Sub Text1_Change()
For x = 0 To List2.ListCount - 1
If Val(Text1.Text) = Val(List2.List(x)) Then
 List2.ListIndex = x
 Label4.FontSize = Val(Text1.Text)
 Exit For
End If
Next
End Sub
