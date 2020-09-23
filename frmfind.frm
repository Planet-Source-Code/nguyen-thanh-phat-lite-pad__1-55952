VERSION 5.00
Begin VB.Form frmfind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find And Replace"
   ClientHeight    =   2085
   ClientLeft      =   5595
   ClientTop       =   4485
   ClientWidth     =   5310
   Icon            =   "frmfind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtReplace 
      Height          =   285
      Left            =   1080
      TabIndex        =   12
      Top             =   480
      Width           =   3135
   End
   Begin VB.TextBox txtFind 
      Height          =   315
      Left            =   1080
      TabIndex        =   11
      Top             =   0
      Width           =   3135
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   315
      Left            =   4275
      TabIndex        =   8
      Top             =   0
      Width           =   990
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   4275
      TabIndex        =   7
      Top             =   375
      Width           =   990
   End
   Begin VB.PictureBox picBar 
      BorderStyle     =   0  'None
      Height          =   1290
      Left            =   0
      ScaleHeight     =   1290
      ScaleWidth      =   5340
      TabIndex        =   0
      Top             =   825
      Width           =   5340
      Begin VB.Frame Frame1 
         Caption         =   "Search Options"
         Height          =   1215
         Left            =   75
         TabIndex        =   3
         Top             =   0
         Width           =   4065
         Begin VB.CheckBox chkNoHighlight 
            Caption         =   "No &Highlight"
            Height          =   240
            Left            =   150
            TabIndex        =   6
            Top             =   900
            Width           =   1965
         End
         Begin VB.CheckBox chkMatchCase 
            Caption         =   "Match Ca&se"
            Height          =   240
            Left            =   120
            TabIndex        =   5
            Top             =   600
            Width           =   1965
         End
         Begin VB.CheckBox chkWholeWord 
            Caption         =   "Find Whole Word &Only"
            Height          =   240
            Left            =   150
            TabIndex        =   4
            Top             =   300
            Width           =   1965
         End
      End
      Begin VB.CommandButton cmdReplace 
         Caption         =   "&Replace..."
         Height          =   315
         Left            =   4275
         TabIndex        =   2
         Top             =   120
         Width           =   990
      End
      Begin VB.CommandButton cmdReplaceAll 
         Caption         =   "Replace &All"
         Height          =   315
         Left            =   4270
         TabIndex        =   1
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Label lblFind 
      Caption         =   "Fin&d What:"
      Height          =   240
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   840
   End
   Begin VB.Label lblReplace 
      Caption         =   "Replace &With:"
      Height          =   240
      Left            =   0
      TabIndex        =   9
      Top             =   450
      Width           =   1065
   End
End
Attribute VB_Name = "frmfind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdFind_Click()
'click cmdfind xay ra
    Dim lngResult As Long
    Dim lngPos As Long
    Dim intOptions As Integer
    If chkNoHighlight.Value = 1 Then intOptions = intOptions + 8
    If chkWholeWord.Value = 1 Then intOptions = intOptions + 2
    If chkMatchCase.Value = 1 Then intOptions = intOptions + 4

    If cmdFind.Caption = "&Find" Then
        lngResult = frmmain.rtb.Find(txtFind.Text, 0, , intOptions)

        If lngResult = -1 Then
            MsgBox "Text not found", , "LitePad - Find"
            cmdFind.Caption = "&Find"
            frmmain.mnufindnext.Enabled = False
        Else
            frmmain.SetFocus
            cmdReplace.Enabled = True
            cmdReplaceAll.Enabled = True
            cmdFind.Caption = "&Find Next"
            frmmain.mnufindnext.Enabled = True
        End If
    Else
        lngPos = frmmain.rtb.SelStart + frmmain.rtb.SelLength
        lngResult = frmmain.rtb.Find(txtFind.Text, lngPos, , intOptions)

        If lngResult = -1 Then
            MsgBox "Text not found", "LitePad - Find"
            cmdFind.Caption = "&Find"
            cmdReplace.Enabled = False
            cmdReplaceAll.Enabled = False
            frmmain.mnufindnext.Enabled = False
        Else
            frmmain.SetFocus
            frmmain.mnufindnext.Enabled = True
        End If
    End If


End Sub
Private Sub cmdReplace_Click()
    On Error Resume Next
    Dim lngResult As Long
    Dim lngPos As Long
    Dim intOptions As Integer
    
    If cmdReplace.Caption = "&Replace..." Then
        cmdReplace.Top = 150
        cmdReplace.Caption = "&Replace"
        lblReplace.Visible = True
        txtReplace.Visible = True
        cmdReplaceAll.Visible = True
        Exit Sub
    End If

   
    If chkNoHighlight.Value = 1 Then intOptions = intOptions + 8
    If chkWholeWord.Value = 1 Then intOptions = intOptions + 2
    If chkMatchCase.Value = 1 Then intOptions = intOptions + 4
    With frmmain.rtb
        .SelText = txtReplace.Text
        lngPos = .SelStart + .SelLength
        lngResult = .Find(txtFind.Text, lngPos, , intOptions)
        
        If lngResult = -1 Then
            MsgBox "Text not found", "LitePad - Replace"
            cmdFind.Caption = "&Find"
            cmdReplace.Enabled = False
            cmdReplaceAll.Enabled = False
        Else
            .SetFocus
        End If
    End With

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdReplaceAll_Click()
    On Error Resume Next
    Dim intCount As Integer
    Dim lngPos As Long
    Dim intOptions As Integer
    
    If chkNoHighlight.Value = 1 Then intOptions = intOptions + 8
    If chkWholeWord.Value = 1 Then intOptions = intOptions + 2
    If chkMatchCase.Value = 1 Then intOptions = intOptions + 4

    intCount = 0
    lngPos = 0
    With frmmain.rtb
        Do
            If .Find(txtFind.Text, lngPos, , intOptions) = -1 Then
                If intCount > 0 Then
                    MsgBox "The specified region has been searched. " & vbCrLf & _
                    intCount & " replacements have been made.", "Litepad - ReplaceAll"
                End If
                cmdFind.Caption = "&Find"
                cmdReplace.Enabled = False
                cmdReplaceAll.Enabled = False
                Exit Do
            Else
                lngPos = .SelStart + .SelLength
                intCount = intCount + 1
                .SelText = txtReplace.Text
            End If
        Loop
    End With
End Sub

Private Sub Form_Load()
    cmdReplace.Top = 525
    lblReplace.Visible = False
    txtReplace.Visible = False
    cmdReplaceAll.Visible = False
    txtFind.Text = frmmain.rtb.SelText
End Sub


