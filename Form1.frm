VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmmain 
   Caption         =   "LitePad "
   ClientHeight    =   5040
   ClientLeft      =   1395
   ClientTop       =   2010
   ClientWidth     =   7620
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   7620
   Begin MSComctlLib.Toolbar tbformat 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   360
      Width           =   7620
      _ExtentX        =   13441
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "img"
      DisabledImageList=   "img"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Font"
            Object.ToolTipText     =   "Font "
            ImageKey        =   "Font"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Object.ToolTipText     =   "Bold"
            ImageKey        =   "Bold"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Object.ToolTipText     =   "Italic"
            ImageKey        =   "Italic"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Underline"
            Object.ToolTipText     =   "Underline"
            ImageKey        =   "Underline"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Strike"
            Object.ToolTipText     =   "Strikethrough"
            ImageKey        =   "Strike"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Left"
            Object.ToolTipText     =   "Left"
            ImageKey        =   "Left"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Center"
            Object.ToolTipText     =   "Center"
            ImageKey        =   "Center"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Right"
            Object.ToolTipText     =   "Right"
            ImageKey        =   "Right"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bullet"
            Object.ToolTipText     =   "Bullet"
            ImageKey        =   "Bullet"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SuperScript"
            Object.ToolTipText     =   "SuperScript"
            ImageKey        =   "SuperScript"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SubScript"
            Object.ToolTipText     =   "SubScript"
            ImageKey        =   "SubScript"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "NoScript"
            Object.ToolTipText     =   "No Script"
            ImageKey        =   "NoScript"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Indent"
            Object.ToolTipText     =   "Indent"
            ImageKey        =   "Outdent"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Outdent"
            Object.ToolTipText     =   "Outdent"
            ImageKey        =   "Indent"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList img 
      Left            =   7560
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   34
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":151A
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1AB6
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2052
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":25EE
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":298A
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2D26
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":32C2
            Key             =   "Bullet"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":385E
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3DFA
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4396
            Key             =   "Redo"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4932
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4ECE
            Key             =   "Strike"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":546A
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5A06
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5FA2
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":653E
            Key             =   "Left"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6ADA
            Key             =   "Right"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7076
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7612
            Key             =   "Replace"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7BAE
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":814A
            Key             =   "Text"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":86E6
            Key             =   "Home"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8C82
            Key             =   "Font"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":921E
            Key             =   "Lower"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":97BA
            Key             =   "Upper"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9D56
            Key             =   "Indent"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A2F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A88E
            Key             =   "SubScript"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":AE2A
            Key             =   "SuperScript"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B3C6
            Key             =   "Normal"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B962
            Key             =   "NoScript"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BEFE
            Key             =   "Outdent"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C49A
            Key             =   "Picture"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":CA36
            Key             =   "Read"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stb 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   4785
      Width           =   7620
      _ExtentX        =   13441
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "LitePad"
            TextSave        =   "LitePad"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1552
            MinWidth        =   1552
            Text            =   "Total Line"
            TextSave        =   "Total Line"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   979
            MinWidth        =   970
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            Object.Width           =   1587
            MinWidth        =   1587
            TextSave        =   "9:59 AM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1482
            MinWidth        =   1482
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   1482
            MinWidth        =   1482
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            Object.Width           =   1658
            MinWidth        =   1658
            TextSave        =   "8/31/04"
         EndProperty
      EndProperty
   End
   Begin LitePad.TrayControl TrayControl1 
      Left            =   7500
      Top             =   4800
      _ExtentX        =   794
      _ExtentY        =   794
   End
   Begin RichTextLib.RichTextBox rtb 
      Height          =   4215
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   7435
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":CFD2
   End
   Begin MSComctlLib.Toolbar tb 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7620
      _ExtentX        =   13441
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "img"
      DisabledImageList=   "img"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New Document"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open Document"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save Document"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Close"
            Object.ToolTipText     =   "Close Document"
            ImageKey        =   "Close"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Object.ToolTipText     =   "Find Text"
            ImageKey        =   "Find"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Replace"
            Object.ToolTipText     =   "Find And Replace Text"
            ImageKey        =   "Replace"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Picture"
            Object.ToolTipText     =   "Insert Picture"
            ImageKey        =   "Picture"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Text"
            Object.ToolTipText     =   "Insert Text"
            ImageKey        =   "Text"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Home"
            Object.ToolTipText     =   "My Homepage"
            ImageKey        =   "Home"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Read"
            Object.ToolTipText     =   "Read VietNamese"
            ImageKey        =   "Read"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.PictureBox picinsert 
      Height          =   495
      Left            =   6840
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   1
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin LitePad.Dialog CmDlg 
      Left            =   6360
      Top             =   2160
      _ExtentX        =   661
      _ExtentY        =   635
      CancelError     =   -1  'True
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   3240
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtbtemp 
      Height          =   255
      Left            =   5280
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":D08C
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
         Shortcut        =   ^Q
      End
      Begin VB.Menu sp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnusave 
         Caption         =   "&Save "
         Shortcut        =   ^S
      End
      Begin VB.Menu mnusaveas 
         Caption         =   "Save As..."
      End
      Begin VB.Menu sp2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "P&rint"
         Shortcut        =   ^P
      End
      Begin VB.Menu sp3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   {F1}
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "Edit"
      Begin VB.Menu mnuundo 
         Caption         =   "Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu sp4 
         Caption         =   "-"
      End
      Begin VB.Menu mnucut 
         Caption         =   "&Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "C&opy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu sp5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuselect 
         Caption         =   "Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuclear 
         Caption         =   "C&lear All"
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu sp6 
         Caption         =   "-"
      End
      Begin VB.Menu mnufind 
         Caption         =   "Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnufindnext 
         Caption         =   "Find Next"
         Enabled         =   0   'False
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnufindreplace 
         Caption         =   "Find and Replace"
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnutoolbar 
         Caption         =   "Toolbar"
         Begin VB.Menu mnustandand 
            Caption         =   "Standand"
         End
         Begin VB.Menu mnustatusbar 
            Caption         =   "Statusbar"
         End
      End
      Begin VB.Menu sp15 
         Caption         =   "-"
      End
      Begin VB.Menu mnuontop 
         Caption         =   "Alway On Top"
         Shortcut        =   ^{F11}
      End
      Begin VB.Menu mnutray 
         Caption         =   "Sent to tray"
         Shortcut        =   ^{F12}
      End
   End
   Begin VB.Menu mnuformat 
      Caption         =   "Format"
      Begin VB.Menu mnuFont 
         Caption         =   "Font..."
         Shortcut        =   ^T
      End
      Begin VB.Menu sp7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBullet 
         Caption         =   "Bullet"
         Shortcut        =   ^B
      End
      Begin VB.Menu sp8 
         Caption         =   "-"
      End
      Begin VB.Menu mnucase 
         Caption         =   "Change Case"
         Begin VB.Menu mnuupper 
            Caption         =   "LITE PAD"
         End
         Begin VB.Menu mnulower 
            Caption         =   "lite pad"
         End
         Begin VB.Menu mnuproper 
            Caption         =   "Lite Pad"
         End
      End
      Begin VB.Menu sp9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuscript 
         Caption         =   "Script"
         Begin VB.Menu mnunoscript 
            Caption         =   "No Scripting"
            Shortcut        =   ^{F1}
         End
         Begin VB.Menu mnuSuperScript 
            Caption         =   "SuperScript"
            Shortcut        =   ^{F2}
         End
         Begin VB.Menu mnuSubScript 
            Caption         =   "SubScript"
            Shortcut        =   ^{F3}
         End
      End
      Begin VB.Menu sp21 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIndent 
         Caption         =   "Indent"
      End
      Begin VB.Menu mnuoutdent 
         Caption         =   "Outdent"
      End
      Begin VB.Menu sp16 
         Caption         =   "-"
      End
      Begin VB.Menu mnuleft 
         Caption         =   "Align Left"
      End
      Begin VB.Menu mnucenter 
         Caption         =   "Align Center"
      End
      Begin VB.Menu mnuright 
         Caption         =   "Align Right"
      End
   End
   Begin VB.Menu mnutool 
      Caption         =   "Tool"
      Begin VB.Menu mnuReadV 
         Caption         =   "Read (VietNamese)"
      End
      Begin VB.Menu sp10 
         Caption         =   "-"
      End
      Begin VB.Menu mnucaulator 
         Caption         =   "Caculator"
      End
   End
   Begin VB.Menu mnuInsert 
      Caption         =   "Insert"
      Begin VB.Menu mnutime 
         Caption         =   "Time"
      End
      Begin VB.Menu mnudate 
         Caption         =   "Date"
      End
      Begin VB.Menu sp11 
         Caption         =   "-"
      End
      Begin VB.Menu mnupicture 
         Caption         =   "Picture"
         Shortcut        =   ^{INSERT}
      End
      Begin VB.Menu mnutext 
         Caption         =   "Text File"
         Shortcut        =   +{INSERT}
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "Help"
      Begin VB.Menu mnuabout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnupopuptray 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnurestore 
         Caption         =   "Restore "
      End
      Begin VB.Menu sp20 
         Caption         =   "-"
      End
      Begin VB.Menu mnuaboutp 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnupopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnupopupundo 
         Caption         =   "Undo"
      End
      Begin VB.Menu sp18 
         Caption         =   "-"
      End
      Begin VB.Menu mnupopupcut 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnupopupcopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnupopuppaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu sp19 
         Caption         =   "-"
      End
      Begin VB.Menu mnupopupall 
         Caption         =   "Select All"
      End
   End
End
Attribute VB_Name = "frmmain"
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
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wp As Long, lp As Any) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Dim blnInTray As Boolean
Dim changed As Boolean
Public rtbas As RichTextBox
Const SW_SHOWNORMAL = 1
Const WM_RBUTTONUP = &H205
Const WM_LBUTTONDBLCLK = &H203
Private Sub Form_Activate()
stb.Panels(3).Text = 1 + rtb.GetLineFromChar(Len(rtb.Text))
End Sub

Private Sub Form_Load()
'thiet lap form.caption
Me.Caption = Me.Caption & "--" & "Untitled"
tbformat.Buttons("Left").Value = tbrPressed
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If changed Then
If MsgBox("The file " & Me.Caption & " has been changed" & vbCrLf & "Are you sure you want to save", vbYesNo, Me.Caption) = vbYes Then mnusaveas_Click
Else
Unload Me
End If
End Sub

Private Sub Form_Resize()
'thay doi kich thuoc rtb khi form thay doi kich thuoc
If Me.WindowState = 1 Then Exit Sub
rtb.Width = Me.Width - 130
rtb.Height = Me.Height - 1390
End Sub

Private Sub mnuAbout_Click()
frmabout.Show
End Sub

Private Sub mnuaboutp_Click()
mnuAbout_Click
End Sub
Private Sub mnucaulator_Click()
    'goi chuong trinh caculator
    ShellExecute Me.hwnd, vbNullString, "calc.exe", vbNullString, "calc.exe", SW_SHOWNORMAL
End Sub

Private Sub mnucenter_Click()
  rtb.SelAlignment = rtfCenter
        tb.Buttons("Left").Value = tbrUnpressed
        tb.Buttons("Center").Value = tbrPressed
        tb.Buttons("Right").Value = tbrUnpressed
        tb.Refresh
        rtb.SetFocus

End Sub

Private Sub mnuclear_Click()
'xoa tat ca trong rtb
rtb.Text = ""
End Sub
Private Sub mnuClose_Click()
'---------msgbox khi mnuclose duoc click
If rtb.Text = "" Then: Exit Sub
If MsgBox("Do are you want so close, without save?", vbYesNo) = vbYes Then
rtb.Text = ""
Else
mnusave_Click
End If
End Sub
Private Sub mnuCopy_Click()
'copy
Clipboard.SetText rtb.SelRTF
End Sub
Private Sub mnucut_Click()
'cut
  Clipboard.SetText rtb.SelRTF
    rtb.SelText = vbNullString
End Sub
Private Sub mnudate_Click()
'chen thoi gian
'co the chen thoi gian ngan nhu sau
'rtb.seltext = format(Date$,"Short Date")
rtb.SelText = Format(Date$, "Long Date")
End Sub
Private Sub mnuExit_Click()
End
End Sub
Private Sub mnufind_Click()
    frmfind.Show , Me
End Sub

Private Sub mnufindnext_Click()

    On Error Resume Next
    Dim lngResult As Integer
    Dim lngPos As Integer
    Dim intOptions As Integer
    
   If frmfind.chkNoHighlight.Value = 1 Then intOptions = intOptions + 8
    If frmfind.chkWholeWord.Value = 1 Then intOptions = intOptions + 2
    If frmfind.chkMatchCase.Value = 1 Then intOptions = intOptions + 4

    lngPos = rtb.SelStart + rtb.SelLength
    lngResult = rtb.Find(frmfind.txtFind.Text, lngPos, , intOptions)

    If lngResult = -1 Then
        MsgBox "Text not found", "LitePad - FindNext", 1, True
        frmfind.cmdFind.Caption = "&Find"
        frmfind.cmdReplace.Enabled = False
        frmfind.cmdReplaceAll.Enabled = False
        mnufindnext.Enabled = False
    Else
        rtb.SetFocus
    End If

End Sub

Private Sub mnufindreplace_Click()
    With frmfind
        .cmdReplace.Top = 150 'Set cmdReplace top
        .cmdReplace.Caption = "&Replace" 'Set caption
        .lblReplace.Visible = True 'Show lblReplace
        .txtReplace.Visible = True 'Show cboReplace
        .cmdReplaceAll.Visible = True 'Show cmdReplaceAll
        .Show , Me
    End With
End Sub
Private Sub mnuFont_Click()
'show frmmfont
frmfont.Show
End Sub

Private Sub mnuIndent_Click()
'tao buoc cach, tuong tu nhu nut Tab
Me.ScaleMode = vbMillimeters
rtb.SelIndent = rtb.SelIndent + 13
meScaleMode = vbTwips
End Sub



Private Sub mnuleft_Click()
   rtb.SelAlignment = rtfLeft
        tb.Buttons("Left").Value = tbrPressed
        tb.Buttons("Center").Value = tbrUnpressed
        tb.Buttons("Right").Value = tbrUnpressed
        tb.Refresh
       End Sub

Private Sub mnulower_Click()
'cho rtb la chu thuong
On Error Resume Next
rtb.SelText = LCase(rtb.SelText)
End Sub
Private Sub mnuBullet_Click()
 'chen ki tu dau dong
    With rtb
          If (IsNull(.SelBullet) = True) Or (.SelBullet = False) Then
            .SelBullet = True
            tbformat.Buttons("Bullet").Value = tbrPressed
        ElseIf .SelBullet = True Then
            .SelBullet = False
            .SelHangingIndent = False
            tbformat.Buttons("Bullet").Value = tbrUnpressed
        End If
    End With
End Sub
Private Sub mnuNew_Click()
If rtb.Text = "" Then
Exit Sub
Else
If MsgBox("The file " & cd1.FileTitle & " has been changed,Do are you want to save", vbYesNo) = vbNo Then
    Cancel = True
    rtb.Text = ""
    Else
    mnusaveas_Click
End If
End If
End Sub

Private Sub mnunoscript_Click()
rtb.SelCharOffset = 0
End Sub

Private Sub mnuontop_Click()
'thiet lap cho form o tren cung
mnuontop.Checked = Not mnuontop.Checked
If mnuontop.Checked Then
AlwaysOnTop Me, True
Else
AlwaysOnTop Me, False
End If
End Sub

Private Sub mnuOpen_Click()
' xem mot tai lieu
On Error GoTo openproblem
Dim o As String
cd1.Filter = rtbfilter
cd1.DialogTitle = "Open file..."
cd1.ShowOpen
cd1.CancelError = True
o = cd1.FileName
rtb.LoadFile o
Me.Caption = "LitePad " & " " & cd1.FileName
openproblem:
If Err.Number = 32755 Then Exit Sub
End Sub

Private Sub mnuoutdent_Click()
Me.ScaleMode = 6
rtb.SelIndent = rtb.SelIndent - 13
Me.ScaleMode = 1
End Sub
Private Sub mnuPaste_Click()
'SendMessage rtb.hwnd, WM_PASTE, 0&, 0& 'Paste
    On Error Resume Next
    Screen.ActiveControl.SelText = Clipboard.GetText
End Sub
Private Sub mnupicture_Click()
'chen hinh vao rtb
    On Error GoTo Errorpic
    cd1.DialogTitle = "Select Picture..."
    cd1.Filter = "Image Files|*.bmp;*.jpg;*.gif|Bitmap (*.bmp)|*.bmp|JPEG (*.jpg)|*.jpg|GIF (*.gif)|*.gif|All Picture|*.bmp;*.jpg;*.gif"
    cd1.CancelError = True
    cd1.ShowOpen
    picinsert.Picture = LoadPicture(cd1.FileName)
    Clipboard.Clear
    Clipboard.SetData picinsert.Picture
    SendMessage rtb.hwnd, WM_PASTE, 0, 0&
Errorpic:
    Cancel = True
End Sub
Private Sub mnupopupall_Click()
mnuselect_Click
End Sub
Private Sub mnupopupcopy_Click()
mnuCopy_Click
End Sub
Private Sub mnupopupcut_Click()
mnucut_Click
End Sub
Private Sub mnupopuppaste_Click()
mnuPaste_Click
End Sub
Private Sub mnupopupundo_Click()
mnuundo_Click
End Sub

Private Sub mnuPrint_Click()
'print
 On Error Resume Next
With cd1
    .PrinterDefault = True
    .flags = cdlPDDisablePrintToFile Or cdlPDNoPageNums
    If rtb.SelLength = 0 Then
        .flags = .flags Or cdlPDNoSelection
    Else
           .flags = .flags Or cdlPDSelection
    End If
    .CancelError = True
    .ShowPrinter
    If Err = 0 Then
        If .flags And cdlPDSelection Then
            Printer.Print rtb.SelText
        Else
            Printer.Print rtb.Text
        End If
    End If
End With
End Sub

Private Sub mnuproper_Click()
rtb.SelText = StrConv(rtb.SelText, vbProperCase)
End Sub

'cach doc mot van ban tieng anh
'nhan Ctrl+T, chon Microsoft Direct Speech Synthesis
'cho control moi vao form voi phan Name: Sp
'Privare sub mnureadE_Click()
'sp.Speak rtb.text
'end sub
'-----------------Read VietNamese---------
Private Sub mnuReadV_Click()
'ban phai compile thi moi co tac dung
On Error GoTo rvproblem 'bay loi
Dim readv As Integer
readv = VietTTS(rtb.Text)
rvproblem:
MsgBox "Can't find file Vnspeech.dll"
End Sub
Private Sub mnurestore_Click()
'cho hien thi from
TrayControl1.RestoreFromTray
End Sub

Private Sub mnuright_Click()
  rtb.SelAlignment = rtfRight
        tb.Buttons("Left").Value = tbrUnpressed
        tb.Buttons("Center").Value = tbrUnpressed
        tb.Buttons("Right").Value = tbrPressed
        tb.Refresh
        rtb.SetFocus

End Sub

Private Sub mnusave_Click()
'cach save mot tai lieu
On Error GoTo SaveError 'bay loi
Dim sType As String
With frmmain
        If Left(.Caption, 8) = "Document" Then
            mnusaveas_Click
        Else
            If UCase(Right(.Caption, 3)) = "rtf" Then
                sType = rtfText
            Else
                sType = rtfText
            End If
            .rtb.SaveFile .Caption, sType
        End If
    End With
    changed = False
SaveError:
    Exit Sub
End Sub
Private Sub mnusaveas_Click()
'cach save as mot tai lieu
On Error GoTo saveasproblem 'bay loi
Dim s As String

cd1.Filter = rtbsave
cd1.CancelError = True
cd1.DialogTitle = "Save As..."
cd1.flags = cdlOFNExplorer And cdlOFNLongNames
cd1.ShowSave

On Error GoTo saveasproblem
s = cd1.FileName

    If cd1.FilterIndex = 1 Then
    
        cd1.DefaultExt = "txt"
        rtb.SaveFile s, rtfText
        
        ElseIf cd1.FilterIndex = 2 Then
        cd1.DefaultExt = "rtf"
        rtb.SaveFile s, rtfRTF
        
        ElseIf cd1.FilterIndex = 3 Then
        cd1.DefaultExt = "log"
        rtb.SaveFile s, rtfText
        
        ElseIf cd1.FilterIndex = 4 Then
        cd1.DefaultExt = "bat"
        rtb.SaveFile s, rtfText
    Else
        cd1.DefaultExt = "ini"
        rtb.SaveFile s, rtfText
  End If
saveasproblem:
If Err.Number = 32755 Then Exit Sub
End Sub
Private Sub mnuselect_Click() 'chon tat ca
rtb.SelStart = 0 'chon diem dau tien
rtb.SelLength = Len(rtb.Text) 'va diem ket thuc
End Sub

Private Sub mnustandand_Click()
mnustandand.Checked = Not mnustandand.Checked

If mnustandand.Checked Then
tb.Visible = False
Else
tb.Visible = True

End If

End Sub

Private Sub mnustatusbar_Click()
mnustatusbar.Checked = Not mnustatusbar.Checked

If mnustatusbar.Checked Then
stb.Visible = False
Else
stb.Visible = True
End If
End Sub
Private Sub mnuSubScript_Click()
rtb.SelCharOffset = -50
End Sub
Private Sub mnuSuperScript_Click()
rtb.SelCharOffset = 50
End Sub
Private Sub mnutext_Click()
'chen mot doan text khac vao van ban
    On Error GoTo InsertError
    Dim sType As String
'title
    CmDlg.DialogTitle = "Select File to Insert"
    CmDlg.Filter = rtbfilter
    CmDlg.CancelError = True
    CmDlg.ShowOpen
    CmDlg.CancelError = True
'lay phan duoi mo rong cua file duoc chon
    Select Case UCase(Right(CmDlg.cFileTitle(1), 3))
        Case "rtf"
            sType = rtfRTF
        Case Else
            sType = rtfText
    End Select
'load file duoc chon vao rtbtemp
    rtbtemp.LoadFile CmDlg.cFileName(1), sType
    rtbtemp.SelStart = 0
    rtbtemp.SelLength = Len(rtbtemp.Text)
'cut doan text trong rtbtemp
    SendMessage rtbtemp.hwnd, WM_CUT, 0, 0&
'paste doan text do vao rtb
    rtb.SelText = SendMessage(rtb.hwnd, WM_PASTE, 0, 0&)
InsertError:
    If Err.Number = 32755 Then Exit Sub
  End Sub
Private Sub mnutime_Click()
'chen thoi gian
rtb.SelText = Format(Time$, "Long Time")
End Sub
Private Sub mnutray_Click()
' cho form vao system tray
mnutray.Checked = Not mnutray.Checked
If mnutray.Checked Then
TrayControl1.SendToTray
Else
TrayControl1.RestoreFromTray
End If
End Sub
Private Sub mnuundo_Click()
'undo
SendMessage rtb.Text, EM_UNDO, 0&, 0&
End Sub
Private Sub mnuupper_Click()
'cho rtb.text bang chu hoa
On Error Resume Next
rtb.SelText = UCase(rtb.SelText)
End Sub

Private Sub rtb_Change()
changed = True
End Sub

Private Sub rtb_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
 If Button = vbRightButton Then
        PopupMenu mnupopup  'show popup menu
    End If
End Sub
Private Sub rtb_SelChange()
stb.Panels(3).Text = 1 + rtb.GetLineFromChar(Len(rtb.Text)) 'lay so dong
If IsNull(rtb.SelBold) Then
        tbformat.Buttons("Bold").MixedState = True
Else
        tbformat.Buttons("Bold").MixedState = False
        tbformat.Buttons("Bold").Value = IIf(rtb.SelBold, tbrPressed, tbrUnpressed)
End If
If IsNull(rtb.SelItalic) Then
        tbformat.Buttons("Italic").MixedState = True
Else
        tbformat.Buttons("Italic").MixedState = False
        tbformat.Buttons("Italic").Value = IIf(rtb.SelItalic, tbrPressed, tbrUnpressed)
End If
If IsNull(rtb.SelUnderline) Then
        tbformat.Buttons("Underline").MixedState = True
Else
        tbformat.Buttons("Underline").MixedState = False
        tbformat.Buttons("Underline").Value = IIf(rtb.SelUnderline, tbrPressed, tbrUnpressed)
End If
If IsNull(rtb.SelStrikeThru) Then
        tbformat.Buttons("Strike").MixedState = True
Else
        tbformat.Buttons("Strike").MixedState = False
        tbformat.Buttons("Strike").Value = IIf(rtb.SelStrikeThru, tbrPressed, tbrUnpressed)
End If
If IsNull(rtb.SelBullet) Then
        tbformat.Buttons("Bullet").MixedState = True
   Else
        tbformat.Buttons("Bullet").MixedState = False
        tbformat.Buttons("Bullet").MixedState = IIf(rtb.SelItalic, tbrPressed, tbrUnpressed)
End If
If rtb.SelAlignment = rtfLeft Then
            tbformat.Buttons("Left").Value = tbrPressed
            tbformat.Buttons("Center").Value = tbrUnpressed
            tbformat.Buttons("Right").Value = tbrUnpressed
ElseIf rtb.SelAlignment = rtfCenter Then
            tbformat.Buttons("Left").Value = tbrUnpressed
           tbformat.Buttons("Center").Value = tbrPressed
           tbformat.Buttons("Right").Value = tbrUnpressed
ElseIf rtb.SelAlignment = rtfRight Then
            tbformat.Buttons("Left").Value = tbrUnpressed
            tbformat.Buttons("Center").Value = tbrUnpressed
            tbformat.Buttons("Right").Value = tbrPressed
End If
End Sub



Private Sub tb_ButtonClick(ByVal Button As MSComctlLib.Button)
 On Error Resume Next
    Select Case Button.Key
        Case "New"
            mnuNew_Click
        Case "Print"
            mnuPrint_Click
        Case "Open"
            mnuOpen_Click
        Case "Save"
            mnusave_Click
        Case "Undo"
            mnuundo_Click
        Case "Replace"
            mnufindreplace_Click
        Case "Cut"
            mnucut_Click
        Case "Copy"
            mnuCopy_Click
        Case "Paste"
            mnuPaste_Click
        Case "Find"
            mnufind_Click
        Case "Insert"
            mnupicture_Click
        Case "Home"
        Dim ret&
            ret& = ShellExecute(Me.hwnd, "Open", "http://ng-phat.hit.as", "", App.Path, 1)
        Case "Text"
            mnutext_Click
        Case "Picture"
            mnupicture_Click
        Case "Read"
            mnuReadV_Click
    End Select
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'hien popupmenu khi r_click
    If Button = 2 Then PopupMenu mnupopuptray
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
        Static Message As Long
    Message = x / Screen.TwipsPerPixelX
            Select Case Message
                Case WM_LBUTTONDBLCLK 'khi nhan chuot trai
                            Call mnurestore_Click 'popup menu restore
                Case WM_RBUTTONUP: 'khi nhan chuot phai
           Me.PopupMenu mnupopuptray ' popup menu tray
    End Select
End Sub

Private Sub tbformat_ButtonClick(ByVal Button As MSComctlLib.Button)
       Select Case Button.Key
            Case "Font"
                mnuFont_Click
            Case "Strike"
                rtb.SelStrikeThru = Not rtb.SelStrikeThru
                Button.Value = IIf(rtb.SelStrikeThru, tbrPressed, tbrubpressed)
            Case "Bold"
                rtb.SelBold = Not rtb.SelBold
                Button.Value = IIf(rtb.SelBold, tbrPressed, tbrUnpressed)
            Case "Italic"
                rtb.SelItalic = Not rtb.SelItalic
                Button.Value = IIf(rtb.SelItalic, tbrPressed, tbrUnpressed)
            Case "Underline"
                rtb.SelUnderline = Not rtb.SelUnderline
                Button.Value = IIf(rtb.SelUnderline, tbrPressed, tbrUnpressed)
            Case "Left"
            rtb.SelAlignment = rtfLeft
            tbformat.Refresh
            rtb.SetFocus
            Case "Center"
            rtb.SelAlignment = rtfCenter
            tbformat.Refresh
            rtb.SetFocus
            Case "Right"
            rtb.SelAlignment = rtfRight
            tbformat.Refresh
            rtb.SetFocus
            Case "Bullet"
                mnuBullet_Click
            Case "Lower"
                mnulower_Click
            Case "Upper"
                mnuupper_Click
            Case "Nomal"
                mnuproper_Click
            Case "NoScript"
                mnunoscript_Click
            Case "SuperScript"
                mnuSuperScript_Click
            Case "SubScript"
                mnuSubScript_Click
            Case "Indent"
                mnuIndent_Click
            Case "Outdent"
                mnuoutdent_Click
    End Select
End Sub
