VERSION 5.00
Begin VB.UserControl Dialog 
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   375
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   360
   ScaleWidth      =   375
   Begin VB.Image imgLogo 
      Height          =   240
      Left            =   75
      Picture         =   "Dialog.ctx":0000
      Stretch         =   -1  'True
      Top             =   75
      Width           =   240
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'              LitePad------------------
 '15-8-2004
'(C)Tiger Ds, Nguyen Thanh Phat
'              http://ng-phat.hit.as
'              tigerproand@ yahoo.com
'neu ban nao co y kien hay phat hien loi xin cu mai ve hop mail cua toi, Xin cam on !!!!!!
'ban co the su dung phan mem va ma nguon nay tu do
'co the sao chep voi nhieu hinh thuc nhung:
'KHONG DUOC SU DUNG NO VAO MUC DICH THUONG MAI. CaM oN bAn Da Su DuNg ChUoNg Nay
'Option Explicit

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function CHOOSEFONT Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSEFONT) As Long
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long
Private Declare Function PageSetupDlg Lib "comdlg32.dll" Alias "PageSetupDlgA" (pPagesetupdlg As PageSetupDlg) As Long
Private Declare Function PrintDialog Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLG_TYPE) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Const LF_FACESIZE = 32
Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32

Private Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type

Private Type CHOOSEFONT
        lStructSize As Long
        hwndOwner As Long
        hdc As Long
        lpLogFont As Long
        iPointSize As Long
        flags As Long
        rgbColors As Long
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
        hInstance As Long
        lpszStyle As String
        nFontType As Integer
        MISSING_ALIGNMENT As Integer
        nSizeMin As Long
        nSizeMax As Long
End Type

Private Type ChooseColor
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        rgbResult As Long
        lpCustColors As Long
        flags As Long
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type

Private Type PRINTDLG_TYPE
        lStructSize As Long
        hwndOwner As Long
        hDevMode As Long
        hDevNames As Long
        hdc As Long
        flags As Long
        nFromPage As Integer
        nToPage As Integer
        nMinPage As Integer
        nMaxPage As Integer
        nCopies As Integer
        hInstance As Long
        lCustData As Long
        lpfnPrintHook As Long
        lpfnSetupHook As Long
        lpPrintTemplateName As String
        lpSetupTemplateName As String
        hPrintTemplate As Long
        hSetupTemplate As Long
End Type

Private Type DEVNAMES_TYPE
        wDriverOffset As Integer
        wDeviceOffset As Integer
        wOutputOffset As Integer
        wDefault As Integer
        extra As String * 100
End Type

Private Type DEVMODE_TYPE
        dmDeviceName As String * CCHDEVICENAME
        dmSpecVersion As Integer
        dmDriverVersion As Integer
        dmSize As Integer
        dmDriverExtra As Integer
        dmFields As Long
        dmOrientation As Integer
        dmPaperSize As Integer
        dmPaperLength As Integer
        dmPaperWidth As Integer
        dmScale As Integer
        dmCopies As Integer
        dmDefaultSource As Integer
        dmPrintQuality As Integer
        dmColor As Integer
        dmDuplex As Integer
        dmYResolution As Integer
        dmTTOption As Integer
        dmCollate As Integer
        dmFormName As String * CCHFORMNAME
        dmUnusedPadding As Integer
        dmBitsPerPel As Integer
        dmPelsWidth As Long
        dmPelsHeight As Long
        dmDisplayFlags As Long
        dmDisplayFrequency As Long
End Type

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type PageSetupDlg
        lStructSize As Long
        hwndOwner As Long
        hDevMode As Long
        hDevNames As Long
        flags As Long
        ptPaperSize As POINTAPI
        rtMinMargin As RECT
        rtMargin As RECT
        hInstance As Long
        lCustData As Long
        lpfnPageSetupHook As Long
        lpfnPagePaintHook As Long
        lpPageSetupTemplateName As String
        hPageSetupTemplate As Long
End Type

Private Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName As String * 31
End Type

 
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_EXPLORER = &H80000
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800
Private Const CF_PRINTERFONTS = &H2
Private Const CF_SCREENFONTS = &H1
Private Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Private Const CF_EFFECTS = &H100&
Private Const CF_FORCEFONTEXIST = &H10000
Private Const CF_INITTOLOGFONTSTRUCT = &H40&
Private Const CF_LIMITSIZE = &H2000&
Private Const DEFAULT_CHARSET = 1
Private Const DEFAULT_PITCH = 0
Private Const DEFAULT_QUALITY = 0
Private Const FW_BOLD = 700
Private Const FF_ROMAN = 16
Private Const FW_NORMAL = 400
Private Const CLIP_DEFAULT_PRECIS = 0
Private Const OUT_DEFAULT_PRECIS = 0
Private Const REGULAR_FONTTYPE = &H400
Private Const DM_DUPLEX = &H1000&
Private Const DM_ORIENTATION = &H1&


Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40

Private Const MAX_PATH = 260
Public cFileName As Collection
Public cFileTitle As Collection

 
Const m_def_CancelError = 0
Const m_def_Filename = ""
Const m_def_DialogTitle = ""
Const m_def_InitialDir = ""
Const m_def_Filter = ""
Const m_def_FilterIndex = 1
Const m_def_MultiSelect = 0
Const m_def_FontName = "Arial"
Const m_def_FontSize = 10
Const m_def_FontColor = 0
Const m_def_FontBold = 0
Const m_def_FontItalic = 0
Const m_def_FontUnderline = 0
Const m_def_FontStrikeThru = 0

 
Dim m_CancelError As Boolean
Dim m_Filename As String
Dim m_DialogTitle As String
Dim m_InitialDir As String
Dim m_Filter As String
Dim m_FilterIndex As Integer
Dim m_MultiSelect As Boolean
Dim m_FontName As String
Dim m_FontSize As Integer
Dim m_FontColor As Long
Dim m_FontBold As Boolean
Dim m_FontItalic As Boolean
Dim m_FontUnderline As Boolean
Dim m_FontStrikeThru As Boolean


Public Property Get CancelError() As Boolean
    CancelError = m_CancelError
End Property
Public Property Let CancelError(ByVal New_CancelError As Boolean)
    m_CancelError = New_CancelError
    PropertyChanged "CancelError"
End Property

Public Property Get MultiSelect() As Boolean
    MultiSelect = m_MultiSelect
End Property
Public Property Let MultiSelect(ByVal New_MultiSelect As Boolean)
    m_MultiSelect = New_MultiSelect
    PropertyChanged "MultiSelect"
End Property

Public Property Get DefaultFilename() As String
Attribute DefaultFilename.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    DefaultFilename = m_Filename
End Property
Public Property Let DefaultFilename(ByVal New_Filename As String)
    m_Filename = New_Filename
    PropertyChanged "DefaultFilename"
End Property

Public Property Get DialogTitle() As String
Attribute DialogTitle.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    DialogTitle = m_DialogTitle
End Property
Public Property Let DialogTitle(ByVal New_DialogTitle As String)
    m_DialogTitle = New_DialogTitle
    PropertyChanged "DialogTitle"
End Property

Public Property Get InitialDir() As String
    InitialDir = m_InitialDir
End Property
Public Property Let InitialDir(ByVal New_InitialDir As String)
    m_InitialDir = New_InitialDir
    PropertyChanged "InitialDir"
End Property

Public Property Get Filter() As String
    Filter = m_Filter
End Property
Public Property Let Filter(ByVal New_Filter As String)
    m_Filter = New_Filter
    PropertyChanged "Filter"
End Property

Public Property Get FilterIndex() As Integer
    FilterIndex = m_FilterIndex
End Property
Public Property Let FilterIndex(ByVal New_FilterIndex As Integer)
    m_FilterIndex = New_FilterIndex
    PropertyChanged "FilterIndex"
End Property

Public Property Get FontName() As String
    FontName = m_FontName
End Property
Public Property Let FontName(ByVal New_FontName As String)
    m_FontName = New_FontName
End Property

Public Property Get FontSize() As Integer
    FontSize = m_FontSize
End Property
Public Property Let FontSize(ByVal New_FontSize As Integer)
    m_FontSize = New_FontSize
End Property

Public Property Get FontColor() As Long
    FontColor = m_FontColor
End Property
Public Property Let FontColor(ByVal New_FontColor As Long)
    m_FontColor = New_FontColor
End Property

Public Property Get FontBold() As Boolean
    FontBold = m_FontBold
End Property
Public Property Let FontBold(ByVal New_FontBold As Boolean)
    m_FontBold = New_FontBold
End Property

Public Property Get FontItalic() As Boolean
    FontItalic = m_FontItalic
End Property
Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    m_FontItalic = New_FontItalic
End Property

Public Property Get FontUnderline() As Boolean
    FontUnderline = m_FontUnderline
End Property
Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    m_FontUnderline = New_FontUnderline
End Property

Public Property Get FontStrikeThru() As Boolean
    FontStrikeThru = m_FontStrikeThru
End Property
Public Property Let FontStrikeThru(ByVal New_FontStrikeThru As Boolean)
    m_FontStrikeThru = New_FontStrikeThru
End Property

Private Sub UserControl_Initialize()
    UserControl_Resize
End Sub

Private Sub UserControl_InitProperties()
    m_CancelError = m_def_CancelError
    m_Filename = m_def_Filename
    m_DialogTitle = m_def_DialogTitle
    m_InitialDir = m_def_InitialDir
    m_Filter = m_def_Filter
    m_FilterIndex = m_def_FilterIndex
    m_MultiSelect = m_def_MultiSelect
    m_FontName = m_def_FontName
    m_FontSize = m_def_FontSize
    m_FontColor = m_def_FontColor
    m_FontBold = m_def_FontBold
    m_FontItalic = m_def_FontItalic
    m_FontUnderline = m_def_FontUnderline
    m_FontStrikeThru = m_def_FontStrikeThru
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_CancelError = PropBag.ReadProperty("CancelError", m_def_CancelError)
    m_Filename = PropBag.ReadProperty("DefaultFilename", m_def_Filename)
    m_DialogTitle = PropBag.ReadProperty("DialogTitle", m_def_DialogTitle)
    m_InitialDir = PropBag.ReadProperty("InitialDir", m_def_InitialDir)
    m_Filter = PropBag.ReadProperty("Filter", m_def_Filter)
    m_FilterIndex = PropBag.ReadProperty("FilterIndex", m_def_FilterIndex)
    m_MultiSelect = PropBag.ReadProperty("MultiSelect", m_def_MultiSelect)
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = 360
    UserControl.Width = 375
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("CancelError", m_CancelError, m_def_CancelError)
    Call PropBag.WriteProperty("DefaultFilename", m_Filename, m_def_Filename)
    Call PropBag.WriteProperty("DialogTitle", m_DialogTitle, m_def_DialogTitle)
    Call PropBag.WriteProperty("InitialDir", m_InitialDir, m_def_InitialDir)
    Call PropBag.WriteProperty("Filter", m_Filter, m_def_Filter)
    Call PropBag.WriteProperty("FilterIndex", m_FilterIndex, m_def_FilterIndex)
    Call PropBag.WriteProperty("MultiSelect", m_MultiSelect, m_def_MultiSelect)
End Sub

Public Function ShowOpen()
On Error Resume Next
    Dim epOFN As OPENFILENAME
    Dim lngRet As Long
    With epOFN
    
        If MultiSelect Then
            .flags = OFN_ALLOWMULTISELECT Or OFN_EXPLORER Or OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY
            .lpstrFile = DefaultFilename & Space(9999 - Len(DefaultFilename)) & vbNullChar
            .lpstrFileTitle = Space(9999) & vbNullChar
        Else
            .flags = OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY
            .lpstrFile = DefaultFilename & String(MAX_PATH - Len(DefaultFilename), 0) & vbNullChar
            .lpstrFileTitle = String(MAX_PATH, 0) & vbNullChar
        End If

        .hwndOwner = UserControl.ContainerHwnd
        .lpstrFilter = SetFilter(Filter) & vbNullChar
        .lpstrInitialDir = InitialDir & vbNullChar
        .lpstrTitle = DialogTitle & vbNullChar
        .lStructSize = Len(epOFN)
        .nFilterIndex = FilterIndex
        .nMaxFile = Len(.lpstrFile)
        .nMaxFileTitle = Len(.lpstrFileTitle)
    End With
    
    lngRet = GetOpenFileName(epOFN)
    
    If lngRet <> 0 Then
        ParseFileName epOFN.lpstrFile
    Else
        If CancelError Then
          Err.Raise 32755, App.EXEName, "Cancel was selected.", "litepad.chm", 32755
        End If
    End If
End Function

Public Function ShowSave()
     Dim epOFN As OPENFILENAME
    Dim lngRet As Long
    With epOFN
        .hwndOwner = UserControl.ContainerHwnd
        .flags = OFN_PATHMUSTEXIST Or OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT
        .lpstrFile = DefaultFilename & String(MAX_PATH - Len(DefaultFilename), 0) & vbNullChar
        .lpstrFileTitle = String(MAX_PATH, 0) & vbNullChar
        .lpstrFilter = SetFilter(Filter) & vbNullChar
        .lpstrInitialDir = InitialDir & vbNullChar
        .lpstrTitle = DialogTitle & vbNullChar
        .lStructSize = Len(epOFN)
        .nFilterIndex = FilterIndex
        .nMaxFile = Len(.lpstrFile)
        .nMaxFileTitle = Len(.lpstrFileTitle)
    End With
    
    lngRet = GetSaveFileName(epOFN)
    
    If lngRet <> 0 Then
        ParseFileName epOFN.lpstrFile
    Else
        If CancelError Then
       
            Err.Raise 32755, App.EXEName, "Cancel was selected.", "litepad.chm", 32755
        End If
    End If
End Function

Public Function ShowFont()
  
    Dim CF As CHOOSEFONT
    Dim LF As LOGFONT
    Dim lMemHandle As Long
    Dim lLogFont As Long
    Dim lngRet As Long
    
    With LF
        .lfCharSet = DEFAULT_CHARSET
        .lfClipPrecision = CLIP_DEFAULT_PRECIS
        .lfFaceName = "Arial" & vbNullChar
        .lfHeight = 13
        .lfOutPrecision = OUT_DEFAULT_PRECIS
        .lfPitchAndFamily = DEFAULT_PITCH Or FF_ROMAN
        .lfQuality = DEFAULT_QUALITY
        .lfWeight = FW_NORMAL
    End With
    
   
    lMemHandle = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(LF))
    lLogFont = GlobalLock(lMemHandle)
    CopyMemory ByVal lLogFont, LF, Len(LF)
        
    With CF
        .flags = CF_BOTH Or CF_EFFECTS Or CF_FORCEFONTEXIST Or CF_INITTOLOGFONTSTRUCT Or CF_LIMITSIZE
        .hdc = Printer.hdc
        .hwndOwner = UserControl.ContainerHwnd
        .iPointSize = 120
        .lpLogFont = lLogFont
        .lStructSize = Len(CF)
        .nFontType = REGULAR_FONTTYPE
        .nSizeMax = 72
        .nSizeMin = 10
        .rgbColors = RGB(0, 0, 0)
    End With
    
    lngRet = CHOOSEFONT(CF)
    If lngRet <> 0 Then
        CopyMemory LF, ByVal lLogFont, Len(LF)

        FontName = Left(LF.lfFaceName, InStr(LF.lfFaceName, vbNullChar) - 1)
        FontSize = CF.iPointSize / 10
        FontColor = CF.rgbColors
        If LF.lfWeight = FW_NORMAL Then
            FontBold = False
            FontItalic = False
            FontUnderline = False
            FontStrikeThru = False
        Else
            If LF.lfWeight = FW_BOLD Then FontBold = True
            If LF.lfItalic <> 0 Then FontItalic = True
            If LF.lfUnderline <> 0 Then FontUnderline = True
            If LF.lfStrikeOut <> 0 Then FontStrikeThru = True
        End If
    Else
        If CancelError Then
  
            Err.Raise 32755, App.EXEName, "Cancel was selected.", "litepad.chm", 32755
        End If
    End If
    
  
    GlobalUnlock lMemHandle
    GlobalFree lMemHandle
End Function

Public Function ShowColor()
  
    Dim epCC As ChooseColor
    Dim lngRet As Long
    Dim CusCol(0 To 16) As Long
    Dim i As Integer
    
   
    For i = 0 To 15
        CusCol(i) = vbWhite
    Next
    
    With epCC
        .hwndOwner = UserControl.ContainerHwnd
        .lStructSize = Len(epCC)
        .lpCustColors = VarPtr(CusCol(0))
        .rgbResult = 0
    End With
    
    lngRet = ChooseColor(epCC)
    If lngRet <> 0 Then
        ShowColor = epCC.rgbResult
    Else
        If CancelError Then
         
            Err.Raise 32755, App.EXEName, "Cancel was selected.", "litepad.chm", 32755
        End If
    End If
End Function

Public Function ShowPageSetup()
   
    Dim epPSD As PageSetupDlg
    Dim lngRet As Long
    
    epPSD.lStructSize = Len(epPSD)
    epPSD.hwndOwner = UserControl.ContainerHwnd
    
    lngRet = PageSetupDlg(epPSD)
    If lngRet <> 0 Then
      
    Else
        If CancelError Then
  
            Err.Raise 32755, App.EXEName, "Cancel was selected.", "litepad.chm", 32755
        End If
    End If
End Function

Public Function ShowPrinter()

    Dim PrintDlg As PRINTDLG_TYPE
    Dim DevMode As DEVMODE_TYPE
    Dim DevName As DEVNAMES_TYPE

    Dim lpDevMode As Long, lpDevName As Long
    Dim bReturn As Integer
    Dim objPrinter As Printer, NewPrinterName As String


    PrintDlg.lStructSize = Len(PrintDlg)
    PrintDlg.hwndOwner = UserControl.ContainerHwnd

 
    DevMode.dmDeviceName = Printer.DeviceName
    DevMode.dmSize = Len(DevMode)
    DevMode.dmFields = DM_ORIENTATION Or DM_DUPLEX
    DevMode.dmPaperWidth = Printer.Width
    DevMode.dmOrientation = Printer.Orientation
    DevMode.dmPaperSize = Printer.PaperSize
    DevMode.dmDuplex = Printer.Duplex
    On Error GoTo 0

    PrintDlg.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevMode))
    lpDevMode = GlobalLock(PrintDlg.hDevMode)
    If lpDevMode > 0 Then
        CopyMemory ByVal lpDevMode, DevMode, Len(DevMode)
        bReturn = GlobalUnlock(PrintDlg.hDevMode)
    End If

    With DevName
        .wDriverOffset = 8
        .wDeviceOffset = .wDriverOffset + 1 + Len(Printer.DriverName)
        .wOutputOffset = .wDeviceOffset + 1 + Len(Printer.Port)
        .wDefault = 0
    End With

    With Printer
        DevName.extra = .DriverName & Chr(0) & .DeviceName & Chr(0) & .Port & Chr(0)
    End With

    PrintDlg.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevName))
    lpDevName = GlobalLock(PrintDlg.hDevNames)
    If lpDevName > 0 Then
        CopyMemory ByVal lpDevName, DevName, Len(DevName)
        bReturn = GlobalUnlock(lpDevName)
    End If

    If PrintDialog(PrintDlg) <> 0 Then

        
        lpDevName = GlobalLock(PrintDlg.hDevNames)
        CopyMemory DevName, ByVal lpDevName, 45
        bReturn = GlobalUnlock(lpDevName)
        GlobalFree PrintDlg.hDevNames

      
        lpDevMode = GlobalLock(PrintDlg.hDevMode)
        CopyMemory DevMode, ByVal lpDevMode, Len(DevMode)
        bReturn = GlobalUnlock(PrintDlg.hDevMode)
        GlobalFree PrintDlg.hDevMode
        NewPrinterName = UCase$(Left(DevMode.dmDeviceName, InStr(DevMode.dmDeviceName, Chr$(0)) - 1))
        If Printer.DeviceName <> NewPrinterName Then
            For Each objPrinter In Printers
                If UCase$(objPrinter.DeviceName) = NewPrinterName Then
                    Set Printer = objPrinter
               
                End If
            Next
        End If

        On Error Resume Next
     
        Printer.Copies = DevMode.dmCopies
        Printer.Duplex = DevMode.dmDuplex
        Printer.Orientation = DevMode.dmOrientation
        Printer.PaperSize = DevMode.dmPaperSize
        Printer.PrintQuality = DevMode.dmPrintQuality
        Printer.ColorMode = DevMode.dmColor
        Printer.PaperBin = DevMode.dmDefaultSource
        On Error GoTo 0
    Else
        If CancelError Then
     
            Err.Raise 32755, App.EXEName, "Cancel was selected.", "litepad.chm", 32755
        End If
    End If
End Function

Private Function ParseFileName(sFileName As String)

    Dim i As Long
    Dim sPath As String
    Dim sFiles() As String
    Dim Pos As Integer
    Dim sFile As String
    Dim sFileTitle As String
    

    Set cFileName = New Collection
    Set cFileTitle = New Collection
   
    Pos = InStr(sFileName, vbNullChar & vbNullChar)
 
    sFile = Left(sFileName, Pos - 1)
    
  
    If InStr(1, sFile, vbNullChar) <> 0 Then
   
        sFile = Left(sFileName, Pos) & vbNullChar
        sPath = Left(sFileName, InStr(1, sFileName, Chr(0)) - 1)
        sFiles = Split(sFile, Chr(0))
        
       
        For i = LBound(sFiles) To UBound(sFiles) - 2
      
            If Right(sPath, 1) = "\" Then
                cFileName.Add sPath & sFiles(i)
            Else
                cFileName.Add sPath & "\" & sFiles(i)
            End If
      
            cFileTitle.Add sFiles(i)
         
            If i = 1 Then cFileName.Remove 1: cFileTitle.Remove 1
        Next
    Else
  
        cFileName.Add sFile
 
        cFileTitle.Add Right(sFile, Len(sFile) - InStrRev(sFile, "\"))
    End If
End Function

Private Function SetFilter(sFlt As String) As String
       Dim sLen As Long
    Dim Pos As Long

    sLen = Len(sFlt)
    Pos = InStr(1, sFlt, "|")

  
    While Pos > 0

        sFlt = Left(sFlt, Pos - 1) & vbNullChar & Mid(sFlt, Pos + 1, sLen - Pos)
       
        Pos = InStr(Pos + 1, sFlt, "|")
    Wend
    SetFilter = sFlt
End Function
