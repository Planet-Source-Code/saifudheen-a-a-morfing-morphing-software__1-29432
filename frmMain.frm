VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Morfy Ver1.0"
   ClientHeight    =   3375
   ClientLeft      =   2520
   ClientTop       =   3285
   ClientWidth     =   9390
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   225
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   626
   Begin VB.CommandButton Command1 
      Caption         =   "&About"
      Height          =   285
      Left            =   7200
      TabIndex        =   6
      Top             =   1680
      Width           =   795
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save Frames"
      Enabled         =   0   'False
      Height          =   315
      Left            =   7230
      TabIndex        =   11
      Top             =   2970
      Width           =   1785
   End
   Begin VB.CommandButton cmdTips 
      Caption         =   "&Next Tip"
      Height          =   285
      Left            =   8550
      TabIndex        =   7
      ToolTipText     =   "Next Tip."
      Top             =   1680
      Width           =   795
   End
   Begin VB.CommandButton cmdSwap 
      Caption         =   "S&wap"
      Height          =   375
      Left            =   2100
      TabIndex        =   12
      ToolTipText     =   "Swap"
      Top             =   2940
      Width           =   615
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6480
      TabIndex        =   5
      ToolTipText     =   "Next"
      Top             =   2970
      Width           =   585
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Sto&p"
      Height          =   315
      Left            =   5910
      TabIndex        =   4
      ToolTipText     =   "Stop Animation"
      Top             =   2970
      Width           =   585
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "&Play"
      Height          =   315
      Left            =   5340
      TabIndex        =   3
      ToolTipText     =   "Play Animation"
      Top             =   2970
      Width           =   585
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4770
      TabIndex        =   2
      ToolTipText     =   "Back "
      Top             =   2970
      Width           =   585
   End
   Begin VB.TextBox txtRate 
      Height          =   285
      Left            =   7230
      TabIndex        =   10
      Text            =   "20"
      Top             =   2610
      Width           =   435
   End
   Begin VB.PictureBox picBuf 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   -3000
      ScaleHeight     =   1035
      ScaleWidth      =   2115
      TabIndex        =   19
      Top             =   4770
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   2580
      Left            =   4770
      ScaleHeight     =   170
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   270
      Width           =   2280
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   2580
      Left            =   2430
      ScaleHeight     =   170
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   270
      Width           =   2280
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   2580
      Left            =   90
      ScaleHeight     =   170
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   270
      Width           =   2280
   End
   Begin VB.CommandButton cmdLImage 
      Caption         =   "&2"
      Height          =   375
      Left            =   3390
      TabIndex        =   1
      ToolTipText     =   "Insert last Image"
      Top             =   2940
      Width           =   615
   End
   Begin VB.CommandButton cmdFImage 
      Caption         =   "&1"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      ToolTipText     =   "Insert First Image"
      Top             =   2940
      Width           =   615
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   9420
      Top             =   3450
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "iles (*.bmp)|*.bmp|JPG /JPEG- JPEG Files|*.jpg|GIF - Compuserve GIF (*.gif) |*.gif|All Files (*.*)|*.*"
   End
   Begin VB.CheckBox chkReverseLoop 
      Caption         =   "Reverse Loop."
      Enabled         =   0   'False
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   7230
      TabIndex        =   9
      Top             =   2310
      Width           =   1425
   End
   Begin VB.CheckBox chkLoop 
      Caption         =   "Loop."
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   7230
      TabIndex        =   8
      Top             =   2070
      Width           =   945
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   41
      Left            =   9450
      Top             =   2970
   End
   Begin VB.Label Label6 
      Caption         =   "Tips :"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   7170
      TabIndex        =   22
      Top             =   60
      Width           =   705
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   1305
      Left            =   7170
      TabIndex        =   21
      Top             =   270
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "Frames/ Second."
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   7710
      TabIndex        =   20
      Top             =   2670
      Width           =   1395
   End
   Begin VB.Image imgFrame 
      Height          =   450
      Index           =   0
      Left            =   -1920
      Top             =   5940
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label Label5 
      Caption         =   "Output :"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   4770
      TabIndex        =   15
      Top             =   30
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Final Image : (150 x 170)"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   2430
      TabIndex        =   14
      Top             =   30
      Width           =   1905
   End
   Begin VB.Label Label2 
      Caption         =   "Start Image : (150 x 170)"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   30
      Width           =   1965
   End
   Begin VB.Menu mnuFirst 
      Caption         =   "Dummy"
      Visible         =   0   'False
      Begin VB.Menu mnuInsertPicture1 
         Caption         =   "Insert Picture"
      End
      Begin VB.Menu mnuShowOriginalSize 
         Caption         =   "Show Original Size"
      End
      Begin VB.Menu mnuSizeTo1 
         Caption         =   "Size All Frames to this Size"
      End
      Begin VB.Menu mnuStretchToStandard 
         Caption         =   "Stretch to Standard Size"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Type BITMAPINFOHEADER '40 bytes
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type
Private Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

Private Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors As RGBQUAD
End Type

'**********************
'**********************
Dim fByte() As Byte
Dim lByte() As Byte

Dim ByteFrame1() As Byte
Dim ByteFrame2() As Byte
Dim ByteFrame3() As Byte
Dim ByteFrame4() As Byte
Dim ByteFrame5() As Byte
Dim ByteFrame6() As Byte
Dim ByteFrame7() As Byte
Dim ByteFrame8() As Byte
Dim ByteFrame9() As Byte
Dim ByteFrame10() As Byte
Dim ByteFrame11() As Byte
Dim ByteFrame12() As Byte

Dim ByteFrame13() As Byte
Dim ByteFrame14() As Byte
Dim ByteFrame15() As Byte
Dim ByteFrame16() As Byte
Dim ByteFrame17() As Byte
Dim ByteFrame18() As Byte
Dim ByteFrame19() As Byte
Dim ByteFrame20() As Byte
Dim ByteFrame21() As Byte
Dim ByteFrame22() As Byte


Dim FrameCount  As Integer
Dim ByteCount As Long
'**************************
'**************************

Dim lpBI As BITMAPINFO
Dim dw As Integer, dh As Integer
Dim FrameIndex As Integer


Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long




Dim Tip(0 To 6) As String



Private Sub chkLoop_Click()
'Animation Loop Options
If chkLoop.Value = 1 Then
    chkReverseLoop.Enabled = True
Else
    chkReverseLoop.Enabled = False
End If


End Sub



Private Sub cmdBack_Click()
    FrameIndex = FrameIndex - 1
    If FrameIndex < 0 Then FrameIndex = 0
    Picture3.Picture = imgFrame(FrameIndex).Picture

End Sub

Private Sub cmdNext_Click()
    
    FrameIndex = FrameIndex + 1
    If FrameIndex > FrameCount - 1 Then FrameIndex = FrameCount - 1
    Picture3.Picture = imgFrame(FrameIndex).Picture

End Sub

Private Sub cmdPlay_Click()
    Timer1.Interval = 1000 / Val(txtRate.Text)
    Timer1.Enabled = True
     

End Sub

Private Sub cmdStop_Click()
    Timer1.Enabled = False
    cmdStop.Enabled = False
End Sub

Private Sub Make()
    Dim fb As Integer
    Dim lb As Integer
    Dim diff As Integer
    Dim Incr As Single
    Dim frmDC As Long
    Dim fsDc As Long, lsDc As Long 'HDC
    Dim fBm As Long, lBm As Long    'HBITMAP
    Timer1.Enabled = False
    
    SetInitialValues  'Sub function to Initialise all arrays
    imgFrame(0).Picture = Picture1.Picture
    imgFrame(FrameCount - 1).Picture = Picture2.Picture
    fsDc = CreateCompatibleDC(Me.hdc)
    fBm = CreateDIBSection(fsDc, lpBI, DIB_RGB_COLORS, ByVal 0&, ByVal 0&, ByVal 0&)
    Call SelectObject(fsDc, fBm)
    BitBlt fsDc, 0, 0, dw, dh, Picture1.hdc, 0, 0, vbSrcCopy
  
    lsDc = CreateCompatibleDC(Me.hdc)
    lBm = CreateDIBSection(lsDc, lpBI, DIB_RGB_COLORS, ByVal 0&, ByVal 0&, ByVal 0&)
    Call SelectObject(lsDc, lBm)
    BitBlt lsDc, 0, 0, dw, dh, Picture2.hdc, 0, 0, vbSrcCopy
  
    Call GetDIBits(fsDc, fBm, 0, dh, fByte(0), lpBI, 0)
    Call GetDIBits(lsDc, lBm, 0, dh, lByte(0), lpBI, 0)
    
    
    DeleteDC fsDc
    DeleteDC lsDc
    DeleteObject fBm
    DeleteObject lBm
   
    Dim i As Long
    Dim pe As Integer
    For i = 0 To ByteCount
        fb = fByte(i)
        lb = lByte(i)
        diff = lb - fb
        Incr = diff / (FrameCount - 2)
        
        ByteFrame1(i) = fb + Incr * 1
        ByteFrame2(i) = fb + Incr * 2
        ByteFrame3(i) = fb + Incr * 3
        ByteFrame4(i) = fb + Incr * 4
        ByteFrame5(i) = fb + Incr * 5
        ByteFrame6(i) = fb + Incr * 6
        ByteFrame7(i) = fb + Incr * 7
        ByteFrame8(i) = fb + Incr * 8
        ByteFrame9(i) = fb + Incr * 9
        ByteFrame10(i) = fb + Incr * 10
        ByteFrame11(i) = fb + Incr * 11
        ByteFrame12(i) = fb + Incr * 12
        
        ByteFrame13(i) = fb + Incr * 13
        ByteFrame14(i) = fb + Incr * 14
        ByteFrame15(i) = fb + Incr * 15
        ByteFrame16(i) = fb + Incr * 16
        ByteFrame17(i) = fb + Incr * 17
        ByteFrame18(i) = fb + Incr * 18
        ByteFrame19(i) = fb + Incr * 19
        ByteFrame20(i) = fb + Incr * 20
        ByteFrame21(i) = fb + Incr * 21
        ByteFrame22(i) = fb + Incr * 22
        
    Next i
    
    MakeFrames
   

    
End Sub

Private Sub MakeFrames()
    Call StretchDIBits(picBuf.hdc, 0, 0, dw, dh, 0, 0, dw, dh, ByteFrame1(0), lpBI, 0, vbSrcCopy)
    imgFrame(1).Picture = picBuf.Image
    Call StretchDIBits(picBuf.hdc, 0, 0, dw, dh, 0, 0, dw, dh, ByteFrame2(0), lpBI, 0, vbSrcCopy)
    imgFrame(2).Picture = picBuf.Image
    Call StretchDIBits(picBuf.hdc, 0, 0, dw, dh, 0, 0, dw, dh, ByteFrame3(0), lpBI, 0, vbSrcCopy)
    imgFrame(3).Picture = picBuf.Image
    Call StretchDIBits(picBuf.hdc, 0, 0, dw, dh, 0, 0, dw, dh, ByteFrame4(0), lpBI, 0, vbSrcCopy)
    imgFrame(4).Picture = picBuf.Image
    Call StretchDIBits(picBuf.hdc, 0, 0, dw, dh, 0, 0, dw, dh, ByteFrame5(0), lpBI, 0, vbSrcCopy)
    imgFrame(5).Picture = picBuf.Image
    Call StretchDIBits(picBuf.hdc, 0, 0, dw, dh, 0, 0, dw, dh, ByteFrame6(0), lpBI, 0, vbSrcCopy)
    imgFrame(6).Picture = picBuf.Image
    Call StretchDIBits(picBuf.hdc, 0, 0, dw, dh, 0, 0, dw, dh, ByteFrame7(0), lpBI, 0, vbSrcCopy)
    imgFrame(7).Picture = picBuf.Image
    Call StretchDIBits(picBuf.hdc, 0, 0, dw, dh, 0, 0, dw, dh, ByteFrame8(0), lpBI, 0, vbSrcCopy)
    imgFrame(8).Picture = picBuf.Image
    Call StretchDIBits(picBuf.hdc, 0, 0, dw, dh, 0, 0, dw, dh, ByteFrame9(0), lpBI, 0, vbSrcCopy)
    imgFrame(9).Picture = picBuf.Image
    Call StretchDIBits(picBuf.hdc, 0, 0, dw, dh, 0, 0, dw, dh, ByteFrame10(0), lpBI, 0, vbSrcCopy)
    imgFrame(10).Picture = picBuf.Image
    Call StretchDIBits(picBuf.hdc, 0, 0, dw, dh, 0, 0, dw, dh, ByteFrame11(0), lpBI, 0, vbSrcCopy)
    imgFrame(11).Picture = picBuf.Image
    Call StretchDIBits(picBuf.hdc, 0, 0, dw, dh, 0, 0, dw, dh, ByteFrame12(0), lpBI, 0, vbSrcCopy)
    imgFrame(12).Picture = picBuf.Image
    
    Call StretchDIBits(picBuf.hdc, 0, 0, dw, dh, 0, 0, dw, dh, ByteFrame13(0), lpBI, 0, vbSrcCopy)
    imgFrame(13).Picture = picBuf.Image
    Call StretchDIBits(picBuf.hdc, 0, 0, dw, dh, 0, 0, dw, dh, ByteFrame14(0), lpBI, 0, vbSrcCopy)
    imgFrame(14).Picture = picBuf.Image
    Call StretchDIBits(picBuf.hdc, 0, 0, dw, dh, 0, 0, dw, dh, ByteFrame15(0), lpBI, 0, vbSrcCopy)
    imgFrame(15).Picture = picBuf.Image
    Call StretchDIBits(picBuf.hdc, 0, 0, dw, dh, 0, 0, dw, dh, ByteFrame16(0), lpBI, 0, vbSrcCopy)
    imgFrame(16).Picture = picBuf.Image
    Call StretchDIBits(picBuf.hdc, 0, 0, dw, dh, 0, 0, dw, dh, ByteFrame17(0), lpBI, 0, vbSrcCopy)
    imgFrame(17).Picture = picBuf.Image
    Call StretchDIBits(picBuf.hdc, 0, 0, dw, dh, 0, 0, dw, dh, ByteFrame18(0), lpBI, 0, vbSrcCopy)
    imgFrame(18).Picture = picBuf.Image
    Call StretchDIBits(picBuf.hdc, 0, 0, dw, dh, 0, 0, dw, dh, ByteFrame19(0), lpBI, 0, vbSrcCopy)
    imgFrame(19).Picture = picBuf.Image
    Call StretchDIBits(picBuf.hdc, 0, 0, dw, dh, 0, 0, dw, dh, ByteFrame20(0), lpBI, 0, vbSrcCopy)
    imgFrame(20).Picture = picBuf.Image
    Call StretchDIBits(picBuf.hdc, 0, 0, dw, dh, 0, 0, dw, dh, ByteFrame21(0), lpBI, 0, vbSrcCopy)
    imgFrame(21).Picture = picBuf.Image
    Call StretchDIBits(picBuf.hdc, 0, 0, dw, dh, 0, 0, dw, dh, ByteFrame22(0), lpBI, 0, vbSrcCopy)
    imgFrame(22).Picture = picBuf.Image
    DisCardArrays

End Sub

Private Sub DisCardArrays()

    'Clean Up Arrays To Save Memory
    ReDim fByte(0)
    ReDim lByte(0)

    ReDim ByteFrame1(0)
    ReDim ByteFrame2(0)
    ReDim ByteFrame3(0)
    ReDim ByteFrame4(0)
    ReDim ByteFrame5(0)
    ReDim ByteFrame6(0)
    ReDim ByteFrame7(0)
    ReDim ByteFrame8(0)
    ReDim ByteFrame9(0)
    ReDim ByteFrame10(0)
    ReDim ByteFrame11(0)
    ReDim ByteFrame12(0)
    ReDim ByteFrame13(0)
    ReDim ByteFrame14(0)
    ReDim ByteFrame15(0)
    ReDim ByteFrame16(0)
    ReDim ByteFrame17(0)
    ReDim ByteFrame18(0)
    ReDim ByteFrame19(0)
    ReDim ByteFrame20(0)
    ReDim ByteFrame21(0)
    ReDim ByteFrame22(0)
End Sub


Private Sub Command2_Click()
    Timer1.Interval = 1000 / Val(txtRate.Text)
    Timer1.Enabled = True
    cmdStop.Enabled = True
End Sub




Private Sub cmdSave_Click()
    Dim filePathName As String
    Dim fileName As String
    Dim dr As String
    Dim slPos As Integer
    Dim bf As String
    Dim Prompt1 As String, Prompt2 As String, Prompt3 As String, Prompt4 As String
    On Error GoTo Er:
    CD1.Filter = "Bitmap Files (*.bmp)|*.bmp"
    CD1.ShowSave
    filePathName = CD1.fileName
    slPos = InStrRev(filePathName, "\")
    dr = Mid(filePathName, 1, slPos)
    filePathName = Left(filePathName, Len(filePathName) - 4)
    slPos = Len(filePathName) - Len(dr)
    fileName = Right(filePathName, slPos)
    Prompt1 = FrameCount & " Images will be saved to "
    Prompt2 = dr
    Prompt3 = fileName & "0.bmp, " & fileName & "1.bmp, ....." & fileName & "23.bmp"
    Prompt4 = "Any files with same names will overwrite." & vbCrLf & "Do you want to continue.?"
    ret = MsgBox(Prompt1 & vbCrLf & Prompt2 & vbCrLf & Prompt3 & vbCrLf & Prompt4, vbYesNoCancel + vbQuestion + vbDefaultButton1, "Confirm..")
    If ret = vbYes Then
        For i = 0 To FrameCount - 1
            SavePicture imgFrame(i).Picture, filePathName & i & ".bmp"
        Next i
    End If
    
Er:
    
End Sub

Private Sub cmdSwap_Click()
    Set picBuf.Picture = Picture1.Picture
    Set Picture1.Picture = Picture2.Picture
    Set Picture2.Picture = picBuf.Picture
    Me.MousePointer = 11
    Make
    Me.MousePointer = vbDefault
    Set Picture3.Picture = imgFrame(0).Picture
End Sub

Private Sub cmdLImage_Click()
    On Error GoTo Er
    CD1.Filter = "Bitmap Files (*.bmp)|*.bmp|JPG /JPEG- JPEG Files|*.jpg|GIF - Compuserve GIF (*.gif) |*.gif|All Files (*.*)|*.*"
    CD1.ShowOpen
    Me.MousePointer = vbHourglass
    Picture2.Picture = LoadPicture(CD1.fileName)
    Me.Refresh
    Make
    cmdSave.Enabled = True
    Set Picture3.Picture = imgFrame(0).Picture
Er:
    Me.MousePointer = vbDefault
Exit Sub
End Sub


Private Sub cmdFImage_Click()
    On Error GoTo Er
    CD1.Filter = "Bitmap Files (*.bmp)|*.bmp|JPG /JPEG- JPEG Files|*.jpg|GIF - Compuserve GIF (*.gif) |*.gif|All Files (*.*)|*.*"
    CD1.ShowOpen
    Me.MousePointer = 11
    Picture1.Picture = LoadPicture(CD1.fileName)
    Me.Refresh
    Make
    cmdSave.Enabled = True
    Set Picture3.Picture = imgFrame(0).Picture
    
Er:
    Me.MousePointer = vbDefault
   Exit Sub
End Sub



Private Sub cmdTips_Click()
    Static n As Integer
    Label1.Caption = Tip(n)
    n = n + 1
    If n > 6 Then n = 0
End Sub

Private Sub Command1_Click()
    frmAbout.Show vbModal
End Sub

Private Sub Command8_Click()
    Timer1.Enabled = False
    

End Sub

Private Sub Form_Load()
    FrameCount = 24
    For i = 1 To FrameCount - 1
        Load imgFrame(i)
    Next i
    Picture2.Width = Picture1.Width
    Picture2.Height = Picture1.Height
    Picture3.Width = Picture1.Width
    Picture3.Height = Picture1.Height
    picBuf.Width = Picture1.Width
    picBuf.Height = Picture1.Height
    
    ByteCount = Picture1.Width * Picture1.Height * 3 ' Number of total resulting bytes

    dw = Picture1.Width  ' image Width
    dh = Picture1.Height  'image Height
    With lpBI.bmiHeader
        .biBitCount = 24
        .biCompression = 0
        .biWidth = dw
        .biHeight = dh
        .biPlanes = 1
        .biSize = 40
    End With
    
    Tip(0) = "Adjust Image Position such that, Common parts if present should coincide each othor." _
                    & vbCrLf & " eg: Eye, nose, Mouth etc..."
    Tip(1) = "Use Same size of Images."
    Tip(2) = "Adjust to match colors of images." & vbCrLf _
            & " eg: If Start Images is in reddish in color, It is best to have Final Image is also in reddish color."
    Tip(3) = "Adjust the images such that the faces in both images are in same inclination."
    Tip(4) = "Use No Final Image  for a diminishing effect."
    Tip(5) = "Use a White blank image as Final image for a Fading Effect."
    Tip(6) = "Use an External Image Editor Such as Paint Shop Pro for adjusting Images."
    Randomize
    n = 6 * Rnd
    Label1.Caption = Tip(n)
    

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Timer1.Enabled = False
End Sub



Private Sub Timer1_Timer()
    Static flg As Integer
    If FrameIndex = 0 Then
        flg = 1
    End If
    If FrameIndex = FrameCount Then
        flg = -1
        FrameIndex = FrameCount - 1
        If chkLoop.Value <> 1 Then
            FrameIndex = 0
            Timer1.Enabled = False
            Exit Sub
        End If
        If chkReverseLoop.Value <> 1 Then
            FrameIndex = 0
            flg = 1
        End If
    End If
    Set Picture3.Picture = imgFrame(FrameIndex).Picture
    FrameIndex = FrameIndex + flg
     
   
    
End Sub

Private Sub txtNumFrames_Change()
    nFrm = Val(txtNumFrames.Text)
End Sub

Private Sub SetInitialValues()
    ' Initialize Values

    ReDim fByte(0 To ByteCount)
    ReDim lByte(0 To ByteCount)

    ReDim ByteFrame1(0 To ByteCount)
    ReDim ByteFrame2(0 To ByteCount)
    ReDim ByteFrame3(0 To ByteCount)
    ReDim ByteFrame4(0 To ByteCount)
    ReDim ByteFrame5(0 To ByteCount)
    ReDim ByteFrame6(0 To ByteCount)
    ReDim ByteFrame7(0 To ByteCount)
    ReDim ByteFrame8(0 To ByteCount)
    ReDim ByteFrame9(0 To ByteCount)
    ReDim ByteFrame10(0 To ByteCount)
    ReDim ByteFrame11(0 To ByteCount)
    ReDim ByteFrame12(0 To ByteCount)
    
    ReDim ByteFrame13(0 To ByteCount)
    ReDim ByteFrame14(0 To ByteCount)
    ReDim ByteFrame15(0 To ByteCount)
    ReDim ByteFrame16(0 To ByteCount)
    ReDim ByteFrame17(0 To ByteCount)
    ReDim ByteFrame18(0 To ByteCount)
    ReDim ByteFrame19(0 To ByteCount)
    ReDim ByteFrame20(0 To ByteCount)
    ReDim ByteFrame21(0 To ByteCount)
    ReDim ByteFrame22(0 To ByteCount)

End Sub


Private Sub txtRate_Change()
    Timer1.Enabled = False
    If Val(txtRate.Text) < 1 Then txtRate.Text = 1
    Timer1.Interval = 1000 / Val(txtRate.Text)
    Timer1.Enabled = True
End Sub
