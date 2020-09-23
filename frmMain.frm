VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "HTML Messages Encoder"
   ClientHeight    =   5925
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10680
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   10680
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlg 
      Left            =   8400
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox txtM 
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   2990
      _Version        =   393217
      HideSelection   =   0   'False
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMain.frx":0CCA
   End
   Begin VB.Frame frameMenu3 
      Height          =   25
      Left            =   -120
      TabIndex        =   14
      Top             =   1750
      Width           =   9015
   End
   Begin VB.ComboBox lstFontColor 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4320
      Style           =   2  'Dropdown List
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1335
   End
   Begin VB.ComboBox lstFontSize 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1335
   End
   Begin VB.ComboBox lstFontName 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmMain.frx":0DB0
      Left            =   120
      List            =   "frmMain.frx":0DB2
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1320
      Width           =   2655
   End
   Begin VB.CommandButton cmdStrikethru 
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   11
      ToolTipText     =   "Strikeout"
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton cmdUnderline 
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   10
      ToolTipText     =   "Underline"
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton cmdItalic 
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   9
      ToolTipText     =   "Italic"
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton cmdBold 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   8
      ToolTipText     =   "Bold"
      Top             =   1320
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   2040
      TabIndex        =   13
      Top             =   120
      Width           =   25
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export Message to Encrypted HTML..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      Picture         =   "frmMain.frx":0DB4
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.Frame frameMenu2 
      Height          =   25
      Left            =   -120
      TabIndex        =   12
      Top             =   1200
      Width           =   9015
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Message..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3360
      Picture         =   "frmMain.frx":1EFE
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2160
      Picture         =   "frmMain.frx":2D34
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame frameMenu1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   25
      Left            =   -120
      TabIndex        =   0
      Top             =   0
      Width           =   9015
   End
   Begin VB.TextBox txtOpen 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   17
      Top             =   3960
      Width           =   8775
   End
   Begin VB.Label lblM 
      AutoSize        =   -1  'True
      Caption         =   "Message (RichText):"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   120
      TabIndex        =   15
      Top             =   1800
      Width           =   1845
   End
   Begin VB.Label lblOpen 
      AutoSize        =   -1  'True
      Caption         =   "Opening Message (PlainText):"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   120
      TabIndex        =   16
      Top             =   3720
      Width           =   2715
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExport 
         Caption         =   "&Export Message to Encrypted HTML..."
      End
      Begin VB.Menu Null0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Sa&ve Message as..."
      End
      Begin VB.Menu Null1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuSelection 
      Caption         =   "&Selection"
      Begin VB.Menu mnuBold 
         Caption         =   "Make &Bold"
      End
      Begin VB.Menu mnuItalic 
         Caption         =   "Make &Italic"
      End
      Begin VB.Menu mnuUnderline 
         Caption         =   "&Underline"
      End
      Begin VB.Menu mnuStrikethru 
         Caption         =   "&Strikeout"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About..."
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBold_Click()
    Call mnuBold_Click
    
End Sub

Private Sub cmdExport_Click()
    Call mnuExport_Click
    
End Sub

Private Sub cmdItalic_Click()
    Call mnuItalic_Click

End Sub

Private Sub cmdOpen_Click()
    Call mnuOpen_Click
    
End Sub

Private Sub cmdSave_Click()
    Call mnuSaveAs_Click
    
End Sub

Private Sub cmdStrikethru_Click()
    Call mnuStrikethru_Click
    
End Sub

Private Sub cmdUnderline_Click()
    Call mnuUnderline_Click

End Sub

Private Sub Form_Load()
    
    For i% = 0 To Screen.FontCount - 1
        lstFontName.AddItem Screen.Fonts(i%)
    Next
    
    lstFontSize.AddItem "8"
    lstFontSize.AddItem "10"
    lstFontSize.AddItem "12"
    lstFontSize.AddItem "14"
    lstFontSize.AddItem "18"
    lstFontSize.AddItem "24"
    lstFontSize.AddItem "36"
    
    If FontNameIndex% > Screen.FontCount - 1 Then
        FontNameIndex% = 0
    End If
    
    lstFontColor.AddItem "Black"
    lstFontColor.AddItem "Blue"
    lstFontColor.AddItem "Red"
    lstFontColor.AddItem "Green"
    lstFontColor.AddItem "Magenta"
    
defFontSet:
    lstFontSize.ListIndex = 1
    lstFontColor.ListIndex = 0
    lstFontName.ListIndex = 0
    txtM.SelFontName = lstFontName.List(lstFontName.ListIndex)
    QT = Chr(34)

End Sub

Private Sub Form_Resize()

    On Error Resume Next
    frameMenu1.Width = frmMain.Width + 30
    On Error Resume Next
    frameMenu2.Width = frmMain.Width + 30
    On Error Resume Next
    frameMenu3.Width = frmMain.Width + 30
    On Error Resume Next
    txtOpen.Top = frmMain.Height - 1710
    On Error Resume Next
    lblOpen.Top = frmMain.Height - 1950
    On Error Resume Next
    txtM.Height = frmMain.Height - 3975
    On Error Resume Next
    txtM.Width = frmMain.Width - 315
    On Error Resume Next
    txtOpen.Width = frmMain.Width - 315
    If txtM.Height <= 75 Then lblOpen.Visible = False: lblM.Visible = False: txtOpen.Visible = False Else lblM.Visible = True: lblOpen.Visible = True: txtOpen.Visible = True
    Exit Sub
    
End Sub

Private Sub lstFontColor_Click()
    
    Select Case lstFontColor.ListIndex
        Case 0
            Call ApplyFormat(3, vbBlack)
        Case 1
            Call ApplyFormat(3, vbBlue)
        Case 2
            Call ApplyFormat(3, vbRed)
        Case 3
            Call ApplyFormat(3, vbGreen)
        Case 4
            Call ApplyFormat(3, vbMagenta)
    End Select

End Sub

Private Sub lstFontName_Click()
    Call ApplyFormat(1, lstFontName.List(lstFontName.ListIndex))

End Sub

Private Sub lstFontSize_Click()
    Call ApplyFormat(2, lstFontSize.List(lstFontSize.ListIndex))

End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show 1
End Sub

Private Sub mnuBold_Click()
    Call ApplyFormat(4)
    
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Public Sub ApplyFormat(FormatIndex As Long, Optional addFormat As String)
    
    Select Case FormatIndex
    Case 1
        txtM.SelFontName = addFormat
    Case 2
        txtM.SelFontSize = Int(addFormat)
    Case 3
        txtM.SelColor = addFormat
    Case 4
        txtM.SelBold = Not (txtM.SelBold)
    Case 5
        txtM.SelItalic = Not (txtM.SelItalic)
    Case 6
        txtM.SelUnderline = Not (txtM.SelUnderline)
    Case 7
        txtM.SelStrikeThru = Not (txtM.SelStrikeThru)
    End Select
    
    On Error Resume Next
    txtM.SetFocus
    Exit Sub

End Sub

Private Sub mnuExport_Click()
        
    Dim sourceHTML As String, destHTML As String
    Dim Keyword As String, arCode() As String
    Dim script As String
    
    If Trim(txtM.Text) = "" Then
        MsgBox "Nothing to export!", 48, MainTitle
        On Error Resume Next
        txtM.SetFocus
        Exit Sub
    End If
    
    destFile = FilePath("Hyper Text Markup Language (*.html)|*.html", "Export")
    If destFile = "" Then Exit Sub
    
    Keyword = InputBox("Please type the Password (Case Sensitive) you'd like to encrypt your message with (16 characters max)." & vbCrLf & "This Password will be needed when decrypting your message.", MainTitle)
    
    If Keyword = "" Then
        MsgBox "Can not proceed with out Password!", 16, MainTitle
        Exit Sub
    ElseIf Len(Keyword) > 16 Then
        MsgBox "The Password should consist of a maximum number of 16 characters.", 16, MainTitle
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    Call InitStatus
    LockWindowUpdate Me.hwnd
    frmStatus.lblOperation.Caption = "Converting RichText to HTML..."
    frmStatus.lblOperation.Refresh
    frmStatus.lblTotal.Caption = ""
    frmStatus.Show
    frmStatus.Refresh
    frmMain.Refresh
    Call UpdateProgress(1, Len(txtM.Text), 1)
    sourceHTML = HTMLCompile("System", 1, vbBlack)
    frmStatus.lblOperation.Caption = "Encrypting message..."
    frmStatus.lblOperation.Refresh
    Call UpdateProgress(1, Len(sourceHTML), 1)
    destHTML = crypt(sourceHTML, Keyword, Len(txtM.Text))
    frmStatus.Hide
    
    Call Export(destFile, destHTML)
    LockWindowUpdate 0 'free any window lock
    Screen.MousePointer = 0

End Sub

Private Sub mnuItalic_Click()
    Call ApplyFormat(5)

End Sub

Private Sub mnuOpen_Click()
    dlg.FileName = ""
    dlg.DialogTitle = "Open File"
    dlg.Filter = "Rich Text Format (*.rtf)|*.rtf|Plain-Text (*.txt)|*.txt"
    dlg.ShowOpen
    If dlg.FileName = "" Then Exit Sub
    
    If Not FileExists(dlg.FileName + "") Then
        MsgBox "Can not open file " & LCase$(dlg.FileName) & ". Please make sure that file exists and not being used by another application.", 48, "Error openning file"
        Exit Sub
    End If
    
    Select Case UCase(Right$(dlg.FileName, 3))
    Case "TXT"
        txtM.LoadFile dlg.FileName, 1
    Case "RTF"
        txtM.LoadFile dlg.FileName, 0
    Case Else
        MsgBox "Unrecognized file format.", 48, MainTitle
        Exit Sub
    End Select
    txtM.SelStart = Len(txtM.Text)
    txtM.SetFocus
    
End Sub

Private Sub mnuSaveAs_Click()
    
    If Trim(txtM.Text) = "" Then
        MsgBox "Nothing to save!", 48, MainTitle
        Exit Sub
    End If
        
    destFile = FilePath(AllFilter, "Save Message")
    If destFile = "" Then Exit Sub
    
    Screen.MousePointer = 11
    Select Case UCase$(Right$(dlg.FileName, 4))
    Case ".TXT"
        txtM.SaveFile dlg.FileName, 1
    Case ".RTF"
        txtM.SaveFile dlg.FileName, 0
    Case "HTML"
        Screen.MousePointer = 11
        Call InitStatus
        LockWindowUpdate Me.hwnd
        frmStatus.lblOperation.Caption = "Converting RichText to HTML..."
        frmStatus.lblTotal.Caption = ""
        frmStatus.Show
        frmStatus.Refresh
        frmMain.Refresh
        Open dlg.FileName For Output As #1
        Print #1, "<HTML><HEAD><TITLE>" & Title & "</TITLE>" & "</HEAD>"
        Print #1, "<BODY>"
        Print #1, ""
        Print #1, HTMLCompile("System", 1, vbBlack)
        Print #1, ""
        Print #1, "</BODY></HTML>"
        Close #1
        frmStatus.Hide
        Screen.MousePointer = 0
        frmStatus.Hide
        
    Case Else
        Screen.MousePointer = 0
        MsgBox "Invalid file format.", 48, MainTitle
        Exit Sub
    End Select
    LockWindowUpdate 0
    Screen.MousePointer = 0
    
End Sub

Private Sub mnuStrikethru_Click()
    Call ApplyFormat(7)
    
End Sub

Private Sub mnuUnderline_Click()
    Call ApplyFormat(6)

End Sub

Private Sub txtM_Change()
    If txtM.Text = "" Then txtM.SelFontName = lstFontName.List(lstFontName.ListIndex)
    
End Sub


Public Sub Export(FileName As String, EncryptedData As String)
    
    Dim arData() As String, arOpen() As String
    Dim curElement As Long, blockLen As Long
    Dim lastEPos As Long, OpenMessage As String
    Dim elementCount As Long
    
    
    curElement = 0
    blockLen = 0
    lastEPos = 1
    ReDim arData(0)
    
    For currentCounter = 1 To Len(EncryptedData) Step 2
        blockLen = blockLen + 2
        If blockLen = 70 Then
            arData(curElement) = Mid$(EncryptedData, lastEPos, 70)
            ReDim Preserve arData(UBound(arData) + 1)
            curElement = curElement + 1
            blockLen = 0
            lastEPos = lastEPos + 70
        ElseIf currentCounter >= Len(EncryptedData) - 1 Then
            arData(curElement) = Mid$(EncryptedData, lastEPos)
        End If
    Next currentCounter

    arOpen = Split(txtOpen.Text, vbCrLf)
    For currentCounter = LBound(arOpen) To UBound(arOpen)
        OpenMessage = OpenMessage & HTMLString(arOpen(currentCounter)) & "<BR>"
    Next currentCounter
    
    ' Open and write to file
        Open FileName For Output As #1
        Print #1, "<HTML><HEAD><TITLE>" & Title & "</TITLE>" & "</HEAD>"
        Print #1, "<BODY><CENTER>"
        Print #1, "<TABLE width=430 border=0 cellspacing=10 cellpadding=10>"
        Print #1, "<TR>"
        Print #1, "<TD><FONT face=" & QT & "Verdana" & QT & " size=1>" & OpenMessage & "</FONT></TD></TR>"
        Print #1, "<TR>"
        Print #1, "<TD align=center>"
        Print #1, "<INPUT type=" & QT & "password" & QT & " name=" & QT & "txtPassword" & QT & ">"
        Print #1, "<INPUT type=" & QT & "button" & QT & " value=" & QT & "Decrypt Message" & QT & " name=" & QT & "cmdDecrypt" & QT & ">"
        Print #1, "</TD></TR></TABLE>"
        Print #1, "</CENTER>"
        Print #1, " "
    
    ' Open script's code file
        Open App.Path & "\Code.txt" For Input As #2
        curElement = 0
        Do While Not EOF(2)
            Line Input #2, script
            If Trim(script) = "//CUSTOM1//" Then
                Print #1, "Public arData(" & UBound(arData) & ")"
            ElseIf Trim(script) = "//CUSTOM2//" Then
                For currentCounter = LBound(arData) To UBound(arData)
                    Print #1, "arData(" & curElement & ") = " & QT & arData(currentCounter) & QT
                    curElement = curElement + 1
                Next currentCounter
            ElseIf Trim(Left(script, 4)) = "<!--" Then ' Ignore comments in Code.txt
            Else
                Print #1, script
            End If
        Loop
        Print #1, "</BODY></HTML>"
        Close #1
        Close #2
        Screen.MousePointer = 0
        Exit Sub
ErrOpenSourceCode:
        Screen.MousePointer = 0
        MsgBox "Can not open " & UCase$(App.Path) & "\CODE.TXT.", 16, MainTitle
        Exit Sub
End Sub

Public Function FilePath(dlgFilter As String, dlgTitle As String) As String
    
    dlg.FileName = ""
    dlg.Filter = dlgFilter
    dlg.DialogTitle = dlgTitle
    dlg.ShowSave

    If dlg.FileName = "" Then FilePath = ""
    If Len(Dir(dlg.FileName)) > 0 Then
        
        Msg = "The file " & UCase$(dlg.FileName) & " already exists." & vbCrLf & "Do you want to replace it?"
        dgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2

        Response = MsgBox(Msg, dgDef, Title)
        If Response = IDYES Then
            If Not FileExists(dlg.FileName + "") Then
                MsgBox "Unable to over-write the file " & UCase$(dlg.FileName) & "." & vbCrLf & "Please make sure that the file is NOT being used by another application.", 16, MainTitle
                Exit Function
            End If
            Kill dlg.FileName
        Else
            Exit Function
        End If
    End If
    
    FilePath = dlg.FileName
    
End Function

Public Sub InitStatus()

On Error Resume Next
Unload frmStatus
Load frmStatus
SetWindowPos frmStatus.hwnd, HWND_TOPMOST, frmStatus.Left / 15, _
    frmStatus.Top / 15, frmStatus.Width / 15, _
    frmStatus.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW

End Sub
