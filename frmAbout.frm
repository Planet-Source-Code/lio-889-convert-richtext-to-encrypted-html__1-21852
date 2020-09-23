VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5415
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDisclaimer 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   1800
      Width           =   5175
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Okay!"
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Image imgTitle 
      Height          =   480
      Left            =   120
      Picture         =   "frmAbout.frx":0000
      Top             =   80
      Width           =   480
   End
   Begin VB.Label lblHomePage 
      AutoSize        =   -1  'True
      Caption         =   "http://www.geocities.com/lio889"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1560
      TabIndex        =   6
      Top             =   1440
      Width           =   2805
   End
   Begin VB.Label lblEmail 
      AutoSize        =   -1  'True
      Caption         =   "lio_889@ziplip.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1560
      TabIndex        =   5
      Top             =   1200
      Width           =   1605
   End
   Begin VB.Label lblHomePageCap 
      AutoSize        =   -1  'True
      Caption         =   "Home Page:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1050
   End
   Begin VB.Label lblEmailCap 
      AutoSize        =   -1  'True
      Caption         =   "E-mail address:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1350
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Version 1.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0FEFE&
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   360
      Width           =   1170
   End
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
      Caption         =   "Copyright © 2001 Khaery Rida"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HTML Messages Encoder"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0FEFE&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5415
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    frmAbout.Hide
    
End Sub

Private Sub Form_Load()
    txtDisclaimer.Text = ""
    txtDisclaimer.Text = txtDisclaimer.Text & "This computer program is FREEWARE. You may copy this program and use it in any way you may find it useful." & vbCrLf
    txtDisclaimer.Text = txtDisclaimer.Text & "However, you may NOT repost modifications or include it in your programs without the permission of author." & vbCrLf & vbCrLf
    txtDisclaimer.Text = txtDisclaimer.Text & "Please note that the use of this program is subject to the following conditions:" & vbCrLf
    txtDisclaimer.Text = txtDisclaimer.Text & "(1) The author can NOT be held responsibility for any damage and/or loss of data of any kind caused by this program. "
    txtDisclaimer.Text = txtDisclaimer.Text & "USE IT AT YOUR OWN RISK!" & vbCrLf
    txtDisclaimer.Text = txtDisclaimer.Text & "(2) It is your responsibility to comply with local of federal laws regarding the use of this program."

End Sub

Private Sub lblEmail_Click()
    iRet = Shell("start.exe mailto:lio_889@ziplip.com?subject=HTMLmsgEnc", vbNormal)
End Sub

Private Sub lblHomePage_Click()
    iRet = Shell("start.exe http://www.geocities.com/lio889/", vbNormal)
End Sub
