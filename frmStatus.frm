VERSION 5.00
Begin VB.Form frmStatus 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4140
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   4140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picProgress 
      Height          =   255
      Left            =   240
      ScaleHeight     =   195
      ScaleWidth      =   3555
      TabIndex        =   1
      Top             =   1080
      Width           =   3615
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      Caption         =   "Total"
      Height          =   195
      Left            =   3480
      TabIndex        =   3
      Top             =   840
      Width           =   420
   End
   Begin VB.Label lblOperation 
      AutoSize        =   -1  'True
      Caption         =   "Operation"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   840
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      Caption         =   "Please stand by..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1725
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
