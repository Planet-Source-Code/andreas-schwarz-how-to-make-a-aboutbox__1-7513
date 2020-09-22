VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About..."
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5580
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   301
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   372
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton blCommand 
      Caption         =   "More"
      Height          =   330
      Index           =   1
      Left            =   3990
      TabIndex        =   10
      Top             =   3675
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.CommandButton blCommand 
      Caption         =   "Ok"
      Height          =   330
      Index           =   0
      Left            =   3990
      TabIndex        =   9
      Top             =   3255
      Width           =   1380
   End
   Begin VB.PictureBox lpTemp 
      Height          =   855
      Left            =   1470
      ScaleHeight     =   795
      ScaleWidth      =   3840
      TabIndex        =   6
      Top             =   2100
      Width           =   3900
      Begin VB.Label uIfoName 
         AutoSize        =   -1  'True
         Caption         =   "FutureProjects"
         Height          =   195
         Left            =   105
         TabIndex        =   8
         Top             =   105
         Width           =   1065
      End
      Begin VB.Label uIfoSerial 
         AutoSize        =   -1  'True
         Caption         =   "Serial-number: xxx-xx-xxx-xxx"
         Height          =   195
         Left            =   105
         TabIndex        =   7
         Top             =   420
         Width           =   2265
      End
   End
   Begin VB.Line Line2 
      X1              =   14
      X2              =   357
      Y1              =   301
      Y2              =   301
   End
   Begin VB.Label lbTemp 
      Caption         =   $"frmAbout.frx":000C
      Height          =   1275
      Index           =   1
      Left            =   210
      TabIndex        =   11
      Top             =   3150
      Width           =   3690
   End
   Begin VB.Label lbTemp 
      AutoSize        =   -1  'True
      Caption         =   "This product is licensed to:"
      Height          =   195
      Index           =   0
      Left            =   1470
      TabIndex        =   5
      Top             =   1785
      Width           =   1905
   End
   Begin VB.Label lbAInfo 
      AutoSize        =   -1  'True
      Caption         =   "written and developed by Andreas Schwarz"
      Height          =   195
      Left            =   1470
      TabIndex        =   4
      Top             =   1365
      Width           =   3135
   End
   Begin VB.Label lbAppVersion 
      AutoSize        =   -1  'True
      Caption         =   "Version"
      Height          =   195
      Left            =   1470
      TabIndex        =   3
      Top             =   840
      Width           =   525
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   14
      X2              =   357
      Y1              =   204
      Y2              =   204
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   14
      X2              =   357
      Y1              =   203
      Y2              =   203
   End
   Begin VB.Label lbAppCopyright 
      AutoSize        =   -1  'True
      Caption         =   "Copyright © 1999 FutureProjects Development"
      Height          =   195
      Left            =   1470
      TabIndex        =   2
      Top             =   1155
      Width           =   3405
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "for 32Bit Database Development"
      Height          =   195
      Left            =   1470
      TabIndex        =   1
      Top             =   525
      Width           =   2340
   End
   Begin VB.Label lbAppTitle 
      AutoSize        =   -1  'True
      Caption         =   "D/4F Database File Management"
      Height          =   195
      Left            =   1470
      TabIndex        =   0
      Top             =   315
      Width           =   2340
   End
   Begin VB.Image lpAppImage 
      Height          =   2730
      Left            =   210
      Picture         =   "frmAbout.frx":012D
      Top             =   210
      Width           =   1050
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
lbAppVersion.Caption = "Version " + AppVersion
Height = 4815

End Sub

Public Sub DisplayAboutBox()
    lbAppTitle.Caption = App.FileDescription
    lbAppCopyright.Caption = App.LegalCopyright
End Sub

Function AppVersion() As String
    AppVersion = Trim(Trim(Str(App.Major)) + "." + Trim(Str(App.Minor)) + "." + Trim(Str(App.Revision)))
End Function
