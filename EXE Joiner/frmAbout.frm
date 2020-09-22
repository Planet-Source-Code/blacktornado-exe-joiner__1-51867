VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "..:: File Binder Class ::.."
   ClientHeight    =   8505
   ClientLeft      =   -150
   ClientTop       =   240
   ClientWidth     =   10725
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5870.302
   ScaleMode       =   0  'User
   ScaleWidth      =   10071.33
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frame 
      Height          =   7215
      Index           =   3
      Left            =   4920
      TabIndex        =   19
      Top             =   240
      Width           =   5535
      Begin VB.Frame frame 
         Caption         =   " Binder "
         Height          =   2895
         Index           =   5
         Left            =   360
         TabIndex        =   24
         Top             =   3960
         Width           =   4815
         Begin VB.Frame frame 
            Height          =   975
            Index           =   6
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   4455
            Begin VB.Label lblLabel 
               Caption         =   "[Binder] will contain the virus with the loader, it will bind the infected EXE file with the [Loader] which already has the virus."
               Height          =   555
               Index           =   18
               Left            =   120
               TabIndex        =   27
               Top             =   240
               Width           =   4215
            End
         End
         Begin VB.Frame frame 
            Caption         =   " Loader "
            Height          =   735
            Index           =   4
            Left            =   120
            TabIndex        =   25
            Top             =   1320
            Width           =   4455
            Begin VB.Label lblLabel 
               Alignment       =   2  'Center
               Caption         =   "[Loader] will contain the virus binded with the infected EXE file, they are binded using the [Binder]"
               Height          =   435
               Index           =   16
               Left            =   120
               TabIndex        =   28
               Top             =   240
               Width           =   4215
            End
         End
         Begin VB.Label lblLabel 
            Caption         =   $"frmAbout.frx":0ECA
            Height          =   675
            Index           =   17
            Left            =   120
            TabIndex        =   29
            Top             =   2160
            Width           =   4455
         End
      End
      Begin VB.Label lblLabel 
         Caption         =   $"frmAbout.frx":0F59
         Height          =   1035
         Index           =   15
         Left            =   360
         TabIndex        =   23
         Top             =   2160
         Width           =   4815
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Making A Virus"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   360
         TabIndex        =   22
         Top             =   1800
         Width           =   1245
      End
      Begin VB.Label lblLabel 
         Caption         =   $"frmAbout.frx":1088
         Height          =   675
         Index           =   13
         Left            =   360
         TabIndex        =   21
         Top             =   720
         Width           =   4815
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Project Description"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   360
         TabIndex        =   20
         Top             =   360
         Width           =   1620
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   240
      TabIndex        =   14
      Top             =   240
      Width           =   4575
      Begin VB.Label lblVersion 
         Caption         =   "x.x.x"
         Height          =   255
         Left            =   1800
         TabIndex        =   18
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblProjectName 
         Caption         =   "-=[ FileBinder Class ]=-"
         Height          =   255
         Left            =   1800
         TabIndex        =   17
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         Caption         =   "-=[ Version ]=-"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         Caption         =   "-=[ Project ]=-"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4635
      TabIndex        =   13
      Top             =   7920
      Width           =   1455
   End
   Begin VB.Frame frame 
      Height          =   3975
      Index           =   2
      Left            =   240
      TabIndex        =   9
      Top             =   1200
      Width           =   4575
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         Caption         =   $"frmAbout.frx":1123
         Height          =   675
         Index           =   11
         Left            =   240
         TabIndex        =   12
         Top             =   3120
         Width           =   4215
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         Caption         =   "The main purpose of this project is helping you easily bind your files,strings into executables or other files."
         Height          =   1155
         Index           =   10
         Left            =   240
         TabIndex        =   11
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         Caption         =   "This project was completly written by Black Tornado, which is a nick name for: Ahmed Adel Sahib"
         Height          =   1035
         Index           =   9
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   1695
      End
      Begin VB.Image imgLogo 
         Height          =   2475
         Left            =   1920
         Picture         =   "frmAbout.frx":11AD
         ToolTipText     =   "Black Tornado's Logo"
         Top             =   480
         Width           =   2400
      End
   End
   Begin VB.Frame frame 
      Height          =   1215
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   6240
      Width           =   4575
      Begin VB.Label lblEmail 
         Alignment       =   2  'Center
         Caption         =   "blacktornado2k@yahoo.com"
         Height          =   315
         Left            =   1920
         TabIndex        =   5
         ToolTipText     =   "Click to email me"
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         Caption         =   "-=[ Website ]=-"
         Height          =   315
         Index           =   6
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         Caption         =   "-=[ Email ]=-"
         Height          =   315
         Index           =   5
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblWebsite 
         Alignment       =   2  'Center
         Caption         =   "http://www.blacktornado.tk"
         Height          =   315
         Left            =   1920
         TabIndex        =   2
         ToolTipText     =   "Click to visit"
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame frame 
      Height          =   615
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   5280
      Width           =   4575
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         Caption         =   "-=[ Nothing Is Impossible ]=-"
         Height          =   315
         Index           =   8
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.Label lblLabel 
      Alignment       =   2  'Center
      Caption         =   "Made in Iraq"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   255
      TabIndex        =   6
      ToolTipText     =   "This project was made in Baghdad, Iraq"
      Top             =   7560
      Width           =   10215
   End
   Begin VB.Label lblLabel 
      Alignment       =   2  'Center
      Caption         =   "For more information, feel free to contact me using :"
      Height          =   315
      Index           =   2
      Left            =   240
      TabIndex        =   0
      Top             =   6000
      Width           =   4575
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
lblVersion.Caption = App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub lblEmail_Click()
Shell "C:\Program Files\Internet Explorer\Iexplore.exe mailto:blacktornado2k@yahoo.com", vbHide
End Sub

Private Sub lblWebsite_Click()
Shell "C:\Program Files\Internet Explorer\Iexplore.exe http://www.blacktornado.tk", vbNormalFocus
End Sub
