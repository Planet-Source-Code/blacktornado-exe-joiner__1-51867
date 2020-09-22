VERSION 5.00
Begin VB.Form frmBinder 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "..:: EXE Joiner : An example of FileBinder usage ::.."
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6705
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBinder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frame 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6135
      Begin VB.Frame frame 
         BackColor       =   &H00000000&
         Height          =   975
         Index           =   1
         Left            =   360
         TabIndex        =   7
         Top             =   1800
         Width           =   5415
         Begin VB.CommandButton cmdClose 
            Caption         =   "Close"
            Height          =   375
            Left            =   3480
            TabIndex        =   10
            ToolTipText     =   "Click to close the program"
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton cmdAbout 
            Caption         =   "About"
            Height          =   375
            Left            =   1920
            TabIndex        =   9
            ToolTipText     =   "Click to see more information about FileBinder"
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton cmdCompile 
            Caption         =   "Compile"
            Height          =   375
            Left            =   360
            TabIndex        =   8
            ToolTipText     =   "Click to bind the two executables into one file"
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.TextBox txtTarget 
         BackColor       =   &H00404040&
         ForeColor       =   &H00849D9F&
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Text            =   "C:\Virus.exe"
         Top             =   1320
         Width           =   4335
      End
      Begin VB.TextBox txtEXE2 
         BackColor       =   &H00404040&
         ForeColor       =   &H00849D9F&
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Text            =   "C:\Windows\Notepad.exe"
         Top             =   720
         Width           =   4335
      End
      Begin VB.TextBox txtEXE1 
         BackColor       =   &H00404040&
         ForeColor       =   &H00849D9F&
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Text            =   "C:\Windows\Regedit.exe"
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Target File:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H003FA3BC&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Second EXE:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H003FA3BC&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "First EXE:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H003FA3BC&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Label lblMadeIN 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Made in Iraq"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H003FA3BC&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   3480
      Width           =   6135
   End
End
Attribute VB_Name = "frmBinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MyFBF As New clsFileBinder

Private Sub cmdAbout_Click()
MyFBF.About
End Sub

Private Sub cmdClose_Click()
End
End Sub

Private Sub cmdCompile_Click()
On Error Resume Next
MyFBF.FileToProperty txtEXE1, "EXE1"
MyFBF.FileToProperty txtEXE2, "EXE2"
FileCopy App.Path & "\Loader.exe", txtTarget
MyFBF.SavePackage txtTarget
MsgBox "Done! Click OK to test the file.", vbInformation
Shell txtTarget, vbNormalFocus
End
End Sub
