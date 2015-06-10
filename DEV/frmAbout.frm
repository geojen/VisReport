VERSION 5.00
Begin VB.Form frmAbout 
   Caption         =   "About Visual Report Designer"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4770
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   4770
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   1960
         Picture         =   "frmAbout.frx":0442
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   2
         Top             =   2040
         Width           =   480
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   1
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "NOAA Fisheries Toolbox"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   720
         TabIndex        =   6
         Top             =   240
         Width           =   2925
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Visual Report Designer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   360
         TabIndex        =   5
         Top             =   840
         Width           =   3705
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "April 2008"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   1320
         TabIndex        =   4
         Top             =   2760
         Width           =   1785
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Version 1.6.1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   1440
         TabIndex        =   3
         Top             =   1680
         Width           =   1665
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
Unload Me
End Sub
