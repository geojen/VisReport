VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Help Using This Feature"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "In this example, you would enter 2000 as the Start Year, 2002 as the End Year, and 3 as the Number of Rows."
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   1920
      Width           =   4095
   End
   Begin VB.Label Label4 
      Caption         =   "X"
      Height          =   255
      Index           =   8
      Left            =   2640
      TabIndex        =   16
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Label4 
      Caption         =   "X"
      Height          =   255
      Index           =   7
      Left            =   2040
      TabIndex        =   15
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Label4 
      Caption         =   "X"
      Height          =   255
      Index           =   6
      Left            =   1440
      TabIndex        =   14
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Label4 
      Caption         =   "X"
      Height          =   255
      Index           =   5
      Left            =   2640
      TabIndex        =   13
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Label4 
      Caption         =   "X"
      Height          =   255
      Index           =   4
      Left            =   2040
      TabIndex        =   12
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Label4 
      Caption         =   "X"
      Height          =   255
      Index           =   3
      Left            =   1440
      TabIndex        =   11
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Label4 
      Caption         =   "X"
      Height          =   255
      Index           =   2
      Left            =   2640
      TabIndex        =   10
      Top             =   960
      Width           =   135
   End
   Begin VB.Label Label4 
      Caption         =   "X"
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   9
      Top             =   960
      Width           =   135
   End
   Begin VB.Label Label4 
      Caption         =   "X"
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   8
      Top             =   960
      Width           =   135
   End
   Begin VB.Label Label3 
      Caption         =   "Age3"
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   7
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Age2"
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   6
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Age1"
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   5
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "2002"
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   4
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "2001"
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   3
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "2000"
      Height          =   255
      Index           =   0
      Left            =   1320
      TabIndex        =   2
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Your data should be arranged with years in columns and categories in rows. For example:"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Me
End Sub
