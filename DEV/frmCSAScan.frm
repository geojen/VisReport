VERSION 5.00
Begin VB.Form frmCSAScan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CSA Scan Summary"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   105
   Icon            =   "frmCSAScan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   9510
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9255
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   5280
         TabIndex        =   12
         Top             =   2040
         Width           =   1575
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   1800
         TabIndex        =   11
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox txtCase 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2400
         TabIndex        =   1
         Top             =   600
         Width           =   6615
      End
      Begin VB.Label lblYear 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   7680
         TabIndex        =   10
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "End Year"
         Height          =   195
         Index           =   3
         Left            =   6600
         TabIndex        =   9
         Top             =   1320
         Width           =   795
      End
      Begin VB.Label lblYear 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   3480
         TabIndex        =   8
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Start Year"
         Height          =   195
         Index           =   2
         Left            =   2400
         TabIndex        =   7
         Top             =   1320
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Case Description"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblFile 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2400
         TabIndex        =   5
         Top             =   240
         Width           =   6615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CSA Input File Selected"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2040
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Model Type"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1005
      End
      Begin VB.Label lblModelType 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2400
         TabIndex        =   2
         Top             =   960
         Width           =   6615
      End
   End
End
Attribute VB_Name = "frmCSAScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Close #5
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim S As String
Dim N As Integer

    'check for commas in the case description
    S = txtCase.Text
    N = InStr(S, ",")
    If N > 0 Then
        MsgBox "Case Description Cannot Contain Commas. Please Fix.", vbInformation, "Visual Report Designer"
        Exit Sub
    End If
    
    ScanCSAResults
End Sub
