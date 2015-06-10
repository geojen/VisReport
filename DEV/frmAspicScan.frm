VERSION 5.00
Begin VB.Form frmAspicScan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ASPIC Scan Summary"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9570
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
   Icon            =   "frmAspicScan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   9570
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9255
      Begin VB.Frame Frame2 
         Caption         =   "How Do You Want to Handle Missing Data?"
         Height          =   1575
         Left            =   2520
         TabIndex        =   15
         Top             =   2280
         Width           =   6495
         Begin VB.TextBox txtMissing 
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   600
            TabIndex        =   18
            Top             =   1080
            Width           =   1695
         End
         Begin VB.OptionButton opMissing 
            Caption         =   "Replace All Missing Values With:"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   17
            Top             =   840
            Width           =   3375
         End
         Begin VB.OptionButton opMissing 
            Caption         =   "Do Not Change"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   16
            Top             =   360
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "In Input: Negative values are missing data. In Output: Missing value are zeros."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   19
            Top             =   600
            Width           =   5775
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   5640
         TabIndex        =   14
         Top             =   4200
         Width           =   1575
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   2040
         TabIndex        =   13
         Top             =   4200
         Width           =   1575
      End
      Begin VB.TextBox txtCase 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2520
         TabIndex        =   1
         Top             =   600
         Width           =   6495
      End
      Begin VB.Label lblModelType 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2520
         TabIndex        =   12
         Top             =   960
         Width           =   6495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Model Type"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ASPIC Input File Selected"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   2220
      End
      Begin VB.Label lblFile 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2520
         TabIndex        =   9
         Top             =   240
         Width           =   6495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Case Description"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Start Year"
         Height          =   195
         Index           =   2
         Left            =   2520
         TabIndex        =   7
         Top             =   1320
         Width           =   870
      End
      Begin VB.Label lblYear 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   3600
         TabIndex        =   6
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "End Year"
         Height          =   195
         Index           =   3
         Left            =   6600
         TabIndex        =   5
         Top             =   1320
         Width           =   795
      End
      Begin VB.Label lblYear 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   7680
         TabIndex        =   4
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Number of Data Series"
         Height          =   195
         Index           =   5
         Left            =   2520
         TabIndex        =   3
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label lblSeries 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4680
         TabIndex        =   2
         Top             =   1800
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmAspicScan"
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
    
    ScanAspicResults
End Sub

Private Sub Form_Load()
txtMissing.Text = NAVal
End Sub

Private Sub opMissing_Click(Index As Integer)
If opMissing(0).Value = True Then
    txtMissing.Enabled = False
Else
    txtMissing.Enabled = True
End If
End Sub
