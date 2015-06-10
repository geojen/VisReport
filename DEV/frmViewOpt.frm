VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmViewOpt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Viewing Options"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10245
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   116
   Icon            =   "frmViewOpt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   10245
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Save"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   2760
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8760
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   360
      TabIndex        =   4
      Top             =   240
      Width           =   9615
      Begin VB.Frame framView 
         Height          =   735
         Left            =   480
         TabIndex        =   6
         Top             =   1200
         Width           =   8775
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "Browse"
            Height          =   255
            Left            =   7560
            TabIndex        =   7
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lblFile 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "None Selected"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   7215
         End
      End
      Begin VB.OptionButton optView 
         Caption         =   "Use My Computer's Default Web Browser"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   4575
      End
      Begin VB.OptionButton optView 
         Caption         =   "Select Another Program"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   1
         Top             =   960
         Width           =   3255
      End
      Begin VB.OptionButton optView 
         Caption         =   "Use Visual Report Designer's Report Viewer"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   5055
      End
   End
End
Attribute VB_Name = "frmViewOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBrowse_Click()
    OpenEditSelect
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim T As String
    
    RptViewer = "GUI"
    If optView(1).Value = True Then
        T = lblFile.Caption
        If Dir(T) <> "" Then
            RptViewer = T
        End If
    ElseIf optView(2).Value = True Then
        RptViewer = "Browser"
    End If
    
    WriteCfg
    
    Unload Me
End Sub

Private Sub Form_Load()
    If RptViewer = "GUI" Then
        optView(0).Value = True
    ElseIf RptViewer = "Browser" Then
        optView(2).Value = True
    Else
        optView(1).Value = True
        lblFile.Caption = RptViewer
    End If
    
End Sub

Private Sub optView_Click(Index As Integer)
    If Index = 1 Then
        framView.Visible = True
    Else
        framView.Visible = False
    End If
End Sub
