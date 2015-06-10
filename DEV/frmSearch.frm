VERSION 5.00
Begin VB.Form frmSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Data Collection"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9945
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   108
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   9945
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9735
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   375
         Left            =   8160
         TabIndex        =   3
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Find Next"
         Default         =   -1  'True
         Height          =   375
         Left            =   6600
         TabIndex        =   2
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtSearch 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Top             =   1440
         Width           =   5295
      End
      Begin VB.Label Label5 
         Caption         =   "Data:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblLine 
         Caption         =   "0"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1440
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Find what:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblData 
         ForeColor       =   &H00FF0000&
         Height          =   805
         Left            =   1440
         TabIndex        =   5
         Top             =   480
         Width           =   8175
      End
      Begin VB.Label Label1 
         Caption         =   "Current Line:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdSearch_Click()
    SearchCollection
End Sub

Private Sub Form_Activate()
    Dim G As MSFlexGrid
    Dim S As String
    Dim J As Integer
    
    Set G = frmData.MSFlexGrid1
    If G.ColSel < G.Cols - 1 Then
        lblLine.Caption = "0"
        lblData.Caption = ""
        JList = 1
    Else
        S = GetDataGridDesc(G.Row) & vbCrLf
        S = S & GetDataGridData(G.Row)
        lblLine.Caption = CStr(G.Row)
        lblData.Caption = S
        If G.Row = G.Rows - 1 Then
            JList = 1
        Else
            JList = G.Row + 1
        End If
    End If
    
    Me.Top = 0
    Me.Left = (Screen.Width - Me.ScaleWidth) \ 2
    
End Sub

