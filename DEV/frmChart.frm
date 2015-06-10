VERSION 5.00
Object = "{E9DF30CA-4B30-4235-BF0C-7150F6466080}#1.0#0"; "ChartFX.ClientServer.Core.dll"
Begin VB.Form frmChart 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chart of Selected Data"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10650
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   110
   Icon            =   "frmChart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   10650
   StartUpPosition =   1  'CenterOwner
   Begin Cfx62ClientServerCtl.Chart Chart1 
      Height          =   6975
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   10335
      _Data_          =   "frmChart.frx":0442
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<  Previous"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next  >"
      Default         =   -1  'True
      Height          =   375
      Left            =   7200
      TabIndex        =   1
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   7200
      Width           =   1455
   End
End
Attribute VB_Name = "frmChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdNext_Click()
Dim G As MSFlexGrid
Dim N As Integer

Set G = frmData.MSFlexGrid1
N = G.Row + 1
If N > G.Rows - 1 Then N = 1
G.Row = N
G.Col = 0
G.ColSel = G.Cols - 1
If G.RowIsVisible(N) = False Then G.TopRow = N 'show row if out of view

ShowChart
End Sub

Private Sub cmdPrevious_Click()
Dim G As MSFlexGrid
Dim N As Integer

Set G = frmData.MSFlexGrid1
N = G.Row - 1
If N < 1 Then N = G.Rows - 1
G.Row = N
G.Col = 0
G.ColSel = G.Cols - 1
If G.RowIsVisible(N) = False Then G.TopRow = N 'show row if out of view

ShowChart
End Sub

