VERSION 5.00
Begin VB.Form frmDataList 
   Caption         =   "Catalogue of Available Data"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11835
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDataList.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8550
   ScaleWidth      =   11835
   Begin VB.Frame Frame1 
      Caption         =   "Data Collection"
      Height          =   7455
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   11265
      Begin VB.CommandButton cmdShowPlot 
         Caption         =   "Show Plot of Selected Item"
         Height          =   375
         Left            =   3840
         TabIndex        =   3
         Top             =   6120
         Width           =   3855
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Height          =   375
         Left            =   3840
         TabIndex        =   2
         Top             =   5520
         Width           =   3855
      End
      Begin VB.ListBox lstData 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   4890
         HelpContextID   =   1001
         Left            =   240
         TabIndex        =   1
         ToolTipText     =   "Double Click to Select Item"
         Top             =   360
         Width           =   10815
      End
   End
End
Attribute VB_Name = "frmDataList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboReport_Click()
    Dim I As Integer
    Dim G As MSFlexGrid
    
    
    
End Sub

Private Sub cmdSearch_Click()
    frmSearch.Show vbModal
End Sub



Private Sub cmdShowPlot_Click()
    ShowChart
End Sub

Private Sub lstData_DblClick()
    SelectItem
End Sub



