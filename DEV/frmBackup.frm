VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBackup 
   Caption         =   "Restore Project from Backup"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11130
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   126
   Icon            =   "frmBackup.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   11130
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Select a Backup Copy to Open"
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10815
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2775
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   4895
         _Version        =   393216
         Rows            =   6
         Cols            =   4
         FixedCols       =   0
         ForeColor       =   16711680
         AllowUserResizing=   3
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "OK"
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   3360
         Width           =   2175
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   6960
         TabIndex        =   1
         Top             =   3360
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOpen_Click()
    If MSFlexGrid1.Row < 1 Or MSFlexGrid1.Col < 1 Then
        MsgBox "Please Select a Backup File to Restore", vbInformation, "Visual Report Designer"
        Exit Sub
    End If
    
    RestoreBackup
End Sub

Private Sub MSFlexGrid1_Click()
Dim I As Integer
Dim J As Integer
Dim G As MSFlexGrid
Dim iRow As Integer
Dim iCol As Integer

'highlight selected grid row

Set G = MSFlexGrid1
'first get current cell
iRow = G.Row
iCol = G.Col

DoEvents
'go through and make all other rows' background color white
If iRow > 1 Then
    For I = 1 To iRow - 1
        For J = 0 To G.Cols - 1
            G.Row = I
            G.Col = J
            G.CellBackColor = &H8000000E
        Next J
    Next I
End If
If iRow < G.Rows - 1 Then
    For I = iRow + 1 To G.Rows - 1
        For J = 0 To G.Cols - 1
            G.Row = I
            G.Col = J
            G.CellBackColor = &H8000000E
        Next J
    Next I
End If

'now highlight the selected row
G.Row = iRow
For J = 0 To G.Cols - 1
    G.Col = J
    G.CellBackColor = &HC0FFFF
Next J

'go back to original column
G.Col = iCol

End Sub
