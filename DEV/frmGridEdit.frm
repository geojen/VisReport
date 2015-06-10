VERSION 5.00
Begin VB.Form frmGridEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grid Input Form"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10740
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1003
   Icon            =   "frmGridEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   10740
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFill 
      Caption         =   "Fill Column"
      Height          =   375
      HelpContextID   =   1026
      Index           =   1
      Left            =   8640
      TabIndex        =   12
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton cmdFill 
      Caption         =   "Fill Row"
      Height          =   375
      HelpContextID   =   1026
      Index           =   0
      Left            =   8640
      TabIndex        =   11
      Top             =   360
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Edit Mode"
      Height          =   1335
      Left            =   4920
      TabIndex        =   7
      Top             =   120
      Width           =   3495
      Begin VB.OptionButton optMode 
         Caption         =   "Continuous Edit by Columnn"
         Height          =   255
         HelpContextID   =   1026
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   2775
      End
      Begin VB.OptionButton optMode 
         Caption         =   "Continuous Edit by Row"
         Height          =   255
         HelpContextID   =   1026
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   2775
      End
      Begin VB.OptionButton optMode 
         Caption         =   "Edit This Cell Only"
         Height          =   255
         HelpContextID   =   1026
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      HelpContextID   =   1026
      Left            =   2280
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtValue 
      ForeColor       =   &H00FF0000&
      Height          =   285
      HelpContextID   =   1026
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label lblCol 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   3840
      TabIndex        =   5
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lblRow 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Column"
      Height          =   195
      Index           =   2
      Left            =   2880
      TabIndex        =   3
      Top             =   720
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Row"
      Height          =   195
      Index           =   1
      Left            =   600
      TabIndex        =   2
      Top             =   720
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Value"
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   495
   End
End
Attribute VB_Name = "frmGridEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Jopt As Integer

Private Sub cmdFill_Click(Index As Integer)
    FillGridEditData Index
End Sub

Private Sub cmdOK_Click()
    PutData
End Sub
Private Sub PutData()

 '   If txtValue.Text = "" Then
 '       Me.Hide
 '       Exit Sub
 '   End If
    PutGridEditData
    If Jopt = 0 Then
        Me.Hide
    ElseIf Jopt = 2 Then
        If JRow < NRow Then
            JRow = JRow + 1
            lblRow.Caption = CStr(JRow)
        ElseIf JCol < NCol Then
            JRow = 1
            lblRow.Caption = CStr(JRow)
            JCol = JCol + 1
            lblCol.Caption = CStr(JCol)
        Else
            Me.Hide
        End If
        GetGridEditData
        'make sure row doesn't fall off page
        If FG.Row = 1 Then
            FG.TopRow = 1
        ElseIf FG.Row = FG.Rows - 1 Then
            FG.TopRow = FG.TopRow + 1
        Else
            If FG.RowIsVisible(FG.Row + 1) = False Then
                FG.TopRow = FG.TopRow + 1
            End If
        End If
        'make sure column doesn't fall off page
        If FG.Col = 1 Then
            FG.LeftCol = 1
        ElseIf FG.Col = FG.Cols - 1 Then
            FG.LeftCol = FG.LeftCol + 1
        Else
            If FG.ColIsVisible(FG.Col + 1) = False Then
                FG.LeftCol = FG.LeftCol + 1
            End If
        End If
    Else
        If JCol < NCol Then
            JCol = JCol + 1
            lblCol.Caption = CStr(JCol)
        ElseIf JRow < NRow Then
            JRow = JRow + 1
            lblRow.Caption = CStr(JRow)
            JCol = 1
            lblCol.Caption = CStr(JCol)
        Else
            FG.CellBackColor = RGB(255, 255, 255)
            Me.Hide
        End If
        GetGridEditData
        'make sure row doesn't fall off page
        If FG.Row = 1 Then
            FG.TopRow = 1
        ElseIf FG.Row = FG.Rows - 1 Then
            FG.TopRow = FG.TopRow + 1
        Else
            If FG.RowIsVisible(FG.Row + 1) = False Then
                FG.TopRow = FG.TopRow + 1
            End If
        End If
        'make sure column doesn't fall off page
        If FG.Col = 1 Then
            FG.LeftCol = 1
        ElseIf FG.Col = FG.Cols - 1 Then
            FG.LeftCol = FG.LeftCol + 1
        Else
            If FG.ColIsVisible(FG.Col + 1) = False Then
                FG.LeftCol = FG.LeftCol + 1
            End If
        End If
    End If
End Sub

Private Sub Form_Activate()
    Me.Top = 0
    Me.Left = (Screen.Width - Me.ScaleWidth) \ 2
    txtValue.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FG.CellBackColor = RGB(255, 255, 255)
End Sub
Private Sub optMode_Click(Index As Integer)
    Jopt = Index
End Sub

Private Sub txtValue_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        PutData
        KeyAscii = 0
    ElseIf KeyAscii = 27 Then
        txtValue.Text = ""
        KeyAscii = 0
    End If
End Sub
