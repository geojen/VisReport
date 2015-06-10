VERSION 5.00
Begin VB.Form frmSpecUserAdded 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Data to Collection"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6825
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSpecUserAdded.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   6825
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   17
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Frame frmSelect 
      Caption         =   "Select Data Source"
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6375
      Begin VB.Frame frmUser 
         Height          =   1695
         Left            =   2400
         TabIndex        =   13
         Top             =   960
         Width           =   3735
         Begin VB.CommandButton cmdInfo 
            Caption         =   "?"
            Height          =   255
            Left            =   3470
            TabIndex        =   18
            ToolTipText     =   "Get Help on This"
            Top             =   120
            Width           =   255
         End
         Begin VB.TextBox txtUser 
            ForeColor       =   &H00FF0000&
            Height          =   285
            HelpContextID   =   106
            Index           =   0
            Left            =   1800
            TabIndex        =   9
            ToolTipText     =   "Enter the First Year in Your Dataset"
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox txtUser 
            ForeColor       =   &H00FF0000&
            Height          =   285
            HelpContextID   =   106
            Index           =   1
            Left            =   1800
            TabIndex        =   10
            ToolTipText     =   "Enter the Last Year in Your Dataset"
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox txtUser 
            ForeColor       =   &H00FF0000&
            Height          =   285
            HelpContextID   =   106
            Index           =   2
            Left            =   1800
            TabIndex        =   11
            Text            =   "1"
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Start Year"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   16
            Top             =   360
            Width           =   870
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "End Year"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   15
            Top             =   720
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Number of Rows"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   14
            Top             =   1080
            Width           =   1410
         End
      End
      Begin VB.OptionButton opType 
         Caption         =   "User Supplied Data"
         Height          =   255
         HelpContextID   =   106
         Index           =   7
         Left            =   2040
         TabIndex        =   8
         Top             =   720
         Width           =   3255
      End
      Begin VB.OptionButton opType 
         Caption         =   "SAGA Sample Length Weight Data"
         Height          =   255
         HelpContextID   =   105
         Index           =   6
         Left            =   2040
         TabIndex        =   7
         Top             =   360
         Width           =   3495
      End
      Begin VB.OptionButton opType 
         Caption         =   "VPA"
         Height          =   255
         HelpContextID   =   105
         Index           =   5
         Left            =   240
         TabIndex        =   6
         Top             =   2160
         Width           =   1455
      End
      Begin VB.OptionButton opType 
         Caption         =   "CSA"
         Height          =   255
         HelpContextID   =   105
         Index           =   4
         Left            =   240
         TabIndex        =   5
         Top             =   1800
         Width           =   1935
      End
      Begin VB.OptionButton opType 
         Caption         =   "ASPIC"
         Height          =   255
         HelpContextID   =   105
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   1935
      End
      Begin VB.OptionButton opType 
         Caption         =   "ASAP"
         Height          =   255
         HelpContextID   =   105
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   1935
      End
      Begin VB.OptionButton opType 
         Caption         =   "AIM"
         Height          =   255
         HelpContextID   =   105
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1935
      End
      Begin VB.OptionButton opType 
         Caption         =   "AgePro"
         Height          =   255
         HelpContextID   =   105
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmSpecUserAdded"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdInfo_Click()
frmHelp.Show vbModal
End Sub

Private Sub cmdOK_Click()
Dim I As Integer
Dim flag As Boolean

flag = False
For I = 0 To 7
    If opType(I).Value = True Then
        flag = True
        Exit For
    End If
Next I
If Not flag Then
    MsgBox "Please Select a Data Source", vbInformation, "Visual Report Designer"
    Exit Sub
End If

Select Case I
    Case 0 'AgePro
        OpenDataFile 6
    Case 1 'AIM
        OpenDataFile 5
    Case 2 'ASAP
        OpenDataFile 4
    Case 3 'ASPIC
        OpenDataFile 3
    Case 4 'CSA
        OpenDataFile 2
    Case 5 'VPA
        OpenDataFile 1
    Case 6 'sample length weight
        OpenDataFile 7
    Case 7 'user supplied
        InitUserDataGrid
End Select

End Sub

Private Sub Form_Load()
frmUser.Visible = False
End Sub

Private Sub opType_Click(Index As Integer)
If opType(7).Value = True Then
    frmUser.Visible = True
Else
    frmUser.Visible = False
End If
End Sub
