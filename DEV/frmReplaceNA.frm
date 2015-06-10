VERSION 5.00
Begin VB.Form frmReplaceNA 
   Caption         =   "Global Replace  - Missing Value Indicator"
   ClientHeight    =   3405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6750
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   130
   Icon            =   "frmReplaceNA.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3405
   ScaleWidth      =   6750
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   3720
         TabIndex        =   9
         Top             =   2640
         Width           =   1575
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   1200
         TabIndex        =   8
         Top             =   2640
         Width           =   1575
      End
      Begin VB.TextBox txtValue 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3480
         TabIndex        =   2
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label5 
         Caption         =   "If the Missing Value Indicator label says ""[blank]"" it means that blank grid cells represent missing data."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   1800
         Width           =   6135
      End
      Begin VB.Label Label4 
         Caption         =   "Note: To replace blank grid cells, delete all characters in the text box above."
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
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   5415
      End
      Begin VB.Label Label3 
         Caption         =   "In the Data Collection Grid,"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lblNAVal 
         Caption         =   "[blank]"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3480
         TabIndex        =   4
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "With the Missing Value Indicator:"
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Replace All Instance of:"
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   600
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmReplaceNA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim S As String

S = NAVal
If Trim(S) = "" Then S = "[blank]"
lblNAVal.Caption = S

End Sub

Private Sub cmdOK_Click()
Dim rtnval As Integer
Dim S As String
Dim I As Integer

rtnval = MsgBox("Globally Replacing the Missing Value Indicator " & _
        vbCrLf & "Will Modify Values in the Data Collection Grid." & _
        vbCrLf & vbCrLf & "Do You Wish to Proceed?", vbQuestion + vbOKCancel, "Visual Report Designer")
If rtnval = vbCancel Then
    Unload Me
Else
    S = Trim(txtValue.Text)
    ReplaceNA S
    'auto-save project
    If DataFlag Then
        WriteGridData
        If DataSavedFlag Then
            If LayoutFlag Then
                WriteLayoutFile
                For I = 0 To 14
                    HTMLFile = Left(FNLOG, Len(FNLOG) - 4) & "_rpt" & CStr(I + 1) & ".html"
                    If SaveLayout Then WriteHTML HTMLFile, I
                Next I
            End If
        End If
    End If
    Unload Me
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

