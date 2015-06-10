VERSION 5.00
Begin VB.Form frmMissing 
   Caption         =   "Missing Value Indicator"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5580
   HelpContextID   =   129
   Icon            =   "frmMissing.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   5580
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmChange 
      Caption         =   "Change the Missing Value Indicator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   9
      Top             =   2160
      Width           =   5055
      Begin VB.TextBox txtMissing 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1680
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "New Value:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "To indicate that blank grid cells should be treated as missing, leave the New Value text box blank."
         Height          =   495
         Left            =   360
         TabIndex        =   11
         Top             =   840
         Width           =   4455
      End
   End
   Begin VB.Frame frmReplace 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   6
      Top             =   3600
      Width           =   5055
      Begin VB.CheckBox chkReplace 
         Caption         =   "Replace Values in Data Collection Grid"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label lblReplace 
         Caption         =   "Checking this option will replace all instances of the current Missing Value Indicator with the new value specified."
         Enabled         =   0   'False
         Height          =   495
         Left            =   360
         TabIndex        =   8
         Top             =   600
         Width           =   4455
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Current Missing Value Indicator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   5055
      Begin VB.Label Label4 
         Caption         =   "Current Value:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblNAVal 
         Caption         =   "[blank]"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1680
         TabIndex        =   4
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label5 
         Caption         =   "Note: If the Current Value label says ""[blank]"" it means that blank grid cells currently represent missing data."
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   4455
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Any datum that matches the Missing Value Indicator will be excluded from the binning calculations."
      Height          =   495
      Left            =   240
      TabIndex        =   13
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmMissing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
    Dim S As String
    
    S = NAVal
    If Trim(S) = "" Then S = "[blank]"
    lblNAVal.Caption = S
    
    If DataFlag Then
        chkReplace.Enabled = True
        lblReplace.Enabled = True
    End If
End Sub
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim S As String
    Dim T As String
    Dim I As Integer
    
    'get old missing value indicator
    T = NAVal
    
    S = Trim(txtMissing.Text)
    If Trim(NAVal) <> S Then
        'change missing value and save to config file
        NAVal = Trim(txtMissing.Text)
        WriteCfg
        
        'change grid data if specified
        If DataFlag And chkReplace.Value = 1 Then
            ReplaceNA T
            'auto-save data
            WriteGridData
        End If
        
        'make sure change is propagated in reports
        If LayoutFlag Then
            WriteLayoutFile
            For I = 0 To 14
                HTMLFile = Left(FNLOG, Len(FNLOG) - 4) & "_rpt" & CStr(I + 1) & ".html"
                If SaveLayout Then WriteHTML HTMLFile, I
            Next I
        End If
        
    End If
    
    Unload Me
End Sub


