VERSION 5.00
Begin VB.Form frmDefaultLayout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Specify Default Layout"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11415
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDefaultLayout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   11415
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   5175
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   10935
      Begin VB.TextBox txtLinesUsed 
         BackColor       =   &H8000000F&
         Height          =   495
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   31
         Text            =   "frmDefaultLayout.frx":0442
         Top             =   1920
         Width           =   4575
      End
      Begin VB.TextBox txtLine 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3120
         TabIndex        =   30
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Frame Frame2 
         Height          =   3015
         Left            =   5280
         TabIndex        =   8
         Top             =   1200
         Width           =   5175
         Begin VB.ComboBox cboBins 
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   3120
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txtMark 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Index           =   0
            Left            =   3120
            TabIndex        =   17
            Top             =   840
            Width           =   1695
         End
         Begin VB.TextBox txtMark 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Index           =   1
            Left            =   3120
            TabIndex        =   16
            Top             =   1200
            Width           =   1695
         End
         Begin VB.CheckBox chkZero 
            Alignment       =   1  'Right Justify
            Caption         =   "Treat Zero Values as Missing Data"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   2520
            Value           =   1  'Checked
            Width           =   4455
         End
         Begin VB.ComboBox cboHighLow 
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   1560
            Width           =   2895
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   4
            Left            =   3840
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   13
            Top             =   1920
            Width           =   250
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   3
            Left            =   3600
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   12
            Top             =   1920
            Width           =   250
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   2
            Left            =   3360
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   11
            Top             =   1920
            Width           =   250
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   1
            Left            =   3120
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   10
            Top             =   1920
            Width           =   250
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   0
            Left            =   2880
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   9
            Top             =   1920
            Width           =   250
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Specify Data Display Type"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   22
            Top             =   480
            Width           =   2265
         End
         Begin VB.Label lblMark 
            AutoSize        =   -1  'True
            Caption         =   "Cut Point Value"
            Height          =   195
            Index           =   0
            Left            =   600
            TabIndex        =   21
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label lblMark 
            AutoSize        =   -1  'True
            Caption         =   "Upper Cut Point Value"
            Height          =   195
            Index           =   1
            Left            =   600
            TabIndex        =   20
            Top             =   1200
            Width           =   1905
         End
         Begin VB.Label lblHighLow 
            Caption         =   "Color Palette"
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   1560
            Width           =   1695
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "CANCEL"
         Height          =   375
         Left            =   6360
         TabIndex        =   6
         Top             =   4560
         Width           =   2295
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   2160
         TabIndex        =   4
         Top             =   4560
         Width           =   2295
      End
      Begin VB.ComboBox cboReport 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label lblNLines 
         Caption         =   "1"
         Height          =   255
         Left            =   3240
         TabIndex        =   33
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lines Currently In Use"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   32
         Top             =   1680
         Width           =   1890
      End
      Begin VB.Label Label6 
         Caption         =   "Number of Rows of Data to Add:"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label5 
         Caption         =   "Last non-blank year in the data"
         Height          =   375
         Left            =   3120
         TabIndex        =   28
         Top             =   3720
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "First non-blank year in the data"
         Height          =   495
         Left            =   3120
         TabIndex        =   27
         Top             =   3240
         Width           =   1815
      End
      Begin VB.Label lblLine 
         Height          =   255
         Left            =   3000
         TabIndex        =   26
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Data Collection Grid, columns 2 - 5  ( Source + Case + Data Type + Item )"
         Height          =   255
         Left            =   3240
         TabIndex        =   25
         Top             =   600
         Width           =   6495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Start Year for Data Analysis:"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   24
         Top             =   3240
         Width           =   2430
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "End Year for Data Analysis:"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   23
         Top             =   3720
         Width           =   2355
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Line Description:  "
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Begin Adding Rows at Line: "
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   2640
         Width           =   2430
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Place in Report #"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   1320
         Width           =   1515
      End
   End
   Begin VB.Label lblKey 
      Caption         =   "Key"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmDefaultLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboBins_Click()
    Dim K As Integer
    Dim I As Integer
    
    K = cboBins.ListIndex
    
    If K = 0 Then
        txtMark(0).Visible = True
        txtMark(1).Visible = False
        lblMark(0).Visible = True
        lblMark(0).Caption = "Cut Point Value"
        lblMark(1).Visible = False
    ElseIf K = 1 Then
        txtMark(0).Visible = True
        txtMark(1).Visible = True
        lblMark(0).Visible = True
        lblMark(0).Caption = "Lower Cut Point Value"
        lblMark(1).Visible = True
    Else
        txtMark(0).Visible = False
        txtMark(1).Visible = False
        lblMark(0).Visible = False
        lblMark(0).Caption = "Cut Point Value"
        lblMark(1).Visible = False
    End If
    
    'set color palette selector
    If K = 2 Then
        lblHighLow.Visible = True
        cboHighLow.Visible = True
        For I = 0 To 4
            Picture1(I).Visible = True
        Next I
    Else
        lblHighLow.Visible = False
        cboHighLow.Visible = False
        For I = 0 To 4
            Picture1(I).Visible = False
        Next I
    End If
End Sub

Private Sub cboReport_Click()
Dim G As MSFlexGrid
Dim N As Integer
Dim I As Integer
Dim S As String

' get the number of lines used in the survey
N = cboReport.ListIndex
Set G = frmCompose.grdReport(N)
S = ""
If G.Rows > 1 Then
    For I = 1 To G.Rows - 1
        If G.TextArray(FGIndex(G, I, 1)) <> "" Then
            If S <> "" Then S = S & ", "
            S = S & CStr(I)
        End If
    Next I
End If
If S = "" Then
    txtLinesUsed.Text = "none"
Else
    txtLinesUsed.Text = S
End If

'initialize line number to start adding data
txtLine.Text = CStr(G.Rows)
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub
Private Sub cmdUpdate_Click()
    Dim I As Integer
    Dim iRpt As Integer
    Dim N As Integer
    Dim iKey As Integer
    Dim iLine As Integer
    
    'check line number
    iLine = Val(txtLine.Text)
    If iLine < 1 Then
        MsgBox "Invalid Line Number", vbExclamation, "Visual Report Designer"
        Exit Sub
    End If
    'check dual cut point values
    If cboBins.ListIndex = 1 Then
        If Val(txtMark(1).Text) <= Val(txtMark(0).Text) Then
            MsgBox "Upper Cut Point is Less than the Lower Cut Point", vbExclamation, "Visual Report Designer"
            Exit Sub
        End If
    End If
    
    iRpt = cboReport.ListIndex
    iKey = Val(lblKey.Caption)
    N = Val(lblNLines.Caption)
    For I = 1 To N
        AddRptDefault iRpt, (iLine + I - 1), (iKey + I - 1)
    Next I
    
    WriteLayoutFile
    HTMLFile = Left(FNLOG, Len(FNLOG) - 4) & "_rpt" & CStr(iRpt + 1) & ".html"
    If SaveLayout Then WriteHTML HTMLFile, iRpt
    
    frmCompose.ZOrder
    frmCompose.SSTab1.Tab = iRpt
    Unload Me
End Sub
Private Sub Form_Load()
    Dim I As Integer

    cboReport.Clear
    For I = 1 To 15
        cboReport.AddItem Space(2) + CStr(I)
    Next I
    cboReport.ListIndex = frmCompose.SSTab1.Tab
    
    cboBins.Clear
    cboBins.AddItem "Single Cut Point"
    cboBins.AddItem "Dual Cut Points"
  '  cboBins.AddItem "Quartiles"
    cboBins.AddItem "Quintiles"
    cboBins.ListIndex = 2
    
    cboHighLow.Clear
    cboHighLow.AddItem "High = Black ----> Low = Red"
    cboHighLow.AddItem "Low = Black ----> High = Red"
    cboHighLow.ListIndex = 0
    Picture1(0).Picture = LoadPicture(App.Path + "\symbols\black.bmp")
    Picture1(1).Picture = LoadPicture(App.Path + "\symbols\halfblack.bmp")
    Picture1(2).Picture = LoadPicture(App.Path + "\symbols\white.bmp")
    Picture1(3).Picture = LoadPicture(App.Path + "\symbols\halfred.bmp")
    Picture1(4).Picture = LoadPicture(App.Path + "\symbols\red.bmp")
    
    txtMark(0).Visible = False
    txtMark(1).Visible = False
    lblMark(0).Visible = False
    lblMark(1).Visible = False
    
    'hide the label to hold the data grid key
    lblKey.Visible = False
            
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'If UnloadMode = vbFormControlMenu Then CancelFlag = True

End Sub

