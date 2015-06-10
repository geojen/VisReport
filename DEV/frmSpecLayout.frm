VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSpecLayout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Display Options"
   ClientHeight    =   7020
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
   Icon            =   "frmSpecLayout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   11415
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10800
      Top             =   6360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   81
      ImageHeight     =   17
      MaskColor       =   -2147483643
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpecLayout.frx":0442
            Key             =   "plusminus"
            Object.Tag             =   "plusminus"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpecLayout.frx":14CA
            Key             =   "greenplusredminus"
            Object.Tag             =   "greenplusredminus"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpecLayout.frx":2552
            Key             =   "redplusgreenminus"
            Object.Tag             =   "redplusgreenminus"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpecLayout.frx":35DA
            Key             =   "greenred"
            Object.Tag             =   "greenred"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpecLayout.frx":4662
            Key             =   "redgreen"
            Object.Tag             =   "redgreen"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpecLayout.frx":56EA
            Key             =   "plustominus"
            Object.Tag             =   "plustominus"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpecLayout.frx":6772
            Key             =   "greenplustoredminus"
            Object.Tag             =   "greenplustoredminus"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpecLayout.frx":77FA
            Key             =   "redplustogreenminus"
            Object.Tag             =   "redplustogreenminus"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpecLayout.frx":8882
            Key             =   "greentored"
            Object.Tag             =   "greentored"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpecLayout.frx":990A
            Key             =   "redtogreen"
            Object.Tag             =   "redtogreen"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpecLayout.frx":A992
            Key             =   "redtoblack5"
            Object.Tag             =   "redtoblack5"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpecLayout.frx":BA1A
            Key             =   "blacktored5"
            Object.Tag             =   "blacktored5"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpecLayout.frx":CAA2
            Key             =   "redtoblue5"
            Object.Tag             =   "redtoblue5"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpecLayout.frx":DB2A
            Key             =   "bluetored5"
            Object.Tag             =   "bluetored5"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpecLayout.frx":EBB2
            Key             =   "whitetoblack5"
            Object.Tag             =   "whitetoblack5"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpecLayout.frx":FC3A
            Key             =   "blacktowhite5"
            Object.Tag             =   "blacktowhite5"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   6735
      Left            =   240
      TabIndex        =   24
      Top             =   120
      Width           =   10935
      Begin VB.Frame frmDesc 
         Height          =   735
         Left            =   0
         TabIndex        =   42
         Top             =   0
         Width           =   10935
         Begin VB.CommandButton cmdMinMax 
            Caption         =   "?"
            Height          =   255
            Left            =   10440
            TabIndex        =   43
            ToolTipText     =   "Check the bounds of the item"
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblDescrip 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   240
            TabIndex        =   2
            Top             =   240
            Width           =   10095
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "End Year for Data Analysis"
         Height          =   1455
         Left            =   360
         TabIndex        =   39
         Top             =   4560
         Width           =   5055
         Begin VB.OptionButton opXYear 
            Caption         =   "Do Not Edit"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   48
            Top             =   1080
            Width           =   1815
         End
         Begin VB.OptionButton opXYear 
            Caption         =   "Last non-blank year in the data:"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   15
            Top             =   360
            Width           =   3135
         End
         Begin VB.ComboBox cboXYear 
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   720
            Width           =   1335
         End
         Begin VB.OptionButton opXYear 
            Caption         =   "Other"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   16
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label lblXYear 
            Caption         =   "Xyear"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   3480
            TabIndex        =   45
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame frmFYear 
         Caption         =   "Start Year for Data Analysis"
         Height          =   1455
         Left            =   360
         TabIndex        =   38
         Top             =   3000
         Width           =   5055
         Begin VB.OptionButton opFYear 
            Caption         =   "Do Not Edit"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   47
            Top             =   1080
            Width           =   1815
         End
         Begin VB.OptionButton opFYear 
            Caption         =   "First non-blank year in the data:"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   12
            Top             =   360
            Width           =   3135
         End
         Begin VB.OptionButton opFYear 
            Caption         =   "Other"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   13
            Top             =   720
            Width           =   975
         End
         Begin VB.ComboBox cboFYear 
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label lblFYear 
            Caption         =   "Fyear"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   3360
            TabIndex        =   44
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Line Description"
         Height          =   1455
         Left            =   360
         TabIndex        =   37
         Top             =   1440
         Width           =   10215
         Begin VB.TextBox txtTag 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1560
            TabIndex        =   9
            Top             =   960
            Width           =   8295
         End
         Begin VB.CheckBox chkTag 
            Caption         =   "Item"
            Height          =   255
            Index           =   3
            Left            =   6600
            TabIndex        =   7
            Top             =   240
            Width           =   975
         End
         Begin VB.CheckBox chkTag 
            Caption         =   "Data Type"
            Height          =   255
            Index           =   2
            Left            =   5280
            TabIndex        =   6
            Top             =   240
            Width           =   1215
         End
         Begin VB.CheckBox chkTag 
            Caption         =   "Case"
            Height          =   255
            Index           =   1
            Left            =   4440
            TabIndex        =   5
            Top             =   240
            Width           =   855
         End
         Begin VB.CheckBox chkTag 
            Caption         =   "Source"
            Height          =   255
            Index           =   0
            Left            =   3480
            TabIndex        =   4
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton opTag 
            Caption         =   "Custom"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   8
            Top             =   960
            Width           =   1815
         End
         Begin VB.OptionButton opTag 
            Caption         =   "Data Collection Grid Titles"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   3
            Top             =   240
            Width           =   2775
         End
         Begin VB.Label lblLine 
            Caption         =   "Line Description:"
            Height          =   255
            Left            =   720
            TabIndex        =   41
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label lblTag 
            Caption         =   "None Specified"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2280
            TabIndex        =   40
            Top             =   600
            Width           =   7815
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Display Options"
         Height          =   2295
         Left            =   5640
         TabIndex        =   30
         Top             =   3000
         Width           =   4935
         Begin VB.CheckBox chkEditDisplay 
            Caption         =   "Do Not Edit Display Type"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   240
            Width           =   2535
         End
         Begin MSComctlLib.ImageCombo cboPalette 
            Height          =   330
            Left            =   3120
            TabIndex        =   21
            Top             =   1680
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483643
            BackColor       =   -2147483643
            Locked          =   -1  'True
         End
         Begin VB.ComboBox cboBins 
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   2760
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   480
            Width           =   1935
         End
         Begin VB.TextBox txtMark 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Index           =   0
            Left            =   2760
            TabIndex        =   19
            Top             =   840
            Width           =   1935
         End
         Begin VB.TextBox txtMark 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Index           =   1
            Left            =   2760
            TabIndex        =   20
            Top             =   1200
            Width           =   1935
         End
         Begin VB.Label lblBins 
            AutoSize        =   -1  'True
            Caption         =   "Specify Data Display Type"
            Height          =   195
            Left            =   360
            TabIndex        =   34
            Top             =   600
            Width           =   2265
         End
         Begin VB.Label lblMark 
            AutoSize        =   -1  'True
            Caption         =   "Cut Point Value"
            Height          =   195
            Index           =   0
            Left            =   600
            TabIndex        =   33
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label lblMark 
            AutoSize        =   -1  'True
            Caption         =   "Upper Cut Point Value"
            Height          =   195
            Index           =   1
            Left            =   600
            TabIndex        =   32
            Top             =   1320
            Width           =   1905
         End
         Begin VB.Label lblPalette 
            Caption         =   "Color Palette"
            Height          =   255
            Left            =   360
            TabIndex        =   31
            Top             =   1800
            Width           =   1695
         End
         Begin VB.Label lblHigh 
            Alignment       =   2  'Center
            Caption         =   "High Values ... Low Values"
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
            Left            =   2880
            TabIndex        =   46
            Top             =   2010
            Width           =   1935
         End
      End
      Begin VB.TextBox txtLine 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4680
         TabIndex        =   11
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtLinesUsed 
         BackColor       =   &H8000000F&
         Height          =   525
         Left            =   7440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Text            =   "frmSpecLayout.frx":10CC2
         Top             =   840
         Width           =   3135
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "CANCEL"
         Height          =   375
         Left            =   6480
         TabIndex        =   23
         Top             =   6120
         Width           =   2295
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "ADD/UPDATE"
         Default         =   -1  'True
         Height          =   375
         Left            =   2280
         TabIndex        =   22
         Top             =   6120
         Width           =   2295
      End
      Begin VB.ComboBox cboReport 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   840
         Width           =   975
      End
      Begin VB.Frame Frame5 
         Height          =   855
         Left            =   5640
         TabIndex        =   50
         Top             =   5160
         Width           =   4935
         Begin VB.CheckBox chkZero 
            Alignment       =   1  'Right Justify
            Caption         =   "Treat Zero Values as Missing Data"
            Height          =   255
            Left            =   360
            TabIndex        =   52
            Top             =   480
            Value           =   1  'Checked
            Width           =   4215
         End
         Begin VB.CheckBox chkEditZero 
            Caption         =   "Do Not Edit Missing Data Criteria"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   120
            Width           =   3855
         End
      End
      Begin VB.Label lblRowsAdd 
         Caption         =   "Number of Rows of Data to Add:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label lblNLines 
         Caption         =   "1"
         Height          =   255
         Left            =   3000
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lines Currently In Use"
         Height          =   390
         Index           =   6
         Left            =   6120
         TabIndex        =   27
         Top             =   840
         Width           =   1335
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Place on Line #"
         Height          =   195
         Index           =   1
         Left            =   3240
         TabIndex        =   26
         Top             =   960
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Place in Report #"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   25
         Top             =   960
         Width           =   1515
      End
   End
   Begin VB.Label lblJRpt 
      Caption         =   "JRpt"
      Height          =   255
      Left            =   720
      TabIndex        =   36
      Top             =   6840
      Width           =   495
   End
   Begin VB.Label lblJLine 
      Caption         =   "JLine"
      Height          =   255
      Left            =   1440
      TabIndex        =   35
      Top             =   6840
      Width           =   615
   End
   Begin VB.Label lblKey 
      Caption         =   "Key"
      Height          =   255
      Left            =   0
      TabIndex        =   29
      Top             =   6840
      Width           =   375
   End
End
Attribute VB_Name = "frmSpecLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
    
    cboFYear.Clear
    cboXYear.Clear
    For I = 1 To NYears
        cboFYear.AddItem CStr(StartYear + I - 1)
        cboXYear.AddItem CStr(StartYear + I - 1)
    Next I
    cboFYear.ListIndex = 0
    cboXYear.ListIndex = NYears - 1
    
    txtMark(0).Visible = False
    txtMark(1).Visible = False
    lblMark(0).Visible = False
    lblMark(1).Visible = False
    
    'hide the label to hold the data grid key
    lblKey.Visible = False
    'hide the label to hold the initial Report number (used in editing the report)
    lblJRpt.Visible = False
    'hide the label to hold the initial Line number (used in editing the report)
    lblJLine.Visible = False
            
End Sub
Private Sub opTag_Click(Index As Integer)
If opTag(0).Value = True Then
    chkTag(0).Enabled = True
    chkTag(1).Enabled = True
    chkTag(2).Enabled = True
    chkTag(3).Enabled = True
    lblLine.Visible = True
    lblTag.Visible = True
    txtTag.Visible = False
Else
    chkTag(0).Enabled = False
    chkTag(1).Enabled = False
    chkTag(2).Enabled = False
    chkTag(3).Enabled = False
    lblLine.Visible = False
    lblTag.Visible = False
    If opTag(1).Caption = "Custom" Then
        txtTag.Visible = True
    Else
        txtTag.Visible = False
    End If
End If
End Sub
Private Sub chkTag_Click(Index As Integer)
Dim G As MSFlexGrid
Dim S As String
Dim N As Integer

N = Val(lblKey.Caption)

Set G = frmData.MSFlexGrid1
S = ""
For Index = 0 To 3
    If chkTag(Index).Value = 1 Then
        S = S & G.TextMatrix(N, Index + 1) & " "
    End If
Next Index

lblTag.Caption = Trim(S)

End Sub
Private Sub opFYear_Click(Index As Integer)
Dim N As Integer
Dim I As Integer

If opFYear(0).Value = True Then
    N = Val(lblKey.Caption)
    I = GetDataMinYear(N)
    lblFYear.Enabled = True
    cboFYear.Visible = False
    cboFYear.ListIndex = I - StartYear
ElseIf opFYear(1).Value = True Then
    lblFYear.Enabled = False
    cboFYear.Visible = True
ElseIf opFYear(2).Value = True Then
    cboFYear.Visible = False
End If
End Sub
Private Sub opXYear_Click(Index As Integer)
Dim N As Integer
Dim I As Integer

If opXYear(0).Value = True Then
    N = Val(lblKey.Caption)
    I = GetDataMaxYear(N)
    lblXYear.Enabled = True
    cboXYear.Visible = False
    cboXYear.ListIndex = I - StartYear
ElseIf opXYear(1).Value = True Then
    lblXYear.Enabled = False
    cboXYear.Visible = True
ElseIf opXYear(2).Value = True Then
    cboXYear.Visible = False
End If
End Sub

Private Sub cboBins_Click()
    Dim K As Integer
    Dim I As Integer
    
    K = cboBins.ListIndex
    
    If K = 0 Then
        'set cut point options
        txtMark(0).Visible = True
        txtMark(1).Visible = False
        lblMark(0).Visible = True
        lblMark(0).Caption = "Cut Point Value"
        lblMark(1).Visible = False
    ElseIf K = 1 Then
        'set cut point options
        txtMark(0).Visible = True
        txtMark(1).Visible = True
        lblMark(0).Visible = True
        lblMark(0).Caption = "Lower Cut Point Value"
        lblMark(1).Visible = True
    Else
        'disable cut point options
        txtMark(0).Visible = False
        txtMark(1).Visible = False
        lblMark(0).Visible = False
        lblMark(0).Caption = "Cut Point Value"
        lblMark(1).Visible = False
    End If
    
    'set color palette selector
    InitPalette K
    
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
        If G.TextMatrix(I, 1) <> "" Then
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

'get default row
txtLine.Text = CStr(G.Rows)
End Sub
Private Sub cmdCancel_Click()
CancelFlag = True
Unload Me

End Sub
Private Sub cmdMinMax_Click()
    GetDataStats Val(lblKey.Caption)
End Sub

Private Sub cmdUpdate_Click()
    Dim flag As Boolean
    
    'check years
    If cboXYear.ListIndex < cboFYear.ListIndex Then
        MsgBox "Invalid Year Range", vbExclamation, "Visual Report Designer"
        Exit Sub
    End If
    'check line number
    If Val(txtLine.Text) < 1 Then
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
    'check that palette options don't conflict
    flag = ColorCheck
    If Not flag Then Exit Sub
    
    If cmdUpdate.Caption = "UPDATE" Then
        If frmDesc.Visible = False Then
            UpdateLayoutBatch 'if user is in batch mode
        Else
            UpdateLayout 'if user is in single-edit mode
        End If
    Else
        AddtoLayout
    End If

End Sub
Private Sub cboPalette_Click()
If chkZero.Enabled = True Then
    chkZero.SetFocus
Else
    chkEditZero.SetFocus
End If
End Sub
Private Sub chkEditDisplay_Click()
If chkEditDisplay.Value = 1 Then
    lblBins.Enabled = False
    cboBins.Enabled = False
    lblMark(0).Enabled = False
    txtMark(0).Enabled = False
    lblMark(1).Enabled = False
    txtMark(1).Enabled = False
    lblPalette.Enabled = False
    cboPalette.Enabled = False
Else
    lblBins.Enabled = True
    cboBins.Enabled = True
    lblMark(0).Enabled = True
    txtMark(0).Enabled = True
    lblMark(1).Enabled = True
    txtMark(1).Enabled = True
    lblPalette.Enabled = True
    cboPalette.Enabled = True
End If
End Sub
Private Sub chkEditZero_Click()
If chkEditZero.Value = 1 Then
    chkZero.Enabled = False
Else
    chkZero.Enabled = True
End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If UnloadMode = vbFormControlMenu Then CancelFlag = True

End Sub
