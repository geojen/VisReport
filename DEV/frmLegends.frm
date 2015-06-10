VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLegends 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Report Display Preferences"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6300
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   120
   Icon            =   "frmLegends.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   6300
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.Frame Frame1 
         Caption         =   "Dispersion Statistic"
         Height          =   975
         Left            =   360
         TabIndex        =   65
         Top             =   5400
         Width           =   5295
         Begin VB.TextBox txtSigDig 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   3480
            TabIndex        =   68
            Text            =   "1"
            Top             =   560
            Width           =   1335
         End
         Begin VB.CheckBox chkDispersion 
            Caption         =   "Display Dispersion Statistic for Each Line of Data"
            Height          =   195
            Left            =   240
            TabIndex        =   66
            Top             =   300
            Value           =   1  'Checked
            Width           =   4815
         End
         Begin VB.Label lblSigDig 
            Caption         =   "Number of Digits After Decimal:"
            Enabled         =   0   'False
            Height          =   255
            Left            =   600
            TabIndex        =   67
            Top             =   600
            Width           =   2775
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   3600
         TabIndex        =   64
         Top             =   6600
         Width           =   1575
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   840
         TabIndex        =   63
         Top             =   6600
         Width           =   1575
      End
      Begin VB.Frame Frame5 
         Caption         =   "Legends for Quintile Reports"
         Height          =   3855
         Left            =   360
         TabIndex        =   5
         Top             =   480
         Width           =   5295
         Begin TabDlg.SSTab SSTab1 
            Height          =   3015
            Left            =   240
            TabIndex        =   7
            Top             =   600
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   5318
            _Version        =   393216
            TabHeight       =   617
            ForeColor       =   16711680
            TabCaption(0)   =   " "
            TabPicture(0)   =   "frmLegends.frx":0442
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "frameLabels(0)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   " "
            TabPicture(1)   =   "frmLegends.frx":14CA
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "frameLabels(1)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   " "
            TabPicture(2)   =   "frmLegends.frx":2552
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "frameLabels(2)"
            Tab(2).ControlCount=   1
            Begin VB.Frame frameLabels 
               Height          =   2295
               Index           =   2
               Left            =   -74760
               TabIndex        =   45
               Top             =   480
               Width           =   4335
               Begin VB.TextBox txtLegend3 
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Index           =   2
                  Left            =   1680
                  TabIndex        =   55
                  Text            =   "Middle"
                  Top             =   1080
                  Width           =   2415
               End
               Begin VB.PictureBox Picture3 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  ForeColor       =   &H80000008&
                  Height          =   250
                  Index           =   4
                  Left            =   1080
                  ScaleHeight     =   225
                  ScaleWidth      =   225
                  TabIndex        =   54
                  Top             =   1800
                  Width           =   250
               End
               Begin VB.PictureBox Picture3 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  ForeColor       =   &H80000008&
                  Height          =   250
                  Index           =   3
                  Left            =   1080
                  ScaleHeight     =   225
                  ScaleWidth      =   225
                  TabIndex        =   53
                  Top             =   1440
                  Width           =   250
               End
               Begin VB.PictureBox Picture3 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  ForeColor       =   &H80000008&
                  Height          =   250
                  Index           =   2
                  Left            =   1080
                  ScaleHeight     =   225
                  ScaleWidth      =   225
                  TabIndex        =   52
                  Top             =   1080
                  Width           =   250
               End
               Begin VB.PictureBox Picture3 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  ForeColor       =   &H80000008&
                  Height          =   250
                  Index           =   1
                  Left            =   1080
                  ScaleHeight     =   225
                  ScaleWidth      =   225
                  TabIndex        =   51
                  Top             =   720
                  Width           =   250
               End
               Begin VB.PictureBox Picture3 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  ForeColor       =   &H80000008&
                  Height          =   250
                  Index           =   0
                  Left            =   1080
                  ScaleHeight     =   225
                  ScaleWidth      =   225
                  TabIndex        =   50
                  Top             =   360
                  Width           =   250
               End
               Begin VB.TextBox txtLegend3 
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Index           =   4
                  Left            =   1680
                  TabIndex        =   49
                  Text            =   "Lowest"
                  Top             =   1800
                  Width           =   2415
               End
               Begin VB.TextBox txtLegend3 
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Index           =   3
                  Left            =   1680
                  TabIndex        =   48
                  Text            =   "2nd Lowest"
                  Top             =   1440
                  Width           =   2415
               End
               Begin VB.TextBox txtLegend3 
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Index           =   1
                  Left            =   1680
                  TabIndex        =   47
                  Text            =   "2nd Highest"
                  Top             =   720
                  Width           =   2415
               End
               Begin VB.TextBox txtLegend3 
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Index           =   0
                  Left            =   1680
                  TabIndex        =   46
                  Text            =   "Highest"
                  Top             =   360
                  Width           =   2415
               End
               Begin VB.Label Label2 
                  Caption         =   "="
                  Height          =   255
                  Index           =   14
                  Left            =   1440
                  TabIndex        =   62
                  Top             =   1800
                  Width           =   135
               End
               Begin VB.Label Label2 
                  Caption         =   "="
                  Height          =   255
                  Index           =   13
                  Left            =   1440
                  TabIndex        =   61
                  Top             =   1440
                  Width           =   135
               End
               Begin VB.Label Label2 
                  Caption         =   "="
                  Height          =   255
                  Index           =   12
                  Left            =   1440
                  TabIndex        =   60
                  Top             =   1080
                  Width           =   135
               End
               Begin VB.Label Label2 
                  Caption         =   "="
                  Height          =   255
                  Index           =   11
                  Left            =   1440
                  TabIndex        =   59
                  Top             =   720
                  Width           =   135
               End
               Begin VB.Label Label2 
                  Caption         =   "="
                  Height          =   255
                  Index           =   10
                  Left            =   1440
                  TabIndex        =   58
                  Top             =   360
                  Width           =   135
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Lowest Value"
                  Height          =   435
                  Index           =   6
                  Left            =   240
                  TabIndex        =   57
                  Top             =   1680
                  Width           =   675
                  WordWrap        =   -1  'True
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Highest Value"
                  Height          =   390
                  Index           =   4
                  Left            =   240
                  TabIndex        =   56
                  Top             =   240
                  Width           =   720
                  WordWrap        =   -1  'True
               End
            End
            Begin VB.Frame frameLabels 
               Height          =   2295
               Index           =   1
               Left            =   -74760
               TabIndex        =   27
               Top             =   480
               Width           =   4335
               Begin VB.TextBox txtLegend2 
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Index           =   0
                  Left            =   1680
                  TabIndex        =   37
                  Text            =   "Highest"
                  Top             =   360
                  Width           =   2415
               End
               Begin VB.TextBox txtLegend2 
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Index           =   1
                  Left            =   1680
                  TabIndex        =   36
                  Text            =   "2nd Highest"
                  Top             =   720
                  Width           =   2415
               End
               Begin VB.TextBox txtLegend2 
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Index           =   3
                  Left            =   1680
                  TabIndex        =   35
                  Text            =   "2nd Lowest"
                  Top             =   1440
                  Width           =   2415
               End
               Begin VB.TextBox txtLegend2 
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Index           =   4
                  Left            =   1680
                  TabIndex        =   34
                  Text            =   "Lowest"
                  Top             =   1800
                  Width           =   2415
               End
               Begin VB.PictureBox Picture2 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  ForeColor       =   &H80000008&
                  Height          =   250
                  Index           =   0
                  Left            =   1080
                  ScaleHeight     =   225
                  ScaleWidth      =   225
                  TabIndex        =   33
                  Top             =   360
                  Width           =   250
               End
               Begin VB.PictureBox Picture2 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  ForeColor       =   &H80000008&
                  Height          =   250
                  Index           =   1
                  Left            =   1080
                  ScaleHeight     =   225
                  ScaleWidth      =   225
                  TabIndex        =   32
                  Top             =   720
                  Width           =   250
               End
               Begin VB.PictureBox Picture2 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  ForeColor       =   &H80000008&
                  Height          =   250
                  Index           =   2
                  Left            =   1080
                  ScaleHeight     =   225
                  ScaleWidth      =   225
                  TabIndex        =   31
                  Top             =   1080
                  Width           =   250
               End
               Begin VB.PictureBox Picture2 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  ForeColor       =   &H80000008&
                  Height          =   250
                  Index           =   3
                  Left            =   1080
                  ScaleHeight     =   225
                  ScaleWidth      =   225
                  TabIndex        =   30
                  Top             =   1440
                  Width           =   250
               End
               Begin VB.PictureBox Picture2 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  ForeColor       =   &H80000008&
                  Height          =   250
                  Index           =   4
                  Left            =   1080
                  ScaleHeight     =   225
                  ScaleWidth      =   225
                  TabIndex        =   29
                  Top             =   1800
                  Width           =   250
               End
               Begin VB.TextBox txtLegend2 
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Index           =   2
                  Left            =   1680
                  TabIndex        =   28
                  Text            =   "Middle"
                  Top             =   1080
                  Width           =   2415
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Highest Value"
                  Height          =   390
                  Index           =   3
                  Left            =   240
                  TabIndex        =   44
                  Top             =   240
                  Width           =   720
                  WordWrap        =   -1  'True
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Lowest Value"
                  Height          =   435
                  Index           =   2
                  Left            =   240
                  TabIndex        =   43
                  Top             =   1680
                  Width           =   675
                  WordWrap        =   -1  'True
               End
               Begin VB.Label Label2 
                  Caption         =   "="
                  Height          =   255
                  Index           =   9
                  Left            =   1440
                  TabIndex        =   42
                  Top             =   360
                  Width           =   135
               End
               Begin VB.Label Label2 
                  Caption         =   "="
                  Height          =   255
                  Index           =   8
                  Left            =   1440
                  TabIndex        =   41
                  Top             =   720
                  Width           =   135
               End
               Begin VB.Label Label2 
                  Caption         =   "="
                  Height          =   255
                  Index           =   7
                  Left            =   1440
                  TabIndex        =   40
                  Top             =   1080
                  Width           =   135
               End
               Begin VB.Label Label2 
                  Caption         =   "="
                  Height          =   255
                  Index           =   6
                  Left            =   1440
                  TabIndex        =   39
                  Top             =   1440
                  Width           =   135
               End
               Begin VB.Label Label2 
                  Caption         =   "="
                  Height          =   255
                  Index           =   5
                  Left            =   1440
                  TabIndex        =   38
                  Top             =   1800
                  Width           =   135
               End
            End
            Begin VB.Frame frameLabels 
               Height          =   2295
               Index           =   0
               Left            =   240
               TabIndex        =   8
               Top             =   480
               Width           =   4335
               Begin VB.TextBox txtLegend1 
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Index           =   2
                  Left            =   1680
                  TabIndex        =   18
                  Text            =   "Middle"
                  Top             =   1080
                  Width           =   2415
               End
               Begin VB.PictureBox Picture1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  ForeColor       =   &H80000008&
                  Height          =   250
                  Index           =   4
                  Left            =   1080
                  ScaleHeight     =   225
                  ScaleWidth      =   225
                  TabIndex        =   17
                  Top             =   1800
                  Width           =   250
               End
               Begin VB.PictureBox Picture1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  ForeColor       =   &H80000008&
                  Height          =   250
                  Index           =   3
                  Left            =   1080
                  ScaleHeight     =   225
                  ScaleWidth      =   225
                  TabIndex        =   16
                  Top             =   1440
                  Width           =   250
               End
               Begin VB.PictureBox Picture1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  ForeColor       =   &H80000008&
                  Height          =   250
                  Index           =   2
                  Left            =   1080
                  ScaleHeight     =   225
                  ScaleWidth      =   225
                  TabIndex        =   15
                  Top             =   1080
                  Width           =   250
               End
               Begin VB.PictureBox Picture1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  ForeColor       =   &H80000008&
                  Height          =   250
                  Index           =   1
                  Left            =   1080
                  ScaleHeight     =   225
                  ScaleWidth      =   225
                  TabIndex        =   14
                  Top             =   720
                  Width           =   250
               End
               Begin VB.PictureBox Picture1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  ForeColor       =   &H80000008&
                  Height          =   250
                  Index           =   0
                  Left            =   1080
                  ScaleHeight     =   225
                  ScaleWidth      =   225
                  TabIndex        =   13
                  Top             =   360
                  Width           =   250
               End
               Begin VB.TextBox txtLegend1 
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Index           =   4
                  Left            =   1680
                  TabIndex        =   12
                  Text            =   "Lowest"
                  Top             =   1800
                  Width           =   2415
               End
               Begin VB.TextBox txtLegend1 
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Index           =   3
                  Left            =   1680
                  TabIndex        =   11
                  Text            =   "2nd Lowest"
                  Top             =   1440
                  Width           =   2415
               End
               Begin VB.TextBox txtLegend1 
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Index           =   1
                  Left            =   1680
                  TabIndex        =   10
                  Text            =   "2nd Highest"
                  Top             =   720
                  Width           =   2415
               End
               Begin VB.TextBox txtLegend1 
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Index           =   0
                  Left            =   1680
                  TabIndex        =   9
                  Text            =   "Highest"
                  Top             =   360
                  Width           =   2415
               End
               Begin VB.Label Label2 
                  Caption         =   "="
                  Height          =   255
                  Index           =   4
                  Left            =   1440
                  TabIndex        =   25
                  Top             =   1800
                  Width           =   135
               End
               Begin VB.Label Label2 
                  Caption         =   "="
                  Height          =   255
                  Index           =   3
                  Left            =   1440
                  TabIndex        =   24
                  Top             =   1440
                  Width           =   135
               End
               Begin VB.Label Label2 
                  Caption         =   "="
                  Height          =   255
                  Index           =   2
                  Left            =   1440
                  TabIndex        =   23
                  Top             =   1080
                  Width           =   135
               End
               Begin VB.Label Label2 
                  Caption         =   "="
                  Height          =   255
                  Index           =   1
                  Left            =   1440
                  TabIndex        =   22
                  Top             =   720
                  Width           =   135
               End
               Begin VB.Label Label2 
                  Caption         =   "="
                  Height          =   255
                  Index           =   0
                  Left            =   1440
                  TabIndex        =   21
                  Top             =   360
                  Width           =   135
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Lowest Value"
                  Height          =   435
                  Index           =   1
                  Left            =   240
                  TabIndex        =   20
                  Top             =   1680
                  Width           =   675
                  WordWrap        =   -1  'True
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Highest Value"
                  Height          =   390
                  Index           =   0
                  Left            =   240
                  TabIndex        =   19
                  Top             =   240
                  Width           =   720
                  WordWrap        =   -1  'True
               End
            End
         End
         Begin VB.Label lblNoPalette 
            Alignment       =   2  'Center
            Caption         =   "No Quintiles in this Report"
            Height          =   255
            Left            =   1440
            TabIndex        =   26
            Top             =   360
            Width           =   2295
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Display Cut Point Notes"
         Height          =   855
         Left            =   360
         TabIndex        =   2
         Top             =   4440
         Width           =   5295
         Begin VB.OptionButton OpCutPts 
            Caption         =   "Below the Data Table"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   4
            Top             =   360
            Value           =   -1  'True
            Width           =   2175
         End
         Begin VB.OptionButton OpCutPts 
            Caption         =   "Beside the Data Table"
            Height          =   255
            Index           =   0
            Left            =   2760
            TabIndex        =   3
            Top             =   360
            Width           =   2295
         End
      End
      Begin VB.Label lblReport 
         Caption         =   "1"
         Height          =   255
         Left            =   960
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Report #"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmLegends"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkDispersion_Click()
If chkDispersion.Value = 1 Then
    lblSigDig.Enabled = True
    txtSigDig.Enabled = True
Else
    lblSigDig.Enabled = False
    txtSigDig.Enabled = False
End If

End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub


Private Sub cmdOK_Click()
    Dim N As Integer
    
    N = Val(lblReport.Caption) - 1
    SaveSelections N
    
    Unload Me
End Sub
Private Sub SaveSelections(iRpt As Integer)
For I = 1 To 5
    ReportLegend(iRpt, 1, I) = txtLegend1(I - 1).Text
    ReportLegend(iRpt, 2, I) = txtLegend2(I - 1).Text
    ReportLegend(iRpt, 3, I) = txtLegend3(I - 1).Text
Next I
If chkDispersion.Value = 1 Then
    DoStat(iRpt) = True
    StatDig(iRpt) = Val(txtSigDig.Text)
Else
    DoStat(iRpt) = False
End If
If OpCutPts(0).Value = True Then
    CutPtLoc(iRpt) = "beside"
Else
    CutPtLoc(iRpt) = "below"
End If
If LayoutFlag Then
    WriteLayoutFile
    HTMLFile = Left(FNLOG, Len(FNLOG) - 4) & "_rpt" & CStr(iRpt + 1) & ".html"
    If SaveLayout Then WriteHTML HTMLFile, iRpt
End If
End Sub

Private Sub Form_Load()
Dim I As Integer

For I = 0 To 2
    frmLegends.SSTab1.TabCaption(I) = ""
Next I
End Sub

