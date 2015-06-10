VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmData 
   Caption         =   "Data Collection Grid"
   ClientHeight    =   8580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11865
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmData.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8580
   ScaleWidth      =   11865
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Data Collection Grid"
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11655
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   6855
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   12091
         _Version        =   393216
         ForeColor       =   16711680
         BackColorSel    =   12648447
         ForeColorSel    =   16711680
         SelectionMode   =   1
         AllowUserResizing=   1
      End
   End
   Begin VB.Frame Frame5 
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   7440
      Width           =   1815
      Begin VB.CommandButton cmdAddData 
         Caption         =   "Add Data to Collection"
         Height          =   495
         HelpContextID   =   105
         Left            =   240
         TabIndex        =   10
         ToolTipText     =   "Add more data to the Data Collection Grid"
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   5880
      TabIndex        =   6
      Top             =   7440
      Width           =   5895
      Begin VB.CommandButton cmdAddDefault 
         Caption         =   "Add Selection to Report   -- Batch Mode -- "
         Height          =   495
         HelpContextID   =   113
         Left            =   3120
         TabIndex        =   8
         ToolTipText     =   "Add selected items to the report and customize display options as a batch"
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton cmdAddEdit 
         Caption         =   "Add Selection to Report -- Individual Mode --  "
         Height          =   495
         HelpContextID   =   114
         Left            =   360
         TabIndex        =   7
         ToolTipText     =   "Add selected items to the report and customize display options row by row"
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   3600
      TabIndex        =   4
      Top             =   7440
      Width           =   2295
      Begin VB.CommandButton cmdPlot 
         Caption         =   "Plot Selected Item"
         Height          =   495
         HelpContextID   =   109
         Left            =   240
         TabIndex        =   5
         ToolTipText     =   "Create a plot of a single row of data in the Data Collection Grid"
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   1920
      TabIndex        =   2
      Top             =   7440
      Width           =   1695
      Begin VB.CommandButton cmdFind 
         Caption         =   "Search"
         Height          =   495
         HelpContextID   =   108
         Left            =   240
         TabIndex        =   3
         ToolTipText     =   "Search the Data Collection Grid"
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAddData_Click()
frmSpecUserAdded.Show vbModal
End Sub

Private Sub cmdAddDefault_Click()
Dim G As MSFlexGrid
Dim I As Integer
Dim K As Integer

Set G = MSFlexGrid1

If G.Rows < 2 Then
    MsgBox "There is No Data to Add to the Report." & vbCrLf & "Please Add Data to the Collection Before Continuing.", vbInformation, "Visual Report Designer"
    Exit Sub
End If

If G.Row < 1 Then
    MsgBox "Please Select an Item", vbInformation, "Visual Report Designer"
    Exit Sub
End If
If G.ColSel < G.Cols - 1 Then
    MsgBox "Please Select an Item", vbInformation, "Visual Report Designer"
    Exit Sub
End If
If G.Row > G.RowSel Then
    MsgBox "Please Select Items by Highlighting them from Top to Bottom" & vbCrLf & _
            "(Instead of from Bottom to Top)", vbInformation, "Visual Report Designer"
    Exit Sub
End If

If G.Row = G.RowSel Then
    MultiFlag = False
    MultiLastRow = 0
    SelectItem G.Row, 1
Else
    SelectItem G.Row, 2
End If

End Sub

Private Sub cmdFind_Click()
If MSFlexGrid1.Rows < 2 Then
    MsgBox "There is No Data to Search For." & vbCrLf & "Please Add Data to the Collection Before Continuing.", vbInformation, "Visual Report Designer"
    Exit Sub
End If
frmSearch.Show 'vbModal
End Sub
Private Sub cmdAddEdit_Click()
Dim G As MSFlexGrid
Dim I As Integer

If MSFlexGrid1.Rows < 2 Then
    MsgBox "There is No Data to Add to the Report." & vbCrLf & "Please Add Data to the Collection Before Continuing.", vbInformation, "Visual Report Designer"
    Exit Sub
End If

CancelFlag = False

Set G = MSFlexGrid1
If G.Row < 1 Then
    MsgBox "Please Select an Item", vbInformation, "Visual Report Designer"
    Exit Sub
End If
If G.ColSel < G.Cols - 1 Then
    MsgBox "Please Select an Item", vbInformation, "Visual Report Designer"
    Exit Sub
End If
If G.Row > G.RowSel Then
    MsgBox "Please Select Items by Highlighting them from Top to Bottom" & vbCrLf & _
            "(Instead of from Bottom to Top)", vbInformation, "Visual Report Designer"
    Exit Sub
End If

If G.Row = G.RowSel Then
    MultiFlag = False
    MultiLastRow = 0
    SelectItem G.Row, 1
Else
    MultiFlag = True
    MultiLastRow = G.RowSel
    SelectItem G.Row, 1
End If

End Sub
Private Sub cmdPlot_Click()
Dim G As MSFlexGrid

Set G = MSFlexGrid1

If G.Rows < 2 Then
    MsgBox "There is No Data to Plot." & vbCrLf & "Please Add Data to the Collection Before Continuing.", vbInformation, "Visual Report Designer"
    Exit Sub
End If

If G.Row < 1 Then
    MsgBox "Please Select an Item to Plot", vbInformation, "Visual Report Designer"
    Exit Sub
End If
If G.ColSel < G.Cols - 1 Then
    MsgBox "Please Select an Item to Plot", vbInformation, "Visual Report Designer"
    Exit Sub
End If
ShowChart
End Sub

Private Sub Form_Load()
Dim I As Integer
Dim J As Integer

ReDim BList(0 To 14, 1 To 1)
ReDim ReportInfo(0 To 14)

'set report defaults
ReDim ReportLegend(0 To 14, 1 To 3, 0 To 5)
ReDim CutPtLoc(0 To 14)
ReDim DoStat(0 To 14)
ReDim StatDig(0 To 14)
For I = 0 To 14
    CutPtLoc(I) = "below"
    DoStat(I) = True
    StatDig(I) = 1
    For J = 1 To 3
        ReportLegend(I, J, 1) = "Highest"
        ReportLegend(I, J, 2) = "2nd Highest"
        ReportLegend(I, J, 3) = "Middle"
        ReportLegend(I, J, 4) = "2nd Lowest"
        ReportLegend(I, J, 5) = "Lowest"
    Next J
Next I
    
MaxLines = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.mnuFileSaveDatabase.Enabled = False
End Sub
