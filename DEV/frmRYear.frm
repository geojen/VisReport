VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRYear 
   Caption         =   "Report Year Range Specification"
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
   Icon            =   "frmRYear.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8580
   ScaleWidth      =   11865
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   6255
      Left            =   600
      TabIndex        =   0
      Top             =   960
      Width           =   9615
      Begin VB.CommandButton cmdOK 
         Caption         =   "Save Report File"
         Height          =   375
         Left            =   3360
         TabIndex        =   2
         Top             =   5280
         Width           =   3135
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   4095
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   7223
         _Version        =   393216
         ForeColor       =   16711680
         BorderStyle     =   0
         Appearance      =   0
      End
   End
End
Attribute VB_Name = "frmRYear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    WriteLayoutFile
End Sub

Private Sub Form_Load()
    Dim G As MSFlexGrid
    Dim I As Integer
    
    Set G = MSFlexGrid1
    G.Rows = NForms + 1
    G.Cols = 4
    G.TextArray(FGIndex(G, 0, 0)) = "Report #"
    G.TextArray(FGIndex(G, 0, 1)) = "Start Year"
    G.TextArray(FGIndex(G, 0, 2)) = "End Year"
    G.TextArray(FGIndex(G, 0, 3)) = "Report Title"
    
    G.ColWidth(1) = 1500
    G.ColWidth(2) = 1500
    G.ColWidth(3) = 5000
    
    For I = 1 To NForms
        G.TextArray(FGIndex(G, I, 0)) = CStr(I)
        G.TextArray(FGIndex(G, I, 1)) = frmGeneral.txtYear(0).Text
        G.TextArray(FGIndex(G, I, 2)) = frmGeneral.txtYear(1).Text
    Next I
    FF2 = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FF2 = False
End Sub

Private Sub MSFlexGrid1_Click()
    InitGridEditForm MSFlexGrid1, 0
End Sub
