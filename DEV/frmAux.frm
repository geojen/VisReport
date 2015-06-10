VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAux 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Added Data"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11805
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   106
   Icon            =   "frmAux.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   11805
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      ForeColor       =   &H00FF0000&
      Height          =   285
      HelpContextID   =   107
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   5400
      Width           =   2415
   End
   Begin VB.CommandButton cmdAddData 
      Caption         =   "Add to Data Collection"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4920
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Height          =   5535
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   6720
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   4920
         Width           =   2295
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   4455
         HelpContextID   =   107
         Left            =   240
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Enter Data in Grid Cells or Select a Cell to Paste Data from Clipboard"
         Top             =   360
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   7858
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   16711680
         BackColorSel    =   12648447
         ForeColorSel    =   16711680
         BackColorBkg    =   -2147483636
         GridLinesFixed  =   1
         AllowUserResizing=   1
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      HelpContextID   =   107
      Begin VB.Menu mnuUndo 
         Caption         =   "Undo"
         HelpContextID   =   107
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuSpc1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
         HelpContextID   =   107
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
         HelpContextID   =   107
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
         HelpContextID   =   107
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuSpc2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFillDown 
         Caption         =   "Fill Down"
         HelpContextID   =   107
      End
      Begin VB.Menu mnuFillRight 
         Caption         =   "Fill Right"
         HelpContextID   =   107
      End
   End
End
Attribute VB_Name = "frmAux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OldData() As String 'array to hold data for one undo
Dim DataDim(1 To 4) As Integer 'grid row and column selection for undo data
Dim UserDataFlag As Boolean 'flag to indicate whether there is data in OldData

Private Sub Form_Load()
    ' set text box properties and hide
    ' Use no border.
    Text1.BorderStyle = vbBSNone
    ' Match the grid's font.
    Text1.FontName = MSFlexGrid1.FontName
    Text1.FontSize = MSFlexGrid1.FontSize
    Text1.Visible = False
    
    Set FG = MSFlexGrid1
    
    'set undo data flag
    UserDataFlag = False
    
    AddUserDataForm = True 'flag to indicate frmAux is loaded
End Sub
Private Sub ReSizeCols()
Dim r As Integer
Dim c As Integer
Dim max_len As Single
Dim new_len As Single

    ' Size the columns.
    Font.Name = MSFlexGrid1.Font.Name
    Font.Size = MSFlexGrid1.Font.Size
    For c = 0 To MSFlexGrid1.Cols - 1
        max_len = 0
        For r = 0 To MSFlexGrid1.Rows - 1
            new_len = TextWidth(MSFlexGrid1.TextMatrix(r, c))
            If max_len < new_len Then max_len = new_len
        Next r
        MSFlexGrid1.ColWidth(c) = max_len + 240
        MSFlexGrid1.ColAlignment(c) = flexAlignLeftCenter
    Next c

    MSFlexGrid1.AllowUserResizing = flexResizeBoth
End Sub
Private Sub mnuCopy_Click()
CopyToClipboard
End Sub
Private Sub mnuCut_Click()
InitUndoData
CutGridData
End Sub

Private Sub mnuFillRight_Click()
Dim rtnval As Integer
Dim S As String

'If FG.ColSel = FG.Col Then
'    FG.ColSel = FG.Cols - 1
'    If FG.ColSel = FG.Col Then
'        S = "Value"
'    Else
'        S = "Values"
'    End If
'    rtnval = MsgBox("Fill Selected " + S + " to the End of the Row?", vbQuestion + vbOKCancel, "Visual Report Designer")
'    If rtnval = vbCancel Then
'        ClearGridSelection
'        Exit Sub
'    End If
'End If
FillRight
End Sub
Private Sub FillRight()
    Dim I As Integer
    Dim J As Integer
    Dim N As Integer
    Dim data() As String

    InitUndoData
    
    GetCellMinMax FG

    'get data to fill cells with
    N = FG.Col
    ReDim data(Rmin To Rmax)
    For I = Rmin To Rmax
        data(I) = FG.TextMatrix(I, N)
    Next I
    
    'if user hasn't selected which cells to fill, assume fill to right edge of grid
    If Cmin = Cmax Then Cmax = FG.Cols - 1
    
    'now fill the cells
    For I = Rmin To Rmax
        For J = Cmin To Cmax
            FG.TextMatrix(I, J) = data(I)
        Next J
    Next I

    ClearGridSelection
End Sub
Private Sub mnuPaste_Click()
GetPasteUndoData
PasteFromClipboard
End Sub
Private Sub mnuFillDown_Click()
Dim rtnval As Integer
Dim S As String

'If FG.RowSel = FG.Row Then
'    FG.RowSel = FG.Rows - 1
'    If FG.RowSel = FG.Row Then
'        S = "Value"
'    Else
'        S = "Values"
'    End If
'    rtnval = MsgBox("Fill Selected " + S + " to the Bottom of Grid?", vbQuestion + vbOKCancel, "Visual Report Designer")
'    If rtnval = vbCancel Then
'        ClearGridSelection
'        Exit Sub
'    End If
'End If
FillDown
End Sub
Private Sub FillDown()
    Dim I As Integer
    Dim J As Integer
    Dim N As Integer
    Dim data() As String
    
    InitUndoData
    
    GetCellMinMax FG
    
    'get data to fill cells with
    N = FG.Row
    ReDim data(Cmin To Cmax)
    For I = Cmin To Cmax
        data(I) = FG.TextMatrix(N, I)
    Next I
    
    'if user hasn't selected which cells to fill, assume fill to bottom of grid
    If Rmin = Rmax Then Rmax = FG.Rows - 1
    
    'now fill the cells
    For I = Rmin To Rmax
        For J = Cmin To Cmax
            FG.TextMatrix(I, J) = data(J)
        Next J
    Next I
    
    ClearGridSelection
End Sub

Private Sub GridEdit1(KeyAscii As Integer)
    InitUndoData
    
    ' Position the TextBox over the cell.
    Text1.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left + Frame1.Left
    Text1.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top + Frame1.Top
    Text1.Width = MSFlexGrid1.CellWidth
    Text1.Height = MSFlexGrid1.CellHeight
    Text1.Visible = True
    Text1.SetFocus
    
    'disable flexgrid to prevent user from scrolling
    'the grid and mis-aligning the text box
    'MSFlexGrid1.Enabled = False

    Select Case KeyAscii
        Case 0 To Asc(" ")
            Text1.Text = MSFlexGrid1.Text
            Text1.SelStart = Len(Text1.Text)
        Case Else
            Text1.Text = Chr$(KeyAscii)
            Text1.SelStart = 1
    End Select
End Sub
Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button And vbRightButton Then PopupMenu mnuEdit
End Sub
Private Sub Text1_LostFocus()
    
'    If vbKeyTab Then
'      MsgBox "This is the TAB Key"
'    End If

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
        
    Select Case KeyCode
        Case vbKeyTab
            ' Finish editing and move to the right one column.
            FG.SetFocus
            DoEvents
            If FG.Col < FG.Cols - 1 Then
                FG.Col = FG.Col + 1
            ElseIf FG.Col = FG.Cols - 1 Then
                If FG.Row < FG.Rows - 1 Then
                    FG.Col = FG.FixedCols
                    FG.Row = FG.Row + 1
                    FixRowVis
                End If
            End If
            FixColVis
            
        Case vbKeyEscape
            ' Leave the text unchanged.
            Text1.Visible = False
            FG.SetFocus

        Case vbKeyReturn
            ' Finish editing and move down one row.
            FG.SetFocus
            DoEvents
            If FG.Row < FG.Rows - 1 Then
                FG.Row = FG.Row + 1
            ElseIf FG.Row = FG.Rows - 1 Then
                If FG.Col < FG.Cols - 1 Then
                    FG.Row = FG.FixedRows
                    FG.Col = FG.Col + 1
                    FixColVis
                End If
            End If
            FixRowVis

        Case vbKeyDown
            ' Move down 1 row.
            FG.SetFocus
            DoEvents
            If FG.Row < FG.Rows - 1 Then
                FG.Row = FG.Row + 1
            End If
            FixRowVis

        Case vbKeyUp
            ' Move up 1 row.
            FG.SetFocus
            DoEvents
            If FG.Row > FG.FixedRows Then
                FG.Row = FG.Row - 1
            End If
                    
    End Select

    
End Sub
Private Sub FixRowVis()
    'make sure row doesn't fall off page
    If FG.Row = FG.FixedRows Then
        FG.TopRow = FG.FixedRows
    ElseIf FG.Row = FG.Rows - 1 Then
        FG.TopRow = FG.TopRow + 1
    Else
        If FG.RowIsVisible(FG.Row + 1) = False Then
            FG.TopRow = FG.TopRow + 1
        End If
    End If
End Sub
Private Sub FixColVis()
    'make sure column doesn't fall off page
    If FG.Col = FG.FixedCols Then
        FG.LeftCol = FG.FixedCols
    ElseIf FG.Col = FG.Cols - 1 Then
        FG.LeftCol = FG.LeftCol + 1
    Else
        If FG.ColIsVisible(FG.Col + 1) = False Then
            FG.LeftCol = FG.LeftCol + 1
        End If
    End If
End Sub
' Do not beep on Return or Escape.
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) Or _
           (KeyAscii = vbKeyTab) Then KeyAscii = 0
End Sub
Private Sub MSFlexGrid1_DblClick()
    GridEdit1 Asc(" ")
End Sub

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyTab Then
        DoEvents
        If FG.Col < FG.Cols - 1 Then
            FG.Col = FG.Col + 1
        ElseIf FG.Col = FG.Cols - 1 Then
            If FG.Row < FG.Rows - 1 Then
                FG.Col = FG.FixedCols
                FG.Row = FG.Row + 1
                FixRowVis
            End If
        End If
        FixColVis
        Exit Sub
    End If
    GridEdit1 KeyAscii
End Sub

Private Sub MSFlexGrid1_LeaveCell()
    If Text1.Visible Then
        MSFlexGrid1.Text = Text1.Text
        Text1.Visible = False
    End If
End Sub
Private Sub MSFlexGrid1_GotFocus()
    If Text1.Visible Then
        MSFlexGrid1.Text = Text1.Text
        Text1.Visible = False
    End If
End Sub
Private Sub mnuUndo_Click()
    Dim I As Integer
    Dim J As Integer
    Dim NR As Integer
    Dim NC As Integer
    
    If UserDataFlag Then
        NR = DataDim(2) - DataDim(1) + 1
        NC = DataDim(4) - DataDim(3) + 1
        For I = 1 To NR
            For J = 1 To NC
                FG.TextMatrix(DataDim(1) + I - 1, DataDim(3) + J - 1) = OldData(I, J)
            Next J
        Next I
        
        ReDim OldData(1 To 1, 1 To 1)
        For I = 1 To 4
            DataDim(I) = 0
        Next I
        UserDataFlag = False
        mnuUndo.Enabled = False
    End If
        
End Sub
Private Sub GetPasteUndoData()
Dim I As Integer
Dim S As String
Dim N As Integer
Dim NR As Integer
Dim NC As Integer
Dim Count As Integer


    S = Clipboard.GetText
    'add carriage return/line feed at end of string, if not present
    If Right(S, 2) <> vbCrLf Then S = S + vbCrLf
    
    NR = 0
    NC = 0
    
    Do While S <> ""
        N = InStr(S, vbCrLf)
        If N > 0 Then
            NR = NR + 1
            Count = 1
            For I = 1 To N - 1
                T = Mid(S, I, 1)
                If T = Chr(9) Then Count = Count + 1
            Next I
            If Count > NC Then NC = Count
            S = Mid(S, N + 2)
            If S = "" Then Exit Do
        Else
            Exit Do
        End If
    Loop
    
    DataDim(1) = FG.Row
    N = FG.Row + NR - 1
    If N > FG.Rows - 1 Then
        N = FG.Rows - 1
        NR = N - FG.Row + 1
    End If
    DataDim(2) = N
    DataDim(3) = FG.Col
    N = FG.Col + NC - 1
    If N > FG.Cols - 1 Then
        N = FG.Cols - 1
        NC = N - FG.Col + 1
    End If
    DataDim(4) = N
    
    GetUndoData
End Sub
Private Sub InitUndoData()
    
    DataDim(1) = FG.Row
    DataDim(2) = FG.RowSel
    DataDim(3) = FG.Col
    DataDim(4) = FG.ColSel
    GetUndoData
End Sub
Private Sub GetUndoData()
    Dim I As Integer
    Dim J As Integer
    Dim NR As Integer
    Dim NC As Integer
    
    NR = DataDim(2) - DataDim(1) + 1
    NC = DataDim(4) - DataDim(3) + 1
    ReDim OldData(1 To NR, 1 To NC)
    
    For I = 1 To NR
        For J = 1 To NC
            OldData(I, J) = FG.TextMatrix(DataDim(1) + I - 1, DataDim(3) + J - 1)
        Next J
    Next I
    UserDataFlag = True
    mnuUndo.Enabled = True
    
End Sub
Private Sub cmdAddData_Click()
Dim S As String
Dim T As String
Dim N As Integer
Dim I As Integer
Dim J As Integer
Dim G As MSFlexGrid

    'make sure last change to grid is captured
    If Text1.Visible Then
        MSFlexGrid1.Text = Text1.Text
        Text1.Visible = False
    End If
    
    'check for commas
    Set G = MSFlexGrid1
    For I = 1 To G.Rows - 1
        For J = 1 To G.Cols - 1
            S = G.TextMatrix(I, J)
            N = InStr(S, ",")
            If N > 0 Then
                T = "Data Cannot Contain Commas." & vbCrLf & "See "
                If J > 3 Then
                    T = T & "Year " & G.TextMatrix(0, J)
                Else
                    T = T & G.TextMatrix(0, J) & " Description"
                End If
                T = T & ", Row " & CStr(I) & "."
                MsgBox T, vbInformation, "Visual Report Designer"
                Exit Sub
            End If
        Next J
    Next I
    
    AddUserDataToGrid
End Sub
Private Sub cmdCancel_Click()
Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
AddUserDataForm = False 'flag to indicate frmAux is not loaded
End Sub

