VERSION 5.00
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Visual Report Designer - Version 1.6.1"
   ClientHeight    =   8745
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11880
   HelpContextID   =   101
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileCreateNew 
         Caption         =   "Create New Project"
         HelpContextID   =   102
      End
      Begin VB.Menu mnuFileSpc5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOpenProjectLog 
         Caption         =   "Open Existing Project"
         HelpContextID   =   103
      End
      Begin VB.Menu mnuFileSpc3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save Project As ..."
         HelpContextID   =   104
      End
      Begin VB.Menu mnuFileSpc2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSaveDatabase 
         Caption         =   "Save Grid Data in Database Format"
         HelpContextID   =   111
      End
      Begin VB.Menu mnuFileSpc1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileBackup 
         Caption         =   "Backup"
         HelpContextID   =   126
         Begin VB.Menu mnuBackupNow 
            Caption         =   "Backup Entire Project Now"
            HelpContextID   =   126
         End
         Begin VB.Menu mnuBackupRestore 
            Caption         =   "Restore from Backup ..."
            HelpContextID   =   127
         End
         Begin VB.Menu mnuBackupSpc1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuBackupClear 
            Caption         =   "Clear All Backups"
            HelpContextID   =   128
         End
      End
      Begin VB.Menu mnuFileSpc4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuData 
      Caption         =   "Data"
      Begin VB.Menu mnuDataAdd 
         Caption         =   "Add Data to the Project"
         Begin VB.Menu mnuDataAgePro 
            Caption         =   "Import AgePro Model Results"
            HelpContextID   =   105
         End
         Begin VB.Menu mnuDataspc1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDataAim 
            Caption         =   "Import AIM Model Results"
            HelpContextID   =   105
         End
         Begin VB.Menu mnuDataspc2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDataASAP 
            Caption         =   "Import ASAP Model Results"
            HelpContextID   =   105
         End
         Begin VB.Menu mnuDataspc3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDataASPIC 
            Caption         =   "Import ASPIC Model Results"
            HelpContextID   =   105
         End
         Begin VB.Menu mnuDataSpc5 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDataCSA 
            Caption         =   "Import CSA Model Results"
            HelpContextID   =   105
         End
         Begin VB.Menu mnuDataSpc6 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDataVPa 
            Caption         =   "Import VPA Model Results"
            HelpContextID   =   105
         End
         Begin VB.Menu mnuDataSpc4 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDataSamp 
            Caption         =   "Import SAGA Sample Length Weight Data"
            HelpContextID   =   105
         End
         Begin VB.Menu mnuDataSp7 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDataAux 
            Caption         =   "Add User Supplied Data"
            HelpContextID   =   106
         End
      End
      Begin VB.Menu mnuDataSp8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDataMissing 
         Caption         =   "Missing Value Indicator"
         Begin VB.Menu mnuDataNAVal 
            Caption         =   "Specify Missing Value Indicator"
            HelpContextID   =   129
         End
         Begin VB.Menu mnuDataMissingReplace 
            Caption         =   "Global Replace"
            HelpContextID   =   130
         End
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuOptionsViewer 
         Caption         =   "Select Report Viewer"
         HelpContextID   =   116
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "Window"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuHelpSpc1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpUsing 
         Caption         =   "Using Visual Report Designer"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
    Dim I As Integer
    Dim J As Integer
    
    App.HelpFile = App.Path + "\VRHELP.CHM"
    
    'disable menu items
    mnuFileSaveDatabase.Enabled = False
    mnuDataAdd.Enabled = False
    mnuDataMissingReplace.Enabled = False
    
    ReadCfg
    
    'hardwire the number of backups for now
    NBackups = 20

    FF1 = False
    FF3 = False
    FF4 = False
    
    LogFlag = False
    
    FNLOG = ""
    
    DataFlag = False 'flag to indicate presence of data loaded
    AddUserDataForm = False 'flag to indicate whether frmAux is loaded
    
    frmGeneral.Show
    
    'set default number of forms
    NForms = 15
    
    'initialize years
    StartYear = 0
    EndYear = 0
    NYears = 0
    
    'set report preferences and defaults
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
    
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim S As String
    
    'Backup the current file before unloading, only if the command did NOT come
    'from the code (e.g, from clicking the X button on the form, etc).
    If UnloadMode <> vbFormCode Then
        'first make sure there is a file available
        S = Trim(frmGeneral.lblFile.Caption)
        If S <> "" And LCase(S) <> "none specified" Then
            'do backup
            If UnloadMode = vbFormControlMenu Then
                CreateBackup S, "File-Exit"
            Else
                CreateBackup S, "Close File"
            End If
        End If
    End If

'    Dim RS As Integer
'
'
'    If DataFlag Then
'        RS = MsgBox("If You Have Added Data to the Project" + vbCrLf + _
'                    "You Should Save the Data Grid to CSV File Before Exit" + vbCrLf + _
'                    "Do You Wish to Save Data Now?", vbQuestion + vbYesNoCancel, "Visual Report Designer")
'        If RS = vbYes Then
'            Cancel = 0
'            WriteGridData
'            If DataSavedFlag Then MsgBox "Grid Data Has Been Saved to File", vbInformation, "Visual Report Designer"
'        ElseIf RS = vbNo Then
'            Cancel = 0
'        Else
'            Cancel = 1
'        End If
'    End If
'
'    If LayoutFlag Then
'        RS = MsgBox("Save Report Layout File Before Exit", vbQuestion + vbYesNoCancel, "Visual Report Designer")
'        If RS = vbYes Then
'            WriteLayoutFile
'            If SaveLayout Then MsgBox "Report Layout File has Been Saved", vbInformation, "Visual Report Designer"
'            If SaveLayout Then
'                Cancel = 0
'            Else
'                Cancel = 1
'            End If
'        ElseIf RS = vbNo Then
'            Cancel = 0
'        Else
'            Cancel = 1
'        End If
'    End If
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    
    Close #3
    
    If FF1 Then
        FF1 = False
        Unload RF1
        Set RF1 = Nothing
    End If
    
        
    If FF3 Then
        FF3 = False
        Unload RF3
        Set RF3 = Nothing
    End If
    
    If FF4 Then
        FF4 = False
        Unload RF4
        Set RF4 = Nothing
    End If
    Unload frmChart
    Unload frmAux
End Sub

Private Sub mnuBackupClear_Click()
Dim rtnval As Integer

rtnval = MsgBox("Warning: This Will Delete All Backups." & vbCrLf & _
        "Do You Wish To Proceed?", vbInformation + vbOKCancel, "Visual Report Designer")
If rtnval = vbCancel Then
    Exit Sub
Else
    DeleteBackups
End If

End Sub

Private Sub mnuBackupNow_Click()
Dim S As String

'first make sure there is a file available
S = Trim(frmGeneral.lblFile.Caption)
If S = "" Or LCase(S) = "none specified" Then
    MsgBox "Please Create a New Project Before Proceeding", vbOKOnly + vbQuestion, "Visual Report Designer"
    Exit Sub
End If

'do backup
CreateBackup S, "User"

'give user feedback upon successful completion
If Not RedFlag Then
    MsgBox "Backup Successfully Completed", vbOKOnly + vbInformation, "Visual Report Designer"
End If
End Sub

Private Sub mnuBackupRestore_Click()
    InitRestore
End Sub

Private Sub mnuDataAgePro_Click()
    OpenDataFile 6
End Sub

Private Sub mnuDataAim_Click()
    OpenDataFile 5
End Sub

Private Sub mnuDataASAP_Click()
    OpenDataFile 4
End Sub

Private Sub mnuDataASPIC_Click()
    OpenDataFile 3
End Sub

Private Sub mnuDataAux_Click()
    Dim S As String
    Dim I As Integer
    
    frmSpecUserAdded.opType(7).Value = True
    frmSpecUserAdded.frmUser.Visible = True
    frmSpecUserAdded.Show vbModal, Me
    
    're-set the focus because otherwise the Window menu loses all the MDI child forms
    S = Me.ActiveForm.Caption
    If S = "General Information" Then
        frmGeneral.cmdOpen.SetFocus
        frmGeneral.Picture1.SetFocus
    ElseIf S = "Data Collection Grid" Then
        frmData.cmdFind.SetFocus
        frmData.MSFlexGrid1.SetFocus
    ElseIf S = "Report Design and Layout" Then
        I = frmCompose.SSTab1.Tab
        frmCompose.txtTitle(I).SetFocus
        frmCompose.SSTab1.SetFocus
    End If
    If AddUserDataForm Then frmAux.ZOrder
End Sub

Private Sub mnuDataCSA_Click()
    OpenDataFile 2
End Sub

Private Sub mnuDataMissingReplace_Click()
frmData.ZOrder 'show the data form
frmReplaceNA.Show vbModal

End Sub

Private Sub mnuDataSamp_Click()
    OpenDataFile 7
End Sub

Private Sub mnuDataVPa_Click()
    OpenDataFile 1
End Sub

Private Sub mnuFileCreateNew_Click()
    Dim S As String
    
    'backup the current file before unloading
    'first make sure there is a file available
    S = Trim(frmGeneral.lblFile.Caption)
    If S <> "" And LCase(S) <> "none specified" Then
        'do backup
        CreateBackup S, "Close File"
    End If
    
    CreateNewLogFile
End Sub

Private Sub mnuFileExit_Click()
    Dim S As String
    
    'backup the current file before unloading
    'first make sure there is a file available
    S = Trim(frmGeneral.lblFile.Caption)
    If S <> "" And LCase(S) <> "none specified" Then
        'do backup
        CreateBackup S, "File-Exit"
    End If
    
    Unload frmMain
End Sub

Private Sub mnuFileOpenProjectLog_Click()
    Dim S As String
    
    'backup the current file before unloading
    'first make sure there is a file available
    S = Trim(frmGeneral.lblFile.Caption)
    If S <> "" And LCase(S) <> "none specified" Then
        'do backup
        CreateBackup S, "Close File"
    End If
    
    OpenProjectLog
End Sub

Private Sub mnuFileSaveAs_Click()
SaveProjectAs
End Sub

Private Sub mnuFileSaveDatabase_Click()
    WriteDatabaseStyle
End Sub

Private Sub mnuHelpAbout_Click()
frmAbout.Show vbModal
End Sub

Private Sub mnuHelpUsing_Click()
Shell "hh.exe " & App.HelpFile, vbMaximizedFocus
End Sub

Private Sub mnuDataNAVal_Click()
Dim S As String
Dim I As Integer

frmMissing.Show vbModal, Me

're-set the focus because otherwise the Window menu loses all the MDI child forms
S = Me.ActiveForm.Caption
If S = "General Information" Then
    frmGeneral.cmdOpen.SetFocus
    frmGeneral.Picture1.SetFocus
ElseIf S = "Data Collection Grid" Then
    frmData.cmdFind.SetFocus
    frmData.MSFlexGrid1.SetFocus
ElseIf S = "Report Design and Layout" Then
    I = frmCompose.SSTab1.Tab
    frmCompose.txtTitle(I).SetFocus
    frmCompose.SSTab1.SetFocus
End If

If LayoutFlag Then
    WriteLayoutFile
    If SaveLayout Then
        For I = 0 To 14
            HTMLFile = Left(FNLOG, Len(FNLOG) - 4) & "_rpt" & CStr(I + 1) & ".html"
            WriteHTML HTMLFile, I
        Next I
    End If
End If

End Sub

Private Sub mnuOptionsViewer_Click()
Dim S As String
Dim I As Integer

frmViewOpt.Show vbModal, Me

're-set the focus because otherwise the Window menu loses all the MDI child forms
S = Me.ActiveForm.Caption
If S = "General Information" Then
    frmGeneral.cmdOpen.SetFocus
    frmGeneral.Picture1.SetFocus
ElseIf S = "Data Collection Grid" Then
    frmData.cmdFind.SetFocus
    frmData.MSFlexGrid1.SetFocus
ElseIf S = "Report Design and Layout" Then
    I = frmCompose.SSTab1.Tab
    frmCompose.txtTitle(I).SetFocus
    frmCompose.SSTab1.SetFocus
End If

End Sub
