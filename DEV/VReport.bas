Attribute VB_Name = "VReport"
Option Explicit
' ======== GENERAL NOTES =======================================
' There is a global variable called "VRVersion" to help automate
' the version information throughout the GUI.
'
' 1. When making updates to the ScanModel.exe programs, be sure to edit
'    the help topic "Supported Model Details" (in the "Adding Data to the
'    Collection" book). Include the most recent version of the NFT model
'    that Visual Report can handle and the list of items that are imported.
' 2. Remember to:
'    - Change the date on "frmAbout".
'    - Update ReadMe.txt to include any changes you make to the
'      GUI. Version history changes go in Section 5 "VERSION CHANGES"
'      towards the bottom of the file.
'    - Change the version number on the Inno Setup script.
'    - Increment the version in the VB Project Properties.
' 3. Be sure the zip file to post on the web site includes the ReadMe.txt
'    file in addition to the setup.exe installer.
' 4. Also, add the text you insert in the ReadMe.txt file to the web
'    page for Visual Report Designer (Visual_Repoort_Designer.htm).
' ================================================================

' ======== NOTES (2/14/08) =======================================
' I made some minor modifications to ScanAspic.exe. Use the C code
' in the folder "ScanAspic" (located in this directory) to modify
' when ASPIC changes. The version on Alan's computer is outdated.
' =================================================================
'
' ======= Jen's To Do List ========================================
'
'HIGH PRIORITY
'
'-For importing SAGA data: Remind Al about the need to export a species list
' ("species.txt") when the user exports length-weight data. (2/26/08)
'
'-Add the ability to delete data collection data. Probably would need to start
' logging or tagging each batch of data added to the collection with a unique
' identifier so that you know which chunks to delete.(2/13/08)
'
'
'MEDIUM PRIORITY
'
'-Add the ability to drag-with-the-mouse to select an area in the report to copy.
'
'-Incorporate a less confusing file management system. Idea 1: Go to a
' single-document interface with 3 tabs instead of MDI with 3 forms?
' Idea 2 (from Michele): "Maybe make a "close project" button.  You know how in
' Excel, if you open another worksheet, the other one stays in the background
' (unless you 'x' it out), I'm always afraid to open another project (or start
' another one) without closing the original one.  I'm sure it's fine, however,
' it's the fear of changing things without meaning to or closing without saving...
' something to that effect.  Basically a way to close the project entirely
' before starting another one." (michelle wish list, 2/5/08))
'
'-In the log file, write out which positions in the data grid were replaced when
' doing a global replace of no data or missing data values. (2/26/08)
'
'-Follow up on ways to implement Paul Rago's suggestion for sparknotes. (2/21/08)
' Could implement the whisker style (aka baseball scores) fairly easily. See the
' file "TO DO symbols sparklines.doc" for details.
'
'-In Page Preferences view, don't show cut point notes location option unless there
' are cut point notes to display.
'
'-Clarify what cut points/notes are; add definitions using Tool Tip text or popups.
'
'-Re-do red-yellow-green symbols so that there is no black border. Think about
' re-doing the report table so that there are no black borders between grid cells,
' or at the very least make the borders a very light shade of gray. (Tip, test out
' the table styles in NVu first, then copy the HTML code to Visual Report).
' Experts say that your eyes are naturally drawn more to the white space than the
' symbol when you have borders around images like these. See
' http://www.edwardtufte.com/bboard/q-and-a-fetch-msg?msg_id=0001OR
' for examples to illustrate why this is so, and info on the theories behind this.
'
'
'LOW PRIORITY
'
'-Finesse the auto-backup feature. Have a user-selected number of backups. Include
' a checkbox for which actions should cause an auto-backup. (1/24/08)
'
'-Don't have the "Add Data to Collection" button on Data collection grid;
' or put elsewhere.
'
'-Make the Report Viewer a fixed window (or an additional tab). Would need to
' implement the mouse-select to copy function. (michelle wish list, 2/5/08)
'
'-Speed up the time it takes users to enter the report title in the text box.
' Network slowness causing severe lag time. (michelle wish list, 2/5/08)
'
'-unable to import AIM 1.5.2 files; files not current. Can't reproduce error.
' (sue wigley, 2/8/08)
'
'-bad file name or number when editing report title. possibly because writing
' to a file on the network which caused the write procedure to slow down, and
' when user clicked view report or page preferences, it couldn't write.
' (Michelle, 2/5/08)
' =================================================================

Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type
Public Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Private Type VPASurvey
        Tag As String * 16
        StartAge As Integer
        EndAge  As Integer
        Time As Integer
        Type As Integer
End Type

Private Type ASAPSurvey
        Tag As String * 16
        StartAge As Integer
        EndAge  As Integer
        Units As Integer
        Month As Integer
        Used As Integer
End Type

Private Type DataLayout
    Tag As String
    StartYear As Integer
    EndYear As Integer
    Key As Integer
    Type As Integer
    'HighFlag As Boolean 'HighFlag = True means high values are good (or black/green) and low values are bad (or red).
    Palette As String
    LowerCut As Double
    UpperCut As Double
    ZeroFlag As Boolean
End Type

Private Type RptList
    Title As String
    NLines As Integer
End Type

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function CreateConsoleProcess Lib "CUTIL.DLL" (ByVal S As String) As Long
'Private Declare Function CreateConsoleProcess Lib "C:\Develop\NFT\VisualReport_1.6\MDIVBDEV\CUTIL.DLL" (ByVal S As String) As Long
Public Declare Function CompareFileTime Lib "kernel32" (lpFileTime1 As FILETIME, lpFileTime2 As FILETIME) As Long
Public Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Public Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const OPEN_EXISTING = 3

Private JToken As Integer


' Global Variables
Public Const VRVersion = "1.6.1" 'version number of Visual Report Designer

'configuration file global variables
'user preferences for what application to view report
'"GUI"=in the GUI, "Browser"=computer's default browser, path=path of selected application
Public RptViewer As String
Public NAVal As String 'the missing-value indicator

Public NBackups As Integer 'number of backup copies to save
Public NYears As Integer
Public StartYear As Integer
Public EndYear As Integer
Public NForms As Integer
Public KRow As Integer
Public CaseID As String
Public FNLOG As String
Public FName As String
Public FNOUT As String
Public FNTXT As String
Public HTMLFile As String 'the name/location of the HTML report file to view
Public KYears As Integer
Public KAges As Integer
Public KFish As Integer
Public KFYear As Integer
Public KXYear As Integer
Public KFAge As Integer
Public KINDX As Integer
Public KModel As Integer
Public KAge1 As Integer
Public KAge2 As Integer
Public BList() As DataLayout 'stores the report layout details, by line (dimensioned with 0 as first element)
Public ReportInfo() As RptList 'stores each reports general details (dimensioned with 0 as first element)
Public MaxLines As Integer 'the temporary maximum number of lines for all reports
Public ModelType As String
Public RedFlag As Boolean 'generic warning flag for something not completed
Public JList As Integer
Public RF1 As Form
Public FF1 As Boolean
Public ReportLegend() As String
'where to put the cut point notes; below ("below") or to the side ("beside") of the table
Public CutPtLoc() As String
'user preference for whether to display the dispersion statistic for each line of data in each report
Public DoStat() As Boolean
'number of significat digits in disperson statsitic
Public StatDig() As Integer
Public NDataEdit As Integer 'number of lines in a report to edit at once
Public DataEdit() As DataLayout 'temporary storage of report data to edit at once

Public AgeFlag As Boolean
Public CSAFlag As Boolean
Public AspicFlag As Boolean
Public ASAPFlag As Boolean
Public AIMFlag As Boolean
Public AgeProFlag As Boolean

Public MinYear As Integer
Public MaxYear As Integer

Public LogFlag As Boolean 'flag for whether or not log file is loaded
Public DataSavedFlag As Boolean 'flag to indicate whether data has been saved to csv or not
Public DataFlag As Boolean 'flag for whether or not data file is loaded
Public LayoutFlag As Boolean 'flag for whether or not layout file is loaded
Public SaveLayout As Boolean
Public AddUserDataForm As Boolean 'flag to indicate whether frmAux is loaded

'flag to cancel out of add data loop
Public CancelFlag As Boolean

'flag to suppress message boxes on initialization of forms
Public InitFlag As Boolean
'flag to indicate data entry is in multi-line mode
Public MultiFlag As Boolean
'keep track of the last row in a multi-select add data mode
Public MultiLastRow As Integer

Public Const KRows = 30
Public Const KCols = 45

Public Function GetFirstToken(S As String) As String
    Dim I As Integer
    Dim L As Integer
    Dim N As Integer
    Dim T As String
    
    L = Len(S)
    N = 0
    JToken = 0
    For I = 1 To L
        If Mid(S, I, 1) <> " " Then
            N = I
            Exit For
        End If
    Next I
    
    If N = 0 Then
        GetFirstToken = ""
        Exit Function
    End If
    
    
    For I = N To L
        If Mid(S, I, 1) = " " Then
            JToken = I
            Exit For
        End If
    Next I
    
    If JToken = 0 Then
        GetFirstToken = Mid(S, N, L - N + 1)
    Else
        GetFirstToken = Mid(S, N, JToken - N)
    End If
    
End Function
Public Function GetNextToken(S As String) As String
    Dim I As Integer
    Dim L As Integer
    Dim N As Integer
    Dim T As String
    
    L = Len(S)
    N = 0
    
    If JToken = 0 Then
        GetNextToken = ""
        Exit Function
    End If
    
    
    For I = JToken To L
        If Mid(S, I, 1) <> " " Then
            N = I
            Exit For
        End If
    Next I
    
    If N = 0 Then
        GetNextToken = ""
        Exit Function
    End If
    
    
    For I = N To L
        If Mid(S, I, 1) = " " Then
            JToken = I
            Exit For
        End If
    Next I
    
    If JToken = 0 Or JToken < N Then
        GetNextToken = Mid(S, N, L - N + 1)
        JToken = 0
    Else
        GetNextToken = Mid(S, N, JToken - N)
    End If
    
    
End Function
Public Function FGIndex(G As MSFlexGrid, r As Integer, c As Integer) As Long
    FGIndex = r * G.Cols + c
End Function

Public Sub OpenDataFile(NType As Integer)
    
    On Error GoTo OpenErr
    
    Unload frmSpecUserAdded
    
    If NType = 3 Then
        frmGeneral.CommonDialog1.Filter = "INP Files (*.inp)|*.inp"
    ElseIf NType = 6 Then
        frmGeneral.CommonDialog1.Filter = "IN Files (*.in)|*.in"
    ElseIf NType = 7 Then
        frmGeneral.CommonDialog1.Filter = "OUT Files (*.out)|*.out"
    Else
        frmGeneral.CommonDialog1.Filter = "DAT Files (*.dat)|*.dat"
    End If
    frmGeneral.CommonDialog1.FileName = ""
    frmGeneral.CommonDialog1.Flags = &H1004
    
    Select Case NType
        Case 1
            frmGeneral.CommonDialog1.DialogTitle = "Open VPA/ADAPT Model Input Data File"
        Case 2
            frmGeneral.CommonDialog1.DialogTitle = "Open CSA Model Input Data File"
        Case 3
            frmGeneral.CommonDialog1.DialogTitle = "Open ASPIC Model Input Data File"
        Case 4
            frmGeneral.CommonDialog1.DialogTitle = "Open ASAP Model Input Data File"
        Case 5
            frmGeneral.CommonDialog1.DialogTitle = "Open AIM Model Input Data File"
        Case 6
            frmGeneral.CommonDialog1.DialogTitle = "Open AgePro Model Input Data File"
        Case 7
            frmGeneral.CommonDialog1.DialogTitle = "Open Sample Length Weight Report File"
    End Select
    
    frmGeneral.CommonDialog1.CancelError = True
    frmGeneral.CommonDialog1.FilterIndex = 0
    frmGeneral.CommonDialog1.ShowOpen
    FName = frmGeneral.CommonDialog1.FileName
    If FName <> "" Then
        ScanData NType
    End If
    Exit Sub
    
OpenErr:
    Exit Sub
End Sub
Private Sub ScanData(NType As Integer)

    Dim Cmdline As String
    Dim FN As String
    Dim KL As Long
    Dim flag As Boolean
    Dim N As Integer
    
    On Error GoTo ScanErr
    
    Select Case NType
        Case 1
            flag = CheckVPAFiles
            If flag = False Then
                Exit Sub
            End If
            Cmdline = App.Path + "\scanvpa.exe " + Chr(34) + FName + Chr(34)
            KL = CreateConsoleProcess(Cmdline)
            ScanVPAGeneralData
        Case 2
            flag = CheckCSAFiles
            If flag = False Then
                Exit Sub
            End If
            Cmdline = App.Path + "\scancsa.exe " + Chr(34) + FName + Chr(34)
            KL = CreateConsoleProcess(Cmdline)
            ScanCSAGeneralData
        Case 3
            flag = CheckAspicFiles
            If flag = False Then
                Exit Sub
            End If
            Cmdline = App.Path + "\scanaspic.exe " + Chr(34) + FName + Chr(34)
            KL = CreateConsoleProcess(Cmdline)
            ScanAspicGeneralData
        Case 4
            flag = CheckASAPFiles
            If flag = False Then
                Exit Sub
            End If
            Cmdline = App.Path + "\scanasap.exe " + Chr(34) + FName + Chr(34)
            KL = CreateConsoleProcess(Cmdline)
            ScanAsapGeneralData
        Case 5
            flag = CheckAIMFiles
            If flag = False Then
                Exit Sub
            End If
            Cmdline = App.Path + "\scanaim.exe " + Chr(34) + FName + Chr(34)
            KL = CreateConsoleProcess(Cmdline)
            ScanAIMGeneralData
        Case 6
            flag = CheckAgeProFiles
            If flag = False Then
                Exit Sub
            End If
            Cmdline = App.Path + "\scanagepro.exe " + Chr(34) + FName + Chr(34)
            KL = CreateConsoleProcess(Cmdline)
            ScanAgeProGeneralData
        Case 7
            N = InStrRev(FName, ".")
            FN = Left(FName, N) & "tmp"
            If Dir(FN) <> "" Then
                Kill FN
            End If
            'check to make sure species list is available
            N = InStrRev(FName, "\")
            FN = Left(FName, N) & "species.txt"
            If Dir(FN) = "" Then
                FileCopy App.Path + "\species.txt", FN
            End If
            'Cmdline = "cmd.exe /c " + App.Path + "\scansamplenwt.exe " + FName + " 2>VR.LOG"
            Cmdline = App.Path + "\scansamplenwt.exe " + Chr(34) + FName + Chr(34)
            KL = CreateConsoleProcess(Cmdline)
            ScanSampleGeneralData
    End Select
    
    Exit Sub
ScanErr:
    MsgBox "Error Scanning in Toolbox Model Data", vbExclamation, "Visual Report Designer"
End Sub
Private Sub ScanVPAGeneralData()
    Dim N As Long
    Dim Buffer As String
    Dim Token As String
    
    On Error GoTo ScanVPAErr
    
    FNOUT = FName
    N = InStrRev(FNOUT, ".")
    Mid(FNOUT, N, 4) = ".tmp"
    
    If Dir(FNOUT) = "" Then
        MsgBox "VPA Data Scan Failed" + vbCrLf + "VPA Output Files May Be Incomplete", vbExclamation, "Visual Report Designer"
        Exit Sub
    End If
    
    Open FNOUT For Input As #5
    
    Line Input #5, Buffer
    
    frmVPAScan.lblFile.Caption = FNOUT
    frmVPAScan.txtCase.Text = Buffer
    
    Line Input #5, Buffer
    
    Token = GetFirstToken(Buffer)
    KYears = Val(Token)
    Token = GetNextToken(Buffer)
    KAges = Val(Token)
    Token = GetNextToken(Buffer)
    Token = GetNextToken(Buffer)
    KINDX = Val(Token)
    Token = GetNextToken(Buffer)
    KFYear = Val(Token)
    Token = GetNextToken(Buffer)
    KFAge = Val(Token)
    Token = GetNextToken(Buffer)
    KAge1 = Val(Token)
    Token = GetNextToken(Buffer)
    KAge2 = Val(Token)

    frmVPAScan.lblAges.Caption = CStr(KAges)
    frmVPAScan.lblIndex.Caption = CStr(KINDX)
    frmVPAScan.lblYear(0).Caption = CStr(KFYear)
    KXYear = KFYear + KYears
    frmVPAScan.lblYear(1).Caption = CStr(KXYear)
    frmVPAScan.Show

    AgeFlag = True
'    If KFYear < StartYear Or KXYear > EndYear Then
'        MsgBox "Invalid Year Range Specification for Data Grid", vbInformation, "Visual Report Designer"
'        Close #5
'        AgeFlag = False
'    Else
'        AgeFlag = True
'    End If
'
'    If KFYear < MinYear Then
'        MinYear = KFYear
'    End If
'    If KXYear > MaxYear Then
'        MaxYear = KXYear
'    End If
    
    Exit Sub
    
ScanVPAErr:
    MsgBox "Error Scanning VPA Results", vbInformation, "Visual Report Designer"
    Close #5
End Sub
Public Sub ScanVPAResults()
    Dim Buffer As String
    Dim Token As String
    Dim G As MSFlexGrid
    Dim I As Integer
    Dim J As Integer
    Dim N As Integer
    Dim K As Integer
    Dim KR As Integer
    Dim Survey() As VPASurvey
    Dim SurveyData() As Double
    Dim SurveyResid() As Double
    Dim DT As String
    
    On Error GoTo ScanVPAError
    
    If AgeFlag = False Then
        Unload frmVPAScan
        Exit Sub
    End If

    KR = KRow
    
    CaseID = Trim(frmVPAScan.txtCase.Text)
    
    'get years and reset forms if needed
    I = Val(frmVPAScan.lblYear(0).Caption)
    J = Val(frmVPAScan.lblYear(1).Caption)
    If StartYear = 0 And EndYear = 0 Then
        frmGeneral.lblStartYr.Caption = CStr(I)
        frmGeneral.lblEndYr.Caption = CStr(J)
        SetParms
    ElseIf I < StartYear Or J > EndYear Then
        If I < StartYear Then frmGeneral.lblStartYr.Caption = CStr(I)
        If J > EndYear Then frmGeneral.lblEndYr.Caption = CStr(J)
        SetParms
    End If
    
    Unload frmVPAScan
    
    Print #3, "VPA"
    Print #3, CaseID
    Print #3, FName
    DT = GetFileWriteTime(FName)
    Print #3, DT
    Print #3, ""
    
    Set G = frmData.MSFlexGrid1
    N = KFYear - StartYear + 4
    
    Line Input #5, Buffer
    G.Rows = G.Rows + KAges
    For I = 1 To KAges
        K = KRow + I
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "VPA"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = "Stock Weight"
        G.TextMatrix(K, 4) = "Age " + CStr(I)
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For J = 1 To KYears + 1
            G.TextMatrix(K, J + N) = Token
            Token = GetNextToken(Buffer)
        Next J
    Next I
    KRow = KRow + KAges
    
    Line Input #5, Buffer
    G.Rows = G.Rows + KAges
    For I = 1 To KAges
        K = KRow + I
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "VPA"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = "Catch Weight"
        G.TextMatrix(K, 4) = "Age " + CStr(I)
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For J = 1 To KYears + 1
            G.TextMatrix(K, J + N) = Token
            Token = GetNextToken(Buffer)
        Next J
    Next I
    KRow = KRow + KAges
    
    Line Input #5, Buffer
    G.Rows = G.Rows + KAges
    For I = 1 To KAges
        K = KRow + I
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "VPA"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = "Spawning Stock Weight"
        G.TextMatrix(K, 4) = "Age " + CStr(I)
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For J = 1 To KYears + 1
            G.TextMatrix(K, J + N) = Token
            Token = GetNextToken(Buffer)
        Next J
    Next I
    KRow = KRow + KAges
    
    Line Input #5, Buffer
    G.Rows = G.Rows + KAges
    For I = 1 To KAges
        K = KRow + I
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "VPA"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = "Maturity"
        G.TextMatrix(K, 4) = "Age " + CStr(I)
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For J = 1 To KYears + 1
            G.TextMatrix(K, J + N) = Token
            Token = GetNextToken(Buffer)
        Next J
    Next I
    KRow = KRow + KAges
    
    Line Input #5, Buffer
    G.Rows = G.Rows + KAges
    For I = 1 To KAges
        K = KRow + I
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "VPA"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = "Natural Mortality"
        G.TextMatrix(K, 4) = "Age " + CStr(I)
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For J = 1 To KYears + 1
            G.TextMatrix(K, J + N) = Token
            Token = GetNextToken(Buffer)
        Next J
    Next I
    KRow = KRow + KAges
    
    Line Input #5, Buffer
    G.Rows = G.Rows + KAges
    For I = 1 To KAges
        K = KRow + I
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "VPA"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = "Stock Numbers"
        G.TextMatrix(K, 4) = "Age " + CStr(I)
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For J = 1 To KYears + 1
            G.TextMatrix(K, J + N) = Token
            Token = GetNextToken(Buffer)
        Next J
    Next I
    KRow = KRow + KAges
    
    Line Input #5, Buffer
    G.Rows = G.Rows + 1
    K = KRow + 1
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "VPA"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Total Stock Numbers"
    G.TextMatrix(K, 4) = ""
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears + 1
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    KRow = KRow + 1
    
    Line Input #5, Buffer
    G.Rows = G.Rows + KAges
    For I = 1 To KAges
        K = KRow + I
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "VPA"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = "Catch Numbers"
        G.TextMatrix(K, 4) = "Age " + CStr(I)
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For J = 1 To KYears + 1
            G.TextMatrix(K, J + N) = Token
            Token = GetNextToken(Buffer)
        Next J
    Next I
    KRow = KRow + KAges
    
    Line Input #5, Buffer
    G.Rows = G.Rows + KAges
    For I = 1 To KAges
        K = KRow + I
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "VPA"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = "JAN-1 Biomass"
        G.TextMatrix(K, 4) = "Age " + CStr(I)
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For J = 1 To KYears + 1
            G.TextMatrix(K, J + N) = Token
            Token = GetNextToken(Buffer)
        Next J
    Next I
    KRow = KRow + KAges

    Line Input #5, Buffer
    G.Rows = G.Rows + 1
    K = KRow + 1
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "VPA"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Total Biomass"
    G.TextMatrix(K, 4) = ""
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears + 1
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    KRow = KRow + 1
    
    Line Input #5, Buffer
    G.Rows = G.Rows + KAges
    For I = 1 To KAges
        K = KRow + I
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "VPA"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = "Spawning Stock Biomass"
        G.TextMatrix(K, 4) = "Age " + CStr(I)
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For J = 1 To KYears
            G.TextMatrix(K, J + N) = Token
            Token = GetNextToken(Buffer)
        Next J
    Next I
    KRow = KRow + KAges

    Line Input #5, Buffer
    G.Rows = G.Rows + 1
    K = KRow + 1
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "VPA"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Total Spawning Stock Biomass"
    G.TextMatrix(K, 4) = ""
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    KRow = KRow + 1
  
    Line Input #5, Buffer
    G.Rows = G.Rows + 1
    K = KRow + 1
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "VPA"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Total Catch Biomass"
    G.TextMatrix(K, 4) = ""
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    KRow = KRow + 1
    
    
    Line Input #5, Buffer
    G.Rows = G.Rows + 1
    K = KRow + 1
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "VPA"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Surplus Production"
    G.TextMatrix(K, 4) = ""
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    KRow = KRow + 1
    
    
    Line Input #5, Buffer
    G.Rows = G.Rows + KAges
    For I = 1 To KAges
        K = KRow + I
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "VPA"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = "Fishing Mortality"
        G.TextMatrix(K, 4) = "Age " + CStr(I)
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For J = 1 To KYears
            G.TextMatrix(K, J + N) = Token
            Token = GetNextToken(Buffer)
        Next J
    Next I
    KRow = KRow + KAges
    
    Line Input #5, Buffer
    G.Rows = G.Rows + 1
    K = KRow + 1
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "VPA"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Average F " + CStr(KAge1) + " - " + CStr(KAge2)
    G.TextMatrix(K, 4) = ""
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    KRow = KRow + 1

    
    Line Input #5, Buffer
    G.Rows = G.Rows + 1
    K = KRow + 1
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "VPA"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Average F " + CStr(KAge1) + " - " + CStr(KAge2)
    G.TextMatrix(K, 4) = "N Wtd"
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    KRow = KRow + 1
    
    Line Input #5, Buffer
    G.Rows = G.Rows + 1
    K = KRow + 1
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "VPA"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Average F " + CStr(KAge1) + " - " + CStr(KAge2)
    G.TextMatrix(K, 4) = "C Wtd"
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    KRow = KRow + 1
    
    
    Line Input #5, Buffer
    G.Rows = G.Rows + 1
    K = KRow + 1
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "VPA"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Average F " + CStr(KAge1) + " - " + CStr(KAge2)
    G.TextMatrix(K, 4) = "B Wtd"
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    KRow = KRow + 1
    
    ReDim Survey(1 To KINDX)
    ReDim SurveyData(1 To KINDX, 1 To KYears + 1)
    ReDim SurveyResid(1 To KINDX, 1 To KYears + 1)
    
    Line Input #5, Buffer
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KINDX
        Survey(J).Tag = Token
        Token = GetNextToken(Buffer)
    Next J
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KINDX
        Survey(J).StartAge = Val(Token)
        Token = GetNextToken(Buffer)
    Next J
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KINDX
        Survey(J).EndAge = Val(Token)
        Token = GetNextToken(Buffer)
    Next J
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KINDX
        Survey(J).Time = Val(Token)
        Token = GetNextToken(Buffer)
    Next J
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KINDX
        Survey(J).Type = Val(Token)
        Token = GetNextToken(Buffer)
    Next J
    
    For J = 1 To KYears + 1
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For I = 1 To KINDX
            SurveyData(I, J) = Val(Token)
            Token = GetNextToken(Buffer)
        Next I
    Next J
    G.Rows = G.Rows + KINDX
    
    Line Input #5, Buffer
    For J = 1 To KYears + 1
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For I = 1 To KINDX
            SurveyResid(I, J) = Val(Token)
            Token = GetNextToken(Buffer)
        Next I
    Next J
    G.Rows = G.Rows + KINDX
    
    For I = 1 To KINDX
        K = KRow + I
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "VPA"
        G.TextMatrix(K, 2) = CaseID
        Token = Trim(Survey(I).Tag) + Space(1) + CStr(Survey(I).StartAge) + "-" + CStr(Survey(I).EndAge)
        If Survey(I).Type = 0 Then
            Token = Token + " N"
        Else
            Token = Token + " W"
        End If
        If Survey(I).Time = 0 Then
            Token = Token + "/Jan-1 "
        Else
            Token = Token + "/Mean "
        End If
        G.TextMatrix(K, 3) = Token
        G.TextMatrix(K, 4) = "Obs. Index"
        For J = 1 To KYears + 1
            G.TextMatrix(K, J + N) = Format(SurveyData(I, J), "0.00000E+00")
        Next J
    Next I
    KRow = KRow + KINDX
    
    For I = 1 To KINDX
        K = KRow + I
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "VPA"
        G.TextMatrix(K, 2) = CaseID
        Token = Trim(Survey(I).Tag) + Space(1) + CStr(Survey(I).StartAge) + "-" + CStr(Survey(I).EndAge)
        If Survey(I).Type = 0 Then
            Token = Token + " Numbers"
        Else
            Token = Token + " Weight"
        End If
        If Survey(I).Time = 0 Then
            Token = Token + "/Jan-1 "
        Else
            Token = Token + "/Mean "
        End If
        G.TextMatrix(K, 3) = Token
        G.TextMatrix(K, 4) = "Residual"
        For J = 1 To KYears + 1
            G.TextMatrix(K, J + N) = Format(SurveyResid(I, J), "0.00000E+00")
        Next J
    Next I
    KRow = KRow + KINDX
    G.Rows = KRow + 1
        
    K = G.Rows
    
    frmCompose.Show
    frmData.ZOrder
    
    MsgBox "VPA Data Scan Completed" + vbCrLf + CStr(KRow - KR) + " Rows Added to Data Grid", vbInformation, "Visual Report Designer"
    Close #5
    
    DataFlag = True
    
    frmMain.mnuDataMissingReplace.Enabled = True
    
    'automatically save data into file
    WriteGridData

    Exit Sub
    
ScanVPAError:
    MsgBox "Error Scanning VPA Results" + vbCrLf + Err.Description, vbInformation, "Visual Report Designer"
    Close #5
End Sub
Private Function CheckVPAFiles() As Boolean
    Dim N As Long
    
    On Error GoTo CheckErr
    
    FNOUT = FName
    N = InStrRev(FNOUT, ".")
    Mid(FNOUT, N, 4) = ".out"
    If CompareFileWriteTimes(FName, FNOUT) Then
        MsgBox "VPA Output Files Are Not Current", vbInformation, "Visual Report Designer"
        CheckVPAFiles = False
        Exit Function
    End If
    Mid(FNOUT, N, 4) = ".tmp"
    If Dir(FNOUT) <> "" Then
        Kill FNOUT
    End If
    CheckVPAFiles = True
    Exit Function
CheckErr:
    MsgBox "Error Checking VPA Files" + vbCrLf + Err.Description, vbExclamation, "Visual Report Designer"
End Function
Private Function CheckCSAFiles() As Boolean
    Dim N As Long
    
    On Error GoTo CheckErr1
    
    FNOUT = FName
    N = InStrRev(FNOUT, ".")
    Mid(FNOUT, N, 4) = ".out"
    If CompareFileWriteTimes(FName, FNOUT) Then
        MsgBox "CSA Output Files Are Not Current", vbInformation, "Visual Report Designer"
        CheckCSAFiles = False
        Exit Function
    End If
    Mid(FNOUT, N, 4) = ".tmp"
    If Dir(FNOUT) <> "" Then
        Kill FNOUT
    End If
    CheckCSAFiles = True
    Exit Function
    
CheckErr1:
    MsgBox "Error Checking CSA Files" + vbCrLf + Err.Description, vbExclamation, "Visual Report Designer"
End Function
Public Sub SetParms()
    Dim KStart As Integer
    Dim KEnd As Integer
    Dim I As Integer
    Dim tmpLayoutFlag As Boolean

    'Unload frmCompose
    Unload frmVPAScan
    Unload frmCSAScan
    Unload frmAspicScan
    Unload frmASAPScan
    Unload frmAgeProScan
    Unload frmAimScan
    Unload frmAux
    
    'if there is report data then save report before unloading
    tmpLayoutFlag = False
    If LayoutFlag Then
        SaveLayoutFile
        tmpLayoutFlag = True
    End If
    Unload frmCompose
    
    KStart = Val(frmGeneral.lblStartYr.Caption)
    KEnd = Val(frmGeneral.lblEndYr.Caption)
    
    StartYear = KStart
    EndYear = KEnd
    NYears = EndYear - StartYear + 1
    
    If DataFlag = False Then
        MaxYear = StartYear
        MinYear = EndYear
    End If
        
    If DataFlag Then
        ResizeGrid
    Else
        Unload frmData
        InitDataCollectionGrid
    End If
    
    frmMain.mnuFileSaveDatabase.Enabled = True
    frmMain.mnuDataAdd.Enabled = True
    
    If tmpLayoutFlag Then ReadLayoutFile
    
    'ReDim BList(1 To 450)
End Sub
Public Function CompareFileWriteTimes(F1 As String, F2 As String) As Boolean
    Dim H1 As Long
    Dim H2 As Long
    Dim XX As Long
    Dim TCreate As FILETIME
    Dim TAccess As FILETIME
    Dim TW1 As FILETIME
    Dim TW2 As FILETIME
    
    ' CompareFileWriteTimes = False = file write times check out OK.
    ' CompareFileWriteTimes = True = file write times are not OK.
    ' F1 should be oldest (earliest) file; F2 should be most recently written file.
    
    ' Create File Handles
    H1 = CreateFile(F1, GENERIC_READ Or GENERIC_WRITE, 0, 0, OPEN_EXISTING, 0, 0)
    H2 = CreateFile(F2, GENERIC_READ Or GENERIC_WRITE, 0, 0, OPEN_EXISTING, 0, 0)
    ' Get File Times (Write Last)
    XX = GetFileTime(H1, TCreate, TAccess, TW1)
    XX = GetFileTime(H2, TCreate, TAccess, TW2)
    XX = CompareFileTime(TW1, TW2)
    If XX = 1 Or XX = 0 Then
        CompareFileWriteTimes = True
    Else
        CompareFileWriteTimes = False
    End If
    XX = CloseHandle(H1)
    XX = CloseHandle(H2)
End Function
Public Function GetFileWriteTime(F1 As String) As String
    Dim H1 As Long
    Dim XX As Long
    Dim TCreate As FILETIME
    Dim TAccess As FILETIME
    Dim TW1 As FILETIME
    Dim ST1 As SYSTEMTIME
    Dim X As Boolean
    Dim DT As String
    
    ' Create File Handle
    H1 = CreateFile(F1, GENERIC_READ Or GENERIC_WRITE, 0, 0, OPEN_EXISTING, 0, 0)
    ' Get File Times (Write Last)
    XX = GetFileTime(H1, TCreate, TAccess, TW1)
    ' Convert Filetime to SystemTime
    X = FileTimeToSystemTime(TW1, ST1)
    DT = Format(ST1.wMonth, "00") + "/" + Format(ST1.wDay, "00") + "/"
    DT = DT + Format(ST1.wYear, "0000") + Space(5)
    DT = DT + Format(ST1.wHour, "00") + ":" + Format(ST1.wMinute, "00") + " (GMT)"
    GetFileWriteTime = DT
    XX = CloseHandle(H1)
End Function
Public Sub InitDataCollectionGrid()
    Dim G As MSFlexGrid
    Dim I As Integer
    
    Set G = frmData.MSFlexGrid1
    
    G.Cols = NYears + 5
    G.FixedCols = 0
    G.Rows = 1
    G.TextMatrix(0, 0) = "Line"
    G.TextMatrix(0, 1) = "Source"
    G.TextMatrix(0, 2) = "Case"
    G.TextMatrix(0, 3) = "Data Type"
    G.TextMatrix(0, 4) = "Item"
    G.ColWidth(0) = 600
    G.ColWidth(1) = 1000
    G.ColWidth(2) = 2500
    G.ColWidth(3) = 2500
    G.ColWidth(4) = 1100
    
    For I = 1 To NYears
        G.ColWidth(I + 4) = 1600
        G.TextMatrix(0, I + 4) = CStr(StartYear + I - 1)
    Next I
    
    KRow = 0
    
    
    
    frmData.Show
End Sub
Public Sub AddtoLayout()
    Dim I As Integer
    Dim J As Integer
    Dim iLine As Integer
    Dim iRpt As Integer
    Dim N As Integer
    Dim M As Integer
    Dim S As String
    Dim T As String
    Dim rtnval As Integer
    Dim G As MSFlexGrid
    Dim NLines As Integer
    Dim BatchMode As Boolean
    
    'set flag for batch mode or individual add mode
    If frmSpecLayout.frmDesc.Visible = True Then
        BatchMode = False 'individual add mode
    Else
        BatchMode = True 'batch mode
    End If
    
    'exit if user has canceled out of a multi-select add
    If MultiFlag And CancelFlag And Not BatchMode Then
        MultiFlag = False
        MultiLastRow = 0
        CancelFlag = False
        Exit Sub
    End If
            
    iRpt = frmSpecLayout.cboReport.ListIndex 'report number
    iLine = Val(frmSpecLayout.txtLine.Text) 'line to begin adding data
    NLines = Val(frmSpecLayout.lblNLines.Caption) 'number of lines to add
    
    'check if user wants to replace an existing item
    If Not BatchMode Then
        S = frmSpecLayout.txtLinesUsed.Text
        T = ", " & CStr(iLine) & ","
        N = InStr(S, T)
        If N > 0 Then
            rtnval = MsgBox("This line already contains data." & vbCrLf & _
            "Do you want to overwrite?", vbQuestion + vbOKCancel, "Visual Report Designer")
            If rtnval = vbCancel Then
                Exit Sub
            End If
        End If
    End If
    
    'intitialize variables and grid to prepare for more lines if necessary
    N = iLine + NLines - 1
    If N > MaxLines Then
        MaxLines = N
        ReDim Preserve BList(0 To 14, 1 To N)
    End If
    If N > ReportInfo(iRpt).NLines Then ReportInfo(iRpt).NLines = N
    Set G = frmCompose.grdReport(iRpt)
    If G.Rows < N + 1 Then
        M = G.Rows
        G.Rows = N + 1
        For I = M To N
            G.TextMatrix(I, 0) = CStr(I)
        Next I
    End If
    
    For I = 1 To NLines
        'key
        M = Val(frmSpecLayout.lblKey.Caption) + I - 1
        BList(iRpt, iLine + I - 1).Key = M
        'line tag
        If BatchMode Then 'need to construct tag from check boxes on form
            Set G = frmData.MSFlexGrid1
            S = ""
            For J = 0 To 3
                If frmSpecLayout.chkTag(J).Value = 1 Then
                    S = S & G.TextMatrix(M, J + 1) & " "
                End If
            Next J
            BList(iRpt, iLine + I - 1).Tag = Trim(S)
        Else 'get line tag from form
            If frmSpecLayout.opTag(0).Value = True Then
                BList(iRpt, iLine + I - 1).Tag = Trim(frmSpecLayout.lblTag.Caption)
            Else
                BList(iRpt, iLine + I - 1).Tag = Trim(frmSpecLayout.txtTag.Text)
            End If
        End If
        'years
        If BatchMode And frmSpecLayout.opFYear(0).Value = True Then
            BList(iRpt, iLine + I - 1).StartYear = GetDataMinYear(M)
        Else
            BList(iRpt, iLine + I - 1).StartYear = frmSpecLayout.cboFYear.ListIndex + StartYear
        End If
        If BatchMode And frmSpecLayout.opXYear(0).Value = True Then
            BList(iRpt, iLine + I - 1).EndYear = GetDataMaxYear(M)
        Else
            BList(iRpt, iLine + I - 1).EndYear = frmSpecLayout.cboXYear.ListIndex + StartYear
        End If
        'type and display options
        BList(iRpt, iLine + I - 1).Type = frmSpecLayout.cboBins.ListIndex + 2
        BList(iRpt, iLine + I - 1).LowerCut = Val(frmSpecLayout.txtMark(0).Text)
        BList(iRpt, iLine + I - 1).UpperCut = Val(frmSpecLayout.txtMark(1).Text)
        If frmSpecLayout.chkZero.Value = 0 Then
            BList(iRpt, iLine + I - 1).ZeroFlag = False
        Else
            BList(iRpt, iLine + I - 1).ZeroFlag = True
        End If
        BList(iRpt, iLine + I - 1).Palette = Trim(frmSpecLayout.cboPalette.Text)
        'assign palette to report legends if necessary
        If BList(iRpt, iLine + I - 1).Type = 4 Then
            Select Case BList(iRpt, iLine + I - 1).Palette
                Case "Black_to_Red", "Red_to_Black": N = 1
                Case "Red_to_Blue", "Blue_to_Red": N = 2
                Case "White_to_Black", "Black_to_White": N = 3
            End Select
            ReportLegend(iRpt, N, 0) = BList(iRpt, iLine + I - 1).Palette
        End If
        
        'put data on grid
        PutReportData iRpt, (iLine + I - 1)
    Next I
    
    Unload frmSpecLayout
    frmCompose.SSTab1.Tab = iRpt
    
    WriteLayoutFile
    HTMLFile = Left(FNLOG, Len(FNLOG) - 4) & "_rpt" & CStr(iRpt + 1) & ".html"
    If SaveLayout Then WriteHTML HTMLFile, iRpt
    frmCompose.ZOrder
    
    If MultiFlag And Not BatchMode Then
        'if at end of multi-select mode
        If BList(iRpt, iLine).Key = MultiLastRow Then
            MultiFlag = False
            MultiLastRow = 0
        Else 'continue adding
            SelectItem (BList(iRpt, iLine).Key + 1), 1
        End If
    End If
    
End Sub
Public Sub GetDataStats(iKey As Integer)
    Dim G As MSFlexGrid
    Dim I As Integer
    Dim Istart
    Dim S As String
    Dim min As Double
    Dim max As Double
    Dim fyr As Integer
    Dim xyr As Integer
    
    Set G = frmData.MSFlexGrid1
    
    'initialize min and max by finding first non-blank value
    Istart = 0 'initialize in case there is no data
    For I = 5 To G.Cols - 1
        S = G.TextMatrix(iKey, I)
        If S <> "" Then
            min = Val(S)
            max = Val(S)
            Istart = I
            Exit For
        End If
    Next I
    
    If Istart > 0 Then
        'find min and max values
        For I = Istart + 1 To G.Cols - 1
            S = G.TextMatrix(iKey, I)
            If Trim(S) <> "" Then
                If Val(S) < min Then min = Val(S)
                If Val(S) > max Then max = Val(S)
            End If
        Next I
        'get min and max years
        fyr = GetDataMinYear(iKey)
        xyr = GetDataMaxYear(iKey)
        'report the stats
        MsgBox "First Year in Data: " & CStr(fyr) & Space(15) & vbCrLf & _
           "Last Year in Data: " & CStr(xyr) & vbCrLf & vbCrLf & _
           "Minimum Value: " & CStr(min) & vbCrLf & _
           "Maximum Value: " & CStr(max), vbInformation, "Visual Report Designer"
    Else
        MsgBox "This item contains no data", vbInformation, "Visual Report Designer"
    End If
        
End Sub

Public Function GetDataMinYear(Line As Integer) As Integer
    Dim G As MSFlexGrid
    Dim I As Integer
    Dim S As String
    
    Set G = frmData.MSFlexGrid1
    
    'initialize min year in case all data is zeroes
    GetDataMinYear = Val(G.TextMatrix(0, 5))
    
    For I = 5 To G.Cols - 1
        S = G.TextMatrix(Line, I)
        If Trim(S) <> "" Then
            GetDataMinYear = Val(G.TextMatrix(0, I))
            Exit Function
        End If
    Next I
    
End Function
Public Function GetDataMaxYear(Line As Integer) As Integer
    Dim G As MSFlexGrid
    Dim I As Integer
    Dim S As String
    
    Set G = frmData.MSFlexGrid1
    
    'initialize max year in case all data is zeroes
    GetDataMaxYear = Val(G.TextMatrix(0, G.Cols - 1))
    
    For I = G.Cols - 1 To 5 Step -1
        S = G.TextMatrix(Line, I)
        If Trim(S) <> "" Then
            GetDataMaxYear = Val(G.TextMatrix(0, I))
            Exit Function
        End If
    Next I
    
End Function
Public Sub InitSpecEdit(Index As Integer, iRow As Integer)
    Dim I As Integer
    Dim N As Integer
    Dim K As Integer
    Dim G As MSFlexGrid
    Dim T As String
    Dim iKey As Integer
    
    Set G = frmCompose.grdReport(Index)
    
    'exit if the row is a blank line or if the user clicked on the header row
    If iRow = 0 Or G.TextMatrix(iRow, 1) = "" Then
        If MultiFlag Then
            'if at end of multi-select mode
            If iRow = MultiLastRow Then
                MultiFlag = False
                MultiLastRow = 0
                Exit Sub
            Else 'continue with editing
                For I = (iRow + 1) To MultiLastRow
                    iRow = iRow + 1
                    If G.TextMatrix(iRow, 1) <> "" Then Exit For
                Next I
            End If
        Else
            Exit Sub
        End If
    End If
    
    'hide options for batch edit mode
    frmSpecLayout.opFYear(2).Visible = False
    frmSpecLayout.opXYear(2).Visible = False
    frmSpecLayout.chkEditDisplay.Visible = False
    frmSpecLayout.chkEditZero.Visible = False

    'get key
    iKey = BList(Index, iRow).Key
    'put the data grid key, report, and line on the form for use later
    frmSpecLayout.lblKey.Caption = CStr(iKey)
    frmSpecLayout.lblJRpt.Caption = CStr(Index)
    frmSpecLayout.lblJLine.Caption = CStr(iRow)
    'frmSpecLayout.lblNLines.Caption = CStr(1)
    
    frmSpecLayout.cmdUpdate.Caption = "UPDATE"
    
    'get item description
    T = Trim(GetDataGridDesc(iKey))
    frmSpecLayout.lblDescrip.Caption = T

    'get tag
    If T = Trim(BList(Index, iRow).Tag) Then
        frmSpecLayout.opTag(0).Value = True
        For I = 0 To 3
            frmSpecLayout.chkTag(I).Value = 1
        Next I
    Else
        frmSpecLayout.opTag(1).Value = True
        frmSpecLayout.txtTag.Text = BList(Index, iRow).Tag
    End If

    'index and line
    frmSpecLayout.cboReport.ListIndex = Index
    frmSpecLayout.txtLine.Text = CStr(iRow)

    'start and end years
    N = GetDataMinYear(iKey)
    frmSpecLayout.lblFYear.Caption = CStr(N)
    If N = BList(Index, iRow).StartYear Then
        frmSpecLayout.opFYear(0).Value = True
    Else
        frmSpecLayout.opFYear(1).Value = True
        frmSpecLayout.cboFYear.ListIndex = BList(Index, iRow).StartYear - StartYear
    End If
    N = GetDataMaxYear(iKey)
    frmSpecLayout.lblXYear.Caption = CStr(N)
    If N = BList(Index, iRow).EndYear Then
        frmSpecLayout.opXYear(0).Value = True
    Else
        frmSpecLayout.opXYear(1).Value = True
        frmSpecLayout.cboXYear.ListIndex = BList(Index, iRow).EndYear - StartYear
    End If
    
    frmSpecLayout.txtMark(0).Text = CStr(BList(Index, iRow).LowerCut)
    frmSpecLayout.txtMark(1).Text = CStr(BList(Index, iRow).UpperCut)
    
    N = BList(Index, iRow).Type
    frmSpecLayout.cboBins.ListIndex = N - 2
    T = BList(Index, iRow).Palette
    
    With frmSpecLayout
    If N = 2 Then 'single cut point
        Select Case T
            Case "Plain_Plus/Minus": Set .cboPalette.SelectedItem = .cboPalette.ComboItems(1)
            Case "Green_Plus/Red_Minus": Set .cboPalette.SelectedItem = .cboPalette.ComboItems(2)
            Case "Red_Plus/Green_Minus": Set .cboPalette.SelectedItem = .cboPalette.ComboItems(3)
            Case "Green/Red": Set .cboPalette.SelectedItem = .cboPalette.ComboItems(4)
            Case "Red/Green": Set .cboPalette.SelectedItem = .cboPalette.ComboItems(5)
        End Select
    ElseIf N = 3 Then 'dual cut point
        Select Case T
            Case "Plain_Plus_to_Minus": Set .cboPalette.SelectedItem = .cboPalette.ComboItems(1)
            Case "Green_Plus_to_Red_Minus": Set .cboPalette.SelectedItem = .cboPalette.ComboItems(2)
            Case "Red_Plus_to_Green_Minus": Set .cboPalette.SelectedItem = .cboPalette.ComboItems(3)
            Case "Green_to_Red": Set .cboPalette.SelectedItem = .cboPalette.ComboItems(4)
            Case "Red_to_Green": Set .cboPalette.SelectedItem = .cboPalette.ComboItems(5)
        End Select
    ElseIf N = 4 Then 'quintiles
        Select Case T
            Case "Red_to_Black": Set .cboPalette.SelectedItem = .cboPalette.ComboItems(1)
            Case "Black_to_Red": Set .cboPalette.SelectedItem = .cboPalette.ComboItems(2)
            Case "Red_to_Blue": Set .cboPalette.SelectedItem = .cboPalette.ComboItems(3)
            Case "Blue_to_Red": Set .cboPalette.SelectedItem = .cboPalette.ComboItems(4)
            Case "White_to_Black": Set .cboPalette.SelectedItem = .cboPalette.ComboItems(5)
            Case "Black_to_White": Set .cboPalette.SelectedItem = .cboPalette.ComboItems(6)
        End Select
    End If
    End With
    
    If BList(Index, iRow).ZeroFlag Then
        frmSpecLayout.chkZero.Value = 1
    Else
        frmSpecLayout.chkZero.Value = 0
    End If
    
    frmSpecLayout.Show

End Sub
Public Sub InitSpecBatch(Index As Integer, iRow As Integer, NRows As Integer)
    Dim I As Integer
    Dim N As Integer
    Dim K As Integer
    Dim T As String
    Dim iKey As Integer
    
    'hide description box
    frmSpecLayout.frmDesc.Visible = False
    'edit number of items description
    frmSpecLayout.lblRowsAdd.Caption = "Number of Rows of Data to Edit:"
    'number of items to add
    frmSpecLayout.lblNLines.Caption = CStr(NRows)
    
    'get key
    iKey = BList(Index, iRow).Key
    'put the data grid key, report, and line on the form for use later
    frmSpecLayout.lblKey.Caption = CStr(iKey)
    frmSpecLayout.lblJRpt.Caption = CStr(Index)
    frmSpecLayout.lblJLine.Caption = CStr(iRow)
    
    frmSpecLayout.cmdUpdate.Caption = "UPDATE"
    
    'item description
    frmSpecLayout.lblLine.Caption = "Example:"
    For I = 0 To 3
        frmSpecLayout.chkTag(I).Value = 1
    Next I
    frmSpecLayout.opTag(1).Caption = "Do Not Edit"
    frmSpecLayout.opTag(1).Value = True
    frmSpecLayout.txtTag.Visible = False

    'index and line
    frmSpecLayout.cboReport.ListIndex = Index
    frmSpecLayout.txtLine.Text = CStr(iRow)

    'start and end years
    frmSpecLayout.opFYear(0).Caption = "First non-blank year in the data"
    frmSpecLayout.lblFYear.Visible = False
    frmSpecLayout.opFYear(2).Visible = True
    frmSpecLayout.opFYear(2).Value = True
    frmSpecLayout.cboFYear.Visible = False
    frmSpecLayout.opXYear(0).Caption = "Last non-blank year in the data"
    frmSpecLayout.lblXYear.Visible = False
    frmSpecLayout.opXYear(2).Visible = True
    frmSpecLayout.opXYear(2).Value = True
    frmSpecLayout.cboXYear.Visible = False
    
    frmSpecLayout.txtMark(0).Text = CStr(BList(Index, iRow).LowerCut)
    frmSpecLayout.txtMark(1).Text = CStr(BList(Index, iRow).UpperCut)
    
    N = BList(Index, iRow).Type
    If N = 0 Then N = 4
    frmSpecLayout.cboBins.ListIndex = N - 2
    T = BList(Index, iRow).Palette
    
    With frmSpecLayout
    If N = 2 Then 'single cut point
        Select Case T
            Case "Plain_Plus/Minus": Set .cboPalette.SelectedItem = .cboPalette.ComboItems(1)
            Case "Green_Plus/Red_Minus": Set .cboPalette.SelectedItem = .cboPalette.ComboItems(2)
            Case "Red_Plus/Green_Minus": Set .cboPalette.SelectedItem = .cboPalette.ComboItems(3)
            Case "Green/Red": Set .cboPalette.SelectedItem = .cboPalette.ComboItems(4)
            Case "Red/Green": Set .cboPalette.SelectedItem = .cboPalette.ComboItems(5)
        End Select
    ElseIf N = 3 Then 'dual cut point
        Select Case T
            Case "Plain_Plus_to_Minus": Set .cboPalette.SelectedItem = .cboPalette.ComboItems(1)
            Case "Green_Plus_to_Red_Minus": Set .cboPalette.SelectedItem = .cboPalette.ComboItems(2)
            Case "Red_Plus_to_Green_Minus": Set .cboPalette.SelectedItem = .cboPalette.ComboItems(3)
            Case "Green_to_Red": Set .cboPalette.SelectedItem = .cboPalette.ComboItems(4)
            Case "Red_to_Green": Set .cboPalette.SelectedItem = .cboPalette.ComboItems(5)
        End Select
    ElseIf N = 4 Then 'quintiles
        Select Case T
            Case "Red_to_Black": Set .cboPalette.SelectedItem = .cboPalette.ComboItems(1)
            Case "Black_to_Red": Set .cboPalette.SelectedItem = .cboPalette.ComboItems(2)
            Case "Red_to_Blue": Set .cboPalette.SelectedItem = .cboPalette.ComboItems(3)
            Case "Blue_to_Red": Set .cboPalette.SelectedItem = .cboPalette.ComboItems(4)
            Case "White_to_Black": Set .cboPalette.SelectedItem = .cboPalette.ComboItems(5)
            Case "Black_to_White": Set .cboPalette.SelectedItem = .cboPalette.ComboItems(6)
        End Select
    End If
    End With
    
    frmSpecLayout.chkEditDisplay.Visible = True
    frmSpecLayout.chkEditDisplay.Value = 1
    
    If BList(Index, iRow).ZeroFlag Then
        frmSpecLayout.chkZero.Value = 1
    Else
        frmSpecLayout.chkZero.Value = 0
    End If
    
    frmSpecLayout.chkEditZero.Visible = True
    frmSpecLayout.chkEditZero.Value = 1
    
    frmSpecLayout.Show

End Sub
Public Sub UpdateLayout()
    Dim I As Integer
    Dim iLine As Integer
    Dim iRpt As Integer
    Dim G As MSFlexGrid
    Dim N As Integer
    Dim S As String
    Dim T As String
    Dim OldRpt As Integer
    Dim OldLine As Integer
    
    'exit if user has canceled out of a multi-select edit
    If MultiFlag And CancelFlag Then
        MultiFlag = False
        MultiLastRow = 0
        CancelFlag = False
        Exit Sub
    End If
    
    'get original report and line number
    OldRpt = Val(frmSpecLayout.lblJRpt.Caption)
    OldLine = Val(frmSpecLayout.lblJLine.Caption)
    
    iRpt = frmSpecLayout.cboReport.ListIndex
    iLine = Val(frmSpecLayout.txtLine.Text)
    
'    'check if user wants to replace another item
'    If iLine <> OldLine Or iRpt <> OldRpt Then
'        S = frmSpecLayout.txtLinesUsed.Text
'        T = ", " & CStr(iLine) & ","
'        N = InStr(S, T)
'        If N > 0 Then
'            rtnval = MsgBox("This line already contains data." & vbCrLf & _
'            "Do you want to overwrite?", vbQuestion + vbOKCancel, "Visual Report Designer")
'            If rtnval = vbCancel Then
'                Exit Sub
'            End If
'        End If
'    End If
    
    If iLine > MaxLines Then
        MaxLines = iLine
        ReDim Preserve BList(0 To 14, 1 To iLine)
    End If
    
    If iLine > ReportInfo(iRpt).NLines Then ReportInfo(iRpt).NLines = iLine
    
    'adjust grid if line doesn't exit yet
    Set G = frmCompose.grdReport(iRpt)
    If G.Rows < iLine + 1 Then
        N = G.Rows
        G.Rows = iLine + 1
        For I = N To iLine
            G.TextMatrix(I, 0) = CStr(I)
        Next I
    End If
    
    If frmSpecLayout.opTag(0).Value = True Then
        BList(iRpt, iLine).Tag = frmSpecLayout.lblTag.Caption
    Else
        BList(iRpt, iLine).Tag = frmSpecLayout.txtTag.Text
    End If
    BList(iRpt, iLine).StartYear = frmSpecLayout.cboFYear.ListIndex + StartYear
    BList(iRpt, iLine).EndYear = frmSpecLayout.cboXYear.ListIndex + StartYear
    BList(iRpt, iLine).Key = Val(frmSpecLayout.lblKey.Caption) '(can move from another line or report)
    BList(iRpt, iLine).Type = frmSpecLayout.cboBins.ListIndex + 2
    BList(iRpt, iLine).Palette = Trim(frmSpecLayout.cboPalette.Text)
    'assign palette to report legends if necessary
    If BList(iRpt, iLine).Type = 4 Then
        Select Case BList(iRpt, iLine).Palette
            Case "Black_to_Red", "Red_to_Black": N = 1
            Case "Red_to_Blue", "Blue_to_Red": N = 2
            Case "White_to_Black", "Black_to_White": N = 3
        End Select
        ReportLegend(iRpt, N, 0) = BList(iRpt, iLine).Palette
    End If
    BList(iRpt, iLine).LowerCut = Val(frmSpecLayout.txtMark(0).Text)
    BList(iRpt, iLine).UpperCut = Val(frmSpecLayout.txtMark(1).Text)
    If frmSpecLayout.chkZero.Value = 0 Then
        BList(iRpt, iLine).ZeroFlag = False
    Else
        BList(iRpt, iLine).ZeroFlag = True
    End If
    
    'put data on grid
    PutReportData iRpt, iLine
    frmCompose.SSTab1.Tab = iRpt
          
    'if user has changed lines or reports, delete old item
    If OldLine <> iLine Or iRpt <> OldRpt Then
        ClearLine OldRpt, OldLine
    End If
    
    'clear report legends if user has changed palette
    CleanReportLegend iRpt
        
    WriteLayoutFile
    HTMLFile = Left(FNLOG, Len(FNLOG) - 4) & "_rpt" & CStr(iRpt + 1) & ".html"
    If SaveLayout Then WriteHTML HTMLFile, iRpt
    If iRpt <> OldRpt And SaveLayout Then
        HTMLFile = Left(FNLOG, Len(FNLOG) - 4) & "_rpt" & CStr(OldRpt + 1) & ".html"
        WriteHTML HTMLFile, OldRpt
    End If
    
    Unload frmSpecLayout
    
    If Not MultiFlag Then
        'un-highlight grid row
        G.Col = 0
        G.ColSel = 0
    End If
    
    If MultiFlag Then
        'if at end of multi-select mode
        If OldLine = MultiLastRow Then
            MultiFlag = False
            MultiLastRow = 0
            'un-highlight grid row
            G.Col = 0
            G.ColSel = 0
        Else 'continue with editing
            InitSpecEdit OldRpt, (OldLine + 1)
        End If
    End If

        
End Sub
Public Sub UpdateLayoutBatch()
    Dim I As Integer
    Dim J As Integer
    Dim iLine As Integer
    Dim iRpt As Integer
    Dim G As MSFlexGrid
    Dim N As Integer
    Dim M As Integer
    Dim S As String
    Dim T As String
    Dim OldRpt As Integer
    Dim OldLine As Integer
    Dim NLines As Integer
    Dim tempData() As DataLayout
        
    'get original report and line number
    OldRpt = Val(frmSpecLayout.lblJRpt.Caption)
    OldLine = Val(frmSpecLayout.lblJLine.Caption)
    
    iRpt = frmSpecLayout.cboReport.ListIndex
    iLine = Val(frmSpecLayout.txtLine.Text)
    NLines = Val(frmSpecLayout.lblNLines.Caption)
    
    'store data in temporary array
    ReDim tempData(1 To NLines)
    For I = 1 To NLines
        tempData(I).Key = BList(OldRpt, OldLine + I - 1).Key
        tempData(I).Tag = BList(OldRpt, OldLine + I - 1).Tag
        tempData(I).StartYear = BList(OldRpt, OldLine + I - 1).StartYear
        tempData(I).EndYear = BList(OldRpt, OldLine + I - 1).EndYear
        tempData(I).Type = BList(OldRpt, OldLine + I - 1).Type
        tempData(I).LowerCut = BList(OldRpt, OldLine + I - 1).LowerCut
        tempData(I).UpperCut = BList(OldRpt, OldLine + I - 1).UpperCut
        tempData(I).Palette = BList(OldRpt, OldLine + I - 1).Palette
        tempData(I).ZeroFlag = BList(OldRpt, OldLine + I - 1).ZeroFlag
    Next I
        
    'adjust variables and grid to accomodate data if needed
    N = iLine + NLines - 1
    If N > MaxLines Then
        MaxLines = N
        ReDim Preserve BList(0 To 14, 1 To N)
    End If
    If N > ReportInfo(iRpt).NLines Then ReportInfo(iRpt).NLines = N
    Set G = frmCompose.grdReport(iRpt)
    If G.Rows < N + 1 Then
        M = G.Rows
        G.Rows = N + 1
        For I = M To N
            G.TextMatrix(I, 0) = CStr(I)
        Next I
    End If
    
    For I = 1 To NLines
        'if line is a blank then clear the line
        If tempData(I).Key = 0 Then
            ClearLine iRpt, (iLine + I - 1)
        Else
            'otherwise, assign new data
            If frmSpecLayout.opTag(0).Value = True Then
                N = tempData(I).Key
                Set G = frmData.MSFlexGrid1
                S = ""
                For J = 0 To 3
                    If frmSpecLayout.chkTag(J).Value = 1 Then
                        S = S & G.TextMatrix(N, J + 1) & " "
                    End If
                Next J
                BList(iRpt, iLine + I - 1).Tag = Trim(S)
            Else
                BList(iRpt, iLine + I - 1).Tag = tempData(I).Tag
            End If
            If frmSpecLayout.opFYear(2).Value = True Then
                BList(iRpt, iLine + I - 1).StartYear = tempData(I).StartYear
            Else
                BList(iRpt, iLine + I - 1).StartYear = frmSpecLayout.cboFYear.ListIndex + StartYear
            End If
            If frmSpecLayout.opXYear(2).Value = True Then
                BList(iRpt, iLine + I - 1).EndYear = tempData(I).EndYear
            Else
                BList(iRpt, iLine + I - 1).EndYear = frmSpecLayout.cboXYear.ListIndex + StartYear
            End If
            BList(iRpt, iLine + I - 1).Key = tempData(I).Key '(can move from another line or report)
            If frmSpecLayout.chkEditDisplay.Value = 0 Then
                BList(iRpt, iLine + I - 1).Type = frmSpecLayout.cboBins.ListIndex + 2
                BList(iRpt, iLine + I - 1).Palette = Trim(frmSpecLayout.cboPalette.Text)
                'assign palette to report legends if necessary
                If BList(iRpt, iLine + I - 1).Type = 4 Then
                    Select Case BList(iRpt, iLine + I - 1).Palette
                        Case "Black_to_Red", "Red_to_Black": N = 1
                        Case "Red_to_Blue", "Blue_to_Red": N = 2
                        Case "White_to_Black", "Black_to_White": N = 3
                    End Select
                    ReportLegend(iRpt, N, 0) = BList(iRpt, iLine + I - 1).Palette
                End If
                BList(iRpt, iLine + I - 1).LowerCut = Val(frmSpecLayout.txtMark(0).Text)
                BList(iRpt, iLine + I - 1).UpperCut = Val(frmSpecLayout.txtMark(1).Text)
            Else
                BList(iRpt, iLine + I - 1).Type = tempData(I).Type
                BList(iRpt, iLine + I - 1).Palette = tempData(I).Palette
                'assign palette to report legends if necessary
                If BList(iRpt, iLine + I - 1).Type = 4 Then
                    Select Case BList(iRpt, iLine + I - 1).Palette
                        Case "Black_to_Red", "Red_to_Black": N = 1
                        Case "Red_to_Blue", "Blue_to_Red": N = 2
                        Case "White_to_Black", "Black_to_White": N = 3
                    End Select
                    ReportLegend(iRpt, N, 0) = BList(iRpt, iLine + I - 1).Palette
                End If
                BList(iRpt, iLine + I - 1).LowerCut = tempData(I).LowerCut
                BList(iRpt, iLine + I - 1).UpperCut = tempData(I).UpperCut
            End If
            If frmSpecLayout.chkEditZero.Value = 0 Then
                If frmSpecLayout.chkZero.Value = 0 Then
                    BList(iRpt, iLine + I - 1).ZeroFlag = False
                Else
                    BList(iRpt, iLine + I - 1).ZeroFlag = True
                End If
            Else
                BList(iRpt, iLine + I - 1).ZeroFlag = tempData(I).ZeroFlag
            End If
            
            'put data on grid
            PutReportData iRpt, (iLine + I - 1)
        End If
    
    Next I
    
    'delete unecessary lines
    For I = NLines To 1 Step -1 'start from bottom of grid to not mess up the line numbering
        N = OldLine + I - 1
        If iRpt <> OldRpt Then
            'if data was on a different report, delete
            DeleteFromLayout OldRpt, N
        Else
            'if data was on the same form, check to make sure
            'that new data didn't overwrite old data before clearing the line
            M = iLine + NLines - 1
            If N < iLine Or N > M Then ClearLine OldRpt, N
        End If
    Next I
          
    'clear report legends if user has changed palette
    If iRpt <> OldRpt Then
        CleanReportLegend OldRpt
    End If
    CleanReportLegend iRpt
        
    WriteLayoutFile
    HTMLFile = Left(FNLOG, Len(FNLOG) - 4) & "_rpt" & CStr(iRpt + 1) & ".html"
    If SaveLayout Then WriteHTML HTMLFile, iRpt
    If iRpt <> OldRpt And SaveLayout Then
        HTMLFile = Left(FNLOG, Len(FNLOG) - 4) & "_rpt" & CStr(OldRpt + 1) & ".html"
        WriteHTML HTMLFile, OldRpt
    End If
    
    Unload frmSpecLayout
    
    'de-select lines
    Set G = frmCompose.grdReport(OldRpt)
    G.Row = 0
    G.RowSel = 0
    
    frmCompose.SSTab1.Tab = iRpt
    
End Sub
Public Sub ClearRptLegend(Index As Integer, Line As Integer)
'Dim I As Integer
'Dim G As MSFlexGrid
'Dim N As Integer
'Dim OpColor As String
'Dim OpFlag As Boolean
'
'Set G = frmCompose.grdReport(Index)
'
''get opposite color palette in case report uses both
'Select Case BList(Index, Line).Palette
'    Case "Black_to_Red":
'        OpColor = "Red_to_Black"
'        N = 1
'    Case "Red_to_Black":
'        OpColor = "Black_to_Red"
'        N = 1
'    Case "Red_to_Blue":
'        OpColor = "Blue_to_Red"
'        N = 2
'    Case "Blue_to_Red":
'        OpColor = "Red_to_Blue"
'        N = 2
'    Case "White_to_Black":
'        OpColor = "Black_to_White"
'        N = 3
'    Case "Black_to_White":
'        OpColor = "White_to_Black"
'        N = 3
'End Select
'
''check to see if there are any other lines with same color palette or opposite color palette
'OpFlag = False
'For I = 1 To Line - 1
'    If BList(Index, I).Palette = BList(Index, Line).Palette Then
'        Exit Sub
'    ElseIf BList(Index, I).Palette = OpColor Then
'        OpFlag = True
'    End If
'Next I
'For I = Line + 1 To G.Rows - 1
'    If BList(Index, I).Palette = BList(Index, Line).Palette Then
'        Exit Sub
'    ElseIf BList(Index, I).Palette = OpColor Then
'        OpFlag = True
'    End If
'Next I
'
''if there are none with the same color palette but one or more with opposite palette
''then reset the report legend index
'If OpFlag Then
'    ReportLegend(Index, N, 0) = OpColor
'Else
'    ReportLegend(Index, N, 0) = ""
'End If

End Sub
Public Sub CleanReportLegend(Index As Integer)
Dim G As MSFlexGrid
Dim I As Integer
Dim N As Integer
Dim S As String

'clear legend palette indicator
For I = 1 To 3
    ReportLegend(Index, I, 0) = ""
Next I
Set G = frmCompose.grdReport(Index)
For I = 1 To G.Rows - 1
    If BList(Index, I).Type = 4 Then
        S = BList(Index, I).Palette
        Select Case S
            Case "Black_to_Red", "Red_to_Black":
                N = 1
            Case "Red_to_Blue", "Blue_to_Red":
                N = 2
            Case "White_to_Black", "Black_to_White":
                N = 3
        End Select
        If S <> "" Then ReportLegend(Index, N, 0) = S
    End If
Next I

End Sub
Public Sub ClearLine(Index As Integer, Line As Integer)
    Dim I As Integer
    Dim G As MSFlexGrid
    Dim flag As Boolean
    
    Set G = frmCompose.grdReport(Index)
    
    BList(Index, Line).Tag = ""
    BList(Index, Line).StartYear = 0
    BList(Index, Line).EndYear = 0
    BList(Index, Line).Key = 0
    'clear report legend if needed
    'If BList(Index, Line).Type = 4 Then ClearRptLegend Index, Line
    BList(Index, Line).Palette = ""
    BList(Index, Line).Type = 0
    BList(Index, Line).LowerCut = 0#
    BList(Index, Line).UpperCut = 0#
    BList(Index, Line).ZeroFlag = True
    For I = 1 To G.Cols - 1
        G.TextMatrix(Line, I) = ""
    Next I
    CleanReportLegend Index
        
End Sub
Public Sub DeleteFromLayout(Index As Integer, Line As Integer)
    Dim I As Integer
    Dim J As Integer
    Dim G As MSFlexGrid
    
    Set G = frmCompose.grdReport(Index)
    
    If Line < G.Rows - 1 Then
        'move data up one row
        For I = Line + 1 To G.Rows - 1
            For J = 1 To G.Cols - 1
                G.TextMatrix(I - 1, J) = G.TextMatrix(I, J)
            Next J
        Next I
        'move data variables
        For I = Line + 1 To G.Rows - 1
            BList(Index, I - 1).Tag = BList(Index, I).Tag
            BList(Index, I - 1).StartYear = BList(Index, I).StartYear
            BList(Index, I - 1).EndYear = BList(Index, I).EndYear
            BList(Index, I - 1).Key = BList(Index, I).Key
            BList(Index, I - 1).Type = BList(Index, I).Type
            BList(Index, I - 1).Palette = BList(Index, I).Palette
            BList(Index, I - 1).LowerCut = BList(Index, I).LowerCut
            BList(Index, I - 1).UpperCut = BList(Index, I).UpperCut
            BList(Index, I - 1).UpperCut = BList(Index, I).UpperCut
        Next I
    End If
    'clear last row's data
    I = G.Rows - 1
    BList(Index, I).Tag = ""
    BList(Index, I).StartYear = 0
    BList(Index, I).EndYear = 0
    BList(Index, I).Key = 0
    'clear report legend if needed
    'If BList(Index, I).Type = 4 Then ClearRptLegend Index, I
    BList(Index, I).Palette = ""
    BList(Index, I).Type = 0
    BList(Index, I).LowerCut = 0#
    BList(Index, I).UpperCut = 0#
    BList(Index, I).ZeroFlag = True
    G.Rows = G.Rows - 1

    ReportInfo(Index).NLines = G.Rows
    
    CleanReportLegend Index
    
End Sub
Public Sub SaveLayoutFile()
    Dim I As Integer
    
    WriteLayoutFile
    If SaveLayout Then
        For I = 0 To 14
            HTMLFile = Left(FNLOG, Len(FNLOG) - 4) & "_rpt" & CStr(I + 1) & ".html"
            WriteHTML HTMLFile, I
        Next I
    End If
    
End Sub
Public Sub WriteLayoutFile()
    Dim I As Integer
    Dim K1 As Integer
    Dim K2 As Integer
    Dim N As Integer
    
    On Error GoTo WriteErr
    
    SaveLayout = False
    
    'check that the last year comes after the first year
    For I = 1 To NForms
        K1 = Val(frmCompose.cboStartYr(I - 1).Text)
        K2 = Val(frmCompose.cboEndYr(I - 1).Text)
        If K1 > K2 Or K1 < StartYear Or K2 > EndYear Then
            MsgBox "Invalid Year Range for Report Form" & CStr(I), vbInformation, "NFT"
            Exit Sub
        End If
    Next I
    
    FNTXT = frmGeneral.lblFile.Caption
    N = InStrRev(FNTXT, ".")
    Mid(FNTXT, N, 4) = ".txt"
    Open FNTXT For Output As #1
    
    Print #1, "$VisualReport Version " & VRVersion
    For I = 1 To NForms
        WriteReportData (I - 1)
    Next I
    Print #1, "$END"
    
    'now print what missing value indicator was used
    Print #1, "$Missing Value Indicator = " + NAVal
    
    Close #1
    
    SaveLayout = True

    Exit Sub
WriteErr:
    Close #1
    MsgBox "Error Writing Layout File", vbExclamation, "Visual Report Designer"
End Sub
Public Sub WriteReportData(Index As Integer)
    Dim I As Integer
    Dim J As Integer
    Dim K As Integer
    Dim K1 As Integer
    Dim K2 As Integer
    Dim NY As Integer
    Dim S As String
    Dim T As String
    Dim G As MSFlexGrid
    Dim H As MSFlexGrid
    
    Set G = frmCompose.grdReport(Index)
    Set H = frmData.MSFlexGrid1
    
    'exit if report doesn't contain any data
    If G.Rows = 1 Then Exit Sub
    
    Print #1, "$Report" + CStr(Index + 1)
    K1 = Val(frmCompose.cboStartYr(Index).Text)
    K2 = Val(frmCompose.cboEndYr(Index).Text)
    NY = K2 - K1 + 1
    S = Trim(frmCompose.txtTitle(Index).Text)
    ReportInfo(Index).Title = S
    Print #1, S
    S = CStr(K1) + Space(2) + CStr(K2)
    Print #1, S
    For I = 1 To 3
        If ReportLegend(Index, I, 0) <> "" Then
            S = "Legend_" & CStr(I) & ": " & ReportLegend(Index, I, 0)
            For J = 1 To 5
                S = S & " Label_" & CStr(J) & ": " & ReportLegend(Index, I, J)
            Next J
            Print #1, S
        End If
    Next I
    Print #1, "Print_Dispersion: " & CStr(DoStat(Index)) & "  " & CStr(StatDig(Index))
    Print #1, "CutPoint_Location: " & CutPtLoc(Index)
    For I = 1 To G.Rows - 1
        If BList(Index, I).Key > 0 Then
            Print #1, BList(Index, I).Tag
            Print #1, BList(Index, I).StartYear, BList(Index, I).EndYear, I, BList(Index, I).Type, _
                BList(Index, I).Palette, BList(Index, I).LowerCut, BList(Index, I).UpperCut, BList(Index, I).Key, BList(Index, I).ZeroFlag
            S = ""
            For J = 1 To NYears
                K = StartYear + J - 1
                If K >= BList(Index, I).StartYear And K <= BList(Index, I).EndYear Then
                    S = S + H.TextMatrix(BList(Index, I).Key, J + 4) + Chr(9)
                End If
            Next J
            Print #1, S
        End If
    Next I
    
End Sub

Private Sub DeleteRow(F As Form, J As Integer)
    Dim I As Integer
    Dim K As Integer
    
    
    F.Label1(J - 1).Visible = False
    
    For I = 0 To KCols - 1
        K = I + (J - 1) * KCols
        F.Picture1(K).Visible = False
    Next I
End Sub
Private Sub SortVectorUp(N As Long, X() As Double)
    Dim I As Long
    Dim J As Long
    Dim T As Double
    
    
    ' Bubble Sort of Vector in Ascending order
    For I = 1 To N - 1
        For J = I + 1 To N
            If X(J) < X(I) Then
                T = X(I)
                X(I) = X(J)
                X(J) = T
            End If
        Next J
    Next I
    
End Sub
Public Sub WriteGridData()
    Dim G As MSFlexGrid
    Dim I As Integer
    Dim J As Integer
    Dim K As Integer
    Dim N As Integer
    Dim S As String
    Dim T As String
    Dim FN As String
    
    On Error GoTo WrtErr
    
    DataSavedFlag = False 'flag to indicate sucessful write
    
    Set G = frmData.MSFlexGrid1
    
    FN = FNLOG
    N = InStrRev(FN, ".")
    Mid(FN, N, 4) = ".csv"
    
    K = G.Rows
    N = G.Cols
    
    Open FN For Output As #2
    
    For I = 0 To K - 1
        S = ""
        For J = 0 To N - 1
            S = S + G.TextMatrix(I, J)
            If J < N - 1 Then
                S = S + ","
            End If
        Next J
        Print #2, S
    Next I

    Close #2
    
    Print #3, ""
    Print #3, "Data Saved to CSV File"
    Print #3, FN
    Print #3, DtStr
    Print #3, ""
    
    DataSavedFlag = True
    
    Exit Sub
WrtErr:
    Close #2
    MsgBox "Error Saving Grid Data: " + vbCrLf + Err.Description, vbExclamation, "Visual Report Designer"
End Sub
Public Sub WriteDatabaseStyle()
    Dim G As MSFlexGrid
    Dim I As Integer
    Dim J As Integer
    Dim K As Integer
    Dim N As Integer
    Dim S As String
    Dim T As String
    Dim X As String
    Dim FN As String
    
    
    Set G = frmData.MSFlexGrid1
    
    FN = FNLOG
    N = InStrRev(FN, ".")
    FN = Left(FNLOG, N - 1) + "_db.csv"
    
    K = G.Rows
    N = G.Cols
    
    Open FN For Output As #2
    
    Print #2, "Source,Case,Data Type,Item,Year,Value"
    
    
    For I = 1 To K - 1
        S = ""
        For J = 1 To 4
            S = S + G.TextMatrix(I, J) + ","
        Next J
        For J = 1 To NYears
            T = G.TextMatrix(I, J + 4)
            If T <> "" Then
                X = S + CStr(J + StartYear - 1) + "," + T
                Print #2, X
            End If
        Next J
    Next I

    Close #2
    
    Print #3, ""
    Print #3, "Data Saved to Database Style File"
    Print #3, FN
    Print #3, DtStr
    Print #3, ""
    
    
    MsgBox "Grid Data Has Been Saved to File", vbInformation, "Visual Report Designer"
End Sub
Public Function DtStr() As String
    Dim D As Date
    Dim S As String
    
    
    D = Now
    S = Format(Month(D), "00") + "-"
    S = S + Format(Day(D), "00") + "-"
    S = S + Format(Year(D), "0000") + Space(2)
    S = S + Format(Hour(D), "00") + ":"
    S = S + Format(Minute(D), "00")
    DtStr = S

End Function
Public Function DtStrCompact() As String
    'formats current date and time as YYMMDDHHmmss, where
    '  YY = year, MM = month, DD = day, HH = hour, mm = minute, and ss = second
    Dim D As Date
    Dim S As String
    
    D = Now
    S = Right(CStr(Year(D)), 2)
    S = S + Format(Month(D), "00")
    S = S + Format(Day(D), "00")
    S = S + Format(Hour(D), "00")
    S = S + Format(Minute(D), "00")
    S = S + Format(Second(D), "00")
    DtStrCompact = S

End Function

Public Sub OpenProjectLog()
    Dim DT As String
    Dim N As Integer
    Dim RS As Integer
    Dim TempLOG As String
    Dim S As String

    On Error Resume Next
        
    frmGeneral.CommonDialog1.FileName = ""
    frmGeneral.CommonDialog1.Flags = &H1004
    frmGeneral.CommonDialog1.DialogTitle = "Open Existing Project Log File"
    frmGeneral.CommonDialog1.Filter = "Log Files (*.log)|*.log"
    frmGeneral.CommonDialog1.CancelError = True
    frmGeneral.CommonDialog1.FilterIndex = 0
    frmGeneral.CommonDialog1.CancelError = False
    frmGeneral.CommonDialog1.ShowOpen
    TempLOG = frmGeneral.CommonDialog1.FileName
    N = InStr(UCase(TempLOG), "ST6UNST.LOG")
    If N > 0 Then
        MsgBox TempLOG + vbCrLf + "Is is Not A Valid Project File", vbInformation, "Visual Report Designer"
        Exit Sub
    End If
    If TempLOG <> "" Then
        'close file before opening new one
        If FNLOG <> "" Then Close #3
        
        Unload frmData
        Unload frmCompose
        
        FF1 = False
        
        FNLOG = TempLOG
        frmGeneral.lblFile.Caption = FNLOG
        S = "Visual Report Designer - Version " & CStr(VRVersion) & " - "
        S = S & FNLOG
        frmMain.Caption = S
        DT = DtStr
        Open FNLOG For Append As #3
        Print #3, "*** Project Log File Appended: " + DT + " ***"
        LogFlag = True
        DataFlag = False
        LayoutFlag = False
        frmMain.mnuDataMissingReplace.Enabled = False
        ReadGridData
        If DataFlag Then
            S = Left(FNLOG, Len(FNLOG) - 3) + "txt"
            If Dir(S) <> "" Then
                RS = MsgBox("Do You Wish to Add Existing Layouts Back into the Project?", vbQuestion + vbYesNo, "Visual Report Designer")
                If RS = vbYes Then
                    ReadLayoutFile
                    If LayoutFlag = True Then MsgBox "All Report Layouts Have Been Restored", vbInformation, "Visual Report Designer"
                End If
            End If
        Else 'if there is no data file, allow users to begin adding data
            StartYear = 0
            EndYear = 0
            NYears = 0
            frmMain.mnuDataAdd.Enabled = True
            'show the data collection form so users have a better idea what to do next
            InitDataCollectionGrid
        End If
        
        'backup the new file
        'first make sure there is a file available
        S = Trim(frmGeneral.lblFile.Caption)
        If S <> "" And LCase(S) <> "none specified" Then
            'do backup
            CreateBackup S, "File-Open"
        End If
    End If
    
    Exit Sub
    
    
End Sub
Public Sub CreateNewLogFile()
    Dim DT As String
    Dim N As Integer
    Dim TempLOG As String
    Dim S As String

    On Error Resume Next
    
    frmGeneral.CommonDialog1.FileName = ""
    frmGeneral.CommonDialog1.Flags = &H806
    frmGeneral.CommonDialog1.DialogTitle = "Create New Project Log File"
    frmGeneral.CommonDialog1.Filter = "Log Files (*.log)|*.log"
    frmGeneral.CommonDialog1.CancelError = True
    frmGeneral.CommonDialog1.FilterIndex = 0
    frmGeneral.CommonDialog1.DefaultExt = "log"
    frmGeneral.CommonDialog1.CancelError = False
    frmGeneral.CommonDialog1.ShowSave
    TempLOG = frmGeneral.CommonDialog1.FileName
    N = InStr(UCase(TempLOG), "ST6UNST.LOG")
    If N > 0 Then
        MsgBox TempLOG + vbCrLf + "Is Not A Valid Project File Name", vbInformation, "Visual Report Designer"
        Exit Sub
    End If

    If TempLOG <> "" Then
        'if file doesn't have ".log" appended (when user types a file name with
        ' a period in it and also doesn't explictly type ".log")
        If LCase(Right(TempLOG, 4)) <> ".log" Then
            TempLOG = TempLOG + ".log"
        End If
        
        'close file before opening new one
        If FNLOG <> "" Then Close #3
    
        Unload frmData
        KRow = 0
        Unload frmCompose
        
        FNLOG = TempLOG
        frmGeneral.lblFile.Caption = FNLOG
        S = "Visual Report Designer - Version " & CStr(VRVersion) & " - "
        S = S & FNLOG
        frmMain.Caption = S
        DT = DtStr
        Open FNLOG For Output As #3
        Print #3, "*** Project Log File Opened: " + DT + " ***"
        LogFlag = True
        DataFlag = False
        LayoutFlag = False
        FF1 = False
        
        StartYear = 0
        EndYear = 0
        NYears = 0
        frmMain.mnuDataAdd.Enabled = True
        frmMain.mnuDataMissingReplace.Enabled = False
        
        'show the data collection grid so users have better idea of what to do next
        InitDataCollectionGrid
    End If
    
    
    Exit Sub
    
    
End Sub
Public Sub SaveProjectAs()
    Dim DT As String
    Dim N As Integer
    Dim TempLOG As String
    Dim OldLog As String
    Dim I As Integer
    Dim S As String

    On Error Resume Next
    
    If FNLOG = "" Or frmGeneral.lblFile.Caption = "None Specified" Then
        CreateNewLogFile
        Exit Sub
    End If
    
    OldLog = FNLOG
    
    frmGeneral.CommonDialog1.FileName = ""
    frmGeneral.CommonDialog1.Flags = &H806
    frmGeneral.CommonDialog1.DialogTitle = "Create New Project Log File"
    frmGeneral.CommonDialog1.Filter = "Log Files (*.log)|*.log"
    frmGeneral.CommonDialog1.CancelError = True
    frmGeneral.CommonDialog1.FilterIndex = 0
    frmGeneral.CommonDialog1.DefaultExt = "log"
    frmGeneral.CommonDialog1.CancelError = False
    frmGeneral.CommonDialog1.ShowSave
    TempLOG = frmGeneral.CommonDialog1.FileName
    N = InStr(UCase(TempLOG), "ST6UNST.LOG")
    If N > 0 Then
        MsgBox TempLOG + vbCrLf + "Is Not A Valid Project File Name", vbInformation, "Visual Report Designer"
        Exit Sub
    End If

    If TempLOG <> "" Then
        'if file doesn't have ".log" appended (when user types a file name with
        ' a period in it and also doesn't explictly type ".log")
        If LCase(Right(TempLOG, 4)) <> ".log" Then
            TempLOG = TempLOG + ".log"
        End If
        'close file before opening new one
        If FNLOG <> "" Then
            Close #3
            FileCopy FNLOG, TempLOG
        End If
        
        FNLOG = TempLOG
        frmGeneral.lblFile.Caption = FNLOG
        S = "Visual Report Designer - Version " & CStr(VRVersion) & " - "
        S = S & FNLOG
        frmMain.Caption = S
        DT = DtStr
        Open FNLOG For Append As #3
        Print #3, ""
        Print #3, "*** Project Renamed: " + DT + " ***"
        Print #3, "Old Log File: " + OldLog
        
        If DataFlag Then
            WriteGridData
            If DataSavedFlag Then
                If Not InitFlag And LayoutFlag Then
                    WriteLayoutFile
                    For I = 0 To 14
                        HTMLFile = Left(FNLOG, Len(FNLOG) - 4) & "_rpt" & CStr(I + 1) & ".html"
                        If SaveLayout Then WriteHTML HTMLFile, I
                    Next I
                End If
            End If
        End If
        
    End If
    
    
    Exit Sub
    
    
End Sub
Public Sub InitUserDataGrid()
    Dim I As Integer
    Dim K1 As Integer
    Dim K2 As Integer
    Dim N As Integer
    Dim NY As Integer
    Dim G As MSFlexGrid
    
    For I = 0 To 2
        If frmSpecUserAdded.txtUser(I).Text = "" Then
            MsgBox "Please Specify User Added Data Parameters", vbInformation, "Visual Report Designer"
            Exit Sub
        End If
    Next I
    K1 = Val(frmSpecUserAdded.txtUser(0).Text)
    K2 = Val(frmSpecUserAdded.txtUser(1).Text)
    N = Val(frmSpecUserAdded.txtUser(2).Text)
    Unload frmSpecUserAdded
    
    NY = K2 - K1 + 1
    
    If K1 > K2 Then
        MsgBox "Invalid Year Range for User Added Data", vbInformation, "Visual Report Designer"
        Exit Sub
    End If
        
    Set G = frmAux.MSFlexGrid1
    
    G.Cols = NY + 4
    G.FixedCols = 1
    G.Rows = N + 1
    G.TextMatrix(0, 0) = "Source"
    G.TextMatrix(0, 1) = "Case"
    G.TextMatrix(0, 2) = "Data Type"
    G.TextMatrix(0, 3) = "Item"
    G.ColWidth(0) = 1000
    G.ColWidth(1) = 2500
    G.ColWidth(2) = 2500
    G.ColWidth(3) = 1100
    
    For I = 1 To NY
        G.ColWidth(I + 3) = 1600
        G.TextMatrix(0, I + 3) = CStr(K1 + I - 1)
    Next I
    
    For I = 1 To N
        G.TextMatrix(I, 0) = "USER"
    Next I

    frmAux.Show
    
End Sub

Public Sub AddUserDataToGrid()
    Dim I As Integer
    Dim J As Integer
    Dim K As Integer
    Dim tmpStartYear As Integer
    Dim tmpEndYear As Integer
    Dim G As MSFlexGrid
    Dim T As String
    Dim flag As Boolean
    Dim iRows As Integer
    Dim iCols As Integer
    Dim data() As String
    
    On Error GoTo AddErr
    
    Set G = frmAux.MSFlexGrid1
    'get user-supplied data first (SetParms will unload frmAux)
    iRows = G.Rows - 1
    iCols = G.Cols - 1
    ReDim data(1 To iRows, 0 To iCols)
    For I = 1 To iRows
        For J = 0 To iCols
            data(I, J) = G.TextMatrix(I, J)
        Next J
    Next I
    
    'get first and last year
    tmpStartYear = Val(G.TextMatrix(0, 4))
    tmpEndYear = Val(G.TextMatrix(0, G.Cols - 1))
    
    'make room in data grid, if needed
    flag = False
    If StartYear = 0 And EndYear = 0 Then
        frmGeneral.lblStartYr.Caption = CStr(tmpStartYear)
        frmGeneral.lblEndYr.Caption = CStr(tmpEndYear)
        flag = True
    Else
        If tmpStartYear < StartYear Then
            frmGeneral.lblStartYr.Caption = CStr(tmpStartYear)
            flag = True
        End If
        If tmpEndYear > EndYear Then
            frmGeneral.lblEndYr.Caption = CStr(tmpEndYear)
            flag = True
        End If
    End If
    If flag Then SetParms
    
    K = KRow
    
    Set G = frmData.MSFlexGrid1
    
    G.Rows = K + iRows + 1
    
    'add line numbers to data grid
    For I = 1 To iRows
        G.TextMatrix(K + I, 0) = CStr(K + I)
    Next I
    
    'add tags from user data to data grid
    For I = 1 To iRows
        For J = 1 To 4
            G.TextMatrix(K + I, J) = data(I, J - 1)
        Next J
    Next I
    
    'add user data to data grid
    For I = 1 To iRows
        For J = 0 To tmpEndYear - tmpStartYear
            T = data(I, J + 4)
            G.TextMatrix(K + I, J + tmpStartYear - StartYear + 5) = T
        Next J
    Next I
    
    KRow = KRow + iRows
    
    Print #3, "USER"
    Print #3, CStr(iRows) + " Lines of User Data Added"
    Print #3, CStr(tmpStartYear) + " to " + CStr(tmpEndYear)
    Print #3, ""
    
    K = G.Rows
    
    Unload frmAux
    
    DataFlag = True
    
    frmMain.mnuDataMissingReplace.Enabled = True
    
    'automatically save data into file
    WriteGridData

    MsgBox "User Data Has Been Added to Data Collection Grid", vbInformation, "Visual Report Designer"
    
    Exit Sub
AddErr:
    MsgBox "Error Adding User Data to Data Collection Grid" + vbCrLf + _
        Err.Description, vbExclamation, "Visual Report Designer"
End Sub
Private Sub ScanCSAGeneralData()
    Dim N As Long
    Dim Buffer As String
    Dim Token As String
    
    On Error GoTo ScanCSAErr
    
    FNOUT = FName
    N = InStrRev(FNOUT, ".")
    Mid(FNOUT, N, 4) = ".tmp"
    
    If Dir(FNOUT) = "" Then
        MsgBox "CSA Data Scan Failed" + vbCrLf + "CSA Output Files May Be Incomplete", vbExclamation, "Visual Report Designer"
        Exit Sub
    End If
    
    Open FNOUT For Input As #5
    
    Line Input #5, Buffer
    
    frmCSAScan.lblFile.Caption = FNOUT
    frmCSAScan.txtCase.Text = Buffer
    
    Line Input #5, Buffer
    
    Token = GetFirstToken(Buffer)
    KYears = Val(Token)
    Token = GetNextToken(Buffer)
    KFYear = Val(Token)
    Token = GetNextToken(Buffer)
    KModel = Val(Token)


    frmCSAScan.lblYear(0).Caption = CStr(KFYear)
    KXYear = KFYear + KYears - 1
    frmCSAScan.lblYear(1).Caption = CStr(KXYear)
    
    If KModel = 1 Then
        frmCSAScan.lblModelType.Caption = "Process Error Only"
    ElseIf KModel = 2 Then
        frmCSAScan.lblModelType.Caption = "Observed Error Only"
    Else
        frmCSAScan.lblModelType.Caption = "Both Process && Observed Error"
    End If
    frmCSAScan.Show

    CSAFlag = True
'    If KFYear < StartYear Or KXYear > EndYear Then
'        MsgBox "Invalid Year Range Specification for Data Grid", vbInformation, "Visual Report Designer"
'        Close #5
'        CSAFlag = False
'    Else
'        CSAFlag = True
'    End If
'
'    If KFYear < MinYear Then
'        MinYear = KFYear
'    End If
'    If KXYear > MaxYear Then
'        MaxYear = KXYear
'    End If
    
    Exit Sub
    
ScanCSAErr:
    MsgBox "Error Scanning CSA Results", vbInformation, "Visual Report Designer"
    Close #5
End Sub
Public Sub ScanCSAResults()
    Dim Buffer As String
    Dim Token As String
    Dim G As MSFlexGrid
    Dim I As Integer
    Dim J As Integer
    Dim N As Integer
    Dim K As Integer
    Dim KR As Integer
    Dim DT As String
    
    If CSAFlag = False Then
        Unload frmCSAScan
        Exit Sub
    End If

    KR = KRow
    
    On Error GoTo ScanCSAError
    
    CaseID = Trim(frmCSAScan.txtCase.Text)
    
    'get years and reset forms if needed
    I = Val(frmCSAScan.lblYear(0).Caption)
    J = Val(frmCSAScan.lblYear(1).Caption)
    If StartYear = 0 And EndYear = 0 Then
        frmGeneral.lblStartYr.Caption = CStr(I)
        frmGeneral.lblEndYr.Caption = CStr(J)
        SetParms
    ElseIf I < StartYear Or J > EndYear Then
        If I < StartYear Then frmGeneral.lblStartYr.Caption = CStr(I)
        If J > EndYear Then frmGeneral.lblEndYr.Caption = CStr(J)
        SetParms
    End If
    
    Unload frmCSAScan
    
    Print #3, "CSA"
    Print #3, CaseID
    Print #3, FName
    DT = GetFileWriteTime(FName)
    Print #3, DT
    Print #3, ""
    
    
    
    Set G = frmData.MSFlexGrid1
    G.ColWidth(4) = 2500
    N = KFYear - StartYear + 4
    
    Line Input #5, Buffer
    G.Rows = G.Rows + 2
    K = KRow + 1
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "CSA"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Survey Weight"
    G.TextMatrix(K, 4) = "Recruit"
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    K = KRow + 2
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "CSA"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Survey Weight"
    G.TextMatrix(K, 4) = "Post Recruit"
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    
    KRow = KRow + 2
    
    Line Input #5, Buffer
    G.Rows = G.Rows + 2
    K = KRow + 1
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "CSA"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Catch Weight"
    G.TextMatrix(K, 4) = "Landing"
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    K = KRow + 2
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "CSA"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Catch Weight"
    G.TextMatrix(K, 4) = "Discard"
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    KRow = KRow + 2
    
    Line Input #5, Buffer
    G.Rows = G.Rows + 6
    
    
    K = KRow + 1
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "CSA"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Relative Abundance"
    G.TextMatrix(K, 4) = "Obs. Recruits"
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    K = KRow + 2
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "CSA"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Relative Abundance"
    G.TextMatrix(K, 4) = "Est. Recruits"
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    K = KRow + 3
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "CSA"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Relative Abundance"
    G.TextMatrix(K, 4) = "Resid. Recruits"
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    K = KRow + 4
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "CSA"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Relative Abundance"
    G.TextMatrix(K, 4) = "Obs. PostRecruits"
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    K = KRow + 5
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "CSA"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Relative Abundance"
    G.TextMatrix(K, 4) = "Est. PostRecruits"
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    K = KRow + 6
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "CSA"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Relative Abundance"
    G.TextMatrix(K, 4) = "Resid. PostRecruits"
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    KRow = KRow + 6
    
    If KModel = 3 Then
        Line Input #5, Buffer
        G.Rows = G.Rows + 3
  
        K = KRow + 1
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "CSA"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = "Process Error"
        G.TextMatrix(K, 4) = "Calculated"
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For J = 1 To KYears
            G.TextMatrix(K, J + N) = Token
            Token = GetNextToken(Buffer)
        Next J
    
        K = KRow + 2
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "CSA"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = "Process Error"
        G.TextMatrix(K, 4) = "Estimated"
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For J = 1 To KYears
            G.TextMatrix(K, J + N) = Token
            Token = GetNextToken(Buffer)
        Next J
    
        K = KRow + 3
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "CSA"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = "Process Error"
        G.TextMatrix(K, 4) = "Residual"
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For J = 1 To KYears
            G.TextMatrix(K, J + N) = Token
            Token = GetNextToken(Buffer)
        Next J
    
        KRow = KRow + 3
    
    End If
    
    Line Input #5, Buffer
    G.Rows = G.Rows + 1
    
    K = KRow + 1
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "CSA"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Catchability Ratio"
    G.TextMatrix(K, 4) = "Input"
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    KRow = KRow + 1
    
    Line Input #5, Buffer
    G.Rows = G.Rows + 3
    
    
    K = KRow + 1
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "CSA"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Population"
    G.TextMatrix(K, 4) = "Recruits"
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    K = KRow + 2
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "CSA"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Population"
    G.TextMatrix(K, 4) = "Post-Recruits"
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    K = KRow + 3
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "CSA"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Population"
    G.TextMatrix(K, 4) = "Total"
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    KRow = KRow + 3
    
    Line Input #5, Buffer
    G.Rows = G.Rows + 3
    
    
    K = KRow + 1
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "CSA"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Biomass"
    G.TextMatrix(K, 4) = "Recruits"
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    K = KRow + 2
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "CSA"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Biomass"
    G.TextMatrix(K, 4) = "Post-Recruits"
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    K = KRow + 3
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "CSA"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Biomass"
    G.TextMatrix(K, 4) = "Total"
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    KRow = KRow + 3
    
    Line Input #5, Buffer
    G.Rows = G.Rows + 3
    
    
    K = KRow + 1
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "CSA"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Catch"
    G.TextMatrix(K, 4) = "Landings"
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    K = KRow + 2
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "CSA"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Catch"
    G.TextMatrix(K, 4) = "Discards"
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    K = KRow + 3
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "CSA"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Catch"
    G.TextMatrix(K, 4) = "Total"
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    KRow = KRow + 3
    
    Line Input #5, Buffer
    G.Rows = G.Rows + 3
    
    
    K = KRow + 1
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "CSA"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Catch Biomass"
    G.TextMatrix(K, 4) = "Landings"
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    K = KRow + 2
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "CSA"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Catch Biomass"
    G.TextMatrix(K, 4) = "Discards"
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    K = KRow + 3
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "CSA"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Catch Biomass"
    G.TextMatrix(K, 4) = "Total"
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    KRow = KRow + 3

    Line Input #5, Buffer
    G.Rows = G.Rows + 3
    
    
    K = KRow + 1
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "CSA"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Mortality"
    G.TextMatrix(K, 4) = "Total"
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    K = KRow + 2
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "CSA"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Mortality"
    G.TextMatrix(K, 4) = "Natural"
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    K = KRow + 3
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "CSA"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Mortality"
    G.TextMatrix(K, 4) = "Fishing"
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    KRow = KRow + 3
    
    Line Input #5, Buffer
    G.Rows = G.Rows + 3
    
    
    K = KRow + 1
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "CSA"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Harvest Rate"
    G.TextMatrix(K, 4) = "Combined"
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    K = KRow + 2
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "CSA"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Harvest Rate"
    G.TextMatrix(K, 4) = "Landings"
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    K = KRow + 3
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "CSA"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Harvest Rate"
    G.TextMatrix(K, 4) = "Harvest-F"
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    KRow = KRow + 3

    Line Input #5, Buffer
    G.Rows = G.Rows + 1
    
    
    K = KRow + 1
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "CSA"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Surplus Production"
    G.TextMatrix(K, 4) = ""
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    KRow = KRow + 1
    
frmCompose.Show
frmData.ZOrder
    K = G.Rows
    
    MsgBox "CSA Data Scan Completed" + vbCrLf + CStr(KRow - KR) + " Rows Added to Data Grid", vbInformation, "Visual Report Designer"
    Close #5
    DataFlag = True
    
    frmMain.mnuDataMissingReplace.Enabled = True
    
    'automatically save data into file
    WriteGridData
    
    Exit Sub
    
ScanCSAError:
    Close #5
    MsgBox "Error Scanning CSA Model Results" + vbCrLf + Err.Description, vbInformation, "Visual Report Designer"
End Sub
Private Function CheckAspicFiles() As Boolean
    Dim N As Long
    
    On Error GoTo CheckErr2
    
    FNOUT = FName
    N = InStrRev(FNOUT, ".")
    Mid(FNOUT, N, 4) = ".fit"
    If Dir(FNOUT) = "" Then
        Mid(FNOUT, N, 4) = ".bot"
    End If
    If CompareFileWriteTimes(FName, FNOUT) Then
        MsgBox "Aspic Output Files Are Not Current", vbInformation, "Visual Report Designer"
        CheckAspicFiles = False
        Exit Function
    End If
    Mid(FNOUT, N, 4) = ".tmp"
    If Dir(FNOUT) <> "" Then
        Kill FNOUT
    End If
    CheckAspicFiles = True
    Exit Function
    
CheckErr2:
    MsgBox "Error Checking Aspic Files" + vbCrLf + Err.Description, vbExclamation, "Visual Report Designer"
End Function
Private Sub ScanAspicGeneralData()
    Dim N As Long
    Dim Buffer As String
    Dim Token As String
    
    On Error GoTo ScanAspicErr
    
    FNOUT = FName
    N = InStrRev(FNOUT, ".")
    Mid(FNOUT, N, 4) = ".tmp"
    
    If Dir(FNOUT) = "" Then
        MsgBox "Aspic Data Scan Failed" + vbCrLf + "Aspic Output Files May Be Incomplete", vbExclamation, "Visual Report Designer"
        Exit Sub
    End If
    
    Open FNOUT For Input As #5
    
    Line Input #5, Buffer
    
    frmAspicScan.lblFile.Caption = FNOUT
    frmAspicScan.txtCase.Text = Buffer
    
    Line Input #5, Buffer
    
    Token = GetFirstToken(Buffer)
    KYears = Val(Token)
    Token = GetNextToken(Buffer)
    KFYear = Val(Token)
    Token = GetNextToken(Buffer)
    KINDX = Val(Token)
    ModelType = GetNextToken(Buffer)

    frmAspicScan.lblYear(0).Caption = CStr(KFYear)
    KXYear = KFYear + KYears - 1
    frmAspicScan.lblYear(1).Caption = CStr(KXYear)
    
    frmAspicScan.lblSeries.Caption = CStr(KINDX)
    frmAspicScan.lblModelType.Caption = ModelType
    
    frmAspicScan.Show

    AspicFlag = True
'    If KFYear < StartYear Or KXYear > EndYear Then
'        MsgBox "Invalid Year Range Specification for Data Grid", vbInformation, "Visual Report Designer"
'        Close #5
'        AspicFlag = False
'    Else
'        AspicFlag = True
'    End If
'
'    If KFYear < MinYear Then
'        MinYear = KFYear
'    End If
'    If KXYear > MaxYear Then
'        MaxYear = KXYear
'    End If
    
    Exit Sub
    
ScanAspicErr:
    MsgBox "Error Scanning Aspic Results", vbInformation, "Visual Report Designer"
    Close #5
End Sub
Public Sub ScanAspicResults()
    Dim Buffer As String
    Dim Token As String
    Dim TX As String
    Dim TY As String
    Dim G As MSFlexGrid
    Dim I As Integer
    Dim J As Integer
    Dim N As Integer
    Dim K As Integer
    Dim NI As Integer
    Dim NK As Integer
    Dim NX As Integer
    Dim KR As Integer
    Dim DT As String
    Dim data() As String
    Dim DoMissing As Boolean 'flag to replace missing values with user selection
    Dim isMissing() As Boolean 'vector of years that are missing
    Dim xNA As String 'user-selected missing value
    Dim S As String
    
    If AspicFlag = False Then
        Unload frmAspicScan
        Exit Sub
    End If

    KR = KRow
    
    On Error GoTo ScanAspicError
    
    CaseID = Trim(frmAspicScan.txtCase.Text)
    
    'get options for replacing missing values
    If frmAspicScan.opMissing(0).Value = True Then
        DoMissing = False
        xNA = ""
    Else
        DoMissing = True
        xNA = frmAspicScan.txtMissing.Text
    End If

    'get years and reset forms if needed
    I = Val(frmAspicScan.lblYear(0).Caption)
    J = Val(frmAspicScan.lblYear(1).Caption)
    If StartYear = 0 And EndYear = 0 Then
        frmGeneral.lblStartYr.Caption = CStr(I)
        frmGeneral.lblEndYr.Caption = CStr(J)
        SetParms
    ElseIf I < StartYear Or J > EndYear Then
        If I < StartYear Then frmGeneral.lblStartYr.Caption = CStr(I)
        If J > EndYear Then frmGeneral.lblEndYr.Caption = CStr(J)
        SetParms
    End If
        
    Unload frmAspicScan
    
    Print #3, "ASPIC"
    Print #3, CaseID
    Print #3, FName
    DT = GetFileWriteTime(FName)
    Print #3, DT
    If DoMissing Then
        Print #3, "Substitute any missing values with: " & xNA
    End If
    Print #3, ""
    
    
    Set G = frmData.MSFlexGrid1
    N = KFYear - StartYear + 4
    
    Line Input #5, Buffer
    G.Rows = G.Rows + 1
    K = KRow + 1
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "ASPIC"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Fishing Mortality"
    G.TextMatrix(K, 4) = ""
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    'Note: bug in previous version of scanaspic.exe forgot to print out starting biomass
    KRow = KRow + 1

    Line Input #5, Buffer
    G.Rows = G.Rows + 1
    K = KRow + 1
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "ASPIC"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Starting Biomass"
    G.TextMatrix(K, 4) = ""
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    KRow = KRow + 1
    
    Line Input #5, Buffer
    G.Rows = G.Rows + 1
    K = KRow + 1
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "ASPIC"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Average Biomass"
    G.TextMatrix(K, 4) = ""
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
       
    KRow = KRow + 1
    
    Line Input #5, Buffer
    G.Rows = G.Rows + 2
    K = KRow + 1
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "ASPIC"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Yield"
    G.TextMatrix(K, 4) = "Observed"
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    
    K = KRow + 2
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "ASPIC"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Yield"
    G.TextMatrix(K, 4) = "Predicted"
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    KRow = KRow + 2
    
    Line Input #5, Buffer
    G.Rows = G.Rows + 1
    K = KRow + 1
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "ASPIC"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Surplus Production"
    G.TextMatrix(K, 4) = ""
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
       
    KRow = KRow + 1
    
    Line Input #5, Buffer
    G.Rows = G.Rows + 1
    K = KRow + 1
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "ASPIC"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "F-Ratio"
    G.TextMatrix(K, 4) = "F / F-Msy"
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
       
    KRow = KRow + 1


    Line Input #5, Buffer
    G.Rows = G.Rows + 1
    K = KRow + 1
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "ASPIC"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "B-Ratio"
    G.TextMatrix(K, 4) = "B / B-Msy"
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
       
    KRow = KRow + 1

    For NI = 1 To KINDX
        Line Input #5, Buffer
        Line Input #5, Buffer
        NK = InStr(Buffer, ",")
        If NK > 1 Then
            Mid(Buffer, NK, 1) = "-"
        End If
        TX = Buffer
        Line Input #5, TY
        If TY = "CC" Or TY = "CE" Then
            NX = 6
        Else
            NX = 5
        End If
        G.Rows = G.Rows + NX
        'prepare for missing value mask
        ReDim isMissing(1 To KYears)
        If NX = 6 Then
            K = KRow + 1
            G.TextMatrix(K, 0) = CStr(K)
            G.TextMatrix(K, 1) = "ASPIC"
            G.TextMatrix(K, 2) = CaseID
            G.TextMatrix(K, 3) = TX + Space(1) + TY
            If TY = "CC" Then
                Token = "Observed Effort"
            Else
                Token = "Observed CPUE"
            End If
            G.TextMatrix(K, 4) = Token
            Line Input #5, Buffer
            Token = GetFirstToken(Buffer)
            For J = 1 To KYears
                'check for negative value (indicates missing) and construct mask
                If DoMissing Then
                    If Val(Token) < 0 Then
                        Token = xNA
                        isMissing(J) = True
                    Else
                        isMissing(J) = False
                    End If
                End If
                G.TextMatrix(K, J + N) = Token
                Token = GetNextToken(Buffer)
            Next J
            K = KRow + 2
            G.TextMatrix(K, 0) = CStr(K)
            G.TextMatrix(K, 1) = "ASPIC"
            G.TextMatrix(K, 2) = CaseID
            G.TextMatrix(K, 3) = TX + Space(1) + TY
            G.TextMatrix(K, 4) = "Estimated CPUE"
            Line Input #5, Buffer
            Token = GetFirstToken(Buffer)
            For J = 1 To KYears
                G.TextMatrix(K, J + N) = Token
                Token = GetNextToken(Buffer)
            Next J
            K = KRow + 3
            G.TextMatrix(K, 0) = CStr(K)
            G.TextMatrix(K, 1) = "ASPIC"
            G.TextMatrix(K, 2) = CaseID
            G.TextMatrix(K, 3) = TX + Space(1) + TY
            G.TextMatrix(K, 4) = "Estimated F"
            Line Input #5, Buffer
            Token = GetFirstToken(Buffer)
            For J = 1 To KYears
                G.TextMatrix(K, J + N) = Token
                Token = GetNextToken(Buffer)
            Next J
            K = KRow + 4
            G.TextMatrix(K, 0) = CStr(K)
            G.TextMatrix(K, 1) = "ASPIC"
            G.TextMatrix(K, 2) = CaseID
            G.TextMatrix(K, 3) = TX + Space(1) + TY
            G.TextMatrix(K, 4) = "Observed Yield"
            Line Input #5, Buffer
            Token = GetFirstToken(Buffer)
            For J = 1 To KYears
                G.TextMatrix(K, J + N) = Token
                Token = GetNextToken(Buffer)
            Next J
            K = KRow + 5
            G.TextMatrix(K, 0) = CStr(K)
            G.TextMatrix(K, 1) = "ASPIC"
            G.TextMatrix(K, 2) = CaseID
            G.TextMatrix(K, 3) = TX + Space(1) + TY
            G.TextMatrix(K, 4) = "Estimated Yield"
            Line Input #5, Buffer
            Token = GetFirstToken(Buffer)
            For J = 1 To KYears
                G.TextMatrix(K, J + N) = Token
                Token = GetNextToken(Buffer)
            Next J
            K = KRow + 6
            G.TextMatrix(K, 0) = CStr(K)
            G.TextMatrix(K, 1) = "ASPIC"
            G.TextMatrix(K, 2) = CaseID
            G.TextMatrix(K, 3) = TX + Space(1) + TY
            G.TextMatrix(K, 4) = "Residual"
            Line Input #5, Buffer
            Token = GetFirstToken(Buffer)
            For J = 1 To KYears
                If DoMissing And isMissing(J) Then Token = xNA
                G.TextMatrix(K, J + N) = Token
                Token = GetNextToken(Buffer)
            Next J
        Else
            ReDim data(1 To 5, 1 To KYears)
            
            Line Input #5, Buffer
            Token = GetFirstToken(Buffer)
            For J = 1 To KYears
                data(1, J) = Token
                Token = GetNextToken(Buffer)
            Next J
            
            Line Input #5, Buffer
            Token = GetFirstToken(Buffer)
            For J = 1 To KYears
                data(2, J) = Token
                Token = GetNextToken(Buffer)
            Next J
            
            Line Input #5, Buffer
            Token = GetFirstToken(Buffer)
            For J = 1 To KYears
                'check for negative value (indicates missing) and construct mask
                If DoMissing Then
                    If Val(Token) < 0 Then
                        isMissing(J) = True
                    Else
                        isMissing(J) = False
                    End If
                End If
                data(3, J) = Token
                Token = GetNextToken(Buffer)
            Next J
            
            Line Input #5, Buffer
            Token = GetFirstToken(Buffer)
            For J = 1 To KYears
                data(4, J) = Token
                Token = GetNextToken(Buffer)
            Next J
            
            Line Input #5, Buffer
            Token = GetFirstToken(Buffer)
            For J = 1 To KYears
                data(5, J) = Token
                Token = GetNextToken(Buffer)
            Next J
            
            K = KRow + 1
            G.TextMatrix(K, 0) = CStr(K)
            G.TextMatrix(K, 1) = "ASPIC"
            G.TextMatrix(K, 2) = CaseID
            G.TextMatrix(K, 3) = TX + Space(1) + TY
            G.TextMatrix(K, 4) = "Observed Effort"
            For J = 1 To KYears
                S = data(1, J)
                If DoMissing And isMissing(J) Then S = xNA
                G.TextMatrix(K, J + N) = S
            Next J
            K = KRow + 2
            G.TextMatrix(K, 0) = CStr(K)
            G.TextMatrix(K, 1) = "ASPIC"
            G.TextMatrix(K, 2) = CaseID
            G.TextMatrix(K, 3) = TX + Space(1) + TY
            G.TextMatrix(K, 4) = "Estimated Effort"
            For J = 1 To KYears
                S = data(2, J)
                If DoMissing And isMissing(J) Then S = xNA
                G.TextMatrix(K, J + N) = S
            Next J
            K = KRow + 3
            G.TextMatrix(K, 0) = CStr(K)
            G.TextMatrix(K, 1) = "ASPIC"
            G.TextMatrix(K, 2) = CaseID
            G.TextMatrix(K, 3) = TX + Space(1) + TY
            If Mid(TY, 1, 1) = "B" Then
                S = "Observed Biomass"
            Else
                S = "Observed Index"
            End If
            G.TextMatrix(K, 4) = S
            For J = 1 To KYears
                S = data(3, J)
                If DoMissing And isMissing(J) Then S = xNA
                G.TextMatrix(K, J + N) = S
            Next J
            K = KRow + 4
            G.TextMatrix(K, 0) = CStr(K)
            G.TextMatrix(K, 1) = "ASPIC"
            G.TextMatrix(K, 2) = CaseID
            G.TextMatrix(K, 3) = TX + Space(1) + TY
            If Mid(TY, 1, 1) = "B" Then
                S = "Model Biomass"
            Else
                S = "Model Index"
            End If
            G.TextMatrix(K, 4) = S
            For J = 1 To KYears
                G.TextMatrix(K, J + N) = data(4, J)
            Next J
            K = KRow + 5
            G.TextMatrix(K, 0) = CStr(K)
            G.TextMatrix(K, 1) = "ASPIC"
            G.TextMatrix(K, 2) = CaseID
            G.TextMatrix(K, 3) = TX + Space(1) + TY
            G.TextMatrix(K, 4) = "Residual"
            For J = 1 To KYears
                S = data(5, J)
                If DoMissing And isMissing(J) Then S = xNA
                G.TextMatrix(K, J + N) = S
            Next J
        End If
        KRow = KRow + NX
    Next NI

    frmCompose.Show
    frmData.ZOrder
    
    K = G.Rows
    
    MsgBox "Aspic Data Scan Completed" + vbCrLf + CStr(KRow - KR) + " Rows Added to Data Grid", vbInformation, "Visual Report Designer"
    Close #5
    DataFlag = True
    
    frmMain.mnuDataMissingReplace.Enabled = True
    
    'automatically save data into file
    WriteGridData

    Exit Sub
    
ScanAspicError:
    Close #5
    MsgBox "Error Scanning Aspic Model Results" + vbCrLf + Err.Description, vbInformation, "Visual Report Designer"
    
End Sub
Private Function CheckASAPFiles() As Boolean
    Dim N As Long
    
    On Error GoTo CheckErr4
    
    FNOUT = FName
    N = InStrRev(FNOUT, ".")
    Mid(FNOUT, N, 4) = ".rep"
    If CompareFileWriteTimes(FName, FNOUT) Then
        MsgBox "ASAP Output Files Are Not Current", vbInformation, "Visual Report Designer"
        CheckASAPFiles = False
        Exit Function
    End If
    Mid(FNOUT, N, 4) = ".tmp"
    If Dir(FNOUT) <> "" Then
        Kill FNOUT
    End If
    CheckASAPFiles = True
    Exit Function
    
CheckErr4:
    MsgBox "Error Checking ASAP Files" + vbCrLf + Err.Description, vbExclamation, "Visual Report Designer"

End Function
Private Sub ScanAsapGeneralData()
    Dim N As Long
    Dim Buffer As String
    Dim Token As String
    
    On Error GoTo ScanasapErr
    
    FNOUT = FName
    N = InStrRev(FNOUT, ".")
    Mid(FNOUT, N, 4) = ".tmp"
    
    If Dir(FNOUT) = "" Then
        MsgBox "ASAP Data Scan Failed" + vbCrLf + "ASAP Output Files May Be Incomplete", vbExclamation, "Visual Report Designer"
        Exit Sub
    End If
    
    Open FNOUT For Input As #5
    
    Line Input #5, Buffer
    
    frmASAPScan.lblFile.Caption = FNOUT
    frmASAPScan.txtCase.Text = Buffer
    
    Line Input #5, Buffer
    
    Token = GetFirstToken(Buffer)
    KYears = Val(Token)
    Token = GetNextToken(Buffer)
    KFYear = Val(Token)
    Token = GetNextToken(Buffer)
    KAges = Val(Token)
    Token = GetNextToken(Buffer)
    KFish = Val(Token)
    Token = GetNextToken(Buffer)
    KINDX = Val(Token)

    frmASAPScan.lblAges.Caption = CStr(KAges)
    frmASAPScan.lblIndex.Caption = CStr(KINDX)
    frmASAPScan.lblYear(0).Caption = CStr(KFYear)
    KXYear = KFYear + KYears - 1
    frmASAPScan.lblYear(1).Caption = CStr(KXYear)
    frmASAPScan.lblFish.Caption = CStr(KFish)
    frmASAPScan.Show

    ASAPFlag = True
    
    Exit Sub
    
ScanasapErr:
    MsgBox "Error Scanning ASAP Results", vbInformation, "Visual Report Designer"
    Close #5
End Sub
Public Sub ScanASAPResults()
    Dim Buffer As String
    Dim Token As String
    Dim G As MSFlexGrid
    Dim I As Integer
    Dim J As Integer
    Dim N As Integer
    Dim K As Integer
    Dim KR As Integer
    Dim NF As Integer
    Dim NSurvUsed As Integer 'the number of surveys used in estimation
    Dim Survey() As ASAPSurvey
    Dim SurveyData() As Double
    Dim SurveyResid() As Double
    Dim DT As String
    Dim TempMat() As String 'temporary matrix
    Dim Tag1 As String
    Dim Tag2 As String
        
    If ASAPFlag = False Then
        Unload frmASAPScan
        Exit Sub
    End If

    KR = KRow
    
  '  On Error GoTo ScanAsapError
    
    CaseID = Trim(frmASAPScan.txtCase.Text)
    
    'get years and reset forms if needed
    I = Val(frmASAPScan.lblYear(0).Caption)
    J = Val(frmASAPScan.lblYear(1).Caption)
    If StartYear = 0 And EndYear = 0 Then
        frmGeneral.lblStartYr.Caption = CStr(I)
        frmGeneral.lblEndYr.Caption = CStr(J)
        SetParms
    ElseIf I < StartYear Or J > EndYear Then
        If I < StartYear Then frmGeneral.lblStartYr.Caption = CStr(I)
        If J > EndYear Then frmGeneral.lblEndYr.Caption = CStr(J)
        SetParms
    End If
    
    Unload frmASAPScan
    
    Print #3, "ASAP"
    Print #3, CaseID
    Print #3, FName
    DT = GetFileWriteTime(FName)
    Print #3, DT
    Print #3, ""
    
    Set G = frmData.MSFlexGrid1
    N = KFYear - StartYear + 4
    
    'Natural Mortality
    Line Input #5, Buffer
    G.Rows = G.Rows + KAges
    For I = 1 To KAges
        K = KRow + I
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "ASAP"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = "Natural Mortality"
        G.TextMatrix(K, 4) = "Age " + CStr(I)
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For J = 1 To KYears
            G.TextMatrix(K, J + N) = Token
            Token = GetNextToken(Buffer)
        Next J
    Next I
    KRow = KRow + KAges
    
    'Maturity
    Line Input #5, Buffer
    G.Rows = G.Rows + KAges
    For I = 1 To KAges
        K = KRow + I
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "ASAP"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = "Maturity"
        G.TextMatrix(K, 4) = "Age " + CStr(I)
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For J = 1 To KYears
            G.TextMatrix(K, J + N) = Token
            Token = GetNextToken(Buffer)
        Next J
    Next I
    KRow = KRow + KAges
    
    'Catch Weight at Age
    Line Input #5, Buffer
    G.Rows = G.Rows + KAges
    For I = 1 To KAges
        K = KRow + I
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "ASAP"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = "Catch Weight"
        G.TextMatrix(K, 4) = "Age " + CStr(I)
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For J = 1 To KYears
            G.TextMatrix(K, J + N) = Token
            Token = GetNextToken(Buffer)
        Next J
    Next I
    KRow = KRow + KAges
    
    'Spawning Stock Weight at Age
    Line Input #5, Buffer
    G.Rows = G.Rows + KAges
    For I = 1 To KAges
        K = KRow + I
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "ASAP"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = "Spawning Stock Weight"
        G.TextMatrix(K, 4) = "Age " + CStr(I)
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For J = 1 To KYears
            G.TextMatrix(K, J + N) = Token
            Token = GetNextToken(Buffer)
        Next J
    Next I
    KRow = KRow + KAges
    
    'Jan-1 Weight at Age
    Line Input #5, Buffer
    G.Rows = G.Rows + KAges
    For I = 1 To KAges
        K = KRow + I
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "ASAP"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = "Jan-1 Weight"
        G.TextMatrix(K, 4) = "Age " + CStr(I)
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For J = 1 To KYears
            G.TextMatrix(K, J + N) = Token
            Token = GetNextToken(Buffer)
        Next J
    Next I
    KRow = KRow + KAges
    
    'Catch Weight by Fleet
    Line Input #5, Buffer
    G.Rows = G.Rows + KFish
    For I = 1 To KFish
        K = KRow + I
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "ASAP"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = "Catch Weight"
        G.TextMatrix(K, 4) = "Fishery # " + CStr(I)
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For J = 1 To KYears
            G.TextMatrix(K, J + N) = Token
            Token = GetNextToken(Buffer)
        Next J
    Next I
    KRow = KRow + KFish
    
    'Discard Weight by Fleet
    Line Input #5, Buffer
    G.Rows = G.Rows + KFish
    For I = 1 To KFish
        K = KRow + I
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "ASAP"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = "Discard Weight"
        G.TextMatrix(K, 4) = "Fishery # " + CStr(I)
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For J = 1 To KYears
            G.TextMatrix(K, J + N) = Token
            Token = GetNextToken(Buffer)
        Next J
    Next I
    KRow = KRow + KFish
    
    'Catch - Input & Effective Sample Size
    'order is: fleet 1 input sample size, fleet 1 effective sample size, fleet 2 input sample size, fleet 2 effective sample size
    Line Input #5, Buffer
    For NF = 1 To KFish
        G.Rows = G.Rows + 2
        
        K = KRow + 1
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "ASAP"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = "Catch Input Sample Size"
        G.TextMatrix(K, 4) = "Fishery # " + CStr(NF)
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For J = 1 To KYears
            G.TextMatrix(K, J + N) = Token
            Token = GetNextToken(Buffer)
        Next J
        KRow = KRow + 1
        
        K = KRow + 1
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "ASAP"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = "Catch Effective Sample Size"
        G.TextMatrix(K, 4) = "Fishery # " + CStr(NF)
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For J = 1 To KYears
            G.TextMatrix(K, J + N) = Token
            Token = GetNextToken(Buffer)
        Next J
        KRow = KRow + 1
        
    Next NF
    
    'Discard - Input & Effective Sample Size
    'order is: fleet 1 input sample size, fleet 1 effective sample size, fleet 2 input sample size, fleet 2 effective sample size
    Line Input #5, Buffer
    For NF = 1 To KFish
        G.Rows = G.Rows + 2
        
        K = KRow + 1
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "ASAP"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = "Discard Input Sample Size"
        G.TextMatrix(K, 4) = "Fishery # " + CStr(NF)
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For J = 1 To KYears
            G.TextMatrix(K, J + N) = Token
            Token = GetNextToken(Buffer)
        Next J
        KRow = KRow + 1
        
        K = KRow + 1
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "ASAP"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = "Discard Effective Sample Size"
        G.TextMatrix(K, 4) = "Fishery # " + CStr(NF)
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For J = 1 To KYears
            G.TextMatrix(K, J + N) = Token
            Token = GetNextToken(Buffer)
        Next J
        KRow = KRow + 1
        
    Next NF
    
    'Index Specification
    ReDim TempMat(1 To 6, 1 To KINDX)

    Line Input #5, Buffer
    
    'tag
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KINDX
        TempMat(1, J) = Token
        Token = GetNextToken(Buffer)
    Next J

    'units
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KINDX
        TempMat(2, J) = Token
        Token = GetNextToken(Buffer)
    Next J

    'month
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KINDX
        TempMat(3, J) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    'skip link and selectivity pattern
    Line Input #5, Buffer
    Line Input #5, Buffer
    
    'start age
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KINDX
        TempMat(4, J) = Token
        Token = GetNextToken(Buffer)
    Next J

    'end age
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KINDX
        TempMat(5, J) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    'used in estimation
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KINDX
        TempMat(6, J) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    'count how many surveys were used
    NSurvUsed = 0
    For I = 1 To KINDX
        If Val(TempMat(6, I)) = 1 Then NSurvUsed = NSurvUsed + 1
    Next I
    
    'populate the Index Spec variable
    ReDim Survey(1 To NSurvUsed)
    For J = 1 To KINDX
        If Val(TempMat(6, J)) = 1 Then
            Survey(J).Tag = TempMat(1, J)
            Survey(J).Units = Val(TempMat(2, J))
            Survey(J).Month = Val(TempMat(3, J))
            Survey(J).StartAge = Val(TempMat(4, J))
            Survey(J).EndAge = Val(TempMat(5, J))
        End If
    Next J
    
    'Index Observed, Predicted, Standardized Residual
    'Order: Index 1 Ob, Index 1 Prd, Index 1 SR, Index 2 Ob, Index 2 Prd, Index 2 SR, etc.
    Line Input #5, Buffer
    For NF = 1 To NSurvUsed
        'construct tags
        Tag1 = Trim(Survey(NF).Tag) + " "
        Tag1 = Tag1 + CStr(Survey(NF).StartAge) + "-" + CStr(Survey(NF).EndAge) + Space(2)
        If Survey(NF).Units = 1 Then
            Tag1 = Tag1 + "B "
        Else
            Tag1 = Tag1 + "N "
        End If
        Tag1 = Tag1 + MonthString(Survey(NF).Month)
        
        G.Rows = G.Rows + 3
        
        K = KRow + 1
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "ASAP"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = Tag1
        G.TextMatrix(K, 4) = "Obs. Index"
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For J = 1 To KYears
            G.TextMatrix(K, J + N) = Token
            Token = GetNextToken(Buffer)
        Next J
        KRow = KRow + 1
        
        K = KRow + 1
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "ASAP"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = Tag1
        G.TextMatrix(K, 4) = "Prd. Index"
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For J = 1 To KYears
            G.TextMatrix(K, J + N) = Token
            Token = GetNextToken(Buffer)
        Next J
        KRow = KRow + 1
        
        K = KRow + 1
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "ASAP"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = Tag1
        G.TextMatrix(K, 4) = "Index Resid."
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For J = 1 To KYears
            G.TextMatrix(K, J + N) = Token
            Token = GetNextToken(Buffer)
        Next J
        KRow = KRow + 1
        
    Next NF
    
    'Index Input & Effective Sample Size
    Line Input #5, Buffer
    For NF = 1 To NSurvUsed
        'construct tags
        Tag1 = Trim(Survey(NF).Tag) + " "
        Tag1 = Tag1 + CStr(Survey(NF).StartAge) + "-" + CStr(Survey(NF).EndAge) + Space(2)
        If Survey(NF).Units = 1 Then
            Tag1 = Tag1 + "B "
        Else
            Tag1 = Tag1 + "N "
        End If
        Tag1 = Tag1 + MonthString(Survey(NF).Month)
        
        G.Rows = G.Rows + 2
        
        K = KRow + 1
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "ASAP"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = Tag1
        G.TextMatrix(K, 4) = "Input Sample Size"
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For J = 1 To KYears
            G.TextMatrix(K, J + N) = Token
            Token = GetNextToken(Buffer)
        Next J
        KRow = KRow + 1
        
        K = KRow + 1
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "ASAP"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = Tag1
        G.TextMatrix(K, 4) = "Effective Sample Size"
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For J = 1 To KYears
            G.TextMatrix(K, J + N) = Token
            Token = GetNextToken(Buffer)
        Next J
        KRow = KRow + 1
        
    Next NF
    
    'Average F - Unweighted, Nweighted, Bweighted
    Line Input #5, Buffer
    'unweighted
    G.Rows = G.Rows + 1
    K = KRow + 1
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "ASAP"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Average F"
    G.TextMatrix(K, 4) = "Unwtd."
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    KRow = KRow + 1
    'n-weighted
    G.Rows = G.Rows + 1
    K = KRow + 1
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "ASAP"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Average F"
    G.TextMatrix(K, 4) = "N Wtd."
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    KRow = KRow + 1
    'b-weighted
    G.Rows = G.Rows + 1
    K = KRow + 1
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "ASAP"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Average F"
    G.TextMatrix(K, 4) = "B Wtd."
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    KRow = KRow + 1
    
    'Stock Numbers
    Line Input #5, Buffer
    G.Rows = G.Rows + KAges
    For I = 1 To KAges
        K = KRow + I
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "ASAP"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = "Stock Numbers"
        G.TextMatrix(K, 4) = "Age " + CStr(I)
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For J = 1 To KYears
            G.TextMatrix(K, J + N) = Token
            Token = GetNextToken(Buffer)
        Next J
    Next I
    KRow = KRow + KAges
    
    'Spawning Stock Biomass
    Line Input #5, Buffer
    G.Rows = G.Rows + 1
    K = KRow + 1
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "ASAP"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Spawning Stock Biomass"
    G.TextMatrix(K, 4) = ""
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    KRow = KRow + 1
    
    'Observed & Predicted Recruitment
    Line Input #5, Buffer
    G.Rows = G.Rows + 2
    K = KRow + 1
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "ASAP"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Recruits"
    G.TextMatrix(K, 4) = "Observed"
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J

    K = KRow + 2
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "ASAP"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Recruits"
    G.TextMatrix(K, 4) = "Predicted"
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J

    KRow = KRow + 2
    
    frmCompose.Show
    frmData.ZOrder
    
    K = G.Rows
    
    MsgBox "ASAP Data Scan Completed" + vbCrLf + CStr(KRow - KR) + " Rows Added to Data Grid", vbInformation, "Visual Report Designer"
    Close #5
    
    DataFlag = True
    
    frmMain.mnuDataMissingReplace.Enabled = True
    
    'automatically save data into file
    WriteGridData
    
    Exit Sub
    
ScanAsapError:
    MsgBox "Error Scanning ASAP Results" + vbCrLf + Err.Description, vbInformation, "Visual Report Designer"
    Close #5

End Sub
Private Function MonthString(K As Integer)

Select Case K
    Case 1
        MonthString = "JAN"
    Case 2
        MonthString = "FEB"
    Case 3
        MonthString = "MAR"
    Case 4
        MonthString = "APR"
    Case 5
        MonthString = "MAY"
    Case 6
        MonthString = "JUN"
    Case 7
        MonthString = "JUL"
    Case 8
        MonthString = "AUG"
    Case 9
        MonthString = "SEP"
    Case 10
        MonthString = "OCT"
    Case 11
        MonthString = "NOV"
    Case 12
        MonthString = "DEC"
    Case -1
        MonthString = "AVG"
End Select
End Function
Private Sub ReadGridData()
    Dim FN As String
    Dim I As Integer
    Dim J As Integer
    Dim K As Integer
    Dim N As Integer
    Dim RS As Integer
    Dim S As String
    Dim T As String
    Dim DT As String
    Dim G As MSFlexGrid
    
    On Error GoTo ReadErr
    
    Unload frmCompose
    
    FN = FNLOG
    N = InStrRev(FN, ".")
    Mid(FN, N, 4) = ".csv"
    
    If Dir(FN) = "" Then
        Exit Sub
    End If
    
    RS = MsgBox("Do You Wish to Restore Existing Grid Data?", vbQuestion + vbYesNo, "Visual Report Designer")
    If RS = vbNo Then
        Exit Sub
    End If
    
    Set G = frmData.MSFlexGrid1
    G.Clear
    Open FN For Input As #2
    Line Input #2, S
    JToken = 0
    
    ' Skip 5 tokens
    For I = 1 To 6
        T = GetToken(S)
    Next I
    StartYear = Val(T)
    frmGeneral.lblStartYr.Caption = CStr(StartYear)
    
    Do
        EndYear = Val(T)
        T = GetToken(S)
    Loop While T <> ""
    NYears = EndYear - StartYear + 1
    frmGeneral.lblEndYr.Caption = CStr(EndYear)
    
    SetParms
    
    N = NYears + 4
    
    I = 0
    Do While Not EOF(2)
        G.Rows = G.Rows + 1
        I = I + 1
        Line Input #2, S
        JToken = 0
        For J = 0 To N
            T = GetToken(S)
            G.TextMatrix(I, J) = T
        Next J
    Loop
    
    KRow = G.Rows - 1
    
    K = G.Rows
    
    Close #2
    
    DT = GetFileWriteTime(FN)
    
    Print #3, ""
    Print #3, "Data Restored from CSV File"
    Print #3, FN
    Print #3, DT
    Print #3, ""
    
    MaxYear = StartYear
    MinYear = EndYear
    
    DataFlag = True
    
    frmMain.mnuDataMissingReplace.Enabled = True
    
    ReDim BList(0 To 14, 1 To 1)
    ReDim ReportInfo(0 To 14)
    MaxLines = 1
    
    MsgBox "Data Grid Has Been Restored", vbInformation, "Visual Report Designer"
    
    Exit Sub
ReadErr:
    Close #2
    MsgBox "Error Reading Grid Data" + vbCrLf + Err.Description, vbExclamation, "Visual Report Designer"
    
End Sub
Private Function GetToken(S As String) As String
    Dim I As Integer
    Dim N As Integer
    Dim K As Integer
    Dim T As String
    
    
    If JToken = K Then
        GetToken = ""
    End If
    
    K = Len(S)
    N = InStr(JToken + 1, S, ",")
    If N > 0 Then
        T = Mid(S, JToken + 1, N - JToken - 1)
        JToken = N
    Else
        T = Mid(S, JToken + 1, K - JToken)
        JToken = K
    End If
    GetToken = T
    
End Function
Private Sub ResizeGrid()
    Dim I As Integer
    Dim J As Integer
    Dim K As Integer
    Dim N As Integer
    Dim N1 As Integer
    Dim N2 As Integer
    Dim NY As Integer
    Dim G As MSFlexGrid
    Dim X() As Variant
    
    
    Set G = frmData.MSFlexGrid1
    
    N = G.Cols - 5
    K = G.Rows - 1
    
    
    ReDim X(1 To K, 1 To N)
    
    N1 = Val(G.TextMatrix(0, 5))
    N2 = Val(G.TextMatrix(0, G.Cols - 1))
    
    For I = 1 To K
        For J = 1 To N
            X(I, J) = G.TextMatrix(I, J + 4)
            G.TextMatrix(I, J + 4) = ""
        Next J
    Next I
    
    G.Cols = NYears + 5
    
    For I = 1 To NYears
        G.ColWidth(I + 4) = 1600
        G.TextMatrix(0, I + 4) = StartYear + I - 1
    Next I
    
    If N1 >= StartYear Then
        If N2 < EndYear Then
            For I = 1 To K
                For J = N1 To N2
                    G.TextMatrix(I, J - StartYear + 5) = X(I, J - N1 + 1)
                Next J
            Next I
        Else
            For I = 1 To K
                For J = N1 To EndYear
                    G.TextMatrix(I, J - StartYear + 5) = X(I, J - N1 + 1)
                Next J
            Next I
        End If
    Else
        If N2 < EndYear Then
            For I = 1 To K
                For J = StartYear To N2
                    G.TextMatrix(I, J - StartYear + 5) = X(I, J - N1 + 1)
                Next J
            Next I
        Else
            For I = 1 To K
                For J = StartYear To EndYear
                    G.TextMatrix(I, J - StartYear + 5) = X(I, J - N1 + 1)
                Next J
            Next I
        End If
    End If
    
End Sub
Public Sub SearchCollection()
    Dim S As String
    Dim S2 As String
    Dim T As String
    Dim T1 As String
    Dim T2 As String
    Dim I As Integer
    Dim J As Integer
    Dim K As Integer
    Dim KX As Integer
    Dim KWild As Integer
    Dim N As Integer
    Dim flag As Boolean
    Dim G As MSFlexGrid
    Dim Istart As Integer
    Dim iStop As Integer
    
    Set G = frmData.MSFlexGrid1
    
    T = UCase(frmSearch.txtSearch.Text)
    N = G.Rows - 1
    
    'determine whether search should continue if reach the end of the grid
    If JList = 1 Then
        Istart = 1
        iStop = N
    Else
        Istart = JList
        iStop = JList - 1
    End If
    
    KWild = InStr(T, "*") 'wildcard search character
    
    flag = False
    
    If KWild > 0 Then 'if wildcard character used
        K = Len(T)
        T1 = Left(T, KWild - 1)
        T2 = Right(T, K - KWild)
        For I = Istart To N
            S = UCase(GetDataGridDesc(I))
            K = InStr(S, T1)
            If K > 0 Then
                S2 = Mid(S, K + Len(T1))
                KX = InStr(S2, T2)
                If KX > 0 Then
                    flag = True
                    Exit For
                End If
            End If
        Next I
        If Not flag And iStop < N Then 'loop through rest of grid
            For I = 1 To iStop
                S = UCase(GetDataGridDesc(I))
                K = InStr(S, T1)
                If K > 0 Then
                    S2 = Mid(S, K + Len(T1))
                    KX = InStr(S2, T2)
                    If KX > 0 Then
                        flag = True
                        Exit For
                    End If
                End If
                Next I
        End If
    Else
        For I = Istart To N
            S = UCase(GetDataGridDesc(I))
            K = InStr(S, T)
            If K > 0 Then
                flag = True
                Exit For
            End If
        Next I
        If Not flag And iStop < N Then
            For I = 1 To iStop 'loop through rest of grid
                S = UCase(GetDataGridDesc(I))
                K = InStr(S, T)
                If K > 0 Then
                    flag = True
                    Exit For
                End If
            Next I
        End If
    End If
    
    If flag Then
        If I <= N Then
            G.Row = I
            G.Col = 0
            G.ColSel = G.Cols - 1
            If G.RowIsVisible(I) = False Then G.TopRow = I 'show row if out of view
            If G.ColIsVisible(0) = False Then G.LeftCol = 0 'show first column if out of view
            frmSearch.lblLine.Caption = CStr(I)
            S = GetDataGridDesc(I) & vbCrLf
            S = S & GetDataGridData(I)
            frmSearch.lblData.Caption = S
            If I = G.Rows - 1 Then
                JList = 1
            Else
                JList = I + 1
            End If
        End If
    Else
        MsgBox "Text Not Found", vbInformation, "Visual Report Designer"
        'JList = 1
        JList = I
    End If
    
    
End Sub
Public Sub ShowChart()
    Dim I As Integer
    Dim J As Integer
    Dim K As Integer
    Dim N As Integer
    Dim G As MSFlexGrid
    Dim T As String
    Dim TY As String
    Dim Tag As String
    Dim XVEC() As Integer
    Dim YVEC() As String
    Dim YMAX As Double
    Dim YMIN As Double
    Dim X As Double
    Dim S As String

    
    Set G = frmData.MSFlexGrid1
    
    K = G.Row
    
    If K = 0 Then
        MsgBox "Please Select an Item to Plot", vbInformation, "Visual Report Designer"
        Exit Sub
    End If
        
    ReDim XVEC(1 To NYears)
    ReDim YVEC(1 To NYears)
    
    Tag = G.TextMatrix(K, 1) + ":  " + G.TextMatrix(K, 2)
    TY = G.TextMatrix(K, 3) + ", " + G.TextMatrix(K, 4)
    
    'get data
    For J = 1 To NYears
        XVEC(J) = StartYear + J - 1
        YVEC(J) = G.TextMatrix(K, J + 4)
    Next J
    
    'assign min and max
    N = 0
    For I = 1 To NYears
        N = N + 1
        If YVEC(I) <> NAVal Then
            YMIN = Val(YVEC(I))
            YMAX = Val(YVEC(I))
            Exit For
        End If
    Next I
    If N = NYears Then
        'no non-missing data
        YMIN = 0
        YMAX = 1
    Else
        For I = N + 1 To NYears
            If YVEC(I) <> NAVal Then
                X = Val(YVEC(I))
                If X < YMIN Then
                    YMIN = X
                End If
                If X > YMAX Then
                    YMAX = X
                End If
            End If
        Next I
    End If
    

    
    With frmChart.Chart1
        .ClearData ClearDataFlag_AllData
        .Gallery = Gallery_Lines
        .OpenData COD_XValues, 1, NYears
        .OpenData COD_Values, 1, NYears
            For I = 1 To NYears
                S = YVEC(I)
                If S = NAVal Then
                    .Value(0, I - 1) = .Hidden
                Else
                    .Value(0, I - 1) = Val(S)
                End If
                .XValue(0, I - 1) = XVEC(I)
            Next I
        .CloseData COD_Values
        .CloseData COD_XValues
        .SerLegBox = False
        .Titles(0).Text = Tag
        .Titles(0).Font.Size = 10
        .Titles(0).Font.Bold = True
        .Titles(1).Text = TY
        .Titles(1).Font.Size = 10
        .AxisX.Title.Text = "Year"
        '.AxisY.Title.Text = TY
        If YMAX > 10000 Then
            .AxisY.LabelsFormat.Decimals = 0
        ElseIf YMAX > 100 Then
            .AxisY.LabelsFormat.Decimals = 2
        ElseIf YMAX < 10 Then
            .AxisY.LabelsFormat.Decimals = 4
        End If
        .AxisY.max = YMAX * 1.1
        If YMIN > 0# Then
            .AxisY.min = YMIN * 0.75
        Else
            .AxisY.min = YMIN * 1.25
        End If
        .AxisX.min = XVEC(1)
        .AxisX.max = XVEC(NYears)
        .Grid = ChartGrid_Horz Or ChartGrid_Vert
        .ToolBar = True
    End With
    
    frmChart.Show


End Sub

Private Function CheckAIMFiles() As Boolean
    Dim N As Long
    Dim flag As Boolean
    
    On Error GoTo CheckErr5
    
    FNOUT = FName
    N = InStrRev(FNOUT, ".")
    Mid(FNOUT, N, 4) = ".out"
    flag = CompareFileWriteTimes(FName, FNOUT)
    If flag Then
        MsgBox "AIM Output Files Are Not Current", vbInformation, "Visual Report Designer"
        CheckAIMFiles = False
        Exit Function
    End If
    Mid(FNOUT, N, 4) = ".tmp"
    If Dir(FNOUT) <> "" Then
        Kill FNOUT
    End If
    CheckAIMFiles = True
    Exit Function
    
CheckErr5:
    MsgBox "Error Checking AIM Files" + vbCrLf + Err.Description, vbExclamation, "Visual Report Designer"

End Function
Private Sub ScanAIMGeneralData()
    Dim N As Long
    Dim Buffer As String
    Dim Token As String
    
    On Error GoTo ScanAIMErr
    
    FNOUT = FName
    N = InStrRev(FNOUT, ".")
    Mid(FNOUT, N, 4) = ".tmp"
    
    If Dir(FNOUT) = "" Then
        MsgBox "AIM Data Scan Failed" + vbCrLf + "AIM Output Files May Be Incomplete", vbExclamation, "Visual Report Designer"
        Exit Sub
    End If
    
    Open FNOUT For Input As #5
    
    Line Input #5, Buffer
    
    frmAimScan.lblFile.Caption = FNOUT
    frmAimScan.txtCase.Text = Buffer
    
    Line Input #5, Buffer
    
    Token = GetFirstToken(Buffer)
    KYears = Val(Token)
    Token = GetNextToken(Buffer)
    KFYear = Val(Token)
    Token = GetNextToken(Buffer)
    KINDX = Val(Token)


    frmAimScan.lblYear(0).Caption = CStr(KFYear)
    KXYear = KFYear + KYears - 1
    frmAimScan.lblYear(1).Caption = CStr(KXYear)
    frmAimScan.lblSeries.Caption = CStr(KINDX)
    
    frmAimScan.Show

    AIMFlag = True

    Exit Sub
    
ScanAIMErr:
    MsgBox "Error Scanning AIM Results", vbInformation, "Visual Report Designer"
    Close #5
End Sub
Public Sub ScanAimResults()
    Dim Buffer As String
    Dim Token As String
    Dim G As MSFlexGrid
    Dim I As Integer
    Dim J As Integer
    Dim N As Integer
    Dim K As Integer
    Dim KR As Integer
    Dim DT As String
    Dim X() As Double
    Dim Tag() As String
    
    On Error GoTo ScanAIMError
    
    If AIMFlag = False Then
        Unload frmAimScan
        Exit Sub
    End If

    KR = KRow
    
    CaseID = Trim(frmAimScan.txtCase.Text)
    
    'get years and reset forms if needed
    I = Val(frmAimScan.lblYear(0).Caption)
    J = Val(frmAimScan.lblYear(1).Caption)
    If StartYear = 0 And EndYear = 0 Then
        frmGeneral.lblStartYr.Caption = CStr(I)
        frmGeneral.lblEndYr.Caption = CStr(J)
        SetParms
    ElseIf I < StartYear Or J > EndYear Then
        If I < StartYear Then frmGeneral.lblStartYr.Caption = CStr(I)
        If J > EndYear Then frmGeneral.lblEndYr.Caption = CStr(J)
        SetParms
    End If
    
    Unload frmAimScan
    
    Print #3, "AIM"
    Print #3, CaseID
    Print #3, FName
    DT = GetFileWriteTime(FName)
    Print #3, DT
    Print #3, ""
    
    
    
    Set G = frmData.MSFlexGrid1
    G.ColWidth(4) = 2500
    N = KFYear - StartYear + 4
    
    Line Input #5, Buffer
    G.Rows = G.Rows + 1
    K = KRow + 1
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "AIM"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Catch"
    G.TextMatrix(K, 4) = "Input"
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    KRow = KRow + 1
    
    ReDim Tag(1 To KINDX)
    ReDim X(1 To KINDX, 1 To KYears)
    
    'get survey names
    Line Input #5, Buffer
    Line Input #5, Buffer
    K = 1
    For J = 1 To KINDX
        Tag(J) = Trim(Mid(Buffer, K, 15))
        K = K + 15
    Next J
    
    'get observed index data
    For J = 1 To KYears
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For I = 1 To KINDX
            X(I, J) = Val(Token)
            Token = GetNextToken(Buffer)
        Next I
    Next J
    
    For I = 1 To KINDX
        G.Rows = G.Rows + 1
        K = KRow + 1
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "AIM"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = "Observed Index"
        G.TextMatrix(K, 4) = Tag(I)
        For J = 1 To KYears
            G.TextMatrix(K, J + N) = Format(X(I, J), "0.000000E+00")
        Next J
        KRow = KRow + 1
    Next I
    
    ReDim X(1 To KINDX, 1 To KYears)
    
    Line Input #5, Buffer
    Line Input #5, Buffer
    
    'get replacement ratio
    For J = 1 To KYears
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For I = 1 To KINDX
            X(I, J) = Val(Token)
            Token = GetNextToken(Buffer)
        Next I
    Next J
    
    For I = 1 To KINDX
        G.Rows = G.Rows + 1
        K = KRow + 1
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "AIM"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = "Replacement Ratio"
        G.TextMatrix(K, 4) = Tag(I)
        For J = 1 To KYears
            G.TextMatrix(K, J + N) = Format(X(I, J), "0.000000E+00")
        Next J
        KRow = KRow + 1
    Next I
    
    
    ReDim X(1 To KINDX, 1 To KYears)
    
    Line Input #5, Buffer
    Line Input #5, Buffer
    
    'get relative F
    For J = 1 To KYears
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For I = 1 To KINDX
            X(I, J) = Val(Token)
            Token = GetNextToken(Buffer)
        Next I
    Next J
    
    For I = 1 To KINDX
        G.Rows = G.Rows + 1
        K = KRow + 1
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "AIM"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = "Relative F"
        G.TextMatrix(K, 4) = Tag(I)
        For J = 1 To KYears
            G.TextMatrix(K, J + N) = Format(X(I, J), "0.000000E+00")
        Next J
        KRow = KRow + 1
    Next I
    
    frmCompose.Show
    frmData.ZOrder
    
    K = G.Rows
    
    
    MsgBox "AIM Data Scan Completed" + vbCrLf + CStr(KRow - KR) + " Rows Added to Data Grid", vbInformation, "Visual Report Designer"
    Close #5
    DataFlag = True
    
    frmMain.mnuDataMissingReplace.Enabled = True
    
    'automatically save data into file
    WriteGridData

    Exit Sub
ScanAIMError:
    Close #5
    MsgBox "Error Scanning AIM Model Results" + vbCrLf + Err.Description, vbInformation, "Visual Report Designer"

End Sub
Private Function CheckAgeProFiles() As Boolean
    Dim N As Long
    Dim flag As Boolean
    
    On Error GoTo CheckErr6
    
    FNOUT = FName
    N = InStrRev(FNOUT, ".")
    FNOUT = Left(FNOUT, N) + "out"
    If CompareFileWriteTimes(FName, FNOUT) Then
        MsgBox "AgePro Output Files Are Not Current", vbInformation, "Visual Report Designer"
        CheckAgeProFiles = False
        Exit Function
    End If
    Mid(FNOUT, N, 4) = ".tmp"
    If Dir(FNOUT) <> "" Then
        Kill FNOUT
    End If
    CheckAgeProFiles = True
    Exit Function
    
CheckErr6:
    MsgBox "Error Checking AgePro Files" + vbCrLf + Err.Description, vbExclamation, "Visual Report Designer"

End Function
Private Sub ScanAgeProGeneralData()
    Dim N As Long
    Dim Buffer As String
    Dim Token As String
    
    On Error GoTo ScanAgeProErr
    
    FNOUT = FName
    N = InStrRev(FNOUT, ".")
    FNOUT = Left(FNOUT, N) + "tmp"
    
    If Dir(FNOUT) = "" Then
        MsgBox "AgePro Data Scan Failed" + vbCrLf + "AgePro Output Files May Be Incomplete", vbExclamation, "Visual Report Designer"
        Exit Sub
    End If
    
    Open FNOUT For Input As #5
    
    Line Input #5, Buffer
    
    frmAgeProScan.lblFile.Caption = FNOUT
    frmAgeProScan.txtCase.Text = Buffer
    
    Line Input #5, Buffer
    
    Token = GetFirstToken(Buffer)
    KYears = Val(Token)
    Token = GetNextToken(Buffer)
    KFYear = Val(Token)
    Token = GetNextToken(Buffer)


    frmAgeProScan.lblYear(0).Caption = CStr(KFYear)
    KXYear = KFYear + KYears - 1
    frmAgeProScan.lblYear(1).Caption = CStr(KXYear)
    
    frmAgeProScan.Show

    AgeProFlag = True
'    If KFYear < StartYear Or KXYear > EndYear Then
'        MsgBox "Invalid Year Range Specification for Data Grid", vbInformation, "Visual Report Designer"
'        Close #5
'        AgeProFlag = False
'    Else
'        AgeProFlag = True
'    End If
'
'    If KFYear < MinYear Then
'        MinYear = KFYear
'    End If
'    If KXYear > MaxYear Then
'        MaxYear = KXYear
'    End If
    
    Exit Sub
    
ScanAgeProErr:
    MsgBox "Error Scanning AgePro Results", vbInformation, "Visual Report Designer"
    Close #5

End Sub
Public Sub ScanAgeProResults()
    Dim Buffer As String
    Dim Token As String
    Dim G As MSFlexGrid
    Dim I As Integer
    Dim J As Integer
    Dim N As Integer
    Dim K As Integer
    Dim KR As Integer
    Dim NR As Integer
    Dim DT As String
    Dim PERC(9) As String
    
    On Error GoTo ScanAgeProError
    
    If AgeProFlag = False Then
        Unload frmAgeProScan
        Exit Sub
    End If

    KR = KRow
    
    PERC(1) = CStr(1)
    PERC(2) = CStr(5)
    PERC(3) = CStr(10)
    PERC(4) = CStr(25)
    PERC(5) = CStr(50)
    PERC(6) = CStr(75)
    PERC(7) = CStr(90)
    PERC(8) = CStr(95)
    PERC(9) = CStr(99)
        
    CaseID = Trim(frmAgeProScan.txtCase.Text)
    
    'get years and reset forms if needed
    I = Val(frmAgeProScan.lblYear(0).Caption)
    J = Val(frmAgeProScan.lblYear(1).Caption)
    If StartYear = 0 And EndYear = 0 Then
        frmGeneral.lblStartYr.Caption = CStr(I)
        frmGeneral.lblEndYr.Caption = CStr(J)
        SetParms
    ElseIf I < StartYear Or J > EndYear Then
        If I < StartYear Then frmGeneral.lblStartYr.Caption = CStr(I)
        If J > EndYear Then frmGeneral.lblEndYr.Caption = CStr(J)
        SetParms
    End If
    
    Unload frmAgeProScan
    
    Print #3, "AGEPRO"
    Print #3, CaseID
    Print #3, FName
    DT = GetFileWriteTime(FName)
    Print #3, DT
    Print #3, ""
    
    Set G = frmData.MSFlexGrid1
    G.ColWidth(4) = 2500
    N = KFYear - StartYear + 4
    
    Line Input #5, Buffer
    G.Rows = G.Rows + 1
    K = KRow + 1
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "AGEPRO"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Average Spawning Stock Biomass"
    G.TextMatrix(K, 4) = ""
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    KRow = KRow + 1
    
    Line Input #5, Buffer
    For I = 1 To 9
        G.Rows = G.Rows + 1
        K = KRow + 1
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "AGEPRO"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = "Percentile Avg SSB"
        G.TextMatrix(K, 4) = PERC(I) + " %"
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For J = 1 To KYears
            G.TextMatrix(K, J + N) = Token
            Token = GetNextToken(Buffer)
        Next J
    
        KRow = KRow + 1
    Next I
    
    Line Input #5, Buffer
    NR = InStr(Buffer, "=")
    G.Rows = G.Rows + 1
    K = KRow + 1
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "AGEPRO"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Probability Avg SSB Exceeds " + Mid(Buffer, NR + 1, 10)
    G.TextMatrix(K, 4) = ""
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    KRow = KRow + 1


    Line Input #5, Buffer
    G.Rows = G.Rows + 1
    K = KRow + 1
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "AGEPRO"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Mean Biomass"
    G.TextMatrix(K, 4) = ""
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    KRow = KRow + 1
    
    
    Line Input #5, Buffer
    For I = 1 To 9
        G.Rows = G.Rows + 1
        K = KRow + 1
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "AGEPRO"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = "Percentile Mean Biomass"
        G.TextMatrix(K, 4) = PERC(I) + " %"
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For J = 1 To KYears
            G.TextMatrix(K, J + N) = Token
            Token = GetNextToken(Buffer)
        Next J
    
        KRow = KRow + 1
    Next I
    
    Line Input #5, Buffer
    NR = InStr(Buffer, "=")
    G.Rows = G.Rows + 1
    K = KRow + 1
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "AGEPRO"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Probability Mean Biomass Exceeds " + Mid(Buffer, NR + 1, 10)
    G.TextMatrix(K, 4) = ""
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    KRow = KRow + 1

    
    Line Input #5, Buffer
    G.Rows = G.Rows + 1
    K = KRow + 1
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "AGEPRO"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "F Weighted by Biomass"
    G.TextMatrix(K, 4) = ""
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    KRow = KRow + 1
    
    
    Line Input #5, Buffer
    For I = 1 To 9
        G.Rows = G.Rows + 1
        K = KRow + 1
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "AGEPRO"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = "Percentile F Wtd by Biomass"
        G.TextMatrix(K, 4) = PERC(I) + " %"
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For J = 1 To KYears
            G.TextMatrix(K, J + N) = Token
            Token = GetNextToken(Buffer)
        Next J
    
        KRow = KRow + 1
    Next I
    
    Line Input #5, Buffer
    NR = InStr(Buffer, "=")
    G.Rows = G.Rows + 1
    K = KRow + 1
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "AGEPRO"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Probability F Weighted by Biomass Exceeds " + Mid(Buffer, NR + 1, 10)
    G.TextMatrix(K, 4) = ""
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    KRow = KRow + 1

    
    Line Input #5, Buffer
    G.Rows = G.Rows + 1
    K = KRow + 1
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "AGEPRO"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Total Biomass"
    G.TextMatrix(K, 4) = ""
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    KRow = KRow + 1
    
    
    Line Input #5, Buffer
    For I = 1 To 9
        G.Rows = G.Rows + 1
        K = KRow + 1
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "AGEPRO"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = "Percentile Total Biomass"
        G.TextMatrix(K, 4) = PERC(I) + " %"
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For J = 1 To KYears
            G.TextMatrix(K, J + N) = Token
            Token = GetNextToken(Buffer)
        Next J
    
        KRow = KRow + 1
    Next I
    
    Line Input #5, Buffer
    NR = InStr(Buffer, "=")
    G.Rows = G.Rows + 1
    K = KRow + 1
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "AGEPRO"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Probability Total Biomass Exceeds " + Mid(Buffer, NR + 1, 10)
    G.TextMatrix(K, 4) = ""
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    KRow = KRow + 1

    Line Input #5, Buffer
    G.Rows = G.Rows + 1
    K = KRow + 1
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "AGEPRO"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Recruits"
    G.TextMatrix(K, 4) = ""
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    KRow = KRow + 1
    
    
    Line Input #5, Buffer
    For I = 1 To 9
        G.Rows = G.Rows + 1
        K = KRow + 1
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "AGEPRO"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = "Percentile Recruits"
        G.TextMatrix(K, 4) = PERC(I) + " %"
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For J = 1 To KYears
            G.TextMatrix(K, J + N) = Token
            Token = GetNextToken(Buffer)
        Next J
    
        KRow = KRow + 1
    Next I
    
    
    Line Input #5, Buffer
    G.Rows = G.Rows + 1
    K = KRow + 1
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "AGEPRO"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Landings"
    G.TextMatrix(K, 4) = ""
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    KRow = KRow + 1
    
    
    Line Input #5, Buffer
    For I = 1 To 9
        G.Rows = G.Rows + 1
        K = KRow + 1
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "AGEPRO"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = "Percentile Landings"
        G.TextMatrix(K, 4) = PERC(I) + " %"
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For J = 1 To KYears
            G.TextMatrix(K, J + N) = Token
            Token = GetNextToken(Buffer)
        Next J
    
        KRow = KRow + 1
    Next I

    
    Line Input #5, Buffer
    NR = InStr(Buffer, "=")
    G.Rows = G.Rows + 1
    K = KRow + 1
    G.TextMatrix(K, 0) = CStr(K)
    G.TextMatrix(K, 1) = "AGEPRO"
    G.TextMatrix(K, 2) = CaseID
    G.TextMatrix(K, 3) = "Probability Fully Recruited F Exceeds " + Mid(Buffer, NR + 1, 10)
    G.TextMatrix(K, 4) = ""
    Line Input #5, Buffer
    Token = GetFirstToken(Buffer)
    For J = 1 To KYears
        G.TextMatrix(K, J + N) = Token
        Token = GetNextToken(Buffer)
    Next J
    
    KRow = KRow + 1
    
    frmCompose.Show
    frmData.ZOrder
    
    K = G.Rows
    
    MsgBox "AgePro Data Scan Completed" + vbCrLf + CStr(KRow - KR) + " Rows Added to Data Grid", vbInformation, "Visual Report Designer"
    Close #5
    DataFlag = True
    
    frmMain.mnuDataMissingReplace.Enabled = True
    
    'automatically save data into file
    WriteGridData

    Exit Sub
    
ScanAgeProError:
    Close #5
    MsgBox "Error Scanning AgePro Model Results" + vbCrLf + Err.Description, vbInformation, "Visual Report Designer"

End Sub
Public Sub SelectItem(iRow As Integer, Mode As Integer)
'iRow is the row in the data grid to add to report, Mode is whether user is adding
' to report in batch mode or individually.
    Dim G As MSFlexGrid
    Dim I As Integer
    Dim J As Integer
    Dim K As Integer
    Dim K1 As Integer
    Dim K2 As Integer
    Dim JJ As Integer
    Dim T As String
    
    'hide options for batch edit mode
    frmSpecLayout.opFYear(2).Visible = False
    frmSpecLayout.opXYear(2).Visible = False
    frmSpecLayout.chkEditDisplay.Visible = False
    frmSpecLayout.chkEditZero.Visible = False

    'put the data grid key on the form for use later
    frmSpecLayout.lblKey.Caption = CStr(iRow)
    'get item description
    T = GetDataGridDesc(iRow)
        
    If Mode = 1 Then 'edit each item individually
        'reset batch mode number of lines to 1
        frmSpecLayout.lblNLines.Caption = "1"
        frmSpecLayout.lblDescrip.Caption = T
        'show individual label box
        frmSpecLayout.frmDesc.Visible = True
        'set row tag options
        frmSpecLayout.txtTag.Text = T
        For I = 0 To 3
            frmSpecLayout.chkTag(I).Value = 1
        Next I
        frmSpecLayout.opTag(0).Value = True
    ElseIf Mode = 2 Then 'edit in batch mode
        Set G = frmData.MSFlexGrid1
        'number of lines to add
        frmSpecLayout.lblNLines.Caption = CStr(G.RowSel - G.Row + 1)
        'hide individual label box
        frmSpecLayout.frmDesc.Visible = False
        'set row tag example
        frmSpecLayout.lblLine.Caption = "Example:"
        frmSpecLayout.lblTag.Caption = T
        For I = 0 To 3
            frmSpecLayout.chkTag(I).Value = 1
        Next I
        frmSpecLayout.opTag(0).Value = True
        'disable custom tag option
        frmSpecLayout.opTag(1).Enabled = False
        'first and last year labels
        T = frmSpecLayout.opFYear(0).Caption
        frmSpecLayout.opFYear(0).Caption = Left(T, Len(T) - 1)
        frmSpecLayout.lblFYear.Visible = False
        T = frmSpecLayout.opXYear(0).Caption
        frmSpecLayout.opXYear(0).Caption = Left(T, Len(T) - 1)
        frmSpecLayout.lblXYear.Visible = False
    End If
    
    frmSpecLayout.cmdUpdate.Caption = "ADD"
    
    'get report and line
    'select a form
    K = frmCompose.SSTab1.Tab
    
    Set G = frmCompose.grdReport(K)
    frmSpecLayout.txtLine.Text = CStr(G.Rows)
    
    frmSpecLayout.chkZero.Value = 1
    
    frmSpecLayout.txtMark(0).Text = "0"
    frmSpecLayout.txtMark(1).Text = "0"
     
    'find start and end years of data
    K1 = GetDataMinYear(iRow)
    K2 = GetDataMaxYear(iRow)
    ' set start and end years
    frmSpecLayout.lblFYear.Caption = CStr(K1)
    frmSpecLayout.lblXYear.Caption = CStr(K2)
    frmSpecLayout.cboFYear.ListIndex = K1 - StartYear
    frmSpecLayout.cboXYear.ListIndex = K2 - StartYear
    frmSpecLayout.opFYear(0).Value = True
    frmSpecLayout.opXYear(0).Value = True
    
    LayoutFlag = True
    frmSpecLayout.Show
End Sub
Public Sub InitPalette(N As Integer)


With frmSpecLayout
    Set .cboPalette.ImageList = .ImageList1
    .cboPalette.ComboItems.Clear
    Select Case N
        Case 0 'single cut point
            .cboPalette.ComboItems.Add , , "    Plain_Plus/Minus", 1, 1
            .cboPalette.ComboItems.Add , , "    Green_Plus/Red_Minus", 2, 2
            .cboPalette.ComboItems.Add , , "    Red_Plus/Green_Minus", 3, 3
            .cboPalette.ComboItems.Add , , "    Green/Red", 4, 4
            .cboPalette.ComboItems.Add , , "    Red/Green", 5, 5
            Set .cboPalette.SelectedItem = .cboPalette.ComboItems(1)
        Case 1 'dual cut point
            .cboPalette.ComboItems.Add , , "    Plain_Plus_to_Minus", 6, 6
            .cboPalette.ComboItems.Add , , "    Green_Plus_to_Red_Minus", 7, 7
            .cboPalette.ComboItems.Add , , "    Red_Plus_to_Green_Minus", 8, 8
            .cboPalette.ComboItems.Add , , "    Green_to_Red", 9, 9
            .cboPalette.ComboItems.Add , , "    Red_to_Green", 10, 10
            Set .cboPalette.SelectedItem = .cboPalette.ComboItems(1)
        Case 2 'quintiles
            .cboPalette.ComboItems.Add , , "    Red_to_Black", 11, 11
            .cboPalette.ComboItems.Add , , "    Black_to_Red", 12, 12
            .cboPalette.ComboItems.Add , , "    Red_to_Blue", 13, 13
            .cboPalette.ComboItems.Add , , "    Blue_to_Red", 14, 14
            .cboPalette.ComboItems.Add , , "    White_to_Black", 15, 15
            .cboPalette.ComboItems.Add , , "    Black_to_White", 16, 16
            Set .cboPalette.SelectedItem = .cboPalette.ComboItems(1)
    End Select
    
End With

End Sub
Public Function ColorCheck() As Boolean
'checks to make sure that user hasn't selected a color palette which contrasts
'with current color palettes in use on the report (e.g., make sure green values
'are consistently the high values and reds are consistently the low values)
Dim rtnval As Integer
Dim iRpt As Integer
Dim iBin As Integer
Dim G As MSFlexGrid
Dim N As Integer
Dim M As Integer
Dim I As Integer
Dim iPalette As String
Dim iType As Integer
Dim HighC As String
Dim S As String
Dim OKflag As Boolean

iRpt = frmSpecLayout.cboReport.ListIndex
Set G = frmCompose.grdReport(iRpt)
N = G.Rows - 1

If N = 0 Then
    ColorCheck = True
    Exit Function
End If

'get color palette selection and bin type
iPalette = Trim(frmSpecLayout.cboPalette.Text)
iBin = frmSpecLayout.cboBins.ListIndex + 2

'if palette is the simple plus/minus style, it doesn't conflict with other palettes, so exit
If InStr(iPalette, "Plain_") > 0 Then
    ColorCheck = True
    Exit Function
End If


'now compare with rest of form

OKflag = True 'initialize flag
'first start with quintiles, and compare only to other quintiles on form
If iBin = 4 Then
    For I = 1 To N
        S = BList(iRpt, I).Palette
        'continue only if item is a quintile and doesn't have the same color palette
        If BList(iRpt, I).Type = iBin And S <> iPalette And S <> "" Then
            Select Case iPalette
                Case "Black_to_Red", "Red_to_Black":
                    If S = "Black_to_Red" Or S = "Red_to_Black" Then OKflag = False
                Case "Red_to_Blue", "Blue_to_Red":
                    If S = "Red_to_Blue" Or S = "Blue_to_Red" Then OKflag = False
                Case "White_to_Black", "Black_to_White":
                    If S = "White_to_Black" Or S = "Black_to_White" Then OKflag = False
            End Select
            If Not OKflag Then Exit For
        End If
    Next I
Else 'now check cut points
    'first get the color palette's high color value (truncated to first 3 letters)
    HighC = Left(iPalette, 3)
    'now do comparison
    For I = 1 To N
        S = BList(iRpt, I).Palette
        S = Left(S, 3)
        'skip if item is a quinitle, if it has the same high color or is a simple plus/minus or the line is blank
        If BList(iRpt, I).Type < 4 And S <> HighC And S <> "Pla" And S <> "" Then
            OKflag = False
            Exit For
        End If
    Next I
End If

If OKflag = True Then
    ColorCheck = True
    Exit Function
End If

rtnval = MsgBox("Color Palette Selection Conflicts with Color Palettes" & vbCrLf & _
                "Currently in Use on Report " & CStr(iRpt + 1) & "." & vbCrLf & vbCrLf & _
                "Do You Wish to Continue?", vbInformation + vbOKCancel, "Visual Report Designer")
If rtnval = vbCancel Then
    ColorCheck = False
Else
    ColorCheck = True
End If

End Function

Public Sub ClearFormCompose(Index As Integer)
Dim I As Integer
Dim G As MSFlexGrid

Set G = frmCompose.grdReport(Index)
G.Rows = 1
frmCompose.cboStartYr(Index).ListIndex = 0
frmCompose.cboEndYr(Index).ListIndex = frmCompose.cboEndYr(Index).ListCount - 1
frmCompose.txtTitle(Index).Text = ""
ReportInfo(Index).Title = ""
ReportInfo(Index).NLines = 0

End Sub
Private Sub ReadLayoutFile()
Dim I As Integer
Dim J As Integer
Dim K As Integer
Dim iRpt As Integer
Dim N As Integer
Dim M As Integer
Dim N1 As Integer
Dim N2 As Integer
Dim K1 As Integer
Dim K2 As Integer
Dim KL As Integer
Dim KT As Integer
Dim HF As Double
Dim LF As Double
Dim KX As Integer
Dim KTag As String
Dim Kzero As Boolean
Dim KPalette As String
Dim S As String
Dim T As String
Dim G As MSFlexGrid
Dim VerNum As String
Dim iForms As Integer
    
On Error GoTo ReadErr

If Not DataFlag Then Exit Sub

FNTXT = frmGeneral.lblFile.Caption
N = InStrRev(FNTXT, ".")
Mid(FNTXT, N, 4) = ".txt"
If Dir(FNTXT) = "" Then
    Exit Sub
End If

Open FNTXT For Input As #1

'set form initialization variable to suppress messages
InitFlag = True

DoEvents
'clear frmCompose in case user had a previous session
For I = 0 To 14
    ClearFormCompose (I)
Next I

iForms = 0

'figure out which version this is
'Versions previous to 1.5 did not contain a version identifying line.
'Versions 1.5 and later contain, as the first line:
'    $VisualReport Version X.XX
'Versions where the layout file changed: 1.5, 1.6
'Use VerNum = "0" for versions previous to 1.5
VerNum = "0"
Line Input #1, S
If InStr(S, "VisualReport Version") > 0 Then
    'get version number
    T = GetFirstToken(S)
    T = GetNextToken(S)
    VerNum = GetNextToken(S)
    Line Input #1, S
End If

Do While Not EOF(1)
    K = InStr(S, "$END")
    If K > 0 Then
        Exit Do
    End If
    N = InStr(S, "$Report")
    If N > 0 Then
        'get report number
        S = Trim(S)
        S = Right(S, Len(S) - 7)
        iForms = Val(S)
        'report title
        Line Input #1, S
        ReportInfo(iForms - 1).Title = Trim(S)
        frmCompose.txtTitle(iForms - 1).Text = Trim(S)
        'report years
        Line Input #1, S
        T = GetFirstToken(S)
        M = Val(T) - StartYear
        frmCompose.cboStartYr(iForms - 1).ListIndex = M
        T = GetNextToken(S)
        M = Val(T) - StartYear
        frmCompose.cboEndYr(iForms - 1).ListIndex = M
        
        'if newer versions (1.5 and later), get legend labels, option for
        ' printing dispersion statistic (ver 1.6 and later), and cut point location
        Line Input #1, S
        If Val(VerNum) > 0 Then
            For I = 1 To 3
                N1 = InStr(S, "Legend_")
                If N1 = 0 Then Exit For
                N2 = Mid(S, 8, 1) 'which legend the labels go with (based on 3 color schemes)
                S = Mid(S, 11)
                For J = 1 To 5
                    T = "Label_" & CStr(J)
                    N1 = InStr(S, T)
                    ReportLegend(iForms - 1, N2, J - 1) = Trim(Left(S, N1 - 1))
                    S = Mid(S, N1 + 9)
                Next J
                ReportLegend(iForms - 1, N2, 5) = Trim(S)
                Line Input #1, S
            Next I
            'get dispersion statistic option and number of significant digits
            ' - for ver. 1.6 and later
            If Val(VerNum) >= 1.6 Then
                N1 = InStr(S, "Print_Dispersion")
                If N1 > 0 Then
                    T = GetFirstToken(S)
                    T = GetNextToken(S)
                    DoStat(iForms - 1) = CBool(T)
                    T = GetNextToken(S)
                    StatDig(iForms - 1) = Val(T)
                    Line Input #1, S
                End If
            End If
            N1 = InStr(S, "CutPoint_Location")
            If N1 > 0 Then
                T = GetFirstToken(S)
                CutPtLoc(iForms - 1) = GetNextToken(S)
                Line Input #1, S
            End If
        End If
        Do While Not EOF(1)
            M = InStr(S, "$")
            If M > 0 Then
                Exit Do
            End If
            'if new version, get line title
            If Val(VerNum) > 0 Then
                KTag = Trim(S)
                Line Input #1, S
            End If
            T = GetFirstToken(S)
            K1 = Val(T) 'first year
            T = GetNextToken(S)
            K2 = Val(T) 'last year
            T = GetNextToken(S)
            KL = Val(T) 'report line number
            T = GetNextToken(S)
            KT = Val(T) 'type
            T = GetNextToken(S)
            KPalette = T 'color palette, or high flag for previous versions (1=black to red, 0=red to black)
            T = GetNextToken(S)
            LF = Val(T) 'lower cut
            T = GetNextToken(S)
            HF = Val(T) 'upper cut
            T = GetNextToken(S)
            KX = Val(T) 'key
            T = GetNextToken(S)
            Kzero = T 'zero flag
            Line Input #1, S
            If VerNum = "0" Then 'if old version, get line title
                KTag = Left(S, 35)
            End If
            If KL > MaxLines Then
                MaxLines = KL
                ReDim Preserve BList(0 To 14, 1 To KL)
            End If
            'now assign values to global variables
            ReportInfo(iForms - 1).NLines = KL
            BList(iForms - 1, KL).Tag = KTag 'tag
            BList(iForms - 1, KL).StartYear = K1 'first year
            BList(iForms - 1, KL).EndYear = K2 'last year
            BList(iForms - 1, KL).Type = KT ' display type
            BList(iForms - 1, KL).LowerCut = LF 'lower cut
            BList(iForms - 1, KL).UpperCut = HF 'upper cut
            BList(iForms - 1, KL).ZeroFlag = Kzero 'zero flag
            'set default palette options, if necessary for older versions
            If UCase(KPalette) = "TRUE" Or UCase(KPalette) = "FALSE" Then
                Select Case KT
                    Case 2 'single cut point
                        KPalette = "Plain_Plus/Minus"
                    Case 3 'dual cut points
                        KPalette = "Plain_Plus_to_Minus"
                    Case 4 'quintiles
                        If UCase(KPalette) = "TRUE" Then
                            KPalette = "Black_to_Red"
                            ReportLegend(iForms - 1, 1, 0) = KPalette
                        Else
                            KPalette = "Red_to_Black"
                            ReportLegend(iForms - 1, 1, 0) = KPalette
                        End If
                End Select
            End If
            BList(iForms - 1, KL).Palette = KPalette
            BList(iForms - 1, KL).Key = KX 'key
            If Not EOF(1) Then Line Input #1, S
        Loop
    End If
Loop


'put data on compose form
For I = 1 To iForms
    Set G = frmCompose.grdReport(I - 1)
    N = ReportInfo(I - 1).NLines
    G.Rows = N + 1
    For J = 1 To N
        G.TextMatrix(J, 0) = CStr(J)
        If BList(I - 1, J).Tag <> "" Then PutReportData (I - 1), J
    Next J
Next I

Close #1
    
InitFlag = False

LayoutFlag = True

Exit Sub
ReadErr:
    MsgBox "Error reading report layout file", vbExclamation, "Visual Report Designer"
    Close #1
End Sub
Public Sub PutReportData(Index As Integer, Item As Integer)
Dim N As Integer
Dim G As MSFlexGrid

Set G = frmCompose.grdReport(Index)

'tag
G.TextMatrix(Item, 1) = BList(Index, Item).Tag
'first year
G.TextMatrix(Item, 2) = CStr(BList(Index, Item).StartYear)
'last year
G.TextMatrix(Item, 3) = CStr(BList(Index, Item).EndYear)
' display type
N = BList(Index, Item).Type
Select Case N
    Case 2 'single cut point
        G.TextMatrix(Item, 4) = "Single Cut Pt"
        G.TextMatrix(Item, 5) = BList(Index, Item).Palette
        If BList(Index, Item).Palette = "Plain_Plus/Minus" Then
            G.TextMatrix(Item, 6) = "Plus"
        ElseIf BList(Index, Item).Palette = "Green_Plus/Red_Minus" Then
            G.TextMatrix(Item, 6) = "Green Plus"
        ElseIf BList(Index, Item).Palette = "Red_Plus/Green_Minus" Then
            G.TextMatrix(Item, 6) = "Red Plus"
        ElseIf BList(Index, Item).Palette = "Green/Red" Then
            G.TextMatrix(Item, 6) = "Green"
        ElseIf BList(Index, Item).Palette = "Red/Green" Then
            G.TextMatrix(Item, 6) = "Red"
        End If
        G.TextMatrix(Item, 7) = CStr(BList(Index, Item).LowerCut)
        G.TextMatrix(Item, 8) = ""
    Case 3 'dual cut point
        G.TextMatrix(Item, 4) = "Dual Cut Pts"
        G.TextMatrix(Item, 5) = BList(Index, Item).Palette
        If BList(Index, Item).Palette = "Plain_Plus_to_Minus" Then
            G.TextMatrix(Item, 6) = "Plus"
        ElseIf BList(Index, Item).Palette = "Green_Plus_to_Red_Minus" Then
            G.TextMatrix(Item, 6) = "Green Plus"
        ElseIf BList(Index, Item).Palette = "Red_Plus_to_Green_Minus" Then
            G.TextMatrix(Item, 6) = "Red Plus"
        ElseIf BList(Index, Item).Palette = "Green_to_Red" Then
            G.TextMatrix(Item, 6) = "Green"
        ElseIf BList(Index, Item).Palette = "Red_to_Green" Then
            G.TextMatrix(Item, 6) = "Red"
        End If
        G.TextMatrix(Item, 7) = CStr(BList(Index, Item).LowerCut)
        G.TextMatrix(Item, 8) = CStr(BList(Index, Item).UpperCut)
    Case 4 'quintiles
        G.TextMatrix(Item, 4) = "Quintiles"
        G.TextMatrix(Item, 5) = BList(Index, Item).Palette
        If BList(Index, Item).Palette = "Red_to_Blue" Then
            G.TextMatrix(Item, 6) = "Red"
        ElseIf BList(Index, Item).Palette = "Blue_to_Red" Then
            G.TextMatrix(Item, 6) = "Blue"
        ElseIf BList(Index, Item).Palette = "White_to_Black" Then
            G.TextMatrix(Item, 6) = "White"
        ElseIf BList(Index, Item).Palette = "Black_to_White" Then
            G.TextMatrix(Item, 6) = "Black"
        ElseIf BList(Index, Item).Palette = "Black_to_Red" Then
            G.TextMatrix(Item, 6) = "Black"
        ElseIf BList(Index, Item).Palette = "Red_to_Black" Then
            G.TextMatrix(Item, 6) = "Red"
        End If
        G.TextMatrix(Item, 7) = ""
        G.TextMatrix(Item, 8) = ""
End Select

'zero flag
If BList(Index, Item).ZeroFlag = True Then
    G.TextMatrix(Item, 9) = "Yes"
Else
    G.TextMatrix(Item, 9) = "No"
End If
End Sub

Public Sub ScanSampleGeneralData()
    Dim N As Long
    Dim Buffer As String
    Dim Token As String
    
    On Error GoTo ScanSampleError1

    FNOUT = FName
    N = InStrRev(FNOUT, ".")
    Mid(FNOUT, N, 4) = ".tmp"
    
    If Dir(FNOUT) = "" Then
        MsgBox "Sample Data Scan Failed" + vbCrLf + "Sample Report Files May Be Incomplete", vbExclamation, "Visual Report Designer"
        Exit Sub
    End If
    
    Open FNOUT For Input As #5
    
    Line Input #5, Buffer
    
    frmSampleScan.lblFile.Caption = FNOUT
    frmSampleScan.txtCase.Text = Buffer
    
    Line Input #5, Buffer
    
    Token = GetFirstToken(Buffer)
    KYears = Val(Token)
    Token = GetNextToken(Buffer)
    KAges = Val(Token)
    Token = GetNextToken(Buffer)
    KFAge = Val(Token)
    Token = GetNextToken(Buffer)
    KFYear = Val(Token)

    frmSampleScan.lblYear(0).Caption = CStr(KFYear)
    KXYear = KFYear + KYears - 1
    frmSampleScan.lblYear(1).Caption = CStr(KXYear)
    frmSampleScan.lblAges.Caption = CStr(KAges)
    frmSampleScan.lblNFage.Caption = CStr(KFAge)
    
    AgeFlag = True
'    If KFYear < StartYear Or KXYear > EndYear Then
'        MsgBox "Invalid Year Range Specification for Data Grid", vbInformation, "Visual Report Designer"
'        Close #5
'        AgeFlag = False
'    Else
'        AgeFlag = True
'    End If
'
'    If KFYear < MinYear Then
'        MinYear = KFYear
'    End If
'    If KXYear > MaxYear Then
'        MaxYear = KXYear
'    End If
    
    frmSampleScan.Show
    
    Exit Sub
ScanSampleError1:
    Close #5
    MsgBox "Error Scanning Sample Data", vbExclamation, "Visual Report Designer"

End Sub
Public Sub ScanSampleData()
    Dim Buffer As String
    Dim Token As String
    Dim G As MSFlexGrid
    Dim I As Integer
    Dim J As Integer
    Dim N As Integer
    Dim K As Integer
    Dim KR As Integer
    Dim SurveyData() As Double
    Dim DT As String
    Dim X As Double
    
    On Error GoTo ScanSampleError
    
    If AgeFlag = False Then
        Unload frmSampleScan
        Exit Sub
    End If

    KR = KRow
    
    CaseID = Trim(frmSampleScan.txtCase.Text)
    
    'get years and reset forms if needed
    I = Val(frmSampleScan.lblYear(0).Caption)
    J = Val(frmSampleScan.lblYear(1).Caption)
    If StartYear = 0 And EndYear = 0 Then
        frmGeneral.lblStartYr.Caption = CStr(I)
        frmGeneral.lblEndYr.Caption = CStr(J)
        SetParms
    ElseIf I < StartYear Or J > EndYear Then
        If I < StartYear Then frmGeneral.lblStartYr.Caption = CStr(I)
        If J > EndYear Then frmGeneral.lblEndYr.Caption = CStr(J)
        SetParms
    End If
    
    Unload frmSampleScan
    
    Print #3, "SAMPLE"
    Print #3, CaseID
    Print #3, FName
    DT = GetFileWriteTime(FName)
    Print #3, DT
    Print #3, ""
    
    Set G = frmData.MSFlexGrid1
    N = KFYear - StartYear + 4
    
    Line Input #5, Buffer
    G.Rows = G.Rows + KAges
    For I = 1 To KAges
        K = KRow + I
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "SAMPLES"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = "Avg Length"
        G.TextMatrix(K, 4) = "Age " + CStr(I + KFAge - 1)
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For J = 1 To KYears
            X = Val(Token)
            If X > 0# Then
                G.TextMatrix(K, J + N) = Token
            End If
            Token = GetNextToken(Buffer)
        Next J
    Next I
    KRow = KRow + KAges
    
    
    Line Input #5, Buffer
    G.Rows = G.Rows + KAges
    For I = 1 To KAges
        K = KRow + I
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "SAMPLES"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = "Avg Weight"
        G.TextMatrix(K, 4) = "Age " + CStr(I + KFAge - 1)
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For J = 1 To KYears
            X = Val(Token)
            If X > 0# Then
                G.TextMatrix(K, J + N) = Token
            End If
            Token = GetNextToken(Buffer)
        Next J
    Next I
    KRow = KRow + KAges

    Line Input #5, Buffer
    G.Rows = G.Rows + KAges
    For I = 1 To KAges
        K = KRow + I
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "SAMPLES"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = "Sex Ratio"
        G.TextMatrix(K, 4) = "Age " + CStr(I + KFAge - 1)
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For J = 1 To KYears
            X = Val(Token)
            If X > 0# Then
                G.TextMatrix(K, J + N) = Token
            End If
            Token = GetNextToken(Buffer)
        Next J
    Next I
    KRow = KRow + KAges
    
    Line Input #5, Buffer
    G.Rows = G.Rows + KAges
    For I = 1 To KAges
        K = KRow + I
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "SAMPLES"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = "Avg Length (M)"
        G.TextMatrix(K, 4) = "Age " + CStr(I + KFAge - 1)
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For J = 1 To KYears
            X = Val(Token)
            If X > 0# Then
                G.TextMatrix(K, J + N) = Token
            End If
            Token = GetNextToken(Buffer)
        Next J
    Next I
    KRow = KRow + KAges
    
    Line Input #5, Buffer
    G.Rows = G.Rows + KAges
    For I = 1 To KAges
        K = KRow + I
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "SAMPLES"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = "Avg Length (F)"
        G.TextMatrix(K, 4) = "Age " + CStr(I + KFAge - 1)
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For J = 1 To KYears
            X = Val(Token)
            If X > 0# Then
                G.TextMatrix(K, J + N) = Token
            End If
            Token = GetNextToken(Buffer)
        Next J
    Next I
    KRow = KRow + KAges
    
    Line Input #5, Buffer
    G.Rows = G.Rows + KAges
    For I = 1 To KAges
        K = KRow + I
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "SAMPLES"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = "Avg Weight (M)"
        G.TextMatrix(K, 4) = "Age " + CStr(I + KFAge - 1)
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For J = 1 To KYears
            X = Val(Token)
            If X > 0# Then
                G.TextMatrix(K, J + N) = Token
            End If
            Token = GetNextToken(Buffer)
        Next J
    Next I
    KRow = KRow + KAges
    
    Line Input #5, Buffer
    G.Rows = G.Rows + KAges
    For I = 1 To KAges
        K = KRow + I
        G.TextMatrix(K, 0) = CStr(K)
        G.TextMatrix(K, 1) = "SAMPLES"
        G.TextMatrix(K, 2) = CaseID
        G.TextMatrix(K, 3) = "Avg Weight (F)"
        G.TextMatrix(K, 4) = "Age " + CStr(I + KFAge - 1)
        Line Input #5, Buffer
        Token = GetFirstToken(Buffer)
        For J = 1 To KYears
            X = Val(Token)
            If X > 0# Then
                G.TextMatrix(K, J + N) = Token
            End If
            Token = GetNextToken(Buffer)
        Next J
    Next I
    KRow = KRow + KAges
    
    Close #5
        
    K = G.Rows
    
    'automatically save data into file
    WriteGridData

    Exit Sub
    
ScanSampleError:
    Close #5
    MsgBox "Error Scanning Sample Data", vbExclamation, "Visual Report Designer"

End Sub
Private Sub GetQuintile(N As Long, K1 As Long, K2 As Long, K3 As Long, K4 As Long)
    Dim XN As Double
    Dim XK As Double
    Dim PK As Double
    Dim K As Long

    XN = N
    
    If N < 5 Then
        K1 = 1
        K2 = 1
        K3 = 1
        K4 = 1
        Exit Sub
    End If
    
    
    For K = 1 To N
        XK = K
        PK = (XK - 0.5) / XN
        If PK < 0.2 Then
            K1 = K
        ElseIf PK < 0.4 And PK >= 0.2 Then
            K2 = K
        ElseIf PK < 0.6 And PK >= 0.4 Then
            K3 = K
        ElseIf PK < 0.8 And PK >= 0.6 Then
            K4 = K
        End If
    Next K

End Sub
Public Sub ReadCfg()
    Dim FN As String
    Dim S As String
    Dim I As Integer
    Dim J As Integer
    
    'set defaults if the configuration file is missing or something goes wrong
    RptViewer = "GUI"
    NAVal = ""
    
    FN = App.Path + "\VisualReport.cfg"
    If Dir(FN) <> "" Then
        Open FN For Input As #6
        
        Do While Not EOF(6)
            Line Input #6, S
            If Trim(S) = "[Report Viewer]" Then
                Line Input #6, S
                'check if path is valid
                If S <> "GUI" And S <> "Browser" Then
                    If Dir(S) = "" Then S = "GUI"
                End If
                RptViewer = S
            ElseIf Trim(S) = "[Missing Value Indicator]" Then
                Line Input #6, S
                NAVal = Trim(S)
            End If
        Loop
        
        Close #6
    End If
    
End Sub
Public Sub WriteCfg()
    Dim FN As String
    Dim I As Integer
    Dim J As Integer
    
    FN = App.Path + "\VisualReport.cfg"
    Open FN For Output As #6
    
    Print #6, "[Report Viewer]"
    Print #6, RptViewer
    
    Print #6, "[Missing Value Indicator]"
    Print #6, NAVal
    
    Close #6
End Sub
Public Sub OpenEditSelect()
    Dim S As String

   On Error GoTo OpenErr
    
    frmViewOpt.CommonDialog1.Filter = "EXE Files (*.exe)|*.exe"
    frmViewOpt.CommonDialog1.Flags = &H1004
    frmViewOpt.CommonDialog1.DialogTitle = "Select Program for Viewing Report Files"
    frmViewOpt.CommonDialog1.CancelError = True
    frmViewOpt.CommonDialog1.ShowOpen
    S = frmViewOpt.CommonDialog1.FileName
    If S <> "" Then
        frmViewOpt.lblFile.Caption = S
    End If
    Exit Sub
    
OpenErr:


End Sub
Public Sub WriteHTML(HTMLFile As String, Index As Integer)
Dim I As Integer
Dim J As Integer
Dim K As Integer
Dim S As String
Dim T As String
Dim G As MSFlexGrid
Dim H As MSFlexGrid
Dim CutPtFlag As Boolean
Dim DT As String
Dim rptStartYr As Integer
Dim rptEndYr As Integer
Dim rptYears As Integer

On Error GoTo WriteErr

SaveLayout = False

'check that the last year comes after the first year
rptStartYr = Val(frmCompose.cboStartYr(Index).Text)
rptEndYr = Val(frmCompose.cboEndYr(Index).Text)
rptYears = rptEndYr - rptStartYr + 1
If rptYears < 1 Then
    MsgBox "Invalid Year Range for Report " & CStr(Index + 1), vbInformation, "NFT"
    Exit Sub
End If

'if report exists but grid has no data, delete report
Set G = frmCompose.grdReport(Index)
If G.Rows = 1 Then
    If Dir(HTMLFile) <> "" Then Kill HTMLFile
    Exit Sub
End If

'check that symbols folder exists, for later use when assigning symbols
I = InStrRev(HTMLFile, "\")
S = Left(HTMLFile, I)
S = S & "VRsymbols\"
If Dir(S) = "" Then MkDir S

'check that file isn't already open
'--- TO DO ----

Open HTMLFile For Output As #1

'check for any cut points
CutPtFlag = False
For I = 1 To G.Rows - 1
    S = Trim(G.TextMatrix(I, 7))
    If S <> "" Then
        CutPtFlag = True
        Exit For
    End If
Next I
    
Print #1, "<html>"
Print #1, "<head>"
' meta data tags
Print #1, "<meta name=" & Chr(34) & "author" & Chr(34) & " content=" & Chr(34) & "NFT Visual Report Designer version " & VRVersion & Chr(34) & ">"
DT = DtStr
Print #1, "<meta name=" & Chr(34) & "date" & Chr(34) & " content=" & Chr(34) & DT & Chr(34) & ">"

' write data
Set H = frmData.MSFlexGrid1
Print #1, "<!-- *** BEGIN DATA SECTION ***"
WriteReportData (Index)
Print #1, "*** END DATA SECTION *** -->"
Print #1, ""

' web page title
Print #1, "<title>"
S = frmCompose.txtTitle(Index).Text
If Trim(S) = "" Then S = "NOAA Fisheries Toolbox Visual Report"
Print #1, S
Print #1, "</title>"
Print #1, ""

' style sheet
Print #1, "<style type=" & Chr(34) & "text/css" & Chr(34) & ">"
Print #1, "<!--"
Print #1, ""
Print #1, "body {"
Print #1, "   font-family: " & Chr(34) & "Helvetica" & Chr(34) & ", " & Chr(34) & "Arial" & Chr(34) & ", sans-serif;"
Print #1, "   font-size: 9pt"
Print #1, "}"
Print #1, ""
Print #1, "table { border-collapse: collapse }"
Print #1, ""
Print #1, "td.corner {"
Print #1, "   border-top: 1pt solid white;"
Print #1, "   border-bottom: 1pt solid black;"
Print #1, "   border-left: 1pt solid white;"
Print #1, "   border-right: 1pt solid black"
Print #1, "}"
Print #1, "td.title {"
Print #1, "   white-space: nowrap;"
Print #1, "   border-top: 1pt solid black;"
Print #1, "   border-bottom: 1pt solid black;"
Print #1, "   border-left: 1pt solid black;"
Print #1, "   border-right: 1pt solid black;"
'Print #1, "   padding: .4pt 2pt .4pt 2pt"
Print #1, "   padding: 0pt 2pt 0pt 2pt"
Print #1, "}"
Print #1, "td.icon {"
Print #1, "   white-space: nowrap;"
Print #1, "   border-top: 1pt solid black;"
Print #1, "   border-bottom: 1pt solid black;"
Print #1, "   border-left: 1pt solid black;"
Print #1, "   border-right: 1pt solid black"
'Print #1, "   padding: .4pt .7pt .4pt .7pt"
Print #1, "}"
Print #1, "td.noLR {"
'Print #1, "   border-top: 1pt solid black;"
'Print #1, "   border-bottom: 1pt solid black;"
Print #1, "   border-left: 1pt solid white;"
Print #1, "   border-right: 1pt solid white"
Print #1, "}"
Print #1, "td.threshold {"
Print #1, "   white-space: nowrap;"
Print #1, "   border-top: 1pt solid white;"
Print #1, "   border-bottom: 1pt solid white;"
Print #1, "   border-left: 1pt solid black;"
Print #1, "   border-right: 1pt solid white;"
'Print #1, "   padding: .4pt 1pt .4pt 2pt"
Print #1, "   padding: 0pt 1pt 0pt 2pt"
Print #1, "}"
Print #1, ""
Print #1, "p.head1 {"
Print #1, "   font-family: " & Chr(34) & "Times New Roman" & Chr(34) & ", " & Chr(34) & "Garamond" & Chr(34) & ", serif;"
Print #1, "   font-size: 12pt;"
Print #1, "   font-weight: bold;"
Print #1, "   text-align: center"
Print #1, "}"
Print #1, ""
Print #1, "p.intable {"
Print #1, "   font-family: " & Chr(34) & "Helvetica" & Chr(34) & ", " & Chr(34) & "Arial" & Chr(34) & ", sans-serif;"
Print #1, "   font-size: 9pt"
Print #1, "}"
Print #1, ""
Print #1, "p.tablecenter {"
Print #1, "   font-family: " & Chr(34) & "Helvetica" & Chr(34) & ", " & Chr(34) & "Arial" & Chr(34) & ", sans-serif;"
Print #1, "   font-size: 9pt;"
Print #1, "   text-align: center"
Print #1, "}"
Print #1, ""
Print #1, "p.smbreak {"
Print #1, "   font-family: " & Chr(34) & "Helvetica" & Chr(34) & ", " & Chr(34) & "Arial" & Chr(34) & ", sans-serif;"
Print #1, "   font-size: 9pt;"
Print #1, "   margin-top: 0pt;"
Print #1, "   padding-top: 3pt"
Print #1, "}"
Print #1, ""
Print #1, "-->"
Print #1, "</style>"
Print #1, ""
Print #1, "</head>"
Print #1, "<body>"

' begin writing the actual report

' report title
S = frmCompose.txtTitle(Index).Text
'If Trim(S) = "" Then S = "NOAA Fisheries Toolbox Visual Report"
If Trim(S) <> "" Then
    Print #1, "<! -- title -->"
    Print #1, "<p class=head1>"
    Print #1, S
    Print #1, "</p>"
End If
Print #1, ""
Print #1, ""
' start writing out the table
Print #1, "<!-- begin Table -->"
Print #1, "<table cellspacing=0 cellpadding=0>"
Print #1, ""
Print #1, "<!-- begin table header, years row -->"
Print #1, "<tr>"
Print #1, "<td class=corner>"
Print #1, "<p class=intable>&nbsp;</p>"
Print #1, "</td>"
' report year headings
For I = 1 To rptYears
    Print #1, "<td class=icon>"
    Print #1, "<p class=intable>"
    S = CStr(rptStartYr + I - 1)
    S = Right(S, 2)
    Print #1, S & "</p>"
    Print #1, "</td>"
Next I
' if user has selected to write out the dispersion statistic
If DoStat(Index) = True Then
    Print #1, "<td class=icon>"
    Print #1, "<p class=tablecenter>"
    Print #1, "D</p>"
    Print #1, "</td>"
End If
' if user has selected to write out cut points next to table
If CutPtLoc(Index) = "beside" And CutPtFlag = True Then
    Print #1, "<td class=threshold>"
    Print #1, "<p style=" & Chr(34) & "text-decoration: underline;" & Chr(34) & " class=intable>"
    Print #1, "Threshold</p>"
    Print #1, "</td>"
End If
Print #1, "</tr>"
Print #1, "<!-- end table header, years row -->"
Print #1, ""

' start writing out each line
Set G = frmCompose.grdReport(Index)
Print #1, "<!-- begin data -->"
For I = 1 To G.Rows - 1
    Print #1, "<!-- row" & CStr(I) & " -->"
    Print #1, "<tr>"
    If G.TextMatrix(I, 1) <> "" Then
        ' item description
        Print #1, "<td class=title>"
        Print #1, "<p class=intable>"
        Print #1, G.TextMatrix(I, 1) & "</p>"
        Print #1, "</td>"
        ' print each icon
        BinData Index, I
        ' if user has selected to write out cut points next to table
        If CutPtLoc(Index) = "beside" And CutPtFlag = True Then
            Print #1, "<td class=threshold>"
            Print #1, "<p class=intable>"
            S = Trim(G.TextMatrix(I, 7))
            T = Trim(G.TextMatrix(I, 8))
            If S <> "" And T <> "" Then
                S = "L=" & S & ", " & "U=" & T
            End If
            Print #1, S & "</p>"
            Print #1, "</td>"
        End If
    Else
        'print blank grid cells
        For J = 1 To rptYears + 1
            Print #1, "<td class=noLR>"
            Print #1, "<p class=intable>&nbsp;</p>"
            Print #1, "</td>"
        Next J
        'add extra dispersion statistic cell, if needed
        If DoStat(Index) = True Then
            Print #1, "<td class=noLR>"
            Print #1, "<p class=intable>&nbsp;</p>"
            Print #1, "</td>"
        End If
        'add extra cut point cell, if needed
        If CutPtLoc(Index) = "beside" And CutPtFlag = True Then
            Print #1, "<td class=noLR>"
            Print #1, "<p class=intable>&nbsp;</p>"
            Print #1, "</td>"
        End If
    End If
    
    Print #1, "</tr>"
    Print #1, ""
Next I
Print #1, "<!-- end data -->"
Print #1, ""
Print #1, "</table>"
Print #1, "<br>"
Print #1, "<!-- end Table -->"
Print #1, ""
Print #1, ""

'write out legend
WriteLegend Index

'write out any notes for cut points or dispersion label
Print #1, "<p class=smbreak>"

' if user has selected to write out cut points below table
If CutPtLoc(Index) = "below" And CutPtFlag = True Then
    Print #1, "<!-- begin Cut Point Notes -->"
    
    'Print #1, "Notes:<br>"
    For I = 1 To G.Rows - 1
        S = Trim(G.TextMatrix(I, 7))
        T = Trim(G.TextMatrix(I, 8))
        If S <> "" And T = "" Then
            S = " Threshold Value = " & S & "<br>"
            S = Trim(G.TextMatrix(I, 1)) & ": " & S
        ElseIf S <> "" And T <> "" Then
            S = "Lower Threshold Value = " & S & ", "
            T = "Upper Threshold Value = " & T & "<br>"
            S = S & T
            S = Trim(G.TextMatrix(I, 1)) & ": " & S
        End If
        Print #1, S
    Next I
    Print #1, ""
    'Print #1, "<br>"
    Print #1, "<!-- end Cut Point Notes -->"
    Print #1, ""
    Print #1, ""
End If

' if user has selected to write out the dispersion statistic
If DoStat(Index) = True Then
    Print #1, "D = Measure of Dispersion: Range/Median"
End If

Print #1, "</p>"

'write out ending syntax
Print #1, ""
Print #1, "</body>"
Print #1, "</html>"

Close #1

SaveLayout = True

Exit Sub
WriteErr:
    Close #1
    MsgBox "Error Writing HTML Report File " & HTMLFile & vbCrLf & _
        Err.Description, vbExclamation
End Sub
Public Sub BinData(Index As Integer, iRow As Integer)
Dim I As Integer
Dim J As Integer
Dim K As Integer
Dim N As Integer
Dim S As String
Dim T As String
Dim G As MSFlexGrid
Dim H As MSFlexGrid
Dim X As Double
Dim na As String 'the no-data or missing-data value
Dim DataVec() As String 'vector of raw data to analize
Dim NValDataVec As Integer 'number of non-blank values in the data vector DataVec (equal to the number of non-blank grid cells in data analysis range)
Dim SortVec() As Double 'vector of data to analyse that has zeros=missing filtered out, if needed
Dim NValSortVec As Integer 'number of non-blank, non-missing (and/or non-zero) values in sorting vector SortVec
Dim icon() As String
Dim AnalysisYears As Integer 'the number of years that the data will be analized
                             ' (can be different than the total number of years
                             '  in the data, or, the number of years to display
                             '  in the report)
Dim rptStartYr As Integer 'the display start year in the report
Dim rptEndYr As Integer 'the display end year in the report
Dim rptYears As Integer 'the total number of years to display in the report
Dim myLongVal As Long 'needed to convert Integer to Long for the SortVectorUp and GetQuintile subroutines
'for GetQuintile fcn
Dim N1 As Long
Dim N2 As Long
Dim N3 As Long
Dim N4 As Long

Dim iDispers As Double 'dispersion statistic
Dim DispersString As String 'dispersion statistic re-formatted into a string
Dim iMin As Double 'the minimum of the row of data
Dim iMax As Double 'the maxiumum of the row of data
Dim tmp As Double
Dim iMed As Double 'the median of the the row of data
Dim iType As Integer
Dim iPalette As String

'get the missing-data or no-data value
na = Trim(NAVal)

Set G = frmCompose.grdReport(Index)
Set H = frmData.MSFlexGrid1

'get report years
rptStartYr = Val(frmCompose.cboStartYr(Index).Text)
rptEndYr = Val(frmCompose.cboEndYr(Index).Text)
rptYears = rptEndYr - rptStartYr + 1

'vector to hold icon type
ReDim icon(1 To rptYears)

'number of years in dataset
AnalysisYears = BList(Index, iRow).EndYear - BList(Index, iRow).StartYear + 1
ReDim DataVec(1 To AnalysisYears) 'vector of raw data to analize
ReDim SortVec(1 To AnalysisYears) 'vector of data to analyse that has zeros=missing filtered out, if needed

'Copy data from data collection grid into the DataVec vector. Remember to only
' copy the data within the selected analysis years.
'Remember that "StartYear" is the starting year of the entire dataset; and
' BList(Index, iRow).StartYear is the year to start the data analysis/binning
N = 0
For J = 1 To NYears
    K = StartYear + J - 1
    'only include data within the selected range
    If K >= BList(Index, iRow).StartYear And K <= BList(Index, iRow).EndYear Then
        S = H.TextArray(FGIndex(H, BList(Index, iRow).Key, J + 4))
        N = N + 1
        DataVec(N) = S
    End If
Next J
NValDataVec = N

'Now, exclude missing values ; and exclude zeros from the data set if user has selected to treat zeros as missing
N = 0
For I = 1 To NValDataVec
    S = DataVec(I)
    If Trim(S) <> na Then
        If BList(Index, iRow).ZeroFlag Then 'zero flag = true = treat zero values as missing
            If Abs(Val(S)) > 0# Then
                N = N + 1
                SortVec(N) = Val(S)
            End If
        Else 'treat zeros as zeros
            N = N + 1
            SortVec(N) = Val(S)
        End If
    End If
Next I
NValSortVec = N

' Notes for GetSymbol(iType, iPalette, iBin): For quintiles iBin is 1 to 5, with 1 being high values for normal case, and 5 being high values
' when HighFlag = False. For single cut points iBin is 1 to 2, and for dual cut points iBin is
' 1 to 3, with 1 being high values. Also note that the Palette is named from high to low (e.g. "Green_to_Red"
' means green is for high values and red is for low values.

iType = BList(Index, iRow).Type
iPalette = BList(Index, iRow).Palette

' Bubble Sort of Vector in Ascending order.
myLongVal = NValSortVec
SortVectorUp myLongVal, SortVec
' Now SortVec() should be in sorted, ascending order.

myLongVal = NValSortVec
Call GetQuintile(myLongVal, N1, N2, N3, N4)
'iType = 4 = quintiles
If iType = 4 And NValSortVec > 1 Then
'    If BList(Index, iRow).HighFlag = True Then
        For I = 1 To NValDataVec
            S = DataVec(I)
            'only assign a symbol type if the data is not a missing value
            If S <> na Then
                ' only assign a symbol type if the data is within the year display range
                If BList(Index, iRow).StartYear + I - 1 >= rptStartYr And _
                BList(Index, iRow).StartYear + I - 1 <= rptEndYr Then
                    K = BList(Index, iRow).StartYear + I - rptStartYr
                    X = Val(S)
                    If X < SortVec(N1 + 1) Then
                        If BList(Index, iRow).ZeroFlag And Abs(X) > 0# Or BList(Index, iRow).ZeroFlag = False Then
                            icon(K) = GetSymbol(iType, iPalette, 5) '"red.bmp"
                        End If
                    ElseIf X < SortVec(N2 + 1) Then
                        icon(K) = GetSymbol(iType, iPalette, 4) '"halfred.bmp"
                    ElseIf X < SortVec(N3 + 1) Then
                        icon(K) = GetSymbol(iType, iPalette, 3) '"white.bmp"
                    ElseIf X < SortVec(N4 + 1) Then
                        icon(K) = GetSymbol(iType, iPalette, 2) '"halfblack.bmp"
                    Else
                        icon(K) = GetSymbol(iType, iPalette, 1) '"black.bmp"
                    End If
                End If
            End If
        Next I
' This part is not needed because color palette name determines symbol choice
'    Else
'        For I = 1 To NValDataVec
'            S = DataVec(I)
'            'only assign a symbol type if the data is not a missing value
'            If S <> na Then
'               If BList(Index, iRow).StartYear + I - 1 >= rptStartYr And _
'               BList(Index, iRow).StartYear + I - 1 <= rptEndYr Then
'                   K = BList(Index, iRow).StartYear + I - rptStartYr
'                   x = val(DataVec(I))
'                   If x < SortVec(N1 + 1) Then
'                       If BList(Index, iRow).ZeroFlag And Abs(x) > 0# Or BList(Index, iRow).ZeroFlag = False Then
'                           icon(K) = "black.bmp"
'                       End If
'                   ElseIf x < SortVec(N2 + 1) Then
'                       icon(K) = "halfblack.bmp"
'                   ElseIf x < SortVec(N3 + 1) Then
'                       icon(K) = "white.bmp"
'                   ElseIf x < SortVec(N4 + 1) Then
'                       icon(K) = "halfred.bmp"
'                   Else
'                       icon(K) = "red.bmp"
'                   End If
'               End If
'           End If
'        Next I
'    End If
'iType = 2 = single cut point
ElseIf iType = 2 Then
    For I = 1 To NValDataVec
        S = DataVec(I)
        'only assign a symbol type if the data is not a missing value
        If S <> na Then
            If BList(Index, iRow).StartYear + I - 1 >= rptStartYr And _
                BList(Index, iRow).StartYear + I - 1 <= rptEndYr Then
                K = BList(Index, iRow).StartYear + I - rptStartYr
                X = Val(DataVec(I))
                If X < BList(Index, iRow).LowerCut Then
                    icon(K) = GetSymbol(iType, iPalette, 2) '"minus.bmp"
                Else
                    icon(K) = GetSymbol(iType, iPalette, 1) '"plus.bmp"
                End If
            End If
        End If
    Next I
'iType = 3 = dual cut points
ElseIf iType = 3 Then
    For I = 1 To NValDataVec
        S = DataVec(I)
        'only assign a symbol type if the data is not a missing value
        If S <> na Then
            If BList(Index, iRow).StartYear + I - 1 >= rptStartYr And _
                BList(Index, iRow).StartYear + I - 1 <= rptEndYr Then
                K = BList(Index, iRow).StartYear + I - rptStartYr
                X = Val(DataVec(I))
                If X < BList(Index, iRow).LowerCut Then
                    icon(K) = GetSymbol(iType, iPalette, 3) '"minus.bmp"
                ElseIf X > BList(Index, iRow).UpperCut Then
                    icon(K) = GetSymbol(iType, iPalette, 1) '"plus.bmp"
                Else
                    icon(K) = GetSymbol(iType, iPalette, 2) '"white.bmp"
                End If
            End If
        End If
    Next I
End If

For J = 1 To rptYears
    'print corresponding icon
    S = "<td class=icon>"
    If icon(J) <> "" Then
        S = S & "<img src=" & Chr(34) & "VRsymbols\" & icon(J) & Chr(34) & " alt=" & Chr(34) & icon(J) & Chr(34) & ">"
    Else
        S = S & "<p class=intable>&nbsp;</p>"
    End If
    S = S & "</td>"
    Print #1, S
Next J

'make sure symbol exists in HTML file's symbol folder
For J = 1 To rptYears
    If icon(J) <> "" Then
        I = InStrRev(HTMLFile, "\")
        S = Left(HTMLFile, I)
        S = S & "VRsymbols\" & icon(J)
        T = App.Path & "\VRsymbols\" & icon(J)
        If Dir(S) = "" Then FileCopy T, S
    End If
Next J

'get dispersion statistic
If DoStat(Index) = True And NValSortVec > 0 Then
    'construct format statement for printing out dispersion statistic
    T = "0"
    If StatDig(Index) > 0 Then
        T = T + "."
        For I = 1 To StatDig(Index)
            T = T + "0"
        Next I
    End If
    
    'find median
    If NValSortVec Mod 2 = 0 Then
        iMed = (SortVec(NValSortVec / 2) + SortVec(Int(NValSortVec / 2) + 1)) / 2
    Else
        iMed = SortVec(Int(NValSortVec / 2) + 1)
    End If
    'Debug.Print (NValSortVec / 2), SortVec(NValSortVec / 2), (Int(NValSortVec / 2) + 1), SortVec(Int(NValSortVec / 2) + 1)
    'now get min and max
    iMin = SortVec(1)
    iMax = SortVec(NValSortVec)
    'calculate dispersion
    'To avoid divide by zero errors, convert dispersion statistic to a string and print
    If iMed = 0 Then
        'do first part of calculation
        iDispers = (iMax - iMin)
        'convert to a string
        DispersString = Format(iDispers, T)
        'append notation to let user know value is divided by zero
        DispersString = DispersString & " / 0"
    Else
        'do the calculation
        iDispers = (iMax - iMin) / iMed
        'convert to a string
        DispersString = Format(iDispers, T)
    End If
    'Debug.Print iMin, iMax, iMed, iDispers
    
    'now print out html
    S = "<td class=title><p class=intable>"
    S = S & DispersString
    S = S & "</p></td>"
    Print #1, S
End If

End Sub
Public Sub WriteLegend(Index As Integer)
Dim I As Integer
Dim J As Integer
Dim N As Integer
Dim S As String
Dim T As String
Dim NoQuintiles As Boolean
Dim NoCutPts As Boolean
Dim NLegends As Integer
Dim Legends() As String
Dim Labels() As String
Dim NewLegend As Boolean

'get the number of legends and the list of legends
ReDim Legends(0 To 1, 1 To 13)
'init first legend
NLegends = 1
Legends(0, 1) = CStr(BList(Index, 1).Type)
Legends(1, 1) = BList(Index, 1).Palette
'go find rest of legends
N = ReportInfo(Index).NLines
If N > 1 Then
    For I = 2 To N
        S = BList(Index, I).Palette
        If S <> "" Then
            'check if palette is already in list
            NewLegend = True
            For J = 1 To NLegends
                If Legends(1, J) = S Then
                    NewLegend = False
                    Exit For
                End If
            Next J
            If NewLegend Then
                'add legend to list
                NLegends = NLegends + 1
                Legends(0, NLegends) = CStr(BList(Index, I).Type)
                Legends(1, NLegends) = S
            End If
        End If
    Next I
End If

'get legend labels
ReDim Labels(1 To NLegends, 1 To 5)
For I = 1 To NLegends
    If Legends(0, I) = "2" Then
        Labels(I, 1) = "Above Cut Point"
        Labels(I, 2) = "Below Cut Point"
    ElseIf Legends(0, I) = "3" Then
        Labels(I, 1) = "Above Cut Point"
        Labels(I, 2) = "Between Cut Points"
        Labels(I, 3) = "Below Cut Point"
    ElseIf Legends(0, I) = "4" Then
        S = Legends(1, I)
        If S = "Black_to_Red" Or S = "Red_to_Black" Then
            N = 1
        ElseIf S = "Red_to_Blue" Or S = "Blue_to_Red" Then
            N = 2
        ElseIf S = "White_to_Black" Or S = "Black_to_White" Then
            N = 3
        End If
        For J = 1 To 5
            Labels(I, J) = ReportLegend(Index, N, J)
        Next J
    End If
Next I

'write out legends
Print #1, "<!-- begin Legend -->"
Print #1, "Legend<br>"
Print #1, "<table cellspacing=0 cellpadding=0>"

For I = 1 To NLegends
    Print #1, "<tr>"
    Print #1, "<td class=title>"
    Print #1, "<p class=intable>"
    
    N = Val(Legends(0, I))
    Select Case N
        Case 2: 'single cut point
            T = GetSymbol(N, Legends(1, I), 1)
            Print #1, ""
            Print #1, "<img src=" & Chr(34) & "VRsymbols\" & T & Chr(34) & " alt=" & Chr(34) & T & Chr(34) & ">"
            Print #1, Labels(I, 1)
            Print #1, "&nbsp; &nbsp;"
            
            T = GetSymbol(N, Legends(1, I), 2)
            Print #1, ""
            Print #1, "<img src=" & Chr(34) & "VRsymbols\" & T & Chr(34) & " alt=" & Chr(34) & T & Chr(34) & ">"
            Print #1, Labels(I, 2)
        Case 3: 'dual cut points
            T = GetSymbol(N, Legends(1, I), 1)
            Print #1, ""
            Print #1, "<img src=" & Chr(34) & "VRsymbols\" & T & Chr(34) & " alt=" & Chr(34) & T & Chr(34) & ">"
            Print #1, Labels(I, 1)
            Print #1, "&nbsp; &nbsp;"
            
            T = GetSymbol(N, Legends(1, I), 2)
            Print #1, ""
            Print #1, "<img src=" & Chr(34) & "VRsymbols\" & T & Chr(34) & " alt=" & Chr(34) & T & Chr(34) & ">"
            Print #1, Labels(I, 2)
            Print #1, "&nbsp; &nbsp;"
            
            T = GetSymbol(N, Legends(1, I), 3)
            Print #1, ""
            Print #1, "<img src=" & Chr(34) & "VRsymbols\" & T & Chr(34) & " alt=" & Chr(34) & T & Chr(34) & ">"
            Print #1, Labels(I, 3)
        Case 4: 'quintiles
            For J = 1 To 5
                T = GetSymbol(N, Legends(1, I), J)
                Print #1, ""
                Print #1, "<img src=" & Chr(34) & "VRsymbols\" & T & Chr(34) & " alt=" & Chr(34) & T & Chr(34) & ">"
                Print #1, Labels(I, J)
                If J < 5 Then Print #1, "&nbsp; &nbsp;"
            Next J
    End Select

    Print #1, ""
    Print #1, "</p>"
    Print #1, "</td></tr>"
Next I

'closing punctuation
Print #1, "</table>"
Print #1, "<!-- end Legend -->"


End Sub
Public Function GetSymbol(iType As Integer, iPalette As String, iBin As Integer) As String
'For quintiles iBin is 1 to 5, with 1 being high values for normal case, and 5 being high values
' when HighFlag = False. For single cut points iBin is 1 to 2, and for dual cut points iBin is
' 1 to 3, with 1 being high values. Note that the Palette is named from high to low (e.g. "Green_to_Red"
' means green is for high values and red is for low values.

If iType = 2 Then 'single cut point
    Select Case iPalette
        Case "Plain_Plus/Minus":
            If iBin = 1 Then
                GetSymbol = "plus.bmp"
            ElseIf iBin = 2 Then
                GetSymbol = "minus.bmp"
            End If
         Case "Green_Plus/Red_Minus":
            If iBin = 1 Then
                GetSymbol = "greenplus.bmp"
            ElseIf iBin = 2 Then
                GetSymbol = "redminus.bmp"
            End If
        Case "Red_Plus/Green_Minus":
            If iBin = 1 Then
                GetSymbol = "redplus.bmp"
            ElseIf iBin = 2 Then
                GetSymbol = "greenminus.bmp"
            End If
        Case "Green/Red":
            If iBin = 1 Then
                GetSymbol = "green.bmp"
            ElseIf iBin = 2 Then
                GetSymbol = "red.bmp"
            End If
        Case "Red/Green":
            If iBin = 1 Then
                GetSymbol = "red.bmp"
            ElseIf iBin = 2 Then
                GetSymbol = "green.bmp"
            End If
    End Select
ElseIf iType = 3 Then 'dual cut points
    Select Case iPalette
        Case "Plain_Plus_to_Minus":
            If iBin = 1 Then
                GetSymbol = "plus.bmp"
            ElseIf iBin = 2 Then
                GetSymbol = "white.bmp"
            ElseIf iBin = 3 Then
                GetSymbol = "minus.bmp"
            End If
        Case "Green_Plus_to_Red_Minus":
            If iBin = 1 Then
                GetSymbol = "greenplus.bmp"
            ElseIf iBin = 2 Then
                GetSymbol = "yellow.bmp"
            ElseIf iBin = 3 Then
                GetSymbol = "redminus.bmp"
            End If
        Case "Red_Plus_to_Green_Minus":
            If iBin = 1 Then
                GetSymbol = "redplus.bmp"
            ElseIf iBin = 2 Then
                GetSymbol = "yellow.bmp"
            ElseIf iBin = 3 Then
                GetSymbol = "greenminus.bmp"
            End If
        Case "Green_to_Red":
            If iBin = 1 Then
                GetSymbol = "green.bmp"
            ElseIf iBin = 2 Then
                GetSymbol = "yellow.bmp"
            ElseIf iBin = 3 Then
                GetSymbol = "red.bmp"
            End If
        Case "Red_to_Green":
            If iBin = 1 Then
                GetSymbol = "red.bmp"
            ElseIf iBin = 2 Then
                GetSymbol = "yellow.bmp"
            ElseIf iBin = 3 Then
                GetSymbol = "green.bmp"
            End If
    End Select
ElseIf iType = 4 Then 'quintiles
    Select Case iPalette
        Case "Red_to_Black": 'normal case
            If iBin = 1 Then
                GetSymbol = "q1a.bmp" 'red
            ElseIf iBin = 2 Then
                GetSymbol = "q1b.bmp" 'half-red
            ElseIf iBin = 3 Then
                GetSymbol = "q1c.bmp" 'white
            ElseIf iBin = 4 Then
                GetSymbol = "q1d.bmp" 'half-black
            ElseIf iBin = 5 Then
                GetSymbol = "q1e.bmp" 'black
            End If
        Case "Black_to_Red": 'reversed case
            If iBin = 1 Then
                GetSymbol = "q1e.bmp" 'black
            ElseIf iBin = 2 Then
                GetSymbol = "q1d.bmp" 'half-black
            ElseIf iBin = 3 Then
                GetSymbol = "q1c.bmp" 'white
            ElseIf iBin = 4 Then
                GetSymbol = "q1b.bmp" 'half-red
            ElseIf iBin = 5 Then
                GetSymbol = "q1a.bmp" 'red
            End If
        Case "Red_to_Blue": 'normal case
            If iBin = 1 Then
                GetSymbol = "q2a.bmp" 'red
            ElseIf iBin = 2 Then
                GetSymbol = "q2b.bmp" 'pink
            ElseIf iBin = 3 Then
                GetSymbol = "q2c.bmp" 'white
            ElseIf iBin = 4 Then
                GetSymbol = "q2d.bmp" 'light blue
            ElseIf iBin = 5 Then
                GetSymbol = "q2e.bmp" 'blue
            End If
        Case "Blue_to_Red": 'reversed case
            If iBin = 1 Then
                GetSymbol = "q2e.bmp" 'blue
            ElseIf iBin = 2 Then
                GetSymbol = "q2d.bmp" 'light blue
            ElseIf iBin = 3 Then
                GetSymbol = "q2c.bmp" 'white
            ElseIf iBin = 4 Then
                GetSymbol = "q2b.bmp" 'pink
            ElseIf iBin = 5 Then
                GetSymbol = "q2a.bmp" 'red
            End If
        Case "White_to_Black": 'normal case
            If iBin = 1 Then
                GetSymbol = "q3a.bmp" 'white
            ElseIf iBin = 2 Then
                GetSymbol = "q3b.bmp" 'lightest gray
            ElseIf iBin = 3 Then
                GetSymbol = "q3c.bmp" 'middle gray
            ElseIf iBin = 4 Then
                GetSymbol = "q3d.bmp" 'darkest gray
            ElseIf iBin = 5 Then
                GetSymbol = "q3e.bmp" 'black
            End If
        Case "Black_to_White": 'reversed case
            If iBin = 1 Then
                GetSymbol = "q3e.bmp" 'black
            ElseIf iBin = 2 Then
                GetSymbol = "q3d.bmp" 'darkest gray
            ElseIf iBin = 3 Then
                GetSymbol = "q3c.bmp" 'middle gray
            ElseIf iBin = 4 Then
                GetSymbol = "q3b.bmp" 'lightest gray
            ElseIf iBin = 5 Then
                GetSymbol = "q3a.bmp" 'white
            End If
    End Select
End If

End Function

Public Function AssignLegend(Tag As String) As Integer

If Tag = "Red_to_Black" Or Tag = "Black_to_Red" Then
    AssignLegend = 1
ElseIf Tag = "Red_to_Blue" Or Tag = "Blue_to_Red" Then
    AssignLegend = 2
ElseIf Tag = "White_to_Black" Or Tag = "Black_to_White" Then
    AssignLegend = 3
End If

End Function
Public Sub InitFrmLegends()
Dim I As Integer
Dim J As Integer
Dim N As Integer
Dim S As String
Dim T As String
Dim iPalette As String
Dim NoPaletteFlag As Boolean

'which report
N = Val(frmLegends.lblReport.Caption) - 1

'get palette selections
NoPaletteFlag = True
For I = 1 To 3
    If ReportLegend(N, I, 0) <> "" Then
        NoPaletteFlag = False
        Exit For
    End If
Next I
If NoPaletteFlag Then
    frmLegends.lblNoPalette.Visible = True
    frmLegends.SSTab1.Visible = False
Else
    frmLegends.lblNoPalette.Visible = False
    frmLegends.SSTab1.Visible = True
    For I = 1 To 3
        iPalette = ReportLegend(N, I, 0)
        If iPalette = "" Then
            frmLegends.SSTab1.TabVisible(I - 1) = False
        Else
            frmLegends.SSTab1.TabVisible(I - 1) = True
            'show symbols and labels
            For J = 1 To 5
                S = GetSymbol(4, iPalette, J)
                S = App.Path + "\VRsymbols\" + S
                T = ReportLegend(N, I, J)
                If I = 1 Then
                    frmLegends.Picture1(J - 1).Picture = LoadPicture(S)
                    If Trim(T) <> "" Then frmLegends.txtLegend1(J - 1).Text = T
                ElseIf I = 2 Then
                    frmLegends.Picture2(J - 1).Picture = LoadPicture(S)
                    If Trim(T) <> "" Then frmLegends.txtLegend2(J - 1).Text = T
                ElseIf I = 3 Then
                    frmLegends.Picture3(J - 1).Picture = LoadPicture(S)
                    If Trim(T) <> "" Then frmLegends.txtLegend3(J - 1).Text = T
                End If
            Next J
        End If
    Next I
End If

'show dispersion calculation option
If DoStat(N) = True Then
    frmLegends.chkDispersion.Value = 1
    frmLegends.lblSigDig.Enabled = True
    frmLegends.txtSigDig.Enabled = True
    frmLegends.txtSigDig.Text = CStr(StatDig(N))
Else
    frmLegends.chkDispersion.Value = 0
    frmLegends.lblSigDig.Enabled = False
    frmLegends.txtSigDig.Enabled = False
End If

'show cut point location options
If CutPtLoc(N) = "beside" Then
    frmLegends.OpCutPts(0).Value = True
    frmLegends.OpCutPts(1).Value = False
Else
    frmLegends.OpCutPts(0).Value = False
    frmLegends.OpCutPts(1).Value = True
End If
        
End Sub
Public Function GetDataGridDesc(iRow As Integer) As String
Dim G As MSFlexGrid
Dim I As Integer
Dim S As String

Set G = frmData.MSFlexGrid1
S = ""
For I = 1 To 4
    S = S & G.TextMatrix(iRow, I) & " "
Next I
GetDataGridDesc = S

End Function
Public Function GetDataGridData(iRow As Integer) As String
Dim G As MSFlexGrid
Dim I As Integer
Dim S As String

Set G = frmData.MSFlexGrid1
S = ""
For I = 5 To G.Cols - 1
    If Trim(G.TextMatrix(iRow, I)) <> "" Then
        S = S & G.TextMatrix(iRow, I) & " "
    End If
Next I
GetDataGridData = S
End Function

Public Sub CreateBackup(LogName As String, BakSource As String)
'Save the existing project as a backup and append file info to backup log.
'Arguements:
'  LogName is the name of the project log file
'  BakSource is a string indicating the source of the backup:
'    If source is the user, BakSource = "User"
'    If source is GUI, BakSource = "File-Open", etc.
'The backup log is a tab-delimited file that contains:
'  (1) backup name (w/o path),
'  (2) original file name (enclosed in quotes),
'  (3) date and time (separated by spaces),
'  (4) source of the backup (user or automatic)
Dim BakLog As String 'text file recording the backup file names
Dim P As String 'backup path
Dim Bak() As String 'backup names and original file names for backups
Dim Bakroot As String 'root file name for backup
Dim Logroot As String 'root file name for original log file
Dim DT As String 'date and time of backup
Dim S As String
Dim T As String
Dim I As Integer
Dim N As Integer
Dim Count As Integer
    
On Error GoTo WriteErr

'initialize flag to check for successful completion
RedFlag = True

'Prepare vector of backup file names. If the max number of backups has already been reached,
' remember to get the very last one so the backup files can be deleted.
' NBackups = global variable for the number of backup copies to keep (hardwired at 15 for now)
ReDim Bak(1 To NBackups + 1)

'get date and time strings
DT = DtStr 'formatted for easy reading
T = DtStrCompact 'compact format

'put current project in the first slot of backup list
Bakroot = "VRbak" & T
S = Bakroot & ".log   " & Chr(9) & LogName & Chr(9) & "   " & DT & Chr(9) & BakSource
Bak(1) = S

'first, check for backups directory
P = App.Path & "\VRbackup"
If Dir(P, vbDirectory) = "" Then MkDir P
P = P & "\"

'assign backup log name
BakLog = P & "VRbackups.txt"

'read backup log
If Dir(BakLog) <> "" Then
    Open BakLog For Input As #4
        'skip header lines
        Line Input #4, S
        Line Input #4, S
        Line Input #4, S
        'get old backup info
        Count = 0
        Do While Not EOF(4)
            Line Input #4, S
            If Trim(S) <> "" Then
                Count = Count + 1
                Bak(Count + 1) = S
            Else
                Exit Do
            End If
            If Count = NBackups Then Exit Do
        Loop
    Close #4
End If

'If the max number of backups has already been reached, delete the very last one.
T = Trim(Bak(NBackups + 1))
If T <> "" Then
    'delete log file name
    N = InStr(T, Chr(9))
    S = Trim(Left(T, N - 1))
    S = P & S
    If Dir(S) <> "" Then Kill S
    'delete data file
    S = Left(S, Len(S) - 3) & "csv"
    If Dir(S) <> "" Then Kill S
    'delete report layout file
    S = Left(S, Len(S) - 3) & "txt"
    If Dir(S) <> "" Then Kill S
End If

'write backup log
Open BakLog For Output As #4
Print #4, "Visual Report Designer Backup Log"
Print #4, ""
S = "BACKUP_NAME" & Chr(9) & "ORIGINAL_NAME" & Chr(9) & "DATE/TIME" & Chr(9) & "BACKUP_SOURCE"
Print #4, S
For I = 1 To NBackups
    Print #4, Bak(I)
Next I
Close #4

'write note to log file, e.g.
' "Backup Created (from File-Open): C:NFT\VisualReport\VRbackup\file.log   at  "
S = "Backup Created (from " & BakSource & "): " & P & Bakroot & ".log   at  " & DT
Print #3, S
Print #3, ""

'copy files to backup directory and rename
'get log file root
Logroot = Left(LogName, Len(LogName) - 4)
'data file
T = Logroot & ".csv"
S = P & Bakroot & ".csv"
If Dir(T) <> "" Then FileCopy T, S
'report layout file
T = Logroot & ".txt"
S = P & Bakroot & ".txt"
If Dir(T) <> "" Then FileCopy T, S
'close log file to copy it
Close #3
'copy log file
S = P & Bakroot & ".log"
FileCopy LogName, S
'open log file back up again to resume writing
Open FNLOG For Append As #3

're-set warning flag
RedFlag = False

Exit Sub

WriteErr:
    Close #4
    MsgBox "Error Creating Backup", vbExclamation, "Visual Report Designer"
End Sub

Public Sub InitRestore()
Dim BakLog As String 'text file recording the backup file names
Dim Bak() As String 'backup names and original file names for backups
Dim Count As Integer
Dim S As String
Dim P As String
Dim I As Integer
Dim G As MSFlexGrid
Dim N As Integer

On Error GoTo WriteErr

'first, check for backups directory
P = App.Path & "\VRbackup"
If Dir(P, vbDirectory) = "" Then
    MsgBox "No Backups are Available", vbInformation, "Visual Report Designer"
    Exit Sub
End If
P = P & "\"

'assign backup log name
BakLog = P & "VRbackups.txt"
If Dir(BakLog) = "" Then
    MsgBox "No Backups are Available", vbInformation, "Visual Report Designer"
    Exit Sub
End If

'prepare vector of backup file names
ReDim Bak(1 To NBackups)

'read backup log
Open BakLog For Input As #4
'skip header lines
Line Input #4, S
Line Input #4, S
Line Input #4, S
Count = 0
'get old backup info
Do While Not EOF(4)
    Line Input #4, S
    If Trim(S) <> "" Then
        Count = Count + 1
        Bak(Count) = S
    Else
        Exit Do
    End If
    If Count = NBackups Then Exit Do
Loop
Close #4

'put file names in backup grid
Set G = frmBackup.MSFlexGrid1
G.Cols = 4
G.Rows = Count + 1
G.ColWidth(0) = 0 'hide backup name so user can view more of the original file name
G.TextMatrix(0, 0) = "Backup Name"
G.ColWidth(1) = 6500
G.TextMatrix(0, 1) = "Original File Name"
G.ColWidth(2) = 1700
G.TextMatrix(0, 2) = "Backup Date/Time"
G.ColWidth(3) = 1600
G.TextMatrix(0, 3) = "Backup Source"
For I = 1 To Count
    'backup log contains (tab-delimited):
    '  (1) backup name (w/o path),
    '  (2) original file name (enclosed in quotes),
    '  (3) date and time (separated by spaces),
    '  (4) source of backup (user or automatic)
    
    'extract backup name
    N = InStr(Bak(I), Chr(9))
    S = Left(Bak(I), N - 1)
    G.TextMatrix(I, 0) = Trim(S)
    
    'extract original file name
    Bak(I) = Mid(Bak(I), N + 1)
    N = InStr(Bak(I), Chr(9))
    S = Left(Bak(I), N - 1)
    G.TextMatrix(I, 1) = Trim(S)
    
    'extract date/time
    Bak(I) = Mid(Bak(I), N + 1)
    N = InStr(Bak(I), Chr(9))
    S = Left(Bak(I), N - 1)
    G.TextMatrix(I, 2) = Trim(S)
    
    'extract backup source
    Bak(I) = Mid(Bak(I), N + 1)
    S = Trim(Bak(I))
    G.TextMatrix(I, 3) = Trim(S)
Next I

frmBackup.Show vbModal


Exit Sub

WriteErr:
    Close #4
    MsgBox "Error Reading Backup Log", vbExclamation, "Visual Report Designer"
End Sub

Public Sub RestoreBackup()
    Dim N As Integer
    Dim G As MSFlexGrid
    Dim BakFile As String
    Dim OldFile As String
    Dim P As String
    Dim TempLOG As String
    Dim CopyFrom As String
    Dim CopyTo As String
    Dim S As String
    Dim DT As String
    Dim BakInfo As String
    Dim RS As Integer

    On Error GoTo CopyErr
    
    
    'Get the name of the backup file and the name of the original file
    Set G = frmBackup.MSFlexGrid1
    N = G.Row 'which row user has selected
    BakFile = Trim(G.TextMatrix(N, 0))
    OldFile = Trim(G.TextMatrix(N, 1))
    
    'create informational text to append to log file
    BakInfo = "Original File:" & OldFile
    S = G.TextMatrix(N, 3)
    BakInfo = BakInfo & " , Saved By: " & Trim(S)
    S = G.TextMatrix(N, 2)
    BakInfo = BakInfo & " , On: " & Trim(S)
    
    'check to see if the backup file still exists
    P = App.Path & "\VRbackup\"
    BakFile = P & BakFile
    If Dir(P) = "" Then
        MsgBox "The Backup Files You Have Selected No Longer Exist." & vbCrLf & _
            "Please Select Another One.", vbExclamation, "Visual Report Designer"
        Exit Sub
    End If
    
    Unload frmBackup
    
    'append "_BAK" to original file name
    N = Len(OldFile) - 4
    OldFile = Left(OldFile, N) & "_BAK.log"
    
    'now get new file name to save as
    frmGeneral.CommonDialog1.FileName = OldFile
    frmGeneral.CommonDialog1.Flags = &H806
    frmGeneral.CommonDialog1.DialogTitle = "Save Project Backup Log File As"
    frmGeneral.CommonDialog1.Filter = "Log Files (*.log)|*.log"
    frmGeneral.CommonDialog1.CancelError = True
    frmGeneral.CommonDialog1.FilterIndex = 0
    frmGeneral.CommonDialog1.DefaultExt = "log"
    frmGeneral.CommonDialog1.CancelError = True
    frmGeneral.CommonDialog1.ShowSave
    TempLOG = frmGeneral.CommonDialog1.FileName
    N = InStr(UCase(TempLOG), "ST6UNST.LOG")
    If N > 0 Then
        MsgBox TempLOG + vbCrLf + "Is Not A Valid Project File Name", vbInformation, "Visual Report Designer"
        Exit Sub
    End If
    
    If TempLOG <> "" Then
        'if file doesn't have ".log" appended (when user types a file name with
        ' a period in it and also doesn't explictly type ".log")
        If LCase(Right(TempLOG, 4)) <> ".log" Then
            TempLOG = TempLOG + ".log"
        End If
        
        'copy files to new location
        FileCopy BakFile, TempLOG
        'get root file name from backup name
        N = Len(BakFile) - 3
        S = Left(BakFile, N)
        'copy data file
        CopyFrom = S & "csv"
        If Dir(CopyFrom) <> "" Then
            N = Len(TempLOG) - 3
            CopyTo = Left(TempLOG, N) & "csv"
            FileCopy CopyFrom, CopyTo
        End If
        'copy report layout file
        CopyFrom = S & "txt"
        If Dir(CopyFrom) <> "" Then
            N = Len(TempLOG) - 3
            CopyTo = Left(TempLOG, N) & "txt"
            FileCopy CopyFrom, CopyTo
        End If
        
        'backup the current file before unloading
        'first make sure there is a file available
        S = Trim(frmGeneral.lblFile.Caption)
        If S <> "" And LCase(S) <> "none specified" Then
            'do backup
            CreateBackup S, "Close File"
        End If
        
        'close file before opening new one
        If FNLOG <> "" Then Close #3
        
        Unload frmData
        Unload frmCompose
        
        FF1 = False
        
        FNLOG = TempLOG
        frmGeneral.lblFile.Caption = FNLOG
        S = "Visual Report Designer - Version " & CStr(VRVersion) & " - "
        S = S & FNLOG
        frmMain.Caption = S
        DT = DtStr
        Open FNLOG For Append As #3
        Print #3, ""
        Print #3, "*** Project Restored from Backup: " + DT + " ***"
        Print #3, BakInfo
        
        LogFlag = True
        DataFlag = False
        LayoutFlag = False
        frmMain.mnuDataMissingReplace.Enabled = False
        ReadGridData
        If DataFlag Then
            RS = MsgBox("Do You Wish to Add Existing Layouts Back into the Project?", vbQuestion + vbYesNo, "Visual Report Designer")
            If RS = vbYes Then
                ReadLayoutFile
                If LayoutFlag = True Then MsgBox "All Report Layouts Have Been Restored", vbInformation, "Visual Report Designer"
            End If
        Else 'if there is no data file, allow users to begin adding data
            StartYear = 0
            EndYear = 0
            NYears = 0
            frmMain.mnuDataAdd.Enabled = True
        End If
        
        'backup the new file
        CreateBackup FNLOG, "Restore Backup"
    End If


Exit Sub

CopyErr:
    'be silent if user canceled out of the file save dialog
    If Not Err.Number = cdlCancel Then
        MsgBox "Error Restoring Project From Backup", vbExclamation, "Visual Report Designer"
    End If
End Sub
Public Sub DeleteBackups()
Dim P As String

P = App.Path & "\VRbackup"
If Dir(P, vbDirectory) = "" Then
    MsgBox "There Are No Backups to Clear", vbInformation, "Visual Report Designer"
    Exit Sub
Else
    P = P & "\*.*"
    Kill P
    MsgBox "All Backups Have Been Cleared", vbInformation, "Visual Report Designer"
End If
End Sub
Public Sub ReplaceNA(matchVal As String)
'In the Data Collection Grid, replace all values equaling 'matchVal'
'with the user-specified missing/no-data value 'NAVal'
Dim I As Integer
Dim J As Integer
Dim G As MSFlexGrid
Dim S As String
Dim DT As String
Dim Count As Integer

Set G = frmData.MSFlexGrid1

Count = 0
'if any grid cells match the specified text, replace it with the missing value indicator
For I = 1 To G.Rows - 1
    For J = 5 To G.Cols - 1
        S = Trim(G.TextMatrix(I, J))
        If S = Trim(matchVal) Then
            G.TextMatrix(I, J) = NAVal
            Count = Count + 1
        End If
    Next J
Next I

Unload frmReplaceNA

S = "All " & Count & " Instances of Data Matching: " & matchVal
S = S & vbCrLf & "Were Replaced with the Missing Value Indicator: " & NAVal
MsgBox S, vbInformation, "Visual Report Designer"

'write a note to the log file
Print #3, "***** DATA IN DATA COLLECTION GRID WAS MODIFIED ***"
S = "All " & Count & " Instances of Data Matching: " & matchVal
Print #3, S
S = "Were Replaced with the Missing Value Indicator: " & NAVal
Print #3, S
'get date and time strings
DT = DtStr
Print #3, DT
Print #3, ""

End Sub
