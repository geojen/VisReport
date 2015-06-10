Attribute VB_Name = "Utils"
Private JRow As Integer
Private JCol As Integer
Private NRow As Integer
Private NCol As Integer
Public FG As MSFlexGrid
Public Rmin As Integer
Public Rmax As Integer
Public Cmin As Integer
Public Cmax As Integer

Public Sub GetCellMinMax(iGrd As MSFlexGrid)
'get row and column selections and assign to
' topmost, bottommost, leftmost and rightmost variables

    With iGrd
        
        If .RowSel < .Row Then
            Rmin = .RowSel
            Rmax = .Row
        Else
            Rmin = .Row
            Rmax = .RowSel
        End If
        If .ColSel < .Col Then
            Cmin = .ColSel
            Cmax = .Col
        Else
            Cmin = .Col
            Cmax = .ColSel
        End If
        
    End With

End Sub

Public Sub CopyToClipboard()
    Dim S As String
    Dim I As Integer
    Dim J As Integer
    
    ' Chr(9) is tab
    
    GetCellMinMax FG
    
    S = ""
    For I = Rmin To Rmax
        For J = Cmin To Cmax
            S = S + FG.TextMatrix(I, J)
            If J < Cmax Then
                S = S + Chr(9)
            End If
        Next J
        S = S + vbCrLf
    Next I
    Clipboard.Clear
    Clipboard.SetText S
    ClearGridSelection
    
End Sub
Public Sub CutGridData()
    Dim S As String
    Dim I As Integer
    Dim J As Integer
    
    ' Chr(9) is tab
    
    GetCellMinMax FG
    
    S = ""
    For I = Rmin To Rmax
        For J = Cmin To Cmax
            S = S + FG.TextMatrix(I, J)
            FG.TextMatrix(I, J) = ""
            If J < Cmax Then
                S = S + Chr(9)
            End If
        Next J
        S = S + vbCrLf
    Next I
    Clipboard.Clear
    Clipboard.SetText S
    ClearGridSelection
    
End Sub
Public Sub PasteFromClipboard()
    Dim I As Integer
    Dim J As Integer
    Dim N As Integer
    Dim M As Integer
    Dim S As String
    Dim T As String
    Dim RStart As Integer
    Dim CStart As Integer
    
    ' Chr(9) is tab
    ' Chr(10) is line feed
    ' Chr(13) is carriage return
    
    RStart = FG.Row
    CStart = FG.Col
    
    S = Clipboard.GetText
    
    'add carriage return/line feed at end of string, if not present
    If Right(S, 2) <> vbCrLf Then S = S + vbCrLf
    
    For I = RStart To FG.Rows - 1
        If S = "" Then Exit For
        For J = CStart To FG.Cols - 1
            'find first value; delimited by tabs
            N = InStr(S, Chr(9))
            M = InStr(S, vbCrLf)
            If N = 0 Or N > M Then
                'this is last value in the row
                T = Left(S, M - 1)
                FG.TextMatrix(I, J) = T
                S = Mid(S, M + 2)
                Exit For
            ElseIf N < M Then
                'not the last value in the row; get value
                T = Left(S, N - 1)
                FG.TextMatrix(I, J) = T
                If J = FG.Cols - 1 Then
                    'if data exceeds column boundary, chop of rest of row
                    S = Mid(S, M + 2)
                    Exit For
                Else
                    'otherwise just chop off current value
                    S = Mid(S, N + 1)
                End If
            End If
        Next J
    Next I
    
    ClearGridSelection
    
End Sub
Public Sub ClearGridSelection()
    FG.CellBackColor = RGB(255, 255, 255)
    FG.RowSel = FG.Row
    FG.ColSel = FG.Col
    'Unload frmCopyPaste
End Sub
