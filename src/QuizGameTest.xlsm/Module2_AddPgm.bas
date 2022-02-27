Attribute VB_Name = "Module2_AddPgm"
Option Explicit
Sub QuizGame_DeleteAllData()
    Dim EndRow As Long
    Dim QuizBook As Workbook
    Dim QuizSheet As Worksheet
    Dim QuizMenu As Worksheet
    Dim QuizRange As Range
    Dim VBAns As Integer
    
    Set QuizBook = ThisWorkbook
    Set QuizSheet = QuizBook.Worksheets("QuizData")
    Set QuizMenu = QuizBook.Worksheets("QuizMenu")
    
    EndRow = QuizSheet.Cells(StartRow, startCol).CurrentRegion.Rows.Count
    
    QuizSheet.Activate
    Set QuizRange = QuizSheet.Range(QuizDataRow & ":" & EndRow)
    QuizRange.Select
    
    VBAns = MsgBox("すべての問題データを削除してもよろしいですか？", vbOKCancel)
    If VBAns = vbOK Then
        VBAns = MsgBox("本当によろしいですか？", vbOKCancel)
        If VBAns = vbOK Then
            QuizRange.Delete
            MsgBox ("すべての問題データを削除しました。")
        End If
    End If
    
    QuizSheet.Cells(StartRow, startCol).Select
    QuizMenu.Activate
End Sub

Sub QuizGame_DeleteAnsData()
    Dim EndRow As Long
    Dim QuizBook As Workbook
    Dim QuizSheet As Worksheet
    Dim QuizRange As Range
    Dim QuizMenu As Worksheet
    Dim VBAns As Integer
    
    Set QuizBook = ThisWorkbook
    Set QuizSheet = QuizBook.Worksheets("QuizData")
    Set QuizMenu = QuizBook.Worksheets("QuizMenu")
    
    EndRow = QuizSheet.Cells(StartRow, startCol).CurrentRegion.Rows.Count
    
    Set QuizRange = QuizSheet.Range(QuizSheet.Cells(QuizDataRow, QuizTrueCol), _
                    QuizSheet.Cells(EndRow, QuizTotalCol))
    
    QuizSheet.Activate
    ActiveWindow.ScrollRow = StartRow
    ActiveWindow.ScrollColumn = startCol
    QuizSheet.Cells(StartRow, startCol).Activate
    
    QuizRange.Select
    
    VBAns = MsgBox("QuizDataシートの3列目から6列目のデータがすべて削除されますがよろしいですか？", vbOKCancel)
    If VBAns = vbOK Then
        VBAns = MsgBox("本当によろしいですか？", vbOKCancel)
        If VBAns = vbOK Then
            QuizSheet.Range(QuizSheet.Cells(QuizDataRow, QuizTrueCol), _
            QuizSheet.Cells(EndRow, QuizTotalCol)).ClearContents
            MsgBox ("解答データの削除が完了しました。処理を終了します。")
            QuizSheet.Cells(StartRow, startCol).Select
            QuizMenu.Activate
            Exit Sub
        Else
            QuizSheet.Cells(StartRow, startCol).Select
            QuizMenu.Activate
            Exit Sub
        End If
        
    Else
        QuizSheet.Cells(StartRow, startCol).Select
        QuizMenu.Activate
        Exit Sub
    End If
    
End Sub

Sub QuizGame_Delete()
    Dim QuizBook As Workbook
    Dim QuizSheet As Worksheet
    Dim QuizMenu As Worksheet
    Dim i As Long
    Dim TempData As String
    Dim VBAns As Integer
    
    Set QuizBook = ThisWorkbook
    Set QuizSheet = QuizBook.Worksheets("QuizData")
    Set QuizMenu = QuizBook.Worksheets("QuizMenu")
    
    Do
    TempData = InputBox("削除したい問題を行数で指定してください。(2～1048576)")
    If StrPtr(TempData) = 0 Then
        VBAns = MsgBox("問題の削除を中止しますか？", vbOKCancel)
        If VBAns = vbOK Then
            MsgBox ("問題の追加を中止しました。")
            QuizMenu.Activate
            Exit Sub
        End If
    End If
    
    If (TempData < 2) Or (TempData > 1048576) Then
        MsgBox ("エラーが発生しました。処理を終了します。")
        QuizMenu.Activate
        Exit Sub
    End If
    
    If QuizSheet.Cells(TempData, startCol) = "" Then
        QuizSheet.Activate
        ActiveWindow.ScrollRow = TempData
        ActiveWindow.ScrollColumn = startCol
        QuizSheet.Cells(TempData, startCol).Activate
        MsgBox ("指定した行に問題が入力されていません。" & vbCrLf _
                & "もう一度やり直してください。")
        QuizMenu.Activate
        Exit Sub
    End If
    
    QuizSheet.Activate
    ActiveWindow.ScrollRow = TempData
    ActiveWindow.ScrollColumn = startCol
    QuizSheet.Cells(TempData, startCol).Activate
    
    VBAns = MsgBox("選択されている行を削除しますか？", vbOKCancel)
    If VBAns = vbOK Then
        QuizSheet.Rows(TempData).Delete
        MsgBox ("削除しました。")
    End If
    
    VBAns = MsgBox("問題の削除を続けますか？", vbOKCancel)
    If VBAns = vbOK Then
    Else
        MsgBox ("問題の削除処理を終了します。")
        QuizMenu.Activate
        Exit Sub
    End If
    
    Loop
    
    
    
End Sub


Sub QuizGame_QuizAdd()
    Dim EndRow As Long
    Dim InsRow As Long
    Dim QuizBook As Workbook
    Dim QuizSheet As Worksheet
    Dim i As Long
    Dim TempData As String
    Dim VBAns As Integer
    
    Set QuizBook = ThisWorkbook
    Set QuizSheet = QuizBook.Worksheets("QuizData")
    
    Do
    
    EndRow = QuizSheet.Cells(StartRow, startCol).CurrentRegion.Rows.Count
    InsRow = EndRow + 1
    
    For i = 1 To 9
        Select Case i
            Case 1
                TempData = InputBox("クイズの問題を入力してください")
                QuizSheet.Cells(InsRow, startCol).Value = TempData
            Case 2
                TempData = InputBox("クイズの答えを入力してください")
                QuizSheet.Cells(InsRow, QuizAnsCol).Value = TempData
            Case 3
                QuizSheet.Cells(InsRow, QuizTrueCol).Value = 0
            Case 4
                QuizSheet.Cells(InsRow, QuizFalseCol).Value = 0
            Case 5
                QuizSheet.Cells(InsRow, QuizTotalCol).Value = 0
            Case 6
                QuizSheet.Cells(InsRow, QuizRateCol).Formula = RateFormula
            Case 7
            Case 8
                TempData = InputBox("クイズのジャンルを入力してください")
                QuizSheet.Cells(InsRow, QuizGenreCol).Value = TempData
            Case 9
                QuizSheet.Activate
                ActiveWindow.ScrollRow = InsRow
                ActiveWindow.ScrollColumn = startCol
                QuizSheet.Cells(InsRow, startCol).Activate
                VBAns = MsgBox("問題の追加が完了しました。" & vbCrLf _
                            & "問題追加を続けますか？", vbOKCancel)
                If VBAns = vbOK Then
                Else
                    Exit Sub
                End If
        End Select
        
        If StrPtr(TempData) = 0 Then
            VBAns = MsgBox("問題の追加を中止しますか？", vbOKCancel)
            If VBAns = vbOK Then
                QuizSheet.Rows(InsRow).Clear
                MsgBox ("問題の追加を中止しました。")
                Exit Sub
            End If
        End If
    Next i
    
    Loop
    
End Sub
