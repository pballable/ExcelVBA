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
    
    VBAns = MsgBox("���ׂĂ̖��f�[�^���폜���Ă���낵���ł����H", vbOKCancel)
    If VBAns = vbOK Then
        VBAns = MsgBox("�{���ɂ�낵���ł����H", vbOKCancel)
        If VBAns = vbOK Then
            QuizRange.Delete
            MsgBox ("���ׂĂ̖��f�[�^���폜���܂����B")
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
    
    VBAns = MsgBox("QuizData�V�[�g��3��ڂ���6��ڂ̃f�[�^�����ׂč폜����܂�����낵���ł����H", vbOKCancel)
    If VBAns = vbOK Then
        VBAns = MsgBox("�{���ɂ�낵���ł����H", vbOKCancel)
        If VBAns = vbOK Then
            QuizSheet.Range(QuizSheet.Cells(QuizDataRow, QuizTrueCol), _
            QuizSheet.Cells(EndRow, QuizTotalCol)).ClearContents
            MsgBox ("�𓚃f�[�^�̍폜���������܂����B�������I�����܂��B")
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
    TempData = InputBox("�폜�����������s���Ŏw�肵�Ă��������B(2�`1048576)")
    If StrPtr(TempData) = 0 Then
        VBAns = MsgBox("���̍폜�𒆎~���܂����H", vbOKCancel)
        If VBAns = vbOK Then
            MsgBox ("���̒ǉ��𒆎~���܂����B")
            QuizMenu.Activate
            Exit Sub
        End If
    End If
    
    If (TempData < 2) Or (TempData > 1048576) Then
        MsgBox ("�G���[���������܂����B�������I�����܂��B")
        QuizMenu.Activate
        Exit Sub
    End If
    
    If QuizSheet.Cells(TempData, startCol) = "" Then
        QuizSheet.Activate
        ActiveWindow.ScrollRow = TempData
        ActiveWindow.ScrollColumn = startCol
        QuizSheet.Cells(TempData, startCol).Activate
        MsgBox ("�w�肵���s�ɖ�肪���͂���Ă��܂���B" & vbCrLf _
                & "������x��蒼���Ă��������B")
        QuizMenu.Activate
        Exit Sub
    End If
    
    QuizSheet.Activate
    ActiveWindow.ScrollRow = TempData
    ActiveWindow.ScrollColumn = startCol
    QuizSheet.Cells(TempData, startCol).Activate
    
    VBAns = MsgBox("�I������Ă���s���폜���܂����H", vbOKCancel)
    If VBAns = vbOK Then
        QuizSheet.Rows(TempData).Delete
        MsgBox ("�폜���܂����B")
    End If
    
    VBAns = MsgBox("���̍폜�𑱂��܂����H", vbOKCancel)
    If VBAns = vbOK Then
    Else
        MsgBox ("���̍폜�������I�����܂��B")
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
                TempData = InputBox("�N�C�Y�̖�����͂��Ă�������")
                QuizSheet.Cells(InsRow, startCol).Value = TempData
            Case 2
                TempData = InputBox("�N�C�Y�̓�������͂��Ă�������")
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
                TempData = InputBox("�N�C�Y�̃W����������͂��Ă�������")
                QuizSheet.Cells(InsRow, QuizGenreCol).Value = TempData
            Case 9
                QuizSheet.Activate
                ActiveWindow.ScrollRow = InsRow
                ActiveWindow.ScrollColumn = startCol
                QuizSheet.Cells(InsRow, startCol).Activate
                VBAns = MsgBox("���̒ǉ����������܂����B" & vbCrLf _
                            & "���ǉ��𑱂��܂����H", vbOKCancel)
                If VBAns = vbOK Then
                Else
                    Exit Sub
                End If
        End Select
        
        If StrPtr(TempData) = 0 Then
            VBAns = MsgBox("���̒ǉ��𒆎~���܂����H", vbOKCancel)
            If VBAns = vbOK Then
                QuizSheet.Rows(InsRow).Clear
                MsgBox ("���̒ǉ��𒆎~���܂����B")
                Exit Sub
            End If
        End If
    Next i
    
    Loop
    
End Sub
