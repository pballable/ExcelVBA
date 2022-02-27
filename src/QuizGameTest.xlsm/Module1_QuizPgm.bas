Attribute VB_Name = "Module1_QuizPgm"
Option Explicit
Sub QuizMenu()
    

End Sub
Sub QuizGame_Main()
    Dim QuizRand As Long
    Dim QuizJudge As String
    Dim QuizAnswer As String
    Dim QuizBook As Workbook
    Dim QuizSheet As Worksheet
    Dim QuizRange As Range
    Dim QuizTrueCnt As Long
    Dim i As Long
    
    
    
    
    Set QuizBook = ThisWorkbook
    Set QuizSheet = QuizBook.Worksheets("QuizData")
    
    QuizTrueCnt = 0
    For i = 1 To 6
    
        QuizRand = Int(Rnd * QuizAmount) + 2
        QuizJudge = QuizSheet.Cells(QuizRand, QuizAnsCol)
        QuizAnswer = InputBox("��" & i & "/6��" & vbCrLf & _
                    "���𐔁@" & QuizTrueCnt & "/6��" & vbCrLf & _
                    QuizSheet.Cells(QuizRand, QuizQueCol))
        
        If StrPtr(QuizAnswer) = 0 Then
            MsgBox ("�g���Ă���Ă��肪�Ƃ��������܂��I")
            Exit Sub
        End If
        
        If QuizAnswer = QuizJudge Then
            MsgBox ("�����ł��I")
            QuizSheet.Cells(QuizRand, QuizTrueCol).Value = _
            QuizSheet.Cells(QuizRand, QuizTrueCol).Value + 1
            QuizTrueCnt = QuizTrueCnt + 1
        Else
            MsgBox ("�s�����ł�..." & vbCrLf & "�����́u" & QuizJudge & "�v�ł����B")
            QuizSheet.Cells(QuizRand, QuizFalseCol).Value = _
            QuizSheet.Cells(QuizRand, QuizFalseCol).Value + 1
        End If
        
        QuizSheet.Cells(QuizRand, QuizTotalCol).Value = _
        QuizSheet.Cells(QuizRand, QuizTotalCol).Value + 1
        
        QuizSheet.Cells(QuizRand, QuizRateCol).Formula = RateFormula
        QuizSheet.Cells(QuizRand, QuizDateCol).Value = Now()
        
    Next i
    
    If QuizTrueCnt >= 4 Then
        MsgBox ("���𐔂�6�⒆" & QuizTrueCnt & "��ł����I" & vbCrLf & "���i�I")
    Else
        MsgBox ("���𐔂�6�⒆" & QuizTrueCnt & "��ł���..." & vbCrLf & _
                "�������Ⴈ��[�I�i�s���i�j")
    End If
    
    
End Sub
Sub QuizGame_StudyTraining()
    Dim QuizJudge As String
    Dim QuizAnswer As String
    Dim QuizTotal As Long
    Dim QuizBook As Workbook
    Dim QuizSheet As Worksheet
    Dim QuizRange As Range
    Dim QuizCount As Long
    Dim i As Long
    
    
    
    
    Set QuizBook = ThisWorkbook
    Set QuizSheet = QuizBook.Worksheets("QuizData")
    
    QuizCount = 0
    For i = QuizDataRow To QuizAmount
         QuizJudge = QuizSheet.Cells(i, QuizAnsCol)
         QuizTotal = QuizSheet.Cells(i, QuizTotalCol)
            
         If QuizTotal = 0 Then
            QuizAnswer = InputBox(QuizSheet.Cells(i, QuizQueCol))
        
        
            If StrPtr(QuizAnswer) = 0 Then
                MsgBox ("�g���Ă���Ă��肪�Ƃ��������܂��I")
                Exit Sub
            End If
        
            If QuizAnswer = QuizJudge Then
                MsgBox ("�����ł��I")
                QuizSheet.Cells(i, QuizTrueCol).Value = _
                QuizSheet.Cells(i, QuizTrueCol).Value + 1
            Else
                MsgBox ("�s�����ł�..." & vbCrLf & "�����́u" & QuizJudge & "�v�ł����B")
                QuizSheet.Cells(i, QuizFalseCol).Value = _
                QuizSheet.Cells(i, QuizFalseCol).Value + 1
            End If
        
            QuizSheet.Cells(i, QuizTotalCol).Value = _
            QuizSheet.Cells(i, QuizTotalCol).Value + 1
        
            QuizSheet.Cells(i, QuizRateCol).Formula = RateFormula
            QuizSheet.Cells(i, QuizDateCol).Value = Now()
            
            QuizCount = QuizCount + 1
        End If
        
    Next i
    
    If QuizCount > 0 Then
        MsgBox ("���񓚂̖��͂�������܂���B����ꂳ�܂ł����B")
    Else
        MsgBox ("���񓚂̖�肪���݂��܂���ł����B")
    End If

End Sub
Sub QuizGame_Training_GenreSelect()
    Dim QuizRand As Long
    Dim QuizJudge As String
    Dim QuizAnswer As String
    Dim QuizBook As Workbook
    Dim QuizSheet As Worksheet
    Dim QuizRange As Range
    Dim QuizGenre As String
    Dim QuizGenreInput As String
    
    
    
    Set QuizBook = ThisWorkbook
    Set QuizSheet = QuizBook.Worksheets("QuizData")
    
    QuizGenreInput = InputBox("�W����������͂��Ă�������")
    
    Do
        QuizRand = Int(Rnd * QuizAmount) + QuizDataRow
        QuizJudge = QuizSheet.Cells(QuizRand, QuizAnsCol)
        QuizGenre = QuizSheet.Cells(QuizRand, QuizGenreCol)
        
        If QuizGenre = QuizGenreInput Then
            QuizAnswer = InputBox(QuizSheet.Cells(QuizRand, QuizQueCol))
        
            If StrPtr(QuizAnswer) = 0 Then
                MsgBox ("�g���Ă���Ă��肪�Ƃ��������܂��I")
                Exit Sub
            End If
        
            If QuizAnswer = QuizJudge Then
                MsgBox ("�����ł��I")
                QuizSheet.Cells(QuizRand, QuizTrueCol).Value = _
                QuizSheet.Cells(QuizRand, QuizTrueCol).Value + 1
            Else
                MsgBox ("�s�����ł�..." & vbCrLf & "�����́u" & QuizJudge & "�v�ł����B")
                QuizSheet.Cells(QuizRand, QuizFalseCol).Value = _
                QuizSheet.Cells(QuizRand, QuizFalseCol).Value + 1
            End If
        
            QuizSheet.Cells(QuizRand, QuizTotalCol).Value = _
            QuizSheet.Cells(QuizRand, QuizTotalCol).Value + 1
        
            QuizSheet.Cells(QuizRand, QuizRateCol).Formula = RateFormula
            QuizSheet.Cells(QuizRand, QuizDateCol).Value = Now()
        End If
    Loop
End Sub

Sub QuizGame_Training_Genre01()
    Dim QuizRand As Long
    Dim QuizJudge As String
    Dim QuizAnswer As String
    Dim QuizBook As Workbook
    Dim QuizSheet As Worksheet
    Dim QuizRange As Range
    Dim QuizGenre As String
    
    
    
    Set QuizBook = ThisWorkbook
    Set QuizSheet = QuizBook.Worksheets("QuizData")
    
    
    Do
        QuizRand = Int(Rnd * QuizAmount) + QuizDataRow
        QuizJudge = QuizSheet.Cells(QuizRand, QuizAnsCol)
        QuizGenre = QuizSheet.Cells(QuizRand, QuizGenreCol)
        
        If QuizGenre = "�A�j�����Q�[��" Then
            QuizAnswer = InputBox(QuizSheet.Cells(QuizRand, QuizQueCol))
        
            If StrPtr(QuizAnswer) = 0 Then
                MsgBox ("�g���Ă���Ă��肪�Ƃ��������܂��I")
                Exit Sub
            End If
        
            If QuizAnswer = QuizJudge Then
                MsgBox ("�����ł��I")
                QuizSheet.Cells(QuizRand, QuizTrueCol).Value = _
                QuizSheet.Cells(QuizRand, QuizTrueCol).Value + 1
            Else
                MsgBox ("�s�����ł�..." & vbCrLf & "�����́u" & QuizJudge & "�v�ł����B")
                QuizSheet.Cells(QuizRand, QuizFalseCol).Value = _
                QuizSheet.Cells(QuizRand, QuizFalseCol).Value + 1
            End If
        
            QuizSheet.Cells(QuizRand, QuizTotalCol).Value = _
            QuizSheet.Cells(QuizRand, QuizTotalCol).Value + 1
        
            QuizSheet.Cells(QuizRand, QuizRateCol).Formula = RateFormula
            QuizSheet.Cells(QuizRand, QuizDateCol).Value = Now()
        End If
    Loop

End Sub

Sub QuizGame_WeakTraining()
    Dim QuizJudge As String
    Dim QuizAnswer As String
    Dim QuizRate As Double
    Dim QuizBook As Workbook
    Dim QuizSheet As Worksheet
    Dim QuizRange As Range
    Dim QuizCount As Long
    Dim QuizTotal As Long
    Dim i As Long
    
    
    
    
    Set QuizBook = ThisWorkbook
    Set QuizSheet = QuizBook.Worksheets("QuizData")
    
    QuizCount = 0
    For i = QuizDataRow To QuizAmount
         QuizJudge = QuizSheet.Cells(i, QuizAnsCol)
         QuizRate = QuizSheet.Cells(i, QuizRateCol)
         QuizTotal = QuizSheet.Cells(i, QuizTotalCol)
            
         If (QuizRate < 0.5) And (QuizTotal <> 0) Then
            QuizAnswer = InputBox(QuizSheet.Cells(i, QuizQueCol))
        
        
            If StrPtr(QuizAnswer) = 0 Then
                MsgBox ("�g���Ă���Ă��肪�Ƃ��������܂��I")
                Exit Sub
            End If
        
            If QuizAnswer = QuizJudge Then
                MsgBox ("�����ł��I")
                QuizSheet.Cells(i, QuizTrueCol).Value = _
                QuizSheet.Cells(i, QuizTrueCol).Value + 1
            Else
                MsgBox ("�s�����ł�..." & vbCrLf & "�����́u" & QuizJudge & "�v�ł����B")
                QuizSheet.Cells(i, QuizFalseCol).Value = _
                QuizSheet.Cells(i, QuizFalseCol).Value + 1
            End If
        
            QuizSheet.Cells(i, QuizTotalCol).Value = _
            QuizSheet.Cells(i, QuizTotalCol).Value + 1
        
            QuizSheet.Cells(i, QuizRateCol).Formula = RateFormula
            QuizSheet.Cells(i, QuizDateCol).Value = Now()
            
            QuizCount = QuizCount + 1
        End If
        
    Next i
    
    If QuizCount > 0 Then
        MsgBox ("��_�����͂���ŏI���ł��B����ꂳ�܂ł����B")
    Else
        MsgBox ("����50�������̖�肪���݂��܂���ł����B")
    End If
        
End Sub

Sub QuizGame_Training()
    Dim QuizRand As Long
    Dim QuizJudge As String
    Dim QuizAnswer As String
    Dim QuizBook As Workbook
    Dim QuizSheet As Worksheet
    Dim QuizRange As Range
    
    
    
    Set QuizBook = ThisWorkbook
    Set QuizSheet = QuizBook.Worksheets("QuizData")
    
    
    Do
        QuizRand = Int(Rnd * QuizAmount) + QuizDataRow
        QuizJudge = QuizSheet.Cells(QuizRand, QuizAnsCol)
        QuizAnswer = InputBox(QuizSheet.Cells(QuizRand, QuizQueCol))
        
        If StrPtr(QuizAnswer) = 0 Then
            MsgBox ("�g���Ă���Ă��肪�Ƃ��������܂��I")
            Exit Sub
        End If
        
        If QuizAnswer = QuizJudge Then
            MsgBox ("�����ł��I")
            QuizSheet.Cells(QuizRand, QuizTrueCol).Value = _
            QuizSheet.Cells(QuizRand, QuizTrueCol).Value + 1
        Else
            MsgBox ("�s�����ł�..." & vbCrLf & "�����́u" & QuizJudge & "�v�ł����B")
            QuizSheet.Cells(QuizRand, QuizFalseCol).Value = _
            QuizSheet.Cells(QuizRand, QuizFalseCol).Value + 1
        End If
        
        QuizSheet.Cells(QuizRand, QuizTotalCol).Value = _
        QuizSheet.Cells(QuizRand, QuizTotalCol).Value + 1
        
        QuizSheet.Cells(QuizRand, QuizRateCol).Formula = RateFormula
        QuizSheet.Cells(QuizRand, QuizDateCol).Value = Now()
        
    Loop
    
End Sub

