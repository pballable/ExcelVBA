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
        QuizAnswer = InputBox("第" & i & "/6問" & vbCrLf & _
                    "正解数　" & QuizTrueCnt & "/6問" & vbCrLf & _
                    QuizSheet.Cells(QuizRand, QuizQueCol))
        
        If StrPtr(QuizAnswer) = 0 Then
            MsgBox ("使ってくれてありがとうございます！")
            Exit Sub
        End If
        
        If QuizAnswer = QuizJudge Then
            MsgBox ("正解です！")
            QuizSheet.Cells(QuizRand, QuizTrueCol).Value = _
            QuizSheet.Cells(QuizRand, QuizTrueCol).Value + 1
            QuizTrueCnt = QuizTrueCnt + 1
        Else
            MsgBox ("不正解です..." & vbCrLf & "正解は「" & QuizJudge & "」でした。")
            QuizSheet.Cells(QuizRand, QuizFalseCol).Value = _
            QuizSheet.Cells(QuizRand, QuizFalseCol).Value + 1
        End If
        
        QuizSheet.Cells(QuizRand, QuizTotalCol).Value = _
        QuizSheet.Cells(QuizRand, QuizTotalCol).Value + 1
        
        QuizSheet.Cells(QuizRand, QuizRateCol).Formula = RateFormula
        QuizSheet.Cells(QuizRand, QuizDateCol).Value = Now()
        
    Next i
    
    If QuizTrueCnt >= 4 Then
        MsgBox ("正解数は6問中" & QuizTrueCnt & "問でした！" & vbCrLf & "合格！")
    Else
        MsgBox ("正解数は6問中" & QuizTrueCnt & "問でした..." & vbCrLf & _
                "えいしゃおらー！（不合格）")
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
                MsgBox ("使ってくれてありがとうございます！")
                Exit Sub
            End If
        
            If QuizAnswer = QuizJudge Then
                MsgBox ("正解です！")
                QuizSheet.Cells(i, QuizTrueCol).Value = _
                QuizSheet.Cells(i, QuizTrueCol).Value + 1
            Else
                MsgBox ("不正解です..." & vbCrLf & "正解は「" & QuizJudge & "」でした。")
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
        MsgBox ("未回答の問題はもうありません。お疲れさまでした。")
    Else
        MsgBox ("未回答の問題が存在しませんでした。")
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
    
    QuizGenreInput = InputBox("ジャンルを入力してください")
    
    Do
        QuizRand = Int(Rnd * QuizAmount) + QuizDataRow
        QuizJudge = QuizSheet.Cells(QuizRand, QuizAnsCol)
        QuizGenre = QuizSheet.Cells(QuizRand, QuizGenreCol)
        
        If QuizGenre = QuizGenreInput Then
            QuizAnswer = InputBox(QuizSheet.Cells(QuizRand, QuizQueCol))
        
            If StrPtr(QuizAnswer) = 0 Then
                MsgBox ("使ってくれてありがとうございます！")
                Exit Sub
            End If
        
            If QuizAnswer = QuizJudge Then
                MsgBox ("正解です！")
                QuizSheet.Cells(QuizRand, QuizTrueCol).Value = _
                QuizSheet.Cells(QuizRand, QuizTrueCol).Value + 1
            Else
                MsgBox ("不正解です..." & vbCrLf & "正解は「" & QuizJudge & "」でした。")
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
        
        If QuizGenre = "アニメ＆ゲーム" Then
            QuizAnswer = InputBox(QuizSheet.Cells(QuizRand, QuizQueCol))
        
            If StrPtr(QuizAnswer) = 0 Then
                MsgBox ("使ってくれてありがとうございます！")
                Exit Sub
            End If
        
            If QuizAnswer = QuizJudge Then
                MsgBox ("正解です！")
                QuizSheet.Cells(QuizRand, QuizTrueCol).Value = _
                QuizSheet.Cells(QuizRand, QuizTrueCol).Value + 1
            Else
                MsgBox ("不正解です..." & vbCrLf & "正解は「" & QuizJudge & "」でした。")
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
                MsgBox ("使ってくれてありがとうございます！")
                Exit Sub
            End If
        
            If QuizAnswer = QuizJudge Then
                MsgBox ("正解です！")
                QuizSheet.Cells(i, QuizTrueCol).Value = _
                QuizSheet.Cells(i, QuizTrueCol).Value + 1
            Else
                MsgBox ("不正解です..." & vbCrLf & "正解は「" & QuizJudge & "」でした。")
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
        MsgBox ("弱点克服はこれで終了です。お疲れさまでした。")
    Else
        MsgBox ("正解率50％未満の問題が存在しませんでした。")
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
            MsgBox ("使ってくれてありがとうございます！")
            Exit Sub
        End If
        
        If QuizAnswer = QuizJudge Then
            MsgBox ("正解です！")
            QuizSheet.Cells(QuizRand, QuizTrueCol).Value = _
            QuizSheet.Cells(QuizRand, QuizTrueCol).Value + 1
        Else
            MsgBox ("不正解です..." & vbCrLf & "正解は「" & QuizJudge & "」でした。")
            QuizSheet.Cells(QuizRand, QuizFalseCol).Value = _
            QuizSheet.Cells(QuizRand, QuizFalseCol).Value + 1
        End If
        
        QuizSheet.Cells(QuizRand, QuizTotalCol).Value = _
        QuizSheet.Cells(QuizRand, QuizTotalCol).Value + 1
        
        QuizSheet.Cells(QuizRand, QuizRateCol).Formula = RateFormula
        QuizSheet.Cells(QuizRand, QuizDateCol).Value = Now()
        
    Loop
    
End Sub

