Attribute VB_Name = "Module3_OtherPgm"
Option Explicit
Sub hyperlink2()
    Dim QuizBook As Workbook
    Dim QuizSheet As Worksheet
    
    Set QuizBook = ThisWorkbook
    Set QuizSheet = QuizBook.Worksheets("QuizMenu")
    
    QuizSheet.Activate
End Sub


Sub hyperlink1()
    Dim QuizBook As Workbook
    Dim QuizSheet As Worksheet
    
    Set QuizBook = ThisWorkbook
    Set QuizSheet = QuizBook.Worksheets("GenreSelect")
    
    QuizSheet.Activate
    
End Sub
