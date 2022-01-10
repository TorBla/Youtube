'Dieses Script ist ergänzend zu meinem Webcast auf YouTube
'https://youtube.com/user/torbla74
'Das Script ist nach bestem Wissen und Gewissen erstellt.

'Autor: Torben Blankertz
'Datum: 10.01.2022

Sub BudgetChecker()

Dim A As Task

For Each A In ActiveProject.Tasks


    If A.Cost3 < 0 Then
        A.Marked = True
    Else
        A.Marked = False
    End If

Next


End Sub