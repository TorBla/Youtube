Public Sub MTA()

'Autor:     Torben Blankertz
'Version:   1.0
'Datum:     03.11.2015
'Infos:     Mit diesem Script und der dazugehörigen Excel Datei kann eine Meilensteintrendanalyse erstellt werden.

Dim FilePfad As String
Dim Datei As Variant
Dim FileName As String
Dim TaskName As String
Dim TaskDatum As Date
Dim oWebApp As Object
Dim oDatei, TaskExcName As String

Datei = Excel.Application.GetOpenFilename(FileFilter:="Excel, *.xls; *.xlsx", Title:="Wählen Sie den Bericht für die Meilensteintrendanalyse aus:", MultiSelect:=False)

'Öffne die Excel Applikation

Set oWebApp = Excel.Workbooks.Open(Datei)

Excel.Application.Visible = True

a = 8
b = 3

MTAVers = Excel.Application.Sheets("Daten").Cells(2, 3).Value
ProjectName = Application.ActiveProject.Name


b = b + MTAVers

For Each T In ActiveProject.Tasks
   
    If T.Milestone = True Then
    
    TaskName = T.Name
    TaskDatum = Format(T.Start, "dd.mm.yyyy")
    
     TaskExcName = Excel.Application.Sheets("Daten").Cells(a, 2).Value
     
            'Prüfe nach, ob der Name in der Excel Datei den gleichen Namen hat, wie die Project Datei
            If TaskExcName = "" Then
        
                'Sollte dem nicht so sein, dann lege die ersten Meilensteine an und setze gleichzeitig das Datum in die nachfolgende SPalte
                                
                Excel.Application.Sheets("Daten").Cells(a, 2).Value = TaskName
                Excel.Application.Sheets("Daten").Cells(a, 3).Value = TaskDatum
                Excel.Application.Sheets("Daten").Cells(2, 3).Value = 1
                Excel.Application.Sheets("Daten").Cells(3, 3).Value = ProjectName
                Excel.Application.Sheets("Daten").Cells(4, 3).Value = (Environ("Username"))
                
                Excel.Application.Sheets("Daten").Cells(7, 3).Value = "Report vom" & Chr(10) & Date
             
                a = a + 1
                              
                     
                'Sollten schon Daten geschrieben worden sein, so prüfe schreibe das Datum in die nächste Zeile
            Else
            
                'Liest den aktuellen Meilensteinnamen in die Variable MileStoneName ein
                'MilestoneName = oWebApp.Application.Sheets("Daten").Cells(a, 2).Value
                
                MilestoneName = Excel.Application.Sheets("Daten").Cells(a, 2).Value
                
                'Frage nach, ob der Meilenstein in der Excel-Datei mit dem in der Project-Datei übereinstimmt.
                If MilestoneName <> T.Name Then
                
                    'Gebe eine Warnmeldung aus, dass die Namen der
                    MsgBox "Achtung: Einige Meilensteinnamen in der Project-Datei stimmt nicht mit den Meilensteinnamen in der Excel-Datei überein! Bitte überprüfen Sie diese und führen Sie danach das Programm erneut aus.", vbInformation, "Meilensteintrendanalyse"
                
                    'Markiere den Meilenstein in gelb, der nicht mit dem in der Excel-Datei übereinstimmt
                    
                    Excel.Sheets("Daten").Activate
                    Excel.ActiveSheet.Cells(a, 2).Select
                    
                   
            
                    With Selection.Interior
                    
                     .Color = 65535
                    
                    End With
                
                End If
                
                'Prüft nach, ob der Meilenstein schon erreicht wurde (100%). Sollte dem so sein, wird das Feld leergelassen
                'Sollte dem nicht so sein, wird das aktuelle Startdatum eingetragen.
                
                If T.PercentComplete = 100 Then
                
                    Excel.Application.Sheets("Daten").Cells(7, b).Value = "Report vom" & Chr(10) & Date
                    Excel.Application.Sheets("Daten").Cells(a, b).Value = ""
                    Excel.Application.Sheets("Daten").Cells(2, 3).Value = MTAVers + 1
                
                Else
                
                    Excel.Application.Sheets("Daten").Cells(7, b).Value = "Report vom" & Chr(10) & Date
                    Excel.Application.Sheets("Daten").Cells(a, b).Value = TaskDatum
                    Excel.Application.Sheets("Daten").Cells(2, 3).Value = MTAVers + 1
                
                End If
                
                a = a + 1
                       
            End If
 
    End If
    
Next
' Variable a muss um eins runtergezählt werden, da in der Schleife Each die letzte leere Zeile mitgezählt wird!
    a = a - 1
    Excel.Application.Sheets("MTA").Select
    
    Excel.Application.ActiveSheet.ChartObjects(1).Select
    
    Excel.Application.ActiveChart.SetSourceData Source:= _
    Excel.Application.Sheets("Daten").Range( _
    Excel.Application.Sheets("Daten").Cells(7, 2), _
    Excel.Application.Sheets("Daten").Cells(a, b))
    
    Excel.ActiveChart.PlotBy = xlRows

Set oWebApp = Nothing

End Sub