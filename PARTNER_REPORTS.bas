Sub PARTNERS_REPORTS_czy_dwa_zero()


Application.ScreenUpdating = False

ActiveSheet.Name = "Consumption_Report"

Range("A:AD").Sort Key1:=Range("C1"), Header:=xlYes

[A:A, B:B, D:D, E:E, F:F, G:G, I:I, J:J, O:O, P:P, Q:Q, R:R, S:S, T:T, Z:Z, AC:AC].Delete

[1:1].Font.Bold = True

    Sheets.Add
    Sheets("Sheet1").Name = "Sonru"
    Sheets.Add
    Sheets("Sheet2").Name = "Workhoppers.com"
    Sheets.Add
    Sheets("Sheet3").Name = "Active Job Board"
    Sheets.Add
    Sheets("Sheet4").Name = "Totallyhired inc."
    Sheets.Add
    Sheets("Sheet5").Name = "SalesGravy"
    Sheets.Add
    Sheets("Sheet6").Name = "Recroup"
    Sheets.Add
    Sheets("Sheet7").Name = "Performance Assessment Network"
    Sheets.Add
    Sheets("Sheet8").Name = "PURE JOBS"
    Sheets.Add
    Sheets("Sheet9").Name = "LevoLeague"
    Sheets.Add
    Sheets("Sheet10").Name = "ITJobCafe"
    Sheets.Add
    Sheets("Sheet11").Name = "GlassDoorPro"
    Sheets.Add
    Sheets("Sheet12").Name = "Geebo"
    Sheets.Add
    Sheets("Sheet13").Name = "Good&Co"
    Sheets.Add
    Sheets("Sheet14").Name = "FashionUnited"
    Sheets.Add
    Sheets("Sheet15").Name = "Engineer Nexus LLC"
    Sheets.Add
    Sheets("Sheet16").Name = "Bio Careers"
    Sheets.Add
    Sheets("Sheet17").Name = "AccountantJobs.com"
    Sheets.Add
    Sheets("Sheet18").Name = "JobTeaser"
    Sheets.Add
    Sheets("Sheet19").Name = "Jobing.com"
    Sheets.Add
    Sheets("Sheet20").Name = "Adaface"
    
    
 
Sheets("Consumption_Report").Select
Application.CutCopyMode = False
Selection.AutoFilter

If Not IsError(Application.Match("Sonru", Sheets("Consumption_Report").Range("C:C"), 0)) Then
    Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("A:M").AutoFilter Field:=3, Criteria1:="Sonru"
    ActiveSheet.Range("A:M").AutoFilter Field:=11, Criteria1:="SUCCESS"
    Range("A1:M" & Cells(Rows.Count, "A").End(xlUp).Row).Select
    Selection.Copy
    Sheets("Sonru").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Columns("A:M").EntireColumn.AutoFit
Else
    Application.DisplayAlerts = False
    Sheets("Sonru").Select
    ActiveSheet.Delete
    Application.DisplayAlerts = True
End If



Sheets("Consumption_Report").Select
Application.CutCopyMode = False
Selection.AutoFilter

If Not IsError(Application.Match("Workhoppers.com", Sheets("Consumption_Report").Range("C:C"), 0)) Then
    Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("A:M").AutoFilter Field:=3, Criteria1:="Workhoppers.com"
    ActiveSheet.Range("A:M").AutoFilter Field:=11, Criteria1:="SUCCESS"
    Range("A1:M" & Cells(Rows.Count, "A").End(xlUp).Row).Select
    Selection.Copy
    Sheets("Workhoppers.com").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Columns("A:M").EntireColumn.AutoFit
Else
    Application.DisplayAlerts = False
    Sheets("Workhoppers.com").Select
    ActiveSheet.Delete
    Application.DisplayAlerts = True
End If



Sheets("Consumption_Report").Select
Application.CutCopyMode = False
Selection.AutoFilter

If Not IsError(Application.Match("Active Job Board", Sheets("Consumption_Report").Range("C:C"), 0)) Then
    Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("A:M").AutoFilter Field:=3, Criteria1:="Active Job Board"
    ActiveSheet.Range("A:M").AutoFilter Field:=11, Criteria1:="SUCCESS"
    Range("A1:M" & Cells(Rows.Count, "A").End(xlUp).Row).Select
    Selection.Copy
    Sheets("Active Job Board").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Columns("A:M").EntireColumn.AutoFit
Else
    Application.DisplayAlerts = False
    Sheets("Active Job Board").Select
    ActiveSheet.Delete
    Application.DisplayAlerts = True
End If



Sheets("Consumption_Report").Select
Application.CutCopyMode = False
Selection.AutoFilter

If Not IsError(Application.Match("Totallyhired inc.", Sheets("Consumption_Report").Range("C:C"), 0)) Then
    Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("A:M").AutoFilter Field:=3, Criteria1:="Totallyhired inc."
    ActiveSheet.Range("A:M").AutoFilter Field:=11, Criteria1:="SUCCESS"
    Range("A1:M" & Cells(Rows.Count, "A").End(xlUp).Row).Select
    Selection.Copy
    Sheets("Totallyhired inc.").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Columns("A:M").EntireColumn.AutoFit
Else
    Application.DisplayAlerts = False
    Sheets("Totallyhired inc.").Select
    ActiveSheet.Delete
    Application.DisplayAlerts = True
End If



Sheets("Consumption_Report").Select
Application.CutCopyMode = False
Selection.AutoFilter

If Not IsError(Application.Match("SalesGravy", Sheets("Consumption_Report").Range("C:C"), 0)) Then
    Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("A:M").AutoFilter Field:=3, Criteria1:="SalesGravy"
    ActiveSheet.Range("A:M").AutoFilter Field:=11, Criteria1:="SUCCESS"
    Range("A1:M" & Cells(Rows.Count, "A").End(xlUp).Row).Select
    Selection.Copy
    Sheets("SalesGravy").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Columns("A:M").EntireColumn.AutoFit
Else
    Application.DisplayAlerts = False
    Sheets("SalesGravy").Select
    ActiveSheet.Delete
    Application.DisplayAlerts = True
End If



Sheets("Consumption_Report").Select
Application.CutCopyMode = False
Selection.AutoFilter

If Not IsError(Application.Match("Recroup", Sheets("Consumption_Report").Range("C:C"), 0)) Then
    Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("A:M").AutoFilter Field:=3, Criteria1:="Recroup"
    ActiveSheet.Range("A:M").AutoFilter Field:=11, Criteria1:="SUCCESS"
    Range("A1:M" & Cells(Rows.Count, "A").End(xlUp).Row).Select
    Selection.Copy
    Sheets("Recroup").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Columns("A:M").EntireColumn.AutoFit
Else
    Application.DisplayAlerts = False
    Sheets("Recroup").Select
    ActiveSheet.Delete
    Application.DisplayAlerts = True
End If



Sheets("Consumption_Report").Select
Application.CutCopyMode = False
Selection.AutoFilter

If Not IsError(Application.Match("Performance Assessment Network", Sheets("Consumption_Report").Range("C:C"), 0)) Then
   Range("A1").Select
   Selection.AutoFilter
   ActiveSheet.Range("A:M").AutoFilter Field:=3, Criteria1:="Performance Assessment Network"
   ActiveSheet.Range("A:M").AutoFilter Field:=11, Criteria1:="SUCCESS"
   Range("A1:M" & Cells(Rows.Count, "A").End(xlUp).Row).Select
   Selection.Copy
   Sheets("Performance Assessment Network").Select
   ActiveSheet.Paste
   Application.CutCopyMode = False
   Columns("A:M").EntireColumn.AutoFit
Else
    Application.DisplayAlerts = False
    Sheets("Performance Assessment Network").Select
    ActiveSheet.Delete
    Application.DisplayAlerts = True
End If



Sheets("Consumption_Report").Select
Application.CutCopyMode = False
Selection.AutoFilter

If Not IsError(Application.Match("PURE JOBS", Sheets("Consumption_Report").Range("C:C"), 0)) Then
   Range("A1").Select
   Selection.AutoFilter
   ActiveSheet.Range("A:M").AutoFilter Field:=3, Criteria1:="PURE JOBS"
   ActiveSheet.Range("A:M").AutoFilter Field:=11, Criteria1:="SUCCESS"
   Range("A1:M" & Cells(Rows.Count, "A").End(xlUp).Row).Select
   Selection.Copy
   Sheets("PURE JOBS").Select
   ActiveSheet.Paste
   Application.CutCopyMode = False
   Columns("A:M").EntireColumn.AutoFit
Else
    Application.DisplayAlerts = False
    Sheets("PURE JOBS").Select
    ActiveSheet.Delete
    Application.DisplayAlerts = True
End If



Sheets("Consumption_Report").Select
Application.CutCopyMode = False
Selection.AutoFilter

If Not IsError(Application.Match("LevoLeague", Sheets("Consumption_Report").Range("C:C"), 0)) Then
   Range("A1").Select
   Selection.AutoFilter
   ActiveSheet.Range("A:M").AutoFilter Field:=3, Criteria1:="LevoLeague"
   ActiveSheet.Range("A:M").AutoFilter Field:=11, Criteria1:="SUCCESS"
   Range("A1:M" & Cells(Rows.Count, "A").End(xlUp).Row).Select
   Selection.Copy
   Sheets("LevoLeague").Select
   ActiveSheet.Paste
   Application.CutCopyMode = False
   Columns("A:M").EntireColumn.AutoFit
Else
    Application.DisplayAlerts = False
    Sheets("LevoLeague").Select
    ActiveSheet.Delete
    Application.DisplayAlerts = True
End If



Sheets("Consumption_Report").Select
Application.CutCopyMode = False
Selection.AutoFilter

If Not IsError(Application.Match("ITJobCafe", Sheets("Consumption_Report").Range("C:C"), 0)) Then
   Range("A1").Select
   Selection.AutoFilter
   ActiveSheet.Range("A:M").AutoFilter Field:=3, Criteria1:="ITJobCafe"
   ActiveSheet.Range("A:M").AutoFilter Field:=11, Criteria1:="SUCCESS"
   Range("A1:M" & Cells(Rows.Count, "A").End(xlUp).Row).Select
   Selection.Copy
   Sheets("ITJobCafe").Select
   ActiveSheet.Paste
   Application.CutCopyMode = False
   Columns("A:M").EntireColumn.AutoFit
Else
    Application.DisplayAlerts = False
    Sheets("ITJobCafe").Select
    ActiveSheet.Delete
    Application.DisplayAlerts = True
End If



Sheets("Consumption_Report").Select
Application.CutCopyMode = False
Selection.AutoFilter

If Not IsError(Application.Match("GlassDoorPro", Sheets("Consumption_Report").Range("C:C"), 0)) Then
   Range("A1").Select
   Selection.AutoFilter
   ActiveSheet.Range("A:M").AutoFilter Field:=3, Criteria1:="GlassDoorPro"
   ActiveSheet.Range("A:M").AutoFilter Field:=11, Criteria1:="SUCCESS"
   Range("A1:M" & Cells(Rows.Count, "A").End(xlUp).Row).Select
   Selection.Copy
   Sheets("GlassDoorPro").Select
   ActiveSheet.Paste
   Application.CutCopyMode = False
   Columns("A:M").EntireColumn.AutoFit
Else
    Application.DisplayAlerts = False
    Sheets("GlassDoorPro").Select
    ActiveSheet.Delete
    Application.DisplayAlerts = True
End If



Sheets("Consumption_Report").Select
Application.CutCopyMode = False
Selection.AutoFilter

If Not IsError(Application.Match("Geebo", Sheets("Consumption_Report").Range("C:C"), 0)) Then
   Range("A1").Select
   Selection.AutoFilter
   ActiveSheet.Range("A:M").AutoFilter Field:=3, Criteria1:="Geebo"
   ActiveSheet.Range("A:M").AutoFilter Field:=11, Criteria1:="SUCCESS"
   Range("A1:M" & Cells(Rows.Count, "A").End(xlUp).Row).Select
   Selection.Copy
   Sheets("Geebo").Select
   ActiveSheet.Paste
   Application.CutCopyMode = False
   Columns("A:M").EntireColumn.AutoFit
Else
    Application.DisplayAlerts = False
    Sheets("Geebo").Select
    ActiveSheet.Delete
    Application.DisplayAlerts = True
End If



Sheets("Consumption_Report").Select
Application.CutCopyMode = False
Selection.AutoFilter

If Not IsError(Application.Match("Good&Co", Sheets("Consumption_Report").Range("C:C"), 0)) Then
   Range("A1").Select
   Selection.AutoFilter
   ActiveSheet.Range("A:M").AutoFilter Field:=3, Criteria1:="Good&Co"
   ActiveSheet.Range("A:M").AutoFilter Field:=11, Criteria1:="SUCCESS"
   Range("A1:M" & Cells(Rows.Count, "A").End(xlUp).Row).Select
   Selection.Copy
   Sheets("Good&Co").Select
   ActiveSheet.Paste
   Application.CutCopyMode = False
   Columns("A:M").EntireColumn.AutoFit
Else
    Application.DisplayAlerts = False
    Sheets("Good&Co").Select
    ActiveSheet.Delete
    Application.DisplayAlerts = True
End If



Sheets("Consumption_Report").Select
Application.CutCopyMode = False
Selection.AutoFilter

If Not IsError(Application.Match("FashionUnited", Sheets("Consumption_Report").Range("C:C"), 0)) Then
   Range("A1").Select
   Selection.AutoFilter
   ActiveSheet.Range("A:M").AutoFilter Field:=3, Criteria1:="FashionUnited"
   ActiveSheet.Range("A:M").AutoFilter Field:=11, Criteria1:="SUCCESS"
   Range("A1:M" & Cells(Rows.Count, "A").End(xlUp).Row).Select
   Selection.Copy
   Sheets("FashionUnited").Select
   ActiveSheet.Paste
   Application.CutCopyMode = False
   Columns("A:M").EntireColumn.AutoFit
Else
    Application.DisplayAlerts = False
    Sheets("FashionUnited").Select
    ActiveSheet.Delete
    Application.DisplayAlerts = True
End If



Sheets("Consumption_Report").Select
Application.CutCopyMode = False
Selection.AutoFilter

If Not IsError(Application.Match("Engineer Nexus LLC", Sheets("Consumption_Report").Range("C:C"), 0)) Then
   Range("A1").Select
   Selection.AutoFilter
   ActiveSheet.Range("A:M").AutoFilter Field:=3, Criteria1:="Engineer Nexus LLC"
   ActiveSheet.Range("A:M").AutoFilter Field:=11, Criteria1:="SUCCESS"
   Range("A1:M" & Cells(Rows.Count, "A").End(xlUp).Row).Select
   Selection.Copy
   Sheets("Engineer Nexus LLC").Select
   ActiveSheet.Paste
   Application.CutCopyMode = False
   Columns("A:M").EntireColumn.AutoFit
Else
    Application.DisplayAlerts = False
    Sheets("Engineer Nexus LLC").Select
    ActiveSheet.Delete
    Application.DisplayAlerts = True
End If



Sheets("Consumption_Report").Select
Application.CutCopyMode = False
Selection.AutoFilter

If Not IsError(Application.Match("Bio Careers", Sheets("Consumption_Report").Range("C:C"), 0)) Then
   Range("A1").Select
   Selection.AutoFilter
   ActiveSheet.Range("A:M").AutoFilter Field:=3, Criteria1:="Bio Careers"
   ActiveSheet.Range("A:M").AutoFilter Field:=11, Criteria1:="SUCCESS"
   Range("A1:M" & Cells(Rows.Count, "A").End(xlUp).Row).Select
   Selection.Copy
   Sheets("Bio Careers").Select
   ActiveSheet.Paste
   Application.CutCopyMode = False
   Columns("A:M").EntireColumn.AutoFit
Else
    Application.DisplayAlerts = False
    Sheets("Bio Careers").Select
    ActiveSheet.Delete
    Application.DisplayAlerts = True
End If



Sheets("Consumption_Report").Select
Application.CutCopyMode = False
Selection.AutoFilter

If Not IsError(Application.Match("AccountantJobs.com", Sheets("Consumption_Report").Range("C:C"), 0)) Then
   Range("A1").Select
   Selection.AutoFilter
   ActiveSheet.Range("A:M").AutoFilter Field:=3, Criteria1:="AccountantJobs.com"
   ActiveSheet.Range("A:M").AutoFilter Field:=11, Criteria1:="SUCCESS"
   Range("A1:M" & Cells(Rows.Count, "A").End(xlUp).Row).Select
   Selection.Copy
   Sheets("AccountantJobs.com").Select
   ActiveSheet.Paste
   Application.CutCopyMode = False
   Columns("A:M").EntireColumn.AutoFit
Else
    Application.DisplayAlerts = False
    Sheets("AccountantJobs.com").Select
    ActiveSheet.Delete
    Application.DisplayAlerts = True
End If



Sheets("Consumption_Report").Select
Application.CutCopyMode = False
Selection.AutoFilter

If Not IsError(Application.Match("JobTeaser", Sheets("Consumption_Report").Range("C:C"), 0)) Then
   Range("A1").Select
   Selection.AutoFilter
   ActiveSheet.Range("A:M").AutoFilter Field:=3, Criteria1:="JobTeaser"
   ActiveSheet.Range("A:M").AutoFilter Field:=11, Criteria1:="SUCCESS"
   Range("A1:M" & Cells(Rows.Count, "A").End(xlUp).Row).Select
   Selection.Copy
   Sheets("JobTeaser").Select
   ActiveSheet.Paste
   Application.CutCopyMode = False
   Columns("A:M").EntireColumn.AutoFit
Else
    Application.DisplayAlerts = False
    Sheets("JobTeaser").Select
    ActiveSheet.Delete
    Application.DisplayAlerts = True
End If



Sheets("Consumption_Report").Select
Application.CutCopyMode = False
Selection.AutoFilter

If Not IsError(Application.Match("Jobing.com", Sheets("Consumption_Report").Range("C:C"), 0)) Then
   Range("A1").Select
   Selection.AutoFilter
   ActiveSheet.Range("A:M").AutoFilter Field:=3, Criteria1:="Jobing.com"
   ActiveSheet.Range("A:M").AutoFilter Field:=11, Criteria1:="SUCCESS"
   Range("A1:M" & Cells(Rows.Count, "A").End(xlUp).Row).Select
   Selection.Copy
   Sheets("Jobing.com").Select
   ActiveSheet.Paste
   Application.CutCopyMode = False
   Columns("A:M").EntireColumn.AutoFit
Else
    Application.DisplayAlerts = False
    Sheets("Jobing.com").Select
    ActiveSheet.Delete
    Application.DisplayAlerts = True
End If



Sheets("Consumption_Report").Select
Application.CutCopyMode = False
Selection.AutoFilter

If Not IsError(Application.Match("Adaface", Sheets("Consumption_Report").Range("C:C"), 0)) Then
    Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("A:M").AutoFilter Field:=3, Criteria1:="Adaface"
    ActiveSheet.Range("A:M").AutoFilter Field:=11, Criteria1:="SUCCESS"
    Range("A1:M" & Cells(Rows.Count, "A").End(xlUp).Row).Select
    Selection.Copy
    Sheets("Adaface").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Columns("A:M").EntireColumn.AutoFit
Else
    Application.DisplayAlerts = False
    Sheets("Adaface").Select
    ActiveSheet.Delete
    Application.DisplayAlerts = True
End If


Application.CutCopyMode = False

Sheets("Consumption_Report").Select

Selection.AutoFilter

Application.ScreenUpdating = True


End Sub

