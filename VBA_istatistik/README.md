# Some Statistical Calculation Examples For Grouped Frequency Series in Excel
Excel has a large function library, including statistical ones varying from arithmetic mean to Gamma probability function. But, one of the problems is that most of these functions accepts data as simple series. And, sometimes, even statisticians may have to work with other series like frequency or grouped frequency series/tables/distributions. Excel also has some functions in its library that can help to do math but again, you have to tell Excel how to do!
So, here, I have some simple code samples to calculate more easily in Excel.

## Some Central Tendency Measures
This small function code can calculate arithmetic (metot=1 ,default) , geometric (2) and harmonic mean (3) for grouped frequency distribution entered  as  below in Excel spreadsheet. Of course all possible situations must be checked, this function do not have yet.

![Deneme](../lterlemez/VBA_istatistik/grup_seri.PNG)

```vba
Function GOrtalama(veri As Range, Optional metot As Integer = 1)
    'Metot=1 Aritmetik Ortalama ve varsayılan
    'Metot=2 Geometrik Ortalama
    'Metot=3 Harmonik Ortalama
    toplam = 0
    If metot = 1 Then
        For i = 1 To veri.Rows.Count
            toplam = toplam + WorksheetFunction.Average(Range(veri.Cells(i, 1), veri.Cells(i, 2))) * veri.Cells(i, 3)
        Next i
        ort = toplam / WorksheetFunction.Sum(veri.Columns(3))
    ElseIf metot = 2 Then
        For i = 1 To veri.Rows.Count
            toplam = toplam + WorksheetFunction.Log(WorksheetFunction.Average(Range(veri.Cells(i, 1), veri.Cells(i, 2)))) * veri.Cells(i, 3)
        Next i
        ort = WorksheetFunction.Power(10, (toplam / WorksheetFunction.Sum(veri.Columns(3))))
    ElseIf metot = 3 Then
        For i = 1 To veri.Rows.Count
            toplam = toplam + veri.Cells(i, 3) / (WorksheetFunction.Average(Range(veri.Cells(i, 1), veri.Cells(i, 2))))
        Next i
        ort = WorksheetFunction.Sum(veri.Columns(3)) / toplam
    ElseIf metot = 4 Then
        For i = 1 To veri.Rows.Count
            toplam = toplam + WorksheetFunction.Power(WorksheetFunction.Average(Range(veri.Cells(i, 1), veri.Cells(i, 2))), 2) * veri.Cells(i, 3)
        Next i
        ort = Sqr(toplam / WorksheetFunction.Sum(veri.Columns(3)))
    End If
    GOrtalama = ort
End Function
```
