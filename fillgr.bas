Attribute VB_Name = "Module4"
Option Explicit
Sub fillgrdates()
Dim main As Worksheet
Set main = Sheets("Main")
Dim grs1 As Worksheet
Set grs1 = Sheets("GRS0")
Dim todaydate As Long
Dim todaymonth As Integer
main.Range("D4:DV390").Clear

Dim today1 As Long
today1 = CLng(Date)
todaydate = CLng(Date)
todaymonth = month(todaydate)

Dim todayyear As Integer
todayyear = year(todaydate)

Dim nr As Integer
nr = WorksheetFunction.CountA(grs1.Range("A:A"))
Dim t1 As Long

Dim i As Integer
Dim grs1index As Integer
Dim grs1index1 As Integer
'todaydate = CLng(DateSerial(todayyear, todaymonth, 1))

todaydate = CLng(main.Range("E2"))

For i = 2 To nr
On Error Resume Next
grs1index = 0
grs1index1 = 0
If CLng(grs1.Range("e" & i)) < todaydate Then 'GR date less then today(e2) which is sept 1"
    If grs1.Range("e" & i) = "" Then
        If Weekday(grs1.Range("D" & i), 2) > 5 Then
                t1 = Application.WorkDay(grs1.Range("D" & i), 1)
            ElseIf Weekday(grs1.Range("D" & i), 2) = 5 Then
                t1 = CLng(grs1.Range("D" & i))
            Else
                t1 = CLng(grs1.Range("D" & i)) + 1
            End If
            grs1index = 0
            grs1index = Application.XMatch(t1, main.Range("2:2"), 0)
            grs1index1 = Application.XMatch(grs1.Range("c" & i), main.Range("b:b"), 0)
            If Not IsError(grs1index) And Not IsError(grs1index1) And grs1index <> 0 Then
                main.Cells(grs1index1, grs1index).Value = "R"
                If t1 < today1 Then
                main.Cells(grs1index1, grs1index).Interior.Color = RGB(0, 0, 255)
                End If
            End If
    Else 'fill in as last gr
        grs1index = Application.XMatch(grs1.Range("c" & i), main.Range("b:b"), 0)
            If Not IsError(grs1index) And grs1index <> 0 Then
                If main.Range("D" & grs1index) = "" Or main.Range("D" & grs1index) <= grs1.Range("e" & i) Then
                    main.Range("d" & grs1index) = grs1.Range("e" & i)
                End If
            End If
    End If
Else 'gr1 >= spet 1
    If grs1.Range("D" & i) <> grs1.Range("e" & i) Then 'if g0 no meet g1 1) >= or 2) <
        If grs1.Range("e" & i) <> "" Then 'g1 is not empty
            t1 = CLng(grs1.Range("e" & i))
            grs1index = 0
            grs1index = Application.XMatch(t1, main.Range("2:2"), 0)
            grs1index1 = Application.XMatch(grs1.Range("c" & i), main.Range("b:b"), 0)
                If Not IsError(grs1index) And Not IsError(grs1index1) And grs1index <> 0 Then
                    If grs1.Range("D" & i) > grs1.Range("e" & i) Then 'early gr1
                        main.Cells(grs1index1, grs1index).Value = "ER"
                        main.Cells(grs1index1, grs1index).Interior.Color = RGB(0, 255, 0)
                    ElseIf grs1.Range("D" & i) = grs1.Range("e" & i) Then 'on time
                        main.Cells(grs1index1, grs1index).Value = "R"
                        main.Cells(grs1index1, grs1index).Interior.Color = RGB(0, 255, 0)
                    Else
                        main.Cells(grs1index1, grs1index).Value = "r" 'late gr1
                        main.Cells(grs1index1, grs1index).Interior.Color = RGB(0, 255, 0)
                        t1 = CLng(grs1.Range("d" & i))
                        grs1index = 0
                        grs1index = Application.XMatch(t1, main.Range("2:2"), 0)
                        grs1index1 = Application.XMatch(grs1.Range("c" & i), main.Range("b:b"), 0)
                            If Not IsError(grs1index) And Not IsError(grs1index1) And grs1index <> 0 Then
                                main.Cells(grs1index1, grs1index).Value = "R"
                                main.Cells(grs1index1, grs1index).Interior.Color = RGB(255, 0, 0)
                            End If
                    End If
                End If
        Else 'gr = ""
            If Weekday(grs1.Range("D" & i), 2) > 5 Or grs1.Range("D" & i) = 45651 Then
                t1 = Application.WorkDay(grs1.Range("D" & i), 1)
            ElseIf Weekday(grs1.Range("D" & i), 2) = 5 Then
                t1 = CLng(grs1.Range("D" & i))
            Else
                t1 = CLng(grs1.Range("D" & i)) + 1
            End If
            grs1index = 0
            grs1index = Application.XMatch(t1, main.Range("2:2"), 0)
            grs1index1 = Application.XMatch(grs1.Range("c" & i), main.Range("b:b"), 0)
            If Not IsError(grs1index) And Not IsError(grs1index1) And grs1index <> 0 Then
                main.Cells(grs1index1, grs1index).Value = "R"
                If t1 < today1 Then
                main.Cells(grs1index1, grs1index).Interior.Color = RGB(0, 0, 255)
                End If
            End If
        End If
    End If
End If
Next i


End Sub
