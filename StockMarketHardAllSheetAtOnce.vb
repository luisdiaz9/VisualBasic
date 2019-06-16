Sub StockMarket()
Dim B() As String
Dim BD() As String

bda = Worksheets.Count
ReDim BD(bda)
WBN = ActiveWorkbook.Name

For bd1 = 1 To bda
    BD(bd1) = Worksheets(bd1).Name
    For x0 = 1 To 256
        lstRow = Workbooks(WBN).Worksheets(BD(bd1)).Cells(1048576, x0).End(xlUp).Row
        If lstRow > LstRowMax Then
            LstRowMax = lstRow
        End If
    Next x0
    For y0 = 1 To 65536
        LstCol = Workbooks(WBN).Worksheets(bd1).Cells(y0, 256).End(xlToLeft).Column
        If LstCol > LstColMax Then
            LstColMax = LstCol
        End If
    Next y0
Next bd1
'Stop
ReDim B(bda, LstRowMax, LstColMax)
For bd1 = 1 To bda
    For lr1 = 2 To LstRowMax
        For lc1 = 1 To LstColMax
            If Workbooks(WBN).Worksheets(bd1).Cells(1, lc1) = "<date>" And Workbooks(WBN).Worksheets(bd1).Cells(lr1, lc1) <> "" Then
'                Stop
                aDate = Workbooks(WBN).Worksheets(bd1).Cells(lr1, lc1)
                B(bd1, lr1, lc1) = Format(DateValue(Mid(aDate, 1, 4) & "-" & Mid(aDate, 5, 2) & "-" & Mid(aDate, 7, 2)), "dd-mmm-yy")
            Else
                B(bd1, lr1, lc1) = Workbooks(WBN).Worksheets(bd1).Cells(lr1, lc1)
            End If
        Next lc1
    Next lr1
Next bd1
'Stop
Dim DBcMax() As String
Dim C() As String
Dim DbC() As Integer
Dim Es() As String
esa = 6
ReDim Es(esa)
ReDim DbC(bda, LstColMax, 7)
MaxLstRowMax = LstRowMax / 4
ReDim C(bda, esa, LstColMax, MaxLstRowMax, 7)
ReDim DBcMax(bda)

Es(1) = "Ticker"
Es(2) = "Total Stock Volume"
Es(3) = "Min Date"
Es(4) = "Open Price (3)"
Es(5) = "Max Date"
Es(6) = "Close Price (6)"




For j = 0 To bda
    DBcMax(j) = 0
Next j
'Stop
For bd1 = 1 To bda
    For y0 = 2 To LstRowMax
        For vpa = 3 To LstColMax
            If B(bd1, y0, vpa) = "" Then
                B(bd1, y0, vpa) = 0
            End If
        Next vpa
        If y0 = 2360 Then
'        Stop
        End If
        For x1 = 1 To DBcMax(bd1)
            If C(bd1, 1, 1, x1, 1) = B(bd1, y0, 1) Then
'            Stop
             For vp1 = 3 To LstColMax
             If B(bd1, y0, 2) <> "" Then
              If DateValue(C(bd1, 3, 1, x1, vp1)) > DateValue(B(bd1, y0, 2)) Then
              Stop
                C(bd1, 3, 1, x1, vp1) = B(bd1, y0, 2)
                C(bd1, 4, 1, x1, vp1) = B(bd1, y0, 3)
              End If
                If DateValue(C(bd1, 5, 1, x1, vp1)) < DateValue(B(bd1, y0, 2)) And B(bd1, y0, 2) <> "" Then
                    C(bd1, 5, 1, x1, vp1) = B(bd1, y0, 2)
                    C(bd1, 6, 1, x1, vp1) = B(bd1, y0, 6)
                End If
                End If
                
                     C(bd1, 2, 1, x1, vp1) = CDbl(C(bd1, 2, 1, x1, vp1)) + CDbl(B(bd1, y0, vp1))
                Next vp1
                Exit For
            End If
        Next x1
'        Stop
        If x1 > Int(DBcMax(bd1)) Then
'Stop
            For vp1 = 3 To LstColMax
                DbC(bd1, 1, vp1) = DbC(bd1, 1, vp1) + 1
                C(bd1, 1, 1, DbC(bd1, 1, vp1), 1) = B(bd1, y0, 1)
                C(bd1, 2, 1, DbC(bd1, 1, vp1), vp1) = CDbl(B(bd1, y0, vp1))
                C(bd1, 3, 1, DbC(bd1, 1, vp1), vp1) = B(bd1, y0, 2)
                C(bd1, 4, 1, DbC(bd1, 1, vp1), vp1) = CDbl(B(bd1, y0, 3))
                C(bd1, 5, 1, DbC(bd1, 1, vp1), vp1) = B(bd1, y0, 2)
                C(bd1, 6, 1, DbC(bd1, 1, vp1), vp1) = CDbl(B(bd1, y0, 6))
                If DbC(bd1, 1, vp1) > DBcMax(bd1) Then
                    DBcMax(bd1) = DbC(bd1, 1, vp1)
                End If
            Next vp1
        End If
    Next y0
'    Stop
Next bd1
'Stop


For bd1 = 1 To bda
    MaxPer = 0
    MinPer = 0
    MaxVol = 0
    SerMax = 0
    SerMin = 0
    SerMax = 0
    Workbooks(WBN).Worksheets(bd1).Cells(1, 15) = "<ticker>"
    Workbooks(WBN).Worksheets(bd1).Cells(1, 16) = "<Total vol>"
    Workbooks(WBN).Worksheets(bd1).Cells(1, 17) = "<Percent Change>"
    Workbooks(WBN).Worksheets(bd1).Cells(1, 18) = "<Yearly Change>"
    Workbooks(WBN).Worksheets(bd1).Cells(1, 19) = "<Min Date>"
    Workbooks(WBN).Worksheets(bd1).Cells(1, 20) = "<Max Date>"
    For vp1 = LstColMax To LstColMax
        For t1 = 2 To DbC(bd1, 1, vp1) + 1

                 
                Workbooks(WBN).Worksheets(bd1).Cells(t1, 15) = C(bd1, 1, 1, t1 - 1, 1)
                Workbooks(WBN).Worksheets(bd1).Cells(t1, 16) = C(bd1, 2, 1, t1 - 1, vp1) '15 + vp1 - 6
                If Workbooks(WBN).Worksheets(bd1).Cells(t1, 16) > MaxVol Then
                    MaxVol = Workbooks(WBN).Worksheets(bd1).Cells(t1, 16)
                    SerVol = Workbooks(WBN).Worksheets(bd1).Cells(t1, 15)
                End If
                Workbooks(WBN).Worksheets(bd1).Cells(t1, 17) = C(bd1, 4, 1, t1 - 1, vp1) - C(bd1, 6, 1, t1 - 1, vp1)
                If Workbooks(WBN).Worksheets(bd1).Cells(t1, 17) > 0 Then
                    Workbooks(WBN).Worksheets(bd1).Cells(t1, 17).Interior.ColorIndex = 4
                Else
                    Workbooks(WBN).Worksheets(bd1).Cells(t1, 17).Interior.ColorIndex = 3
                End If
                If C(bd1, 4, 1, t1 - 1, vp1) = 0 Then
                    Workbooks(WBN).Worksheets(bd1).Cells(t1, 18) = 0
                Else
                    Workbooks(WBN).Worksheets(bd1).Cells(t1, 18) = (C(bd1, 4, 1, t1 - 1, vp1) - C(bd1, 6, 1, t1 - 1, vp1)) / C(bd1, 4, 1, t1 - 1, vp1)
                End If
                If Workbooks(WBN).Worksheets(bd1).Cells(t1, 18) > MaxPer Then
                    MaxPer = Workbooks(WBN).Worksheets(bd1).Cells(t1, 18)
                    SerMax = Workbooks(WBN).Worksheets(bd1).Cells(t1, 15)
                End If
                If Workbooks(WBN).Worksheets(bd1).Cells(t1, 18) < MinPer Then
                    MinPer = Workbooks(WBN).Worksheets(bd1).Cells(t1, 18)
                    SerMin = Workbooks(WBN).Worksheets(bd1).Cells(t1, 15)
                End If
                Workbooks(WBN).Worksheets(bd1).Cells(t1, 19) = C(bd1, 3, 1, t1 - 1, vp1)
                Workbooks(WBN).Worksheets(bd1).Cells(t1, 20) = C(bd1, 5, 1, t1 - 1, vp1)

        Next t1
    Next vp1
    Workbooks(WBN).Worksheets(bd1).Cells(2, 22) = "<Greatest % Increase>"
    Workbooks(WBN).Worksheets(bd1).Cells(3, 22) = "<Greatest % Decrease>"
    Workbooks(WBN).Worksheets(bd1).Cells(4, 22) = "<Greatest Total Value>"
    Workbooks(WBN).Worksheets(bd1).Cells(1, 23) = "Ticker"
    Workbooks(WBN).Worksheets(bd1).Cells(2, 23) = SerMax
    Workbooks(WBN).Worksheets(bd1).Cells(3, 23) = SerMin
    Workbooks(WBN).Worksheets(bd1).Cells(4, 23) = SerVol
    Workbooks(WBN).Worksheets(bd1).Cells(1, 24) = "Value"
    Workbooks(WBN).Worksheets(bd1).Cells(2, 24) = MaxPer
    Workbooks(WBN).Worksheets(bd1).Cells(3, 24) = MinPer
    Workbooks(WBN).Worksheets(bd1).Cells(4, 24) = MaxVol
Next bd1

'Stop
End Sub