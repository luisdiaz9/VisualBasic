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
esa = 4
ReDim Es(esa)
ReDim DbC(bda, LstColMax, 7)
ReDim C(bda, esa, LstColMax, LstRowMax, 7)
ReDim DBcMax(bda)

Es(1) = "Ticker"
Es(2) = "Yearly Change"
Es(3) = "Percent Change"
Es(4) = "Total Stock Volume"


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
                For vp1 = 3 To LstColMax
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
                C(bd1, 3, 1, DbC(bd1, 1, vp1), vp1) = CDbl(B(bd1, y0, vp1))
                C(bd1, 4, 1, DbC(bd1, 1, vp1), vp1) = CDbl(B(bd1, y0, vp1))
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
    Workbooks(WBN).Worksheets(bd1).Cells(1, 15) = "<ticker>"
    Workbooks(WBN).Worksheets(bd1).Cells(1, 16) = "<vol>"
    For vp1 = LstColMax To LstColMax
        For t1 = 2 To DbC(bd1, 1, vp1)
            For es1 = 2 To 2
                Workbooks(WBN).Worksheets(bd1).Cells(t1, 15) = C(bd1, 1, 1, t1 - 1, 1)
                Workbooks(WBN).Worksheets(bd1).Cells(t1, 16) = C(bd1, es1, 1, t1 - 1, vp1) '15 + vp1 - 6
            Next es1
        Next t1
    Next vp1
Next bd1

'Stop
End Sub