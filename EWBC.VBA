Sub CompareBooks()
' ver. 0.3 îò 28.01.2019
' ver. 0.4 îò 25.02.2019 äîáàâëåíà ïå÷àòü ïðîâåðåííûõ ñòðàíèö, êîëè÷åñòâî ñòîëáöîâ è ñòðîê, èñïðàâëåíû îøèáêè ïî âûõîäó çà ïðåäåëû ñóùåñòâóþùèõ ëèñòîâ
' ver. 0.5 jn 21.03.2019
Dim myName As String, curPage As String, wB As Workbook, curRow As Integer, upBound As Integer
Dim thisWorkBookPages() As String 'ìàññèâ äëÿ ñòðàíèö ýòîãî ôàéëà
Dim otherWorkBookPages() As String 'ìàññèâ äëÿ ñòðàíèö ñðàâíèâàåìîãî ôàéëà
Dim WorkBookPages() As String 'ìàññèâ ëèñòîâ îáùèõ äëÿ äâóõ ôàéëîâ.
Dim HyperLink As String 'ãèïåðññûëêà íà ýëåìåíò
Dim a As Long, b As Long, i As Long, j As Long

'îòêðûòèå ôàéëà, èìÿ äîëæíî îòëè÷àòüñÿ îò óæå îòêðûòîé òàáëèöû
With Application.FileDialog(msoFileDialogFilePicker)
            .Title = "Âûáåðèòå ôàéë äëÿ ñðàâíåíèÿ"
            .Show
                If .SelectedItems.Count = 0 Then Exit Sub
            myName = .SelectedItems(1)
        End With
    Application.ScreenUpdating = False
    Workbooks.Open Filename:=myName: Set wB = Workbooks(ActiveWorkbook.Name)
    
    Windows(ThisWorkbook.Name).Activate: ActiveSheet.Unprotect 'ñíèìàåì çàùèòó, åñëè åñòü
 curRow = 6
 'Ïðîâåðêà ñóùåñòâîâàíèÿ ëèñòà Compare,  â ñëó÷àå åñëè íå ñóùåñòâóåò - äîáàâëÿåòñÿ.
  Dim wsSh As Worksheet
  On Error Resume Next
  Set wsSh = Sheets("Compare")
  If wsSh Is Nothing Then Sheets.Add(, Sheets(Sheets.Count)).Name = "Compare"
 'Çàïîëíåíèå ìàññèâîâ íàçâàíèé ñòðàíèö
 ReDim thisWorkBookPages(ActiveWorkbook.Worksheets.Count)
 ReDim otherWorkBookPages(wB.Worksheets.Count)
 
 For i = 1 To ActiveWorkbook.Worksheets.Count
    thisWorkBookPages(i - 1) = ActiveWorkbook.Worksheets(i).Name
 Next i
 
 For i = 1 To wB.Worksheets.Count
    otherWorkBookPages(i - 1) = wB.Worksheets(i).Name
 Next i
 
 upBound = 0
 For i = 0 To ActiveWorkbook.Worksheets.Count - 1
    For j = 0 To wB.Worksheets.Count - 1
      If thisWorkBookPages(i) = otherWorkBookPages(j) Then
        ReDim Preserve WorkBookPages(upBound)  'Óâåëè÷èâàåì íà åäèíèöó äèíàìè÷åñêèé ìàññèâ
        WorkBookPages(upBound) = thisWorkBookPages(j) 'Äîáàâëÿåì â íåãî íîâîå íàçâàíèå ñòðàíèöû
        upBound = upBound + 1
        Exit For
      End If
    Next j
 Next i
      
      
    Sheets("Compare").Cells.Clear
    'ïðîâåðêà íà íàëè÷èå #ÑÑÛËÊÀ! çàãîëîâîê
   Sheets("Compare").Cells(curRow, 1).Value = "Ïðîâåðêà íà íàëè÷èå ÿ÷ååê òèïà #ÑÑÛËÊÀ"
   Sheets("Compare").Cells(curRow, 2).Value = "Êîëè÷åñòâî ïðîâåðåííûõ ñòîëáöîâ"
   Sheets("Compare").Cells(curRow, 3).Value = "Êîëè÷åñòâî ïðîâåðåííûõ ñòðîê"
   Sheets("Compare").Cells(curRow, 1).Font.Bold = True
   Sheets("Compare").Cells(curRow, 2).Font.Bold = True
   Sheets("Compare").Cells(curRow, 3).Font.Bold = True
   curRow = curRow + 1
    'ïðîâåðêà íà íàëè÷èå #ÑÑÛËÊÀ!
   For j = 0 To UBound(thisWorkBookPages) - 1
        curPage = thisWorkBookPages(j)
        numCol = Sheets(curPage).Cells.SpecialCells(xlLastCell).Column
        numRow = Sheets(curPage).Cells.SpecialCells(xlLastCell).Row
        
        Sheets("Compare").Cells(curRow, 1).Value = "Ïðîâåðêà ñòðàíèöû " + curPage
        Sheets("Compare").Cells(curRow, 2).Value = numCol
        Sheets("Compare").Cells(curRow, 3).Value = numRow
        curRow = curRow + 1
        
        For i = 1 To numRow
        
            For y = 1 To numCol
                If Sheets(curPage).Cells(i, y).Value = "#ÑÑÛËÊÀ!" Then
                                  HyperLink = "=ÃÈÏÅÐÑÑÛËÊÀ(""#'" + curPage + "'!" + Replace(Sheets(curPage).Cells(i, y).Address, "$", "") + """;"""
                                  Sheets("Compare").Cells(curRow, 1).Value = "Ëèñò: " + curPage + " ß÷åéêà: " + Replace(Sheets(curPage).Cells(i, y).Address, "$", "")
                                  Sheets("Compare").Cells(curRow, 1).Font.Bold = True
                                  Sheets("Compare").Cells(curRow, 2).FormulaR1C1Local = HyperLink + Replace(Sheets(curPage).Cells(i, y).Address, "$", "") + """)"
                                  Sheets("Compare").Cells(curRow, 3).Value = "Âûÿâëåíà ïåðåìåííàÿ òèïà #ÑÑÛËÊÀ!"
                                  curRow = curRow + 1
                End If
            Next
        Next
    Next
    
    'çàãîëîâîê ñðàâíåíèÿ
    curRow = curRow + 1
    Sheets("Compare").Cells(curRow, 1).Value = "Ïðîâåðêà íà íåñîâïàäåíèå ÿ÷ååê"
    Sheets("Compare").Cells(curRow, 2).Value = "Òåêóùèé ôàéë"
    Sheets("Compare").Cells(curRow, 3).Value = "Îòêðûâàåìûé ôàéë"
    Sheets("Compare").Cells(curRow, 1).Font.Bold = True
    Sheets("Compare").Cells(curRow, 2).Font.Bold = True
    Sheets("Compare").Cells(curRow, 3).Font.Bold = True
    curRow = curRow + 1
    
    For j = 0 To UBound(WorkBookPages) - 1
        curPage = WorkBookPages(j)
        a = Sheets(curPage).Cells.SpecialCells(xlLastCell).Column
        b = wB.Sheets(curPage).Cells.SpecialCells(xlLastCell).Column
        numCol = MinOf(a, b)
        a = Sheets(curPage).Cells.SpecialCells(xlLastCell).Row
        b = wB.Sheets(curPage).Cells.SpecialCells(xlLastCell).Row
        numRow = MinOf(a, b)
        Sheets("Compare").Cells(curRow, 1).Value = "Ïðîâåðêà ñòðàíèöû " + curPage
        Sheets("Compare").Cells(curRow, 2).Value = numCol
        Sheets("Compare").Cells(curRow, 3).Value = numRow
        Sheets("Compare").Cells(curRow, 1).Font.Bold = True
        Sheets("Compare").Cells(curRow, 2).Font.Bold = True
        Sheets("Compare").Cells(curRow, 3).Font.Bold = True
        curRow = curRow + 1
        
        For i = 1 To numRow
             For y = 1 To numCol
                 If Sheets(curPage).Cells(i, y) <> wB.Sheets(curPage).Cells(i, y) Then
                           Sheets("Compare").Cells(curRow, 1).Value = "Ëèñò: " + curPage + " ß÷åéêà: " + Replace(Sheets(curPage).Cells(i, y).Address, "$", "")
                           HyperLink = "=ÃÈÏÅÐÑÑÛËÊÀ(""#'" + curPage + "'!" + Replace(Sheets(curPage).Cells(i, y).Address, "$", "") + """;"""
                           Sheets("Compare").Cells(curRow, 2).FormulaR1C1Local = HyperLink + Sheets(curPage).Cells(i, y).Value + """)"
                                  
                           HyperLink = "=ÃÈÏÅÐÑÑÛËÊÀ(""#'" + wB.Path + "\[" + wB.Name + "]" + curPage + "'!" + Replace(Sheets(curPage).Cells(i, y).Address, "$", "") + """;"""
                           Sheets("Compare").Cells(curRow, 3).FormulaR1C1Local = HyperLink + wB.Sheets(curPage).Cells(i, y) + """)"
                           curRow = curRow + 1
                 End If
            Next
        Next

    Next
    ActiveSheet.Protect: Application.ScreenUpdating = True
    ActiveSheet.Unprotect
    'wB.Close
End Sub

Private Sub Workbook_Open()

End Sub
Private Function MaxOf(i As Long, j As Long) As Long
 If i > j Then MaxOf = i
 Else: Max = j
 End If
 
End Function

Private Function MinOf(i As Long, j As Long) As Long
 If i < j Then
 MinOf = i
 Else: MinOf = j
 End If
 
End Function

