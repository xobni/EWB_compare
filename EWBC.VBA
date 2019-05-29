Sub CompareBooks()
' ver. 0.3 îò 28.01.2019
' ver. 0.4 îò 25.02.2019 äîáàâëåíà ïå÷àòü ïðîâåðåííûõ ñòðàíèö, êîëè÷åñòâî ñòîëáöîâ è ñòðîê, èñïðàâëåíû îøèáêè ïî âûõîäó çà ïðåäåëû ñóùåñòâóþùèõ ëèñòîâ
' ver. 0.5 21.03.2019
' ver. 0.6 16.04.2019
Dim myName As String, curPage As String, wB As Workbook, curRow As Integer, upBound As Integer
Dim thisWorkBookPages() As String 'ìàññèâ äëÿ ñòðàíèö ýòîãî ôàéëà
Dim otherWorkBookPages() As String 'ìàññèâ äëÿ ñòðàíèö ñðàâíèâàåìîãî ôàéëà
Dim WorkBookPages() As String 'ìàññèâ ëèñòîâ îáùèõ äëÿ äâóõ ôàéëîâ.
Dim HyperLink As String 'ãèïåðññûëêà íà ýëåìåíò
Dim a As Long, b As Long, i As Long, j As Long
Const START_ROW = 5, FIRST_COLUMN = 1, SECOND_COLUMN = 2, THIRD_COLUMN = 3
Dim ProblemFound As Boolean

'opening workbook
With Application.FileDialog(msoFileDialogFilePicker)
                .Title = "Âûáåðèòå ôàéë äëÿ ñðàâíåíèÿ"
                .Show
     If .SelectedItems.Count = 0 Then Exit Sub
     myName = .SelectedItems(1)
End With
Application.ScreenUpdating = False
Workbooks.Open Filename:=myName: Set wB = Workbooks(ActiveWorkbook.Name)
Windows(ThisWorkbook.Name).Activate: ActiveSheet.Unprotect

'ïðîâåðêà íà ñóùåñòâîâàíèå ëèñòà Compare, ñîçäàíèå ïðè íåîáõîäèìîñòè---------------------------------------------
Dim wsSh As Worksheet
On Error Resume Next
Set wsSh = Sheets("Compare")
If wsSh Is Nothing Then Sheets.Add(, Sheets(Sheets.Count)).Name = "Compare"

'Çàïîëíåíèå ìàññèâîâ íàçâàíèÿìè ñòðàíèö, ïðîâåðêà íà íàëè÷èå ëèñòîâ â îáîèõ ôàéëàõ--------------------------------
curRow = START_ROW
Sheets("Compare").Cells.Clear
Call AddStrToJournal(curRow, "", "----------------------------Ïîèñê ëèñòîâ ïðèñóòñòâóþùèõ òîëüêî â îäíîì ôàéëå---------------", "", True)

ReDim thisWorkBookPages(ActiveWorkbook.Worksheets.Count)
ReDim otherWorkBookPages(wB.Worksheets.Count)
 
For i = 1 To ActiveWorkbook.Worksheets.Count
   thisWorkBookPages(i - 1) = ActiveWorkbook.Worksheets(i).Name
   ProblemFound = True
   For j = 1 To wB.Worksheets.Count
     If wB.Worksheets(j).Name = ActiveWorkbook.Worksheets(i).Name Then ProblemFound = False
   Next j
   If ProblemFound Then
      Call AddStrToJournal(curRow, ActiveWorkbook.Worksheets(i).Name, "Åñòü â òåêóùåì ôàéëå è îòñóòñòâóåò â îòêðûâàåìîì", "", True)
   End If
Next i
 
For i = 1 To wB.Worksheets.Count
   otherWorkBookPages(i - 1) = wB.Worksheets(i).Name
   ProblemFound = True
   For j = 1 To ActiveWorkbook.Worksheets.Count
        If ActiveWorkbook.Worksheets(j).Name = wB.Worksheets(i).Name Then ProblemFound = False
   Next j
   If ProblemFound Then
        Call AddStrToJournal(curRow, wB.Worksheets(i).Name, "åñòü â îòêðûâàåìîì ôàéëå è îòñóòñòâóåò â òåêóùåì", "", True)
   End If
Next i
Call AddStrToJournal(curRow, "", "-------------------------Ïîèñê ëèñòîâ ïðèñóòñòâóþùèõ òîëüêî â îäíîì ôàéëå------------------", "", True)

'Çàïîëíÿåì ìàññèâ ñòðàíèö òîëüêî òåìè ñòðàíèöàìè, êîòîðûå åñòü â îáîèõ ôàéëàõ
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
      
   
  
    'ïðîâåðêà íà íàëè÷èå #ÑÑÛËÊÀ! çàãîëîâîê
   Call AddStrToJournal(curRow, "", "-------------------------Ïîèñê ÿ÷ååê òèïà #ÑÑÛËÊÀ------------------", "", True)
   Call AddStrToJournal(curRow, "Íàçâàíèå ëèñòà", "Êîëè÷åñòâî ïðîâåðåííûõ ñòîëáöîâ", "Êîëè÷åñòâî ïðîâåðåííûõ ñòðîê", True)
   ProblemFound = False
   For j = 0 To UBound(thisWorkBookPages) - 1
        curPage = thisWorkBookPages(j)
        numCol = Sheets(curPage).Cells.SpecialCells(xlLastCell).Column
        numRow = Sheets(curPage).Cells.SpecialCells(xlLastCell).Row
        Call AddStrToJournal(curRow, CStr(curPage), CStr(numCol), CStr(numRow), True)
        
        For i = 1 To numRow
        
            For y = 1 To numCol
                If Sheets(curPage).Cells(i, y).Value = "#ÑÑÛËÊÀ!" Then
                                  HyperLink = "=ÃÈÏÅÐÑÑÛËÊÀ(""#'" + curPage + "'!" + Replace(Sheets(curPage).Cells(i, y).Address, "$", "") + """;"""
                                  Sheets("Compare").Cells(curRow, FIRST_COLUMN).Value = "Ëèñò: " + curPage + " ß÷åéêà: " + Replace(Sheets(curPage).Cells(i, y).Address, "$", "")
                                  Sheets("Compare").Cells(curRow, FIRST_COLUMN).Font.Bold = True
                                  Sheets("Compare").Cells(curRow, SECOND_COLUMN).FormulaR1C1Local = HyperLink + Replace(Sheets(curPage).Cells(i, y).Address, "$", "") + """)"
                                  Sheets("Compare").Cells(curRow, 3).Value = "Âûÿâëåíà ïåðåìåííàÿ òèïà #ÑÑÛËÊÀ!"
                                  curRow = curRow + 1
                                  ProblemFound = True
                End If
            Next
        Next
    Next
    If Not ProblemFound Then
        Call AddStrToJournal(curRow, "", "-------------------------Ïîèñê çíà÷åíèé òèïà #ÑÑÛËÊÀ çàêîí÷åí, ïðîáëåì íå âûÿâëåíî--------------------------", "", True)
    End If
    'çàãîëîâîê ñðàâíåíèÿ
    Call AddStrToJournal(curRow, "", "-------------------------Ïîèñê íåñîâïàäàþùèõ ÿ÷ååê â ëèñòàõ----------------------------", "", True)
    Call AddStrToJournal(curRow, "Íàçâàíèå ëèñòà", "Çíà÷åíèå â òåêóùåì ôàéëå", "Çíà÷åíèå â îòêðûâàåìîì ôàéëå", True)
   
    
    For j = 0 To UBound(WorkBookPages) - 1
        curPage = WorkBookPages(j)
        a = Sheets(curPage).Cells.SpecialCells(xlLastCell).Column
        b = wB.Sheets(curPage).Cells.SpecialCells(xlLastCell).Column
        numCol = MinOf(a, b)
        a = Sheets(curPage).Cells.SpecialCells(xlLastCell).Row
        b = wB.Sheets(curPage).Cells.SpecialCells(xlLastCell).Row
        numRow = MinOf(a, b)
        Call AddStrToJournal(curRow, "Ñòðàíèöà " + curPage, " -  ïðîâåðêà", "", True)
          
        
        ProblemFound = False

        For i = 1 To numRow
             For y = 1 To numCol
                 If Sheets(curPage).Cells(i, y) <> wB.Sheets(curPage).Cells(i, y) Then
                           ProblemFound = True
                           Sheets("Compare").Cells(curRow, FIRST_COLUMN).Value = "Ëèñò: " + curPage + " ß÷åéêà: " + Replace(Sheets(curPage).Cells(i, y).Address, "$", "")
                           HyperLink = "=ÃÈÏÅÐÑÑÛËÊÀ(""#'" + curPage + "'!" + Replace(Sheets(curPage).Cells(i, y).Address, "$", "") + """;"""
                           Sheets("Compare").Cells(curRow, SECOND_COLUMN).FormulaR1C1Local = HyperLink + Sheets(curPage).Cells(i, y).Value + """)"
                                 
                           HyperLink = "=ÃÈÏÅÐÑÑÛËÊÀ(""#'" + wB.Path + "\[" + wB.Name + "]" + curPage + "'!" + Replace(Sheets(curPage).Cells(i, y).Address, "$", "") + """;"""
                           Sheets("Compare").Cells(curRow, THIRD_COLUMN).FormulaR1C1Local = HyperLink + wB.Sheets(curPage).Cells(i, y) + """)"
                           curRow = curRow + 1
                 End If
            Next
        Next
        If Not ProblemFound Then
          Call AddStrToJournal(curRow, "Ñòðàíèöà " + curPage, " -  ðàçëè÷èé íå íàéäåíî", "", True)
        End If

    Next
    Call AddStrToJournal(curRow, "", "-------------------------Ïîèñê íåñîâïàäàþùèõ ÿ÷ååê â ëèñòàõ çàêîí÷åí-------------------", "", True)
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
Private Sub AddStrToJournal(ByRef curRow As Integer, First As String, Second As String, Third As String, Optional Bold As Boolean = False)

            Const FIRST_COLUMN = 1, SECOND_COLUMN = 2, THIRD_COLUMN = 3
            Sheets("Compare").Cells(curRow, FIRST_COLUMN).Value = First
            Sheets("Compare").Cells(curRow, SECOND_COLUMN).Value = Second
            Sheets("Compare").Cells(curRow, THIRD_COLUMN).Value = Third
     
            Sheets("Compare").Cells(curRow, FIRST_COLUMN).Font.Bold = Bold
            Sheets("Compare").Cells(curRow, SECOND_COLUMN).Font.Bold = Bold
            Sheets("Compare").Cells(curRow, THIRD_COLUMN).Font.Bold = Bold
            curRow = curRow + 1
End Sub










