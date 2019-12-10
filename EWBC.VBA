VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ЭтаКнига"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub CompareBooks()
' ver. 0.3 от 28.01.2019
' ver. 0.4 от 25.02.2019 добавлена печать проверенных страниц, количество столбцов и строк, исправлены ошибки по выходу за пределы существующих листов
' ver. 0.5 21.03.2019
' ver. 0.6 16.04.2019
Dim myName As String, curPage As String, wB As Workbook, curRow As Integer, upBound As Integer
Dim thisWorkBookPages() As String 'массив для страниц этого файла
Dim otherWorkBookPages() As String 'массив для страниц сравниваемого файла
Dim WorkBookPages() As String 'массив листов общих для двух файлов.
Dim HyperLink As String 'гиперссылка на элемент
Dim a As Long, b As Long, i As Long, j As Long
Const START_ROW = 5, FIRST_COLUMN = 1, SECOND_COLUMN = 2, THIRD_COLUMN = 3
Dim ProblemFound As Boolean

'opening workbook
With Application.FileDialog(msoFileDialogFilePicker)
                .Title = "Выберите файл для сравнения"
                .Show
     If .SelectedItems.Count = 0 Then Exit Sub
     myName = .SelectedItems(1)
End With
Application.ScreenUpdating = False
Workbooks.Open Filename:=myName: Set wB = Workbooks(ActiveWorkbook.Name)
Windows(ThisWorkbook.Name).Activate: ActiveSheet.Unprotect

'проверка на существование листа Compare, создание при необходимости---------------------------------------------
Dim wsSh As Worksheet
On Error Resume Next
Set wsSh = Sheets("Compare")
If wsSh Is Nothing Then Sheets.Add(, Sheets(Sheets.Count)).Name = "Compare"

'Заполнение массивов названиями страниц, проверка на наличие листов в обоих файлах--------------------------------
curRow = START_ROW
Sheets("Compare").Cells.Clear
Call AddStrToJournal(curRow, "", "----------------------------Поиск листов присутствующих только в одном файле---------------", "", True)

ReDim thisWorkBookPages(ActiveWorkbook.Worksheets.Count)
ReDim otherWorkBookPages(wB.Worksheets.Count)
 
For i = 1 To ActiveWorkbook.Worksheets.Count
   thisWorkBookPages(i - 1) = ActiveWorkbook.Worksheets(i).Name
   ProblemFound = True
   For j = 1 To wB.Worksheets.Count
     If wB.Worksheets(j).Name = ActiveWorkbook.Worksheets(i).Name Then ProblemFound = False
   Next j
   If ProblemFound Then
      Call AddStrToJournal(curRow, ActiveWorkbook.Worksheets(i).Name, "Есть в текущем файле и отсутствует в открываемом", "", True)
   End If
Next i
 
For i = 1 To wB.Worksheets.Count
   otherWorkBookPages(i - 1) = wB.Worksheets(i).Name
   ProblemFound = True
   For j = 1 To ActiveWorkbook.Worksheets.Count
        If ActiveWorkbook.Worksheets(j).Name = wB.Worksheets(i).Name Then ProblemFound = False
   Next j
   If ProblemFound Then
        Call AddStrToJournal(curRow, wB.Worksheets(i).Name, "есть в открываемом файле и отсутствует в текущем", "", True)
   End If
Next i
Call AddStrToJournal(curRow, "", "-------------------------Поиск листов присутствующих только в одном файле------------------", "", True)

'Заполняем массив страниц только теми страницами, которые есть в обоих файлах
 upBound = 0
 For i = 0 To ActiveWorkbook.Worksheets.Count - 1
    For j = 0 To wB.Worksheets.Count - 1
      If thisWorkBookPages(i) = otherWorkBookPages(j) Then
        ReDim Preserve WorkBookPages(upBound)  'Увеличиваем на единицу динамический массив
        WorkBookPages(upBound) = thisWorkBookPages(j) 'Добавляем в него новое название страницы
        upBound = upBound + 1
        Exit For
      End If
    Next j
 Next i
      
   
  
    'проверка на наличие #ССЫЛКА! заголовок
   Call AddStrToJournal(curRow, "", "-------------------------Поиск ячеек типа #ССЫЛКА------------------", "", True)
   Call AddStrToJournal(curRow, "Название листа", "Количество проверенных столбцов", "Количество проверенных строк", True)
   ProblemFound = False
   For j = 0 To UBound(thisWorkBookPages) - 1
        curPage = thisWorkBookPages(j)
        numCol = Sheets(curPage).Cells.SpecialCells(xlLastCell).Column
        numRow = Sheets(curPage).Cells.SpecialCells(xlLastCell).Row
        Call AddStrToJournal(curRow, CStr(curPage), CStr(numCol), CStr(numRow), True)
        
        For i = 1 To numRow
        
            For y = 1 To numCol
                If Sheets(curPage).Cells(i, y).Value = "#ССЫЛКА!" Then
                                  HyperLink = "=ГИПЕРССЫЛКА(""#'" + curPage + "'!" + Replace(Sheets(curPage).Cells(i, y).Address, "$", "") + """;"""
                                  Sheets("Compare").Cells(curRow, FIRST_COLUMN).Value = "Лист: " + curPage + " Ячейка: " + Replace(Sheets(curPage).Cells(i, y).Address, "$", "")
                                  Sheets("Compare").Cells(curRow, FIRST_COLUMN).Font.Bold = True
                                  Sheets("Compare").Cells(curRow, SECOND_COLUMN).FormulaR1C1Local = HyperLink + Replace(Sheets(curPage).Cells(i, y).Address, "$", "") + """)"
                                  Sheets("Compare").Cells(curRow, 3).Value = "Выявлена переменная типа #ССЫЛКА!"
                                  curRow = curRow + 1
                                  ProblemFound = True
                End If
            Next
        Next
    Next
    If Not ProblemFound Then
        Call AddStrToJournal(curRow, "", "-------------------------Поиск значений типа #ССЫЛКА закончен, проблем не выявлено--------------------------", "", True)
    End If
    'заголовок сравнения
    Call AddStrToJournal(curRow, "", "-------------------------Поиск несовпадающих ячеек в листах----------------------------", "", True)
    Call AddStrToJournal(curRow, "Название листа", "Значение в текущем файле", "Значение в открываемом файле", True)
   
    
    For j = 0 To UBound(WorkBookPages) - 1
        curPage = WorkBookPages(j)
        a = Sheets(curPage).Cells.SpecialCells(xlLastCell).Column
        b = wB.Sheets(curPage).Cells.SpecialCells(xlLastCell).Column
        numCol = MinOf(a, b)
        a = Sheets(curPage).Cells.SpecialCells(xlLastCell).Row
        b = wB.Sheets(curPage).Cells.SpecialCells(xlLastCell).Row
        numRow = MinOf(a, b)
        Call AddStrToJournal(curRow, "Страница " + curPage, " -  проверка", "", True)
          
        
        ProblemFound = False

        For i = 1 To numRow
             For y = 1 To numCol
                 If Sheets(curPage).Cells(i, y) <> wB.Sheets(curPage).Cells(i, y) Then
                           ProblemFound = True
                           Sheets("Compare").Cells(curRow, FIRST_COLUMN).Value = "Лист: " + curPage + " Ячейка: " + Replace(Sheets(curPage).Cells(i, y).Address, "$", "")
                           HyperLink = "=ГИПЕРССЫЛКА(""#'" + curPage + "'!" + Replace(Sheets(curPage).Cells(i, y).Address, "$", "") + """;"""
                           Sheets("Compare").Cells(curRow, SECOND_COLUMN).FormulaR1C1Local = HyperLink + Sheets(curPage).Cells(i, y).Value + """)"
                                 
                           HyperLink = "=ГИПЕРССЫЛКА(""#'" + wB.Path + "\[" + wB.Name + "]" + curPage + "'!" + Replace(Sheets(curPage).Cells(i, y).Address, "$", "") + """;"""
                           Sheets("Compare").Cells(curRow, THIRD_COLUMN).FormulaR1C1Local = HyperLink + wB.Sheets(curPage).Cells(i, y) + """)"
                           curRow = curRow + 1
                 End If
            Next
        Next
        If Not ProblemFound Then
          Call AddStrToJournal(curRow, "Страница " + curPage, " -  различий не найдено", "", True)
        End If

    Next
    Call AddStrToJournal(curRow, "", "-------------------------Поиск несовпадающих ячеек в листах закончен-------------------", "", True)
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









