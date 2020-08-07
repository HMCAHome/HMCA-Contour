VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} userform1 
   Caption         =   "Data Processing"
   ClientHeight    =   6675
   ClientLeft      =   6480
   ClientTop       =   2700
   ClientWidth     =   10050
   OleObjectBlob   =   "userform1.dsx":0000
End
Attribute VB_Name = "userform1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public wb As Workbook
Public sh As Worksheet
Public origin_a As Variant
Public origin_b As Variant
Public strafile As String        '数组，提取文件名fileName时使用
Public strafilearr As Variant
Public strafileorigin As Variant

Private Sub CommandButton1_Click()
Dim filenameobj As Variant
Dim afile()
Application.Visible = True
Dim a()
Dim filename As String '打开文件对话框返回的文件名，是一个全路径文件名，其值也可能是False，因此类型为Variant
  Dim FullName As String
    Dim i As Integer
    filenameobj = Application.GetOpenFilename("所有文件 (*.*),*.*")
    '调用Windows打开文件对话框
    If filenameobj <> False Then                       '如果未按“取消”键
        afile = Split(filenameobj, "\")
        filename = afile(UBound(afile))            '数组的最后一个元素为文件名
        FullName = afile(0)
        For i = 1 To UBound(afile)                 '循环合成全路径
            FullName = FullName & "\" & afile(i)
        Next i
    Else
      MsgBox "请选择文件"
    End
    End If
Application.ScreenUpdating = False
Set wb = CreateObject(FullName)
Set sh = wb.Worksheets(1)
Application.ScreenUpdating = False
a = sh.Range(sh.Range("a1"), sh.Range("a1").End(xlToRight)) '需要加入父
a = Application.Transpose(a) '将纵向Select '导入元素符号
ListBox1.List = a '导入列表框
End Sub



Private Sub CommandButton4_Click()
If OptionButton1.Value = True Then
    Call SingleIndex_Click
    MsgBox "The calculation process has been done"
End If

If OptionButton2.Value = True Then
    Call Igeo_Click
    MsgBox "The calculation process has been done", , "Igeo_method"
End If
If OptionButton3.Value = True Then
    Call EF_Click
      MsgBox "The calculation process has been done", , "EF_method"
End If

If OptionButton4.Value = True Then
    Call RI_Click
    MsgBox "The calculation process has been done", , "RI_method"
End If
If OptionButton5.Value = True Then
  Call PLI_Click
  MsgBox "The calculation process has been done", , "PLI_method"
End If
If OptionButton6.Value = True Then
    Call SInemero
      MsgBox "The calculation process has been done", , "singlenemero_method"
End If
End Sub

Private Sub CommandButton5_Click()
userform1.Hide
UserForm2.Show
End Sub




Private Sub CommandButton6_Click()
 Dim path1 As String
 path1 = BrowseForFolder
 outputpath.text = path1
    If path1 = "" Then
       MsgBox " No output path", , " Choosing output path"
       Exit Sub
    Else
       outputpath.text = path1 & "\"
    End If
End Sub

Private Sub openfile_Click()

Dim filenameobj As Variant
Dim i As Integer
Dim afile As Variant
Dim FullName As String
filenameobj = Application.GetOpenFilename
If filenameobj <> False Then                      '如果未按“取消”键
    afile = Split(filenameobj, "\")
    FullName = afile(0)
    For i = 1 To UBound(afile)             '循环合成全路径
    FullName = FullName & "\" & afile(i)
        If i = UBound(afile) - 1 Then outputpath.text = FullName & "\"
    Next i
    filepathname.text = FullName
    FullName = afile(0)
Else
       filepathname.text = " "
       outputpath.text = " "
End If
Application.ScreenUpdating = False
If filepathname.text = " " Then
   MsgBox " No file has been selected"
   Exit Sub
Else
Set wb = CreateObject(filepathname.text)
End If
Set sh = wb.Worksheets(1)
Application.ScreenUpdating = False
origin_a = sh.Range(sh.Range("a1"), sh.Range("a1").End(xlToRight)) '需要加入父
origin_a = Application.Transpose(origin_a)   '将纵向Select '导入元素符号
origin_b = sh.Range(sh.Range(sh.Range("a1"), sh.Range("a65536").End(xlUp)), sh.Range(sh.Range("a1"), sh.Range("a65536").End(xlUp)).Offset(0, 1))
ListBox2.Clear
ListBox1.List = origin_a  '导入列表框
strafile = afile(i - 1)
strafileorigin = Split(strafile, ".")

End Sub

Private Sub SInemero()
Dim temp1(), temp(), tempnemero(), tempelement(), statistic1() '定义装入拟选中元素的数组
Dim i, j, average, max, sum, ll
Dim wbnew As Workbook
temp1 = origin(sh)
temp = nemero_calculation(temp1)
ReDim tempnemero(1 To UBound(temp))
Set wbnew = Workbooks.Add(1)
ReDim tempelement(1 To UBound(temp, 2)) 'temp(1961,9)
For ll = LBound(temp, 2) To UBound(temp, 2)
    tempelement(ll) = temp(1, ll)
Next ll
tempnemero(1) = "Nemero"
wbnew.Sheets(1).Range(Sheets(1).Cells(1, 1), Sheets(1).Cells(UBound(temp, 1), 2)) = origin_b
wbnew.Sheets(1).Range(Sheets(1).Cells(1, 3), Sheets(1).Cells(UBound(temp, 1), UBound(temp, 2))) = temp '需要加入父
sum = 0
For i = 2 To UBound(temp, 1)
    max = temp(i, 1)
    For j = LBound(temp, 2) To UBound(temp, 2)
        sum = sum + temp(i, j)
            If max < temp(i, j) Then max = temp(i, j)
    Next j
    average = sum / ListBox2.ListCount
    tempnemero(i) = ((average ^ 2 + max ^ 2) / 2) ^ 0.5
    sum = 0
Next i
tempnemero = Application.Transpose(tempnemero)
wbnew.Sheets(1).Range(Sheets(1).Cells(1, ListBox2.ListCount + 2), Sheets(1).Cells(UBound(temp, 1), ListBox2.ListCount + 2)).Offset(0, 1) = tempnemero
ActiveWorkbook.Sheets.Add after:=Worksheets(1)
wbnew.Sheets(2).Range(Sheets(2).Cells(1, 2), Sheets(2).Cells(1, UBound(tempelement))) = tempelement
statistic1 = statistic(temp)
wbnew.Sheets(2).Range(Sheets(2).Cells(2, 2), Sheets(2).Cells(UBound(statistic1, 1) + 1, UBound(statistic1, 2) + 1)) = statistic1
wbnew.Sheets(2).Cells(1, 1) = "Variables"
wbnew.Sheets(2).Cells(2, 1) = "Mean"
wbnew.Sheets(2).Cells(3, 1) = "Max"
wbnew.Sheets(2).Cells(4, 1) = "Min"
wbnew.Sheets(2).Cells(5, 1) = "Count"
wbnew.Sheets(2).Cells(6, 1) = "S.D."
wbnew.Sheets(2).Cells(7, 1) = "C.V."
Sheets(2).Name = "Descriptive statistics"
Sheets(1).Activate
strafilearr = Split(strafile, ".", -1)
wbnew.SaveAs outputpath.text & strafileorigin(0) & "_" & "nemero.xlsx"
wbnew.Close
End Sub


Private Sub right_Click() '实现列表框1选择元素数据符号并到第二个列表框中
Dim i As Integer
For i = 0 To ListBox1.ListCount - 1
    If i < ListBox1.ListCount Then
        If ListBox1.Selected(i) Then
            ListBox2.AddItem ListBox1.List(i)
            ListBox1.RemoveItem (i)
             i = i - 1
        End If
    End If
Next i
End Sub
Private Sub allright_Click()
ListBox2.Clear
Dim i As Integer
For i = 0 To ListBox1.ListCount - 1
    ListBox2.AddItem ListBox1.List(i)
Next i
ListBox1.Clear
End Sub
Private Sub allleft_Click()
Dim i As Integer
ListBox1.Clear
For i = 0 To ListBox2.ListCount - 1
    ListBox1.AddItem ListBox2.List(i)
Next i
ListBox2.Clear
End Sub
Private Sub left_Click()
Dim i As Integer
For i = 0 To ListBox2.ListCount - 1
    If i < ListBox2.ListCount Then
        If ListBox2.Selected(i) Then
            ListBox1.AddItem ListBox2.List(i)
            ListBox2.RemoveItem (i)
            i = i - 1
        End If
    End If
Next i
End Sub
Private Sub EF_Click()
Dim temp(), tempelement(), statistic1() '定义装入拟选中元素的数组
Dim BB()
Dim hh As Variant
Dim ii, jj, kk, B_refvalue, ll
Dim wbnew As Workbook
Dim C_refvalue As String
kk = 1
temp = origin(sh) '通用public
C_refvalue = C_ref.text
B_refvalue = B_ref.text
hh = choose_Cref(C_refvalue)
BB = Application.InputBox("input background value for each item", "EF Calculation", "{1,2,3,4}", , , , , 64) '定义地累积指数的参考值
For ii = LBound(temp, 2) To UBound(temp, 2)
    For jj = LBound(temp, 1) + 1 To UBound(temp, 1)
         If kk <= UBound(BB) Then
            temp(jj, ii) = temp(jj, ii) / (BB(kk)) / (hh(jj) / B_refvalue)
         End If
    Next jj
kk = kk + 1
Next ii
Set wbnew = Workbooks.Add(1)
ReDim tempelement(1 To UBound(temp, 2)) 'temp(1961,9)
For ll = LBound(temp, 2) To UBound(temp, 2)
    tempelement(ll) = temp(1, ll)
Next ll
wbnew.Sheets(1).Range(Sheets(1).Cells(1, 1), Sheets(1).Cells(UBound(temp, 1), 2)) = origin_b
wbnew.Sheets(1).Range(Sheets(1).Cells(1, 3), Sheets(1).Cells(UBound(temp, 1), UBound(temp, 2))) = temp '需要加入父
ActiveWorkbook.Sheets.Add after:=Worksheets(1)
wbnew.Sheets(2).Range(Sheets(2).Cells(1, 2), Sheets(2).Cells(1, UBound(tempelement))) = tempelement
statistic1 = statistic(temp)
wbnew.Sheets(2).Range(Sheets(2).Cells(2, 2), Sheets(2).Cells(UBound(statistic1, 1) + 1, UBound(statistic1, 2) + 1)) = statistic1
wbnew.Sheets(2).Cells(1, 1) = "Variables"
wbnew.Sheets(2).Cells(2, 1) = "Mean"
wbnew.Sheets(2).Cells(3, 1) = "Max"
wbnew.Sheets(2).Cells(4, 1) = "Min"
wbnew.Sheets(2).Cells(5, 1) = "Count"
wbnew.Sheets(2).Cells(6, 1) = "S.D."
wbnew.Sheets(2).Cells(7, 1) = "C.V."
Sheets(2).Name = "Descriptive statistics"
Application.AlertBeforeOverwriting = False
wbnew.SaveAs outputpath.text & strafileorigin(0) & "_" & "EF.xlsx"
wbnew.Close
End Sub
Function choose_Cref(a As String)
Dim i, j, m
Dim temp1()
Dim arr()
arr = sh.Range("a1", sh.Range("a65536").End(xlUp).End(xlToRight))
ReDim temp1(1 To UBound(arr, 1))
    For j = LBound(arr, 2) + 2 To UBound(arr, 2)
        If arr(1, j) = a Then
            For i = LBound(arr, 1) To UBound(arr, 1)
                temp1(i) = arr(i, j)
            Next i
        End If
    Next j
choose_Cref = temp1
End Function
Private Sub Igeo_Click()
Dim BB()
Dim temp(), tempelement(), statistic1()
Dim wbnew As Workbook
Dim ii, ll, jj, kk, formular_k '定义地累积指数的k值
kk = 1
temp = origin(sh)
formular_k = TextBox1.text '获得地累积指数的k值
BB = Application.InputBox("input background value for each item", "Igeo calculation", "{1,2,3,4}", , , , , 64) '定义地累积指数的参考值
For ii = LBound(temp, 2) To UBound(temp, 2)
    For jj = LBound(temp, 1) + 1 To UBound(temp, 1)
         If kk <= UBound(BB) Then
            temp(jj, ii) = Log(temp(jj, ii) / (BB(kk) * formular_k)) / Log(2)
         End If
    Next jj
kk = kk + 1
Next ii
Set wbnew = Workbooks.Add(1)
ReDim tempelement(1 To UBound(temp, 2)) 'temp(1961,9)
For ll = LBound(temp, 2) To UBound(temp, 2)
    tempelement(ll) = temp(1, ll)
Next ll
wbnew.Sheets(1).Range(Sheets(1).Cells(1, 1), Sheets(1).Cells(UBound(temp, 1), 2)) = origin_b
wbnew.Sheets(1).Range(Sheets(1).Cells(1, 3), Sheets(1).Cells(UBound(temp, 1), UBound(temp, 2))) = temp '需要加入父
ActiveWorkbook.Sheets.Add after:=Worksheets(1)
wbnew.Sheets(2).Range(Sheets(2).Cells(1, 2), Sheets(2).Cells(1, UBound(tempelement))) = tempelement
statistic1 = statistic(temp)
wbnew.Sheets(2).Range(Sheets(2).Cells(2, 2), Sheets(2).Cells(UBound(statistic1, 1) + 1, UBound(statistic1, 2) + 1)) = statistic1
wbnew.Sheets(2).Cells(1, 1) = "Variables"
wbnew.Sheets(2).Cells(2, 1) = "Mean"
wbnew.Sheets(2).Cells(3, 1) = "Max"
wbnew.Sheets(2).Cells(4, 1) = "Min"
wbnew.Sheets(2).Cells(5, 1) = "Count"
wbnew.Sheets(2).Cells(6, 1) = "S.D."
wbnew.Sheets(2).Cells(7, 1) = "C.V."
Sheets(2).Name = "Descriptive statistics"
Sheets(1).Activate
Application.AlertBeforeOverwriting = False
wbnew.SaveAs outputpath.text & strafileorigin(0) & "_" & "Igeo.xlsx"
wbnew.Close
End Sub
Function statistic(temporigin())
Dim BB()
Dim ii, jj, kk, sum, average, max, min, sum1, std, iii
ReDim BB(1 To 6, 1 To ListBox2.ListCount)
sum = 0
For ii = 1 To ListBox2.ListCount
    max = temporigin(2, ii)
    min = temporigin(2, ii)
    For jj = LBound(temporigin, 1) + 1 To UBound(temporigin, 1)
        sum = temporigin(jj, ii) + sum
        If max <= temporigin(jj, ii) Then max = temporigin(jj, ii)
        If min >= temporigin(jj, ii) Then min = temporigin(jj, ii)
    Next jj
    average = sum / (UBound(temporigin, 1) - 1)
    BB(1, ii) = sum / (UBound(temporigin, 1) - 1)
    sum1 = 0
    For iii = LBound(temporigin, 1) + 1 To UBound(temporigin, 1)
        sum1 = sum1 + (temporigin(iii, ii) - average) * (temporigin(iii, ii) - average)
    Next iii
    BB(5, ii) = Sqr(sum1 / (jj - 2))
    BB(2, ii) = max
    BB(3, ii) = min
    BB(4, ii) = jj - 2
    BB(6, ii) = (Sqr(sum1 / (jj - 3))) / (average)
    sum = 0
Next ii
statistic = BB
End Function
Function origin(wks As Worksheet)
Dim brr()
Dim temp1()
Dim arr()
Dim i, j, m, k
m = 0
arr = wks.Range("a1", wks.Range("a65536").End(xlUp).End(xlToRight))
brr = userform1.ListBox2.List
i = 3
 ReDim temp1(1 To UBound(arr, 1), 1 To UBound(arr, 2)) '重新定义temp1()的行列维
Do While i <= UBound(arr, 2) '吸收重金属元素列号
    For j = LBound(brr, 1) To UBound(brr, 1) 'listbox2的行号
        If arr(1, i) = brr(j, 0) Then '是否有匹配的重金属元素符号
            m = m + 1
                For k = 1 To UBound(arr, 1)
                    temp1(k, m) = arr(k, i)
                Next k
        End If
    Next j
i = i + 1
Loop
origin = temp1
End Function
Private Sub SingleIndex_Click()
Dim temp1(), temp() '定义装入拟选中元素的数组
Dim wbnew As Workbook
Dim tempelement(), statistic1(), ll
temp1 = origin(sh)
temp = nemero_calculation(temp1)
Set wbnew = Workbooks.Add(1)
ReDim tempelement(1 To UBound(temp, 2)) 'temp(1961,9)
For ll = LBound(temp, 2) To UBound(temp, 2)
    tempelement(ll) = temp(1, ll)
Next ll
wbnew.Sheets(1).Range(Sheets(1).Cells(1, 1), Sheets(1).Cells(UBound(temp, 1), 2)) = origin_b
wbnew.Sheets(1).Range(Sheets(1).Cells(1, 3), Sheets(1).Cells(UBound(temp, 1), UBound(temp, 2))) = temp '需要加入父
ActiveWorkbook.Sheets.Add after:=Worksheets(1)
wbnew.Sheets(2).Range(Sheets(2).Cells(1, 2), Sheets(2).Cells(1, UBound(tempelement))) = tempelement
statistic1 = statistic(temp)
wbnew.Sheets(2).Range(Sheets(2).Cells(2, 2), Sheets(2).Cells(UBound(statistic1, 1) + 1, UBound(statistic1, 2) + 1)) = statistic1
wbnew.Sheets(2).Cells(1, 1) = "Variables"
wbnew.Sheets(2).Cells(2, 1) = "Mean"
wbnew.Sheets(2).Cells(3, 1) = "Max"
wbnew.Sheets(2).Cells(4, 1) = "Min"
wbnew.Sheets(2).Cells(5, 1) = "Count"
wbnew.Sheets(2).Cells(6, 1) = "S.D."
wbnew.Sheets(2).Cells(7, 1) = "C.V."
Sheets(2).Name = "Descriptive statistics"
Sheets(1).Activate
strafilearr = Split(strafile, ".", -1)
wbnew.SaveAs outputpath.text & strafileorigin(0) & "_" & "SingleIndex.xlsx"
wbnew.Close
End Sub


Function nemero_calculation(crr())
Dim BB()
Dim ii, jj, kk

kk = 1
BB = Application.InputBox("input background value for each item", "single index calculation", "{1,2,3,4}", , , , , 64)
For ii = LBound(crr, 2) To UBound(crr, 2)
    For jj = LBound(crr, 1) + 1 To UBound(crr, 1)
         If kk <= UBound(BB) Then
            crr(jj, ii) = crr(jj, ii) / BB(kk)
         End If
    Next jj
kk = kk + 1
Next ii
nemero_calculation = crr
End Function
Private Sub PLI_Click()
Dim temp() '定义装入拟选中元素的数组
Dim BB(), PLI()
Dim ii, jj, kk, product, n, m, ll
Dim tempelement(), statistic1()
Dim wbnew As Workbook
temp = origin(sh)
kk = 1
BB = Application.InputBox("input background value for each item", "PLI calculation", "{1,2,3,4}", , , , , 64)
For ii = LBound(temp, 2) To UBound(temp, 2)
    For jj = LBound(temp, 1) + 1 To UBound(temp, 1)
         If kk <= UBound(BB) Then
            temp(jj, ii) = temp(jj, ii) / BB(kk)
         End If
    Next jj
kk = kk + 1
Next ii
ReDim PLI(1 To UBound(temp, 1)) '计算PLI
PLI(1) = "PLI"
product = 1
For m = 1 + 1 To UBound(temp, 1)
     For n = 1 To ListBox2.ListCount
         product = product * temp(m, n)
     Next n
PLI(m) = WorksheetFunction.Power(product, 1 / ListBox2.ListCount)
product = 1 'product重新变成1，不然会使得product值累乘
Next m
PLI = Application.Transpose(PLI)
Set wbnew = Workbooks.Add(1)
ReDim tempelement(1 To UBound(temp, 2)) 'temp(1961,9)
For ll = LBound(temp, 2) To UBound(temp, 2)
    tempelement(ll) = temp(1, ll)
Next ll
wbnew.Sheets(1).Range(Sheets(1).Cells(1, 1), Sheets(1).Cells(UBound(temp, 1), 2)) = origin_b
wbnew.Sheets(1).Range(Sheets(1).Cells(1, 3), Sheets(1).Cells(UBound(temp, 1), UBound(temp, 2))) = temp '需要加入父
wbnew.Sheets(1).Range(Sheets(1).Cells(1, ListBox2.ListCount + 2), Sheets(1).Cells(UBound(temp, 1), ListBox2.ListCount + 2)).Offset(0, 1) = PLI
ActiveWorkbook.Sheets.Add after:=Worksheets(1)
wbnew.Sheets(2).Range(Sheets(2).Cells(1, 2), Sheets(2).Cells(1, UBound(tempelement))) = tempelement
statistic1 = statistic(temp)
wbnew.Sheets(2).Range(Sheets(2).Cells(2, 2), Sheets(2).Cells(UBound(statistic1, 1) + 1, UBound(statistic1, 2) + 1)) = statistic1
wbnew.Sheets(2).Cells(1, 1) = "Variables"
wbnew.Sheets(2).Cells(2, 1) = "Mean"
wbnew.Sheets(2).Cells(3, 1) = "Max"
wbnew.Sheets(2).Cells(4, 1) = "Min"
wbnew.Sheets(2).Cells(5, 1) = "Count"
wbnew.Sheets(2).Cells(6, 1) = "S.D."
wbnew.Sheets(2).Cells(7, 1) = "C.V."
Sheets(2).Name = "Descriptive statistics"
Sheets(1).Activate

wbnew.SaveAs outputpath.text & strafileorigin(0) & "_" & "PLI.xlsx"
wbnew.Close
End Sub
Private Sub RI_Click()
Dim temp() '定义装入拟选中元素的数组
Dim BB(), TR_I(), RI()
Dim ii, jj, kk, sum, n, m, tempelement(), statistic1(), ll
Dim wbnew As Workbook
kk = 1
temp = origin(sh)
BB = Application.InputBox("input background value for each item", "RI Calculation", "{1,2,3,4}", , , , , 64)
TR_I = Application.InputBox("input TR_I", "RI calculation", "{1,2,3,4}", , , , , 64)
For ii = LBound(temp, 2) To UBound(temp, 2)
    For jj = LBound(temp, 1) + 1 To UBound(temp, 1)
         If kk <= UBound(BB) Then
            temp(jj, ii) = temp(jj, ii) / BB(kk) * TR_I(kk)
         End If
    Next jj
kk = kk + 1
Next ii
ReDim RI(1 To UBound(temp, 1)) '计算RI总风险指数
RI(1) = "RI"
For m = 1 + 1 To UBound(temp, 1)
     For n = 1 To UBound(temp, 2)
         sum = sum + temp(m, n)
     Next n
RI(m) = sum
sum = 0 'sum重新清零，不然会使得sum值累加
Next m
RI = Application.Transpose(RI)
Set wbnew = Workbooks.Add(1)
ReDim tempelement(1 To UBound(temp, 2)) 'temp(1961,9)
For ll = LBound(temp, 2) To UBound(temp, 2)
    tempelement(ll) = temp(1, ll)
Next ll
wbnew.Sheets(1).Range(Sheets(1).Cells(1, 1), Sheets(1).Cells(UBound(temp, 1), 2)) = origin_b
wbnew.Sheets(1).Range(Sheets(1).Cells(1, 3), Sheets(1).Cells(UBound(temp, 1), UBound(temp, 2))) = temp '需要加入父
wbnew.Sheets(1).Range(Sheets(1).Cells(1, ListBox2.ListCount + 2), Sheets(1).Cells(UBound(temp, 1), ListBox2.ListCount + 2)).Offset(0, 1) = RI
ActiveWorkbook.Sheets.Add after:=Worksheets(1)
wbnew.Sheets(2).Range(Sheets(2).Cells(1, 2), Sheets(2).Cells(1, UBound(tempelement))) = tempelement
statistic1 = statistic(temp)
wbnew.Sheets(2).Range(Sheets(2).Cells(2, 2), Sheets(2).Cells(UBound(statistic1, 1) + 1, UBound(statistic1, 2) + 1)) = statistic1
wbnew.Sheets(2).Cells(1, 1) = "Variables"
wbnew.Sheets(2).Cells(2, 1) = "Mean"
wbnew.Sheets(2).Cells(3, 1) = "Max"
wbnew.Sheets(2).Cells(4, 1) = "Min"
wbnew.Sheets(2).Cells(5, 1) = "Count"
wbnew.Sheets(2).Cells(6, 1) = "S.D."
wbnew.Sheets(2).Cells(7, 1) = "C.V."
Sheets(2).Name = "Descriptive statistics"
Sheets(1).Activate
wbnew.SaveAs outputpath.text & strafileorigin(0) & "_" & "RI.xlsx"
wbnew.Close
End Sub



