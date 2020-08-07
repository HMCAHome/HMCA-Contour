VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Plotting"
   ClientHeight    =   8970
   ClientLeft      =   5430
   ClientTop       =   1320
   ClientWidth     =   11970
   OleObjectBlob   =   "UserForm2.dsx":0000
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public wb As Workbook
Public sh As Worksheet
Public origin_a As Variant
Public origin_b As Variant
Public strafile As String        '数组，提取文件名fileName时使用
Public strafilearr As Variant
Public X As Variant
Public Y As Variant
Public mv As Integer '定义gridding method value
Public title As String
Public imagetypestr As String
Public iqvalue As Integer
Public destinfile As String
Public filenameobj1 As Variant
Public filenameobj2 As Variant
Public filter_min
Public sss As Integer




Private Sub blncheck_Click()
If blncheck.Value = True Then
   blnposition.Enabled = True
   blnposition.BackColor = vbWhite
Else
   blnposition.Enabled = False
    blnposition.BackColor = &HE0E0E0
End If
End Sub


Private Sub CheckBox1_Click()
If CheckBox1.Value = True Then
    TextBox11.Enabled = True
    TextBox11.BackColor = vbWhite
Else
    TextBox11.Enabled = False
    TextBox11.BackColor = &HE0E0E0
End If
End Sub

Private Sub CheckBox2_Click()
If CheckBox2.Value = True Then
    TextBox12.Enabled = True
    TextBox12.BackColor = vbWhite
Else
    TextBox12.Enabled = False
    TextBox12.BackColor = &HE0E0E0
End If
End Sub

Private Sub ComboBox1_Change()
Select Case ComboBox1.Value
    Case "Inverse Distance to a Power"
        mv = 1
    Case "Kriging"
        mv = 2
    Case "Minimum curvature"
        mv = 3
    Case "Natural Neighbor"
        mv = 4
    Case "Nearest Neighbor"
        mv = 5
    Case "Polynomial Regression"
        mv = 6
    Case "Radial Basis Function"
        mv = 7
    Case "Modified Shepards Method"
        mv = 8
    Case "Triangulation with Linear Interpolation"
        mv = 9
    Case "Moving Average"
        mv = 10
    Case "Data Metrics"
        mv = 11
    Case "Local Polynomial"
        mv = 12
End Select
   
End Sub



Private Sub ComboBox7_DropButtonClick()
 For i = 1 To Printer.FontCount - 1
 ComboBox7.AddItem Printer.Fonts(i)
 Next i
End Sub

Private Sub CommandButton1_Click()
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
        If i = UBound(afile) - 1 Then TextBox10.text = FullName & "\"
    Next i
    TextBox9.text = FullName
    FullName = afile(0)
Else
    TextBox9.text = " "
    TextBox10.text = " "
    Exit Sub
End If
Application.ScreenUpdating = False
Set wb = CreateObject(TextBox9.text)
Set sh = wb.Worksheets(1)
Application.ScreenUpdating = False
origin_a = sh.Range(sh.Range("a1"), sh.Range("a1").End(xlToRight)) '需要加入父
origin_a = Application.Transpose(origin_a)   '将纵向Select '导入元素符号
origin_b = sh.Range(sh.Range(sh.Range("a1"), sh.Range("a65536").End(xlUp)), sh.Range(sh.Range("a1"), sh.Range("a65536").End(xlUp)).Offset(0, 1))
X = sh.Range(sh.Range("a2"), sh.Range("a65536").End(xlUp))
Y = sh.Range(sh.Range("B2"), sh.Range("B65536").End(xlUp))
ListBox2.Clear
ListBox1.List = origin_a  '导入列表框
strafile = afile(i - 1)
xmin.text = Application.min(X)
xmax.text = Application.max(X)
ymin.text = Application.min(Y)
ymax.text = Application.max(Y)
Xnodes.text = 100
xspacing.text = Round((xmax.text - xmin.text) / (Xnodes.text - 1), 4)
yspacing.text = Round((ymax.text - ymin.text) / (Int((ymax.text - ymin.text) / xspacing.text)), 4)
ynodes.text = Int((ymax.text - ymin.text) / xspacing.text) + 1
ComboBox1.text = "kriging"
imagetype.text = "*.jpg"
imagequality.text = "Medium"

End Sub

Private Sub CommandButton10_Click()
Dim i As Integer
Dim origin_aa()
Dim origin_bb()
Dim postfile As Variant
Dim postfullname As String
filenameobj2 = Application.GetOpenFilename
If filenameobj2 <> False Then                      '如果未按“取消”键
    postfile = Split(filenameobj2, "\")
    postfullname = postfile(0)
    For i = 1 To UBound(postfile)             '循环合成全路径
    postfullname = postfullname & "\" & postfile(i)
        If i = UBound(postfile) - 1 Then postposition.text = filenameobj2
    Next i
    postfullname = postfile(0)
Else
   postposition.text = " "
    Exit Sub
End If

End Sub

Private Sub CommandButton2_Click()
 Dim path1 As String
 path1 = BrowseForFolder
 TextBox10.text = path1
    If path1 = "" Then
       MsgBox " No output path"
       Exit Sub
    Else
       TextBox10.text = path1 & "\"
    End If
End Sub

Private Sub CommandButton3_Click()
Dim filter()
If ListBox2.ListCount = 0 Then
   MsgBox "Please choose items from left box to right box "
   ListBox1.SetFocus
   Exit Sub
End If


filter = origin(sh)
Set wbnew = Workbooks.Add(1)
wbnew.Sheets(1).Range(Sheets(1).Cells(1, 1), Sheets(1).Cells(UBound(filter, 1), 2)) = origin_b
wbnew.Sheets(1).Range(Sheets(1).Cells(1, 3), Sheets(1).Cells(UBound(filter, 1), UBound(filter, 2))) = filter '需要加入父

If UserForm3.OptionButton2.Value = True Then
    For p = 1 To ListBox2.ListCount
        filter_min = filter(2, p)
        For q = 2 To UBound(filter, 1)
            If Val(filter(q, p)) <= filter_min Then filter_min = Val(filter(q, p))
        Next q
    Next p
    For p = 1 To ListBox2.ListCount
        filter_max = filter(2, p)
        For q = 2 To UBound(filter, 1)
            If Val(filter(q, p)) >= filter_max Then filter_max = Val(filter(q, p))
        Next q
    Next p
Else

End If
Application.DisplayAlerts = False
wbnew.SaveAs TextBox10.text & "filterbyHMCAcontour.xls"
Application.DisplayAlerts = True
wbnew.Close
If TextBox9.text = " " Then
    MsgBox "Please choose source file"
    Exit Sub
End If
Label14.Width = 0
Label15.Visible = True
Label15.Caption = " It starts to draw, please wait..."
Dim SurferApp, retvalue, doc, plotwindow, contourmapframe, wks, symbol1, shapes1 As Object
Dim contourmap As Object
Dim contourlayer As Object
Dim contourlevels As Object
Dim infile As String
Dim outfile As String
Dim appnum As Integer
Dim text As Object
Dim pageup As Object
Set SurferApp = CreateObject("surfer.application")
infile = TextBox10.text & "filterbyHMCAcontour.xls"
Set wks = SurferApp.Documents.Open(infile)
Dim grid As Object
Dim r As Integer
 Dim Plot As Object
 Set Plot = SurferApp.Documents.Add(1)
For appnum = 0 To ListBox2.ListCount - 1
    title = wks.Cells(1, appnum + 3)
    outfile = TextBox10.text & TextBox11.text & title & TextBox12.text
    outfile1 = TextBox10.text & TextBox11.text & title & TextBox12.text & "transform"
    outfile2 = TextBox10.text & TextBox11.text & title & TextBox12.text & "blank"
    outfile3 = TextBox10.text & TextBox11.text & title & TextBox12.text & "smooth"
    If blncheck.Value = True Then
        If blnposition.text = "" Then
            MsgBox " Please open BLN File"
            CommandButton4.SetFocus
            Exit Sub
        End If
        SurferApp.gridData datafile:=infile, xcol:=1, ycol:=2, zcol:=appnum + 3, algorithm:=mv, showreport:=False, xmin:=xmin.text, _
        xmax:=xmax.text, ymin:=ymin.text, ymax:=ymax.text, OutGrid:=outfile
        SurferApp.gridmath Function:="c=max(a,0)", ingrida:=outfile, outgridc:=outfile1
        SurferApp.GridSplineSmooth InGrid:=outfile1 & ".grd", _
        nRow:=15, nCol:=15, Method:=1, _
        OutGrid:=TextBox10.text & TextBox11.text & title & TextBox12.text & "smooth"
        SurferApp.gridblank InGrid:=outfile3 & ".grd", blankfile:=filenameobj1, OutGrid:=outfile2
        
    Else
        SurferApp.gridData datafile:=infile, xcol:=1, ycol:=2, zcol:=appnum + 3, numcols:=Xnodes.text, numrows:=ynodes.text, algorithm:=mv, showreport:=False, xmin:=xmin.text, _
        xmax:=xmax.text, ymin:=ymin.text, ymax:=ymax.text, OutGrid:=outfile
        
    End If
    Label15.Caption = " It is gridding" + " " + title + "..."
    Set doc = SurferApp.Documents.Add()
    Set plotwindow = doc.Windows(1)
    plotwindow.AutoRedraw = True
    If blncheck.Value = True Then
        Set contourmapframe = doc.Shapes.addcontourmap(outfile2 & ".grd")
    Else
        Set contourmapframe = doc.Shapes.addcontourmap(outfile & ".grd")
    End If
    Set contourmap = contourmapframe.Overlays(1) ' it is also named contourlayer
    Set PageSetup = doc.PageSetup
    If UserForm3.OptionButton2.Value = True Then
        Open UserForm2.TextBox10 + UserForm2.title + "test.lvl" For Output As #1
        Print #1, "LVL2"
        Close #1
        Dim s As String
        For X = 0 To 9
             If Form2.Check1(X).Value = 1 Then
                MsgBox "hello"
                s = Form2.Text1(X).text & " " & "0" & " " & Chr(34) & "Black" & Chr(34) & " " & Chr(34) & "Solid" & Chr(34) & " " & "0.55" & " " _
                & Chr(34) & Form2.Text2(X).text & Chr(34) & " " & Chr(34) & Form2.Text2(X).text & Chr(34) & " " & Chr(34) & "Solid" & Chr(34) & " " & "2"
                Open UserForm2.TextBox10 + title + "test.lvl" For Append As #1
                Print #1, s
                Close #1
             End If
        Next X
    Else
       
    End If
       
  With contourmap
    If UserForm3.colorscaleTotal.Value = True Then
        .showcolorscale = True
    Else
        .showcolorscale = False
    End If
        .LabelTolerance = 1.015
        .LabelLabelDist = 2
        .LabelEdgeDist = 0.5
        .labelformat.Type = 1
        .labelformat.NumDigits = 2
        .OrientLabelsUphill = True
        .LabelFont.Face = "Arial"
        .LabelFont.Size = 5
        .LabelFont.Bold = False
        .FillContours = True
        If UserForm3.OptionButton2.Value = True Then
        .Levels.LoadFile UserForm2.TextBox10 + title + "test.lvl"
        End If
        If UserForm3.OptionButton1.Value = True Then
        .FillForegroundColorMap.LoadPreset (UserForm3.presets)
        .ApplyFillToLevels FirstIndex:=1, NumberToSet:=3, NumberToSkip:=0
        End If
        .Name = title + " "
        .smoothcontours = 4
      
    End With
        Dim colorscale1 As Object
        Set colorscale1 = contourmap.ColorScale
        With colorscale1
            .title = UserForm3.CSTPT.text + " " + title + " " + UserForm3.CSTST.text
            .titlefont.Face = UserForm3.colorscaleTF.text
            .titleangle = UserForm3.colorscaleTA.text
            .titlefont.Bold = UserForm3.ColorscaleTB
            .titlefont.Italic = UserForm3.ColorscaleTI
            .titlefont.Size = UserForm3.ColorscaleTS.text
            .TitlePosition = UserForm3.colorscaleTPvalue  '1 refers to right,2 refers to above, 3 represents below
            .LabelFont.Face = UserForm3.colorscaleLF
            .LabelFont.Size = UserForm3.ColorscaleLS
            .LabelFont.Italic = UserForm3.ColorscaleLI
            .labelangle = UserForm3.colorscaleLA.text
            .firstlabel = 1
            If contourmapframe.Height > .Height Then
               .Top = contourmapframe.Top - (contourmapframe.Height - .Height) / 2
            Else
            .Height = contourmapframe.Height
            .Top = contourmapframe.Top
            '.LabelFormat.NumDigits = 4
            End If
            
     End With
   
    
    Dim axes1, axis1, axis2, axis3, axis4 As Object
    Set axes1 = contourmapframe.Axes
    Set axis1 = axes1(1)
    Set axis2 = axes1(3) '1 refer to bottom, 2 refers to top, 3 refers to left, 4 represents right
    Set axis3 = axes1(2)
    Set axis4 = axes1(4)
    With axis1
        
        .MajorTicktype = UserForm3.BATMAT 'none,2 out,3 in,4 cross
        .MajorTickLength = UserForm3.BMATL.text
        .MinorTicktype = UserForm3.BATMIT
        .MinorTickLength = UserForm3.BMITL.text
        .MinorTicksPerMajor = UserForm3.BMPM.text
        If UserForm3.BALvalue = True Then
            .showlabels = True
        Else
            .showlabels = False
        End If
        .ShowMajorGridLines = UserForm3.BATMAL
        .showminorgridlines = UserForm3.BATMIL
        .AxisLine.Width = UserForm3.BATAW
        .title = UserForm3.TextBox14
        .titlefont.Size = UserForm3.BATsize.text
        .titlefont.Bold = UserForm3.BB.text
        .titlefont.Face = UserForm3.BATF.text
        .titlefont.Italic = UserForm3.BI.text
        .titleangle = UserForm3.BATA.text
        .titleoffset1 = UserForm3.BATAO
        .titleoffset2 = UserForm3.BATVO
        .LabelFont.Size = UserForm3.BALS.text
        .LabelFont.Face = UserForm3.BALF.text
        .LabelFont.Bold = UserForm3.BALB.text
        .LabelFont.Italic = UserForm3.BALI.text
        .LabelOffset = UserForm3.BALoffset.text
        .labelangle = UserForm3.BALA.text
        
        
    End With
    With axis2
        
        .MajorTicktype = UserForm3.LATMAT 'none,out,in,cross
        .MajorTickLength = UserForm3.LMATL.text
        .MinorTicktype = UserForm3.LATMIT 'none
        .MinorTickLength = UserForm3.LMITL.text
        .MinorTicksPerMajor = UserForm3.LMPM.text
        If UserForm3.LALvalue = True Then
            .showlabels = True
        Else
            .showlabels = False
        End If
        .ShowMajorGridLines = UserForm3.LATMAL
        .showminorgridlines = UserForm3.LATMIL
        .AxisLine.Width = UserForm3.LATAW
        .title = UserForm3.TextBox13
        .titlefont.Size = UserForm3.LATsize
        .titlefont.Bold = UserForm3.LB.text
        .titlefont.Face = UserForm3.LATF.text
        .titlefont.Italic = UserForm3.LI.text
        .titleangle = UserForm3.LATA.text
        .titleoffset1 = UserForm3.LATAO
        .titleoffset2 = UserForm3.LATVO
        .LabelFont.Size = UserForm3.LALS.text
        .LabelFont.Face = UserForm3.LALF.text
        .LabelFont.Bold = UserForm3.LALB.text
        .LabelFont.Italic = UserForm3.LALI.text
        .LabelOffset = UserForm3.LALoffset.text
        .labelangle = UserForm3.LALA.text
        
    End With
    
      
    With axis3

    .MajorTicktype = UserForm3.TATMAT 'none,out,in,cross
        .MajorTickLength = UserForm3.TMATL.text
        .MinorTicktype = UserForm3.TATMIT 'none
        .MinorTickLength = UserForm3.TMITL.text
        .MinorTicksPerMajor = UserForm3.TMPM.text
        .showminorgridlines = False
        .ShowMajorGridLines = False
        If UserForm3.TALvalue = True Then
            .showlabels = True
        Else
            .showlabels = False
        End If
        .ShowMajorGridLines = UserForm3.TATMAL
        .showminorgridlines = UserForm3.TATMIL
        .AxisLine.Width = UserForm3.TATAW
        .title = UserForm3.TextBox24
        .titlefont.Size = UserForm3.TATsize.text
        .titlefont.Bold = UserForm3.TB.text
        .titlefont.Face = UserForm3.TATF.text
        .titlefont.Italic = UserForm3.TI.text
        .titleangle = UserForm3.TATA.text
        .titleoffset1 = UserForm3.TATAO
        .titleoffset2 = UserForm3.TATVO
        .LabelFont.Size = UserForm3.TALS.text
        .LabelFont.Face = UserForm3.TALF.text
        .LabelFont.Bold = UserForm3.TALB.text
        .LabelFont.Italic = UserForm3.LALI.text
        .LabelOffset = UserForm3.TALoffset.text
        .labelangle = UserForm3.TALA.text
    End With
    
      
    With axis4
        .MajorTicktype = UserForm3.RATMAT 'none,out,in,cross
        .MajorTickLength = UserForm3.RMATL.text
        .MinorTicktype = UserForm3.RATMIT 'none
        .MinorTickLength = UserForm3.RMITL.text
        .MinorTicksPerMajor = UserForm3.RMPM.text
        .ShowMajorGridLines = UserForm3.TATMAL
        .showminorgridlines = UserForm3.TATMIL
        .AxisLine.Width = UserForm3.RATAW
        If UserForm3.RALvalue = True Then
            .showlabels = True
        Else
            .showlabels = False
        End If
        .title = UserForm3.TextBox25
        .titlefont.Size = UserForm3.RATsize.text
        .titlefont.Bold = UserForm3.RB.text
        .titlefont.Face = UserForm3.RATF.text
        .titlefont.Italic = UserForm3.RI.text
        .titleangle = UserForm3.RATA.text
        .titleoffset1 = UserForm3.RATAO
        .titleoffset2 = UserForm3.RATVO
        .LabelFont.Size = UserForm3.RALS.text
        .LabelFont.Face = UserForm3.RALF.text
        .LabelFont.Bold = UserForm3.RALB.text
        .LabelFont.Italic = UserForm3.RALI.text
        .LabelOffset = UserForm3.RALoffset.text
        .labelangle = UserForm3.RALA.text
    End With
    Set text = doc.Shapes.AddText(X:=PageSetup.Width / 2, Y:=contourmapframe.Top + 0.2, text:=UserForm3.MTPtext + " " + title + " " + UserForm3.MTStext)
    With text
         .Font.Size = UserForm3.MTsize.text
         .Font.Face = UserForm3.MTF.text
         .Font.Bold = UserForm3.MB.text
         .Font.Italic = UserForm3.MI.text
    End With
    text.left = PageSetup.Width / 2 - text.Width / 2
    
    
    Dim Scalebar As Object
    Dim scalebars1 As Object
    Set scalebars1 = contourmapframe.ScalebarS
    Set Scalebar = scalebars1.Add
    With Scalebar
        .labelrotation = UserForm3.ScalebarR.text
        .Top = contourmapframe.Top - contourmapframe.Height - 0.1 '!!!! needs to be changed
        .left = PageSetup.Width / 2
        .NumCycles = UserForm3.NumCycles.text
        .LabelFont.Face = UserForm3.ScalebarF.text
        .LabelFont.Bold = UserForm3.ScalebarB.text
        .LabelFont.Italic = UserForm3.ScalebarI.text
        .LabelFont.Size = UserForm3.ScalebarS.text
        .cyclespacing = UserForm3.cyclespacing
        .Rotation = UserForm3.ScalebarR.text
        .labelIncrement = UserForm3.Increment

    End With
     Scalebar.left = PageSetup.Width / 2 - Scalebar.Width / 2
    If UserForm3.Northarrow.Value = True Then
        Set symbol1 = doc.Shapes.addsymbol(X:=(PageSetup.Width - contourmapframe.Width) / 2, Y:=contourmapframe.Top + UserForm3.NAsymbolsize / 2)  ' ！！！！need to be changed
        symbol1.marker.Size = UserForm3.NAsymbolsize.text
        symbol1.marker.Set = "GSI North Arrows"
        symbol1.marker.Index = UserForm3.symboli
    Else
    End If
    Label15.Caption = " It is composing" + " " + title + " " + "contourmap and postmap..."
    Dim basemapframe As Object
    Dim postmapframe As Object
    Dim pmvalue
    If postmapcheck.Value = True Then
        pmvalue = 1
    Else
        pmvalue = 2
    End If
 Select Case pmvalue
 Case 1
    If postposition.text = "" Then
        MsgBox "Please choose postmapfile"
        postposition.SetFocus
        Exit Sub
    End If
    Set postmapframe = doc.Shapes.AddPostMap(datafilename:=filenameobj2, xcol:=1, ycol:=2, LabCol:=3)
    Dim PostLayer As Object
    Set PostLayer = postmapframe.Overlays(1)
    PostLayer.LabelFont.Face = ComboBox7.text
    PostLayer.LabelFont.Size = symfontsize.text
    PostLayer.symbol.Size = symbolsize.text
    PostLayer.symbol.Index = sss
    If blncheck.Value = True Then
        Set basemapframe = doc.Shapes.addbasemap(importfilename:=filenameobj1)
        doc.Shapes.SelectAll
        doc.Selection.overlaymaps
    Else
        doc.Shapes.SelectAll
    doc.Selection.overlaymaps
    End If

 Case 2

    If blncheck.Value = True Then
        Set basemapframe = doc.Shapes.addbasemap(importfilename:=filenameobj1)
        doc.Shapes.SelectAll
        doc.Selection.overlaymaps
    Else
    
    End If

 End Select
 doc.Export filename:=TextBox10.text + TextBox11.text + title + TextBox12.text + "." + imagetypestr, Options:="Defaults=1,width=int(iqvalue * contourmapframe.width),height= int(iqvalue *contourmapframe.height)       ColorDepth=24"
    plotwindow.Close savechanges:=1, filename:=TextBox10.text + TextBox11.text + title + TextBox12.text + ".srf"
    Label15.Visible = True
    Label15.Caption = " It is generating" + " " + title + " " + "map ..."
    Label14.Width = Label14.Width + Frame6.Width / (ListBox2.ListCount)
    
Next appnum
Label14.Width = Frame6.Width
Label15.Caption = "The batching process has completed !"

SurferApp.Visible = False
Shell "explorer.exe " & " " & TextBox10.text, vbNormalFocus
End Sub

Private Sub CommandButton4_Click()

Dim i As Integer
Dim blnfile As Variant
Dim blnfullname As String
filenameobj1 = Application.GetOpenFilename
If filenameobj1 <> False Then                      '如果未按“取消”键
    blnfile = Split(filenameobj1, "\")
    blnfullname = blnfile(0)
    For i = 1 To UBound(blnfile)             '循环合成全路径
    blnfullname = blnfullname & "\" & blnfile(i)
        If i = UBound(blnfile) - 1 Then blnposition.text = filenameobj1
    Next i
    blnfullname = blnfile(0)
Else
    blnposition.text = " "
    Exit Sub
End If

End Sub

Private Sub CommandButton7_Click()
UserForm2.Hide

UserForm3.Show
End Sub

Private Sub CommandButton8_Click()
userform1.Show
End Sub

Private Sub imagequality_Change()
If imagequality.Value = "Best" Then
   iqvalue = 1200
ElseIf imagequality.Value = "High" Then
   iqvalue = 600
ElseIf imagequality.Value = "Medium" Then
   iqvalue = 300
ElseIf imagequality.Value = "Low" Then
   iqvalue = 100
End If
End Sub

Private Sub imagetype_Change()
Select Case imagetype.Value
    Case "*.jpg"
        imagetypestr = "jpg"
    Case "*.bmp"
        imagetypestr = "bmp"
    Case "*.tif"
        imagetypestr = "tif"
End Select
End Sub




Private Sub postmapcheck_Click()
If postmapcheck.Value = True Then
    postposition.Enabled = True
    postposition.BackColor = vbWhite
Else
    postposition.Enabled = False
    postposition.BackColor = &HE0E0E0
End If
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
ListBox1.Clear
If TextBox9.text = "" Then
   MsgBox "Please choose source file"
   Exit Sub
Else
    ListBox1.List = origin_a  '导入列表框
End If
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

Function origin(wks As Worksheet)
Dim brr()
Dim temp1()
Dim arr()
Dim i, j, m, k
m = 0
arr = wks.Range("a1", wks.Range("a65536").End(xlUp).End(xlToRight))
brr = UserForm2.ListBox2.List
i = 3
 ReDim temp1(1 To UBound(arr, 1), 1 To UBound(arr, 2))
Do While i <= UBound(arr, 2)
    For j = LBound(brr, 1) To UBound(brr, 1)
        If arr(1, i) = brr(j, 0) Then
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



Private Sub symbolsize_Change()
symbolsize.text = symbolsize.Value

End Sub

Private Sub symboltype_Change()
symboltype.text = symboltype.Value
For i = 1 To Len(symboltype.text)
b = Mid(symboltype.text, i, 1)
If IsNumeric(b) Then
sss = Val(Mid(symboltype.text, i))
Exit For
End If
Next

End Sub

Private Sub symfontsize_Change()
symfontsize.text = symfontsize.Value
End Sub

Sub post_ini()
symfontsize.text = 7
For i = 1 To 32
    symfontsize.AddItem i
    symboltype.AddItem "symbol" & i
Next i
For j = 0 To 2 Step 0.1
    symbolsize.AddItem j
Next j
symfontsize.text = 6
symbolsize.text = 0.1
symboltype.text = "symbol32"
ComboBox7.text = "calibri"
End Sub


Private Sub UserForm_Initialize()
Call post_ini
ComboBox1.AddItem "Inverse Distance to a Power"
ComboBox1.AddItem "Kriging"
ComboBox1.AddItem "Minimum curvature"
ComboBox1.AddItem "Natural Neighbor"
ComboBox1.AddItem "Nearest Neighbor"
ComboBox1.AddItem "Polynomial Regression"
ComboBox1.AddItem "Radial Basis Function"
ComboBox1.AddItem "Modified Shepards Method"
ComboBox1.AddItem "Triangulation with Linear Interpolation"
ComboBox1.AddItem "Moving Average"
ComboBox1.AddItem "Data Metrics"
ComboBox1.AddItem "Local Polynomial"
imagetype.AddItem "*.jpg"
imagetype.AddItem "*.bmp"
imagetype.AddItem "*.tif"
imagequality.AddItem "Best"
imagequality.AddItem "High"
imagequality.AddItem "Medium"
imagequality.AddItem "Low"

End Sub



Private Sub xspacing_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii = 46 And Not CBool(InStr(xspacing, ".")) Then Exit Sub
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
If KeyAscii = 8 Then Exit Sub
End Sub

Private Sub ynodes_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyAscii = 46 And Not CBool(InStr(ynodes, ".")) Then Exit Sub
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
If KeyAscii = 8 Then Exit Sub
If ynodes.text <= 1 Then
    yspacing.text = " "
Else
    yspacing.text = Round((ymax.text - ymin.text) / (ynodes.text - 1), 5)
End If
End Sub

Private Sub yspacing_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If yspacing.text <> "" Then
    ynodes.text = Round((ymax.text - ymin.text) / yspacing.text, 0) + 1
Else
    Exit Sub
End If
End Sub


Private Sub Xnodes_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyAscii = 46 And Not CBool(InStr(Xnodes, ".")) Then Exit Sub
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
If KeyAscii = 8 Then Exit Sub
If xspacing.text <> "" Or Xnodes.text > 1 Then
    xspacing.text = Round((xmax.text - xmin.text) / (Xnodes.text - 1), 5)
Else
    Exit Sub
End If
End Sub

Private Sub xspacing_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If xspacing.text <> "" Then
    Xnodes.text = Round((xmax.text - xmin.text) / xspacing.text, 0) + 1
Else
    Exit Sub
End If
End Sub
