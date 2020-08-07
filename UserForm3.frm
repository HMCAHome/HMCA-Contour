VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "Template Setting"
   ClientHeight    =   9360
   ClientLeft      =   5430
   ClientTop       =   1260
   ClientWidth     =   9345
   OleObjectBlob   =   "UserForm3.dsx":0000
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Majorvalue As Integer
Public minorvalue As Integer

Public BATMAT As Integer
Public BATMIT As Integer
Public LATMAT As Integer
Public LATMIT As Integer
Public RATMAT As Integer
Public RATMIT As Integer
Public TATMAT As Integer
Public TATMIT As Integer
Public symboli As Integer
Public colorscaleTPvalue As Integer

Private Sub AxisLabels_add()
BALB.AddItem "True"
BALB.AddItem "False"
BALI.AddItem "True"
BALI.AddItem "False"
LALB.AddItem "True"
LALB.AddItem "False"
LALI.AddItem "True"
LALI.AddItem "False"
RALB.AddItem "True"
RALB.AddItem "False"
RALI.AddItem "True"
RALI.AddItem "False"
TALB.AddItem "True"
TALB.AddItem "False"
TALI.AddItem "True"
TALI.AddItem "False"
End Sub

Private Sub AxisLabels_addvalue()
BALB.text = "False"
BALI.text = "False"
LALB.text = "False"
LALI.text = "False"
TALB.text = "False"
TALI.text = "False"
RALB.text = "False"
RALI.text = "False"
End Sub

Private Sub Axislabel_angle_value()
BALA.text = 0
RALA.text = 0
LALA.text = 0
TALA.text = 0
Dim i As Integer
For i = 0 To 360 Step 45
    BALA.AddItem i
    LALA.AddItem i
    RALA.AddItem i
    TALA.AddItem i
Next i

End Sub

Private Sub BALA_Change()
BALA.text = BALA.Value
End Sub

Private Sub BALB_Change()
BALB.text = BALB.Value
End Sub

Private Sub BALF_Change()
BALF.text = BALF.Value
End Sub

Private Sub BALF_DropButtonClick()
 For i = 1 To Printer.FontCount - 1
 BALF.AddItem Printer.Fonts(i)
 Next i
End Sub

Private Sub BALI_Change()
BALI.text = BALI.Value
End Sub

Private Sub BALoffset_Change()
BALoffset.text = BALoffset.Value
End Sub

Private Sub BALS_Change()
BALS.text = BALS.Value
End Sub
Private Sub LabelSize_Item_value()
For i = 1 To 20
    BALS.AddItem i
    TALS.AddItem i
    LALS.AddItem i
    RALS.AddItem i
Next i
BALS.text = 9
TALS.text = 9
LALS.text = 9
RALS.text = 9
End Sub



Private Sub BALSB_SpinDown()
If BALoffset.text <= 0 Then
    MsgBox "cannot be lower than 0"
Else
    BALoffset.text = BALoffset.text - 0.01
End If
End Sub
Private Sub BALSB_SpinUp()
BALoffset.text = BALoffset.text + 0.01
End Sub



Private Sub BAT_Click()
If BAT.Value = True Then
End If


End Sub



Private Sub BATA_Change()
BATA.text = BATA.Value
End Sub

Private Sub BATAO_Change()
BATAO.text = BATAO.Value
End Sub

Private Sub BATF_Change()
BATF.text = BATF.Value
End Sub

Sub majorlineminorlines()
BATMAL.AddItem "True"
BATMAL.AddItem "False "
BATMIL.AddItem "True"
BATMIL.AddItem "False "
LATMAL.AddItem "True"
LATMAL.AddItem "False "
LATMIL.AddItem "True"
LATMIL.AddItem "False "
RATMAL.AddItem "True"
RATMAL.AddItem "False "
RATMIL.AddItem "True"
RATMIL.AddItem "False "
TATMAL.AddItem "True"
TATMAL.AddItem "False "
TATMIL.AddItem "True"
TATMIL.AddItem "False "
BATMAL.text = "False"
BATMIL.text = "False"
LATMAL.text = "False"
LATMIL.text = "False"
TATMAL.text = "False"
TATMIL.text = "False"
RATMAL.text = "False"
RATMIL.text = "False"
End Sub

Private Sub BATF_DropButtonClick()
For i = 1 To Printer.FontCount - 1
 BATF.AddItem Printer.Fonts(i)
 Next i
End Sub

Private Sub BATMAL_Change()
BATMAL.text = BATMAL.Value
End Sub

Private Sub BATMAT_Change()

End Sub

Private Sub BATMIL_Change()
BATMIL.text = BATMIL.Value
End Sub


Private Sub BATsize_Change()
BATsize.text = BATsize.Value
End Sub

Private Sub ColorscaleLBTB()
ColorscaleLB.AddItem "True"
ColorscaleLB.AddItem "False"

ColorscaleTB.AddItem "True"
ColorscaleTB.AddItem "False"
ColorscaleTI.AddItem "True"
ColorscaleTI.AddItem "False"
ColorscaleLB.text = "False"

ColorscaleTB.text = "False"
ColorscaleTI.text = "False"
End Sub



Sub axesvtoffset()

For i = 0 To 2 Step 0.1
    BATVO.AddItem i
    BATAO.AddItem i
    LATVO.AddItem i
    LATAO.AddItem i
    RATVO.AddItem i
    RATAO.AddItem i
    TATVO.AddItem i
    TATAO.AddItem i
Next i
    BATVO.text = 0
    BATAO.text = 0
    LATVO.text = 0
    LATAO.text = 0
    RATVO.text = 0
    RATAO.text = 0
    TATVO.text = 0
    TATAO.text = 0
End Sub

Private Sub BATVO_Change()
BATVO.text = BATVO.Value
End Sub

Private Sub BMATL_Change()
BMATL.text = BMATL.Value
End Sub
Private Sub BMPM_Change()
BMPM.text = BMPM.Value
End Sub

Private Sub colorscaleLA_Change()
colorscaleLA.text = colorscaleLA.Value
End Sub

Private Sub ColorscaleLB_Change()
ColorscaleLB.text = ColorscaleLB.Value
End Sub

Private Sub colorscaleLF_Change()
colorscaleLF.text = colorscaleLF.Value
End Sub

Private Sub colorscaleLF_DropButtonClick()
For i = 1 To Printer.FontCount - 1
 colorscaleLF.AddItem Printer.Fonts(i)
 Next i
End Sub

Private Sub ColorscaleLI_Change()
ColorscaleLI.text = ColorscaleLI.Value
End Sub

Private Sub ColorscaleLS_Change()
ColorscaleLS.text = ColorscaleLS.Value
End Sub


Private Sub colorscaleTA_Change()
colorscaleTA.text = colorscaleTA.Value
End Sub

Private Sub ColorscaleTB_Change()
ColorscaleTB.text = ColorscaleTB.Value
End Sub

Private Sub colorscaleTF_Change()
colorscaleTF.text = colorscaleTF.Value
End Sub

Private Sub colorscaleTF_DropButtonClick()
For i = 1 To Printer.FontCount - 1
colorscaleTF.AddItem Printer.Fonts(i)
 Next i
End Sub

Private Sub ColorscaleTI_Change()
ColorscaleTI.text = ColorscaleTI.Value
End Sub

Private Sub ComboBox110_Change()

End Sub

Sub ColorscaleTBTITP()
ColorscaleTB.AddItem "True"
ColorscaleTB.AddItem "False"
ColorscaleTI.AddItem "True"
ColorscaleTI.AddItem "False"
colorscaleTP.AddItem "Top"
colorscaleTP.AddItem "Left"
colorscaleTP.AddItem "Right"
colorscaleTP.AddItem "Bottom"
ColorscaleLB.AddItem "True"
ColorscaleLB.AddItem "False"
ColorscaleLI.AddItem "True"
ColorscaleLI.AddItem "False"
ColorscaleTB.text = "False"
ColorscaleTI.text = "False"
colorscaleTP.text = "Left"
colorscaleTA.text = "0"
colorscaleLA.text = "0"
ColorscaleLS.text = "9"
ColorscaleLB.text = "False"
ColorscaleLI.text = "False"
For i = 1 To 30
    ColorscaleTS.AddItem i
    ColorscaleLS.AddItem i
Next i
ColorscaleTS.text = 9
For j = 0 To 360 Step 45
    colorscaleTA.AddItem j
    colorscaleLA.AddItem j
Next j
End Sub












Private Sub colorscaleTP_Change()
colorscaleTP.text = colorscaleTP.Value
If colorscaleTP.text = "Left" Then
    colorscaleTPvalue = 0
ElseIf colorscaleTP.text = "Right" Then
    colorscaleTPvalue = 1
ElseIf colorscaleTP.text = "Top" Then
    colorscaleTPvalue = 2
Else
    colorscaleTPvalue = 3
End If
End Sub
Private Sub ColorscaleTS_Change()
ColorscaleTS.text = ColorscaleTS.Value
End Sub
Private Sub CommandButton2_Click()
MTprefix.Value = True
MTF.text = "Calibri"
presets.text = "Geology"
MTsuffix.Value = False
MTsize.text = 9
MB.text = "False"
MI.text = "False"
MajorTicktype.text = "out"
MinorTicktype.text = "none"
BATMAL.text = "False"
BMATL.text = 0.07
BMITL.text = 0.05
BATMIL.text = "False"
BMPM.text = 5
BATAW.text = 0.008
LMajorTicktype.text = "out"
Lminorticktype.text = "none"
LATMAL.text = "False"
LMATL.text = 0.07
LMITL.text = 0.05
LATMIL.text = "False"
LMPM.text = 5
LATAW.text = 0.008
RMajorTicktype.text = "none"
Rminorticktype.text = "none"
RATMAL.text = "False"
RMATL.text = 0.07
RMITL.text = 0.05
RATMIL.text = "False"
RMPM.text = 5
RATAW.text = 0.008
TMajorTicktype.text = "none"
Tminorticktype.text = "none"
TATMAL.text = "False"
TMATL.text = 0.07
TMITL.text = 0.05
TATMIL.text = "False"
TMPM.text = 5
TATAW.text = 0.008
BALF.text = "Calibri"
BALS.text = 9
BALB.text = "False"
BALI.text = "False"
BALoffset.text = 0.01
BALA.text = 0
LALF.text = "Calibri"
LALS.text = 9
LALB.text = "False"
LALI.text = "False"
LALoffset.text = 0.01
LALA.text = 0
RALF.text = "Calibri"
RALS.text = 9
RALB.text = "False"
RALI.text = "False"
RALoffset.text = 0.01
TALA.text = 0
TALF.text = "Calibri"
TALS.text = 9
TALB.text = "False"
TALI.text = "False"
TALoffset.text = 0.01
TALA.text = 0
BALvalue.Value = True
LALvalue.Value = True
TALvalue.Value = False
RALvalue.Value = False
TextBox14.text = "Easting"
TextBox13.text = "Northing"
TextBox24.text = ""
TextBox25.text = ""
BATA.text = 0
BATF.text = "Calibri"
BATsize.text = 9
BB.text = "False"
BI.text = "False"
BATVO.text = 0
BATAO.text = 0
LATA.text = 0
LATF.text = "Calibri"
LATsize.text = 9
LB.text = "False"
LI.text = "False"
LATVO.text = 0
LATAO.text = 0
RATA.text = 0
RATF.text = "Calibri"
RATsize.text = 9
RB.text = "False"
RI.text = "False"
RATVO.text = 0
RATAO.text = 0
TATA.text = 0
TATF.text = "Calibri"
TATsize.text = 9
TB.text = "False"
TI.text = "False"
TATVO.text = 0
TATAO.text = 0
CSTP.Value = False
CSTS.Value = False
colorscaleTF.text = "Calibri"
ColorscaleTS.text = 9
ColorscaleTB.text = "False"
ColorscaleTI.text = "False"
colorscaleTP.text = "Left"
colorscaleTA.text = 0
colorscaleLF.text = "Calibri"
ColorscaleLS.text = 9
ColorscaleLB.text = "False"
ColorscaleLI.text = "False"
colorscaleLA.text = 0
ScalebarF.text = "Calibri"
ScalebarS.text = 8
ScalebarI.text = "False"
ScalebarB.text = "False"
ScalebarR.text = 0
Increment.text = 3000
cyclespacing.text = 3000
NumCycles.text = 2
symboltype.text = "symbol32"
NAsymbolsize.text = 0.4
symbolxp.text = ""
symbolyp.text = ""
End Sub

Private Sub CommandButton5_Click()
UserForm2.Show
UserForm3.Hide
End Sub

Private Sub CommandButton6_Click()
Form2.Show
For i = 0 To 9
    Form2.Text3(Index).text = Form2.Text1(Index).text
Next i
End Sub





Private Sub CSTP_Click()
If CSTP.Value = True Then
    CSTPT.BackColor = vbWhite
    CSTPT.Enabled = True
Else
    CSTPT.BackColor = &HE0E0E0
    CSTPT.Enabled = False
End If
End Sub

Private Sub CSTS_Click()
If CSTS.Value = True Then
    CSTST.BackColor = vbWhite
    CSTST.Enabled = True
Else
    CSTST.BackColor = &HE0E0E0
    CSTST.Enabled = False
End If
End Sub




Private Sub cyclespacing_Change()
cyclespacing.text = cyclespacing.Value
End Sub

Private Sub Increment_Change()
Increment.text = Increment.Value
End Sub




Private Sub LALA_Change()
LALA.text = LALA.Value
End Sub

Private Sub LALB_Change()
LALB.text = LALB.Value
End Sub

Private Sub LALF_Change()
LALF.text = LALF.Value
End Sub

Private Sub LALF_DropButtonClick()
For i = 1 To Printer.FontCount - 1
 LALF.AddItem Printer.Fonts(i)
 Next i
End Sub

Private Sub LALI_Change()
LALI.text = LALI.Value
End Sub

Private Sub LALoffset_Change()
LALoffset.text = LALoffset.Value
End Sub

Private Sub LALSB_SpinDown()
If LALoffset.text <= 0 Then
    MsgBox "cannot be lower than 0"
Else
    LALoffset.text = LALoffset.text - 0.01
End If
End Sub
Private Sub LALSB_SpinUp()
LALoffset.text = LALoffset.text + 0.01
End Sub

Private Sub LATAO_Change()
LATAO.text = LATAO.Value
End Sub

Private Sub LATF_Change()
LATF.text = LATF.Value
End Sub

Private Sub LATMAL_Change()
LATMAL.text = LATMAL.Value
End Sub

Private Sub LATMAL_DropButtonClick()
For i = 1 To Printer.FontCount - 1
 LATMAL.AddItem Printer.Fonts(i)
 Next i
End Sub

Private Sub LATMIL_Change()
LATMIL.text = LATMIL.Value
End Sub



Private Sub LATsize_Change()
LATsize.text = LATsize.Value
End Sub


Private Sub LATVO_Change()
LATVO.text = LATVO.Value
End Sub

Private Sub LMajorTicktype_Change()
LMajorTicktype.text = LMajorTicktype.Value
If LMajorTicktype.text = "none" Then
     LATMAT = 1
ElseIf LMajorTicktype.text = "out" Then
     LATMAT = 2
ElseIf LMajorTicktype.text = "in" Then
     LATMAT = 3
Else
     LATMAT = 4
End If

End Sub

Private Sub Lminorticktype_Change()
Lminorticktype.text = Lminorticktype.Value
If Lminorticktype.text = "none" Then
     LATMIT = 1
ElseIf Lminorticktype.text = "out" Then
     LATMIT = 2
ElseIf Lminorticktype.text = "in" Then
     LATMIT = 3
Else
     LATMIT = 4
End If
End Sub

Private Sub MinorTicktype_Change()
MinorTicktype.text = MinorTicktype.Value
If MinorTicktype.text = "none" Then
     BATMIT = 1
ElseIf MinorTicktype.text = "out" Then
     BATMIT = 2
ElseIf MinorTicktype.text = "in" Then
     BATMIT = 3
Else
     BATMIT = 4
End If

End Sub

Private Sub MTF_Change()
MTF.text = MTF.Value
End Sub

Private Sub MTF_DropButtonClick()
 For i = 1 To Printer.FontCount - 1
 MTF.AddItem Printer.Fonts(i)
 Next i
End Sub

Private Sub MTsize_Change()
MTsize.text = MTsize.Value
End Sub

Private Sub NAsymbolsizesb_SpinDown()
If NAsymbolsize.text <= 0 Then
    MsgBox "cannot be lower than 0"
Else
    NAsymbolsize.text = NAsymbolsize.text - 0.01
End If
End Sub
Private Sub NAsymbolsizesb_SpinUp()
NAsymbolsize.text = NAsymbolsize.text + 0.01
End Sub
Private Sub NumCycles_Change()
NumCycles.text = NumCycles.Value
End Sub

Private Sub OptionButton2_Change()
If OptionButton2.Value = True Then
   CommandButton6.Enabled = True
Else
   CommandButton6.Enabled = False
End If
End Sub


Private Sub RALA_Change()
RALA.text = RALA.Value

End Sub

Private Sub RALB_Change()
RALB.text = RALB.Value
End Sub

Private Sub RALF_Change()
RALF.text = RALF.Value
End Sub

Private Sub RALF_DropButtonClick()
For i = 1 To Printer.FontCount - 1
 RALF.AddItem Printer.Fonts(i)
 Next i
End Sub

Private Sub RALI_Change()
RALI.text = RALI.Value
End Sub

Private Sub RALoffset_Change()
RALoffset.text = RALoffset.Value
End Sub

Private Sub RALSB_SpinDown()
If RALoffset.text <= 0 Then
    MsgBox "cannot be lower than 0"
Else
    RALoffset.text = RALoffset.text - 0.01
End If
End Sub
Private Sub RALSB_SpinUp()
RALoffset.text = RALoffset.text + 0.01
End Sub

Private Sub RATAO_Change()
RATAO.text = RATAO.Value
End Sub

Private Sub RATF_Change()
RATF.text = RATF.Value
End Sub

Private Sub RATF_DropButtonClick()
For i = 1 To Printer.FontCount - 1
 RATF.AddItem Printer.Fonts(i)
 Next i
End Sub

Private Sub RATMAL_Change()
RATMAL.text = RATMAL.Value
End Sub

Private Sub RATMIL_Change()
RATMIL.text = RATMIL.Value
End Sub

Private Sub RATsize_Change()
RATsize.text = RATsize.Value
End Sub

Private Sub Scalebarcollection()
ScalebarB.AddItem "False"
ScalebarI.AddItem "False"
ScalebarB.AddItem "True"
ScalebarI.AddItem "True"
cyclespacing.text = "3000"
ScalebarB.text = "False"
ScalebarI.text = "False"
ScalebarR.text = "0"
NumCycles.text = 2
Increment.text = "3000"
For p = 1000 To 10000 Step 1000
cyclespacing.AddItem p
Increment.AddItem p
Next p
For i = 1 To 30
    ScalebarS.AddItem i
Next i
    ScalebarS.text = 8
For j = 0 To 360 Step 45
    ScalebarR.AddItem j
Next j
For k = 1 To 10
    NumCycles.AddItem k
Next k
End Sub

Private Sub RATVO_Change()
RATVO.text = RATVO.Value
End Sub

Private Sub RMajorTicktype_Change()
RMajorTicktype.text = RMajorTicktype.Value
If RMajorTicktype.text = "none" Then
     RATMAT = 1
ElseIf RMajorTicktype.text = "out" Then
     RATMAT = 2
ElseIf RMajorTicktype.text = "in" Then
     RATMAT = 3
Else
     RATMAT = 4
End If

End Sub

Private Sub Rminorticktype_Change()
Rminorticktype.text = Rminorticktype.Value
If Rminorticktype.text = "none" Then
     RATMIT = 1
ElseIf Rminorticktype.text = "out" Then
     RATMIT = 2
ElseIf Rminorticktype.text = "in" Then
     RATMIT = 3
Else
     RATMIT = 4
End If
End Sub



Private Sub ScalebarB_Change()
ScalebarB.text = ScalebarB.Value
End Sub

Private Sub ScalebarF_Change()
ScalebarF.text = ScalebarF.Value
End Sub



Private Sub ScalebarF_DropButtonClick()
For i = 1 To Printer.FontCount - 1
 ScalebarF.AddItem Printer.Fonts(i)
 Next i
End Sub

Private Sub ScalebarI_Change()
ScalebarI.text = ScalebarI.Value
End Sub




Private Sub ScalebarNumCycle_Change()
ScalebarNumCycle.text = ScalebarNumCycle.Value
End Sub

Private Sub ScalebarR_Change()
ScalebarR.text = ScalebarR.Value
End Sub

Private Sub ScalebarS_Change()
ScalebarS.text = ScalebarS.Value
End Sub



Private Sub symboltype_Change()
symboltype.text = symboltype.Value
Select Case symboltype.text
    Case "symbol1"
        symboli = 1
    Case "symbol2"
        symboli = 2
    Case "symbol3"
        symboli = 3
    Case "symbol4"
        symboli = 4
    Case "symbol5"
        symboli = 5
    Case "symbol6"
        symboli = 6
    Case "symbol7"
        symboli = 7
    Case "symbol8"
        symboli = 8
End Select
End Sub
Sub symboltypevalue()
For i = 1 To 8
    symboltype.AddItem "symbol" & i
Next i
symboltype.text = "symbol1"
End Sub


Private Sub symbolxp_Change()
symbolxp.text = symbolxp.Value
End Sub

Private Sub TALA_Change()
TALA.text = TALA.Value
End Sub

Private Sub TALB_Change()
TALB.text = TALB.Value
End Sub

Private Sub TALF_Change()
TALF.text = TALF.Value
End Sub

Private Sub TALF_DropButtonClick()
For i = 1 To Printer.FontCount - 1
 TALF.AddItem Printer.Fonts(i)
 Next i
End Sub

Private Sub TALI_Change()
TALI.text = TALI.Value
End Sub

Private Sub TALoffset_Change()
TALoffset.text = TALoffset.Value
End Sub

Private Sub TALSB_SpinDown()
If TALoffset.text <= 0 Then
    MsgBox "cannot be lower than 0"
Else
    TALoffset.text = TALoffset.text - 0.01
End If
End Sub
Private Sub TALSB_SpinUp()
TALoffset.text = TALoffset.text + 0.01
End Sub

Sub titlesize()
For i = 1 To 30
    MTsize.AddItem i
    BATsize.AddItem i
    LATsize.AddItem i
    RATsize.AddItem i
    TATsize.AddItem i
Next i
    MTsize.text = 9
    BATsize.text = 9
    LATsize.text = 9
    RATsize.text = 9
    TATsize.text = 9
End Sub

Sub labeloffsetvalue()
BALoffset.text = 0.01
RALoffset.text = 0.01
LALoffset.text = 0.01
TALoffset.text = 0.01
End Sub


Private Sub BMATLSB_SpinUp()
BMATL.text = BMATL.text + 0.01
End Sub
Private Sub BMATLSB_Spindown()
If BMATL.text <= 0 Then
    MsgBox "cannot be lower than 0"
Else
    BMATL.text = BMATL.text - 0.01
End If
End Sub

Private Sub LALS_Change()
LALS.text = LALS.Value
End Sub
Private Sub MB_Change()
MB.text = MB.Value
End Sub
Private Sub MI_Change()
MI.text = MI.Value
End Sub
Private Sub BB_Change()
BB.text = BB.Value
End Sub
Private Sub BI_Change()
BI.text = BI.Value
End Sub

Private Sub LB_Change()
LB.text = LB.Value
End Sub
Private Sub LI_Change()
LI.text = LI.Value
End Sub
Private Sub MTprefix_Click()
If MTprefix.Value = True Then
    MTPtext.BackColor = vbWhite
    MTPtext.Enabled = True
Else
    MTPtext.BackColor = &HE0E0E0
    MTPtext.Enabled = False
End If
End Sub
Private Sub MTsuffix_Click()
If MTsuffix.Value = True Then
   MTStext.BackColor = vbWhite
   MTStext.Enabled = True
Else
    MTStext.BackColor = &HE0E0E0
    MTStext.Enabled = False
End If
End Sub
Private Sub RALS_Change()
RALS.text = RALS.Value
End Sub

Private Sub RB_Change()
RB.text = RB.Value
End Sub
Private Sub RI_Change()
RI.text = RI.Value
End Sub

Private Sub TALS_Change()
TALS.text = TALS.Value
End Sub

Private Sub TATAO_Change()
TATAO.text = TATAO.Value
End Sub

Private Sub TATF_Change()
TATF.text = TATF.Value
End Sub

Private Sub TATF_DropButtonClick()
For i = 1 To Printer.FontCount - 1
 TATF.AddItem Printer.Fonts(i)
 Next i
End Sub

Private Sub TATMAL_Change()
TATMAL.text = TATMAL.Value
End Sub

Private Sub TATMIL_Change()
TATMIL.text = TATMIL.Value
End Sub

Private Sub TATsize_Change()
TATsize.text = TATsize.Value
End Sub

Private Sub TATVO_Change()
TATVO.text = TATVO.Value
End Sub

Private Sub TB_Change()
TB.text = TB.Value
End Sub





Private Sub TI_Change()
TI.text = TI.Value
End Sub

Private Sub SpinButton12_SpinDown()

If BMPM.text <= 0 Then
    MsgBox "cannot be lower than 0"
Else
    BMPM.text = BMPM.text - 1
End If
End Sub

Private Sub SpinButton12_SpinUp()
    BMPM.text = BMPM.text + 1
End Sub

Private Sub SpinButton14_SpinDown()
If LMPM.text <= 0 Then
    MsgBox "cannot be lower than 0"
Else
    LMPM.text = LMPM.text - 1
End If
End Sub

Private Sub SpinButton14_SpinUp()
    LMPM.text = LMPM.text + 1
End Sub

Private Sub SpinButton15_SpinDown()
If RMPM.text <= 0 Then
    MsgBox "cannot be lower than 0"
Else
    RMPM.text = RMPM.text - 1
End If
End Sub

Private Sub SpinButton15_SpinUp()
  RMPM.text = RMPM.text + 1
End Sub

Private Sub SpinButton16_SpinDown()
If TMPM.text <= 0 Then
    MsgBox "cannot be lower than 0"
Else
    TMPM.text = TMPM.text - 1
End If
End Sub

Private Sub SpinButton16_SpinUp()
 TMPM.text = TMPM.text + 1
End Sub

Private Sub SpinButton6_Spindown()
If BMITL.text <= 0 Then
    MsgBox "cannot be lower than 0"
Else
    BMITL.text = BMITL.text - 0.01
End If
End Sub
Private Sub SpinButton6_SpinUp()
BMITL.text = BMITL.text + 0.01
End Sub
Private Sub SpinButton5_Spindown()
If LMATL.text <= 0 Then
    MsgBox "cannot be lower than 0"
Else
    LMATL.text = LMATL.text - 0.01
End If
End Sub
Private Sub SpinButton5_SpinUp()
LMATL.text = LMATL.text + 0.01
End Sub
Private Sub SpinButton4_SpinDown()
If LMITL.text <= 0 Then
    MsgBox "cannot be lower than 0"
Else
    LMITL.text = LMITL.text - 0.01
End If
End Sub
Private Sub SpinButton4_SpinUp()
LMITL.text = LMITL.text + 0.01
End Sub

Private Sub SpinButton3_SpinDown()
If RMATL.text <= 0 Then
    MsgBox "cannot be lower than 0"
Else
    RMATL.text = RMATL.text - 0.01
End If
End Sub
Private Sub SpinButton3_SpinUp()
RMATL.text = RMATL.text + 0.01
End Sub
Private Sub SpinButton2_SpinDown()
If RMITL.text <= 0 Then
    MsgBox "cannot be lower than 0"
Else
    RMITL.text = RMITL.text - 0.01
End If
End Sub
Private Sub SpinButton2_SpinUp()
RMITL.text = RMITL.text + 0.01
End Sub
  
Private Sub SpinButton7_SpinUp()
BATAW.text = BATAW.text + 0.001
End Sub
Private Sub SpinButton7_Spindown()
If BATAW.text <= 0 Then
    MsgBox "cannot be lower than 0"
Else
   BATAW.text = BATAW.text - 0.001
End If
End Sub

Private Sub SpinButton8_SpinUp()
LATAW.text = LATAW.text + 0.001
End Sub
Private Sub SpinButton8_Spindown()
If LATAW.text <= 0 Then
    MsgBox "cannot be lower than 0"
Else
   LATAW.text = LATAW.text - 0.001
End If
End Sub

Private Sub SpinButton9_SpinUp()
RATAW.text = RATAW.text + 0.001
End Sub
Private Sub SpinButton9_Spindown()
If RATAW.text <= 0 Then
    MsgBox "cannot be lower than 0"
Else
   RATAW.text = RATAW.text - 0.001
End If
End Sub

Private Sub SpinButton11_SpinUp()
   TATAW.text = TATAW.text + 0.001
End Sub
Private Sub SpinButton11_Spindown()
If TATAW.text <= 0 Then
    MsgBox "cannot be lower than 0"
Else
   TATAW.text = TATAW.text - 0.001
End If
End Sub

Private Sub TMajorTicktype_Change()
TMajorTicktype.text = TMajorTicktype.Value
If TMajorTicktype.text = "none" Then
     TATMAT = 1
ElseIf TMajorTicktype.text = "out" Then
     TATMAT = 2
ElseIf TMajorTicktype.text = "in" Then
     TATMAT = 3
Else
     TATMAT = 4
End If
End Sub

Private Sub TMATLSB_SpinDown()
If TMATL.text <= 0 Then
    MsgBox "cannot be lower than 0"
Else
    TMATL.text = TMATL.text - 0.01
End If
End Sub
Private Sub TMATLSB_SpinUp()
TMATL.text = TMATL.text + 0.01
End Sub

Private Sub Tminorticktype_Change()
Tminorticktype.text = Tminorticktype.Value
If Tminorticktype.text = "none" Then
     TATMIT = 1
ElseIf Tminorticktype.text = "out" Then
     TATMIT = 2
ElseIf Tminorticktype.text = "in" Then
     TATMIT = 3
Else
     TATMIT = 4
End If
End Sub

Private Sub TMITLSB_SpinDown()
If TMITL.text <= 0 Then
    MsgBox "cannot be lower than 0"
Else
    TMITL.text = TMITL.text - 0.01
End If
End Sub
Private Sub TMITLSB_SpinUp()
TMITL.text = TMITL.text + 0.01
End Sub

Private Sub RATSB_SpinDown()
If RATA.text <= 0 Then
   MsgBox "cannot be lower than 0"
   RATA.text = 0
Else
    RATA.text = RATA.text - 1
End If
End Sub
Private Sub RATSB_SpinUp()
RATA.text = RATA.text + 1
End Sub

Private Sub TATSB_SpinUp()
TATA.text = TATA.text + 1
End Sub
Private Sub BATSB_SpinDown()
If BATA.text <= 0 Then
   MsgBox "cannot be lower than 0"
   RATA.text = 0
Else
    BATA.text = BATA.text - 1
End If
End Sub
Private Sub BATSB_SpinUp()
BATA.text = BATA.text + 1
End Sub

Private Sub TATSB_SpinDown()
If TATA.text <= 0 Then
   MsgBox "cannot be lower than 0"
   TATA.text = 0
Else
    TATA.text = TATA.text - 1
End If
End Sub
Private Sub CheckBox1_Click()
If CheckBox1.Value = True Then
    TextBox1.BackColor = vbWhite
    TextBox1.Enabled = True
Else
    TextBox1.BackColor = &HE0E0E0
    TextBox1.Enabled = False
End If
End Sub

Private Sub CommandButton3_Click()
MsgBox Majorvalue
UserForm2.Show
End Sub
Private Sub LATSB_SpinDown()
If LATA.text <= 0 Then
   MsgBox "cannot be lower than 0"
   LATA.text = 0
Else
    LATA.text = LATA.text - 1
End If
End Sub
Private Sub LATSB_SpinUp()
LATA.text = LATA.text + 1
End Sub

Private Sub MajorTicktype_Change()
MajorTicktype.text = MajorTicktype.Value
If MajorTicktype.text = "none" Then
     BATMAT = 1
ElseIf MajorTicktype.text = "out" Then
     BATMAT = 2
ElseIf MajorTicktype.text = "in" Then
     BATMAT = 3
Else
     BATMAT = 4
End If

End Sub
Private Sub presets_Change()
presets.text = presets.Value
End Sub

Sub colorpresets()
presets.text = "Geology"
With presets
    .AddItem "Grayscale"
    .AddItem "Terrian"
    .AddItem "Rainbow"
    .AddItem "Geology"
    .AddItem "Geology2"
    .AddItem "Gravity"
    .AddItem "Gravity2"
    .AddItem "Exploration"
    .AddItem "Exploration2"
    .AddItem "Forecast"
    .AddItem "Soil"
    .AddItem "Sea2"
    .AddItem "LandSea"
    .AddItem "Land"
    .AddItem "Land2"
    .AddItem "Landarid"
    .AddItem "Heat"
    .AddItem "Sea2"
    .AddItem "Forest2"
End With
End Sub
Sub getFont()
    BALF.text = "Calibri"
    LALF.text = "Calibri"
    RALF.text = "Calibri"
    TALF.text = "Calibri"
    MTF.text = "Calibri"
    LATF.text = "Calibri"
    BATF.text = "Calibri"
    RATF.text = "Calibri"
    TATF.text = "Calibri"
    ScalebarF.text = "Calibri"
    colorscaleTF.text = "Calibri"
    colorscaleLF.text = "Calibri"

End Sub

Sub loading_maiticktype()
MajorTicktype.AddItem "none"
MajorTicktype.AddItem "in"
MajorTicktype.AddItem "out"
MajorTicktype.AddItem "cross"
MinorTicktype.AddItem "none"
MinorTicktype.AddItem "in"
MinorTicktype.AddItem "out"
MinorTicktype.AddItem "cross"
LMajorTicktype.AddItem "none"
LMajorTicktype.AddItem "in"
LMajorTicktype.AddItem "out"
LMajorTicktype.AddItem "cross"
Lminorticktype.AddItem "none"
Lminorticktype.AddItem "in"
Lminorticktype.AddItem "out"
Lminorticktype.AddItem "cross"
RMajorTicktype.AddItem "none"
RMajorTicktype.AddItem "in"
RMajorTicktype.AddItem "out"
RMajorTicktype.AddItem "cross"
Rminorticktype.AddItem "none"
Rminorticktype.AddItem "in"
Rminorticktype.AddItem "out"
Rminorticktype.AddItem "cross"
TMajorTicktype.AddItem "none"
TMajorTicktype.AddItem "in"
TMajorTicktype.AddItem "out"
TMajorTicktype.AddItem "cross"
Tminorticktype.AddItem "none"
Tminorticktype.AddItem "in"
Tminorticktype.AddItem "out"
Tminorticktype.AddItem "cross"
End Sub

Sub loading_maiticktypevalue()
MajorTicktype.text = "out"
MinorTicktype.text = "none"
LMajorTicktype.text = "out"
Lminorticktype.text = "none"
RMajorTicktype.text = "none"
Rminorticktype.text = "none"
TMajorTicktype.text = "none"
Tminorticktype.text = "none"
End Sub

Private Sub UserForm_Initialize()

Call CSTS_Click
Call CSTP_Click
Call getFont
Call loading_maiticktype
Call loading_maiticktypevalue
Call AxisLabels_add
Call AxisLabels_addvalue
Call Axislabel_angle_value
Call LabelSize_Item_value
Call labeloffsetvalue
Call titlesize
Call symboltypevalue
Call ColorscaleTBTITP
Call colorpresets
Call Scalebarcollection
Call axesvtoffset
Call majorlineminorlines


BATA.text = 0
LATA.text = 0
RATA.text = 0
TATA.text = 0
BMATL.text = 0.07
BMITL.text = 0.05
LMATL.text = 0.07
LMITL.text = 0.05
RMATL.text = 0.07
RMITL.text = 0.05
TMATL.text = 0.07
TMITL.text = 0.05
NAsymbolsize.text = 0.4
TATAW.text = 0.008
RATAW.text = 0.008
LATAW.text = 0.008
BATAW.text = 0.008

BMPM.text = 5
LMPM.text = 5
TMPM.text = 5
RMPM.text = 5
RB.text = False
TB.text = False
LB.text = False
BB.text = False
MB.text = False
MI.text = False
RI.text = False
TI.text = False
LI.text = False
BI.text = False
BB.AddItem "True"
BB.AddItem "False"
BI.AddItem "True"
BI.AddItem "False"
MB.AddItem "True"
MB.AddItem "False"
MI.AddItem "True"
MI.AddItem "False"
RB.AddItem "True"
RB.AddItem "False"
RI.AddItem "True"
RI.AddItem "False"
TB.AddItem "True"
TB.AddItem "False"
TI.AddItem "True"
TI.AddItem "False"
LB.AddItem "True"
LB.AddItem "False"
LI.AddItem "True"
LI.AddItem "False"

End Sub
