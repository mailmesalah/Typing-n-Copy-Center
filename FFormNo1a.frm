VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FFormNo1a 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form 1"
   ClientHeight    =   9195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15330
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FFormNo1a.frx":0000
   ScaleHeight     =   9195
   ScaleWidth      =   15330
   StartUpPosition =   1  'CenterOwner
   Tag             =   "a"
   Begin VB.CommandButton CNext 
      Caption         =   "Next"
      Height          =   525
      Left            =   3450
      TabIndex        =   34
      Top             =   7890
      Width           =   1440
   End
   Begin VB.CommandButton CPrint 
      Height          =   500
      Left            =   1725
      Picture         =   "FFormNo1a.frx":1FEC42
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7875
      Width           =   1365
   End
   Begin VB.CommandButton CClose 
      Cancel          =   -1  'True
      Height          =   500
      Left            =   9765
      Picture         =   "FFormNo1a.frx":2010A4
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7950
      Width           =   1365
   End
   Begin MSForms.TextBox TextBox17 
      Height          =   420
      Left            =   4125
      TabIndex        =   33
      Top             =   6315
      Width           =   10485
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      Size            =   "18494;741"
      BorderColor     =   8421504
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TextBox16 
      Height          =   420
      Left            =   4125
      TabIndex        =   32
      Top             =   5880
      Width           =   10485
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      Size            =   "18494;741"
      BorderColor     =   8421504
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label13 
      Height          =   300
      Index           =   13
      Left            =   180
      TabIndex        =   31
      Top             =   5520
      Width           =   3435
      ForeColor       =   -2147483640
      VariousPropertyBits=   8388627
      Caption         =   "In country of domicile"
      Size            =   "6059;529"
      FontName        =   "Sylfaen"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TextBox15 
      Height          =   420
      Left            =   4125
      TabIndex        =   30
      Top             =   5445
      Width           =   10485
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      Size            =   "18494;741"
      BorderColor     =   8421504
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TextBox14 
      Height          =   420
      Left            =   4125
      TabIndex        =   29
      Top             =   5010
      Width           =   10485
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      Size            =   "18494;741"
      BorderColor     =   8421504
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TextBox13 
      Height          =   420
      Left            =   4125
      TabIndex        =   28
      Top             =   4575
      Width           =   10485
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      Size            =   "18494;741"
      BorderColor     =   8421504
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label13 
      Height          =   300
      Index           =   12
      Left            =   180
      TabIndex        =   27
      Top             =   4215
      Width           =   3435
      ForeColor       =   -2147483640
      VariousPropertyBits=   8388627
      Caption         =   "3. Permanent address in India"
      Size            =   "6059;529"
      FontName        =   "Sylfaen"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TextBox12 
      Height          =   420
      Left            =   4125
      TabIndex        =   26
      Top             =   4140
      Width           =   10485
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      Size            =   "18494;741"
      BorderColor     =   8421504
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label13 
      Height          =   300
      Index           =   11
      Left            =   135
      TabIndex        =   25
      Top             =   3660
      Width           =   7785
      ForeColor       =   -2147483640
      VariousPropertyBits=   8388627
      Caption         =   "Visible distinguishing marks, if any"
      Size            =   "13732;529"
      FontName        =   "Sylfaen"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TextBox11 
      Height          =   420
      Left            =   8115
      TabIndex        =   24
      Top             =   3600
      Width           =   3810
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      Size            =   "6720;741"
      BorderColor     =   8421504
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label13 
      Height          =   285
      Index           =   10
      Left            =   10710
      TabIndex        =   23
      Top             =   3210
      Width           =   1440
      ForeColor       =   -2147483640
      VariousPropertyBits=   8388627
      Caption         =   "Color of Hair"
      Size            =   "2540;503"
      FontName        =   "Sylfaen"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TextBox10 
      Height          =   420
      Left            =   12195
      TabIndex        =   22
      Top             =   3165
      Width           =   3105
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      Size            =   "5477;741"
      BorderColor     =   8421504
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label13 
      Height          =   285
      Index           =   9
      Left            =   5430
      TabIndex        =   21
      Top             =   3210
      Width           =   1920
      ForeColor       =   -2147483640
      VariousPropertyBits=   8388627
      Caption         =   "cms Color of eyes"
      Size            =   "3387;503"
      FontName        =   "Sylfaen"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TextBox9 
      Height          =   420
      Left            =   7410
      TabIndex        =   20
      Top             =   3165
      Width           =   3030
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      Size            =   "5345;741"
      BorderColor     =   8421504
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label13 
      Height          =   270
      Index           =   8
      Left            =   0
      TabIndex        =   19
      Top             =   3225
      Width           =   1380
      ForeColor       =   -2147483640
      VariousPropertyBits=   8388627
      Caption         =   "Height"
      Size            =   "2434;476"
      FontName        =   "Sylfaen"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TextBox8 
      Height          =   420
      Left            =   1470
      TabIndex        =   18
      Top             =   3165
      Width           =   3810
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      Size            =   "6720;741"
      BorderColor     =   8421504
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label13 
      Height          =   300
      Index           =   7
      Left            =   1185
      TabIndex        =   17
      Top             =   2790
      Width           =   1515
      ForeColor       =   -2147483640
      VariousPropertyBits=   8388627
      Caption         =   "Country"
      Size            =   "2672;529"
      FontName        =   "Sylfaen"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TextBox7 
      Height          =   420
      Left            =   2730
      TabIndex        =   16
      Top             =   2730
      Width           =   3810
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      Size            =   "6720;741"
      BorderColor     =   8421504
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label13 
      Height          =   285
      Index           =   6
      Left            =   6690
      TabIndex        =   15
      Top             =   2340
      Width           =   1575
      ForeColor       =   -2147483640
      VariousPropertyBits=   8388627
      Caption         =   "Place of birth"
      Size            =   "2778;503"
      FontName        =   "Sylfaen"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TextBox6 
      Height          =   420
      Left            =   8880
      TabIndex        =   14
      Top             =   2295
      Width           =   3810
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      Size            =   "6720;741"
      BorderColor     =   8421504
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label13 
      Height          =   270
      Index           =   5
      Left            =   765
      TabIndex        =   13
      Top             =   2355
      Width           =   1800
      ForeColor       =   -2147483640
      VariousPropertyBits=   8388627
      Caption         =   "2. Date of birth"
      Size            =   "3175;476"
      FontName        =   "Sylfaen"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TextBox5 
      Height          =   420
      Left            =   2730
      TabIndex        =   12
      Top             =   2295
      Width           =   3810
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      Size            =   "6720;741"
      BorderColor     =   8421504
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label13 
      Height          =   300
      Index           =   4
      Left            =   180
      TabIndex        =   11
      Top             =   1905
      Width           =   7890
      ForeColor       =   -2147483640
      VariousPropertyBits=   8388627
      Caption         =   "Maiden name, if applicant is a married woman"
      Size            =   "13917;529"
      FontName        =   "Sylfaen"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TextBox4 
      Height          =   420
      Left            =   8160
      TabIndex        =   10
      Top             =   1860
      Width           =   3810
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      Size            =   "6720;741"
      BorderColor     =   8421504
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label13 
      Height          =   300
      Index           =   3
      Left            =   180
      TabIndex        =   9
      Top             =   1470
      Width           =   7890
      ForeColor       =   -2147483640
      VariousPropertyBits=   8388627
      Caption         =   "Has applicant ever changed his/her name? Is so, give previous name in full"
      Size            =   "13917;529"
      FontName        =   "Sylfaen"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TextBox3 
      Height          =   420
      Left            =   8160
      TabIndex        =   8
      Top             =   1425
      Width           =   3810
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      Size            =   "6720;741"
      BorderColor     =   8421504
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label13 
      Height          =   300
      Index           =   2
      Left            =   195
      TabIndex        =   7
      Top             =   900
      Width           =   1515
      ForeColor       =   -2147483640
      VariousPropertyBits=   8388627
      Caption         =   "Aliases, if any"
      Size            =   "2672;529"
      FontName        =   "Sylfaen"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TextBox2 
      Height          =   420
      Left            =   1740
      TabIndex        =   6
      Top             =   825
      Width           =   3810
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      Size            =   "6720;741"
      BorderColor     =   8421504
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label13 
      Height          =   285
      Index           =   1
      Left            =   5625
      TabIndex        =   5
      Top             =   465
      Width           =   1380
      ForeColor       =   -2147483640
      VariousPropertyBits=   8388627
      Caption         =   "Surname"
      Size            =   "2434;503"
      FontName        =   "Sylfaen"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TextBox1 
      Height          =   420
      Left            =   7095
      TabIndex        =   4
      Top             =   420
      Width           =   3810
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      Size            =   "6720;741"
      BorderColor     =   8421504
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TTransactionNo 
      Height          =   420
      Left            =   1665
      TabIndex        =   0
      Top             =   420
      Width           =   3810
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      Size            =   "6720;741"
      BorderColor     =   8421504
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label13 
      Height          =   270
      Index           =   0
      Left            =   195
      TabIndex        =   1
      Top             =   480
      Width           =   1380
      ForeColor       =   -2147483640
      VariousPropertyBits=   8388627
      Caption         =   "1. Full Name"
      Size            =   "2434;476"
      FontName        =   "Sylfaen"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
End
Attribute VB_Name = "FFormNo1a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CClose_Click()
    Unload Me
End Sub

'Private Sub printSales()
'
'On Error GoTo GoOut
'    Dim i, x, y As Double
'    Dim ITaxAmount, IGrossvalue, INetValue, IDiscount, IQty, ITotalValue, IRate As Double
'
'    ITaxAmount = 0
'    IGrossvalue = 0
'    INetValue = 0
'    IDiscount = 0
'    IQty = 0
'    ITotalValue = 0
'    IRate = 0
'
'    i = 0
'    x = 500
'
'    y = NewPage + 400
'
'    While (i < MGrid.Rows)
'
'        Printer.FontSize = 9
'        Printer.FontBold = False
'
'        x = 550
'        Printer.CurrentX = x
'        Printer.CurrentY = y
'        Printer.Print Trim(MGrid.TextMatrix(i, gSerialNo))
'
'        x = 1100
'        Printer.CurrentX = x
'        Printer.CurrentY = y
'        Printer.Print Trim(MGrid.TextMatrix(i, gBillingName))
'
'        x = 5250
'        Printer.CurrentX = x
'        Printer.CurrentY = y
'        Printer.Print Format(MGrid.TextMatrix(i, gTax), "0.00")
'
'        x = 5950
'        Printer.CurrentX = x
'        Printer.CurrentY = y
'        Printer.Print Format(MGrid.TextMatrix(i, gSaleRate), "0.00")
'
'        x = 7250
'        Printer.CurrentX = x
'        Printer.CurrentY = y
'        Printer.Print Trim(MGrid.TextMatrix(i, gQuantity))
'
''        x = 7250
''        Printer.CurrentX = x
''        Printer.CurrentY = y
''        Printer.Print Format(MGrid.TextMatrix(i, gGrossValue), "0.00")
'
'        x = 7950
'        Printer.CurrentX = x
'        Printer.CurrentY = y
'        Printer.Print Format(MGrid.TextMatrix(i, gItemDiscount), "0.00")
'
''        x = 9100
''        Printer.CurrentX = x
''        Printer.CurrentY = y
''        Printer.Print Format(MGrid.TextMatrix(i, gNetValue), "0.00")
'
'         x = 8950
'        Printer.CurrentX = x
'        Printer.CurrentY = y
'        Printer.Print Format(MGrid.TextMatrix(i, gTaxAmount), "0.00")
'
'         x = 9950
'        Printer.CurrentX = x
'        Printer.CurrentY = y
'        Printer.Print Format(MGrid.TextMatrix(i, gToTalAmount), "0.00")
'
'        ITaxAmount = ITaxAmount + Val(MGrid.TextMatrix(i, gTaxAmount))
'        IDiscount = IDiscount + Val(MGrid.TextMatrix(i, gItemDiscount))
'        IGrossvalue = IGrossvalue + Val(MGrid.TextMatrix(i, gGrossValue))
'        INetValue = INetValue + Val(MGrid.TextMatrix(i, gNetValue))
'        IQty = IQty + Val(MGrid.TextMatrix(i, gQuantity))
'        IRate = IRate + Val(MGrid.TextMatrix(i, gSaleRate))
'
''        If Val(MGrid.TextMatrix(i, gTax)) = 4 Then
''            Check = True
''            TGrossvalue = Tgross + Val(MGrid.TextMatrix(i, gGrossValue))
''            Taxamt = Tex + Val(MGrid.TextMatrix(i, gTaxAmount))
''            TNetamt = Tnet + Val(MGrid.TextMatrix(i, gTotalAmount))
''        Else
''            Check1 = True
''            TGrossvalue1 = Tgross1 + Val(MGrid.TextMatrix(i, gGrossValue))
''            Taxamt1 = Tex1 + Val(MGrid.TextMatrix(i, gTaxAmount))
''            TNetamt1 = Tnet1 + Val(MGrid.TextMatrix(i, gTotalAmount))
''        End If
'
'        i = i + 1
'        y = y + 300
'        If (y > 13000) Then
'            y = NewPage + 400
'        End If
'    Wend
'
'        y = 9800
'        x = 1700
'        Printer.CurrentX = x
'        Printer.CurrentY = y
'        Printer.Print "TOTAL"
'
''        x = 5550
''        Printer.CurrentX = x
''        Printer.CurrentY = y
''        Printer.Print Format(IRate, "0.00")
'
'        x = 7250
'        Printer.CurrentX = x
'        Printer.CurrentY = y
'        Printer.Print IQty
'
''        x = 7250
''        Printer.CurrentX = x
''        Printer.CurrentY = y
''        Printer.Print Format(IGrossvalue, "0.00")
'
'        x = 7950
'        Printer.CurrentX = x
'        Printer.CurrentY = y
'        Printer.Print Format(IDiscount, "0.00")
'
''        x = 9100
''        Printer.CurrentX = x
''        Printer.CurrentY = y
''        Printer.Print Format(INetValue, "0.00")
'
'         x = 8950
'        Printer.CurrentX = x
'        Printer.CurrentY = y
'        Printer.Print Format(ITaxAmount, "0.00")
'
'         x = 9950
'        Printer.CurrentX = x
'        Printer.CurrentY = y
'        Printer.Print Format(INetValue + ITaxAmount, "0.00")
'
'    y = 10100
'    Printer.FontBold = True
'    Printer.CurrentX = 8750
'    Printer.CurrentY = y
'    Printer.Print "  Servicing Charge:"
'
'    Printer.CurrentX = 10600
'    Printer.CurrentY = y
'    Printer.Print Format(Val(TServiceCharge.Text), "0.00")
'
'    y = y + 500
'    Printer.FontBold = True
'    Printer.CurrentX = 8750
'    Printer.CurrentY = y
'    Printer.Print "        Discount Amt:"
'
'    Printer.CurrentX = 10600
'    Printer.CurrentY = y
'    Printer.Print Format(Val(TDiscount.Text), "0.00")
'
'    y = y + 500
'    Printer.FontSize = 16
'    Printer.FontBold = True
'    Printer.CurrentX = 9200
'    Printer.Font = "Rupee"
'    Printer.CurrentY = y
'    Printer.Print "`"
'
'    Printer.CurrentX = 9900
'    Printer.CurrentY = y
'    Printer.Font = "Arial"
'    Printer.FontSize = 12
'    Printer.Print Format(Val(LGrandAmount.Caption), "0.00")
'
''    Printer.Print Tab(5); String(110, "-")
''    Printer.Print Tab(10); "Tax"; Tab(20); "Gross Value"; Tab(35); "Tax Amt"; Tab(50); "Cess Amt"; Tab(65); "Net Amount";
''    Printer.Print Tab(5); String(110, "-")
''
''    If Check = True Then
''        Printer.Print Tab(10); "4.00"; Tab(20); Format(TGrossvalue, "0.00"); Tab(35); Format(Taxamt, "0.00"); Tab(50); Format(Taxamt * 0.01, "0.00"); Tab(67); Format(TNetamt, "0.00");
''    End If
''    If Check1 = True Then
''        Printer.Print Tab(10); "12.50"; Tab(20); Format(TGrossvalue1, "0.00"); Tab(35); Format(Taxamt1, "0.00"); Tab(50); Format(Taxamt1 * 0.01, "0.00"); Tab(67); Format(TNetamt1, "0.00");
''    End If
''    Printer.Print Tab(5); String(110, "-")
'
'    x = 500
'    y = 13900
'    Printer.FontSize = 10
'    Printer.FontUnderline = False
'    Printer.CurrentX = x
'    Printer.CurrentY = y
'    Printer.Print "Customer Signature"
'
'    x = 9300
'    y = 13900
'    Printer.FontSize = 10
'    Printer.FontUnderline = False
'    Printer.CurrentX = x
'    Printer.CurrentY = y
'    Printer.Print "For REDLINES"
'
'    x = 500
'    y = 14200
'    Printer.FontSize = 10
'    Printer.FontBold = False
'    Printer.FontUnderline = False
'    Printer.CurrentX = x
'    Printer.CurrentY = y
'    Printer.Print "Received With Good Condition"
'
'    Printer.EndDoc
'
'    x = MsgBox("Successfully Printed !", vbInformation)
'
'GoOut:
'End Sub
'
'Private Function NewPage() As Long
'
'    Dim i, j, x, y, D, M, YR, DT1, TOPH As Double
'    Dim Declaration(10) As String
'
'    Printer.ScaleMode = 1
'    Printer.FontName = "Arial"
'    Printer.FontBold = False
'    y = 400
'    x = 450
'
'    Printer.FontSize = 10
'    Printer.CurrentX = x
'    Printer.CurrentY = y
'    Printer.Print "CST NO :"
'
'    x = x + 8500
'    Printer.CurrentX = x
'    Printer.CurrentY = y
'    Printer.Print "TIN NO :"
'
'    Printer.FontBold = True
'    Printer.FontUnderline = True
'    Printer.FontSize = 14
'    Printer.CurrentX = (Val(Printer.Width) / 2) - (Val(Printer.TextWidth("RED LINES")) / 2)
'    Printer.CurrentY = 400
'    Printer.Print "RED LINES"
'    Printer.FontUnderline = False
'    Printer.FontBold = False
'    x = 400
'    y = 800
''
''    Printer.FontUnderline = True
''    Printer.FontSize = 16
''    Printer.CurrentX = x
''    Printer.CurrentY = y
''    Printer.FontUnderline = True
''    Printer.Print "Ink - Opening Stock"
''
''    Printer.FontBold = False
'
'    Printer.FontSize = 10
'    Printer.CurrentX = (Val(Printer.Width) / 2) - (Val(Printer.TextWidth("Peruvazhiyambalam , Tirur ")) / 2)
'    Printer.CurrentY = 800
'    Printer.Print "Peruvazhiyambalam , Tirur "
'
'    Printer.CurrentX = (Val(Printer.Width) / 2) - (Val(Printer.TextWidth("THE KERALA VALUE ADDED TAX RULES 2005 FORM NO.8B")) / 2)
'    Printer.CurrentY = 1000
'    Printer.Print "THE KERALA VALUE ADDED TAX RULES 2005 FORM NO.8B"
'
'    Printer.CurrentX = (Val(Printer.Width) / 2) - (Val(Printer.TextWidth("(For Customers When input tax credit is not required)[See Rule 58(10)]")) / 2)
'    Printer.CurrentY = 1200
'    Printer.Print "(For Customers When input tax credit is not required)[See Rule 58(10)]"
'
'    Printer.CurrentX = (Val(Printer.Width) / 2) - (Val(Printer.TextWidth("TAX INVOICE")) / 2)
'    Printer.CurrentY = 1400
'    Printer.Print "TAX INVOICE"
'
'    Printer.FontSize = 10
'    Printer.FontUnderline = False
'    Printer.CurrentX = x
'    y = y + 1100
'    Printer.CurrentY = y
'    Printer.Print "Invoice No"
'
'    x = x + 1100
'    Printer.CurrentX = x
'    Printer.CurrentY = y
'    Printer.Print ": "
'
'    x = x + 200
'    Printer.FontSize = 10
'    Printer.CurrentX = x
'    Printer.CurrentY = y
'    Printer.Print Trim(TTransactionNo.Text)
'
'    x = x + 6500
'    Printer.FontBold = False
'    Printer.FontSize = 10
'    Printer.FontUnderline = False
'    Printer.CurrentX = x
'    Printer.CurrentY = y
'    Printer.Print "Customer"
'
'    x = x + 1000
'    Printer.CurrentX = x
'    Printer.CurrentY = y
'    Printer.Print ": "
'
'    x = x + 200
'    Printer.FontSize = 10
'    Printer.CurrentX = x
'    Printer.CurrentY = y
'    Printer.Print Trim(CoCustomer.Text)
'
'    D = Trim(Day(DTPDate))
'    M = Trim(Month(DTPDate))
'    YR = Trim(Year(DTPDate))
'    If Len(D) = 1 Then D = "0" & D
'    If Len(M) = 1 Then M = "0" & M
'    DT1 = D & "-" & M & "-" & YR
'
'    x = 600
'    y = y + 200
'    Printer.FontSize = 10
'    Printer.CurrentX = x
'    Printer.CurrentY = y
'    Printer.Print "Date"
'
'    x = x + 900
'    Printer.CurrentX = x
'    Printer.CurrentY = y
'    Printer.Print ": "
'
'    x = x + 200
'    Printer.FontSize = 10
'    Printer.CurrentX = x
'    Printer.CurrentY = y
'    Printer.Print Trim(DT1)
'
'    x = x + 6500
'    Printer.FontSize = 10
'    Printer.CurrentX = x
'    Printer.CurrentY = y
'    Printer.Print "Address"
'
'    x = x + 1000
'    Printer.CurrentX = x
'    Printer.CurrentY = y
'    Printer.Print ": "
'
'    x = x + 100
'    Printer.FontSize = 10
'    Printer.CurrentX = x
'    Printer.CurrentY = y
'    Printer.Print Trim(TAddress.Text)
'
'    x = 500
'    y = y + 1600
'
'
'    Printer.FontBold = True
'    Printer.FontSize = 9
'    Printer.CurrentX = x
'    Printer.CurrentY = y
'    Printer.Print "SNo"
'
'    x = 100 + 1000
'    Printer.FontSize = 9
'    Printer.CurrentX = x
'    Printer.CurrentY = y
'    Printer.Print "Particulars"
'
'    x = 100 + 5200
'    Printer.FontSize = 9
'    Printer.CurrentX = x
'    Printer.CurrentY = y
'    Printer.Print "Tax % "
'
'    x = 100 + 5900
'    Printer.FontSize = 9
'    Printer.CurrentX = x
'    Printer.CurrentY = y
'    Printer.Print "Rate"
'
'    x = 100 + 7200
'    Printer.FontSize = 9
'    Printer.CurrentX = x
'    Printer.CurrentY = y
'    Printer.Print "Qty"
'
''    x = 100 + 7100
''    Printer.FontSize = 9
''    Printer.CurrentX = x
''    Printer.CurrentY = y
''    Printer.Print "GR Value"
'
'    x = 100 + 7900
'    Printer.FontSize = 9
'    Printer.CurrentX = x
'    Printer.CurrentY = y
'    Printer.Print "Disc"
'
''    x = 100 + 9000
''    Printer.FontSize = 9
''    Printer.CurrentX = x
''    Printer.CurrentY = y
''    Printer.Print "Net Amt"
'
'    x = 100 + 8900
'    Printer.FontSize = 9
'    Printer.CurrentX = x
'    Printer.CurrentY = y
'    Printer.Print "Tax Amt"
'
'    x = 100 + 9900
'    Printer.FontSize = 9
'    Printer.CurrentX = x
'    Printer.CurrentY = y
'    Printer.Print "Total"
'
'    'HORIZONTAL LINES
'    Printer.Line (400, 3600)-(11000, 3600)
'    Printer.Line (400, 4000)-(11000, 4000)
'    Printer.Line (400, 10000)-(11000, 10000)
'    Printer.Line (400, 9700)-(11000, 9700)
'
'
'    'FIRST AND LAST VERTICAL LINE
'    Printer.Line (400, 3600)-(400, 10000)
'    Printer.Line (11000, 3600)-(11000, 10000)
'
'    'INNER LINES
'    Printer.Line (1000, 3600)-(1000, 10000)
'    Printer.Line (5200, 3600)-(5200, 10000)
'    Printer.Line (5900, 3600)-(5900, 10000)
'    Printer.Line (7200, 3600)-(7200, 10000)
'    Printer.Line (7900, 3600)-(7900, 10000)
'    Printer.Line (8900, 3600)-(8900, 10000)
'    Printer.Line (9900, 3600)-(9900, 10000)
''    Printer.Line (10000, 3600)-(10000, 10000)
''    Printer.Line (10900, 3600)-(10900, 10000)
'
'
'
'
'    Printer.FontSize = 10
'    Printer.FontItalic = True
'    Printer.FontBold = False
'    Printer.CurrentY = 11600
'    Printer.CurrentX = 1000
'    Printer.Print (NumberToWords(Val(LGrandAmount.Caption & "")))
'    Printer.FontItalic = False
''    Print #1, Chr(27) & "!" & Chr(4) & "|Amount in Words:" & Left(NumberToWords(Val(LGrandAmount.Caption & "")) & Space(66), 66) & " Balance                                |" & Chr(0) & Chr(27) & "!" & Chr(29) & Right(Space(13) & Format("0" & LBalance.Caption, "0.00"), 13) & "|" & Chr(0)
'
'    Printer.FontSize = 10
'    Printer.FontUnderline = True
'    Printer.FontBold = True
'    Printer.CurrentX = (Val(Printer.Width) / 2) - (Val(Printer.TextWidth("Declaration")) / 2)
'    Printer.CurrentY = 12600
'    Printer.Print "Declaration"
'    Printer.FontBold = False
'    Printer.FontUnderline = False
'
'    Declaration(0) = "DECLARATION : Certified that all the particulars shown in the above Tax Invoice are true and correct and that my/our registration under"
'    Declaration(1) = "KVAT ACT is valid as on the date of this bill"
'
'
'TOPH = 200
'
'    For i = 0 To 2
'        Printer.FontSize = 9
'        Printer.CurrentX = 550
'        Printer.CurrentY = Printer.CurrentY + TOPH
'        If i = 2 Then
'            Printer.Print Declaration(i)
'        Else
'            Printer.Print Declaration(i);
'        End If
'    Next
'NewPage = y
'End Function
'
Private Sub CNext_Click()
    FFormNo1a.Hide
    FFormNo1b.Show
End Sub
