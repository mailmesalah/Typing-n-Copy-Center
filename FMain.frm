VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FMain 
   BackColor       =   &H00EFEFEF&
   Caption         =   "DeXtop - Accessories Software"
   ClientHeight    =   9585
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14130
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9585
   ScaleWidth      =   14130
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2535
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSComDlg.CommonDialog CoDialog 
      Left            =   -45
      Top             =   -30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lSequenceI As Long

Private Sub backUp()
On Error GoTo GoOut
Dim x As Long
    
    'BACKUP DATA
    Dim fso As Object, s As String
    
    CoDialog.CancelError = True
    
    CoDialog.FileName = "DeXtop_" & Day(Date) & "_" & Month(Date) & "_" & Year(Date)
    CoDialog.Filter = "mdb"
    CoDialog.ShowSave
    Set fso = CreateObject("Scripting.FileSystemObject")
    x = fso.CopyFile(App.Path & "/Storage.mdb", CoDialog.FileName & ".mdb", True)
    
    x = MsgBox("Successfully Exported !", vbInformation)
    Exit Sub
GoOut:
    x = MsgBox("Backup was Failed : " & Err.Description, vbInformation)
End Sub

Private Sub reStore()
On Error GoTo GoOut
Dim x As Long
    
    If (MsgBox("Are you sure to Restore ? ,Current Data will be Overwritten !", vbDefaultButton2 Or vbYesNo) = vbNo) Then
        Exit Sub
    End If
    
    'RESTORE DATA
    Dim fso As Object
    
    CoDialog.CancelError = True
    CoDialog.Filter = "mdb"
    CoDialog.ShowOpen
    Set fso = CreateObject("Scripting.FileSystemObject")
    x = fso.CopyFile(CoDialog.FileName, App.Path & "/Storage.mdb", True)
    
    x = MsgBox("Successfully Restored !", vbInformation)
    Exit Sub
    
GoOut:
    x = MsgBox("Restore was Failed : " & Err.Description, vbInformation)
End Sub

Private Sub Form_Load()
    setUserAccess
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub MExit_Click()
    End
End Sub

Private Sub MMAccountMaster_Click()
    FAccountRegister.Show
End Sub

Private Sub MMCustomerMaster_Click()
    FCustomerRegister.Show
End Sub

Private Sub MMItemMaster_Click()
    FItemMaster.Show
End Sub

Private Sub MMSupplierMaster_Click()
    FSupplierRegister.Show
End Sub

Private Sub MOCustomerWiseSaleReport_Click()
    FCustomerWiseSaleRegister.Show
End Sub

Private Sub MOItemWiseSaleReport_Click()
    FItemWiseReport.Show
End Sub

Private Sub MOOutstandingSummaryReport_Click()
    'FOutStandingReport.Show
End Sub

Private Sub MRCashBook_Click()
    FCashBook.Show
End Sub

Private Sub MRCashInHand_Click()
    FCashInHand.Show
End Sub

Private Sub MRDayBook_Click()
    FDayBook.Show
End Sub

Private Sub MRLedgerReport_Click()
    FLedgerReport.Show
End Sub

Private Sub MRMinimumStock_Click()
    FMinimumStock.Show
End Sub

Private Sub MROutstandingPayable_Click()
    FOutstandingPayable.Show
End Sub

Private Sub MROutstandingReceivable_Click()
    FOutstandingReceivable.Show
End Sub

Private Sub MRPurchaseSummary_Click()
    FPurchaseRegister.Show
End Sub

Private Sub MRRateComparison_Click()
    FRateComparison.Show
End Sub

Private Sub MRSaleProfitOrLoss_Click()
    FSaleProfitAndLoss.Show
End Sub

Private Sub MRSaleSummary_Click()
    FSaleRegister.Show
End Sub

Private Sub MRStockReport_Click()
    FStockRegister.Show
End Sub

Private Sub MRTaxReport_Click()
    FTaxReport.Show
End Sub

Private Sub MRWarantyReport_Click()
    FWarrantyDetails.Show
End Sub

Private Sub MSAbout_Click()
    FAbout.Show
End Sub
Private Sub MSBackup_Click()
    backUp
End Sub

Private Sub MSRestore_Click()
    reStore
End Sub

Private Sub MSUserAccounts_Click()
  FUserAccounts.Show
End Sub

Private Sub MTOpeningBalance_Click()
    FOpeningBalance.Show
End Sub

Private Sub MTOpeningStock_Click()
    FOpeningStock.Show
End Sub

Private Sub MTPayment_Click()
    FPayment.Show
End Sub

Private Sub MTPurchase_Click()
    FPurchase.Show
End Sub

Private Sub MTPurchaseReturn_Click()
    FPurchaseReturn.Show
End Sub

Private Sub MTReceipt_Click()
    FReceipt.Show
End Sub

Private Sub MTRetailSales_Click()
    FRetailSales.Show
End Sub

Private Sub MTSaleReturn_Click()
    FSaleReturn.Show
End Sub

Private Sub setUserAccess()
Dim rs As Recordset, r As Long
    
    Set rs = db.OpenRecordset("Select Rights.RightDescription,Rights.MapName,Rights.Status,Users.RightCode From Rights,Users Where (Users.Code = '" & sCurrentUserCode & "' ) And (Rights.Code = Users.RightCode )")
    If rs.RecordCount > 0 Then
        If Trim(rs!RightDescription) = "Administrator" Then
            'SHOW ALL
            r = 0
            Do While r < Me.Controls.Count
                'SKIPPING MENU DIVIDERS
                If Left(Me.Controls(r).Name, 1) = "M" And Len(Me.Controls(r).Name) > 5 Then
                    Me.Controls(r).Visible = True
                End If
                r = r + 1
            Loop
            MSettings.Visible = True
            MSAbout.Visible = True
            
        ElseIf Trim(rs!RightDescription) = "None" Then
            'SHOW NONE
            r = 0
            Do While r < Me.Controls.Count
                If Left(Me.Controls(r).Name, 1) = "M" And Len(Me.Controls(r).Name) > 4 Then
                    Me.Controls(r).Visible = False
                End If
                r = r + 1
            Loop
            MSettings.Visible = True
            MSAbout.Visible = True
                        
        Else
            While rs.EOF = False
'
                If Left(rs!MapName, 1) = "B" Then

                Else
                    r = 0
                    Do While r < Me.Controls.Count
                        If Trim(Me.Controls(r).Name) = Trim(rs!MapName) Then
                            Me.Controls(r).Visible = rs!Status
                            Exit Do
                      End If
                        r = r + 1
                    Loop
                End If
                rs.MoveNext
            Wend
            
            MSettings.Visible = True
            MSAbout.Visible = True
        End If
    Else
        'SHOW NONE
        
    End If
    rs.Close
End Sub

