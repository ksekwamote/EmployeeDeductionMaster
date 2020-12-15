Public Class MainForm

    Public usertype As String

    Private Sub MainForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        mform = Me

        Dim frmLogin As New LogIn
        frmLogin.vrl = Me
        If frmLogin.ShowDialog = Windows.Forms.DialogResult.OK Then
            frmLogin.vrl.Show()
        ElseIf frmLogin.ShowDialog = Windows.Forms.DialogResult.Cancel Then
            Me.Close()
            frmLogin.vrl.Close()
        End If

        If usertype = "Front Office" Then
            MenuStrip1.Items(0).Enabled = True
            MenuStrip1.Items(1).Enabled = False
            MenuStrip1.Items(2).Enabled = False
            MenuStrip1.Items(3).Enabled = False
            MenuStrip1.Items(4).Enabled = False
            MenuStrip1.Items(5).Enabled = False
            MenuStrip1.Items(6).Enabled = False
        ElseIf usertype = "Front Office Receipts And Cheques" Then
            MenuStrip1.Items(0).Enabled = True
            MenuStrip1.Items(1).Enabled = True
            MenuStrip1.Items(2).Enabled = False
            MenuStrip1.Items(3).Enabled = False
            MenuStrip1.Items(4).Enabled = False
            MenuStrip1.Items(5).Enabled = False
            MenuStrip1.Items(6).Enabled = False
        ElseIf usertype = "Front Office Supervisor" Then
            MenuStrip1.Items(0).Enabled = True
            MenuStrip1.Items(1).Enabled = True
            MenuStrip1.Items(2).Enabled = True
            MenuStrip1.Items(3).Enabled = False
            MenuStrip1.Items(4).Enabled = False
            MenuStrip1.Items(5).Enabled = False
            MenuStrip1.Items(6).Enabled = False
        ElseIf usertype = "Back Office Viewer" Then
            MenuStrip1.Items(0).Enabled = True
            MenuStrip1.Items(1).Enabled = True
            MenuStrip1.Items(2).Enabled = False
            MenuStrip1.Items(3).Enabled = True
            MenuStrip1.Items(4).Enabled = False
            MenuStrip1.Items(5).Enabled = False
            MenuStrip1.Items(6).Enabled = False
        ElseIf usertype = "Back Office Editor" Then
            MenuStrip1.Items(0).Enabled = True
            MenuStrip1.Items(1).Enabled = True
            MenuStrip1.Items(2).Enabled = False
            MenuStrip1.Items(3).Enabled = True
            MenuStrip1.Items(4).Enabled = True
            MenuStrip1.Items(5).Enabled = False
            MenuStrip1.Items(6).Enabled = False
        ElseIf usertype = "Back Office Supervisor" Then
            MenuStrip1.Items(0).Enabled = True
            MenuStrip1.Items(1).Enabled = True
            MenuStrip1.Items(2).Enabled = True
            MenuStrip1.Items(3).Enabled = True
            MenuStrip1.Items(4).Enabled = True
            MenuStrip1.Items(5).Enabled = True
            MenuStrip1.Items(6).Enabled = False
        ElseIf usertype = "Administrator" Then
            MenuStrip1.Items(0).Enabled = True
            MenuStrip1.Items(1).Enabled = True
            MenuStrip1.Items(2).Enabled = True
            MenuStrip1.Items(3).Enabled = True
            MenuStrip1.Items(4).Enabled = True
            MenuStrip1.Items(5).Enabled = True
            MenuStrip1.Items(6).Enabled = True
        End If

    End Sub

    Private Sub ProgressReportToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim vrl As New Progress
        vrl.MdiParent = mform
        vrl.Text = vrl.Text & " : Employer Name = " & employername
        vrl.Show()
    End Sub

    Private Sub UploadExpectedPerProductToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim vrl As New UploadExpectedPerProduct
        vrl.MdiParent = mform
        vrl.Text = vrl.Text & " : Employer Name = " & employername
        vrl.Show()
    End Sub

    Private Sub LodgingFileToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LodgingFileToolStripMenuItem.Click
        Dim vrl As New ViewLodgingReport
        vrl.MdiParent = mform
        vrl.Text = vrl.Text & " : Employer Name = " & employername
        vrl.Show()
    End Sub

    Private Sub ProgressReportToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProgressReportToolStripMenuItem1.Click
        Dim vrl As New Progress
        vrl.MdiParent = mform
        vrl.Text = vrl.Text & " : Employer Name = " & employername
        vrl.Show()
    End Sub

    Private Sub UploadCustomerDetailsToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UploadCustomerDetailsToolStripMenuItem1.Click
        Dim vrl As New UploadingCustomerDetails
        vrl.MdiParent = mform
        vrl.Text = vrl.Text & " : Employer Name = " & employername
        vrl.Show()
    End Sub

    Private Sub AddUsersToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddUsersToolStripMenuItem.Click
        Dim vrl As New AddUser
        vrl.MdiParent = mform
        vrl.Text = vrl.Text & " : Employer Name = " & employername
        vrl.Show()
    End Sub

    Private Sub ProductsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProductsToolStripMenuItem.Click
        Dim vrl As New Add_Products
        vrl.MdiParent = mform
        vrl.Text = vrl.Text & " : Employer Name = " & employername
        vrl.Show()
    End Sub

    Private Sub ViewAmendmentsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ViewAmendmentsToolStripMenuItem.Click
        Dim vrl As New ViewAmendments
        vrl.MdiParent = mform
        vrl.Text = vrl.Text & " : Employer Name = " & employername
        vrl.Show()
    End Sub

    Private Sub EmployersToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EmployersToolStripMenuItem.Click
        Dim vrl As New EmployerDetails
        vrl.MdiParent = mform
        vrl.Text = vrl.Text & " : Employer Name = " & employername
        vrl.Show()
    End Sub

    Private Sub ListOfCustomersToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListOfCustomersToolStripMenuItem.Click
        Dim vrl As New ListOfCustomers
        vrl.MdiParent = mform
        vrl.Text = vrl.Text & " : Employer Name = " & employername
        vrl.Show()
    End Sub

    Private Sub DefaultersListToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DefaultersListToolStripMenuItem.Click
        Dim vrl As New Defaulters_List
        vrl.MdiParent = mform
        vrl.Text = vrl.Text & " : Employer Name = " & employername
        vrl.Show()
    End Sub

    Private Sub SwitchEmployersToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SwitchEmployersToolStripMenuItem.Click
        Dim vrl As New SwitchEmployers
        vrl.MdiParent = mform
        vrl.Text = vrl.Text & " : Employer Name = " & employername
        vrl.mainw = Me
        vrl.Show()
    End Sub

    Private Sub CommissionProductsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CommissionProductsToolStripMenuItem.Click
        Dim vrl As New CommissionProducts
        vrl.MdiParent = mform
        vrl.Text = vrl.Text & " : Employer Name = " & employername
        vrl.Show()
    End Sub

    Private Sub AddRequestedChequeRefundToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddRequestedChequeRefundToolStripMenuItem.Click
        Dim vrl As New Add_Requested_Refund
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub ApprovePendingRefundToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ApprovePendingRefundToolStripMenuItem.Click
        Dim vrl As New ApprovePendingRefunds
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub RequestChequeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RequestChequeToolStripMenuItem.Click
        Dim vrl As New PaymentProcessing
        vrl.MdiParent = mform
        vrl.Show()
    End Sub


    Private Sub DailyChequesReportToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DailyChequesReportToolStripMenuItem.Click
        Dim vrl As New DailyChequesReport
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub AddCustomerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddCustomerToolStripMenuItem.Click
        Dim vrl As New AddCustomer
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub CompanyDetailsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CompanyDetailsToolStripMenuItem.Click
        Dim vrl As New CompanyDetails
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub SuperviorsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SuperviorsToolStripMenuItem.Click
        Dim vrl As New Supervisors
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub SettlementLetterDesignToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SettlementLetterDesignToolStripMenuItem.Click
        Dim vrl As New SettlementLettersDesign
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub ClearanceLedterDesignToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClearanceLedterDesignToolStripMenuItem.Click
        Dim vrl As New ClearanceLettersDesign
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub GenerateSettlementLettersToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GenerateSettlementLettersToolStripMenuItem.Click
        Dim vrl As New SettlementBalance
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub SettlementBalancesToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SettlementBalancesToolStripMenuItem1.Click
        Dim vrl As New SettlementBalancesList
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub ConfirmWrittenChequesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ConfirmWrittenChequesToolStripMenuItem.Click
        Dim vrl As New WrittenChequesProcessing
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub CancelledChequesToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CancelledChequesToolStripMenuItem1.Click
        Dim vrl As New CancelCheques
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub ViewCustomerLatestDeductionToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ViewCustomerLatestDeductionToolStripMenuItem.Click
        Dim vrl As New ViewActualDeductionForCustomer
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub ViewExpectedDeductionsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ViewExpectedDeductionsToolStripMenuItem.Click
        Dim vrl As New ViewLodgingReportForCustomer
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub UploadActualDeductionsToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UploadActualDeductionsToolStripMenuItem1.Click
        Dim vrl As New Uploading_Deductions
        vrl.MdiParent = mform
        vrl.Text = vrl.Text & " : Employer Name = " & employername
        vrl.Show()
    End Sub

    Private Sub GenerateActualDeductionBreakdownToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GenerateActualDeductionBreakdownToolStripMenuItem.Click
        Dim vrl As New Deduction_File
        vrl.MdiParent = mform
        vrl.Text = vrl.Text & " : Employer Name = " & employername
        vrl.Show()
    End Sub

    Private Sub UploadAmendmentAndTerminationsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UploadAmendmentAndTerminationsToolStripMenuItem.Click
        Dim vrl As New Uploading_Amendments
        vrl.MdiParent = mform
        vrl.Text = vrl.Text & " : Employer Name = " & employername
        vrl.Show()
    End Sub

    Private Sub GenerateLodgingFileToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GenerateLodgingFileToolStripMenuItem1.Click
        Dim vrl As New CustomersPaidForList
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub UpdateAmendmentOrTerminationToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UpdateAmendmentOrTerminationToolStripMenuItem.Click
        Dim vrl As New UpdateLodgingReport
        vrl.MdiParent = mform
        vrl.Text = vrl.Text & " : Employer Name = " & employername
        vrl.Show()
    End Sub

    Private Sub ProcessStopRefundsToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProcessStopRefundsToolStripMenuItem1.Click
        Dim vrl As New PotentialRefundsToStop
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub RemoveDuplicatesCustomerRecordsToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RemoveDuplicatesCustomerRecordsToolStripMenuItem1.Click
        Dim vrl As New RemoveDuplicateCustomerRecords
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub ChangeCustomersToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChangeCustomersToolStripMenuItem.Click
        Dim vrl As New ChangeCustomerEmployer
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub UploadEmployersToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UploadEmployersToolStripMenuItem.Click
        Dim vrl As New UploadEmployers
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub UploadProductsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UploadProductsToolStripMenuItem.Click
        Dim vrl As New UploadProducts
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub UploadSupervisorsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UploadSupervisorsToolStripMenuItem.Click
        Dim vrl As New UploadSupervisors
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub LocalDatabaseRefilToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LocalDatabaseRefilToolStripMenuItem.Click
        Dim vrl As New LocalDatabaseRefill
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub UploadTerminationsAndAmendmentsForCorrectionToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UploadTerminationsAndAmendmentsForCorrectionToolStripMenuItem.Click
        Dim vrl As New UploadTerminationsAndAmendments
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub ReverseChequeToRequestedToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReverseChequeToRequestedToolStripMenuItem.Click
        Dim vrl As New ReverseCancelledChequeToRequested
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub ConfirmCutOfDateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ConfirmCutOfDateToolStripMenuItem.Click
        Dim vrl As New CutOfDate
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub ModeOfPaymentToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ModeOfPaymentToolStripMenuItem.Click
        Dim vrl As New ModeOfPayment
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub CompanyBooksToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CompanyBooksToolStripMenuItem.Click
        Dim vrl As New CompanyBooks
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub CompanyBranchesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CompanyBranchesToolStripMenuItem.Click
        Dim vrl As New Branches
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub MyCapturedReceiptsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyCapturedReceiptsToolStripMenuItem.Click
        Dim vrl As New ApproveReceipts
        vrl.MdiParent = mform
        vrl.Text = "My Captured Receipts"
        vrl.btapprove.Enabled = False
        vrl.btdecline.Enabled = False
        vrl.whichreport = "My Captured Receipts"
        vrl.Show()
    End Sub

    Private Sub ApproveReceiptToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ApproveReceiptToolStripMenuItem1.Click
        Dim vrl As New ApproveReceipts
        vrl.MdiParent = mform
        vrl.whichreport = "New Receipts"
        vrl.Show()
    End Sub

    Private Sub ReceiptCustomerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReceiptCustomerToolStripMenuItem.Click
        Dim vrl As New ReceiptBook
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub ConfirmBankingToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ConfirmBankingToolStripMenuItem1.Click
        Dim vrl As New ConfirmBanking
        vrl.MdiParent = mform
        vrl.choice = "Banking"
        vrl.Show()
    End Sub

    Private Sub ExportReceiptsToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExportReceiptsToolStripMenuItem1.Click
        Dim vrl As New ReceiptsToExport
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub ReceiptsReportToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReceiptsReportToolStripMenuItem1.Click
        Dim vrl As New ReceiptsReport
        vrl.MdiParent = mform
        vrl.GroupBox3.Visible = True
        vrl.Show()
    End Sub

    Private Sub ReceiptsReportToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReceiptsReportToolStripMenuItem.Click
        Dim vrl As New ReceiptsReport
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub BankingReceiptsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BankingReceiptsToolStripMenuItem.Click
        Dim vrl As New BankingReceipts
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub ProductsToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProductsToolStripMenuItem1.Click
        Dim vrl As New Add_Products
        vrl.MdiParent = mform
        vrl.Text = vrl.Text & " : Employer Name = " & employername
        vrl.Button1.Enabled = False
        vrl.Button2.Enabled = False
        vrl.Button3.Enabled = False
        vrl.Show()
    End Sub

    Private Sub MyaBToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyaBToolStripMenuItem.Click
        Dim vrl As New ApproveReceiptBatch
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub MyReceiptsBatchPendingConfirmationToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyReceiptsBatchPendingConfirmationToolStripMenuItem.Click
        Dim vrl As New ReceiptBatch
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub ChangeReceiptPayerIDToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChangeReceiptPayerIDToolStripMenuItem.Click
        Dim vrl As New CorrectReceiptPayerID
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub CustomerBalanceReconToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CustomerBalanceReconToolStripMenuItem.Click
        Dim vrl As New CustomerBalanceRecon
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub BankingBatchesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BankingBatchesToolStripMenuItem.Click
        Dim vrl As New ConfirmBanking
        vrl.MdiParent = mform
        vrl.choice = "All"
        vrl.Text = "All Batches"
        vrl.Button1.Enabled = False
        vrl.Show()
    End Sub

    Private Sub CustomerListToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CustomerListToolStripMenuItem.Click
        Dim vrl As New ListOfCustomers
        vrl.MdiParent = mform
        vrl.Text = vrl.Text & " : Employer Name = " & employername
        vrl.Show()
    End Sub

    Private Sub DeleteCustomersToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteCustomersToolStripMenuItem.Click
        Dim vrl As New DeleteCustomers
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub Form1ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Form1ToolStripMenuItem.Click
        Dim vrl As New DeductionsBreakdown
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub TransferReceiptsBetweenCompanyBooksToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TransferReceiptsBetweenCompanyBooksToolStripMenuItem1.Click
        Dim vrl As New TransferReceiptBetweenBooks
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub TransferReceiptsBetweenCompanyBooksToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TransferReceiptsBetweenCompanyBooksToolStripMenuItem.Click
        Dim vrl As New TransferReceiptBetweenPaymentModes
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub UpdateReceiptNumberToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UpdateReceiptNumberToolStripMenuItem1.Click
        Dim vrl As New UpdateReceiptNo
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub TotalReceiptsReportToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TotalReceiptsReportToolStripMenuItem.Click
        Dim vrl As New Total_Receipts_Report
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub DailyDisbursementsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DailyDisbursementsToolStripMenuItem.Click
        Dim vrl As New DailyTotals
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub CovidRegisterToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CovidRegisterToolStripMenuItem.Click
        Dim vrl As New FileConfirmation
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub CovidRegisterToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CovidRegisterToolStripMenuItem1.Click
        Dim vrl As New CovidRegister
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub CustomersPaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CustomersPaToolStripMenuItem.Click
        Dim vrl As New CustomersPaidFor
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub ListOfCustomersPaidForToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListOfCustomersPaidForToolStripMenuItem.Click
        Dim vrl As New CustomersPaidForList
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub CustomerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CustomerToolStripMenuItem.Click
        Dim vrl As New SearchCustomerDetails
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub CountriesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CountriesToolStripMenuItem.Click
        Dim vrl As New Countries
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub DepartmentsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DepartmentsToolStripMenuItem.Click
        Dim vrl As New Department
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub IdentityTypesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IdentityTypesToolStripMenuItem.Click
        Dim vrl As New IdentityType
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub GenderToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GenderToolStripMenuItem.Click
        Dim vrl As New Gender
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub RegionToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RegionToolStripMenuItem.Click
        Dim vrl As New Region
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub SchoolToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SchoolToolStripMenuItem.Click
        Dim vrl As New School
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub EmployementTypeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EmployementTypeToolStripMenuItem.Click
        Dim vrl As New EmploymentType
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub LocationToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LocationToolStripMenuItem.Click
        Dim vrl As New Location
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub UploadLocationsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UploadLocationsToolStripMenuItem.Click
        Dim vrl As New UploadLocations
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub GroupToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupToolStripMenuItem.Click
        Dim vrl As New Group
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub BankToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BankToolStripMenuItem.Click
        Dim vrl As New Bank
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub BankBranchToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BankBranchToolStripMenuItem.Click
        Dim vrl As New BankBranch
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub ReceiptsReportPerModeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReceiptsReportPerModeToolStripMenuItem.Click
        Dim vrl As New ReceiptsReportForModeOfPayment
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub RequestPrintingToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RequestPrintingToolStripMenuItem.Click
        Dim vrl As New PrintingRequests
        vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub ConfirmBarcodePrintingToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ConfirmBarcodePrintingToolStripMenuItem.Click
        Dim vrl As New PrintingStatus
        vrl.MdiParent = mform
        vrl.Show()
    End Sub
End Class
