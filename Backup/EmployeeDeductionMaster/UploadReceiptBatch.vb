Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data.Odbc
Imports System.IO
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports CrystalDecisions.ReportSource

Public Class UploadReceiptBatch

    Dim batchname As String
    Public receiptbookid As Integer
    Dim go As Boolean

    Dim crrpl As New ReportDocument

    Dim PostalAddress, PhysicalAddress, ContactAddresses, ReceivedFromName, ReceivedFromIDNumber, ReceivedFromCellNumber, PayerName, PayerIDNumber, PayerCellNumber, Description, PaymentMode, Amount, ReceiptDate, ReceiptNo, ProductsName, ProductAmount As String

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            DataGridView1.AllowUserToAddRows = False

            Button1.Enabled = False
            Button2.Enabled = True
            Dim csvfile As String
            Dim OpenFileDialog As New OpenFileDialog
            OpenFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            OpenFileDialog.Filter = "CSV Files (*.csv)|*.csv"

            If (OpenFileDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK) Then

                Dim fi As New FileInfo(OpenFileDialog.FileName)
                Dim FileName As String = OpenFileDialog.FileName
                csvfile = fi.FullName
                Dim counter As Integer = 1
                For Each line As String In System.IO.File.ReadAllLines(csvfile)
                    If counter = 1 Then
                    Else
                        DataGridView1.Rows.Add(line.Split(","))
                    End If
                    counter = counter + 1
                Next

                For Each row As DataGridViewRow In DataGridView1.Rows
                    If checkCustomerInDatabase(row.Cells(0).Value.ToString) = True Then

                    Else
                        If row.Cells(7).Value = "" Then
                            row.Cells(7).Value = "Customer Does Not Exists"
                        Else
                            row.Cells(7).Value = row.Cells(7).Value & "/ Customer Does Not Exists"
                        End If
                        Button2.Enabled = False
                    End If

                    If Decimal.Parse(row.Cells(1).Value.ToString) > 0 Then

                    Else
                        If row.Cells(7).Value = "" Then
                            row.Cells(7).Value = "Amount should be greator than 0"
                        Else
                            row.Cells(7).Value = row.Cells(7).Value & "/ Amount should be greator than 0"
                        End If
                        Button2.Enabled = False
                    End If

                    If row.Cells(2).Value.ToString = "" Then

                    Else
                        If row.Cells(7).Value = "" Then
                            row.Cells(7).Value = "ModeID should be blank"
                        Else
                            row.Cells(7).Value = row.Cells(7).Value & "/ ModeID should be blank"
                        End If
                        Button2.Enabled = False
                    End If

                    If row.Cells(3).Value.ToString = "" Then

                    Else
                        If row.Cells(7).Value = "" Then
                            row.Cells(7).Value = "PaymentMode should be blank"
                        Else
                            row.Cells(7).Value = row.Cells(7).Value & "/ PaymentMode should be blank"
                        End If
                        Button2.Enabled = False
                    End If

                    If row.Cells(4).Value.ToString = "" Then

                    Else
                        If row.Cells(7).Value = "" Then
                            row.Cells(7).Value = "ProductID should be blank"
                        Else
                            row.Cells(7).Value = row.Cells(7).Value & "/ ProductID should be blank"
                        End If
                        Button2.Enabled = False
                    End If

                    If row.Cells(5).Value.ToString = "" Then

                    Else
                        If row.Cells(7).Value = "" Then
                            row.Cells(7).Value = "ProductName should be blank"
                        Else
                            row.Cells(7).Value = row.Cells(7).Value & "/ ProductName should be blank"
                        End If
                        Button2.Enabled = False
                    End If

                    tbtotalreceipt.Text = Decimal.Parse(tbtotalreceipt.Text.Trim) + Decimal.Parse(row.Cells(1).Value)
                    row.Cells(3).Value = getModeOfPaymentByModeID(Integer.Parse(row.Cells(2).Value))
                    row.Cells(5).Value = getProductNameByProductID(Integer.Parse(row.Cells(4).Value))
                Next

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally

        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If tbdate.Text.Trim = "" Or tbreceivedfrom.Text.Trim = "" Or tbreceivedidnumber.Text.Trim = "" Or tbreceivedcellnumber.Text.Trim = "" Or tbdescription.Text.Trim = "" Or tbamount.Text.Trim = "" Then
            MessageBox.Show("Enter receipt number and date first", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            If decision("Are you to proceed ?") = DialogResult.Yes Then
                go = True
                For Each row As DataGridViewRow In DataGridView1.Rows
                    If Decimal.Parse(row.Cells(1).Value) <= 0 Then
                        go = False
                    End If
                Next

                If go = True Then
                    If decision("Is this the Amount Paid, " & getCurrency(Math.Round(Decimal.Parse(tbtotalreceipt.Text.Trim), 2)) & " ?") = DialogResult.Yes Then
                        Dim vrl As New ReceiptAuthorization
                        vrl.MdiParent = mform
                        vrl.urb = Me
                        vrl.whichside = "Upload"
                        vrl.Show()
                    End If
                ElseIf go = False Then
                    MessageBox.Show("Please ensure that you have put the Receipt Type", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End If
        End If
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        ExportToExcel(DataGridView1, "")
    End Sub

    Public Sub saveBulkReceipt(ByVal authorizer As String)
        For Each row As DataGridViewRow In DataGridView1.Rows
            If Decimal.Parse(row.Cells(1).Value) > 0 Then
                PaymentMode = row.Cells(3).Value.ToString()
                batchname = row.Cells(0).Value.ToString() & ":" & Date.Now
                savePaymentDetails(tbreceivedfrom.Text.Trim, tbreceivedidnumber.Text.Trim, tbreceivedcellnumber.Text.Trim, getCustomerName(row.Cells(0).Value.ToString()), row.Cells(0).Value.ToString(), "", Date.Now.Date, Decimal.Parse(row.Cells(1).Value.ToString()), batchname, tbdescription.Text.Trim, getNextReceipNo())
                saveReceipt(row.Cells(0).Value.ToString(), Decimal.Parse(row.Cells(1).Value.ToString()), Integer.Parse(row.Cells(2).Value.ToString()), DateTimePicker1.Value, authorizer, Date.Now, Date.Now, "", Date.Now, Date.Now, batchname, "New", "", "", getNextReceipNo(), branchid, receiptbookid)
                saveReceiptBreakdown(row.Cells(0).Value.ToString(), Decimal.Parse(row.Cells(1).Value.ToString()), DateTimePicker1.Value, Integer.Parse(row.Cells(4).Value), batchname, "", "New", "", "Repayment", "", Date.Now, getNextReceipNo(), Integer.Parse(row.Cells(2).Value.ToString()), branchid, receiptbookid, row.Cells(6).Value)
                wrapper.saveAction("Receipt Saved", Date.Now, authorizer, row.Cells(0).Value.ToString())
            End If
        Next
        MessageBox.Show("Bulk Receipts Saved", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
        loadReport(PaymentMode)
        Me.TabControl1.SelectTab(1)
    End Sub

    Private Sub tbtotalreceipt_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbtotalreceipt.TextChanged
        Label2.Text = getCurrency(Math.Round(Decimal.Parse(tbtotalreceipt.Text.Trim), 2))
        tbamount.Text = Math.Round(Decimal.Parse(tbtotalreceipt.Text.Trim), 2)
    End Sub

    Private Sub UploadReceiptBatch_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        tbdate.Text = DateTimePicker1.Value
    End Sub

    Private Sub CrystalReportViewer1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CrystalReportViewer1.Load
        Try
            For Each ctrl As Control In CrystalReportViewer1.Controls
                If TypeOf ctrl Is Windows.Forms.ToolStrip Then
                    Dim btnNew As New ToolStripButton
                    btnNew.Text = "Print Report"
                    ''btnNew.Image = Image.FromFile("../../Resources/printer.ico")
                    CType(ctrl, ToolStrip).Items.Add(btnNew)
                    AddHandler btnNew.Click, AddressOf Printlaf
                    'index = 6
                End If
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally

        End Try
    End Sub

    Private Sub Printlaf(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            crrpl.PrintToPrinter(1, False, 0, 0)
            MessageBox.Show("Done printing", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally

        End Try
    End Sub

    Public Sub loadReport(ByVal paymentmode As String)
        Dim dtcrystalreport As New DataTable

        ' start report
        dtcrystalreport.Columns.Add("PostalAddress", Type.GetType("System.String"))
        dtcrystalreport.Columns.Add("PhysicalAddress", Type.GetType("System.String"))
        dtcrystalreport.Columns.Add("ContactAddresses", Type.GetType("System.String"))
        dtcrystalreport.Columns.Add("ReceivedFromName", Type.GetType("System.String"))
        dtcrystalreport.Columns.Add("ReceivedFromIDNumber", Type.GetType("System.String"))
        dtcrystalreport.Columns.Add("ReceivedFromCellNumber", Type.GetType("System.String"))
        dtcrystalreport.Columns.Add("PayerName", Type.GetType("System.String"))
        dtcrystalreport.Columns.Add("PayerIDNumber", Type.GetType("System.String"))
        dtcrystalreport.Columns.Add("PayerCellNumber", Type.GetType("System.String"))
        dtcrystalreport.Columns.Add("Description", Type.GetType("System.String"))
        dtcrystalreport.Columns.Add("PaymentMode", Type.GetType("System.String"))
        dtcrystalreport.Columns.Add("Amount", Type.GetType("System.String"))
        dtcrystalreport.Columns.Add("ReceiptDate", Type.GetType("System.String"))
        dtcrystalreport.Columns.Add("ReceiptNo", Type.GetType("System.String"))
        dtcrystalreport.Columns.Add("ProductName", Type.GetType("System.String"))
        dtcrystalreport.Columns.Add("ProductAmount", Type.GetType("System.String"))

        PostalAddress = getPostalAddress()
        PhysicalAddress = getPhysicalAddress()
        ContactAddresses = "Tel = " & getCompanyContactNumber() & " : Fax = " & getCompanyFaxNumber() & " : Email = " & getCompanyEmailAdress()
        ReceivedFromName = tbreceivedfrom.Text.Trim
        ReceivedFromIDNumber = tbreceivedidnumber.Text.Trim
        ReceivedFromCellNumber = tbreceivedcellnumber.Text.Trim
        PayerName = "N/A"
        PayerIDNumber = "N/A"
        PayerCellNumber = "N/A"
        Description = tbdescription.Text.Trim
        paymentmode = paymentmode
        Amount = getCurrency(Decimal.Parse(tbamount.Text.Trim))
        ReceiptDate = Date.Now.Date
        ReceiptNo = getNextReceipNo()

        For Each row As DataGridViewRow In DataGridView1.Rows
            If Decimal.Parse(row.Cells(1).Value) > 0 Then
                PayerName = getCustomerName(row.Cells(0).Value)
                PayerIDNumber = row.Cells(0).Value
                ProductsName = row.Cells(5).Value
                ProductAmount = getCurrency(Decimal.Parse(row.Cells(1).Value))
                dtcrystalreport.Rows.Add(PostalAddress, PhysicalAddress, ContactAddresses, ReceivedFromName, ReceivedFromIDNumber, ReceivedFromCellNumber, PayerName, PayerIDNumber, PayerCellNumber, Description, paymentmode, Amount, ReceiptDate, ReceiptNo, ProductsName, ProductAmount)
            End If
        Next

        Dim objRpt As New ReceiptReport
        objRpt.SetDataSource(dtcrystalreport)
        Me.CrystalReportViewer1.ReportSource = objRpt
        Me.CrystalReportViewer1.DisplayBackgroundEdge = True
        Me.CrystalReportViewer1.DisplayToolbar = True
        Me.CrystalReportViewer1.DisplayGroupTree = True
        Me.CrystalReportViewer1.Refresh()
        Me.Cursor = Windows.Forms.Cursors.Default
        'Me.TabControl1.SelectTab(1)
        crrpl = objRpt

    End Sub

End Class