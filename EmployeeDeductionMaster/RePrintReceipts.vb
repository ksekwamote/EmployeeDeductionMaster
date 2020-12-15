Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports CrystalDecisions.ReportSource
Public Class RePrintReceipts

    Dim PostalAddress, PhysicalAddress, ContactAddresses, ReceivedFromName, ReceivedFromIDNumber, ReceivedFromCellNumber, PayerName, PayerIDNumber, PayerCellNumber, Description, PaymentMode, Amount, ReceiptDate, ReceiptNo, ProductsName, ProductAmount As String
    Dim crrpl As New ReportDocument
    Dim ds As New DataSet
    Public BatchName As String
    Public rr As ReceiptsReport
    Private Sub RePrintReceipts_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        loadReport()
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

    Public Sub loadReport()
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

        ds.Clear()
        ds = getDataset("select * from ReceiptPayerInfor where BatchName = '" & batchname & "'", con)
        If ds.Tables(0).Rows.Count > 0 Then
            ReceivedFromName = ds.Tables(0).Rows(0).Item(4).ToString
            ReceivedFromIDNumber = ds.Tables(0).Rows(0).Item(3).ToString
            ReceivedFromCellNumber = ds.Tables(0).Rows(0).Item(5).ToString
            PayerName = ds.Tables(0).Rows(0).Item(8).ToString
            PayerIDNumber = ds.Tables(0).Rows(0).Item(1).ToString
            PayerCellNumber = ds.Tables(0).Rows(0).Item(2).ToString
            Description = ds.Tables(0).Rows(0).Item(6).ToString
        Else
            ReceivedFromName = ""
            ReceivedFromIDNumber = ""
            ReceivedFromCellNumber = ""
            PayerName = ""
            PayerIDNumber = ""
            PayerCellNumber = ""
            Description = ""
        End If

        ds.Clear()
        ds = getDataset("select ModeName from ModeOfPayment where ModeID in (select ModeID from Receipts where BatchName = '" & BatchName & "')", con)
        PaymentMode = ds.Tables(0).Rows(0).Item(0).ToString

        ds.Clear()
        ds = getDataset("select * from Receipts where BatchName = '" & BatchName & "'", con)
        Amount = getCurrency(Decimal.Parse(ds.Tables(0).Rows(0).Item(2).ToString))

        ReceiptDate = ds.Tables(0).Rows(0).Item(4).ToString
        ReceiptNo = ds.Tables(0).Rows(0).Item(15).ToString & " : Re-Print"

        For Each row As DataGridViewRow In rr.DataGridView2.Rows
            PayerName = getCustomerName(row.Cells(0).Value)
            PayerIDNumber = row.Cells(0).Value
            ProductsName = row.Cells(5).Value
            ProductAmount = getCurrency(Decimal.Parse(row.Cells(1).Value))
            dtcrystalreport.Rows.Add(PostalAddress, PhysicalAddress, ContactAddresses, ReceivedFromName, ReceivedFromIDNumber, ReceivedFromCellNumber, PayerName, PayerIDNumber, PayerCellNumber, Description, PaymentMode, Amount, ReceiptDate, ReceiptNo, ProductsName, ProductAmount)
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

    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
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
End Class