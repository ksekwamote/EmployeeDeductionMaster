Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports CrystalDecisions.ReportSource

Public Class BankingReport

    Dim crrpsl As New ReportDocument
    Dim dt As New DataTable

    Public br As BankingReceipts
    Public BankingBatch As String

    Dim PayerID, Amount, TransactionDate, ModeName, CapturedBy, CapturedTime, BatchName As String
    Dim totalamount As Decimal

    Public Sub displayReport()

        dt.Columns.Add("PayerID", Type.GetType("System.String"))
        dt.Columns.Add("Amount", Type.GetType("System.String"))
        dt.Columns.Add("TransactionDate", Type.GetType("System.String"))
        dt.Columns.Add("ModeName", Type.GetType("System.String"))
        dt.Columns.Add("CapturedBy", Type.GetType("System.String"))
        dt.Columns.Add("CapturedTime", Type.GetType("System.String"))
        dt.Columns.Add("BatchName", Type.GetType("System.String"))
        dt.Columns.Add("BankingBatch", Type.GetType("System.String"))

        totalamount = 0
        For Each row As DataGridViewRow In br.DataGridView1.Rows
            PayerID = row.Cells(0).Value & " (" & getCustomerName(row.Cells(0).Value) & ")"
            Amount = row.Cells(1).Value
            TransactionDate = row.Cells(2).Value
            ModeName = row.Cells(3).Value
            CapturedBy = row.Cells(4).Value
            CapturedTime = row.Cells(5).Value
            BatchName = row.Cells(6).Value

            dt.Rows.Add(PayerID, getCurrency(Decimal.Parse(Amount)), TransactionDate, ModeName, CapturedBy, CapturedTime, BatchName, BankingBatch)
            totalamount = totalamount + Decimal.Parse(row.Cells(1).Value)
        Next

        dt.Rows.Add("__________________________________________", "__________________________________________", "__________________________________________", "__________________________________________", "__________________________________________", "__________________________________________", "__________________________________________", "__________________________________________")
        dt.Rows.Add("Total Amount", getCurrency(totalamount), "", "", "", "", "", "")
        dt.Rows.Add("", "____________________________________________________________________", "", "", "", "", "", "")
        ''dt.Rows.Add("", "", "", "", "", "", "", "")
        ''dt.Rows.Add("Confirmed By :_____________________________________", "_________________________________________________", "_________________________________________________", "Date :_____________________________________", "_____________________________________________________", "", "", "")
        ''dt.Rows.Add("Collected By :_____________________________________", "_________________________________________________", "_________________________________________________", "Date :_____________________________________", "_____________________________________________________", "", "", "")

        Dim objRpt As New CRBanking
        objRpt.SetDataSource(dt)
        Me.CrystalReportViewer1.ReportSource = objRpt
        Me.CrystalReportViewer1.DisplayBackgroundEdge = True
        Me.CrystalReportViewer1.DisplayToolbar = True
        Me.CrystalReportViewer1.DisplayGroupTree = True
        Me.CrystalReportViewer1.Refresh()
        Me.Cursor = Windows.Forms.Cursors.Default
        crrpsl = objRpt
    End Sub

    Private Sub BankingReport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            displayReport()

            For Each ctrl As Control In CrystalReportViewer1.Controls
                If TypeOf ctrl Is Windows.Forms.ToolStrip Then
                    Dim btnNew As New ToolStripButton
                    btnNew.Text = "Print Report"
                    ''btnNew.Image = Image.FromFile("../../Resources/printer.ico")
                    CType(ctrl, ToolStrip).Items.Add(btnNew)
                    AddHandler btnNew.Click, AddressOf Printdomf
                    'index = 5
                End If
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally

        End Try
    End Sub

    Private Sub Printdomf(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            If decision("Are you sure you want to print ?") = DialogResult.Yes Then
                crrpsl.PrintToPrinter(1, False, 0, 0)
                br.Close()
                Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally

        End Try
    End Sub

End Class