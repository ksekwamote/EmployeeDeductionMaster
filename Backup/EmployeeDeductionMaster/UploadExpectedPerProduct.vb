Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc

Public Class UploadExpectedPerProduct

    Dim customerid As String

    Private Sub Upload_Last_Expected_Deductions_Per_Product_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If decision("Do you want to load the deductions ?") = DialogResult.Yes Then
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
                    If checkCustomerDetails(row.Cells(0).Value.ToString, employerid) = True Then

                    Else
                        If row.Cells(15).Value = "" Then
                            row.Cells(15).Value = "Customer Does Not Exists"
                        Else
                            row.Cells(15).Value = row.Cells(15).Value & "/ Customer Does Not Exists"
                        End If
                        Button2.Enabled = False
                    End If
                Next

                MessageBox.Show(counter, "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            If decision("YES = Save and Export Expected Deductions, NO = Export without Saving") = DialogResult.Yes Then
                If DataGridView1.Rows.Count > 0 Then
                    ProgressBar1.Maximum = DataGridView1.Rows.Count - 1
                    For Each row As DataGridViewRow In DataGridView1.Rows
                        For Each col As DataGridViewColumn In DataGridView1.Columns
                            If col.Index = 0 Or col.Index = 1 Or col.Index = 15 Then

                            Else
                                customerid = row.Cells(0).Value
                                If Decimal.Parse(row.Cells(col.Index).Value) > 0 Then
                                    saveExpectedDeductionsPerProduct(row.Cells(0).Value.ToString, Decimal.Parse(row.Cells(col.Index).Value), dtpformonth.Value, Date.Now.Date, Integer.Parse(col.HeaderText.Trim.ToString), employerid)
                                End If
                            End If
                        Next
                    Next
                    ProgressBar1.Value = 0
                    MessageBox.Show("Done Saving", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show("Nothing to Save", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            Else
                MessageBox.Show("Nothing done", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch
            MessageBox.Show(customerid, "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub dtpformonth_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpformonth.ValueChanged
        getLastDateOfMonth(dtpformonth.Value)
    End Sub
End Class