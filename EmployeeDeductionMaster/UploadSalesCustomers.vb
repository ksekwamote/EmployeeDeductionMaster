Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc

Public Class UploadSalesCustomers

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        DataGridView1.AllowUserToAddRows = False
        Button1.Enabled = False
        Button2.Enabled = True

        If decision("Do you want to load the sales agents customers ?") = DialogResult.Yes Then
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
                    If getValueBooleanOtherSystems("select * from SalesAgents where IDNumber = '" & row.Cells(4).Value.ToString & "'", con) = False Then
                        If row.Cells(5).Value = "" Then
                            row.Cells(5).Value = "Sales Agent of ID (" & row.Cells(4).Value & ") Does Not Exist"
                        Else
                            row.Cells(5).Value = row.Cells(5).Value & "/ Sales Agent of ID (" & row.Cells(4).Value & ") Does Not Exist"
                        End If
                        Button2.Enabled = False
                    End If

                    If getValueBooleanOtherSystems("select * from SalesCustomers where CustomerID = '" & row.Cells(0).Value.ToString & "'", con) = True Then
                        If row.Cells(5).Value = "" Then
                            row.Cells(5).Value = "Customer of ID (" & row.Cells(0).Value & ") already exists"
                        Else
                            row.Cells(5).Value = row.Cells(5).Value & "/ Customer of ID (" & row.Cells(0).Value & ") already exists"
                        End If
                        Button2.Enabled = False
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If decision("Are you sure you want to proceed?") = DialogResult.Yes Then
            If DataGridView1.Rows.Count > 0 Then
                For Each row As DataGridViewRow In DataGridView1.Rows
                    saveSalesCustomers(row.Cells(0).Value.ToString, row.Cells(1).Value.ToString, row.Cells(2).Value.ToString, Date.Parse(row.Cells(3).Value.ToString), row.Cells(4).Value.ToString, logedinname, Date.Now, "Active", Date.Parse(row.Cells(3).Value.ToString).AddMonths(getSalesCustomerValidity()))
                Next
                MessageBox.Show("Done saving", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Close()
            End If
        Else
            MessageBox.Show("Nothing to save", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        ExportToExcel(DataGridView1, "")
    End Sub
End Class