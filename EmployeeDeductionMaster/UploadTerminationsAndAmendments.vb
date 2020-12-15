Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc

Public Class UploadTerminationsAndAmendments

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        DataGridView1.AllowUserToAddRows = False
        Button1.Enabled = False

        If decision("Do you want to load data ?") = DialogResult.Yes Then
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
            End If
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If decision("Are you sure you want to proceed?") = DialogResult.Yes Then
            If DataGridView1.Rows.Count > 0 Then
                ProgressBar1.Maximum = DataGridView1.Rows.Count - 1
                For Each row As DataGridViewRow In DataGridView1.Rows
                    If CheckBox1.Checked = True Then
                        saveAmendments(row.Cells(0).Value, Decimal.Parse(row.Cells(1).Value.ToString), Date.Parse(row.Cells(2).Value), Date.Parse(row.Cells(3).Value), Integer.Parse(row.Cells(4).Value), Integer.Parse(row.Cells(5).Value))
                    Else
                        saveExpectedDeductionsPerProduct(row.Cells(0).Value, Decimal.Parse(row.Cells(1).Value.ToString), Date.Parse(row.Cells(2).Value), Date.Parse(row.Cells(3).Value), Integer.Parse(row.Cells(4).Value), Integer.Parse(row.Cells(5).Value))
                    End If
                    ProgressBar1.Value = row.Index
                Next
                ProgressBar1.Value = 0
                MessageBox.Show("All amendments has been saved successfully", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Else
            MessageBox.Show("Nothing to save", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            CheckBox1.Checked = True
            CheckBox2.Checked = False
        ElseIf CheckBox1.Checked = False Then
            CheckBox1.Checked = False
            CheckBox2.Checked = True
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            CheckBox2.Checked = True
            CheckBox1.Checked = False
        ElseIf CheckBox2.Checked = False Then
            CheckBox2.Checked = False
            CheckBox1.Checked = True
        End If
    End Sub

    Private Sub UploadTerminationsAndAmendments_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class