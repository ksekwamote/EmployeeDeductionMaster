Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data.Odbc
Imports System.IO

Public Class UploadDepartments

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
                    If CheckDepartmentName(row.Cells(0).Value.ToString()) = True Then
                        If row.Cells(1).Value = "" Then
                            row.Cells(1).Value = "Department Code Exists"
                        Else
                            row.Cells(1).Value = row.Cells(1).Value & "/ Department Code Exists"
                        End If
                        Button2.Enabled = False
                    End If
                Next

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally

        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If decision("Are you to proceed ?") = DialogResult.Yes Then

            For Each row As DataGridViewRow In DataGridView1.Rows
                TextBox1.Text = getNextDepartmentID(row.Cells(0).Value.ToString.Substring(0, 3))
                saveDepartment(getNextDepartmentID(row.Cells(0).Value.ToString.Substring(0, 3)), row.Cells(0).Value, "Active", row.Cells(2).Value, row.Cells(3).Value, row.Cells(4).Value, row.Cells(5).Value, row.Cells(6).Value)
            Next
            MessageBox.Show("Department Uploaded", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Close()
        End If
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        ExportToExcel(DataGridView1, "")
    End Sub
End Class