Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data.Odbc
Imports System.IO

Public Class UploadBankBranches

    Private Sub UploadLocations_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

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
            End If

            For Each row As DataGridViewRow In DataGridView1.Rows
                If getValueBooleanOtherSystems("select * from Bank where BankID = '" & row.Cells(2).Value & "'", con) = False Then
                    If row.Cells(2).Value = "" Then
                        row.Cells(2).Value = "Bank of the ID (" & row.Cells(2).Value & ") Does Not Exist EDM"
                    Else
                        row.Cells(2).Value = row.Cells(2).Value & "/ Bank of the ID (" & row.Cells(2).Value & ") Does Not Exist EDM"
                    End If
                    Button2.Enabled = False
                End If
            Next

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally

        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If decision("Are you to proceed ?") = DialogResult.Yes Then
            For Each row As DataGridViewRow In DataGridView1.Rows
                saveBankBranch(getNextBankBranchID(row.Cells(0).Value.ToString().Substring(0, 3) & getValueString("select BankName from Bank where BankID = '" & row.Cells(2).Value.ToString() & "'").Substring(0, 3)), row.Cells(0).Value.ToString(), row.Cells(2).Value.ToString(), "Active", row.Cells(1).Value.ToString())
            Next
            MessageBox.Show("Bank Branches Uploaded", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Close()
        End If
    End Sub

End Class