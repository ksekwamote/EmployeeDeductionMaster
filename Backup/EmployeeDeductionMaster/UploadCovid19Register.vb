Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc

Public Class UploadCovid19Register

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        DataGridView1.AllowUserToAddRows = False
        Button1.Enabled = False
        tdpcaptureddate.Enabled = True

        If decision("Do you want to load the employers ?") = DialogResult.Yes Then
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
        Try
            If decision("Is the date correct : " & tdpcaptureddate.Value & " ?") = DialogResult.Yes Then
                If decision("Are you sure you want to proceed?") = DialogResult.Yes Then
                    If DataGridView1.Rows.Count > 0 Then
                        For Each row As DataGridViewRow In DataGridView1.Rows
                            saveCovidDetails(row.Cells(0).Value.ToString, row.Cells(1).Value.ToString, row.Cells(2).Value.ToString, row.Cells(3).Value.ToString, row.Cells(4).Value.ToString, row.Cells(5).Value.ToString, Decimal.Parse(row.Cells(6).Value), tdpcaptureddate.Value, tdpcaptureddate.Value, Date.Now, Date.Now, logedinname, row.Cells(7).Value.ToString)
                        Next
                        MessageBox.Show("Done saving", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Close()
                    End If
                Else
                    MessageBox.Show("Nothing to save", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
           
        End Try

    End Sub

    Private Sub tdpcaptureddate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdpcaptureddate.ValueChanged
        Button2.Enabled = True
    End Sub
End Class