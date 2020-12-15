Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data.Odbc
Imports System.IO

Public Class UploadLocations

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
                If checkLocationID("select Count(LocationID) from Location where LocationID = '" & row.Cells(1).Value & "'", conRMOrange) = False Then
                    If row.Cells(4).Value = "" Then
                        row.Cells(4).Value = "Location of the ID (" & row.Cells(1).Value & ") Does Not Exist RM Orange"
                    Else
                        row.Cells(4).Value = row.Cells(4).Value & "/ Location of the ID (" & row.Cells(1).Value & ") Does Not Exist RM Orange"
                    End If
                    Button2.Enabled = False
                End If
                If checkLocationID("select Count(LocationID) from Location where LocationID = '" & row.Cells(2).Value & "'", conRMBeMobile) = False Then
                    If row.Cells(4).Value = "" Then
                        row.Cells(4).Value = "Location of the ID (" & row.Cells(2).Value & ") Does Not Exist RM Bemobile"
                    Else
                        row.Cells(4).Value = row.Cells(4).Value & "/ Location of the ID (" & row.Cells(2).Value & ") Does Not Exist RM Bemobile"
                    End If
                    Button2.Enabled = False
                End If
                If checkLocationID("select Count(AreaCode) from Area where AreaCode = '" & row.Cells(3).Value & "'", conFPM) = False Then
                    If row.Cells(4).Value = "" Then
                        row.Cells(4).Value = "Location of the ID (" & row.Cells(3).Value & ") Does Not Exist FPM"
                    Else
                        row.Cells(4).Value = row.Cells(4).Value & "/ Location of the ID (" & row.Cells(3).Value & ") Does Not Exist FPM"
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
                saveLocation(getNextLocationID(row.Cells(0).Value.Substring(0, 3)), row.Cells(0).Value.ToString(), True, row.Cells(1).Value.ToString(), row.Cells(2).Value.ToString(), row.Cells(3).Value.ToString())
            Next
            MessageBox.Show("Location Uploaded", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Close()
        End If
    End Sub

    Public Function getNextLocationID(ByVal str As String) As String
        Dim qry As String = "select * from Location where LocationID like '" & str & "%' order by LocationID DESC"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "MS")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                Dim int As Integer = 0
                int = Integer.Parse(ds.Tables(0).Rows(0).Item(0).ToString.Substring(3)) + 1
                Dim go As Boolean = False
                While go = False
                    perc = str & int
                    If CheckID("select * from Location where LocationID = '" & perc & "'") = True Then
                        int = int + 1
                    Else
                        go = True
                    End If
                End While
            Else
                perc = str & "1"
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "getNextLocationID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getNextLocationID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Sub saveLocation(ByVal LocationID As String, ByVal LocationName As String, ByVal Status As String, ByVal RMOLocation As String, ByVal RMBLocation As String, ByVal FPMLocation As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try
            Dim qryinsert As String = "Insert into Location values ('" & LocationID & "','" & LocationName & "','" & Status & "','" & RMOLocation & "','" & RMBLocation & "','" & FPMLocation & "')"
            Dim cmdinsert As New SqlClient.SqlCommand(qryinsert, con)
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            transactioninsert = con.BeginTransaction()
            cmdinsert.Transaction = transactioninsert

            If cmdinsert.ExecuteNonQuery() Then
                transactioninsert.Commit()
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "saveLocation()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "saveLocation()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub


End Class