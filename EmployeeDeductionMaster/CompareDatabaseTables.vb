Imports System.Data.Odbc
Imports System.IO

Public Class CompareDatabaseTables

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        DataGridView5.Rows.Clear()

        loadTables(conRMOrange, "Recharge Master_Orange", DataGridView1, BindingNavigator1)
        GroupBox1.Text = "Database Tables For Recharge Master_Orange"
        loadTables(conRMBeMobile, "Recharge Master_BeMobile", DataGridView2, BindingNavigator2)
        GroupBox2.Text = "Database Tables For Recharge Master_BeMobile"

        If DataGridView1.Rows.Count > 0 And DataGridView2.Rows.Count > 0 Then
            If DataGridView1.Rows.Count = DataGridView2.Rows.Count Then
                For Each tablerow As DataGridViewRow In DataGridView1.Rows
                    If DataGridView1.Rows(tablerow.Index).Cells(0).Value.ToString() = DataGridView2.Rows(tablerow.Index).Cells(0).Value.ToString() Then
                        loadColumns(conRMOrange, DataGridView1.Rows(tablerow.Index).Cells(0).Value.ToString(), DataGridView4, BindingNavigator4)
                        loadColumns(conRMBeMobile, DataGridView2.Rows(tablerow.Index).Cells(0).Value.ToString(), DataGridView3, BindingNavigator3)
                        If DataGridView4.Rows.Count > 0 And DataGridView3.Rows.Count > 0 Then
                            If DataGridView4.Rows.Count = DataGridView3.Rows.Count Then
                                For Each columnrow As DataGridViewRow In DataGridView4.Rows
                                    If DataGridView4.Rows(columnrow.Index).Cells(0).Value.ToString() = DataGridView3.Rows(columnrow.Index).Cells(0).Value.ToString() Then
                                        If DataGridView4.Rows(columnrow.Index).Cells(1).Value.ToString() = DataGridView3.Rows(columnrow.Index).Cells(1).Value.ToString() Then

                                        Else
                                            DataGridView5.Rows.Add()
                                            DataGridView5.Rows(DataGridView5.Rows.Count - 1).Cells(0).Value = "Column datatypes not the same : " & DataGridView4.Rows(columnrow.Index).Cells(1).Value.ToString() & " : " & DataGridView3.Rows(columnrow.Index).Cells(1).Value.ToString() & " : Table " & DataGridView1.Rows(tablerow.Index).Cells(0).Value.ToString()
                                        End If
                                    Else
                                        DataGridView5.Rows.Add()
                                        DataGridView5.Rows(DataGridView5.Rows.Count - 1).Cells(0).Value = "Column names are not the same : " & DataGridView4.Rows(columnrow.Index).Cells(0).Value.ToString() & " : " & DataGridView3.Rows(columnrow.Index).Cells(0).Value.ToString() & " : Table " & DataGridView1.Rows(tablerow.Index).Cells(0).Value.ToString()
                                    End If
                                Next
                            Else
                                DataGridView5.Rows.Add()
                                DataGridView5.Rows(DataGridView5.Rows.Count - 1).Cells(0).Value = "Missing Column on table " & DataGridView1.Rows(tablerow.Index).Cells(0).Value.ToString()
                            End If
                        Else
                            DataGridView5.Rows.Add()
                            DataGridView5.Rows(DataGridView5.Rows.Count - 1).Cells(0).Value = "Table names are not the same : " & DataGridView1.Rows(tablerow.Index).Cells(0).Value.ToString() & " : " & DataGridView2.Rows(tablerow.Index).Cells(0).Value.ToString()
                        End If
                    End If
                Next
            Else
                For Each row As DataGridViewRow In DataGridView1.Rows
                    If DataGridView1.Rows(row.Index).Cells(0).Value.ToString() = DataGridView2.Rows(row.Index).Cells(0).Value.ToString() Then

                    Else
                        DataGridView5.Rows.Add()
                        DataGridView5.Rows(DataGridView5.Rows.Count - 1).Cells(0).Value = DataGridView1.Rows(row.Index).Cells(0).Value.ToString() & "  |  " & DataGridView2.Rows(row.Index).Cells(0).Value.ToString() & " : There is a missing table in one of the tables"
                        Exit For
                    End If
                Next
            End If
        End If
        MessageBox.Show("Done", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub

    Private Sub loadTables(ByVal con As SqlClient.SqlConnection, ByVal Database As String, ByVal dgv As DataGridView, ByVal bn As BindingNavigator)
        dgv.DataSource = Nothing

        Dim qry As String = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_CATALOG='" & Database & "' order by TABLE_NAME"
        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)
        Dim dt As DataTable

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim ctr As Integer = 0
            SQLAdapter.Fill(ds, "IDs")
            If ds.Tables(0).Rows.Count > 0 Then
                dt = ds.Tables(0)
                Dim bs As New BindingSource
                bs.DataSource = dt.DefaultView
                bn.BindingSource = bs
                dgv.DataSource = bs
                dgv.Columns(dgv.ColumnCount - 1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

            Else
                dgv.DataSource = Nothing
            End If
        Catch ex As Exception
            MessageBox.Show(ex.StackTrace, "loadTables()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub loadColumns(ByVal con As SqlClient.SqlConnection, ByVal table As String, ByVal dgv As DataGridView, ByVal bn As BindingNavigator)
        dgv.DataSource = Nothing

        Dim qry As String = "select COLUMN_NAME,DATA_TYPE from INFORMATION_SCHEMA.COLUMNS where TABLE_NAME='" & table & "'"
        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)
        Dim dt As DataTable

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim ctr As Integer = 0
            SQLAdapter.Fill(ds, "IDs")
            If ds.Tables(0).Rows.Count > 0 Then
                dt = ds.Tables(0)
                Dim bs As New BindingSource
                bs.DataSource = dt.DefaultView
                bn.BindingSource = bs
                dgv.DataSource = bs
                dgv.Columns(dgv.ColumnCount - 1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

            Else
                dgv.DataSource = Nothing
            End If
        Catch ex As Exception
            MessageBox.Show(ex.StackTrace, "loadTables()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub


    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        DataGridView5.Rows.Clear()

        loadTables(conMLMAdvances, "MicroLoansMaster_new_MLM", DataGridView1, BindingNavigator1)
        GroupBox1.Text = "Database Tables For MicroLoansMaster_new_MLM"
        loadTables(conMLMEducational, "MicroLoansMaster_Educational", DataGridView2, BindingNavigator2)
        GroupBox2.Text = "Database Tables For MicroLoansMaster_Educational"

        If DataGridView1.Rows.Count > 0 And DataGridView2.Rows.Count > 0 Then
            If DataGridView1.Rows.Count = DataGridView2.Rows.Count Then
                For Each tablerow As DataGridViewRow In DataGridView1.Rows
                    If DataGridView1.Rows(tablerow.Index).Cells(0).Value.ToString() = DataGridView2.Rows(tablerow.Index).Cells(0).Value.ToString() Then
                        loadColumns(conMLMAdvances, DataGridView1.Rows(tablerow.Index).Cells(0).Value.ToString(), DataGridView4, BindingNavigator4)
                        loadColumns(conMLMEducational, DataGridView2.Rows(tablerow.Index).Cells(0).Value.ToString(), DataGridView3, BindingNavigator3)
                        If DataGridView4.Rows.Count > 0 And DataGridView3.Rows.Count > 0 Then
                            If DataGridView4.Rows.Count = DataGridView3.Rows.Count Then
                                For Each columnrow As DataGridViewRow In DataGridView4.Rows
                                    If DataGridView4.Rows(columnrow.Index).Cells(0).Value.ToString() = DataGridView3.Rows(columnrow.Index).Cells(0).Value.ToString() Then
                                        If DataGridView4.Rows(columnrow.Index).Cells(1).Value.ToString() = DataGridView3.Rows(columnrow.Index).Cells(1).Value.ToString() Then

                                        Else
                                            DataGridView5.Rows.Add()
                                            DataGridView5.Rows(DataGridView5.Rows.Count - 1).Cells(0).Value = "Column datatypes not the same : " & DataGridView4.Rows(columnrow.Index).Cells(1).Value.ToString() & " : " & DataGridView3.Rows(columnrow.Index).Cells(1).Value.ToString() & " : Table " & DataGridView1.Rows(tablerow.Index).Cells(0).Value.ToString()
                                        End If
                                    Else
                                        DataGridView5.Rows.Add()
                                        DataGridView5.Rows(DataGridView5.Rows.Count - 1).Cells(0).Value = "Column names are not the same : " & DataGridView4.Rows(columnrow.Index).Cells(0).Value.ToString() & " : " & DataGridView3.Rows(columnrow.Index).Cells(0).Value.ToString() & " : Table " & DataGridView1.Rows(tablerow.Index).Cells(0).Value.ToString()
                                    End If
                                Next
                            Else
                                DataGridView5.Rows.Add()
                                DataGridView5.Rows(DataGridView5.Rows.Count - 1).Cells(0).Value = "Missing Column on table " & DataGridView1.Rows(tablerow.Index).Cells(0).Value.ToString()
                            End If
                        Else
                            DataGridView5.Rows.Add()
                            DataGridView5.Rows(DataGridView5.Rows.Count - 1).Cells(0).Value = "Table names are not the same : " & DataGridView1.Rows(tablerow.Index).Cells(0).Value.ToString() & " : " & DataGridView2.Rows(tablerow.Index).Cells(0).Value.ToString()
                        End If
                    End If
                Next
            Else
                For Each row As DataGridViewRow In DataGridView1.Rows
                    If DataGridView1.Rows(row.Index).Cells(0).Value.ToString() = DataGridView2.Rows(row.Index).Cells(0).Value.ToString() Then

                    Else
                        DataGridView5.Rows.Add()
                        DataGridView5.Rows(DataGridView5.Rows.Count - 1).Cells(0).Value = DataGridView1.Rows(row.Index).Cells(0).Value.ToString() & "  |  " & DataGridView2.Rows(row.Index).Cells(0).Value.ToString() & " : There is a missing table in one of the tables"
                        Exit For
                    End If
                Next
            End If
        End If
        MessageBox.Show("Done", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
End Class