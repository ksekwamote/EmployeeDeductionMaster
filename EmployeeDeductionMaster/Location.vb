Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports DgvFilterPopup

Public Class Location

    Dim selected As Boolean = False

    Private Sub loadLocation()
        DataGridView1.DataSource = Nothing
        Dim qry As String = "select * from Location Order By LocationName ASC"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)
        Dim dt As DataTable

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim ctr As Integer = 0
            SQLAdapter.Fill(ds, "DIM")
            If ds.Tables(0).Rows.Count > 0 Then
                dt = ds.Tables(0)
                Dim bs As New BindingSource
                bs.DataSource = dt.DefaultView
                DataGridView1.DataSource = bs
                BindingNavigator1.BindingSource = bs
                DataGridView1.Columns(DataGridView1.ColumnCount - 1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                Dim filterManager As New DgvFilterManager
                filterManager = New DgvFilterManager(DataGridView1)
            Else
                DataGridView1.DataSource = Nothing
                BindingNavigator1.BindingSource = Nothing
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "loadLocation()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadLocation()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If tblocationid.Text.Trim = "" Or tblocation.Text.Trim = "" Or tbrmoid.Text.Trim = "" Or tbrmbid.Text.Trim = "" Or tbfpmid.Text.Trim = "" Then
                MessageBox.Show("Enter all the details before saving", "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                If decision("Are you sure to proceed ?") = DialogResult.Yes Then
                    saveLocation(tblocationid.Text.Trim, tblocation.Text.Trim, "Active", tbrmoid.Text.Trim, tbrmbid.Text.Trim, tbfpmid.Text.Trim)
                    loadLocation()
                    clearText()
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally

        End Try
    End Sub

    Private Sub DataGridView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Click
        If DataGridView1.Rows.Count > 0 Then
            selected = True
            tblocationid.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(0).Value.ToString()
            tblocation.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(1).Value.ToString()
            tbrmoid.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(3).Value.ToString()
            tbrmo.Text = getValueStringOtherSystems("select LocationName from Location where LocationID = '" & tbrmoid.Text.Trim & "'", conRMOrange)
            tbrmbid.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(4).Value.ToString()
            tbrmb.Text = getValueStringOtherSystems("select LocationName from Location where LocationID = '" & tbrmbid.Text.Trim & "'", conRMBeMobile)
            tbfpmid.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(5).Value.ToString()
            tbfpm.Text = getValueStringOtherSystems("select AreaName from Area where AreaCode = '" & tbfpmid.Text.Trim & "'", conFPM)
            Button2.Enabled = True
            Button1.Enabled = False
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            If tblocationid.Text = "" Or tblocation.Text.Trim = "" Or tbrmoid.Text.Trim = "" Or tbrmbid.Text.Trim = "" Or tbfpmid.Text.Trim = "" Then
                MessageBox.Show("Make Sure to select the location to update changes and enter the changes", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                If decision("Are you sure to proceed ?") = DialogResult.Yes Then
                    updateLocation(tblocationid.Text.Trim, tblocation.Text.Trim, tbrmoid.Text.Trim, tbrmbid.Text.Trim, tbfpmid.Text.Trim)
                    loadLocation()
                    clearText()
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally

        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        clearText()
    End Sub

    Public Sub clearText()
        Button2.Enabled = False
        Button1.Enabled = True
        tblocationid.Text = ""
        tblocation.Text = ""
        selected = False

        tbrmoid.Text = ""
        tbrmo.Text = ""
        tbrmbid.Text = ""
        tbrmb.Text = ""
        tbfpmid.Text = ""
        tbfpm.Text = ""
    End Sub


    Private Sub tbbankname_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tblocation.TextChanged
        If Button1.Enabled = True Then
            If tblocation.Text.Trim.Length > 3 Then
                If selected = True Then

                Else
                    tblocationid.Text = getNextLocationID(tblocation.Text.Substring(0, 3))
                End If
            End If
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

    Public Sub updateLocation(ByVal LocationID As String, ByVal LocationName As String, ByVal RMOLocation As String, ByVal RMBLocation As String, ByVal FPMLocation As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update Location set LocationName = '" & LocationName & "',RMOLocation = '" & RMOLocation & "',RMBLocation = '" & RMBLocation & "',FPMLocation = '" & FPMLocation & "' where LocationID = '" & LocationID & "'"
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
            MessageBox.Show(ex.Message, "updateLocation()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateLocation()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Private Sub Location_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadLocation()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim vrl As New LocationList
        vrl.MdiParent = mform
        vrl.whichform = "RMO"
        vrl.loc = Me
        vrl.Show()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim vrl As New LocationList
        vrl.MdiParent = mform
        vrl.whichform = "RMB"
        vrl.loc = Me
        vrl.Show()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim vrl As New LocationList
        vrl.MdiParent = mform
        vrl.whichform = "FPM"
        vrl.loc = Me
        vrl.Show()
    End Sub
End Class