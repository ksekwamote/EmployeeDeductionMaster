Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc

Public Class AddUser

    Dim useravailable As String
    Dim choice As Integer
    Dim wrapper As New Simple3Des("ThithoH")
    Public logedinname As String
    Dim dsrs As New DataSet

    Public Sub decision(ByVal str As String)
        choice = MessageBox.Show(str, "Decision", MessageBoxButtons.YesNo)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        decision("Do you want to add the new User?")
        If choice = DialogResult.No Then
            'MessageBox.Show("No pressed")
        ElseIf choice = DialogResult.Yes Then
            If tbusername.Text = "" Or tbpassword.Text = "" Or tbusertype.Text = "" Or tbname.Text = "" Or tbsupervisorid.Text = "" Or tbcompanybranchid.Text = "" Then
                MessageBox.Show("Enter all the details first", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                Dim cipherText As String = wrapper.EncryptData(tbpassword.Text.Trim)
                saveUser(tbusername.Text.Trim, cipherText, tbusertype.Text, tbname.Text.Trim, Integer.Parse(tbsupervisorid.Text.Trim), Integer.Parse(tbcompanybranchid.Text.Trim))
                MessageBox.Show("Details saved", "saveUser()", MessageBoxButtons.OK, MessageBoxIcon.Information)
                wrapper.saveAction("New User added", Date.Now, logedinname, tbusername.Text)
            End If
        End If
    End Sub


    Public Sub updateUser()
        Dim wrapper As New Simple3Des("ThithoH")
        Dim cipherText As String = wrapper.EncryptData(tbpassword.Text.Trim)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try
            If tbusername.Text.Trim = "" Or tbpassword.Text.Trim = "" Or tbusertype.Text = "" Or tbname.Text.Trim = "" Or tbsupervisorid.Text.Trim = "" Then
                MessageBox.Show("Enter all the details before saving")
            Else
                getUser(tbusername.Text)
                If useravailable = "Yes" Then
                    If con.State = ConnectionState.Closed Then
                        con.Open()
                    End If
                    transactioninsert = con.BeginTransaction()
                    Dim qryinsert As String = "Update Users set Username = '" & tbusername.Text.Trim & "',PasswordValue = '" & cipherText & "',UserType = '" & tbusertype.Text & "',Name = '" & tbname.Text.Trim & "',SuperviorsID = '" & Integer.Parse(tbsupervisorid.Text.Trim) & "' where Username = '" & tbusername.Text.Trim & "'"
                    Dim cmdinsert As New SqlClient.SqlCommand(qryinsert, con)
                    cmdinsert.Transaction = transactioninsert
                    If cmdinsert.ExecuteNonQuery() Then
                        transactioninsert.Commit()
                    End If
                    
                Else
                    MessageBox.Show("Make sure the username is correct", "updateUser()", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "updateUser()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateUser()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Private Function getUser(ByVal str As String) As DataSet
        Dim qry As String = "select * from Users inner join CompanyBranches on Users.CompanyBranchID = CompanyBranches.AutoID where Username = '" & str & "'"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)
        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim ctr As Integer = 0
            SQLAdapter.Fill(ds, "Users")
            If ds.Tables(0).Rows.Count > 0 Then
                useravailable = "Yes"
            Else
                useravailable = "No"
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getUser()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return ds
    End Function

    Private Sub AddUser_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadSupervisors()
        tbsupervisor.Text = ""
        tbsupervisorid.Text = ""
        loadBranches()
        tbbranch.Text = ""
        tbcompanybranchid.Text = ""
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        tbusertype.Text = ComboBox1.Text
    End Sub

    Public Sub loadSupervisors()
        Dim qry As String = "select AutoID,SupervisorName from Supervisors"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim ctr As Integer = 0
            SQLAdapter.Fill(ds, "Emp")
            If ds.Tables(0).Rows.Count > 0 Then
                Me.ComboBox2.DataSource = ds
                Me.ComboBox2.DisplayMember = "Emp.SupervisorName"
                Me.ComboBox2.ValueMember = "Emp.AutoID"
                Me.ComboBox2.SelectedIndex = 0
            Else

            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "loadSupervisors()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        tbsupervisor.Text = ComboBox2.Text
        tbsupervisorid.Text = ComboBox2.SelectedValue.ToString
    End Sub

    Private Sub tbusername_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbusername.Leave
        fillDetails()
    End Sub

    Public Sub fillDetails()
        dsrs = getUser(tbusername.Text)
        If dsrs.Tables(0).Rows.Count > 0 Then
            tbusername.Text = dsrs.Tables(0).Rows(0).Item(0).ToString
            tbusertype.Text = dsrs.Tables(0).Rows(0).Item(2).ToString
            ComboBox1.SelectedIndex = ComboBox1.FindStringExact(tbusertype.Text)
            tbname.Text = dsrs.Tables(0).Rows(0).Item(3).ToString
            tbsupervisorid.Text = dsrs.Tables(0).Rows(0).Item(4).ToString
            tbsupervisor.Text = getSupervisorName(Integer.Parse(tbsupervisorid.Text.Trim))
            ComboBox2.SelectedIndex = ComboBox2.FindStringExact(tbsupervisor.Text)
            tbcompanybranchid.Text = dsrs.Tables(0).Rows(0).Item(5).ToString
            tbbranch.Text = dsrs.Tables(0).Rows(0).Item(7).ToString
            cbbranch.SelectedIndex = cbbranch.FindStringExact(tbbranch.Text)
            Button1.Enabled = False
            Button2.Enabled = True
        Else
            'tbusername.Text = ""
            tbusertype.Text = ""
            tbname.Text = ""
            tbsupervisorid.Text = ""
            tbsupervisor.Text = ""
            tbbranch.Text = ""
            tbcompanybranchid.Text = ""
            Button1.Enabled = True
            Button2.Enabled = False
        End If
    End Sub

    Public Sub loadBranches()
        Dim qry As String = "select * from CompanyBranches"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim ctr As Integer = 0
            SQLAdapter.Fill(ds, "Emp")
            If ds.Tables(0).Rows.Count > 0 Then
                Me.cbbranch.DataSource = ds
                Me.cbbranch.DisplayMember = "Emp.BranchName"
                Me.cbbranch.ValueMember = "Emp.AutoID"
                Me.cbbranch.SelectedIndex = 0
            Else

            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "loadBranches()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub cbbranch_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbbranch.SelectedIndexChanged
        Try
            tbbranch.Text = cbbranch.Text
            tbcompanybranchid.Text = cbbranch.SelectedValue.ToString
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If tbusername.Text = "" Or tbpassword.Text = "" Or tbusertype.Text = "" Or tbname.Text = "" Or tbsupervisorid.Text = "" Or tbcompanybranchid.Text = "" Then
            MessageBox.Show("Enter all the details first", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            updateUser()
            MessageBox.Show("Details saved", "updateUser()", MessageBoxButtons.OK, MessageBoxIcon.Information)
            wrapper.saveAction("User details updated", Date.Now, logedinname, tbusername.Text)
        End If
    End Sub
End Class