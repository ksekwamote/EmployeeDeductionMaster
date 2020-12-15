Public Class SwitchEmployers

    Public mainw As New MainForm

    Private Sub SwitchEmployers_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadEmployers()
        tbemployername.Text = ""
        tbemployerid.Text = ""
    End Sub

    Public Sub loadEmployers()
        Dim qry As String = "select * from Employers"

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
                Me.ComboBox1.DataSource = ds
                Me.ComboBox1.DisplayMember = "Emp.EmployerName"
                Me.ComboBox1.ValueMember = "Emp.AutoID"
                Me.ComboBox1.SelectedIndex = 0
            Else

            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "loadEmployers()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        tbemployername.Text = ComboBox1.Text
        tbemployerid.Text = ComboBox1.SelectedValue.ToString
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If tbemployername.Text = "" Then
            MessageBox.Show("Select Employer first", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            employername = tbemployername.Text.Trim
            employerid = Integer.Parse(tbemployerid.Text.Trim)
            mainw.Text = "Main Form : Employer Name = " & employername
            mainw.ToolStripMenuItem1.Text = "Employer Name = " & employername
            Close()
        End If
    End Sub
End Class