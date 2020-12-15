Imports System.Deployment.Application
Imports System.Reflection

Public Class LogIn

    Dim useravailable, closewindow As String
    Dim ds As New DataSet
    Public vrl As New MainForm

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If TextBox1.Text = "" Or TextBox2.Text = "" Then
            MessageBox.Show("Enter all the details", "Button1_Click()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            getUser(TextBox1.Text)
            If useravailable = "Yes" Then
                If tbemployername.Text = "" Then
                    MessageBox.Show("Select Employer first", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Else
                    Dim cipherText As String = wrapper.EncryptData(TextBox2.Text.Trim)
                    If cipherText = ds.Tables(0).Rows(0).Item(1).ToString Then

                        If ds.Tables(0).Rows(0).Item(2).ToString = "Front Office" Then
                            vrl.usertype = "Front Office"
                        ElseIf ds.Tables(0).Rows(0).Item(2).ToString = "Front Office Receipts And Cheques" Then
                            vrl.usertype = "Front Office Receipts And Cheques"
                        ElseIf ds.Tables(0).Rows(0).Item(2).ToString = "Front Office Supervisor" Then
                            vrl.usertype = "Front Office Supervisor"
                        ElseIf ds.Tables(0).Rows(0).Item(2).ToString = "Back Office Viewer" Then
                            vrl.usertype = "Back Office Viewer"
                        ElseIf ds.Tables(0).Rows(0).Item(2).ToString = "Back Office Editor" Then
                            vrl.usertype = "Back Office Editor"
                        ElseIf ds.Tables(0).Rows(0).Item(2).ToString = "Back Office Supervisor" Then
                            vrl.usertype = "Back Office Supervisor"
                        ElseIf ds.Tables(0).Rows(0).Item(2).ToString = "Administrator" Then
                            vrl.usertype = "Administrator"
                        End If
                        closewindow = "Yes"
                        Me.DialogResult = Windows.Forms.DialogResult.OK
                        wrapper.saveAction("Logged in", Date.Now, ds.Tables(0).Rows(0).Item(3).ToString, "")
                        logedinname = Me.ds.Tables(0).Rows(0).Item(3).ToString
                        supervisorid = Integer.Parse(ds.Tables(0).Rows(0).Item(4).ToString)
                        branchid = Integer.Parse(ds.Tables(0).Rows(0).Item(5).ToString)
                        RMOBranchID = getValueStringOtherSystems("select RMOBranch from CompanyBranches where AutoID = '" & branchid & "'", con)
                        RMBBranchID = getValueStringOtherSystems("select RMBBranch from CompanyBranches where AutoID = '" & branchid & "'", con)
                        MLMABranchID = getValueStringOtherSystems("select MLMABranchID from CompanyBranches where AutoID = '" & branchid & "'", con)
                        MLMEBranchID = getValueStringOtherSystems("select MLMEBranchID from CompanyBranches where AutoID = '" & branchid & "'", con)

                        ''MessageBox.Show(branchid, "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)

                        employername = tbemployername.Text.Trim
                        employerid = Integer.Parse(tbemployerid.Text.Trim)

                        vrl.Text = "Main Form : Employer Name = " & employername
                        vrl.ToolStripMenuItem1.Text = "Employer Name = " & employername
                        vrl.ToolStripStatusLabel1.Text = "Client logged in : " & logedinname
                        vrl.ToolStripStatusLabel2.Text = " Logged in at : " & Date.Now

                        Dim versionNumber As Version
                        versionNumber = Assembly.GetExecutingAssembly().GetName().Version
                        vrl.ToolStripStatusLabel3.Text = versionNumber.ToString
                        If getSYstemVersion() = versionNumber.ToString Then

                        Else
                            MessageBox.Show("The System needs to be updated. Contact Support Team", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            vrl.MenuStrip1.Enabled = False
                            If ds.Tables(0).Rows(0).Item(2).ToString = "Administrator" Then
                                Dim vrl As New CompanyDetails
                                vrl.Show()
                            End If
                        End If
                    Else
                        MessageBox.Show("The password is not correct", "Button1_Click()", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                End If

            ElseIf useravailable = "No" Then
                MessageBox.Show("The username is not correct", "Button1_Click()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
    End Sub

    Private Sub getUser(ByVal str As String)
        ds.Clear()
        Dim qry As String = "select * from Users inner join CompanyBranches on Users.CompanyBranchID = CompanyBranches.AutoID where Username = '" & str & "'"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
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
    End Sub

    Private Sub LogIn_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If closewindow = "Yes" Then

        Else
            e.Cancel = True
        End If

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        closewindow = "Yes"
        End
    End Sub

    Private Sub LogIn_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

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

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Dim vrl As New EmployerDetails
        ''vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If CheckBox1.Checked = True Then
            con = New SqlClient.SqlConnection(ConnString.Default.ConnLocal)
            conFPM = New SqlClient.SqlConnection(ConnString.Default.ConnFPMLocal)
            conMLMAdvances = New SqlClient.SqlConnection(ConnString.Default.ConnMLMLocalAdvances)
            conMLMEducational = New SqlClient.SqlConnection(ConnString.Default.ConnMLMLocalEducational)
            conRMOrange = New SqlClient.SqlConnection(ConnString.Default.ConnRMOrangeLocal)
            conRMBeMobile = New SqlClient.SqlConnection(ConnString.Default.ConnRMBeMobileLocal)
            conServer = New SqlClient.SqlConnection(ConnString.Default.Conn)
            DatabaseLocation = "Local"
        ElseIf CheckBox1.Checked = False Then
            con = New SqlClient.SqlConnection(ConnString.Default.Conn)
            conFPM = New SqlClient.SqlConnection(ConnString.Default.ConnFPM)
            conMLMAdvances = New SqlClient.SqlConnection(ConnString.Default.ConnMLMAdvances)
            conMLMEducational = New SqlClient.SqlConnection(ConnString.Default.ConnMLMEducational)
            conRMOrange = New SqlClient.SqlConnection(ConnString.Default.ConnRMOrange)
            conRMBeMobile = New SqlClient.SqlConnection(ConnString.Default.ConnRMBeMobile)
            conServer = New SqlClient.SqlConnection("")
            DatabaseLocation = "Server"
        End If

        GroupBox1.Enabled = True

        'vrl.MenuStrip1.Items(4).Visible = True
        'vrl.MenuStrip1.Items(4).Enabled = True
        loadEmployers()
        tbemployername.Text = ""
        tbemployerid.Text = ""

        If checkSettlementDetails() = False Then
            saveSettlementLetterDesignDetails("TO: WHOM IT MAY CONCERN", "RE: LOAN ACCOUNT:", "OF ID NO:", "This serves to confirm that", "has an account with us.", "Please be informed that his/her outstanding balance as end of ", "If you intend to clear this account, kindly be requested to credit the following account.", "Thito holdings (Pty) Ltd, First National Bank, Account No: 62031160500 Main Branch Gaborone, or write a cheque payable to Thito Holdings (Pty) LTD.", "Please use your ID number as reference when making payment to enable your account to be credited accordingly.", "For more information do not hesitate to contact the undersigned.", "Yours faithfully", "Name Of Supervisor", "SENIOR CLIENT SERVICE OFFICER", "tshegofatso@thitoholdings.co.bw", "Please be informed that this letter is valid until ", 25, 15.0, "Please ensure the payment is made before", "to avoid additional interest charges.", 7)
        End If

        If checkClearanceDetails() = False Then
            saveClearanceLetterDesignDetails("TO: WHOM IT MAY CONCERN", "RE: ACCOUNT - ", "OF ID NO. ", "This serves to confirm that ", " has fully cleared his/her loan account with us as end of ", "Therefore he/she does not owe us.", "For more information do not hesitate to contact the undersigned.", "Yours Faithfully", "Senior Client Services Officer")
        End If

        If checkEmployers() = False Then
            LinkLabel1.Visible = True
        End If

        If checkUsers() = False Then
            LinkLabel2.Visible = True
        End If

        If checkProducts() = False Then
            saveProduct("Commission", 1, "Active", 1, 0, "Commission", "Yes", "Yes", "Yes", "No")
        End If

        If checkSupervisors() = False Then
            LinkLabel3.Visible = True
        End If

        If checkStatus() = False Then
            saveStatus("Active")
            saveStatus("Cancelled")
        End If

    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Dim vrl As New AddUser
        ''vrl.MdiParent = mform
        vrl.Show()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged

        If CheckBox1.Checked = True Then
            con = New SqlClient.SqlConnection(ConnString.Default.ConnLocal)
            conFPM = New SqlClient.SqlConnection(ConnString.Default.ConnFPMLocal)
            conMLMAdvances = New SqlClient.SqlConnection(ConnString.Default.ConnMLMLocalAdvances)
            conMLMEducational = New SqlClient.SqlConnection(ConnString.Default.ConnMLMLocalEducational)
            conRMOrange = New SqlClient.SqlConnection(ConnString.Default.ConnRMOrangeLocal)
            conRMBeMobile = New SqlClient.SqlConnection(ConnString.Default.ConnRMBeMobileLocal)
            conServer = New SqlClient.SqlConnection(ConnString.Default.Conn)
            DatabaseLocation = "Local"
        ElseIf CheckBox1.Checked = False Then
            con = New SqlClient.SqlConnection(ConnString.Default.Conn)
            conFPM = New SqlClient.SqlConnection(ConnString.Default.ConnFPM)
            conMLMAdvances = New SqlClient.SqlConnection(ConnString.Default.ConnMLMAdvances)
            conMLMEducational = New SqlClient.SqlConnection(ConnString.Default.ConnMLMEducational)
            conRMOrange = New SqlClient.SqlConnection(ConnString.Default.ConnRMOrange)
            conRMBeMobile = New SqlClient.SqlConnection(ConnString.Default.ConnRMBeMobile)
            conServer = New SqlClient.SqlConnection("")
            DatabaseLocation = "Server"
        End If

        GroupBox1.Enabled = True

        'vrl.MenuStrip1.Items(4).Visible = True
        'vrl.MenuStrip1.Items(4).Enabled = True
        loadEmployers()
        tbemployername.Text = ""
        tbemployerid.Text = ""

        If checkSettlementDetails() = False Then
            saveSettlementLetterDesignDetails("TO: WHOM IT MAY CONCERN", "RE: LOAN ACCOUNT:", "OF ID NO:", "This serves to confirm that", "has an account with us.", "Please be informed that his/her outstanding balance as end of ", "If you intend to clear this account, kindly be requested to credit the following account.", "Thito holdings (Pty) Ltd, First National Bank, Account No: 62031160500 Main Branch Gaborone, or write a cheque payable to Thito Holdings (Pty) LTD.", "Please use your ID number as reference when making payment to enable your account to be credited accordingly.", "For more information do not hesitate to contact the undersigned.", "Yours faithfully", "Name Of Supervisor", "SENIOR CLIENT SERVICE OFFICER", "tshegofatso@thitoholdings.co.bw", "Please be informed that this letter is valid until ", 25, 15.0, "Please ensure the payment is made before", "to avoid additional interest charges.", 7)
        End If

        If checkClearanceDetails() = False Then
            saveClearanceLetterDesignDetails("TO: WHOM IT MAY CONCERN", "RE: ACCOUNT - ", "OF ID NO. ", "This serves to confirm that ", " has fully cleared his/her loan account with us as end of ", "Therefore he/she does not owe us.", "For more information do not hesitate to contact the undersigned.", "Yours Faithfully", "Senior Client Services Officer")
        End If

        If checkEmployers() = False Then
            LinkLabel1.Visible = True
        End If

        If checkUsers() = False Then
            LinkLabel2.Visible = True
        End If

        If checkProducts() = False Then
            saveProduct("Commission", 1, "Active", 1, 0, "Commission", "Yes", "Yes", "Yes", "No")
        End If

        If checkSupervisors() = False Then
            LinkLabel3.Visible = True
        End If

        If checkStatus() = False Then
            saveStatus("Active")
            saveStatus("Cancelled")
        End If

    End Sub

    Private Sub LinkLabel3_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel3.LinkClicked
        Dim vrl As New Supervisors
        ''vrl.MdiParent = mform
        vrl.Show()
    End Sub
End Class