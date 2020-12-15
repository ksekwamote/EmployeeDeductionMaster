Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports DgvFilterPopup

Public Class CompanyDetails

    Private Sub loadCompanyDetails()
        DataGridView1.DataSource = Nothing
        Dim qry As String = "select * from CompanyDetails"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)
        Dim dt As DataTable

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim ctr As Integer = 0
            SQLAdapter.Fill(ds, "LoanDetails")
            If ds.Tables(0).Rows.Count > 0 Then
                dt = ds.Tables(0)
                Dim bs As New BindingSource
                bs.DataSource = dt.DefaultView
                DataGridView1.DataSource = bs
                BindingNavigator1.BindingSource = bs
                DataGridView1.Columns(DataGridView1.ColumnCount - 1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

                'filterManager = New DgvFilterManager(DataGridView1)
            Else
                DataGridView1.DataSource = Nothing
                BindingNavigator1.BindingSource = Nothing
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "loadBranches()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadBranches()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub Branches_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadCompanyDetails()
    End Sub

    Public Sub saveCompanyDetails(ByVal LenderName As String, ByVal LicenseNumber As String, ByVal AuthorizedAdress As String, ByVal ContactNumber As String, ByVal EmailAdress As String, ByVal FaxNumber As String, ByVal ShortName As String, ByVal SystemName As String, ByVal AffordabilityPercentage As Decimal, ByVal PhysicalAddress As String, ByVal ChequeFee As Decimal, ByVal PasswordUpdateInMonths As Integer, ByVal SystemVersion As String, ByVal CutOfDate As Integer, ByVal ProductIDForPaidCustomers As Integer, ByVal UpdateWithin As Integer, ByVal Integrate As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            transactioninsert = con.BeginTransaction()
            Dim qryinsert As String = "Insert into CompanyDetails values ('" & LenderName & "','" & LicenseNumber & "','" & AuthorizedAdress & "','" & ContactNumber & "','" & EmailAdress & "','" & FaxNumber & "','" & ShortName & "','" & SystemName & "','" & AffordabilityPercentage & "','" & PhysicalAddress & "','" & ChequeFee & "','" & PasswordUpdateInMonths & "','" & SystemVersion & "','" & CutOfDate & "','" & ProductIDForPaidCustomers & "','" & UpdateWithin & "','" & Integrate & "')"
            Dim cmdinsert As New SqlClient.SqlCommand(qryinsert, con)
            cmdinsert.Transaction = transactioninsert
            If cmdinsert.ExecuteNonQuery() Then
                transactioninsert.Commit()
            End If
            MessageBox.Show("Details saved", "saveBranch()", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "saveCompanyDetails()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "saveCompanyDetails()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If tblicensenumber.Text = "" Or tbpostaladdress.Text = "" Or tbcontactnumber.Text = "" Or tbemail.Text = "" Or tbfaxnumber.Text = "" Or tbshortname.Text = "" Or tbsystemname.Text = "" Or tbaffordability.Text = "" Or tbphysicaladdress.Text = "" Or tbchequefee.Text = "" Or tbpasswordreset.Text = "" Or tbproductid.Text.Trim = "" Or tbupdatewithin.Text.Trim = "" Or cbintegrate.Text.Trim = "" Then
            MessageBox.Show("Enter all details", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            saveCompanyDetails(tbcompanyname.Text.Trim, tblicensenumber.Text.Trim, tbpostaladdress.Text.Trim, tbcontactnumber.Text.Trim, tbemail.Text.Trim, tbfaxnumber.Text.Trim, tbshortname.Text.Trim, tbsystemname.Text.Trim, Decimal.Parse(tbaffordability.Text.Trim), tbphysicaladdress.Text.Trim, Decimal.Parse(tbchequefee.Text.Trim), Integer.Parse(tbpasswordreset.Text.Trim), tbsystemversion.Text.Trim, Integer.Parse(tbcutofdate.Text.Trim), Integer.Parse(tbproductid.Text.Trim), Integer.Parse(tbupdatewithin.Text.Trim), cbintegrate.Text.Trim)
            loadCompanyDetails()
            clearTexts
        End If
    End Sub

    Private Sub DataGridView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Click
        If DataGridView1.Rows.Count > 0 Then
            tbautoid.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(0).Value.ToString()
            tbcompanyname.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(1).Value.ToString()
            tblicensenumber.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(2).Value.ToString()
            tbpostaladdress.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(3).Value.ToString()
            tbcontactnumber.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(4).Value.ToString()
            tbemail.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(5).Value.ToString()
            tbfaxnumber.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(6).Value.ToString()
            tbshortname.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(7).Value.ToString()
            tbsystemname.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(8).Value.ToString()
            tbaffordability.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(9).Value.ToString()
            tbphysicaladdress.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(10).Value.ToString()
            tbchequefee.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(11).Value.ToString()
            tbpasswordreset.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(12).Value.ToString()
            tbsystemversion.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(13).Value.ToString()
            tbcutofdate.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(14).Value.ToString()
            tbproductid.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(15).Value.ToString()
            tbproductname.Text = getProductName(Integer.Parse(tbproductid.Text.Trim))
            tbupdatewithin.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(16).Value.ToString()
            cbintegrate.SelectedIndex = cbintegrate.FindStringExact(DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(17).Value.ToString())
            Button2.Enabled = True
            Button1.Enabled = False
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If tbcompanyname.Text = "" Then
            MessageBox.Show("Make Sure to select the branch to update changes", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            If tbautoid.Text = "" Or tblicensenumber.Text = "" Or tbpostaladdress.Text = "" Or tbcontactnumber.Text = "" Or tbemail.Text = "" Or tbfaxnumber.Text = "" Or tbshortname.Text = "" Or tbsystemname.Text = "" Or tbaffordability.Text = "" Or tbphysicaladdress.Text = "" Or tbchequefee.Text = "" Or tbpasswordreset.Text = "" Or tbcutofdate.Text.Trim = "" Or tbupdatewithin.Text.Trim = "" Or cbintegrate.Text.Trim = "" Then
                MessageBox.Show("Enter all details", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                updateCompanyDetails(tbcompanyname.Text.Trim, tblicensenumber.Text.Trim, tbpostaladdress.Text.Trim, tbcontactnumber.Text.Trim, tbemail.Text.Trim, tbfaxnumber.Text.Trim, tbshortname.Text.Trim, tbsystemname.Text.Trim, Decimal.Parse(tbaffordability.Text.Trim), tbphysicaladdress.Text.Trim, Decimal.Parse(tbchequefee.Text.Trim), Integer.Parse(tbpasswordreset.Text.Trim), tbsystemversion.Text.Trim, Integer.Parse(tbcutofdate.Text.Trim), Integer.Parse(tbproductid.Text.Trim), Integer.Parse(tbupdatewithin.Text.Trim), cbintegrate.Text.Trim, Integer.Parse(tbautoid.Text.Trim))
                loadCompanyDetails()
                clearTexts()
            End If
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        clearTexts()
    End Sub

    Public Sub clearTexts()
        Button2.Enabled = False
        Button1.Enabled = True
        tbautoid.Text = ""
        tbcompanyname.Text = ""
        tblicensenumber.Text = ""
        tbpostaladdress.Text = ""
        tbcontactnumber.Text = ""
        tbemail.Text = ""
        tbfaxnumber.Text = ""
        tbshortname.Text = ""
        tbsystemname.Text = ""
        tbaffordability.Text = ""
        tbphysicaladdress.Text = ""
        tbchequefee.Text = ""
        tbpasswordreset.Text = ""
        tbsystemversion.Text = ""
        tbcutofdate.Text = ""
        tbproductid.Text = ""
        tbproductname.Text = ""
        tbupdatewithin.Text = ""
        cbintegrate.SelectedIndex = -1
    End Sub
    Public Sub updateCompanyDetails(ByVal LenderName As String, ByVal LicenseNumber As String, ByVal AuthorizedAdress As String, ByVal ContactNumber As String, ByVal EmailAdress As String, ByVal FaxNumber As String, ByVal ShortName As String, ByVal SystemName As String, ByVal AffordabilityPercentage As Decimal, ByVal PhysicalAddress As String, ByVal ChequeFee As Decimal, ByVal PasswordUpdateInMonths As Integer, ByVal SystemVersion As String, ByVal CutOfDate As Integer, ByVal ProductIDForPaidCustomers As Integer, ByVal UpdateWithin As Integer, ByVal Integrate As String, ByVal AutoID As Integer)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update CompanyDetails set LenderName = '" & LenderName & "',LicenseNumber = '" & LicenseNumber & "',AuthorizedAdress = '" & AuthorizedAdress & "',ContactNumber = '" & ContactNumber & "',EmailAdress = '" & EmailAdress & "',FaxNumber = '" & FaxNumber & "',ShortName = '" & ShortName & "',SystemName = '" & SystemName & "',AffordabilityPercentage = '" & AffordabilityPercentage & "',PhysicalAddress = '" & PhysicalAddress & "',ChequeFee = '" & ChequeFee & "',PasswordUpdateInMonths = '" & PasswordUpdateInMonths & "',SystemVersion = '" & SystemVersion & "',CutOfDate = '" & CutOfDate & "',ProductIDForPaidCustomers = '" & ProductIDForPaidCustomers & "',UpdateWithin = '" & UpdateWithin & "',Integrate = '" & Integrate & "' where AutoID = '" & AutoID & "'"
            Dim cmdinsert As New SqlClient.SqlCommand(qryinsert, con)
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            transactioninsert = con.BeginTransaction()
            cmdinsert.Transaction = transactioninsert

            If cmdinsert.ExecuteNonQuery() Then
                transactioninsert.Commit()
            End If
            MessageBox.Show("Details updated", "updateBranch()", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "updateCompanyDetails()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateCompanyDetails()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try

    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        tbproductid.Text = "0"
        tbproductname.Text = "Not Available"
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim vrl As New ProductsList
        vrl.MdiParent = mform
        vrl.whichform = "CompanyDetails"
        vrl.cd = Me
        vrl.Show()
    End Sub

    Private Sub GroupBox1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox1.Enter

    End Sub
End Class