Public Class Customer

    Dim ds As New DataSet
    Dim clearing As Boolean

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim vrl As New CountriesList
        vrl.MdiParent = mform
        vrl.whichform = "Customer"
        vrl.cust = Me
        vrl.Show()
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Dim vrl As New DepartmentList
        vrl.MdiParent = mform
        vrl.whichform = "Customer"
        vrl.cust = Me
        vrl.Show()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim vrl As New IdentityTypeList
        vrl.MdiParent = mform
        vrl.whichform = "Customer"
        vrl.cust = Me
        vrl.Show()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim vrl As New GenderList
        vrl.MdiParent = mform
        vrl.whichform = "Customer"
        vrl.cust = Me
        vrl.Show()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim vrl As New RegionList
        vrl.MdiParent = mform
        vrl.whichform = "Customer"
        vrl.cust = Me
        vrl.Show()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim vrl As New SchoolList
        vrl.MdiParent = mform
        vrl.whichform = "Customer"
        vrl.cust = Me
        vrl.Show()
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Dim vrl As New EmploymentTypeList
        vrl.MdiParent = mform
        vrl.whichform = "Customer"
        vrl.cust = Me
        vrl.Show()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim vrl As New LocationList
        vrl.MdiParent = mform
        vrl.whichform = "Customer"
        vrl.cust = Me
        vrl.Show()
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Dim vrl As New EmployerList
        vrl.MdiParent = mform
        vrl.whichform = "Customer"
        vrl.cust = Me
        vrl.Show()
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Dim vrl As New GroupList
        vrl.MdiParent = mform
        vrl.whichform = "Customer"
        vrl.cust = Me
        vrl.Show()
    End Sub

    Private Sub btsavenew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btsavenew.Click
        If tbtitle.Text.Trim = "" Or tbcountry.Text.Trim = "" Or tbidentitytype.Text.Trim = "" Or tbidnumber.Text.Trim = "" Or tbfirstnames.Text.Trim = "" Or tbsurname.Text.Trim = "" Or tbgender.Text.Trim = "" Or tbcellnumber.Text.Trim = "" Or tbemail.Text.Trim = "" Or tbpostaladdress.Text.Trim = "" Or tbhometelephone.Text.Trim = "" Or tbregion.Text.Trim = "" Or tbtown.Text.Trim = "" Or tbarea.Text.Trim = "" Or tbward.Text.Trim = "" Or tbchief.Text.Trim = "" Or tbhomephysicaladdress.Text.Trim = "" Or tbschool.Text.Trim = "" Or tbemployer.Text.Trim = "" Or tbdepartment.Text.Trim = "" Or tbgroup.Text.Trim = "" Or tboccupation.Text.Trim = "" Or tbworkpostaladdress.Text.Trim = "" Or tbworkphysicaladdress.Text.Trim = "" Or tbworkemailaddress.Text.Trim = "" Or tbemploymenttype.Text.Trim = "" Or tbemployercontactperson.Text.Trim = "" Or tbemptelephone1.Text.Trim = "" Or tbemptelephone2.Text.Trim = "" Or tbempfax1.Text.Trim = "" Or tbempfax2.Text.Trim = "" Or tbnextofkinname1.Text.Trim = "" Or tbnextofkinid1.Text.Trim = "" Or tbnextofkincontact1.Text.Trim = "" Or tbnextofkinaddress1.Text.Trim = "" Or tbnextofkinname2.Text.Trim = "" Or tbnextofkinid2.Text.Trim = "" Or tbnextofkincontact2.Text.Trim = "" Or tbnextofkinaddress2.Text.Trim = "" Or tbnetsalary.Text.Trim = "" Or tbbankname.Text.Trim = "" Or tbbranchname.Text.Trim = "" Or tbaccountnumber.Text.Trim = "" Or tbpayrolnumber.Text.Trim = "" Then
            MessageBox.Show("Enter all the details before saving")
        Else
            If decision("Are you sure to proceed ?") = DialogResult.Yes Then
                saveCustomer(tbtitle.Text.Trim, tbcountryid.Text.Trim, tbidentitytypeid.Text.Trim, dtpindentityexpirydate.Value, tbidnumber.Text.Trim, tbfirstnames.Text.Trim, tbsurname.Text.Trim, dtpdateofbirth.Value, tbgenderid.Text.Trim, tbcellnumber.Text.Trim, tbemail.Text.Trim, tbpostaladdress.Text.Trim, tbhometelephone.Text.Trim, tbregionid.Text.Trim, tbtownid.Text.Trim, tbarea.Text.Trim, tbward.Text.Trim, tbchief.Text.Trim, tbhomephysicaladdress.Text.Trim, tbschoolid.Text.Trim, tbemployerid.Text.Trim, tbdepartmentid.Text.Trim, tboccupation.Text.Trim, tbworkpostaladdress.Text.Trim, tbworkphysicaladdress.Text.Trim, tbworkemailaddress.Text.Trim, tbemploymenttypeid.Text.Trim, tbemployercontactperson.Text.Trim, dtpdateofengagement.Value, tbemptelephone1.Text.Trim, tbemptelephone2.Text.Trim, tbempfax1.Text.Trim, tbempfax2.Text.Trim, tbnextofkinname1.Text.Trim, tbnextofkinid1.Text.Trim, tbnextofkincontact1.Text.Trim, tbnextofkinaddress1.Text.Trim, tbnextofkinname2.Text.Trim, tbnextofkinid2.Text.Trim, tbnextofkincontact2.Text.Trim, tbnextofkinaddress2.Text.Trim, Decimal.Parse(tbnetsalary.Text.Trim), tbbankname.Text.Trim, tbbranchname.Text.Trim, tbaccountnumber.Text.Trim, branchid, Date.Now, "Active", Date.Now, "", Date.Now, tbgroupid.Text.Trim, tbpayrolnumber.Text.Trim)
                MessageBox.Show("Done Saving")
                clearText()
            End If
        End If
    End Sub

    Private Sub Customer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Public Sub clearText()
        clearing = True
        tbtitle.Text = ""
        tbcountry.Text = ""
        tbcountryid.Text = ""
        tbidentitytype.Text = ""
        tbidentitytypeid.Text = ""
        dtpindentityexpirydate.Value = Date.Now
        tbidnumber.Text = ""
        tbidnumber.Enabled = True
        tbfirstnames.Text = ""
        tbsurname.Text = ""
        dtpdateofbirth.Value = Date.Now
        tbgender.Text = ""
        tbgenderid.Text = ""
        tbcellnumber.Text = ""
        tbemail.Text = ""
        tbpostaladdress.Text = ""
        tbhometelephone.Text = ""
        tbregion.Text = ""
        tbregionid.Text = ""
        tbtown.Text = ""
        tbtownid.Text = ""
        tbarea.Text = ""
        tbward.Text = ""
        tbchief.Text = ""
        tbhomephysicaladdress.Text = ""
        tbschool.Text = ""
        tbschoolid.Text = ""
        tbemployer.Text = ""
        tbemployerid.Text = ""
        tbdepartment.Text = ""
        tbdepartmentid.Text = ""
        tboccupation.Text = ""
        tbworkpostaladdress.Text = ""
        tbworkphysicaladdress.Text = ""
        tbworkemailaddress.Text = ""
        tbemploymenttype.Text = ""
        tbemploymenttypeid.Text = ""
        tbemployercontactperson.Text = ""
        dtpdateofengagement.Value = Date.Now
        tbemptelephone1.Text = ""
        tbemptelephone2.Text = ""
        tbempfax1.Text = ""
        tbempfax2.Text = ""
        tbnextofkinname1.Text = ""
        tbnextofkinid1.Text = ""
        tbnextofkincontact1.Text = ""
        tbnextofkinaddress1.Text = ""
        tbnextofkinname2.Text = ""
        tbnextofkinid2.Text = ""
        tbnextofkincontact2.Text = ""
        tbnextofkinaddress2.Text = ""
        tbnetsalary.Text = ""
        tbbankname.Text = ""
        tbbranchname.Text = ""
        tbaccountnumber.Text = ""
        tbgroup.Text = ""
        tbgroupid.Text = ""
        tbpayrolnumber.Text = ""
        tbautoid.Text = ""
        btsavenew.Enabled = True
        btupdatechanges.Enabled = False
        clearing = False
        tbpayrolnumber.Enabled = True
        LinkLabel1.Enabled = True
    End Sub

    Private Sub btcleartexts_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btcleartexts.Click
        clearText()
    End Sub

    Private Sub tbidnumber_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbidnumber.Leave
        getCustomerDetails()
    End Sub

    Public Sub getCustomerDetails()
        If clearing = False Then
            If tbidnumber.Text.Trim = "" Then

            Else
                ds.Clear()
                ds = getCustomerDetailsFormCustomer(tbidnumber.Text.Trim)
                If ds.Tables(0).Rows.Count > 0 Then
                    tbautoid.Text = ds.Tables(0).Rows(0).Item(0).ToString
                    tbtitle.Text = ds.Tables(0).Rows(0).Item(1).ToString
                    tbcountry.Text = getValueString("select CountryName from Country where CountryID = '" & ds.Tables(0).Rows(0).Item(2).ToString & "'")
                    tbcountryid.Text = ds.Tables(0).Rows(0).Item(2).ToString
                    tbidentitytype.Text = getValueString("select IdentityTypeName from IdentityType where IdentityTypeID = '" & ds.Tables(0).Rows(0).Item(3).ToString & "'")
                    tbidentitytypeid.Text = ds.Tables(0).Rows(0).Item(3).ToString
                    dtpindentityexpirydate.Value = Date.Parse(ds.Tables(0).Rows(0).Item(4).ToString)
                    tbidnumber.Text = ds.Tables(0).Rows(0).Item(5).ToString
                    tbidnumber.Enabled = False
                    tbfirstnames.Text = ds.Tables(0).Rows(0).Item(6).ToString
                    tbsurname.Text = ds.Tables(0).Rows(0).Item(7).ToString
                    dtpdateofbirth.Value = Date.Parse(ds.Tables(0).Rows(0).Item(8).ToString)
                    tbgender.Text = getValueString("select GenderName from Gender where GenderID = '" & ds.Tables(0).Rows(0).Item(9).ToString & "'")
                    tbgenderid.Text = ds.Tables(0).Rows(0).Item(9).ToString
                    tbcellnumber.Text = ds.Tables(0).Rows(0).Item(10).ToString
                    tbemail.Text = ds.Tables(0).Rows(0).Item(11).ToString
                    tbpostaladdress.Text = ds.Tables(0).Rows(0).Item(12).ToString
                    tbhometelephone.Text = ds.Tables(0).Rows(0).Item(13).ToString
                    tbregion.Text = getValueString("select RegionName from Region where RegionID = '" & ds.Tables(0).Rows(0).Item(14).ToString & "'")
                    tbregionid.Text = ds.Tables(0).Rows(0).Item(14).ToString
                    tbtown.Text = getValueString("select LocationName from Location where LocationID = '" & ds.Tables(0).Rows(0).Item(15).ToString & "'")
                    tbtownid.Text = ds.Tables(0).Rows(0).Item(15).ToString
                    tbarea.Text = ds.Tables(0).Rows(0).Item(16).ToString
                    tbward.Text = ds.Tables(0).Rows(0).Item(17).ToString
                    tbchief.Text = ds.Tables(0).Rows(0).Item(18).ToString
                    tbhomephysicaladdress.Text = ds.Tables(0).Rows(0).Item(19).ToString
                    tbschool.Text = getValueString("select SchoolName from School where SchoolID = '" & ds.Tables(0).Rows(0).Item(20).ToString & "'")
                    tbschoolid.Text = ds.Tables(0).Rows(0).Item(20).ToString
                    tbemployer.Text = getValueString("select EmployerName from Employers where AutoID = '" & Integer.Parse(ds.Tables(0).Rows(0).Item(21).ToString) & "'")
                    tbemployerid.Text = ds.Tables(0).Rows(0).Item(21).ToString
                    tbdepartment.Text = getValueString("select DepartmentName from Department where DepartmentID = '" & ds.Tables(0).Rows(0).Item(22).ToString & "'")
                    tbdepartmentid.Text = ds.Tables(0).Rows(0).Item(22).ToString
                    tboccupation.Text = ds.Tables(0).Rows(0).Item(23).ToString
                    tbworkpostaladdress.Text = ds.Tables(0).Rows(0).Item(24).ToString
                    tbworkphysicaladdress.Text = ds.Tables(0).Rows(0).Item(25).ToString
                    tbworkemailaddress.Text = ds.Tables(0).Rows(0).Item(26).ToString
                    tbemploymenttype.Text = getValueString("select EmploymentTypeName from EmploymentType where EmploymentTypeID = '" & ds.Tables(0).Rows(0).Item(27).ToString & "'")
                    tbemploymenttypeid.Text = ds.Tables(0).Rows(0).Item(27).ToString
                    tbemployercontactperson.Text = ds.Tables(0).Rows(0).Item(28).ToString
                    dtpdateofengagement.Value = Date.Parse(ds.Tables(0).Rows(0).Item(29).ToString)
                    tbemptelephone1.Text = ds.Tables(0).Rows(0).Item(30).ToString
                    tbemptelephone2.Text = ds.Tables(0).Rows(0).Item(31).ToString
                    tbempfax1.Text = ds.Tables(0).Rows(0).Item(32).ToString
                    tbempfax2.Text = ds.Tables(0).Rows(0).Item(33).ToString
                    tbnextofkinname1.Text = ds.Tables(0).Rows(0).Item(34).ToString
                    tbnextofkinid1.Text = ds.Tables(0).Rows(0).Item(35).ToString
                    tbnextofkincontact1.Text = ds.Tables(0).Rows(0).Item(36).ToString
                    tbnextofkinaddress1.Text = ds.Tables(0).Rows(0).Item(37).ToString
                    tbnextofkinname2.Text = ds.Tables(0).Rows(0).Item(38).ToString
                    tbnextofkinid2.Text = ds.Tables(0).Rows(0).Item(39).ToString
                    tbnextofkincontact2.Text = ds.Tables(0).Rows(0).Item(40).ToString
                    tbnextofkinaddress2.Text = ds.Tables(0).Rows(0).Item(41).ToString
                    tbnetsalary.Text = ds.Tables(0).Rows(0).Item(42).ToString
                    tbbankname.Text = ds.Tables(0).Rows(0).Item(43).ToString
                    tbbranchname.Text = ds.Tables(0).Rows(0).Item(44).ToString
                    tbaccountnumber.Text = ds.Tables(0).Rows(0).Item(45).ToString
                    tbgroup.Text = getValueString("select GroupName from Groups where GroupID = '" & ds.Tables(0).Rows(0).Item(52).ToString & "'")
                    tbgroupid.Text = ds.Tables(0).Rows(0).Item(52).ToString
                    tbpayrolnumber.Text = ds.Tables(0).Rows(0).Item(53).ToString
                    btsavenew.Enabled = False
                    btupdatechanges.Enabled = True
                    tbpayrolnumber.Enabled = False
                    LinkLabel1.Enabled = False
                    MessageBox.Show("Customer available")
                End If
            End If
        End If
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Dim vrl As New BankList
        vrl.MdiParent = mform
        vrl.whichform = "Customer"
        vrl.cust = Me
        vrl.Show()
        tbbranchname.Text = ""
        tbbranchid.Text = ""
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        Dim vrl As New BankBranchList
        vrl.MdiParent = mform
        vrl.whichform = "Customer"
        vrl.autoid = tbbankid.Text.Trim
        vrl.cust = Me
        vrl.Show()
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        tbpayrolnumber.Text = tbidnumber.Text.Trim
    End Sub
End Class