Public Class Customer

    Dim ds As New DataSet
    Dim clearing As Boolean

    Dim EDMGender, MLMEGender, MLMAGender, RMOGender, RMBGender, FPMGender As String
    Public RMORegionID, RMBRegionID, FPMRegionID As String
    Public RMOLocation, RMBLocation, FPMLocation, MLMALocation, MLMELocation As String
    Public FPMSchool As String
    Public EDMEmployer, RMOEmployer, RMBEmployer, MLMAEmployer, MLMEEmployer, MLMAEmployerName, MLMEEmployerName As String
    Public RMODepartment, RMBDepartment, MLMADepartment, MLMEDepartment, FPMDepartment As String
    Public MLMAEmpType, MLMEEmpType As String
    Public FPMGroup, MLMAGroup, MLMEGroup As String

    Public MLMABankName, MLMABankBranchName, MLMEBankName, MLMEBankBranchName As String
    Dim PreEmployerName, PreGroupName As String

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim vrl As New CountriesList
        vrl.MdiParent = mform
        vrl.whichform = "Customer"
        vrl.cust = Me
        vrl.Show()
    End Sub

    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Me.TabControl1.SelectTab(1)
    End Sub

    Private Sub LinkLabel9_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel9.LinkClicked
        Me.TabControl1.SelectTab(0)
    End Sub

    Private Sub LinkLabel3_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel3.LinkClicked
        Me.TabControl1.SelectTab(2)
    End Sub

    Private Sub LinkLabel8_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel8.LinkClicked
        Me.TabControl1.SelectTab(1)
    End Sub

    Private Sub LinkLabel4_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel4.LinkClicked
        Me.TabControl1.SelectTab(3)
    End Sub

    Private Sub LinkLabel7_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel7.LinkClicked
        Me.TabControl1.SelectTab(2)
    End Sub

    Private Sub LinkLabel5_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel5.LinkClicked
        Me.TabControl1.SelectTab(4)
    End Sub

    Private Sub LinkLabel6_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel6.LinkClicked
        Me.TabControl1.SelectTab(3)
    End Sub

    Private Sub tbbankname_TextChanged(sender As Object, e As EventArgs) Handles tbbankname.TextChanged
        MLMABankName = tbbankname.Text.Trim
        MLMEBankName = tbbankname.Text.Trim
    End Sub

    Private Sub tbbranchname_TextChanged(sender As Object, e As EventArgs) Handles tbbranchname.TextChanged
        MLMABankBranchName = tbbranchname.Text.Trim
        MLMEBankBranchName = tbbranchname.Text.Trim
    End Sub

    Private Sub tbidnumber_TextChanged(sender As Object, e As EventArgs) Handles tbidnumber.TextChanged

    End Sub

    Private Sub LinkLabel10_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel10.LinkClicked
        Me.TabControl1.SelectTab(5)
    End Sub

    Private Sub LinkLabel11_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel11.LinkClicked
        Me.TabControl1.SelectTab(4)
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Dim vrl As New DepartmentList
        vrl.MdiParent = mform
        vrl.whichform = "Customer"
        vrl.cust = Me
        vrl.Show()
    End Sub

    Private Sub tbemployerid_TextChanged(sender As Object, e As EventArgs) Handles tbemployerid.TextChanged
        EDMEmployer = tbemployerid.Text.Trim
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
        setValuesForNewEntryThroughUpdate()
        If checkFieldsEntered() = True Then
            If decision("Are you sure to proceed ?") = DialogResult.Yes Then
                saveCustomer(tbtitle.Text.Trim, tbcountryid.Text.Trim, tbidentitytypeid.Text.Trim, dtpindentityexpirydate.Value, tbidnumber.Text.Trim, tbfirstnames.Text.Trim, tbsurname.Text.Trim, dtpdateofbirth.Value, tbgenderid.Text.Trim, tbcellnumber.Text.Trim, tbemail.Text.Trim, tbpostaladdress.Text.Trim, tbhometelephone.Text.Trim, tbregionid.Text.Trim, tbtownid.Text.Trim, tbarea.Text.Trim, tbward.Text.Trim, tbchief.Text.Trim, tbhomephysicaladdress.Text.Trim, tbschoolid.Text.Trim, tbemployerid.Text.Trim, tbdepartmentid.Text.Trim, tboccupation.Text.Trim, tbworkpostaladdress.Text.Trim, tbworkphysicaladdress.Text.Trim, tbworkemailaddress.Text.Trim, tbemploymenttypeid.Text.Trim, tbemployercontactperson.Text.Trim, dtpdateofengagement.Value, tbemptelephone1.Text.Trim, tbemptelephone2.Text.Trim, tbempfax1.Text.Trim, tbempfax2.Text.Trim, tbnextofkinname1.Text.Trim, tbnextofkinid1.Text.Trim, tbnextofkincontact1.Text.Trim, tbnextofkinaddress1.Text.Trim, tbnextofkinname2.Text.Trim, tbnextofkinid2.Text.Trim, tbnextofkincontact2.Text.Trim, tbnextofkinaddress2.Text.Trim, Decimal.Parse(tbnetsalary.Text.Trim), tbbankid.Text.Trim, tbbranchid.Text.Trim, tbaccountnumber.Text.Trim, branchid, Date.Now, "Active", Date.Now, "", Date.Now, tbgroupid.Text.Trim, tbpayrolnumber.Text.Trim)
                If getValueStringOtherSystems("select Integrate from CompanyDetails", con) = "Yes" Then
                    UpdateOtherSystems()
                End If
                MessageBox.Show("Done Saving")
                Close()
            End If
        End If
    End Sub


    Private Sub Customer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ToolTip1.SetToolTip(tbarea, "The area in the village or town. Not the ward")
        ToolTip1.SetToolTip(tboccupation, "/ Work Position")
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

        EDMGender = ""
        MLMEGender = ""
        MLMAGender = ""
        RMOGender = ""
        RMBGender = ""
        FPMGender = ""

        RMORegionID = ""
        RMBRegionID = ""
        FPMRegionID = ""

        RMOLocation = ""
        RMBLocation = ""
        FPMLocation = ""
        MLMALocation = ""
        MLMELocation = ""

        FPMSchool = ""

        EDMEmployer = ""
        RMOEmployer = ""
        RMBEmployer = ""
        MLMAEmployer = ""
        MLMEEmployer = ""
        MLMAEmployerName = ""
        MLMEEmployerName = ""

        RMODepartment = ""
        RMBDepartment = ""
        MLMADepartment = ""
        MLMEDepartment = ""
        FPMDepartment = ""

        MLMAEmpType = ""
        MLMEEmpType = ""

        FPMGroup = ""
        MLMAGroup = ""
        MLMEGroup = ""

        MLMABankName = ""
        MLMABankBranchName = ""
        MLMEBankName = ""
        MLMEBankBranchName = ""

    End Sub

    Private Sub btcleartexts_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btcleartexts.Click
        clearText()
    End Sub

    Private Sub tbidnumber_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbidnumber.Leave
        getCustomerDetails()
    End Sub

    Private Sub tbgender_TextChanged(sender As Object, e As EventArgs) Handles tbgender.TextChanged
        EDMGender = tbgender.Text.Trim
        MLMEGender = tbgender.Text.Trim
        MLMAGender = tbgender.Text.Trim
        RMOGender = tbgender.Text.Trim
        RMBGender = tbgender.Text.Trim
        FPMGender = tbgender.Text.Trim
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
                    PreEmployerName = tbemployer.Text.Trim
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
                    tbbankname.Text = getValueString("select BankNAme from Bank where BankID = '" & ds.Tables(0).Rows(0).Item(43).ToString & "'")
                    tbbankid.Text = ds.Tables(0).Rows(0).Item(43).ToString
                    tbbranchname.Text = getValueString("select BankBranchName from BankBranch where BankBranchID = '" & ds.Tables(0).Rows(0).Item(44).ToString & "'")
                    tbbranchid.Text = ds.Tables(0).Rows(0).Item(44).ToString
                    tbbranchcode.Text = getValueString("select BranchCode from BankBranch where BankBranchID = '" & ds.Tables(0).Rows(0).Item(44).ToString & "'")
                    tbaccountnumber.Text = ds.Tables(0).Rows(0).Item(45).ToString
                    tbgroup.Text = getValueString("select GroupName from Groups where GroupID = '" & ds.Tables(0).Rows(0).Item(52).ToString & "'")
                    PreGroupName = tbgroup.Text.Trim
                    tbgroupid.Text = ds.Tables(0).Rows(0).Item(52).ToString
                    tbpayrolnumber.Text = ds.Tables(0).Rows(0).Item(53).ToString
                    btsavenew.Enabled = False
                    btupdatechanges.Enabled = True
                    tbpayrolnumber.Enabled = False
                    LinkLabel1.Enabled = False

                    EDMGender = getValueStringOtherSystems("select Gender from CustomerDetails where CustomerID = '" & tbidnumber.Text.Trim & "'", con)
                    MLMEGender = getValueStringOtherSystems("select Gender from CustomerDetails where ID_Number = '" & tbidnumber.Text.Trim & "'", conMLMEducational)
                    MLMAGender = getValueStringOtherSystems("select Gender from CustomerDetails where ID_Number = '" & tbidnumber.Text.Trim & "'", conMLMAdvances)
                    RMOGender = getValueStringOtherSystems("select Gender from CustomerMaster where Account = '" & tbidnumber.Text.Trim & "'", conRMOrange)
                    RMBGender = getValueStringOtherSystems("select Gender from CustomerMaster where Account = '" & tbidnumber.Text.Trim & "'", conRMBeMobile)
                    FPMGender = getValueStringOtherSystems("select Gender from CustomerDetails where IDNumber = '" & tbidnumber.Text.Trim & "'", conFPM)

                    RMORegionID = getValueStringOtherSystems("select Region from CustomerMaster where Account = '" & tbidnumber.Text.Trim & "'", conRMOrange)
                    RMBRegionID = getValueStringOtherSystems("select Region from CustomerMaster where Account = '" & tbidnumber.Text.Trim & "'", conRMBeMobile)
                    FPMRegionID = getValueStringOtherSystems("select RegionID from CustomerDetails where IDNumber = '" & tbidnumber.Text.Trim & "'", conFPM)

                    RMOLocation = getValueStringOtherSystems("select Town_Village from CustomerMaster where Account = '" & tbidnumber.Text.Trim & "'", conRMOrange)
                    RMBLocation = getValueStringOtherSystems("select Town_Village from CustomerMaster where Account = '" & tbidnumber.Text.Trim & "'", conRMBeMobile)
                    FPMLocation = getValueStringOtherSystems("select AreaCode from CustomerDetails where IDNumber = '" & tbidnumber.Text.Trim & "'", conFPM)
                    MLMALocation = getValueStringOtherSystems("select Home_Village from CustomerDetails where ID_Number = '" & tbidnumber.Text.Trim & "'", conMLMAdvances)
                    MLMELocation = getValueStringOtherSystems("select Home_Village from CustomerDetails where ID_Number = '" & tbidnumber.Text.Trim & "'", conMLMEducational)

                    FPMSchool = getValueStringOtherSystems("select SchoolID from CustomerDetails where IDNumber = '" & tbidnumber.Text.Trim & "'", conFPM)

                    EDMEmployer = getValueStringOtherSystems("select EmployerID from CustomerDetails where CustomerID = '" & tbidnumber.Text.Trim & "'", con)
                    RMOEmployer = getValueStringOtherSystems("select EmployerCode from CustomerMaster where Account = '" & tbidnumber.Text.Trim & "'", conRMOrange)
                    RMBEmployer = getValueStringOtherSystems("select EmployerCode from CustomerMaster where Account = '" & tbidnumber.Text.Trim & "'", conRMBeMobile)
                    MLMAEmployer = getValueStringOtherSystems("select EmployerID from CustomerDetails where ID_Number = '" & tbidnumber.Text.Trim & "'", conMLMAdvances)
                    MLMEEmployer = getValueStringOtherSystems("select EmployerID from CustomerDetails where ID_Number = '" & tbidnumber.Text.Trim & "'", conMLMEducational)
                    MLMAEmployerName = getValueStringOtherSystems("select Name_Of_Employer from CustomerDetails where ID_Number = '" & tbidnumber.Text.Trim & "'", conMLMAdvances)
                    MLMEEmployerName = getValueStringOtherSystems("select Name_Of_Employer from CustomerDetails where ID_Number = '" & tbidnumber.Text.Trim & "'", conMLMEducational)

                    RMODepartment = getValueStringOtherSystems("select EmployerDeptCode from CustomerMaster where Account = '" & tbidnumber.Text.Trim & "'", conRMOrange)
                    RMBDepartment = getValueStringOtherSystems("select EmployerDeptCode from CustomerMaster where Account = '" & tbidnumber.Text.Trim & "'", conRMBeMobile)
                    MLMADepartment = getValueStringOtherSystems("select DepartmentID from CustomerDetails where ID_Number = '" & tbidnumber.Text.Trim & "'", conMLMAdvances)
                    MLMEDepartment = getValueStringOtherSystems("select DepartmentID from CustomerDetails where ID_Number = '" & tbidnumber.Text.Trim & "'", conMLMEducational)
                    FPMDepartment = getValueStringOtherSystems("select DepID from CustomerDetails where IDNumber = '" & tbidnumber.Text.Trim & "'", conFPM)

                    MLMAEmpType = getValueStringOtherSystems("select EmployerType from CustomerDetails where ID_Number = '" & tbidnumber.Text.Trim & "'", conMLMAdvances)
                    MLMEEmpType = getValueStringOtherSystems("select EmployerType from CustomerDetails where ID_Number = '" & tbidnumber.Text.Trim & "'", conMLMEducational)

                    FPMGroup = getValueStringOtherSystems("select GroupID from CustomerDetails where IDNumber = '" & tbidnumber.Text.Trim & "'", conFPM)
                    MLMAGroup = getValueStringOtherSystems("select GroupID from CustomerDetails where ID_Number = '" & tbidnumber.Text.Trim & "'", conMLMAdvances)
                    MLMEGroup = getValueStringOtherSystems("select GroupID from CustomerDetails where ID_Number = '" & tbidnumber.Text.Trim & "'", conMLMEducational)

                    MLMABankName = getValueStringOtherSystems("select BankName from CustomerDetails where ID_Number = '" & tbidnumber.Text.Trim & "'", conMLMAdvances)
                    MLMABankBranchName = getValueStringOtherSystems("select BranchName from CustomerDetails where ID_Number = '" & tbidnumber.Text.Trim & "'", conMLMAdvances)
                    MLMEBankName = getValueStringOtherSystems("select BankName from CustomerDetails where ID_Number = '" & tbidnumber.Text.Trim & "'", conMLMEducational)
                    MLMEBankBranchName = getValueStringOtherSystems("select BranchName from CustomerDetails where ID_Number = '" & tbidnumber.Text.Trim & "'", conMLMEducational)

                    checkCustomerInOtherSYstems()

                    MessageBox.Show("Customer available")
                End If
            End If
        End If
    End Sub

    Public Sub checkCustomerInOtherSYstems()
        If getValueBooleanOtherSystems("select * from CustomerDetails where CustomerID = '" & tbidnumber.Text.Trim & "'", con) = True Then
            cbedm.Checked = True
            cbedm.Enabled = False
        Else
            cbedm.Enabled = True
        End If

        If getValueBooleanOtherSystems("select * from CustomerDetails where IDNumber = '" & tbidnumber.Text.Trim & "'", conFPM) = True Then
            cbfpm.Checked = True
            cbfpm.Enabled = False
        Else
            cbfpm.Enabled = True
        End If

        If getValueBooleanOtherSystems("select * from CustomerDetails where ID_Number = '" & tbidnumber.Text.Trim & "'", conMLMAdvances) = True Then
            cbmlma.Checked = True
            cbmlma.Enabled = False
        Else
            cbmlma.Enabled = True
        End If

        If getValueBooleanOtherSystems("select * from CustomerDetails where ID_Number = '" & tbidnumber.Text.Trim & "'", conMLMEducational) = True Then
            cbmlme.Checked = True
            cbmlme.Enabled = False
        Else
            cbmlme.Enabled = True
        End If

        If getValueBooleanOtherSystems("select * from CustomerMaster where Account = '" & tbidnumber.Text.Trim & "'", conRMOrange) = True Then
            cbrmo.Checked = True
            cbrmo.Enabled = False
        Else
            cbrmo.Enabled = True
        End If

        If getValueBooleanOtherSystems("select * from CustomerMaster where Account = '" & tbidnumber.Text.Trim & "'", conRMBeMobile) = True Then
            cbrmb.Checked = True
            cbrmb.Enabled = False
        Else
            cbrmb.Enabled = True
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

    Private Sub btupdatechanges_Click(sender As Object, e As EventArgs) Handles btupdatechanges.Click

        If tbautoid.Text.Trim = "" Then
            MessageBox.Show("Nothing selected to be updated", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            setValuesForNewEntryThroughUpdate()
            If checkFieldsEntered() = True Then
                If decision("Are you sure to proceed ?") = DialogResult.Yes Then
                    updateCustomer(tbtitle.Text.Trim, tbcountryid.Text.Trim, tbidentitytypeid.Text.Trim, dtpindentityexpirydate.Value, tbidnumber.Text.Trim, tbfirstnames.Text.Trim, tbsurname.Text.Trim, dtpdateofbirth.Value, tbgenderid.Text.Trim, tbcellnumber.Text.Trim, tbemail.Text.Trim, tbpostaladdress.Text.Trim, tbhometelephone.Text.Trim, tbregionid.Text.Trim, tbtownid.Text.Trim, tbarea.Text.Trim, tbward.Text.Trim, tbchief.Text.Trim, tbhomephysicaladdress.Text.Trim, tbschoolid.Text.Trim, tbemployerid.Text.Trim, tbdepartmentid.Text.Trim, tboccupation.Text.Trim, tbworkpostaladdress.Text.Trim, tbworkphysicaladdress.Text.Trim, tbworkemailaddress.Text.Trim, tbemploymenttypeid.Text.Trim, tbemployercontactperson.Text.Trim, dtpdateofengagement.Value, tbemptelephone1.Text.Trim, tbemptelephone2.Text.Trim, tbempfax1.Text.Trim, tbempfax2.Text.Trim, tbnextofkinname1.Text.Trim, tbnextofkinid1.Text.Trim, tbnextofkincontact1.Text.Trim, tbnextofkinaddress1.Text.Trim, tbnextofkinname2.Text.Trim, tbnextofkinid2.Text.Trim, tbnextofkincontact2.Text.Trim, tbnextofkinaddress2.Text.Trim, Decimal.Parse(tbnetsalary.Text.Trim), tbbankid.Text.Trim, tbbranchid.Text.Trim, tbaccountnumber.Text.Trim, Date.Now, tbgroupid.Text.Trim, tbpayrolnumber.Text.Trim, Integer.Parse(tbautoid.Text.Trim))
                    If PreEmployerName = tbemployer.Text.Trim Then

                    Else
                        wrapper.saveAction("Employer Changed from " & PreEmployerName & " to " & tbemployer.Text.Trim, Date.Now, logedinname, tbidnumber.Text.Trim)
                    End If

                    If PreGroupName = tbgroup.Text.Trim Then

                    Else
                        wrapper.saveAction("Group Changed from " & PreGroupName & " to " & tbgroup.Text.Trim, Date.Now, logedinname, tbidnumber.Text.Trim)
                    End If

                    If getValueStringOtherSystems("select Integrate from CompanyDetails", con) = "Yes" Then
                        UpdateOtherSystems()
                    End If
                    MessageBox.Show("Done Saving")
                    Close()
                End If
            End If
        End If
    End Sub

    Public Function checkFieldsEntered() As Boolean
        Dim chk As Boolean = True
        If tbtitle.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Title Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbcountry.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Country Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbcountryid.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Country ID Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbidentitytype.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Identity Type Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbidentitytypeid.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Identity Type ID Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbidnumber.Text.Trim = "" Then
            chk = False
            MessageBox.Show("ID Number Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbfirstnames.Text.Trim = "" Then
            chk = False
            MessageBox.Show("First Name Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbsurname.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Surname Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbgender.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Gender Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbgenderid.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Gender ID Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbcellnumber.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Cedll Number Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbemail.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Email Address Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbpostaladdress.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Postal Address Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbhometelephone.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Telephone Number Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbregion.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Region Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbregionid.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Region ID Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbtown.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Town Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbtownid.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Town ID Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbarea.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Area Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbward.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Ward Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbchief.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Chief Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbhomephysicaladdress.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Home Address Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbschool.Text.Trim = "" Then
            chk = False
            MessageBox.Show("School Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbschoolid.Text.Trim = "" Then
            chk = False
            MessageBox.Show("School ID Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbemployer.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Employer Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbemployerid.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Employer ID Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbdepartment.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Department Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbdepartmentid.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Department ID Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tboccupation.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Occupation Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbworkpostaladdress.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Work Postal Address Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbworkphysicaladdress.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Work Physical Address Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbworkemailaddress.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Work Email Address Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbemploymenttype.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Employment Type Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbemploymenttypeid.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Employment Type ID Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbemployercontactperson.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Employer Contract Person Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbemptelephone1.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Employer Telephone 1 Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbemptelephone2.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Employer Telephone 2 Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbempfax1.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Employer Fax 1 Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbempfax2.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Employer Fax 2 Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbnextofkinname1.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Next Of Kin Name 1 Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbnextofkinid1.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Next Of Kin ID 1 Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbnextofkincontact1.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Next Of Kin Contact 1 Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbnextofkinaddress1.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Next Of Kin Address 1 Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbnextofkinname2.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Next Of Kin Name 2 Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbnextofkinid2.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Next Of Kin ID 2 Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbnextofkincontact2.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Next Of Kin Contact 2 Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbnextofkinaddress2.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Next Of Kin Address 2 Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbnetsalary.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Net Salary Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbbankname.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Bank Name Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbbranchname.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Branch Name Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbaccountnumber.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Account Number Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbgroup.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Group Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbgroupid.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Group ID Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If tbpayrolnumber.Text.Trim = "" Then
            chk = False
            MessageBox.Show("Payrol Number Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If

        '#######################################################################################################

        If cbedm.Checked = True Then
            If EDMGender = "" Then
                chk = False
                MessageBox.Show("EDM Gender Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            If EDMEmployer = "" Then
                chk = False
                MessageBox.Show("EDM Employer Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
        'END OF EDM #######################################################################################################
        If cbfpm.Checked = True Then
            If FPMGender = "" Then
                chk = False
                MessageBox.Show("FPM Gender Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            If FPMRegionID = "" Then
                chk = False
                MessageBox.Show("FPM Region Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            If FPMLocation = "" Then
                chk = False
                MessageBox.Show("FPM Location Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            If FPMSchool = "" Then
                chk = False
                MessageBox.Show("FPM School Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            If FPMDepartment = "" Then
                chk = False
                MessageBox.Show("FPM Department Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            If FPMGroup = "" Then
                chk = False
                MessageBox.Show("FPM Group Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
        'END OF FPM #######################################################################################################
        If cbmlma.Checked = True Then
            If MLMAGender = "" Then
                chk = False
                MessageBox.Show("MLM Advances Gender Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            If MLMALocation = "" Then
                chk = False
                MessageBox.Show("MLM Advances Location Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            If MLMAEmployer = "" Then
                chk = False
                MessageBox.Show("MLM Advances Employer Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            If MLMAEmployerName = "" Then
                chk = False
                MessageBox.Show("MLM Advances Employer Name Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            If MLMADepartment = "" Then
                chk = False
                MessageBox.Show("MLM Advances Department Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            If MLMAEmpType = "" Then
                chk = False
                MessageBox.Show("MLM Advances Employment Type Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            If MLMAGroup = "" Then
                chk = False
                MessageBox.Show("MLM Advances Group Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            If MLMABankName = "" Then
                chk = False
                MessageBox.Show("MLM Advances Bank Name Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            If MLMABankBranchName = "" Then
                chk = False
                MessageBox.Show("MLM Advances Bank Branch Name Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
        'END OF MLMA #######################################################################################################
        If cbmlme.Checked = True Then
            If MLMEGender = "" Then
                chk = False
                MessageBox.Show("MLM Educational Gender Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            If MLMELocation = "" Then
                chk = False
                MessageBox.Show("MLM Educational Location Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            If MLMEEmployer = "" Then
                chk = False
                MessageBox.Show("MLM Educational Employer Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            If MLMEEmployerName = "" Then
                chk = False
                MessageBox.Show("MLM Educational Employer Name Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            If MLMEDepartment = "" Then
                chk = False
                MessageBox.Show("MLM Educational Department Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            If MLMEEmpType = "" Then
                chk = False
                MessageBox.Show("MLM Educational Employment Type Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            If MLMEGroup = "" Then
                chk = False
                MessageBox.Show("MLM Educational Group Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            If MLMEBankBranchName = "" Then
                chk = False
                MessageBox.Show("MLM Educational Bank Branch Name Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            If MLMEBankName = "" Then
                chk = False
                MessageBox.Show("MLM Educational Bank Name Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
        'END OF MLME #######################################################################################################
        If cbrmo.Checked = True Then
            If RMOGender = "" Then
                chk = False
                MessageBox.Show("RM Orange Gender Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            If RMORegionID = "" Then
                chk = False
                MessageBox.Show("RM Orange Region Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            If RMOLocation = "" Then
                chk = False
                MessageBox.Show("RM Orange Location Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            If RMOEmployer = "" Then
                chk = False
                MessageBox.Show("RM Orange Employer Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            If RMODepartment = "" Then
                chk = False
                MessageBox.Show("RM Orange Department Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
        'END OF RMO #######################################################################################################
        If cbrmb.Checked = True Then
            If RMBGender = "" Then
                chk = False
                MessageBox.Show("RM Bemobile Gender Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            If RMBRegionID = "" Then
                chk = False
                MessageBox.Show("RM Bemobile Region Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            If RMBLocation = "" Then
                chk = False
                MessageBox.Show("RM Bemobile Location Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            If RMBEmployer = "" Then
                chk = False
                MessageBox.Show("RM Bemobile Employer Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            If RMBDepartment = "" Then
                chk = False
                MessageBox.Show("RM Bemobile Department Empty", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
        'ENDI OF RMB #######################################################################################################

        Return chk
    End Function

    Public Sub setValuesForNewEntryThroughUpdate()

        If cbedm.Checked = True And cbedm.Enabled = True Then
            If EDMGender = "" Then
                EDMGender = tbgender.Text.Trim
            End If
            If EDMEmployer = "" Then
                EDMEmployer = tbemployerid.Text.Trim
            End If
        End If
        'END OF EDM #######################################################################################################
        If cbfpm.Checked = True And cbfpm.Enabled = True Then

            If FPMGender = "" Then
                FPMGender = tbgender.Text.Trim
            End If
            If FPMRegionID = "" Then
                FPMRegionID = getValueStringOtherSystems("select FPMRegion from Region where RegionID = '" & tbregionid.Text.Trim & "'", con)
            End If
            If FPMLocation = "" Then
                FPMLocation = getValueStringOtherSystems("select FPMLocation from Location where LocationID = '" & tbtownid.Text.Trim & "'", con)
            End If
            If FPMSchool = "" Then
                FPMSchool = getValueStringOtherSystems("select FPMSchool from School where SchoolID = '" & tbschoolid.Text.Trim & "'", con)
            End If
            If FPMDepartment = "" Then
                FPMDepartment = getValueStringOtherSystems("select FPMDepartment from Department where DepartmentID = '" & tbschoolid.Text.Trim & "'", con)
            End If
            If FPMGroup = "" Then
                FPMGroup = getValueStringOtherSystems("select FPMGroup from Groups where GroupID = '" & tbschoolid.Text.Trim & "'", con)
            End If
        End If
        'END OF FPM #######################################################################################################
        If cbmlma.Checked = True And cbmlma.Enabled = True Then
            If MLMAGender = "" Then
                MLMAGender = tbgender.Text.Trim
            End If
            If MLMALocation = "" Then
                MLMALocation = tbtown.Text.Trim
            End If
            If MLMAEmployer = "" Then
                MLMAEmployer = getValueStringOtherSystems("select MLMAEmployer from Employers where AutoID = '" & tbemployerid.Text.Trim & "'", con)
            End If
            If MLMAEmployerName = "" Then
                MLMAEmployerName = getValueStringOtherSystems("select Employer from Employer where AutoID = '" & MLMAEmployer & "'", conMLMAdvances)
            End If
            If MLMADepartment = "" Then
                MLMADepartment = getValueStringOtherSystems("select MLMADepartment from Department where DepartmentID = '" & tbdepartmentid.Text.Trim & "'", con)
            End If
            If MLMAEmpType = "" Then
                MLMAEmpType = getValueStringOtherSystems("select MLMAEmpType from EmploymentType where EmploymentTypeID = '" & tbemploymenttypeid.Text.Trim & "'", con)
            End If
            If MLMAGroup = "" Then
                MLMAGroup = getValueStringOtherSystems("select MLMAGroup from Groups where GroupID = '" & tbgroupid.Text.Trim & "'", con)
            End If
            If MLMABankName = "" Then
                MLMABankName = tbbankname.Text.Trim
            End If
            If MLMABankBranchName = "" Then
                MLMABankBranchName = tbbranchname.Text.Trim
            End If
        End If
        'END OF MLMA #######################################################################################################
        If cbmlme.Checked = True And cbmlme.Enabled = True Then
            If MLMEGender = "" Then
                MLMEGender = tbgender.Text.Trim
            End If
            If MLMELocation = "" Then
                MLMELocation = tbtown.Text.Trim
            End If
            If MLMEEmployer = "" Then
                MLMEEmployer = getValueStringOtherSystems("select MLMEEmployer from Employers where AutoID = '" & tbemployerid.Text.Trim & "'", con)
            End If
            If MLMEEmployerName = "" Then
                MLMEEmployerName = getValueStringOtherSystems("select Employer from Employer where AutoID = '" & MLMEEmployer & "'", conMLMEducational)
            End If
            If MLMEDepartment = "" Then
                MLMEDepartment = getValueStringOtherSystems("select MLMEDepartment from Department where DepartmentID = '" & tbdepartmentid.Text.Trim & "'", con)
            End If
            If MLMEEmpType = "" Then
                MLMEEmpType = getValueStringOtherSystems("select MLMEEmpType from EmploymentType where EmploymentTypeID = '" & tbemploymenttypeid.Text.Trim & "'", con)
            End If
            If MLMEGroup = "" Then
                MLMEGroup = getValueStringOtherSystems("select MLMEGroup from Groups where GroupID = '" & tbgroupid.Text.Trim & "'", con)
            End If
            If MLMEBankName = "" Then
                MLMEBankName = tbbankname.Text.Trim
            End If
            If MLMEBankBranchName = "" Then
                MLMEBankBranchName = tbbranchname.Text.Trim
            End If
        End If
        'END OF MLME #######################################################################################################
        If cbrmo.Checked = True And cbrmo.Enabled = True Then
            If RMOGender = "" Then
                RMOGender = tbgender.Text.Trim
            End If
            If RMORegionID = "" Then
                RMORegionID = getValueStringOtherSystems("select RMORegion from Region where RegionID = '" & tbregionid.Text.Trim & "'", con)
            End If
            If RMOLocation = "" Then
                RMOLocation = getValueStringOtherSystems("select RMOLocation from Location where LocationID = '" & tbtownid.Text.Trim & "'", con)
            End If
            If RMOEmployer = "" Then
                RMOEmployer = getValueStringOtherSystems("select RMOEmployer from Employers where AutoID = '" & tbemployerid.Text.Trim & "'", con)
            End If
            If RMODepartment = "" Then
                RMODepartment = getValueStringOtherSystems("select RMODepartment from Department where DepartmentID = '" & tbdepartmentid.Text.Trim & "'", con)
            End If
        End If
        'END OF RMO #######################################################################################################
        If cbrmb.Checked = True And cbrmb.Enabled = True Then
            If RMBGender = "" Then
                RMBGender = tbgender.Text.Trim
            End If
            If RMBRegionID = "" Then
                RMBRegionID = getValueStringOtherSystems("select RMBRegion from Region where RegionID = '" & tbregionid.Text.Trim & "'", con)
            End If
            If RMBLocation = "" Then
                RMBLocation = getValueStringOtherSystems("select RMBLocation from Location where LocationID = '" & tbtownid.Text.Trim & "'", con)
            End If
            If RMBEmployer = "" Then
                RMBEmployer = getValueStringOtherSystems("select RMBEmployer from Employers where AutoID = '" & tbemployerid.Text.Trim & "'", con)
            End If
            If RMBDepartment = "" Then
                RMBDepartment = getValueStringOtherSystems("select RMBDepartment from Department where DepartmentID = '" & tbdepartmentid.Text.Trim & "'", con)
            End If
        End If
        'ENDI OF RMB #######################################################################################################
    End Sub

    Public Sub UpdateOtherSystems()
        If cbedm.Checked = True Then
            If cbedm.Enabled = False Then
                updateCustomerDetails(tbidnumber.Text.Trim, tbfirstnames.Text.Trim & " " & tbsurname.Text.Trim, EDMEmployer, EDMGender)
            ElseIf cbedm.Enabled = True Then
                saveCustomerDetails(tbidnumber.Text.Trim, tbfirstnames.Text.Trim & " " & tbsurname.Text.Trim, EDMEmployer, EDMGender)
            End If
        End If

        If cbfpm.Checked = True Then
            If cbfpm.Enabled = False Then
                UpdateFPMCustomerDetails(tbtitle.Text.Trim, tbidnumber.Text.Trim, tbfirstnames.Text.Trim, tbsurname.Text.Trim, dtpdateofbirth.Value, FPMGender, tbhometelephone.Text.Trim, tbpostaladdress.Text.Trim, tbcellnumber.Text.Trim, tbemail.Text.Trim, FPMLocation, FPMGroup, tboccupation.Text.Trim, tbcountry.Text.Trim, tbhomephysicaladdress.Text.Trim, FPMSchool, FPMDepartment, FPMRegionID, conFPM)
            ElseIf cbfpm.Enabled = True Then
                SaveFPMCustomerDetails(tbtitle.Text.Trim, tbidnumber.Text.Trim, tbfirstnames.Text.Trim, tbsurname.Text.Trim, dtpdateofbirth.Value, FPMGender, tbhometelephone.Text.Trim, tbpostaladdress.Text.Trim, tbcellnumber.Text.Trim, tbemail.Text.Trim, FPMLocation, FPMGroup, tboccupation.Text.Trim, "Active", Date.Now, "", Date.Now, tbcountry.Text.Trim, tbhomephysicaladdress.Text.Trim, FPMSchool, FPMDepartment, FPMRegionID, conFPM)
            End If
        End If

        If cbmlma.Checked = True Then
            If cbmlma.Enabled = False Then
                updateMLMCustomerDetails(tbfirstnames.Text.Trim, tbsurname.Text.Trim, tbidnumber.Text.Trim, dtpdateofbirth.Value, tbcellnumber.Text.Trim, tbhometelephone.Text.Trim, MLMAEmployerName, tbemptelephone1.Text.Trim, tbpostaladdress.Text.Trim, tbhomephysicaladdress.Text.Trim, MLMALocation, tbward.Text.Trim, tbchief.Text.Trim, tbnextofkinname1.Text.Trim, tbnextofkinid1.Text.Trim, tbnextofkincontact1.Text.Trim, tbnextofkinaddress1.Text.Trim, tbnextofkinname2.Text.Trim, tbnextofkinid2.Text.Trim, tbnextofkincontact2.Text.Trim, tbnextofkinaddress2.Text.Trim, MLMAGender, tbemployercontactperson.Text.Trim, MLMAEmpType, dtpdateofengagement.Value, Decimal.Parse(tbnetsalary.Text.Trim), MLMABankName, MLMABankBranchName, tbaccountnumber.Text.Trim, Integer.Parse(MLMAEmployer), Integer.Parse(MLMADepartment), Integer.Parse(MLMAGroup), conMLMAdvances)
            ElseIf cbmlma.Enabled = True Then
                saveMLMCustomerDetails(tbfirstnames.Text.Trim, tbsurname.Text.Trim, tbidnumber.Text.Trim, dtpdateofbirth.Value, tbcellnumber.Text.Trim, tbhometelephone.Text.Trim, MLMAEmployerName, tbemptelephone1.Text.Trim, tbpostaladdress.Text.Trim, tbhomephysicaladdress.Text.Trim, MLMALocation, tbward.Text.Trim, tbchief.Text.Trim, tbnextofkinname1.Text.Trim, tbnextofkinid1.Text.Trim, tbnextofkincontact1.Text.Trim, tbnextofkinaddress1.Text.Trim, tbnextofkinname2.Text.Trim, tbnextofkinid2.Text.Trim, tbnextofkincontact2.Text.Trim, tbnextofkinaddress2.Text.Trim, MLMAGender, tbemployercontactperson.Text.Trim, MLMAEmpType, dtpdateofengagement.Value, Decimal.Parse(tbnetsalary.Text.Trim), MLMABankName, MLMABankBranchName, tbaccountnumber.Text.Trim, Integer.Parse(MLMABranchID), Integer.Parse(MLMAEmployer), Integer.Parse(MLMADepartment), Integer.Parse(MLMAGroup), conMLMAdvances)
            End If
        End If

        If cbmlme.Checked = True Then
            If cbmlme.Enabled = False Then
                updateMLMCustomerDetails(tbfirstnames.Text.Trim, tbsurname.Text.Trim, tbidnumber.Text.Trim, dtpdateofbirth.Value, tbcellnumber.Text.Trim, tbhometelephone.Text.Trim, MLMEEmployerName, tbemptelephone1.Text.Trim, tbpostaladdress.Text.Trim, tbhomephysicaladdress.Text.Trim, MLMELocation, tbward.Text.Trim, tbchief.Text.Trim, tbnextofkinname1.Text.Trim, tbnextofkinid1.Text.Trim, tbnextofkincontact1.Text.Trim, tbnextofkinaddress1.Text.Trim, tbnextofkinname2.Text.Trim, tbnextofkinid2.Text.Trim, tbnextofkincontact2.Text.Trim, tbnextofkinaddress2.Text.Trim, MLMEGender, tbemployercontactperson.Text.Trim, MLMEEmpType, dtpdateofengagement.Value, Decimal.Parse(tbnetsalary.Text.Trim), MLMEBankName, MLMEBankBranchName, tbaccountnumber.Text.Trim, Integer.Parse(MLMEEmployer), Integer.Parse(MLMEDepartment), Integer.Parse(MLMEGroup), conMLMEducational)
            ElseIf cbmlme.Enabled = True Then
                saveMLMCustomerDetails(tbfirstnames.Text.Trim, tbsurname.Text.Trim, tbidnumber.Text.Trim, dtpdateofbirth.Value, tbcellnumber.Text.Trim, tbhometelephone.Text.Trim, MLMEEmployerName, tbemptelephone1.Text.Trim, tbpostaladdress.Text.Trim, tbhomephysicaladdress.Text.Trim, MLMELocation, tbward.Text.Trim, tbchief.Text.Trim, tbnextofkinname1.Text.Trim, tbnextofkinid1.Text.Trim, tbnextofkincontact1.Text.Trim, tbnextofkinaddress1.Text.Trim, tbnextofkinname2.Text.Trim, tbnextofkinid2.Text.Trim, tbnextofkincontact2.Text.Trim, tbnextofkinaddress2.Text.Trim, MLMEGender, tbemployercontactperson.Text.Trim, MLMEEmpType, dtpdateofengagement.Value, Decimal.Parse(tbnetsalary.Text.Trim), MLMEBankName, MLMEBankBranchName, tbaccountnumber.Text.Trim, Integer.Parse(MLMEBranchID), Integer.Parse(MLMEEmployer), Integer.Parse(MLMEDepartment), Integer.Parse(MLMEGroup), conMLMEducational)
            End If
        End If

        If cbrmo.Checked = True Then
            If cbrmo.Enabled = False Then
                updateRMCustomerDetails(tbcountry.Text.Trim, tbidnumber.Text.Trim, tbsurname.Text.Trim, tbfirstnames.Text.Trim, RMOGender, RMOEmployer, RMODepartment, tbemptelephone1.Text.Trim, tbemptelephone2.Text.Trim, tbempfax1.Text.Trim, tbempfax2.Text.Trim, tboccupation.Text.Trim, RMOLocation, tbpostaladdress.Text.Trim, tbhomephysicaladdress.Text.Trim, tbhomephysicaladdress.Text.Trim, tbnextofkinname1.Text.Trim, tbnextofkinid1.Text.Trim, tbnextofkincontact1.Text.Trim, tbcellnumber.Text.Trim, RMORegionID, RMOLocation, tbarea.Text.Trim, tbward.Text.Trim, tbchief.Text.Trim, tbhometelephone.Text.Trim, tbnextofkinname2.Text.Trim, tbnextofkinid2.Text.Trim, tbnextofkincontact2.Text.Trim, dtpdateofbirth.Value, tbworkpostaladdress.Text.Trim, tbworkphysicaladdress.Text.Trim, tbworkemailaddress.Text.Trim, tbemail.Text.Trim, tbnextofkinaddress1.Text.Trim, tbnextofkinaddress2.Text.Trim, 1, dtpindentityexpirydate.Value, conRMOrange)
            ElseIf cbrmo.Enabled = True Then
                saveRMCustomerDetails(tbcountry.Text.Trim, tbidnumber.Text.Trim, tbsurname.Text.Trim, tbfirstnames.Text.Trim, RMOGender, RMOEmployer, RMODepartment, tbemptelephone1.Text.Trim, tbemptelephone2.Text.Trim, tbempfax1.Text.Trim, tbempfax2.Text.Trim, tboccupation.Text.Trim, RMOLocation, tbpostaladdress.Text.Trim, tbhomephysicaladdress.Text.Trim, tbhomephysicaladdress.Text.Trim, tbnextofkinname1.Text.Trim, tbnextofkinid1.Text.Trim, tbnextofkincontact1.Text.Trim, tbcellnumber.Text.Trim, RMORegionID, RMOLocation, tbarea.Text.Trim, tbward.Text.Trim, tbchief.Text.Trim, tbhometelephone.Text.Trim, tbnextofkinname2.Text.Trim, tbnextofkinid2.Text.Trim, tbnextofkincontact2.Text.Trim, dtpdateofbirth.Value, tbworkpostaladdress.Text.Trim, tbworkphysicaladdress.Text.Trim, tbworkemailaddress.Text.Trim, tbemail.Text.Trim, tbnextofkinaddress1.Text.Trim, tbnextofkinaddress2.Text.Trim, 1, Integer.Parse(RMOBranchID), dtpindentityexpirydate.Value, conRMOrange)
            End If
        End If

        If cbrmb.Checked = True Then
            If cbrmb.Enabled = False Then
                updateRMCustomerDetails(tbcountry.Text.Trim, tbidnumber.Text.Trim, tbsurname.Text.Trim, tbfirstnames.Text.Trim, RMBGender, RMBEmployer, RMBDepartment, tbemptelephone1.Text.Trim, tbemptelephone2.Text.Trim, tbempfax1.Text.Trim, tbempfax2.Text.Trim, tboccupation.Text.Trim, RMBLocation, tbpostaladdress.Text.Trim, tbhomephysicaladdress.Text.Trim, tbhomephysicaladdress.Text.Trim, tbnextofkinname1.Text.Trim, tbnextofkinid1.Text.Trim, tbnextofkincontact1.Text.Trim, tbcellnumber.Text.Trim, RMBRegionID, RMBLocation, tbarea.Text.Trim, tbward.Text.Trim, tbchief.Text.Trim, tbhometelephone.Text.Trim, tbnextofkinname2.Text.Trim, tbnextofkinid2.Text.Trim, tbnextofkincontact2.Text.Trim, dtpdateofbirth.Value, tbworkpostaladdress.Text.Trim, tbworkphysicaladdress.Text.Trim, tbworkemailaddress.Text.Trim, tbemail.Text.Trim, tbnextofkinaddress1.Text.Trim, tbnextofkinaddress2.Text.Trim, 1, dtpindentityexpirydate.Value, conRMBeMobile)
            ElseIf cbrmb.Enabled = True Then
                saveRMCustomerDetails(tbcountry.Text.Trim, tbidnumber.Text.Trim, tbsurname.Text.Trim, tbfirstnames.Text.Trim, RMBGender, RMBEmployer, RMBDepartment, tbemptelephone1.Text.Trim, tbemptelephone2.Text.Trim, tbempfax1.Text.Trim, tbempfax2.Text.Trim, tboccupation.Text.Trim, RMBLocation, tbpostaladdress.Text.Trim, tbhomephysicaladdress.Text.Trim, tbhomephysicaladdress.Text.Trim, tbnextofkinname1.Text.Trim, tbnextofkinid1.Text.Trim, tbnextofkincontact1.Text.Trim, tbcellnumber.Text.Trim, RMBRegionID, RMBLocation, tbarea.Text.Trim, tbward.Text.Trim, tbchief.Text.Trim, tbhometelephone.Text.Trim, tbnextofkinname2.Text.Trim, tbnextofkinid2.Text.Trim, tbnextofkincontact2.Text.Trim, dtpdateofbirth.Value, tbworkpostaladdress.Text.Trim, tbworkphysicaladdress.Text.Trim, tbworkemailaddress.Text.Trim, tbemail.Text.Trim, tbnextofkinaddress1.Text.Trim, tbnextofkinaddress2.Text.Trim, 1, Integer.Parse(RMBBranchID), dtpindentityexpirydate.Value, conRMBeMobile)
            End If
        End If
    End Sub

End Class