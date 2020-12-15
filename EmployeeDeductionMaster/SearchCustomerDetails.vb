Public Class SearchCustomerDetails

    Dim ds As New DataSet

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        DataGridView1.Rows.Clear()
        'If checkCustomerInCustomer(tbcustomerid.Text.Trim) = True Then
        '    Dim vrl As New Customer
        '    vrl.MdiParent = mform
        '    vrl.tbidnumber.Text = tbcustomerid.Text.Trim
        '    vrl.Show()
        '    vrl.getCustomerDetails()
        '    Me.Close()
        'Else
        ''MLM Advances
        ds.Clear()
            ds = getCustomerDetailsFromSystems(conMLMAdvances, "select '' as Title,'' as CountryID,'' as IdentityTypeID,'' as IdentityExpiryDate,ID_Number,First_Name,Surname,Date_Of_Birth,Gender,Mobile_Number,'' as Email,Postal_Address,Home_Telephone,'' as Region,Home_Village,'' as Area,Home_Ward,Chief,Residential_Address,'' as SchoolID,Name_Of_Employer,DepartmentID,'' as Occupation,'' as WorkPostalAddress,'' as WorkPhysicalAddress,'' as WorkEmailAddress,EmployerType,EmployerContactPerson,DateOfEngagement,Employer_Telephone,'' as EmpTelephone2,'' as EmpFax_1,'' as EmpFax_2,Next_Of_Kin_1_Name,Next_Of_Kin_1_ID_Number,Next_Of_Kin_1_Mobile_Number1,Next_Of_Kin_1_Address,Next_Of_Kin_2_Name,Next_Of_Kin_2_ID_Number,Next_Of_Kin_2_Mobile_Number1,Next_Of_Kin_2_Address,NetSalary,BankName,BranchName,AccountNumber,'' as GroupID,ID_Number from CustomerDetails where ID_Number = '" & tbcustomerid.Text.Trim & "'")
            For count As Integer = 0 To ds.Tables(0).Rows.Count - 1
                DataGridView1.Rows.Add()
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = "MLM Advances"
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = getDateValue("select CreationDate from LoanDetails where Customer_ID = '" & tbcustomerid.Text.Trim & "' Order by CreationDate DESC", conMLMAdvances)
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = ds.Tables(0).Rows(count).Item(0).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = ds.Tables(0).Rows(count).Item(1).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(4).Value = ds.Tables(0).Rows(count).Item(2).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(5).Value = ds.Tables(0).Rows(count).Item(3).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(6).Value = ds.Tables(0).Rows(count).Item(4).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(7).Value = ds.Tables(0).Rows(count).Item(5).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(8).Value = ds.Tables(0).Rows(count).Item(6).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(9).Value = ds.Tables(0).Rows(count).Item(7).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(10).Value = ds.Tables(0).Rows(count).Item(8).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(11).Value = ds.Tables(0).Rows(count).Item(9).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(12).Value = ds.Tables(0).Rows(count).Item(10).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(13).Value = ds.Tables(0).Rows(count).Item(11).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(14).Value = ds.Tables(0).Rows(count).Item(12).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(15).Value = ds.Tables(0).Rows(count).Item(13).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(16).Value = ds.Tables(0).Rows(count).Item(14).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(17).Value = ds.Tables(0).Rows(count).Item(15).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(18).Value = ds.Tables(0).Rows(count).Item(16).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(19).Value = ds.Tables(0).Rows(count).Item(17).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(20).Value = ds.Tables(0).Rows(count).Item(18).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(21).Value = ds.Tables(0).Rows(count).Item(19).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(22).Value = ds.Tables(0).Rows(count).Item(20).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(23).Value = getStringValue("select DepartmentName from EmployerDepartment where EmployerID = '" & ds.Tables(0).Rows(count).Item(21).ToString & "'", conMLMAdvances)
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(24).Value = ds.Tables(0).Rows(count).Item(22).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(25).Value = ds.Tables(0).Rows(count).Item(23).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(26).Value = ds.Tables(0).Rows(count).Item(24).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(27).Value = ds.Tables(0).Rows(count).Item(25).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(28).Value = ds.Tables(0).Rows(count).Item(26).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(29).Value = ds.Tables(0).Rows(count).Item(27).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(30).Value = ds.Tables(0).Rows(count).Item(28).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(31).Value = ds.Tables(0).Rows(count).Item(29).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(32).Value = ds.Tables(0).Rows(count).Item(30).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(33).Value = ds.Tables(0).Rows(count).Item(31).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(34).Value = ds.Tables(0).Rows(count).Item(32).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(35).Value = ds.Tables(0).Rows(count).Item(33).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(36).Value = ds.Tables(0).Rows(count).Item(34).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(37).Value = ds.Tables(0).Rows(count).Item(35).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(38).Value = ds.Tables(0).Rows(count).Item(36).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(39).Value = ds.Tables(0).Rows(count).Item(37).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(40).Value = ds.Tables(0).Rows(count).Item(38).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(41).Value = ds.Tables(0).Rows(count).Item(39).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(42).Value = ds.Tables(0).Rows(count).Item(40).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(43).Value = ds.Tables(0).Rows(count).Item(41).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(44).Value = ds.Tables(0).Rows(count).Item(42).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(45).Value = ds.Tables(0).Rows(count).Item(43).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(46).Value = ds.Tables(0).Rows(count).Item(44).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(47).Value = ds.Tables(0).Rows(count).Item(45).ToString
            Next

            ''MLM Educational
            ds.Clear()
            ds = getCustomerDetailsFromSystems(conMLMEducational, "select '' as Title,'' as CountryID,'' as IdentityTypeID,'' as IdentityExpiryDate,ID_Number,First_Name,Surname,Date_Of_Birth,Gender,Mobile_Number,'' as Email,Postal_Address,Home_Telephone,'' as Region,Home_Village,'' as Area,Home_Ward,Chief,Residential_Address,'' as SchoolID,Name_Of_Employer,DepartmentID,'' as Occupation,'' as WorkPostalAddress,'' as WorkPhysicalAddress,'' as WorkEmailAddress,EmployerType,EmployerContactPerson,DateOfEngagement,Employer_Telephone,'' as EmpTelephone2,'' as EmpFax_1,'' as EmpFax_2,Next_Of_Kin_1_Name,Next_Of_Kin_1_ID_Number,Next_Of_Kin_1_Mobile_Number1,Next_Of_Kin_1_Address,Next_Of_Kin_2_Name,Next_Of_Kin_2_ID_Number,Next_Of_Kin_2_Mobile_Number1,Next_Of_Kin_2_Address,NetSalary,BankName,BranchName,AccountNumber,'' as GroupID,ID_Number from CustomerDetails where ID_Number = '" & tbcustomerid.Text.Trim & "'")
            For count As Integer = 0 To ds.Tables(0).Rows.Count - 1
                DataGridView1.Rows.Add()
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = "MLM Educational"
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = getDateValue("select CreationDate from LoanDetails where Customer_ID = '" & tbcustomerid.Text.Trim & "' Order by CreationDate DESC", conMLMEducational)
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = ds.Tables(0).Rows(count).Item(0).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = ds.Tables(0).Rows(count).Item(1).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(4).Value = ds.Tables(0).Rows(count).Item(2).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(5).Value = ds.Tables(0).Rows(count).Item(3).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(6).Value = ds.Tables(0).Rows(count).Item(4).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(7).Value = ds.Tables(0).Rows(count).Item(5).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(8).Value = ds.Tables(0).Rows(count).Item(6).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(9).Value = ds.Tables(0).Rows(count).Item(7).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(10).Value = ds.Tables(0).Rows(count).Item(8).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(11).Value = ds.Tables(0).Rows(count).Item(9).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(12).Value = ds.Tables(0).Rows(count).Item(10).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(13).Value = ds.Tables(0).Rows(count).Item(11).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(14).Value = ds.Tables(0).Rows(count).Item(12).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(15).Value = ds.Tables(0).Rows(count).Item(13).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(16).Value = ds.Tables(0).Rows(count).Item(14).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(17).Value = ds.Tables(0).Rows(count).Item(15).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(18).Value = ds.Tables(0).Rows(count).Item(16).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(19).Value = ds.Tables(0).Rows(count).Item(17).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(20).Value = ds.Tables(0).Rows(count).Item(18).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(21).Value = ds.Tables(0).Rows(count).Item(19).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(22).Value = ds.Tables(0).Rows(count).Item(20).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(23).Value = getStringValue("select DepartmentName from EmployerDepartment where EmployerID = '" & ds.Tables(0).Rows(count).Item(21).ToString & "'", conMLMEducational)
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(24).Value = ds.Tables(0).Rows(count).Item(22).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(25).Value = ds.Tables(0).Rows(count).Item(23).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(26).Value = ds.Tables(0).Rows(count).Item(24).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(27).Value = ds.Tables(0).Rows(count).Item(25).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(28).Value = ds.Tables(0).Rows(count).Item(26).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(29).Value = ds.Tables(0).Rows(count).Item(27).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(30).Value = ds.Tables(0).Rows(count).Item(28).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(31).Value = ds.Tables(0).Rows(count).Item(29).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(32).Value = ds.Tables(0).Rows(count).Item(30).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(33).Value = ds.Tables(0).Rows(count).Item(31).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(34).Value = ds.Tables(0).Rows(count).Item(32).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(35).Value = ds.Tables(0).Rows(count).Item(33).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(36).Value = ds.Tables(0).Rows(count).Item(34).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(37).Value = ds.Tables(0).Rows(count).Item(35).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(38).Value = ds.Tables(0).Rows(count).Item(36).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(39).Value = ds.Tables(0).Rows(count).Item(37).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(40).Value = ds.Tables(0).Rows(count).Item(38).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(41).Value = ds.Tables(0).Rows(count).Item(39).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(42).Value = ds.Tables(0).Rows(count).Item(40).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(43).Value = ds.Tables(0).Rows(count).Item(41).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(44).Value = ds.Tables(0).Rows(count).Item(42).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(45).Value = ds.Tables(0).Rows(count).Item(43).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(46).Value = ds.Tables(0).Rows(count).Item(44).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(47).Value = ds.Tables(0).Rows(count).Item(45).ToString
            Next

            ''Orange
            ds.Clear()
            ds = getCustomerDetailsFromSystems(conRMOrange, "select '' as Title,'' as CountryID,'' as IdentityTypeID,'' as IdentityExpiryDate,Account,FirstName,Surname,DateOFBirth,Gender,Contact,YourEmailAddress,PostalAddress,Homephone,Region,Town_Village,Area,Ward,Chief,CurrentPhysicalAddress,'' as SchoolID,EmployerCode,EmployerDeptCode, Position, WorkPostalAddress,WorkPhysicalAddress,WorkEmailAddress,'' as EmployerType,'' as EmployerContactPerson,'' as DateOfEngagement,EmployerTel_1,EmployerTel_2,EmployerFax_1,EmployerFax_2,NextOfKinName,NextOfKinID,NextOfKinContact,NextOfKinAddress,NextOfKinName2,NextOfKinID2,NextOfKinContact2,NextOfKinAddress2,'' as NetSalary,'' as BankName,'' as BranchName,'' as AccountNumber,'' as GroupID,Account from CustomerMaster where Account = '" & tbcustomerid.Text.Trim & "'")
            For count As Integer = 0 To ds.Tables(0).Rows.Count - 1
                DataGridView1.Rows.Add()
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = "RM Orange"
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = getDateValue("select CreationDate from CustomerDeals where Account = '" & tbcustomerid.Text.Trim & "' Order by CreationDate DESC", conRMOrange)
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = ds.Tables(0).Rows(count).Item(0).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = ds.Tables(0).Rows(count).Item(1).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(4).Value = ds.Tables(0).Rows(count).Item(2).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(5).Value = ds.Tables(0).Rows(count).Item(3).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(6).Value = ds.Tables(0).Rows(count).Item(4).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(7).Value = ds.Tables(0).Rows(count).Item(5).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(8).Value = ds.Tables(0).Rows(count).Item(6).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(9).Value = ds.Tables(0).Rows(count).Item(7).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(10).Value = ds.Tables(0).Rows(count).Item(8).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(11).Value = ds.Tables(0).Rows(count).Item(9).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(12).Value = ds.Tables(0).Rows(count).Item(10).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(13).Value = ds.Tables(0).Rows(count).Item(11).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(14).Value = ds.Tables(0).Rows(count).Item(12).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(15).Value = getStringValue("select RegionName from Region where RegionID = '" & ds.Tables(0).Rows(count).Item(13).ToString & "'", conRMOrange)
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(16).Value = getStringValue("select LocationName from Location where LocationID = '" & ds.Tables(0).Rows(count).Item(14).ToString & "'", conRMOrange)
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(17).Value = ds.Tables(0).Rows(count).Item(15).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(18).Value = ds.Tables(0).Rows(count).Item(16).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(19).Value = ds.Tables(0).Rows(count).Item(17).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(20).Value = ds.Tables(0).Rows(count).Item(18).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(21).Value = ds.Tables(0).Rows(count).Item(19).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(22).Value = getStringValue("select Employer_Name from Employer where Code = '" & ds.Tables(0).Rows(count).Item(20).ToString & "'", conRMOrange)
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(23).Value = getStringValue("select Department_Name from Department where DepartmentCode = '" & ds.Tables(0).Rows(count).Item(21).ToString & "'", conRMOrange)
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(24).Value = ds.Tables(0).Rows(count).Item(22).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(25).Value = ds.Tables(0).Rows(count).Item(23).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(26).Value = ds.Tables(0).Rows(count).Item(24).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(27).Value = ds.Tables(0).Rows(count).Item(25).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(28).Value = ds.Tables(0).Rows(count).Item(26).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(29).Value = ds.Tables(0).Rows(count).Item(27).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(30).Value = ds.Tables(0).Rows(count).Item(28).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(31).Value = ds.Tables(0).Rows(count).Item(29).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(32).Value = ds.Tables(0).Rows(count).Item(30).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(33).Value = ds.Tables(0).Rows(count).Item(31).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(34).Value = ds.Tables(0).Rows(count).Item(32).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(35).Value = ds.Tables(0).Rows(count).Item(33).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(36).Value = ds.Tables(0).Rows(count).Item(34).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(37).Value = ds.Tables(0).Rows(count).Item(35).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(38).Value = ds.Tables(0).Rows(count).Item(36).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(39).Value = ds.Tables(0).Rows(count).Item(37).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(40).Value = ds.Tables(0).Rows(count).Item(38).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(41).Value = ds.Tables(0).Rows(count).Item(39).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(42).Value = ds.Tables(0).Rows(count).Item(40).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(43).Value = ds.Tables(0).Rows(count).Item(41).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(44).Value = ds.Tables(0).Rows(count).Item(42).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(45).Value = ds.Tables(0).Rows(count).Item(43).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(46).Value = ds.Tables(0).Rows(count).Item(44).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(47).Value = ds.Tables(0).Rows(count).Item(45).ToString
            Next

            ''Bemobile
            ds.Clear()
            ds = getCustomerDetailsFromSystems(conRMBeMobile, "select '' as Title,'' as CountryID,'' as IdentityTypeID,'' as IdentityExpiryDate,Account,FirstName,Surname,DateOFBirth,Gender,Contact,YourEmailAddress,PostalAddress,Homephone,Region,Town_Village,Area,Ward,Chief,CurrentPhysicalAddress,'' as SchoolID,EmployerCode,EmployerDeptCode, Position, WorkPostalAddress,WorkPhysicalAddress,WorkEmailAddress,'' as EmployerType,'' as EmployerContactPerson,'' as DateOfEngagement,EmployerTel_1,EmployerTel_2,EmployerFax_1,EmployerFax_2,NextOfKinName,NextOfKinID,NextOfKinContact,NextOfKinAddress,NextOfKinName2,NextOfKinID2,NextOfKinContact2,NextOfKinAddress2,'' as NetSalary,'' as BankName,'' as BranchName,'' as AccountNumber,'' as GroupID,Account from CustomerMaster where Account = '" & tbcustomerid.Text.Trim & "'")
            For count As Integer = 0 To ds.Tables(0).Rows.Count - 1
                DataGridView1.Rows.Add()
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = "RM Bemobile"
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = getDateValue("select CreationDate from CustomerDeals where Account = '" & tbcustomerid.Text.Trim & "' Order by CreationDate DESC", conRMBeMobile)
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = ds.Tables(0).Rows(count).Item(0).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = ds.Tables(0).Rows(count).Item(1).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(4).Value = ds.Tables(0).Rows(count).Item(2).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(5).Value = ds.Tables(0).Rows(count).Item(3).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(6).Value = ds.Tables(0).Rows(count).Item(4).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(7).Value = ds.Tables(0).Rows(count).Item(5).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(8).Value = ds.Tables(0).Rows(count).Item(6).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(9).Value = ds.Tables(0).Rows(count).Item(7).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(10).Value = ds.Tables(0).Rows(count).Item(8).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(11).Value = ds.Tables(0).Rows(count).Item(9).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(12).Value = ds.Tables(0).Rows(count).Item(10).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(13).Value = ds.Tables(0).Rows(count).Item(11).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(14).Value = ds.Tables(0).Rows(count).Item(12).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(15).Value = getStringValue("select RegionName from Region where RegionID = '" & ds.Tables(0).Rows(count).Item(13).ToString & "'", conRMBeMobile)
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(16).Value = getStringValue("select LocationName from Location where LocationID = '" & ds.Tables(0).Rows(count).Item(14).ToString & "'", conRMBeMobile)
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(17).Value = ds.Tables(0).Rows(count).Item(15).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(18).Value = ds.Tables(0).Rows(count).Item(16).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(19).Value = ds.Tables(0).Rows(count).Item(17).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(20).Value = ds.Tables(0).Rows(count).Item(18).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(21).Value = ds.Tables(0).Rows(count).Item(19).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(22).Value = getStringValue("select Employer_Name from Employer where Code = '" & ds.Tables(0).Rows(count).Item(20).ToString & "'", conRMBeMobile)
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(23).Value = getStringValue("select Department_Name from Department where DepartmentCode = '" & ds.Tables(0).Rows(count).Item(21).ToString & "'", conRMBeMobile)
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(24).Value = ds.Tables(0).Rows(count).Item(22).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(25).Value = ds.Tables(0).Rows(count).Item(23).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(26).Value = ds.Tables(0).Rows(count).Item(24).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(27).Value = ds.Tables(0).Rows(count).Item(25).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(28).Value = ds.Tables(0).Rows(count).Item(26).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(29).Value = ds.Tables(0).Rows(count).Item(27).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(30).Value = ds.Tables(0).Rows(count).Item(28).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(31).Value = ds.Tables(0).Rows(count).Item(29).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(32).Value = ds.Tables(0).Rows(count).Item(30).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(33).Value = ds.Tables(0).Rows(count).Item(31).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(34).Value = ds.Tables(0).Rows(count).Item(32).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(35).Value = ds.Tables(0).Rows(count).Item(33).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(36).Value = ds.Tables(0).Rows(count).Item(34).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(37).Value = ds.Tables(0).Rows(count).Item(35).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(38).Value = ds.Tables(0).Rows(count).Item(36).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(39).Value = ds.Tables(0).Rows(count).Item(37).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(40).Value = ds.Tables(0).Rows(count).Item(38).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(41).Value = ds.Tables(0).Rows(count).Item(39).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(42).Value = ds.Tables(0).Rows(count).Item(40).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(43).Value = ds.Tables(0).Rows(count).Item(41).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(44).Value = ds.Tables(0).Rows(count).Item(42).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(45).Value = ds.Tables(0).Rows(count).Item(43).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(46).Value = ds.Tables(0).Rows(count).Item(44).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(47).Value = ds.Tables(0).Rows(count).Item(45).ToString
            Next

            ''FPM
            ds.Clear()
        ds = getCustomerDetailsFromSystems(conFPM, "select '' as Title,'' as CountryID,'' as IdentityTypeID,'' as IdentityExpiryDate,IDNumber,FirstName,Surname,DateOFBirth,Gender,CellNumber,Email,PostalAddress,Telephone,RegionID,AreaCode,'' as Area,'' as Ward,'' as Chief,HomeAddress,SchoolID,'' as Employer,DepID, Occupation, '' as WorkPostalAddress,'' as WorkPhysicalAddress,'' as WorkEmailAddress,'' as EmployerType,'' as EmployerContactPerson,'' as DateOfEngagement,'' as EmployerTel_1,'' as EmployerTel_2,'' as EmployerFax_1,'' as EmployerFax_2,'' as NextOfKinName,'' as NextOfKinID,'' as NextOfKinContact,'' as NextOfKinAddress,'' as NextOfKinName2,'' as NextOfKinID2,'' as NextOfKinContact2,'' as NextOfKinAddress2,'' as NetSalary,'' as BankName,'' as BranchName,'' as AccountNumber,GroupID,IDNumber from CustomerDetails where IDNumber = '" & tbcustomerid.Text.Trim & "'")
        For count As Integer = 0 To ds.Tables(0).Rows.Count - 1
                DataGridView1.Rows.Add()
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = "FPM"
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = getDateValue("select CoverStartDate from PolicyCovers where IDNumber = '" & tbcustomerid.Text.Trim & "' Order by CoverStartDate DESC", conFPM)
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = ds.Tables(0).Rows(count).Item(0).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = ds.Tables(0).Rows(count).Item(1).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(4).Value = ds.Tables(0).Rows(count).Item(2).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(5).Value = ds.Tables(0).Rows(count).Item(3).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(6).Value = ds.Tables(0).Rows(count).Item(4).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(7).Value = ds.Tables(0).Rows(count).Item(5).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(8).Value = ds.Tables(0).Rows(count).Item(6).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(9).Value = ds.Tables(0).Rows(count).Item(7).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(10).Value = ds.Tables(0).Rows(count).Item(8).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(11).Value = ds.Tables(0).Rows(count).Item(9).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(12).Value = ds.Tables(0).Rows(count).Item(10).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(13).Value = ds.Tables(0).Rows(count).Item(11).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(14).Value = ds.Tables(0).Rows(count).Item(12).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(15).Value = getStringValue("select RegionName from Regions where RegionID = '" & ds.Tables(0).Rows(count).Item(13).ToString & "'", conFPM)
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(16).Value = getStringValue("select AreaName from Area where AreaCode = '" & ds.Tables(0).Rows(count).Item(14).ToString & "'", conFPM)
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(17).Value = ds.Tables(0).Rows(count).Item(15).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(18).Value = ds.Tables(0).Rows(count).Item(16).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(19).Value = ds.Tables(0).Rows(count).Item(17).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(20).Value = ds.Tables(0).Rows(count).Item(18).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(21).Value = getStringValue("select SchoolName from Schools where SchoolID = '" & ds.Tables(0).Rows(count).Item(19).ToString & "'", conFPM)
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(22).Value = ds.Tables(0).Rows(count).Item(20).ToString
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(23).Value = getStringValue("select DepartmentName from Departments where DepID = '" & ds.Tables(0).Rows(count).Item(21).ToString & "'", conFPM)
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(24).Value = ds.Tables(0).Rows(count).Item(22).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(25).Value = ds.Tables(0).Rows(count).Item(23).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(26).Value = ds.Tables(0).Rows(count).Item(24).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(27).Value = ds.Tables(0).Rows(count).Item(25).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(28).Value = ds.Tables(0).Rows(count).Item(26).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(29).Value = ds.Tables(0).Rows(count).Item(27).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(30).Value = ds.Tables(0).Rows(count).Item(28).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(31).Value = ds.Tables(0).Rows(count).Item(29).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(32).Value = ds.Tables(0).Rows(count).Item(30).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(33).Value = ds.Tables(0).Rows(count).Item(31).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(34).Value = ds.Tables(0).Rows(count).Item(32).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(35).Value = ds.Tables(0).Rows(count).Item(33).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(36).Value = ds.Tables(0).Rows(count).Item(34).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(37).Value = ds.Tables(0).Rows(count).Item(35).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(38).Value = ds.Tables(0).Rows(count).Item(36).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(39).Value = ds.Tables(0).Rows(count).Item(37).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(40).Value = ds.Tables(0).Rows(count).Item(38).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(41).Value = ds.Tables(0).Rows(count).Item(39).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(42).Value = ds.Tables(0).Rows(count).Item(40).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(43).Value = ds.Tables(0).Rows(count).Item(41).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(44).Value = ds.Tables(0).Rows(count).Item(42).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(45).Value = ds.Tables(0).Rows(count).Item(43).ToString
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(47).Value = getStringValue("select GroupName from Groups where GroupID = '" & ds.Tables(0).Rows(count).Item(45).ToString & "'", conFPM)
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(48).Value = ds.Tables(0).Rows(count).Item(46).ToString
        Next

            ''MLMOnline
            ds.Clear()
            ds = getCustomerDetailsFromSystems(conMLMAdvances, "select '' as Title,'' as CountryID,'' as IdentityTypeID,'' as IdentityExpiryDate,ID_Number,First_Name,Surname,Date_Of_Birth,Gender,Mobile_Number,'' as Email,Postal_Address,Home_Telephone,'' as Region,Home_Village,'' as Area,Home_Ward,Chief,Residential_Address,'' as SchoolID,Name_Of_Employer,Department,'' as Occupation,'' as WorkPostalAddress,'' as WorkPhysicalAddress,'' as WorkEmailAddress,EmployerType,EmployerContactPerson,DateOfEngagement,Employer_Telephone,'' as EmpTelephone2,'' as EmpFax_1,'' as EmpFax_2,Next_Of_Kin_1_Name,Next_Of_Kin_1_ID_Number,Next_Of_Kin_1_Mobile_Number1,Next_Of_Kin_1_Address,Next_Of_Kin_2_Name,Next_Of_Kin_2_ID_Number,Next_Of_Kin_2_Mobile_Number1,Next_Of_Kin_2_Address,NetSalary,BankName,BranchName,AccountNumber,'' as GroupID,ID_Number from CustomerDetailsOnline where ID_Number = '" & tbcustomerid.Text.Trim & "'")
            For count As Integer = 0 To ds.Tables(0).Rows.Count - 1
                DataGridView1.Rows.Add()
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = "MLM Online"
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = getDateValue("select CreationDate from LoansOnline where CustomerID = '" & tbcustomerid.Text.Trim & "' Order by CreationDate DESC", conMLMAdvances)
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = ds.Tables(0).Rows(count).Item(0).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = ds.Tables(0).Rows(count).Item(1).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(4).Value = ds.Tables(0).Rows(count).Item(2).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(5).Value = ds.Tables(0).Rows(count).Item(3).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(6).Value = ds.Tables(0).Rows(count).Item(4).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(7).Value = ds.Tables(0).Rows(count).Item(5).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(8).Value = ds.Tables(0).Rows(count).Item(6).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(9).Value = ds.Tables(0).Rows(count).Item(7).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(10).Value = ds.Tables(0).Rows(count).Item(8).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(11).Value = ds.Tables(0).Rows(count).Item(9).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(12).Value = ds.Tables(0).Rows(count).Item(10).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(13).Value = ds.Tables(0).Rows(count).Item(11).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(14).Value = ds.Tables(0).Rows(count).Item(12).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(15).Value = ds.Tables(0).Rows(count).Item(13).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(16).Value = ds.Tables(0).Rows(count).Item(14).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(17).Value = ds.Tables(0).Rows(count).Item(15).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(18).Value = ds.Tables(0).Rows(count).Item(16).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(19).Value = ds.Tables(0).Rows(count).Item(17).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(20).Value = ds.Tables(0).Rows(count).Item(18).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(21).Value = ds.Tables(0).Rows(count).Item(19).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(22).Value = ds.Tables(0).Rows(count).Item(20).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(23).Value = ds.Tables(0).Rows(count).Item(21).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(24).Value = ds.Tables(0).Rows(count).Item(22).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(25).Value = ds.Tables(0).Rows(count).Item(23).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(26).Value = ds.Tables(0).Rows(count).Item(24).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(27).Value = ds.Tables(0).Rows(count).Item(25).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(28).Value = ds.Tables(0).Rows(count).Item(26).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(29).Value = ds.Tables(0).Rows(count).Item(27).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(30).Value = ds.Tables(0).Rows(count).Item(28).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(31).Value = ds.Tables(0).Rows(count).Item(29).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(32).Value = ds.Tables(0).Rows(count).Item(30).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(33).Value = ds.Tables(0).Rows(count).Item(31).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(34).Value = ds.Tables(0).Rows(count).Item(32).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(35).Value = ds.Tables(0).Rows(count).Item(33).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(36).Value = ds.Tables(0).Rows(count).Item(34).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(37).Value = ds.Tables(0).Rows(count).Item(35).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(38).Value = ds.Tables(0).Rows(count).Item(36).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(39).Value = ds.Tables(0).Rows(count).Item(37).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(40).Value = ds.Tables(0).Rows(count).Item(38).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(41).Value = ds.Tables(0).Rows(count).Item(39).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(42).Value = ds.Tables(0).Rows(count).Item(40).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(43).Value = ds.Tables(0).Rows(count).Item(41).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(44).Value = ds.Tables(0).Rows(count).Item(42).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(45).Value = ds.Tables(0).Rows(count).Item(43).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(46).Value = ds.Tables(0).Rows(count).Item(44).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(47).Value = ds.Tables(0).Rows(count).Item(45).ToString
            Next
            If DataGridView1.Rows.Count = 0 Then
                Dim vrl As New Customer
                vrl.MdiParent = mform
                vrl.tbidnumber.Text = tbcustomerid.Text.Trim
                vrl.Show()
                Me.Close()
            Else
                DataGridView1.Sort(DataGridView1.Columns(1), System.ComponentModel.ListSortDirection.Descending)
            End If
        ''End If
    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        Dim vrl As New Customer
        vrl.MdiParent = mform
        vrl.tbtitle.Text = ""
        vrl.tbothercountry.Text = "Country = " & DataGridView1.CurrentRow.Cells(3).Value
        vrl.tbotheridentitytype.Text = "Identity Type = " & DataGridView1.CurrentRow.Cells(4).Value
        If DataGridView1.CurrentRow.Cells(5).Value = "" Then
            vrl.dtpindentityexpirydate.Value = Date.Now
        Else
            vrl.dtpindentityexpirydate.Value = Date.Parse(DataGridView1.CurrentRow.Cells(5).Value)
        End If
        vrl.tbidnumber.Text = DataGridView1.CurrentRow.Cells(6).Value
        vrl.tbidnumber.Enabled = False
        vrl.tbfirstnames.Text = DataGridView1.CurrentRow.Cells(7).Value
        vrl.tbsurname.Text = DataGridView1.CurrentRow.Cells(8).Value
        If DataGridView1.CurrentRow.Cells(9).Value = "" Then
            vrl.dtpdateofbirth.Value = Date.Now
        Else
            vrl.dtpdateofbirth.Value = Date.Parse(DataGridView1.CurrentRow.Cells(9).Value)
        End If
        vrl.tbothergender.Text = "Gender = " & DataGridView1.CurrentRow.Cells(10).Value
        vrl.tbcellnumber.Text = DataGridView1.CurrentRow.Cells(11).Value
        vrl.tbemail.Text = DataGridView1.CurrentRow.Cells(12).Value
        vrl.tbpostaladdress.Text = DataGridView1.CurrentRow.Cells(13).Value
        vrl.tbhometelephone.Text = DataGridView1.CurrentRow.Cells(14).Value
        vrl.tbotherregion.Text = "Region = "
        For Each row As DataGridViewRow In DataGridView1.Rows
            If row.Index = 0 Then
                vrl.tbotherregion.Text = vrl.tbotherregion.Text & row.Cells(15).Value
            Else
                If row.Cells(15).Value.trim = "" Then

                Else
                    If vrl.tbotherregion.Text.Trim = "" Then
                        vrl.tbotherregion.Text = vrl.tbotherregion.Text & row.Cells(15).Value
                    Else
                        vrl.tbotherregion.Text = vrl.tbotherregion.Text & "/ " & row.Cells(15).Value
                    End If
                End If
            End If
        Next
        vrl.tbothervillage.Text = "Village = "
        For Each row As DataGridViewRow In DataGridView1.Rows
            If row.Index = 0 Then
                vrl.tbothervillage.Text = vrl.tbothervillage.Text & row.Cells(16).Value
            Else
                If row.Cells(16).Value.trim = "" Then

                Else
                    If vrl.tbothervillage.Text.Trim = "" Then
                        vrl.tbothervillage.Text = vrl.tbothervillage.Text & row.Cells(16).Value
                    Else
                        vrl.tbothervillage.Text = vrl.tbothervillage.Text & "/ " & row.Cells(16).Value
                    End If
                End If
            End If
        Next
        vrl.tbarea.Text = DataGridView1.CurrentRow.Cells(17).Value
        vrl.tbward.Text = DataGridView1.CurrentRow.Cells(18).Value
        vrl.tbchief.Text = DataGridView1.CurrentRow.Cells(19).Value
        vrl.tbhomephysicaladdress.Text = DataGridView1.CurrentRow.Cells(20).Value
        vrl.tbotherschool.Text = "School = "
        For Each row As DataGridViewRow In DataGridView1.Rows
            If row.Index = 0 Then
                vrl.tbotherschool.Text = vrl.tbotherschool.Text & row.Cells(21).Value
            Else
                If row.Cells(21).Value.trim = "" Then

                Else
                    If vrl.tbotherschool.Text.Trim = "" Then
                        vrl.tbotherschool.Text = vrl.tbotherschool.Text & row.Cells(21).Value
                    Else
                        vrl.tbotherschool.Text = vrl.tbotherschool.Text & "/ " & row.Cells(21).Value
                    End If
                End If
            End If
        Next
        vrl.tbotheremployer.Text = "Employer = "
        For Each row As DataGridViewRow In DataGridView1.Rows
            If row.Index = 0 Then
                vrl.tbotheremployer.Text = vrl.tbotheremployer.Text & row.Cells(22).Value
            Else
                If row.Cells(22).Value.trim = "" Then

                Else
                    If vrl.tbotheremployer.Text.Trim = "" Then
                        vrl.tbotheremployer.Text = vrl.tbotheremployer.Text & row.Cells(22).Value
                    Else
                        vrl.tbotheremployer.Text = vrl.tbotheremployer.Text & "/ " & row.Cells(22).Value
                    End If
                End If
            End If
        Next
        vrl.tbotherdepartment.Text = "Department = "
        For Each row As DataGridViewRow In DataGridView1.Rows
            If row.Index = 0 Then
                vrl.tbotherdepartment.Text = vrl.tbotherdepartment.Text & row.Cells(23).Value
            Else
                If row.Cells(23).Value.trim = "" Then

                Else
                    If vrl.tbotherdepartment.Text.Trim = "" Then
                        vrl.tbotherdepartment.Text = vrl.tbotherdepartment.Text & row.Cells(23).Value
                    Else
                        vrl.tbotherdepartment.Text = vrl.tbotherdepartment.Text & "/ " & row.Cells(23).Value
                    End If
                End If
            End If
        Next
        vrl.tboccupation.Text = DataGridView1.CurrentRow.Cells(24).Value
        vrl.tbworkpostaladdress.Text = DataGridView1.CurrentRow.Cells(25).Value
        vrl.tbworkphysicaladdress.Text = DataGridView1.CurrentRow.Cells(26).Value
        vrl.tbworkemailaddress.Text = DataGridView1.CurrentRow.Cells(27).Value
        vrl.tbemployercontactperson.Text = DataGridView1.CurrentRow.Cells(29).Value
        If DataGridView1.CurrentRow.Cells(30).Value = "" Then
            vrl.dtpdateofengagement.Value = Date.Now
        Else
            vrl.dtpdateofengagement.Value = Date.Parse(DataGridView1.CurrentRow.Cells(30).Value)
        End If
        vrl.tbemptelephone1.Text = DataGridView1.CurrentRow.Cells(31).Value
        vrl.tbemptelephone2.Text = DataGridView1.CurrentRow.Cells(32).Value
        vrl.tbempfax1.Text = DataGridView1.CurrentRow.Cells(33).Value
        vrl.tbempfax2.Text = DataGridView1.CurrentRow.Cells(34).Value
        vrl.tbnextofkinname1.Text = DataGridView1.CurrentRow.Cells(35).Value
        vrl.tbnextofkinid1.Text = DataGridView1.CurrentRow.Cells(36).Value
        vrl.tbnextofkincontact1.Text = DataGridView1.CurrentRow.Cells(37).Value
        vrl.tbnextofkinaddress1.Text = DataGridView1.CurrentRow.Cells(38).Value
        vrl.tbnextofkinname2.Text = DataGridView1.CurrentRow.Cells(39).Value
        vrl.tbnextofkinid2.Text = DataGridView1.CurrentRow.Cells(40).Value
        vrl.tbnextofkincontact2.Text = DataGridView1.CurrentRow.Cells(41).Value
        vrl.tbnextofkinaddress2.Text = DataGridView1.CurrentRow.Cells(42).Value
        vrl.tbnetsalary.Text = DataGridView1.CurrentRow.Cells(43).Value
        vrl.tbotherbank.Text = "Bank = " & DataGridView1.CurrentRow.Cells(44).Value
        vrl.tbotherbranch.Text = "Branch = " & DataGridView1.CurrentRow.Cells(45).Value
        vrl.tbaccountnumber.Text = DataGridView1.CurrentRow.Cells(46).Value
        vrl.tbothergroup.Text = "Group = "
        For Each row As DataGridViewRow In DataGridView1.Rows
            If row.Index = 0 Then
                vrl.tbothergroup.Text = vrl.tbothergroup.Text & row.Cells(47).Value
            Else
                If row.Cells(47).Value.trim = "" Then

                Else
                    If vrl.tbothergroup.Text.Trim = "" Then
                        vrl.tbothergroup.Text = vrl.tbothergroup.Text & row.Cells(47).Value
                    Else
                        vrl.tbothergroup.Text = vrl.tbothergroup.Text & "/ " & row.Cells(47).Value
                    End If
                End If
            End If
        Next
        vrl.tbpayrolnumber.Text = DataGridView1.CurrentRow.Cells(48).Value
        ''vrl.tbpayrolnumber.Enabled = False
        ''vrl.LinkLabel1.Enabled = False
        vrl.Show()
        vrl.checkCustomerInOtherSYstems()

    End Sub

    Private Sub SearchCustomerDetails_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class