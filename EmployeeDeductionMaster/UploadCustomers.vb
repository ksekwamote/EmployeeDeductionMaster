Imports System.IO

Public Class UploadCustomers

    Dim EDMGender, MLMEGender, MLMAGender, RMOGender, RMBGender, FPMGender As String
    Public RMORegionID, RMBRegionID, FPMRegionID As String
    Public RMOLocation, RMBLocation, FPMLocation, MLMALocation, MLMELocation As String
    Public FPMSchool As String
    Public EDMEmployer, RMOEmployer, RMBEmployer, MLMAEmployer, MLMEEmployer, MLMAEmployerName, MLMEEmployerName As String

    Private Sub LinkLabel4_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel4.LinkClicked
        loadData("select *,'' as EndCol from Gender")
        GroupBox2.Text = "Genders"
    End Sub

    Private Sub LinkLabel5_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel5.LinkClicked
        loadData("select *,'' as EndCol from Region")
        GroupBox2.Text = "Regions"
    End Sub

    Private Sub LinkLabel6_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel6.LinkClicked
        loadData("select *,'' as EndCol from Location")
        GroupBox2.Text = "Locations"
    End Sub

    Private Sub LinkLabel7_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel7.LinkClicked
        loadData("select *,'' as EndCol from Employers")
        GroupBox2.Text = "Employers"
    End Sub

    Private Sub LinkLabel8_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel8.LinkClicked
        loadData("select *,'' as EndCol from Department")
        GroupBox2.Text = "Departments"
    End Sub

    Private Sub LinkLabel9_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel9.LinkClicked
        loadData("select *,'' as EndCol from EmploymentType")
        GroupBox2.Text = "Employment Types"
    End Sub

    Private Sub LinkLabel10_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel10.LinkClicked
        loadData("select *,'' as EndCol from Groups")
        GroupBox2.Text = "Groups"
    End Sub

    Private Sub LinkLabel3_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel3.LinkClicked
        loadData("select *,'' as EndCol from IdentityType")
        GroupBox2.Text = "Identity Types"
    End Sub

    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        loadData("select *,'' as EndCol from Country")
        GroupBox2.Text = "Countries"
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        ExportToExcel(DataGridView1, "")
    End Sub

    Public RMODepartment, RMBDepartment, MLMADepartment, MLMEDepartment, FPMDepartment As String
    Public MLMAEmpType, MLMEEmpType As String
    Public FPMGroup, MLMAGroup, MLMEGroup As String

    Public MLMABankName, MLMABankBranchName, MLMEBankName, MLMEBankBranchName As String
    Dim Nationality As String
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
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

                If getValueBooleanOtherSystems("select * from Customers where IDNumber = '" & row.Cells(4).Value & "'", con) = True Then
                    If row.Cells(47).Value = "" Then
                        row.Cells(47).Value = "Customer of the ID (" & row.Cells(4).Value & ") Exist In EDM"
                    Else
                        row.Cells(47).Value = row.Cells(47).Value & "/ Customer of the ID (" & row.Cells(4).Value & ") Exist In EDM"
                    End If
                    ''Button2.Enabled = False
                End If

                '###################### Region
                ''EDM
                If checkItem("select * from Region where RegionID = '" & row.Cells(13).Value & "'", con) = False Then
                    If row.Cells(47).Value = "" Then
                        row.Cells(47).Value = "Region of the ID (" & row.Cells(13).Value & ") Does Not Exist In EDM"
                    Else
                        row.Cells(47).Value = row.Cells(47).Value & "/ Region of the ID (" & row.Cells(13).Value & ") Does Not Exist IN EDM"
                    End If
                    Button2.Enabled = False
                End If

                ''RMO
                If checkItem("select * from Region where RegionID = '" & getStringValue("select RMORegion from Region where RegionID = '" & row.Cells(13).Value & "'", con) & "'", conRMOrange) = False Then
                    If row.Cells(47).Value = "" Then
                        row.Cells(47).Value = "Region of the ID (" & getStringValue("select RMORegion from Region where RegionID = '" & row.Cells(13).Value & "'", con) & ") Does Not Exist In RM Orange"
                    Else
                        row.Cells(47).Value = row.Cells(47).Value & "/ Region of the ID (" & getStringValue("select RMORegion from Region where RegionID = '" & row.Cells(13).Value & "'", con) & ") Does Not Exist IN RM Orange"
                    End If
                    Button2.Enabled = False
                End If

                ''RMB
                If checkItem("select * from Region where RegionID = '" & getStringValue("select RMBRegion from Region where RegionID = '" & row.Cells(13).Value & "'", con) & "'", conRMBeMobile) = False Then
                    If row.Cells(47).Value = "" Then
                        row.Cells(47).Value = "Region of the ID (" & getStringValue("select RMBRegion from Region where RegionID = '" & row.Cells(13).Value & "'", con) & ") Does Not Exist In RM Bemobile"
                    Else
                        row.Cells(47).Value = row.Cells(47).Value & "/ Region of the ID (" & getStringValue("select RMBRegion from Region where RegionID = '" & row.Cells(13).Value & "'", con) & ") Does Not Exist IN RM Bemobile"
                    End If
                    Button2.Enabled = False
                End If

                ''FPM
                If checkItem("select * from Regions where RegionID = '" & getStringValue("select FPMRegion from Region where RegionID = '" & row.Cells(13).Value & "'", con) & "'", conFPM) = False Then
                    If row.Cells(47).Value = "" Then
                        row.Cells(47).Value = "Region of the ID (" & getStringValue("select FPMRegion from Region where RegionID = '" & row.Cells(13).Value & "'", con) & ") Does Not Exist In FPM"
                    Else
                        row.Cells(47).Value = row.Cells(47).Value & "/ Region of the ID (" & getStringValue("select FPMRegion from Region where RegionID = '" & row.Cells(13).Value & "'", con) & ") Does Not Exist IN FPM"
                    End If
                    Button2.Enabled = False
                End If

                '###################### Village
                ''EDM
                If checkItem("select * from Location where LocationID = '" & row.Cells(14).Value & "'", con) = False Then
                    If row.Cells(47).Value = "" Then
                        row.Cells(47).Value = "Location of the ID (" & row.Cells(14).Value & ") Does Not Exist In EDM"
                    Else
                        row.Cells(47).Value = row.Cells(47).Value & "/ Location of the ID (" & row.Cells(14).Value & ") Does Not Exist IN EDM"
                    End If
                    Button2.Enabled = False
                End If

                ''RM Orange
                If checkItem("select * from Location where LocationID = '" & getStringValue("select RMOLocation from Location where LocationID = '" & row.Cells(14).Value & "'", con) & "'", conRMOrange) = False Then
                    If row.Cells(47).Value = "" Then
                        row.Cells(47).Value = "Location of the ID (" & getStringValue("select RMOLocation from Location where LocationID = '" & row.Cells(14).Value & "'", con) & ") Does Not Exist In RM Orange"
                    Else
                        row.Cells(47).Value = row.Cells(47).Value & "/ Location of the ID (" & getStringValue("select RMOLocation from Location where LocationID = '" & row.Cells(14).Value & "'", con) & ") Does Not Exist IN RM Orange"
                    End If
                    Button2.Enabled = False
                End If

                ''RM Bemobile
                If checkItem("select * from Location where LocationID = '" & getStringValue("select RMBLocation from Location where LocationID = '" & row.Cells(14).Value & "'", con) & "'", conRMBeMobile) = False Then
                    If row.Cells(47).Value = "" Then
                        row.Cells(47).Value = "Location of the ID (" & getStringValue("select RMBLocation from Location where LocationID = '" & row.Cells(14).Value & "'", con) & ") Does Not Exist In RM BeMobile"
                    Else
                        row.Cells(47).Value = row.Cells(47).Value & "/ Location of the ID (" & getStringValue("select RMBLocation from Location where LocationID = '" & row.Cells(14).Value & "'", con) & ") Does Not Exist IN RM BeMobile"
                    End If
                    Button2.Enabled = False
                End If

                ''FPM
                If checkItem("select * from Area where AreaCode = '" & getStringValue("select FPMLocation from Location where LocationID = '" & row.Cells(14).Value & "'", con) & "'", conFPM) = False Then
                    If row.Cells(47).Value = "" Then
                        row.Cells(47).Value = "Location of the ID (" & getStringValue("select FPMLocation from Location where LocationID = '" & row.Cells(14).Value & "'", con) & ") Does Not Exist In FPM"
                    Else
                        row.Cells(47).Value = row.Cells(47).Value & "/ Location of the ID (" & getStringValue("select FPMLocation from Location where LocationID = '" & row.Cells(14).Value & "'", con) & ") Does Not Exist IN FPM"
                    End If
                    Button2.Enabled = False
                End If

                '###################### School
                ''EDM
                If checkItem("select * from School where SchoolID = '" & row.Cells(19).Value & "'", con) = False Then
                    If row.Cells(47).Value = "" Then
                        row.Cells(47).Value = "School of the ID (" & row.Cells(19).Value & ") Does Not Exist In EDM"
                    Else
                        row.Cells(47).Value = row.Cells(47).Value & "/ School of the ID (" & row.Cells(19).Value & ") Does Not Exist IN EDM"
                    End If
                    Button2.Enabled = False
                End If

                ''FPM
                If checkItem("select * from Schools where SchoolID = '" & getStringValue("select FPMSchool from School where SchoolID = '" & row.Cells(19).Value & "'", con) & "'", conFPM) = False Then
                    If row.Cells(47).Value = "" Then
                        row.Cells(47).Value = "School of the ID (" & getStringValue("select FPMSchool from School where SchoolID = '" & row.Cells(19).Value & "'", con) & ") Does Not Exist In FPM"
                    Else
                        row.Cells(47).Value = row.Cells(47).Value & "/ School of the ID (" & getStringValue("select FPMSchool from School where SchoolID = '" & row.Cells(19).Value & "'", con) & ") Does Not Exist IN FPM"
                    End If
                    Button2.Enabled = False
                End If

                '###################### Employer
                ''EDM
                If checkItem("select * from Employers where AutoID = '" & row.Cells(20).Value & "'", con) = False Then
                    If row.Cells(47).Value = "" Then
                        row.Cells(47).Value = "Employer of the ID (" & row.Cells(20).Value & ") Does Not Exist In EDM"
                    Else
                        row.Cells(47).Value = row.Cells(47).Value & "/ Employer of the ID (" & row.Cells(20).Value & ") Does Not Exist IN EDM"
                    End If
                    Button2.Enabled = False
                End If

                ''RMO
                If checkItem("select * from Employer where Code = '" & getStringValue("select RMOEmployer from Employers where AutoID = '" & row.Cells(20).Value & "'", con) & "'", conRMOrange) = False Then
                    If row.Cells(47).Value = "" Then
                        row.Cells(47).Value = "Employer of the ID (" & getStringValue("select RMOEmployer from Employers where AutoID = '" & row.Cells(20).Value & "'", con) & ") Does Not Exist In RM Orange"
                    Else
                        row.Cells(47).Value = row.Cells(47).Value & "/ Employer of the ID (" & getStringValue("select RMOEmployer from Employers where AutoID = '" & row.Cells(20).Value & "'", con) & ") Does Not Exist IN RM Orange"
                    End If
                    Button2.Enabled = False
                End If

                ''RMB
                If checkItem("select * from Employer where Code = '" & getStringValue("select RMBEmployer from Employers where AutoID = '" & row.Cells(20).Value & "'", con) & "'", conRMBeMobile) = False Then
                    If row.Cells(47).Value = "" Then
                        row.Cells(47).Value = "Employer of the ID (" & getStringValue("select RMBEmployer from Employers where AutoID = '" & row.Cells(20).Value & "'", con) & ") Does Not Exist In RM Bemobile"
                    Else
                        row.Cells(47).Value = row.Cells(47).Value & "/ Employer of the ID (" & getStringValue("select RMBEmployer from Employers where AutoID = '" & row.Cells(20).Value & "'", con) & ") Does Not Exist IN RM Bemobile"
                    End If
                    Button2.Enabled = False
                End If

                ''MLMA
                If checkItem("select * from Employer where AutoID = '" & getStringValue("select MLMAEmployer from Employers where AutoID = '" & row.Cells(20).Value & "'", con) & "'", conMLMAdvances) = False Then
                    If row.Cells(47).Value = "" Then
                        row.Cells(47).Value = "Employer of the ID (" & getStringValue("select MLMAEmployer from Employers where AutoID = '" & row.Cells(20).Value & "'", con) & ") Does Not Exist In MLM Advances"
                    Else
                        row.Cells(47).Value = row.Cells(47).Value & "/ Employer of the ID (" & getStringValue("select MLMAEmployer from Employers where AutoID = '" & row.Cells(20).Value & "'", con) & ") Does Not Exist IN MLM Advances"
                    End If
                    Button2.Enabled = False
                End If

                ''MLME
                If checkItem("select * from Employer where AutoID = '" & getStringValue("select MLMEEmployer from Employers where AutoID = '" & row.Cells(20).Value & "'", con) & "'", conMLMEducational) = False Then
                    If row.Cells(47).Value = "" Then
                        row.Cells(47).Value = "Employer of the ID (" & getStringValue("select MLMEEmployer from Employers where AutoID = '" & row.Cells(20).Value & "'", con) & ") Does Not Exist In MLM Educational"
                    Else
                        row.Cells(47).Value = row.Cells(47).Value & "/ Employer of the ID (" & getStringValue("select MLMEEmployer from Employers where AutoID = '" & row.Cells(20).Value & "'", con) & ") Does Not Exist IN MLM Educational"
                    End If
                    Button2.Enabled = False
                End If

                '###################### Department
                ''EDM
                If checkItem("select * from Department where DepartmentID = '" & row.Cells(21).Value & "'", con) = False Then
                    If row.Cells(47).Value = "" Then
                        row.Cells(47).Value = "Department of the ID (" & row.Cells(21).Value & ") Does Not Exist In EDM"
                    Else
                        row.Cells(47).Value = row.Cells(47).Value & "/ Department of the ID (" & row.Cells(21).Value & ") Does Not Exist IN EDM"
                    End If
                    Button2.Enabled = False
                End If

                ''RMO
                If checkItem("select * from Department where DepartmentCode = '" & getStringValue("select RMODepartment from Department where DepartmentID = '" & row.Cells(21).Value & "'", con) & "'", conRMOrange) = False Then
                    If row.Cells(47).Value = "" Then
                        row.Cells(47).Value = "Department of the ID (" & getStringValue("select RMODepartment from Department where DepartmentID = '" & row.Cells(21).Value & "'", con) & ") Does Not Exist In RM Orange"
                    Else
                        row.Cells(47).Value = row.Cells(47).Value & "/ Department of the ID (" & getStringValue("select RMODepartment from Department where DepartmentID = '" & row.Cells(21).Value & "'", con) & ") Does Not Exist IN RM Orange"
                    End If
                    Button2.Enabled = False
                End If

                ''RMB
                If checkItem("select * from Department where DepartmentCode = '" & getStringValue("select RMBDepartment from Department where DepartmentID = '" & row.Cells(21).Value & "'", con) & "'", conRMBeMobile) = False Then
                    If row.Cells(47).Value = "" Then
                        row.Cells(47).Value = "Department of the ID (" & getStringValue("select RMBDepartment from Department where DepartmentID = '" & row.Cells(21).Value & "'", con) & ") Does Not Exist In RM Bemobile"
                    Else
                        row.Cells(47).Value = row.Cells(47).Value & "/ Department of the ID (" & getStringValue("select RMBDepartment from Department where DepartmentID = '" & row.Cells(21).Value & "'", con) & ") Does Not Exist IN RM Bemobile"
                    End If
                    Button2.Enabled = False
                End If

                ''FPM
                If checkItem("select * from Departments where DepID = '" & getStringValue("select FPMDepartment from Department where DepartmentID = '" & row.Cells(21).Value & "'", con) & "'", conFPM) = False Then
                    If row.Cells(47).Value = "" Then
                        row.Cells(47).Value = "Department of the ID (" & getStringValue("select FPMDepartment from Department where DepartmentID = '" & row.Cells(21).Value & "'", con) & ") Does Not Exist In FPM"
                    Else
                        row.Cells(47).Value = row.Cells(47).Value & "/ Department of the ID (" & getStringValue("select FPMDepartment from Department where DepartmentID = '" & row.Cells(21).Value & "'", con) & ") Does Not Exist IN FPM"
                    End If
                    Button2.Enabled = False
                End If

                ''MLMA
                If checkItem("select * from EmployerDepartment where DepartmentID = '" & getStringValue("select MLMADepartment from Department where DepartmentID = '" & row.Cells(21).Value & "'", con) & "'", conMLMAdvances) = False Then
                    If row.Cells(47).Value = "" Then
                        row.Cells(47).Value = "Department of the ID (" & getStringValue("select MLMADepartment from Department where DepartmentID = '" & row.Cells(21).Value & "'", con) & ") Does Not Exist In MLM Advances"
                    Else
                        row.Cells(47).Value = row.Cells(47).Value & "/ Department of the ID (" & getStringValue("select MLMADepartment from Department where DepartmentID = '" & row.Cells(21).Value & "'", con) & ") Does Not Exist IN MLM Advances"
                    End If
                    Button2.Enabled = False
                End If

                ''MLMB
                If checkItem("select * from EmployerDepartment where DepartmentID = '" & getStringValue("select MLMEDepartment from Department where DepartmentID = '" & row.Cells(21).Value & "'", con) & "'", conMLMEducational) = False Then
                    If row.Cells(47).Value = "" Then
                        row.Cells(47).Value = "Department of the ID (" & getStringValue("select MLMEDepartment from Department where DepartmentID = '" & row.Cells(21).Value & "'", con) & ") Does Not Exist In MLM Educational"
                    Else
                        row.Cells(47).Value = row.Cells(47).Value & "/ Department of the ID (" & getStringValue("select MLMEDepartment from Department where DepartmentID = '" & row.Cells(21).Value & "'", con) & ") Does Not Exist IN MLM Educational"
                    End If
                    Button2.Enabled = False
                End If

                '###################### Group
                ''EDM
                If checkItem("select * from Groups where GroupID = '" & row.Cells(45).Value & "'", con) = False Then
                    If row.Cells(47).Value = "" Then
                        row.Cells(47).Value = "Group of the ID (" & row.Cells(45).Value & ") Does Not Exist In EDM"
                    Else
                        row.Cells(47).Value = row.Cells(47).Value & "/ Group of the ID (" & row.Cells(45).Value & ") Does Not Exist IN EDM"
                    End If
                    Button2.Enabled = False
                End If

                ''FPM
                If checkItem("select * from Groups where GroupID = '" & getStringValue("select FPMGroup from Groups where GroupID = '" & row.Cells(45).Value & "'", con) & "'", conFPM) = False Then
                    If row.Cells(47).Value = "" Then
                        row.Cells(47).Value = "Group of the ID (" & getStringValue("select FPMGroup from Groups where GroupID = '" & row.Cells(45).Value & "'", con) & ") Does Not Exist In FPM"
                    Else
                        row.Cells(47).Value = row.Cells(47).Value & "/ Group of the ID (" & getStringValue("select FPMGroup from Groups where GroupID = '" & row.Cells(45).Value & "'", con) & ") Does Not Exist IN FPM"
                    End If
                    Button2.Enabled = False
                End If

                ''MLMA
                If checkItem("select * from Groups where GroupID = '" & getStringValue("select MLMAGroup from Groups where GroupID = '" & row.Cells(45).Value & "'", con) & "'", conMLMAdvances) = False Then
                    If row.Cells(47).Value = "" Then
                        row.Cells(47).Value = "Group of the ID (" & getStringValue("select MLMAGroup from Groups where GroupID = '" & row.Cells(45).Value & "'", con) & ") Does Not Exist In MLM Advances"
                    Else
                        row.Cells(47).Value = row.Cells(47).Value & "/ Group of the ID (" & getStringValue("select MLMAGroup from Groups where GroupID = '" & row.Cells(45).Value & "'", con) & ") Does Not Exist IN MLM Advances"
                    End If
                    Button2.Enabled = False
                End If

                ''MLME
                If checkItem("select * from Groups where GroupID = '" & getStringValue("select MLMEGroup from Groups where GroupID = '" & row.Cells(45).Value & "'", con) & "'", conMLMEducational) = False Then
                    If row.Cells(47).Value = "" Then
                        row.Cells(47).Value = "Group of the ID (" & getStringValue("select MLMEGroup from Groups where GroupID = '" & row.Cells(45).Value & "'", con) & ") Does Not Exist In MLM Educational"
                    Else
                        row.Cells(47).Value = row.Cells(47).Value & "/ Group of the ID (" & getStringValue("select MLMEGroup from Groups where GroupID = '" & row.Cells(45).Value & "'", con) & ") Does Not Exist IN MLM Educational"
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

    Private Sub cbrmo_CheckedChanged(sender As Object, e As EventArgs) Handles cbrmo.CheckedChanged
        If cbrmo.Checked = True Then
            cbfpm.Checked = False
            cbmlma.Checked = False
            cbmlme.Checked = False
            cbrmo.Checked = True
            cbrmb.Checked = False
        ElseIf cbrmo.Checked = False Then
            cbfpm.Checked = False
            cbmlma.Checked = False
            cbmlme.Checked = False
            cbrmo.Checked = False
            cbrmb.Checked = False
        End If
    End Sub

    Private Sub cbrmb_CheckedChanged(sender As Object, e As EventArgs) Handles cbrmb.CheckedChanged
        If cbrmb.Checked = True Then
            cbfpm.Checked = False
            cbmlma.Checked = False
            cbmlme.Checked = False
            cbrmo.Checked = False
            cbrmb.Checked = True
        ElseIf cbrmb.Checked = False Then
            cbfpm.Checked = False
            cbmlma.Checked = False
            cbmlme.Checked = False
            cbrmo.Checked = False
            cbrmb.Checked = False
        End If
    End Sub

    Private Sub cbmlma_CheckedChanged(sender As Object, e As EventArgs) Handles cbmlma.CheckedChanged
        If cbmlma.Checked = True Then
            cbfpm.Checked = False
            cbmlma.Checked = True
            cbmlme.Checked = False
            cbrmo.Checked = False
            cbrmb.Checked = False
        ElseIf cbmlma.Checked = False Then
            cbfpm.Checked = False
            cbmlma.Checked = False
            cbmlme.Checked = False
            cbrmo.Checked = False
            cbrmb.Checked = False
        End If
    End Sub

    Private Sub cbmlme_CheckedChanged(sender As Object, e As EventArgs) Handles cbmlme.CheckedChanged
        If cbmlme.Checked = True Then
            cbfpm.Checked = False
            cbmlma.Checked = False
            cbmlme.Checked = True
            cbrmo.Checked = False
            cbrmb.Checked = False
        ElseIf cbmlme.Checked = False Then
            cbfpm.Checked = False
            cbmlma.Checked = False
            cbmlme.Checked = False
            cbrmo.Checked = False
            cbrmb.Checked = False
        End If
    End Sub

    Private Sub cbfpm_CheckedChanged(sender As Object, e As EventArgs) Handles cbfpm.CheckedChanged
        If cbfpm.Checked = True Then
            cbfpm.Checked = True
            cbmlma.Checked = False
            cbmlme.Checked = False
            cbrmo.Checked = False
            cbrmb.Checked = False
        ElseIf cbfpm.Checked = False Then
            cbfpm.Checked = False
            cbmlma.Checked = False
            cbmlme.Checked = False
            cbrmo.Checked = False
            cbrmb.Checked = False
        End If
    End Sub

    Private Sub UploadCustomers_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If getValueBooleanOtherSystems("select * from CompanyDetails where Integrate = 'Yes'", con) = True Then
            Me.Enabled = True
        Else
            MessageBox.Show("Cannot use the form until the Integration has been enabled", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Close()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If decision("Are you to proceed ?") = DialogResult.Yes Then
            For Each row As DataGridViewRow In DataGridView1.Rows
                If getValueBooleanOtherSystems("select * from Customers where IDNumber = '" & row.Cells(4).Value & "'", con) = True Then

                Else
                    saveCustomer(row.Cells(0).Value, row.Cells(1).Value, row.Cells(2).Value, Date.Parse(row.Cells(3).Value), row.Cells(4).Value, row.Cells(5).Value, row.Cells(6).Value, Date.Parse(row.Cells(7).Value), row.Cells(8).Value, row.Cells(9).Value, row.Cells(10).Value, row.Cells(11).Value, row.Cells(12).Value, row.Cells(13).Value, row.Cells(14).Value, row.Cells(15).Value, row.Cells(16).Value, row.Cells(17).Value, row.Cells(18).Value, row.Cells(19).Value, row.Cells(20).Value, row.Cells(21).Value, row.Cells(22).Value, row.Cells(23).Value, row.Cells(24).Value, row.Cells(25).Value, row.Cells(26).Value, row.Cells(27).Value, Date.Parse(row.Cells(28).Value), row.Cells(29).Value, row.Cells(30).Value, row.Cells(31).Value, row.Cells(32).Value, row.Cells(33).Value, row.Cells(34).Value, row.Cells(35).Value, row.Cells(36).Value, row.Cells(37).Value, row.Cells(38).Value, row.Cells(39).Value, row.Cells(40).Value, Decimal.Parse(row.Cells(41).Value), row.Cells(42).Value, row.Cells(43).Value, row.Cells(44).Value, branchid, Date.Now, "Active", Date.Now, "", Date.Now, row.Cells(45).Value, row.Cells(46).Value)
                End If

                If cbfpm.Checked = True Then
                    If getValueBooleanOtherSystems("select * from CustomerDetails where IDNumber = '" & row.Cells(4).Value & "'", conFPM) = False Then

                        FPMGender = "select GenderName from Gender where GenderID = '" & row.Cells(8).Value & "'"
                        FPMRegionID = getValueStringOtherSystems("select FPMRegion from Region where RegionID = '" & row.Cells(13).Value & "'", con)
                        FPMLocation = getValueStringOtherSystems("select FPMLocation from Location where LocationID = '" & row.Cells(14).Value & "'", con)
                        FPMSchool = getValueStringOtherSystems("select FPMSchool from School where SchoolID = '" & row.Cells(19).Value & "'", con)
                        FPMDepartment = getValueStringOtherSystems("select FPMDepartment from Department where DepartmentID = '" & row.Cells(21).Value & "'", con)
                        FPMGroup = getValueStringOtherSystems("select FPMGroup from Groups where GroupID = '" & row.Cells(45).Value & "'", con)
                        Nationality = getValueStringOtherSystems("select CountryName from Country where CountryID = '" & row.Cells(1).Value & "'", con)

                        SaveFPMCustomerDetails(row.Cells(0).Value, row.Cells(4).Value, row.Cells(5).Value, row.Cells(6).Value, Date.Parse(row.Cells(7).Value), FPMGender, row.Cells(12).Value, row.Cells(11).Value, row.Cells(9).Value, row.Cells(10).Value, FPMLocation, FPMGroup, row.Cells(22).Value, "Active", Date.Now, "", Date.Now, Nationality, row.Cells(18).Value, FPMSchool, FPMDepartment, FPMRegionID, conFPM)

                    End If
                ElseIf cbrmo.Checked = True Then
                    If getValueBooleanOtherSystems("select * from CustomerMaster where Account = '" & row.Cells(4).Value & "'", conRMOrange) = False Then

                        RMOGender = "select GenderName from Gender where GenderID = '" & row.Cells(8).Value & "'"
                        RMORegionID = getValueStringOtherSystems("select RMORegion from Region where RegionID = '" & row.Cells(13).Value & "'", con)
                        RMOLocation = getValueStringOtherSystems("select RMOLocation from Location where LocationID = '" & row.Cells(14).Value & "'", con)
                        RMOEmployer = getValueStringOtherSystems("select RMOEmployer from Employers where AutoID = '" & row.Cells(20).Value & "'", con)
                        RMODepartment = getValueStringOtherSystems("select RMODepartment from Department where DepartmentID = '" & row.Cells(21).Value & "'", con)
                        Nationality = getValueStringOtherSystems("select CountryName from Country where CountryID = '" & row.Cells(1).Value & "'", con)

                        saveRMCustomerDetails(Nationality, row.Cells(4).Value, row.Cells(6).Value, row.Cells(5).Value, RMOGender, RMOEmployer, RMODepartment, row.Cells(29).Value, row.Cells(30).Value, row.Cells(31).Value, row.Cells(32).Value, row.Cells(22).Value, RMOLocation, row.Cells(11).Value, row.Cells(18).Value, row.Cells(18).Value, row.Cells(33).Value, row.Cells(34).Value, row.Cells(35).Value, row.Cells(9).Value, RMORegionID, RMOLocation, row.Cells(15).Value, row.Cells(16).Value, row.Cells(17).Value, row.Cells(12).Value, row.Cells(37).Value, row.Cells(38).Value, row.Cells(39).Value, Date.Parse(row.Cells(7).Value), row.Cells(23).Value, row.Cells(24).Value, row.Cells(25).Value, row.Cells(10).Value, row.Cells(36).Value, row.Cells(40).Value, 1, branchid, Date.Parse(row.Cells(3).Value), conRMOrange)
                    End If
                ElseIf cbrmb.Checked = True Then
                    If getValueBooleanOtherSystems("select * from CustomerMaster where Account = '" & row.Cells(4).Value & "'", conRMBeMobile) = False Then

                        RMBGender = "select GenderName from Gender where GenderID = '" & row.Cells(8).Value & "'"
                        RMBRegionID = getValueStringOtherSystems("select RMORegion from Region where RegionID = '" & row.Cells(13).Value & "'", con)
                        RMBLocation = getValueStringOtherSystems("select RMOLocation from Location where LocationID = '" & row.Cells(14).Value & "'", con)
                        RMBEmployer = getValueStringOtherSystems("select RMOEmployer from Employers where AutoID = '" & row.Cells(20).Value & "'", con)
                        RMBDepartment = getValueStringOtherSystems("select RMODepartment from Department where DepartmentID = '" & row.Cells(21).Value & "'", con)
                        Nationality = getValueStringOtherSystems("select CountryName from Country where CountryID = '" & row.Cells(1).Value & "'", con)

                        saveRMCustomerDetails(Nationality, row.Cells(4).Value, row.Cells(6).Value, row.Cells(5).Value, RMBGender, RMBEmployer, RMBDepartment, row.Cells(29).Value, row.Cells(30).Value, row.Cells(31).Value, row.Cells(32).Value, row.Cells(22).Value, RMBLocation, row.Cells(11).Value, row.Cells(18).Value, row.Cells(18).Value, row.Cells(33).Value, row.Cells(34).Value, row.Cells(35).Value, row.Cells(9).Value, RMBRegionID, RMBLocation, row.Cells(15).Value, row.Cells(16).Value, row.Cells(17).Value, row.Cells(12).Value, row.Cells(37).Value, row.Cells(38).Value, row.Cells(39).Value, Date.Parse(row.Cells(7).Value), row.Cells(23).Value, row.Cells(24).Value, row.Cells(25).Value, row.Cells(10).Value, row.Cells(36).Value, row.Cells(40).Value, 1, branchid, Date.Parse(row.Cells(3).Value), conRMBeMobile)
                    End If
                ElseIf cbmlma.Checked = True Then
                    If getValueBooleanOtherSystems("select * from CustomerDetails where ID_Number = '" & row.Cells(4).Value & "'", conMLMAdvances) = False Then

                        MLMAGender = "select GenderName from Gender where GenderID = '" & row.Cells(8).Value & "'"
                        MLMALocation = getValueStringOtherSystems("select LocationName from Location where LocationID = '" & row.Cells(14).Value & "'", con)
                        MLMAEmployer = getValueStringOtherSystems("select MLMAEmployer from Employers where AutoID = '" & row.Cells(20).Value & "'", con)
                        MLMAEmployerName = getValueStringOtherSystems("select Employer from Employer where AutoID = '" & MLMAEmployer & "'", conMLMAdvances)
                        MLMADepartment = getValueStringOtherSystems("select MLMADepartment from Department where DepartmentID = '" & row.Cells(21).Value & "'", con)
                        MLMAEmpType = getValueStringOtherSystems("select MLMAEmpType from EmploymentType where EmploymentTypeID = '" & row.Cells(26).Value & "'", con)
                        MLMAGroup = getValueStringOtherSystems("select MLMAGroup from Groups where GroupID = '" & row.Cells(45).Value & "'", con)
                        MLMABankName = row.Cells(42).Value
                        MLMABankBranchName = row.Cells(43).Value

                        saveMLMCustomerDetails(row.Cells(5).Value, row.Cells(6).Value, row.Cells(4).Value, Date.Parse(row.Cells(7).Value), row.Cells(9).Value, row.Cells(12).Value, MLMAEmployerName, row.Cells(29).Value, row.Cells(11).Value, row.Cells(18).Value, MLMALocation, row.Cells(16).Value, row.Cells(17).Value, row.Cells(33).Value, row.Cells(34).Value, row.Cells(35).Value, row.Cells(36).Value, row.Cells(37).Value, row.Cells(38).Value, row.Cells(39).Value, row.Cells(40).Value, MLMAGender, row.Cells(27).Value, MLMAEmpType, Date.Parse(row.Cells(28).Value), Decimal.Parse(row.Cells(41).Value), MLMABankName, MLMABankBranchName, row.Cells(44).Value, branchid, MLMAEmployer, MLMADepartment, MLMAGroup, conMLMAdvances)
                    End If
                ElseIf cbmlme.Checked = True Then
                    If getValueBooleanOtherSystems("select * from CustomerDetails where ID_Number = '" & row.Cells(4).Value & "'", conMLMEducational) = False Then

                        MLMEGender = "select GenderName from Gender where GenderID = '" & row.Cells(8).Value & "'"
                        MLMELocation = getValueStringOtherSystems("select LocationName from Location where LocationID = '" & row.Cells(14).Value & "'", con)
                        MLMEEmployer = getValueStringOtherSystems("select MLMAEmployer from Employers where AutoID = '" & row.Cells(20).Value & "'", con)
                        MLMEEmployerName = getValueStringOtherSystems("select Employer from Employer where AutoID = '" & MLMEEmployer & "'", conMLMEducational)
                        MLMEDepartment = getValueStringOtherSystems("select MLMADepartment from Department where DepartmentID = '" & row.Cells(21).Value & "'", con)
                        MLMEEmpType = getValueStringOtherSystems("select MLMAEmpType from EmploymentType where EmploymentTypeID = '" & row.Cells(26).Value & "'", con)
                        MLMEGroup = getValueStringOtherSystems("select MLMAGroup from Groups where GroupID = '" & row.Cells(45).Value & "'", con)
                        MLMEBankName = row.Cells(42).Value
                        MLMEBankBranchName = row.Cells(43).Value

                        saveMLMCustomerDetails(row.Cells(5).Value, row.Cells(6).Value, row.Cells(4).Value, Date.Parse(row.Cells(7).Value), row.Cells(9).Value, row.Cells(12).Value, MLMEEmployerName, row.Cells(29).Value, row.Cells(11).Value, row.Cells(18).Value, MLMELocation, row.Cells(16).Value, row.Cells(17).Value, row.Cells(33).Value, row.Cells(34).Value, row.Cells(35).Value, row.Cells(36).Value, row.Cells(37).Value, row.Cells(38).Value, row.Cells(39).Value, row.Cells(40).Value, MLMEGender, row.Cells(27).Value, MLMEEmpType, Date.Parse(row.Cells(28).Value), Decimal.Parse(row.Cells(41).Value), MLMEBankName, MLMEBankBranchName, row.Cells(44).Value, branchid, MLMEEmployer, MLMEDepartment, MLMEGroup, conMLMEducational)
                    End If
                End If
            Next
            MessageBox.Show("Customer details Uploaded", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Close()
        End If

    End Sub

    Private Sub loadData(ByVal str As String)
        DataGridView2.DataSource = Nothing
        Dim qry As String = str

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
                DataGridView2.DataSource = bs
                BindingNavigator1.BindingSource = bs
                DataGridView2.Columns(DataGridView2.ColumnCount - 1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            Else
                DataGridView2.DataSource = Nothing
                BindingNavigator1.BindingSource = Nothing
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "loadData()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadData()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

End Class