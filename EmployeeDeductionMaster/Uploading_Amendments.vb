Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc

Public Class Uploading_Amendments

    Dim totalamendments, totalterminations As Decimal
    Dim index As Integer = 0
    Dim productid As Integer = 0
    
    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        ExportToExcel(DataGridView1, "Format")
        DataGridView1.AllowUserToAddRows = False
        LinkLabel1.Text = "Export"
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        DataGridView1.AllowUserToAddRows = False
        Button1.Enabled = False

        If decision("Do you want to load data ?") = DialogResult.Yes Then
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

                For Each row As DataGridViewRow In DataGridView1.Rows
                    If row.Cells(2).Value = "Amendment" Then
                        If Decimal.Parse(row.Cells(1).Value.ToString) > 0 Then

                        Else
                            If row.Cells(3).Value = "" Then
                                row.Cells(3).Value = "Amount not correct"
                            Else
                                row.Cells(3).Value = row.Cells(3).Value & "/ Amount not correct"
                            End If
                            Button2.Enabled = False
                            Label2.Enabled = True
                        End If
                    ElseIf row.Cells(2).Value = "Termination" Then
                        If Decimal.Parse(row.Cells(1).Value.ToString) = 0 Then

                        Else
                            If row.Cells(3).Value = "" Then
                                row.Cells(3).Value = "Amount should be 0"
                            Else
                                row.Cells(3).Value = row.Cells(3).Value & "/ Amount should be 0"
                            End If
                            Button2.Enabled = False
                            Label2.Enabled = True
                        End If
                    Else
                        If row.Cells(3).Value = "" Then
                            row.Cells(3).Value = "Type is wrong"
                        Else
                            row.Cells(3).Value = row.Cells(3).Value & "/ Type is wrong"
                        End If
                        Button2.Enabled = False
                        Label2.Enabled = True
                    End If

                    If checkCustomerDetails(row.Cells(0).Value.ToString, employerid) = True Then

                    Else
                        If row.Cells(3).Value = "" Then
                            row.Cells(3).Value = "Customer Does Not Exists"
                        Else
                            row.Cells(3).Value = row.Cells(3).Value & "/ Customer Does Not Exists"
                        End If
                        Button2.Enabled = False
                        Label2.Enabled = True
                    End If

                    If row.Cells(2).Value = "Amendment" Then
                        If checkAmendment(row.Cells(0).Value.ToString, DateTimePicker1.Value, productid, employerid) = True Then
                            If row.Cells(3).Value = "" Then
                                row.Cells(3).Value = "Amendment available. Use Update Amendment Function (" & getAmendmentAmount(row.Cells(0).Value.ToString, DateTimePicker1.Value, productid, employerid) & ")"
                            Else
                                row.Cells(3).Value = row.Cells(3).Value & "/ Amendment available. Use Update Amendment Function (" & getAmendmentAmount(row.Cells(0).Value.ToString, DateTimePicker1.Value, productid, employerid) & ")"
                            End If
                            Button2.Enabled = False
                            Label2.Enabled = True
                        Else
                            If checkTermination(row.Cells(0).Value.ToString, DateTimePicker1.Value, productid, employerid) = True Then
                                If row.Cells(3).Value = "" Then
                                    row.Cells(3).Value = "Termination available. Use Update Termination Function"
                                Else
                                    row.Cells(3).Value = row.Cells(3).Value & "/ Termination available. Use Update Termination Function"
                                End If
                                Button2.Enabled = False
                                Label2.Enabled = True
                            End If
                        End If
                    ElseIf row.Cells(2).Value = "Termination" Then
                        If checkTermination(row.Cells(0).Value.ToString, DateTimePicker1.Value, productid, employerid) = True Then
                            If row.Cells(3).Value = "" Then
                                row.Cells(3).Value = "Termination available. Use Update Termination Function"
                            Else
                                row.Cells(3).Value = row.Cells(3).Value & "/ Termination available. Use Update Termination Function"
                            End If
                            Button2.Enabled = False
                            Label2.Enabled = True
                        Else
                            If checkAmendment(row.Cells(0).Value.ToString, DateTimePicker1.Value, productid, employerid) = True Then
                                If row.Cells(3).Value = "" Then
                                    row.Cells(3).Value = "Amendment available. Use Update Amendment Function (" & getAmendmentAmount(row.Cells(0).Value.ToString, DateTimePicker1.Value, productid, employerid) & ")"
                                Else
                                    row.Cells(3).Value = row.Cells(3).Value & "/ Amendment available. Use Update Amendment Function (" & getAmendmentAmount(row.Cells(0).Value.ToString, DateTimePicker1.Value, productid, employerid) & ")"
                                End If
                                Button2.Enabled = False
                                Label2.Enabled = True
                            End If
                        End If
                    End If
                Next

                totalamendments = 0
                totalterminations = 0
                TextBox1.Text = ""
                TextBox2.Text = ""
                For Each row As DataGridViewRow In DataGridView1.Rows
                    If row.Cells(2).Value = "Amendment" Then
                        totalamendments = totalamendments + Decimal.Parse(row.Cells(1).Value.ToString)
                    ElseIf row.Cells(2).Value = "Termination" Then
                        totalterminations = totalterminations + 1
                    End If
                Next
                TextBox1.Text = totalamendments
                TextBox2.Text = totalterminations
            End If
        End If
    End Sub

    Public Sub ExportToExcel(ByVal DataGridView1 As DataGridView, ByVal title As String)
        'verfying the datagridview having data or not
        If ((DataGridView1.Columns.Count = 0) Or (DataGridView1.Rows.Count = 0)) Then
            Exit Sub
        End If

        'Creating dataset to export
        Dim dset As New DataSet
        'add table to dataset
        dset.Tables.Add()
        'add column to that table
        For i As Integer = 0 To DataGridView1.ColumnCount - 1
            dset.Tables(0).Columns.Add(DataGridView1.Columns(i).HeaderText)
        Next
        'add rows to the table
        Dim dr1 As DataRow
        For i As Integer = 0 To DataGridView1.RowCount - 1
            dr1 = dset.Tables(0).NewRow
            For j As Integer = 0 To DataGridView1.Columns.Count - 1
                dr1(j) = DataGridView1.Rows(i).Cells(j).Value
            Next
            dset.Tables(0).Rows.Add(dr1)
        Next

        Dim excel As New Microsoft.Office.Interop.Excel.ApplicationClass
        Dim wBook As Microsoft.Office.Interop.Excel.Workbook
        Dim wSheet As Microsoft.Office.Interop.Excel.Worksheet

        wBook = excel.Workbooks.Add()
        wSheet = wBook.ActiveSheet()

        Dim dt As System.Data.DataTable = dset.Tables(0)
        Dim dc As System.Data.DataColumn
        Dim dr As System.Data.DataRow
        Dim colIndex As Integer = 0
        Dim rowIndex As Integer = 1

        For Each dc In dt.Columns
            colIndex = colIndex + 1
            excel.Cells(2, colIndex) = dc.ColumnName
        Next
        'put excel heading
        excel.Cells(1, 1) = title
        excel.Range("A1:D1").Merge()

        For Each dr In dt.Rows
            rowIndex = rowIndex + 1
            colIndex = 0
            For Each dc In dt.Columns
                colIndex = colIndex + 1
                excel.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName)

            Next
        Next

        wSheet.Columns.AutoFit()
        excel.Visible = True

    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If decision("Are you sure you want to proceed?") = DialogResult.Yes Then
            If DataGridView1.Rows.Count > 0 Then
                ProgressBar1.Maximum = DataGridView1.Rows.Count - 1
                For Each row As DataGridViewRow In DataGridView1.Rows
                    If cbchangelodgingmonth.Checked = True Then
                        If checkAmendment(row.Cells(0).Value.ToString, getLastDateOfMonth(DateTimePicker1.Value), productid, employerid) = True Then
                            updateAmendment(row.Cells(0).Value.ToString, getLastDateOfMonth(DateTimePicker1.Value), productid, Decimal.Parse(row.Cells(1).Value.ToString), employerid)
                        ElseIf checkTermination(row.Cells(0).Value.ToString, getLastDateOfMonth(DateTimePicker1.Value), productid, employerid) = True Then
                            updateTermination(row.Cells(0).Value.ToString, getLastDateOfMonth(DateTimePicker1.Value), productid, Decimal.Parse(row.Cells(1).Value.ToString), employerid)
                        Else
                            saveAmendments(row.Cells(0).Value.ToString, Decimal.Parse(row.Cells(1).Value.ToString), getLastDateOfMonth(DateTimePicker1.Value), Date.Now, productid, employerid)
                        End If
                        ProgressBar1.Value = row.Index
                    Else
                        saveAmendments(row.Cells(0).Value.ToString, Decimal.Parse(row.Cells(1).Value.ToString), getLastDateOfMonth(DateTimePicker1.Value), Date.Now, productid, employerid)
                        ProgressBar1.Value = row.Index
                    End If
                Next
                ProgressBar1.Value = 0
                MessageBox.Show("All amendments has been saved successfully", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
                wrapper.saveAction("Amendments and Terminations uploaded", Date.Now, logedinname, "")

                'Close()
                Dim vrl As New Uploading_Amendments
                vrl.MdiParent = mform
                vrl.Text = vrl.Text & " : Employer Name = " & employername
                vrl.Show()
                vrl.DateTimePicker1.Value = DateTimePicker1.Value
                vrl.loadProducts(employerid)
                Me.Close()
            End If
        Else
            MessageBox.Show("Nothing to save", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        DateTimePicker1.Value = getLastDateOfMonth(DateTimePicker1.Value)
        Button3.Enabled = True
    End Sub

    Private Sub Uploading_Amendments_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DateTimePicker1.Value = getLastDateOfMonth(addMonth(getLastLodgingDate(employerid), 1))
    End Sub

    Private Sub loadProducts(ByVal EmployerID As Integer)
        ''Dim qry As String = "select AutoID,Products, Products as TotalAmendments, Products as Count, Products as TerminationsCount from Products Where Status = 'Active' order by AllocationOrder ASC"
        Dim qry As String = "select EmployerProducts.AutoID,Products, Products as TotalAmendments, Products as Count, Products as TerminationsCount from EmployerProducts inner join Products on EmployerProducts.ProductID = Products.AutoID where EmployerProducts.Status = 'Active' and Products.Status = 'Active' and EmployerID = '" & EmployerID & "' order by OrderNumber ASC"
        DataGridView2.DataSource = Nothing
        BindingNavigator1.BindingSource = Nothing

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)
        Dim dt As DataTable

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim ctr As Integer = 0
            SQLAdapter.Fill(ds, "Defaulters")

            If ds.Tables(0).Rows.Count > 0 Then
                dt = ds.Tables(0)
                Dim bs As New BindingSource
                bs.DataSource = dt.DefaultView
                DataGridView2.DataSource = bs
                BindingNavigator1.BindingSource = bs

                Dim chk1 As New DataGridViewCheckBoxColumn()
                chk1.HeaderText = "Choose?"
                chk1.Name = "confirm"
                If DataGridView2.Columns("confirm") Is Nothing Then
                    DataGridView2.Columns.Add(chk1)
                End If

                For Each row As DataGridViewRow In DataGridView2.Rows
                    row.Cells(2).Value = getProductAmendment(DateTimePicker1.Value, Integer.Parse(row.Cells(0).Value.ToString), EmployerID)
                    row.Cells(3).Value = getProductsCount(DateTimePicker1.Value, Integer.Parse(row.Cells(0).Value.ToString), EmployerID)
                    row.Cells(4).Value = getTerminationsCount(DateTimePicker1.Value, Integer.Parse(row.Cells(0).Value.ToString), EmployerID)
                Next

                DataGridView2.Columns(0).ReadOnly = True
                DataGridView2.Columns(1).ReadOnly = True
                DataGridView2.Columns(2).ReadOnly = True
                DataGridView2.Columns(3).ReadOnly = True

            Else
                DataGridView2.DataSource = Nothing
                BindingNavigator1.BindingSource = Nothing
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "loadProducts()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadProducts()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If cbchangedeductionendndate.Checked = True Then
            loadProducts(employerid)
            Button3.Enabled = False
            DateTimePicker1.Enabled = False
        Else
            If checkExpectedDeductions(DateTimePicker1.Value, employerid) = True Then
                MessageBox.Show("There are expected deductions for this month", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                loadProducts(employerid)
                Button3.Enabled = False
                DateTimePicker1.Enabled = False
            End If
        End If
        
    End Sub

    Private Sub DataGridView2_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick
        If DataGridView2.Rows.Count > 0 Then
            index = DataGridView2.CurrentRow.Index
            productid = Integer.Parse(DataGridView2.Rows(DataGridView2.CurrentRow.Index).Cells(0).Value.ToString)
            Button1.Enabled = True

            For Each row As DataGridViewRow In DataGridView2.Rows
                If row.Cells("confirm").Value = True Then
                    If row.Index = index Then

                    Else
                        row.Cells("confirm").Value = False
                    End If
                End If
            Next
        End If
    End Sub

    Private Sub cbchangelodgingmonth_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbchangelodgingmonth.CheckedChanged
        If cbchangelodgingmonth.Checked = True Then
            Button2.Enabled = True
        Else
            Button2.Enabled = False
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbchangedeductionendndate.CheckedChanged
        If cbchangedeductionendndate.Checked = True Then
            DateTimePicker1.Enabled = True
        Else
            DateTimePicker1.Enabled = False
        End If
    End Sub
End Class