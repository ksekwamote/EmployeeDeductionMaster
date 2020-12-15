Imports DgvFilterPopup
Module Methods

    Public con As New SqlClient.SqlConnection("")
    Public conFPM As New SqlClient.SqlConnection("")
    Public conMLMAdvances As New SqlClient.SqlConnection("")
    Public conMLMEducational As New SqlClient.SqlConnection("")
    Public conRMOrange As New SqlClient.SqlConnection("")
    Public conRMBeMobile As New SqlClient.SqlConnection("")
    Public conServer As New SqlClient.SqlConnection("")
    
    Public mform As MainForm
    Public wrapper As New Simple3Des("ThithoH")
    Public logedinname As String
    Public employername As String
    Public employerid As Integer
    Public supervisorid As Integer
    Public DatabaseLocation As String

    Public branchid As Integer

    Public filterManager As New DgvFilterManager

    Public Function decision(ByVal str As String) As Integer
        Return MessageBox.Show(str, "Decision", MessageBoxButtons.YesNo)
    End Function


    Public Function checkDateFormat() As Boolean
        Dim currentdate, thedate, year, month, day As String
        currentdate = Date.Now.Date
        year = Date.Now.Date.Year
        month = Date.Now.Date.Month

        If month.Length = 1 Then
            month = "0" & month
        End If

        day = Date.Now.Date.Day
        If day.Length = 1 Then
            day = "0" & day
        End If

        thedate = year & "/" & month & "/" & day

        If currentdate = thedate Then
            Return True
        Else
            Return False
        End If

    End Function

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

    Public Function addMonth(ByVal dat As Date, ByVal int As Integer) As Date
        Dim dt As DateTime
        If int = 0 Then
            dt = dat
        Else
            dt = dat.AddMonths(int)
        End If

        Return dt
    End Function


    Public Function getLastDateOfMonth(ByVal dat As Date)
        Dim dt As DateTime = dat
        Dim year As Integer = dt.Year
        Dim month As Integer = dt.Month
        Dim days As Integer = DateTime.DaysInMonth(year, month)
        Dim lastDayOfMonth As New DateTime(year, month, days)
        Return lastDayOfMonth
    End Function

    Public Function getFirstDateOfMonth(ByVal dat As Date)
        Dim dt As DateTime = dat
        Dim year As Integer = dt.Year
        Dim month As Integer = dt.Month
        Dim days As Integer = DateTime.DaysInMonth(year, month)
        Dim lastDayOfMonth As New DateTime(year, month, 1)
        Return lastDayOfMonth
    End Function

    Public Sub saveProduct(ByVal Products As String, ByVal AllocationOrder As Integer, ByVal Status As String, ByVal AutoCalculate As Integer, ByVal HasSystem As Integer, ByVal SystemAbreviatedName As String, ByVal UseInCheques As String, ByVal UseInSettlement As String, ByVal UseInReceipts As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "Insert into Products values ('" & Products & "','" & AllocationOrder & "','" & Status & "','" & AutoCalculate & "','" & HasSystem & "','" & SystemAbreviatedName & "','" & UseInCheques & "','" & UseInSettlement & "','" & UseInReceipts & "')"
            Dim cmdinsert As New SqlClient.SqlCommand(qryinsert, con)
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            transactioninsert = con.BeginTransaction()
            cmdinsert.Transaction = transactioninsert

            If cmdinsert.ExecuteNonQuery() Then
                transactioninsert.Commit()
            End If

            MessageBox.Show("Saved", "saveProduct()", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "saveProduct()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "saveProduct()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub updateProduct(ByVal autoid As Integer, ByVal Products As String, ByVal Status As String, ByVal AllocationOrder As Integer, ByVal AutoCalculate As Integer, ByVal HasSystem As Integer, ByVal SystemAbreviatedName As String, ByVal UseInCheques As String, ByVal UseInSettlement As String, ByVal UseInReceipts As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update Products set Products = '" & Products & "',AllocationOrder = '" & AllocationOrder & "',Status = '" & Status & "',AutoAllocate = '" & AutoCalculate & "',HasSystem = '" & HasSystem & "',SystemAbreviatedName = '" & SystemAbreviatedName & "',UseInCheques = '" & UseInCheques & "',UseInSettlement = '" & UseInSettlement & "',UseInReceipts = '" & UseInReceipts & "' where AutoID = '" & autoid & "'"
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
            MessageBox.Show(ex.Message, "updateProduct()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateProduct()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub saveExpectedDeduction(ByVal CustomerID As String, ByVal ExpectedDeduction As Decimal, ByVal ForMonth As Date, ByVal GeneratedDate As Date, ByVal EmployerID As Integer)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try
            Dim qryinsert As String = "Insert into ExpectedDeductions values ('" & CustomerID & "','" & ExpectedDeduction & "','" & ForMonth & "','" & GeneratedDate & "','" & EmployerID & "')"
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
            MessageBox.Show(ex.Message, "saveExpectedDeduction()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "saveExpectedDeduction()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub saveActualDeduction(ByVal CustomerID As String, ByVal ActualDeduction As Decimal, ByVal ForMonth As Date, ByVal TransactionDate As Date, ByVal SavedDate As Date, ByVal EmployerID As Integer)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try
            Dim qryinsert As String = "Insert into ActualDeductions values ('" & CustomerID & "','" & ActualDeduction & "','" & ForMonth & "','" & TransactionDate & "','" & SavedDate & "','" & EmployerID & "')"
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
            MessageBox.Show(ex.Message, "saveActualDeduction()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "saveActualDeduction()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Function getActualDeduction(ByVal adate As Date, ByVal custid As String, ByVal EmployerID As Integer) As Decimal
        Dim qry As String = "select SUM(ActualDeduction) from ActualDeductions where ForMonth = '" & adate & "' and CustomerID = '" & custid & "' and EmployerID = '" & EmployerID & "'"

        Dim perc As Decimal = 0

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = 0
                Else
                    perc = Decimal.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
                End If
            Else
                perc = 0
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getActualDeduction()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getActualDeduction()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getProductAmendment(ByVal adate As Date, ByVal productid As String, ByVal EmployerID As Integer) As Decimal
        Dim qry As String = "select SUM(Amount) from TerminationsAndAmendments where ProductID = '" & productid & "' and ForMonth = '" & adate & "' and EmployerID = '" & EmployerID & "'"

        Dim perc As Decimal = 0

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = 0
                Else
                    perc = Decimal.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
                End If
            Else
                perc = 0
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getProductAmendment()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getProductAmendment()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Sub saveAmendments(ByVal CustomerID As String, ByVal Amount As Decimal, ByVal ForMonth As Date, ByVal SavedTime As Date, ByVal ProductID As Integer, ByVal EmployerID As Integer)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try
            Dim qryinsert As String = "Insert into TerminationsAndAmendments values ('" & CustomerID & "','" & Amount & "','" & ForMonth & "','" & SavedTime & "','" & ProductID & "','" & EmployerID & "')"
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
            MessageBox.Show(ex.Message, " saveAmendments()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, " saveAmendments()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub


    Public Function getProductsCount(ByVal adate As Date, ByVal productid As String, ByVal EmployerID As Integer) As Integer
        Dim qry As String = "select Count(Amount) from TerminationsAndAmendments where ProductID = '" & productid & "' and ForMonth = '" & adate & "' and Amount > 0.00 and EmployerID = '" & EmployerID & "'"

        Dim perc As Integer = 0

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = 0
                Else
                    perc = Integer.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
                End If
            Else
                perc = 0
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getProductsCount()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getProductsCount()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getTerminationsCount(ByVal adate As Date, ByVal productid As String, ByVal EmployerID As Integer) As Integer
        Dim qry As String = "select Count(Amount) from TerminationsAndAmendments where ProductID = '" & productid & "' and ForMonth = '" & adate & "' and Amount = 0.00 and EmployerID = '" & EmployerID & "'"

        Dim perc As Integer = 0

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = 0
                Else
                    perc = Integer.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
                End If
            Else
                perc = 0
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getTerminationsCount()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getTerminationsCount()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getExpectedPerProduct(ByVal formonthdate As Date, ByVal productid As Integer, ByVal customerid As String, ByVal predate As Date, ByVal EmployerID As Integer) As Decimal
        Dim qry As String = "select Amount from TerminationsAndAmendments where ProductID = '" & productid & "' and ForMonth = '" & formonthdate & "' and CustomerID = '" & customerid & "' and EmployerID = '" & EmployerID & "'"

        Dim perc As Decimal = 0

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = getLastExpectedPerProduct(predate, productid, customerid, EmployerID)
                Else
                    perc = Decimal.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
                End If
            Else
                perc = getLastExpectedPerProduct(predate, productid, customerid, EmployerID)
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getExpectedPerProduct()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getExpectedPerProduct()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getLastExpectedPerProduct(ByVal adate As Date, ByVal productid As Integer, ByVal customerid As String, ByVal EmployerID As Integer) As Decimal
        Dim qry As String = "select Amount from ExpectedDeductionsPerProduct where ProductID = '" & productid & "' and ForMonth = '" & adate & "' and CustomerID = '" & customerid & "' and EmployerID = '" & EmployerID & "'"

        Dim perc As Decimal = 0

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = 0
                Else
                    perc = Decimal.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
                End If
            Else
                perc = 0
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getLastExpectedPerProduct()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getLastExpectedPerProduct()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getExpectedDeductionPerCustomer(ByVal formonth As Date, ByVal customerid As String, ByVal EmployerID As Integer) As Decimal
        Dim qry As String = "select ExpectedDeduction from ExpectedDeductions where CustomerID = '" & customerid & "' and ForMonth = '" & formonth & "' and EmployerID = '" & EmployerID & "'"

        Dim perc As Decimal = 0

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = 0
                Else
                    perc = Decimal.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
                End If
            Else
                perc = 0
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getExpectedDeductionPerCustomer()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getExpectedDeductionPerCustomer()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function checkCustomerDetails(ByVal custid As String, ByVal EmployerID As Integer) As Boolean
        Dim qry As String = "select * from CustomerDetails where CustomerID = '" & custid & "' and EmployerID = '" & EmployerID & "'"

        Dim perc As Boolean = False

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = True
            Else
                perc = False
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "checkCustomerDetails()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "checkCustomerDetails()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function checkCustomerInDatabase(ByVal custid As String) As Boolean
        Dim qry As String = "select * from CustomerDetails where CustomerID = '" & custid & "'"

        Dim perc As Boolean = False

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = True
            Else
                perc = False
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "checkCustomerInDatabase()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "checkCustomerInDatabase()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function


    Public Sub saveCustomerDetails(ByVal CustomerID As String, ByVal CustomerName As String, ByVal EmployerID As Integer, ByVal Gender As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try
            Dim qryinsert As String = "Insert into CustomerDetails values ('" & CustomerID & "','" & CustomerName & "','" & EmployerID & "','" & Gender & "')"
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
            MessageBox.Show(ex.Message, "saveCustomerDetails()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "saveCustomerDetails()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub updateCustomerDetails(ByVal CustomerID As String, ByVal CustomerName As String, ByVal EmployerID As Integer, ByVal Gender As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update CustomerDetails set CustomerName = '" & CustomerName & "',EmployerID = '" & EmployerID & "',Gender = '" & Gender & "' where CustomerID = '" & CustomerID & "'"
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
            MessageBox.Show(ex.Message, "updateCustomerDetails()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateCustomerDetails()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Function checkExpectedDeductions(ByVal ddate As Date, ByVal employerid As Integer) As Boolean
        Dim qry As String = "select Top 1 * from ExpectedDeductions where ForMonth = '" & ddate & "' and CustomerID in (select CustomerID from CustomerDetails where EmployerID = '" & employerid & "') and EmployerID = '" & employerid & "'"

        Dim perc As Boolean = False

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = True
            Else
                perc = False
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "checkExpectedDeductions()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "checkExpectedDeductions()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function checkAmendment(ByVal customerid As String, ByVal formonth As Date, ByVal ProductID As Integer, ByVal EmployerID As Integer) As Boolean
        Dim qry As String = "select * from TerminationsAndAmendments where CustomerID = '" & customerid & "' and ForMonth = '" & formonth & "' and Amount > 0 and ProductID = '" & ProductID & "' and EmployerID = '" & EmployerID & "'"

        Dim perc As Boolean = False

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = True
            Else
                perc = False
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "checkAmendment()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "checkAmendment()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function checkTermination(ByVal customerid As String, ByVal formonth As Date, ByVal ProductID As Integer, ByVal EmployerID As Integer) As Boolean
        Dim qry As String = "select * from TerminationsAndAmendments where CustomerID = '" & customerid & "' and ForMonth = '" & formonth & "' and Amount = 0 and ProductID = '" & ProductID & "' and EmployerID = '" & EmployerID & "'"

        Dim perc As Boolean = False

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = True
            Else
                perc = False
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "checkTermination()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "checkTermination()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function checkAmendmentsAndTermiantionsForProduct(ByVal customerid As String, ByVal formonth As Date, ByVal ProductID As Integer, ByVal EmployerID As Integer) As Boolean
        Dim qry As String = "select * from TerminationsAndAmendments where CustomerID = '" & customerid & "' and ForMonth = '" & formonth & "' and ProductID = '" & ProductID & "' and EmployerID = '" & EmployerID & "'"

        Dim perc As Boolean = False

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = True
            Else
                perc = False
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "checkAmendmentsAndTermiantionsForProduct()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "checkAmendmentsAndTermiantionsForProduct()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function checkAmendmentsAndTermiantionsForCustomer(ByVal customerid As String, ByVal formonth As Date, ByVal EmployerID As Integer) As Boolean
        Dim qry As String = "select * from TerminationsAndAmendments where CustomerID = '" & customerid & "' and ForMonth = '" & formonth & "' and EmployerID = '" & EmployerID & "' and ProductID not in (select AutoID from Products where AutoAllocate = 1)"

        Dim perc As Boolean = False

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = True
            Else
                perc = False
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "checkAmendmentsAndTermiantionsForCustomer()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "checkAmendmentsAndTermiantionsForCustomer()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function


    Public Sub updateAmendment(ByVal customerid As String, ByVal formonth As Date, ByVal ProductID As Integer, ByVal Amount As Decimal, ByVal EmployerID As Integer)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update TerminationsAndAmendments set Amount = '" & Amount & "' where CustomerID = '" & customerid & "' and ForMonth = '" & formonth & "' and Amount > 0 and ProductID = '" & ProductID & "' and EmployerID = '" & EmployerID & "'"
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
            MessageBox.Show(ex.Message, "updateAmendment()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateAmendment()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub updateTermination(ByVal customerid As String, ByVal formonth As Date, ByVal ProductID As Integer, ByVal Amount As Decimal, ByVal EmployerID As Integer)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update TerminationsAndAmendments set Amount = '" & Amount & "' where CustomerID = '" & customerid & "' and ForMonth = '" & formonth & "' and Amount = 0 and ProductID = '" & ProductID & "' and EmployerID = '" & EmployerID & "'"
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
            MessageBox.Show(ex.Message, "updateTermination()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateTermination()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub saveExpectedDeductionsPerProduct(ByVal CustomerID As String, ByVal Amount As Decimal, ByVal ForMonth As Date, ByVal SavedTime As Date, ByVal ProductID As Integer, ByVal EmployerID As Integer)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try
            Dim qryinsert As String = "Insert into ExpectedDeductionsPerProduct values ('" & CustomerID & "','" & Amount & "','" & ForMonth & "','" & SavedTime & "','" & ProductID & "','" & EmployerID & "')"
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
            MessageBox.Show(ex.Message, "saveExpectedDeductionsPerProduct()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "saveExpectedDeductionsPerProduct()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Function checkActualDeduction(ByVal formonth As Date, ByVal employerid As Integer) As Boolean
        Dim qry As String = "select Top 1 * from ActualDeductions where Formonth = '" & formonth & "' and EmployerID = '" & employerid & "' and CustomerID in (select CustomerID from CustomerDetails where EmployerID = '" & employerid & "')"

        Dim perc As Boolean = False

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = True
            Else
                perc = False
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "checkActualDeduction()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "checkActualDeduction()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getLastLodgingDate(ByVal employerid As Integer) As Date
        Dim qry As String = "select Top 1 Formonth from ExpectedDeductions where CustomerID in (select CustomerID from CustomerDetails where EmployerID = '" & employerid & "') and EmployerID = '" & employerid & "' order by ForMonth DESC"

        Dim perc As Date

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = Date.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
            Else
                perc = Date.Now
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getLastLodgingDate()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getLastLodgingDate()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function checkUsers() As Boolean
        Dim qry As String = "select Top 1 * from Users"

        Dim perc As Boolean = False

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = True
            Else
                perc = False
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "checkUsers()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "checkUsers()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getAmendmentAmount(ByVal customerid As String, ByVal formonth As Date, ByVal ProductID As Integer, ByVal EmployerID As Integer) As Decimal
        Dim qry As String = "select SUM(Amount) from TerminationsAndAmendments where CustomerID = '" & customerid & "' and ForMonth = '" & formonth & "' and Amount > 0 and ProductID = '" & ProductID & "' and EmployerID = '" & EmployerID & "'"

        Dim perc As Decimal = 0.0

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = 0.0
                Else
                    perc = Decimal.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
                End If
            Else
                perc = 0.0
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getAmendmentAmount()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getAmendmentAmount()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function checkExpectedDeductionsPerCustomerAndForMonth(ByVal customerid As String, ByVal formonth As Date, ByVal EmployerID As Integer) As Boolean
        Dim qry As String = "select * from ExpectedDeductions where CustomerID = '" & customerid & "' and ForMonth = '" & formonth & "' and EmployerID = '" & EmployerID & "'"

        Dim perc As Boolean = False

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = True
            Else
                perc = False
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "checkExpectedDeductionsPerCustomerAndForMonth()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "checkExpectedDeductionsPerCustomerAndForMonth()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getExpectedDeductionID(ByVal customerid As String, ByVal formonth As Date, ByVal EmployerID As Integer) As Integer
        Dim qry As String = "select AutoID from ExpectedDeductions where CustomerID = '" & customerid & "' and ForMonth = '" & formonth & "' and EmployerID = '" & EmployerID & "' order by AutoID DESC"

        Dim perc As Boolean = False

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = True
            Else
                perc = False
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getExpectedDeductionID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getExpectedDeductionID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Sub updateExpectedDeductions(ByVal ExpectedDeduction As Decimal, ByVal AutoID As Integer, ByVal EmployerID As Integer)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update ExpectedDeductions set ExpectedDeduction = '" & ExpectedDeduction & "' where AutoID = '" & AutoID & "' and EmployerID = '" & EmployerID & "'"
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
            MessageBox.Show(ex.Message, "updateExpectedDeductions()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateExpectedDeductions()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Function checkExpectedDeductionsPerProduct(ByVal customerid As String, ByVal formonth As Date, ByVal ProductID As Integer, ByVal EmployerID As Integer) As Boolean
        Dim qry As String = "select * from ExpectedDeductionsPerProduct where CustomerID = '" & customerid & "' and Formonth = '" & formonth & "' and ProductID = '" & ProductID & "' and EmployerID = '" & EmployerID & "'"

        Dim perc As Boolean = False

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = True
            Else
                perc = False
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "checkExpectedDeductionsPerProduct()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "checkExpectedDeductionsPerProduct()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getExpectedDeductionIDPerProduct(ByVal customerid As String, ByVal formonth As Date, ByVal ProductID As Integer, ByVal EmployerID As Integer) As Integer
        Dim qry As String = "select AutoID from ExpectedDeductionsPerProduct where CustomerID = '" & customerid & "' and Formonth = '" & formonth & "' and ProductID = '" & ProductID & "' and EmployerID = '" & EmployerID & "' order by AutoID DESC"

        Dim perc As Boolean = False

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = True
            Else
                perc = False
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getExpectedDeductionIDPerProduct()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getExpectedDeductionIDPerProduct()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Sub updateExpectedDeductionsPerProduct(ByVal Amount As Decimal, ByVal AutoID As Integer, ByVal EmployerID As Integer)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update ExpectedDeductionsPerProduct set Amount = '" & Amount & "' where AutoID = '" & AutoID & "' and EmployerID = '" & EmployerID & "'"
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
            MessageBox.Show(ex.Message, "updateExpectedDeductionsPerProduct()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateExpectedDeductionsPerProduct()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Function checkEmployer(ByVal employerid As Integer) As Boolean
        Dim qry As String = "select * from Employers where AutoID = '" & employerid & "'"

        Dim perc As Boolean = False

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = True
            Else
                perc = False
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "checkEmployer()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "checkEmployer()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function checkEmployersAvailable() As Boolean
        Dim qry As String = "select * from Employers"

        Dim perc As Boolean = False

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = True
            Else
                perc = False
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "checkEmployersAvailable()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "checkEmployersAvailable()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Sub updateCustomerEmployer(ByVal Employerid As Integer, ByVal CustomerID As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update CustomerDetails set EmployerID = '" & Employerid & "' where CustomerID = '" & CustomerID & "'"
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
            MessageBox.Show(ex.Message, "updateCustomerEmployer()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateCustomerEmployer()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Function checkCustomerDetailsInDatabase(ByVal custid As String) As Boolean
        Dim qry As String = "select * from CustomerDetails where CustomerID = '" & custid & "'"

        Dim perc As Boolean = False

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = True
            Else
                perc = False
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "checkCustomerDetails()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "checkCustomerDetails()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getLastActualDeductionDate(ByVal employerid As Integer) As Date
        Dim qry As String = "select Top 1 ForMonth from ActualDeductions where CustomerID in (select CustomerID from CustomerDetails where EmployerID = '" & employerid & "') and  EmployerID = '" & employerid & "' order by ForMonth DESC"

        Dim perc As Date

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = Date.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
            Else
                perc = Date.Now
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getLastActualDeductionDate()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getLastActualDeductionDate()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getLastActualDeductionDateByCustomerID(ByVal customerid As String) As Date
        Dim qry As String = "select Top 1 ForMonth from ActualDeductions where CustomerID = '" & customerid & "' order by ForMonth DESC"

        Dim perc As Date

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = Date.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
            Else
                perc = Date.Now
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getLastActualDeductionDateByCustomerID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getLastActualDeductionDateByCustomerID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function checkExpectedDeductionDate(ByVal formonth As Date, ByVal EmployerID As Integer) As Boolean
        Dim qry As String = "select Top 1 * from ExpectedDeductions where ForMonth = '" & formonth & "' and EmployerID = '" & EmployerID & "'"

        Dim perc As Boolean = False

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = True
            Else
                perc = False
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "checkExpectedDeductionDate()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "checkExpectedDeductionDate()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function checkActualDeductionDate(ByVal formonth As Date, ByVal EmployerID As Integer) As Boolean
        Dim qry As String = "select Top 1 * from ActualDeductions where ForMonth = '" & formonth & "' and EmployerID = '" & EmployerID & "'"

        Dim perc As Boolean = False

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = True
            Else
                perc = False
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "checkActualDeductionDate()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "checkActualDeductionDate()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getEmployerNameByCustomerID(ByVal CustomerID As String) As String
        Dim qry As String = "select EmployerName from Employers where AutoID in (select EmployerID from CustomerDetails where CustomerID = '" & CustomerID & "')"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = ds.Tables(0).Rows(0).Item(0).ToString
            Else
                perc = ""
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getEmployerNameByCustomerID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getEmployerNameByCustomerID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getEmployerIDByCustomerID(ByVal CustomerID As String) As Integer
        Dim qry As String = "select EmployerID from CustomerDetails where CustomerID = '" & CustomerID & "'"

        Dim perc As Integer = 0

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = Integer.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
            Else
                perc = 0
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getEmployerIDByCustomerID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getEmployerIDByCustomerID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getCustomerName(ByVal customerid As String) As String
        Dim qry As String = "select CustomerName from CustomerDetails where CustomerID = '" & customerid & "'"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = ""
                Else
                    perc = ds.Tables(0).Rows(0).Item(0).ToString
                End If
            Else
                perc = ""
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getCustomerName()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getCustomerName()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getCustomerDetails(ByVal customerid As String) As DataSet
        Dim qry As String = "select CustomerName from CustomerDetails where CustomerID = '" & customerid & "'"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getCustomerDetails()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getCustomerDetails()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return ds
    End Function

    Public Function getSupervisorDetails(ByVal autoid As Integer) As DataSet
        Dim qry As String = "select * from Supervisors where AutoID = '" & autoid & "'"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getSupervisorDetails()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getSupervisorDetails()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return ds
    End Function

    Public Function getSettlementDetails() As DataSet
        Dim qry As String = "select * from SettlementDetails"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getSettlementDetails()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getSettlementDetails()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return ds
    End Function

    Public Function getNumberOfPeriodsToAdd() As Integer
        Dim qry As String = "select Count(AutoID) from Products where IncludePaymentPeriod = 1"

        Dim perc As Integer = 0

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = 0
                Else
                    perc = Integer.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
                End If
            Else
                perc = 0
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getNumberOfPeriodsToAdd()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getNumberOfPeriodsToAdd()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getProductsPeriod() As DataSet
        Dim qry As String = "select * from Products where IncludePaymentPeriod = 1"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getProductsPeriod()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getProductsPeriod()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return ds
    End Function

    Public Function getProductID(ByVal EmployerID As Integer) As Integer
        Dim qry As String = "select ProductID from Employers where AutoID = '" & EmployerID & "'"

        Dim perc As Integer = 0

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = 0
                Else
                    perc = Integer.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
                End If
            Else
                perc = 0
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getProductID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getProductID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getCommissionPercent(ByVal EmployerID As Integer) As Decimal
        Dim qry As String = "select CommissionPercent from Employers where AutoID = '" & EmployerID & "'"

        Dim perc As Decimal

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = 0
                Else
                    perc = Decimal.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
                End If
            Else
                perc = 0
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getCommissionPercent()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getCommissionPercent()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getCommissionProdcuts(ByVal employerid As Integer) As DataSet
        Dim qry As String = "select * from CommissionProducts where Status = 1 and EmployerID = '" & employerid & "'"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getCommissionProdcuts()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getCommissionProdcuts()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return ds
    End Function

    Public Function getProductSystem(ByVal productid As Integer) As String
        Dim qry As String = "select SystemAbreviatedName from Products where AutoID = '" & productid & "'"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = ""
                Else
                    perc = ds.Tables(0).Rows(0).Item(0).ToString
                End If
            Else
                perc = ""
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getProductSystem()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getProductSystem()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getFPMFuneralRefund(ByVal customerid As String, ByVal paymentdate As Date, ByVal refundby As String, ByVal con As SqlClient.SqlConnection) As Decimal
        Dim qry As String = "select SUM(Amount) from CustomerRefunds where RefundBy = '" & refundby & "' and PaymentDate = '" & paymentdate & "' and CustomerID = '" & customerid & "' and Status = 'Approved'"

        Dim perc As Decimal

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = 0
                Else
                    perc = Decimal.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
                End If
            Else
                perc = 0
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getFPMFuneralRefund()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getFPMFuneralRefund()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getFPMMembershipRefund(ByVal PayerID As String, ByVal paymentdate As Date, ByVal PayBy As String, ByVal con As SqlClient.SqlConnection) As Decimal
        Dim qry As String = "select SUM(Amount) from MembershipRefunds where PayBy = '" & PayBy & "' and PaymentDate = '" & paymentdate & "' and PayerID = '" & PayerID & "' and Status = 'Approved'"

        Dim perc As Decimal

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = 0
                Else
                    perc = Decimal.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
                End If
            Else
                perc = 0
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getFPMMembershipRefund()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getFPMMembershipRefund()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getFPMSavingsRefund(ByVal PayerID As String, ByVal paymentdate As Date, ByVal PayBy As String, ByVal con As SqlClient.SqlConnection) As Decimal
        Dim qry As String = "select SUM(Amount) from SavingsRefunds where PayBy = '" & PayBy & "' and PaymentDate = '" & paymentdate & "' and PayerID = '" & PayerID & "' and Status = 'Approved'"

        Dim perc As Decimal

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = 0
                Else
                    perc = Decimal.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
                End If
            Else
                perc = 0
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getFPMSavingsRefund()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getFPMSavingsRefund()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getFPMFuneralRefundPayByReference(ByVal customerid As String, ByVal paymentdate As Date, ByVal refundby As String, ByVal con As SqlClient.SqlConnection) As DataSet
        Dim qry As String = "select Bank,BranchCode,AccountNumber,PayeeName from CustomerRefunds where RefundBy = '" & refundby & "' and PaymentDate = '" & paymentdate & "' and CustomerID = '" & customerid & "' and Status = 'Approved'"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getFPMFuneralRefundPayByReference()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getFPMFuneralRefundPayByReference()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return ds
    End Function

    Public Function getFPMMembeshipRefundPayByReference(ByVal PayerID As String, ByVal paymentdate As Date, ByVal PayBy As String, ByVal con As SqlClient.SqlConnection) As DataSet
        Dim qry As String = "select Bank,BranchCode,AccountNumber,PayeeName from MembershipRefunds where PayBy = '" & PayBy & "' and PaymentDate = '" & paymentdate & "' and PayerID = '" & PayerID & "' and Status = 'Approved'"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getFPMMembeshipRefundPayByReference()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getFPMMembeshipRefundPayByReference()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return ds
    End Function

    Public Function getFPMSavingsRefundPayByReference(ByVal PayerID As String, ByVal paymentdate As Date, ByVal PayBy As String, ByVal con As SqlClient.SqlConnection) As DataSet
        Dim qry As String = "select Bank,BranchCode,AccountNumber,PayeeName from SavingsRefunds where PayBy = '" & PayBy & "' and PaymentDate = '" & paymentdate & "' and PayerID = '" & PayerID & "' and Status = 'Approved'"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getFPMSavingsRefundPayByReference()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getFPMSavingsRefundPayByReference()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return ds
    End Function

    Public Function getFPMFuneralRefundCustomer(ByVal paymentdate As Date, ByVal refundby As String, ByVal con As SqlClient.SqlConnection, ByVal CustomerID As String) As DataSet
        Dim qry As String = "select * from CustomerRefunds where RefundBy = '" & refundby & "' and PaymentDate = '" & paymentdate & "' and Status = 'Approved' and CustomerID = '" & CustomerID & "'"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getFPMFuneralRefundCustomer()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getFPMFuneralRefundCustomer()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return ds
    End Function

    Public Function getFPMMembershipRefundCustomer(ByVal paymentdate As Date, ByVal PayBy As String, ByVal con As SqlClient.SqlConnection, ByVal PayerID As String) As DataSet
        Dim qry As String = "select * from MembershipRefunds where PayBy = '" & PayBy & "' and PaymentDate = '" & paymentdate & "' and Status = 'Approved' and PayerID = '" & PayerID & "'"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getFPMMembershipRefundCustomer()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getFPMMembershipRefundCustomer()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return ds
    End Function

    Public Function getFPMSavingsRefundCustomer(ByVal paymentdate As Date, ByVal PayBy As String, ByVal con As SqlClient.SqlConnection, ByVal PayerID As String) As DataSet
        Dim qry As String = "select * from SavingsRefunds where PayBy = '" & PayBy & "' and PaymentDate = '" & paymentdate & "' and Status = 'Approved' and PayerID = '" & PayerID & "'"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getFPMSavingsRefundCustomer()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getFPMSavingsRefundCustomer()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return ds
    End Function

    Public Function getMLMRefund(ByVal customerid As String, ByVal paymentdate As Date, ByVal loantype As String, ByVal payby As String, ByVal con As SqlClient.SqlConnection) As Decimal
        Dim qry As String = "select SUM(Amount) from Refunds where PayBy = '" & payby & "' and Status = 'Posted' and LoanType = '" & loantype & "' and CustomerID = '" & customerid & "' and AutoID in (select TypeID from PaymentDates where Type = 'Refund' and CustomerID = '" & customerid & "' and PaymentDate = '" & paymentdate & "')"

        Dim perc As Decimal

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = 0
                Else
                    perc = Decimal.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
                End If
            Else
                perc = 0
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getMLMRefund()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getMLMRefund()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getMLMRefundPayByReference(ByVal customerid As String, ByVal paymentdate As Date, ByVal loantype As String, ByVal payby As String, ByVal con As SqlClient.SqlConnection) As DataSet
        Dim qry As String = "select BankName,BranchCode,AccountNumber,PayeeName from Refunds where PayBy = '" & payby & "' and Status = 'Posted' and LoanType = '" & loantype & "' and CustomerID = '" & customerid & "' and AutoID in (select TypeID from PaymentDates where Type = 'Refund' and CustomerID = '" & customerid & "' and PaymentDate = '" & paymentdate & "')"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getMLMRefundPayByReference()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getMLMRefundPayByReference()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return ds
    End Function


    Public Function getMLMRefundCustomer(ByVal paymentdate As Date, ByVal loantype As String, ByVal payby As String, ByVal con As SqlClient.SqlConnection, ByVal CustomerID As String) As DataSet
        Dim qry As String = "select * from Refunds where PayBy = '" & payby & "' and Status = 'Posted' and LoanType = '" & loantype & "' and AutoID in (select TypeID from PaymentDates where Type = 'Refund' and PaymentDate = '" & paymentdate & "') and CustomerID = '" & CustomerID & "'"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getMLMRefundCustomer()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getMLMRefundCustomer()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return ds
    End Function

    Public Function getMLMRLoan(ByVal customerid As String, ByVal paymentdate As Date, ByVal loantype As String, ByVal payby As String, ByVal con As SqlClient.SqlConnection) As Decimal
        Dim qry As String = "select SUM(Disbursement_Amount) from LoanDetails where PayBy = '" & payby & "' and LoanStatus = 'Active' and LoanDescription = '" & loantype & "' and Customer_ID = '" & customerid & "' and AutoLoanID in (select TypeID from PaymentDates where Type = 'Loan' and CustomerID = '" & customerid & "' and PaymentDate = '" & paymentdate & "')"

        Dim perc As Decimal

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = 0
                Else
                    perc = Decimal.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
                End If
            Else
                perc = 0
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getMLMRLoan()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getMLMRLoan()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getMLMRLoanPayByReference(ByVal customerid As String, ByVal paymentdate As Date, ByVal loantype As String, ByVal payby As String, ByVal con As SqlClient.SqlConnection) As DataSet
        Dim qry As String = "select BankName,BranchCode,AccountNumber,PayeeName from LoanDetails where PayBy = '" & payby & "' and LoanStatus = 'Active' and LoanDescription = '" & loantype & "' and Customer_ID = '" & customerid & "' and AutoLoanID in (select TypeID from PaymentDates where Type = 'Loan' and CustomerID = '" & customerid & "' and PaymentDate = '" & paymentdate & "')"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getMLMRLoanPayByReference()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getMLMRLoanPayByReference()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return ds
    End Function

    Public Function getMLMRLoanCustomer(ByVal paymentdate As Date, ByVal loantype As String, ByVal payby As String, ByVal con As SqlClient.SqlConnection, ByVal Customer_ID As String) As DataSet
        Dim qry As String = "select * from LoanDetails where PayBy = '" & payby & "' and LoanStatus = 'Active' and LoanDescription = '" & loantype & "' and AutoLoanID in (select TypeID from PaymentDates where Type = 'Loan' and PaymentDate = '" & paymentdate & "') and Customer_ID = '" & Customer_ID & "'"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getMLMRLoanCustomer()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getMLMRLoanCustomer()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return ds
    End Function

    Public Function getMLMRLoanDetails(ByVal customerid As String, ByVal paymentdate As Date, ByVal loantype As String, ByVal payby As String, ByVal con As SqlClient.SqlConnection) As DataSet
        Dim qry As String = "select Principal_Amount,Disbursement_Amount,Rollover_Amount from LoanDetails where PayBy = '" & payby & "' and LoanStatus = 'Active' and LoanDescription = '" & loantype & "' and Customer_ID = '" & customerid & "' and AutoLoanID in (select TypeID from PaymentDates where Type = 'Loan' and CustomerID = '" & customerid & "' and PaymentDate = '" & paymentdate & "')"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getMLMRLoanDetails()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getMLMRLoanDetails()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return ds
    End Function

    Public Function getRMRefund(ByVal customerid As String, ByVal paymentdate As Date, ByVal refundby As String, ByVal con As SqlClient.SqlConnection) As Decimal
        Dim qry As String = "select SUM(Amount) from CustomerRefunds where RefundBy = '" & refundby & "' and PaymentDate = '" & paymentdate & "' and Status = 'Posted' and CustomerID = '" & customerid & "'"

        Dim perc As Decimal

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = 0
                Else
                    perc = Decimal.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
                End If
            Else
                perc = 0
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getRMRefund()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getRMRefund()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getRMRefundPayByReference(ByVal customerid As String, ByVal paymentdate As Date, ByVal refundby As String, ByVal con As SqlClient.SqlConnection) As DataSet
        Dim qry As String = "select Bank,BranchCode,AccountNumber,PayeeName from CustomerRefunds where RefundBy = '" & refundby & "' and PaymentDate = '" & paymentdate & "' and Status = 'Posted' and CustomerID = '" & customerid & "'"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getRMRefundPayByReference()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getRMRefundPayByReference()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return ds
    End Function

    Public Function getRMRefundCustomer(ByVal paymentdate As Date, ByVal refundby As String, ByVal con As SqlClient.SqlConnection, ByVal CustomerID As String) As DataSet
        Dim qry As String = "select * from CustomerRefunds where RefundBy = '" & refundby & "' and PaymentDate = '" & paymentdate & "' and Status = 'Posted' and CustomerID = '" & CustomerID & "'"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getRMRefundCustomer()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getRMRefundCustomer()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return ds
    End Function

    Public Function getSystemCheque(ByVal customerid As String, ByVal paymentdate As Date, ByVal productid As Integer, ByVal payby As String) As Decimal
        Dim qry As String = "select SUM(Amount) from ProductsCheques where CustomerID = '" & customerid & "' and ProductID = '" & productid & "' and PaymentDate = '" & paymentdate & "' and Status = 'Approved' and PayBy = '" & payby & "'"

        Dim perc As Decimal

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = 0
                Else
                    perc = Decimal.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
                End If
            Else
                perc = 0
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getSystemCheque()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getSystemCheque()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getSystemChequePayByReference(ByVal customerid As String, ByVal paymentdate As Date, ByVal productid As Integer, ByVal payby As String) As DataSet
        Dim qry As String = "select Bank,BranchCode,AccountNumber,PayeeName from ProductsCheques where CustomerID = '" & customerid & "' and ProductID = '" & productid & "' and PaymentDate = '" & paymentdate & "' and Status = 'Approved' and PayBy = '" & payby & "'"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getSystemChequePayByReference()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getSystemChequePayByReference()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return ds
    End Function

    Public Function getSystemChequeCustomer(ByVal paymentdate As Date, ByVal payby As String, ByVal CustomerID As String) As DataSet
        Dim qry As String = "select * from ProductsCheques where PaymentDate = '" & paymentdate & "' and Status = 'Approved' and PayBy = '" & payby & "' and CustomerID = '" & customerid & "'"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getSystemChequeCustomer()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getSystemChequeCustomer()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return ds
    End Function

    Public Sub saveRefund(ByVal CustomerID As String, ByVal ProductID As Integer, ByVal Amount As Decimal, ByVal PaymentDate As Date, ByVal CapturedBy As String, ByVal CapturedDate As Date, ByVal ChequeNumber As String, ByVal Status As String, ByVal PayBy As String, ByVal Bank As String, ByVal BranchCode As String, ByVal AccountNumber As String, ByVal PayeeName As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try
            Dim qryinsert As String = "Insert into ProductsCheques values ('" & CustomerID & "','" & ProductID & "','" & Amount & "','" & PaymentDate & "','" & CapturedBy & "','" & CapturedDate & "','" & ChequeNumber & "','" & Status & "','" & PayBy & "','" & Bank & "','" & BranchCode & "','" & AccountNumber & "','" & PayeeName & "')"
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
            MessageBox.Show(ex.Message, "saveRefund()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "saveRefund()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub updateRefundStatus(ByVal autoid As Integer, ByVal Status As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update ProductsCheques set Status = '" & Status & "' where AutoID = '" & autoid & "'"
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
            MessageBox.Show(ex.Message, "updateRefundStatus()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateRefundStatus()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Function checkFPMPendingRefund(ByVal customerid As String,byval con as SqlClient.SqlConnection) As Boolean
        Dim qry As String = "select TOP 1 * from CustomerRefunds where CustomerID = '" & customerid & "' and Status = 'New'"

        Dim perc As Boolean = False

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = True
            Else
                perc = False
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "checkFPMPendingRefund()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "checkFPMPendingRefund()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function checkMLMPendingRefund(ByVal customerid As String, ByVal con As SqlClient.SqlConnection) As Boolean
        Dim qry As String = "select TOP 1 * from Refunds where Status = 'New' and CustomerID = '" & customerid & "'"

        Dim perc As Boolean = False

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = True
            Else
                perc = False
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "checkFPMPendingRefund()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "checkFPMPendingRefund()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function checkMLMPendingLoan(ByVal customerid As String, ByVal con As SqlClient.SqlConnection) As Boolean
        Dim qry As String = "select TOP 1 * from LoanDetails where LoanStatus = 'Pending' and Customer_ID = '" & customerid & "'"

        Dim perc As Boolean = False

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = True
            Else
                perc = False
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "checkMLMPendingLoan()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "checkMLMPendingLoan()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function checkRMPendingRefund(ByVal customerid As String, ByVal con As SqlClient.SqlConnection) As Boolean
        Dim qry As String = "select TOP 1 * from CustomerRefunds where CustomerID = '" & customerid & "' and Status = 'New'"

        Dim perc As Boolean = False

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = True
            Else
                perc = False
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "checkRMPendingRefund()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "checkRMPendingRefund()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function checkLocalPendingRefund(ByVal customerid As String) As Boolean
        Dim qry As String = "select TOP 1 * from ProductsCheques where Status = 'New' and CustomerID = '" & customerid & "'"

        Dim perc As Boolean = False

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = True
            Else
                perc = False
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "checkLocalPendingRefund()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "checkLocalPendingRefund()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Sub saveEmployers(ByVal EmployerName As String, ByVal CommissionPercent As Decimal, ByVal ProductID As Integer, ByVal AllowToEdit As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            transactioninsert = con.BeginTransaction()
            Dim qryinsert As String = "Insert into Employers values ('" & EmployerName & "','" & CommissionPercent & "','" & ProductID & "','" & AllowToEdit & "')"
            Dim cmdinsert As New SqlClient.SqlCommand(qryinsert, con)
            cmdinsert.Transaction = transactioninsert
            If cmdinsert.ExecuteNonQuery() Then
                transactioninsert.Commit()
            End If
            MessageBox.Show("Details saved", "saveEmployers()", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "saveEmployers()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "saveEmployers()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Function checkChequeCustomerEmployer() As Boolean
        Dim qry As String = "select * from Employers where EmployerName = 'Cheque Customers'"

        Dim perc As Boolean = False

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = True
            Else
                perc = False
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "checkChequeCustomerEmployer()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "checkChequeCustomerEmployer()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getChequeCustomerEmployerID() As Integer
        Dim qry As String = "select AutoID from Employers where EmployerName = 'Cheque Customers'"

        Dim perc As Integer = 0

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = Integer.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
            Else
                perc = 0
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getChequeCustomerEmployerID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getChequeCustomerEmployerID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getFPMCustomerName(ByVal customerid As String, ByVal con As SqlClient.SqlConnection) As String
        Dim qry As String = "select FirstName + ' ' + Surname from CustomerDetails where IDNumber = '" & customerid & "'"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = ""
                Else
                    perc = ds.Tables(0).Rows(0).Item(0).ToString
                End If
            Else
                perc = ""
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getFPMCustomerName()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getFPMCustomerName()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getMLMCustomerName(ByVal customerid As String, ByVal con As SqlClient.SqlConnection) As String
        Dim qry As String = "select First_Name + ' ' + Surname from CustomerDetails where ID_Number = '" & customerid & "'"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = ""
                Else
                    perc = ds.Tables(0).Rows(0).Item(0).ToString
                End If
            Else
                perc = ""
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getFPMCustomerName()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getFPMCustomerName()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getRMCustomerName(ByVal customerid As String, ByVal con As SqlClient.SqlConnection) As String
        Dim qry As String = "select FirstName + ' ' + Surname from CustomerMaster where Account = '" & customerid & "'"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = ""
                Else
                    perc = ds.Tables(0).Rows(0).Item(0).ToString
                End If
            Else
                perc = ""
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getRMCustomerName()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getRMCustomerName()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getRMCustomerGender(ByVal customerid As String, ByVal con As SqlClient.SqlConnection) As String
        Dim qry As String = "select Gender from CustomerMaster where Account = '" & customerid & "'"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = ""
                Else
                    perc = ds.Tables(0).Rows(0).Item(0).ToString
                End If
            Else
                perc = ""
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getRMCustomerGender()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getRMCustomerGender()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getMLMCustomerGender(ByVal customerid As String, ByVal con As SqlClient.SqlConnection) As String
        Dim qry As String = "select Gender from CustomerDetails where ID_Number = '" & customerid & "'"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = ""
                Else
                    perc = ds.Tables(0).Rows(0).Item(0).ToString
                End If
            Else
                perc = ""
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getMLMCustomerGender()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getMLMCustomerGender()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getFPMCustomerGender(ByVal customerid As String, ByVal con As SqlClient.SqlConnection) As String
        Dim qry As String = "select Gender from CustomerDetails where IDNumber = '" & customerid & "'"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = ""
                Else
                    perc = ds.Tables(0).Rows(0).Item(0).ToString
                End If
            Else
                perc = ""
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getFPMCustomerGender()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getFPMCustomerGender()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getGender(ByVal customerid As String) As String
        Dim qry As String = "select Gender from CustomerDetails where CustomerID = '" & customerid & "'"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = ""
                Else
                    perc = ds.Tables(0).Rows(0).Item(0).ToString
                End If
            Else
                perc = ""
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getGender()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getGender()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function checkIFProductHasSystem(ByVal productid As Integer) As Boolean
        Dim qry As String = "select * from Products where AutoID = '" & productid & "' and HasSystem = 1"

        Dim perc As Boolean = False

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = True
            Else
                perc = False
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "checkIFProductHasSystem()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "checkIFProductHasSystem()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function checkPendingCheques(ByVal paymentdate As Date, ByVal customerid As String) As Boolean
        Dim qry As String = "select * from ProductsCheques where PaymentDate = '" & paymentdate & "' and Status in ('New','Approved','Requested') and CustomerID = '" & customerid & "'"

        Dim perc As Boolean = False

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = True
            Else
                perc = False
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "checkPendingCheques()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "checkPendingCheques()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function checkRequestedCheques(ByVal paymentdate As Date, ByVal customerid As String) As Boolean
        Dim qry As String = "select * from ProductsCheques where PaymentDate = '" & paymentdate & "' and Status in ('Requested') and CustomerID = '" & customerid & "'"

        Dim perc As Boolean = False

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = True
            Else
                perc = False
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "checkRequestedCheques()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "checkRequestedCheques()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Sub updateSystemRefundStatusToRequested(ByVal productid As Integer, ByVal customerid As String, ByVal paymentdate As Date, ByVal Status As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update ProductsCheques set Status = '" & Status & "' where ProductID = '" & productid & "' and CustomerID = '" & customerid & "' and PaymentDate = '" & paymentdate & "' and Status = 'Approved'"
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
            MessageBox.Show(ex.Message, "updateSystemRefundStatusToRequested()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateSystemRefundStatusToRequested()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub updateSystemRefundStatusToWritten(ByVal customerid As String, ByVal paymentdate As Date, ByVal checknumber As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update ProductsCheques set Status = 'Written',ChequeNumber = '" & checknumber & "' where Status = 'Requested' and CustomerID = '" & customerid & "' and PaymentDate = '" & paymentdate & "'"
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
            MessageBox.Show(ex.Message, "updateSystemRefundStatusToWritten()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateSystemRefundStatusToWritten()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub updateSystemRefundChequeNumber(ByVal productid As Integer, ByVal customerid As String, ByVal paymentdate As Date, ByVal ChequeNumber As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update ProductsCheques set ChequeNumber = '" & ChequeNumber & "' where ProductID = '" & productid & "' and CustomerID = '" & customerid & "' and PaymentDate = '" & paymentdate & "' and Status = 'Approved'"
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
            MessageBox.Show(ex.Message, "updateSystemRefundChequeNumber()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateSystemRefundChequeNumber()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub saveBalance(ByVal CustomerID As String, ByVal Principal As Decimal, ByVal Disbursement As Decimal, ByVal Rollover As Decimal, ByVal PaymentDate As Date, ByVal LoanType As String, ByVal Status As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try
            Dim qryinsert As String = "Insert into Balances values ('" & CustomerID & "','" & Principal & "','" & Disbursement & "','" & Rollover & "','" & PaymentDate & "','" & LoanType & "','" & Status & "')"
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
            MessageBox.Show(ex.Message, "saveBalance()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "saveBalance()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Function getPrincipal(ByVal customerid As String, ByVal paymentdate As Date, ByVal loantype As String) As Decimal
        Dim qry As String = "select Principal from Balances where CustomerID = '" & customerid & "' and PaymentDate = '" & paymentdate & "' and LoanType = '" & loantype & "'"

        Dim perc As Decimal

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = 0
                Else
                    perc = Decimal.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
                End If
            Else
                perc = 0
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getPrincipal()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getPrincipal()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function


    Public Sub saveCancelledCheque(ByVal CustomerID As String, ByVal ChequeNumber As String, ByVal CapturedDate As Date, ByVal CapturedBy As String, ByVal PaymentDate As Date, ByVal Status As String, ByVal PayBy As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try
            Dim qryinsert As String = "Insert into CancelledCheques values ('" & CustomerID & "','" & ChequeNumber & "','" & CapturedDate & "','" & CapturedBy & "','" & PaymentDate & "','" & Status & "','" & PayBy & "')"
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
            MessageBox.Show(ex.Message, "saveCancelledCheque()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "saveCancelledCheque()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub updateSystemRefundStatusToRequestedAfterCancellation(ByVal customerid As String, ByVal paymentdate As Date, ByVal checknumber As String, ByVal status As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update ProductsCheques set Status = '" & status & "',ChequeNumber = '' where Status = 'Written' and CustomerID = '" & customerid & "' and PaymentDate = '" & paymentdate & "' and ChequeNumber = '" & checknumber & "'"
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
            MessageBox.Show(ex.Message, "updateSystemRefundStatusToRequestedAfterCancellation()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateSystemRefundStatusToRequestedAfterCancellation()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Function checkSystemCheque(ByVal customerid As String, ByVal paymentdate As Date, ByVal productid As Integer, ByVal amount As Decimal, ByVal payby As String) As Boolean
        Dim qry As String = "select * from ProductsCheques where CustomerID = '" & customerid & "' and ProductID = '" & productid & "' and PaymentDate = '" & paymentdate & "' and Amount = '" & amount & "' and PayBy = '" & payby & "'"

        Dim perc As Boolean = True

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = False
                Else
                    perc = True
                End If
            Else
                perc = False
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "checkSystemCheque()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "checkSystemCheque()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getSYstemVersion() As String
        Dim qry As String = "select SystemVersion from CompanyDetails"
        Dim ds As New DataSet
        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)
        Dim str As String = ""

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim ctr As Integer = 0
            SQLAdapter.Fill(ds, "Users")

            str = ds.Tables(0).Rows(0).Item(0).ToString

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getSYstemVersion()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ds.Clear()
            con.Close()
        End Try
        Return str
    End Function

    Public Function getSupervisorName(ByVal autoid As Integer) As String
        Dim qry As String = "select SupervisorName from Supervisors where AutoID = '" & autoid & "'"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = ds.Tables(0).Rows(0).Item(0).ToString
            Else
                perc = ""
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getSupervisorName()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getSupervisorName()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getCutOfDate() As Integer
        Dim qry As String = "select CutOfDate from CompanyDetails"

        Dim perc As Integer = 0

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = 0
                Else
                    perc = Integer.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
                End If
            Else
                perc = 0
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getCutOfDate()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getCutOfDate()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Sub updateSettlementLetterDesign(ByVal ToWhom As String, ByVal Re As String, ByVal OfID As String, ByVal ThisServes As String, ByVal HasAccount As String, ByVal PleaseBe As String, ByVal IfYouIntend As String, ByVal ThitoAccount As String, ByVal ThereAfter As String, ByVal ForMoreInfo As String, ByVal Yours As String, ByVal SupervisorName As String, ByVal JobDescription As String, ByVal EmailAddress As String, ByVal PleaseBeInformed As String, ByVal CutOfDate As Integer, ByVal EarlyPayPercentage As Decimal, ByVal PleaseEnsure As String, ByVal ToAvoid As String, ByVal ValidityDate As Integer, ByVal con As SqlClient.SqlConnection)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update SettlementDetails set ToWhom = '" & ToWhom & "',Re = '" & Re & "',OfID = '" & OfID & "',ThisServes = '" & ThisServes & "',HasAccount = '" & HasAccount & "',PleaseBe = '" & PleaseBe & "',IfYouIntend = '" & IfYouIntend & "',ThitoAccount = '" & ThitoAccount & "',ThereAfter = '" & ThereAfter & "',ForMoreInfo = '" & ForMoreInfo & "',Yours = '" & Yours & "',SupervisorName = '" & SupervisorName & "',JobDescription = '" & JobDescription & "',EmailAddress = '" & EmailAddress & "',PleaseBeInformed = '" & PleaseBeInformed & "',CutOfDate = '" & CutOfDate & "',EarlyPayPercentage = '" & EarlyPayPercentage & "',PleaseEnsure = '" & PleaseEnsure & "',ToAvoid = '" & ToAvoid & "',ValidityDate = '" & ValidityDate & "'"
            Dim cmdinsert As New SqlClient.SqlCommand(qryinsert, con)
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            transactioninsert = con.BeginTransaction()
            cmdinsert.Transaction = transactioninsert

            If cmdinsert.ExecuteNonQuery() Then
                transactioninsert.Commit()
            End If
            MessageBox.Show("Done Saving updates", "updateSettlementLetterDesign()", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "updateSettlementLetterDesign()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateSettlementLetterDesign()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub updateGender(ByVal gender As String, ByVal AutoID As Integer)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update CustomerDetails set Gender = '" & gender & "' where AutoCustID = '" & AutoID & "'"
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
            MessageBox.Show(ex.Message, "updateGender()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateGender()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Function getCurrency(ByVal num As Decimal) As String
        Return Format(num, "P #,##0.00")
    End Function

    Public Function getClearanceDetails() As DataSet
        Dim qry As String = "select * from ClearanceDetails"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)
        Dim ds As New DataSet

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim ctr As Integer = 0
            SQLAdapter.Fill(ds, "LoanDetails")

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getClearanceDetails()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return ds
    End Function

    Public Sub saveSettlementBalance(ByVal CustomerID As String, ByVal ProductID As Integer, ByVal Amount As Decimal, ByVal RequestedDate As Date, ByVal CapturedBy As String, ByVal CapturedDate As Date, ByVal CapturedTime As Date, ByVal Status As String, ByVal BatchName As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try
            Dim qryinsert As String = "Insert into SettlementBalances values ('" & CustomerID & "','" & ProductID & "','" & Amount & "','" & RequestedDate & "','" & CapturedBy & "','" & CapturedDate & "','" & CapturedTime & "','" & Status & "','" & BatchName & "')"
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
            MessageBox.Show(ex.Message, "saveSettlementBalance()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "saveSettlementBalance()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub updateClearanceLetterDesign(ByVal ToWhom As String, ByVal Re As String, ByVal OfID As String, ByVal ThisServes As String, ByVal HasFully As String, ByVal Therefore As String, ByVal ForMoreInfor As String, ByVal Yours As String, ByVal JobDescription As String, ByVal AutoID As Integer)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update ClearanceDetails set ToWhom = '" & ToWhom & "', Re = '" & Re & "', OfID = '" & OfID & "', ThisServes = '" & ThisServes & "', HasFully = '" & HasFully & "', Therefore = '" & Therefore & "', ForMoreInfor = '" & ForMoreInfor & "', Yours = '" & Yours & "', JobDescription = '" & JobDescription & "' where AutoID = '" & AutoID & "'"
            Dim cmdinsert As New SqlClient.SqlCommand(qryinsert, con)
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            transactioninsert = con.BeginTransaction()
            cmdinsert.Transaction = transactioninsert

            If cmdinsert.ExecuteNonQuery() Then
                transactioninsert.Commit()
            End If
            MessageBox.Show("Done Saving updates", "updateClearanceLetterDesign()", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "updateClearanceLetterDesign()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateClearanceLetterDesign()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Function getTotalBalance(ByVal customerid As String, ByVal BatchName As String) As Decimal
        Dim qry As String = "select SUM(Amount) from SettlementBalances where CustomerID = '" & customerid & "' and BatchName = '" & BatchName & "'"

        Dim perc As Decimal = 0

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = 0
                Else
                    perc = Decimal.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
                End If
            Else
                perc = 0
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getTotalBalance()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getTotalBalance()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getBatchesForCustomer(ByVal customerid As String) As DataSet
        Dim qry As String = "select BatchName from SettlementBalances inner join Products on SettlementBalances.ProductID = Products.AutoID where CustomerID = '" & customerid & "' order by SettlementBalances.AutoID Desc"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getBatchesForCustomer()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getBatchesForCustomer()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return ds
    End Function

    Public Function getInstalment(ByVal customerid As String, ByVal productid As Integer) As Decimal
        Dim qry As String = "select Amount from ExpectedDeductionsPerProduct where CustomerID = '" & customerid & "' and ProductID = '" & productid & "' order by AutoID DESC"

        Dim perc As Decimal = 0

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = 0
                Else
                    perc = Decimal.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
                End If
            Else
                perc = 0
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getInstalment()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getInstalment()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getAllTotalReceipts(ByVal custid As String, ByVal edate As Date, ByVal con As SqlClient.SqlConnection) As Decimal
        Dim qry As String = "select Sum(Amount) from CustomerLedgerEntry where TransactionDescription not in ('Billing') and CustomerID = '" & custid & "' and TransactionDate <= '" & edate & "'"

        Dim perc As Decimal = 0

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "CustomerLedgerEntry")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) Then
                    perc = 0
                Else
                    perc = Decimal.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
                End If
            Else
                perc = 0
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getAllTotalReceipts()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getAllTotalReceipts()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getAllTotalRechargesAsReceipts(ByVal custid As String, ByVal edate As Date, ByVal con As SqlClient.SqlConnection) As Decimal
        Dim qry As String = "select Sum(FreebieAmount) from Recharges where DealNo in (select DealNo from CustomerDeals where Account = '" & custid & "' and DealStatus not in ('Declined','New','Requested','Received')) and RechargeDate <= '" & edate & "'"

        Dim perc As Decimal = 0

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "Recharges")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) Then
                    perc = 0
                Else
                    perc = Decimal.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
                End If
            Else
                perc = 0
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getAllTotalRechargesAsReceipts()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getAllTotalRechargesAsReceipts()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getTotalContractualAmount(ByVal custid As String, ByVal con As SqlClient.SqlConnection) As Decimal
        Dim qry As String = "select CustomerDeals.DTCode,DealAmount from CustomerDeals INNER JOIN DealTemplate ON CustomerDeals.DTCode = DealTemplate.DTCode where Account = '" & custid & "' and DealStatus not in ('Declined','New','Requested','Received','Default','Cancelled')"

        Dim perc As Decimal = 0

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "DealTemplate")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) Then
                    perc = 0
                Else
                    perc = 0
                    For index As Integer = 0 To ds.Tables(0).Rows.Count - 1
                        perc = perc + Decimal.Parse(ds.Tables(0).Rows(index).Item(1).ToString)
                    Next
                    perc = perc * 24
                End If
            Else
                perc = 0
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getTotalContractualAmount()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getTotalContractualAmount()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getAllTotalBilling(ByVal custid As String, ByVal edate As Date, ByVal con As SqlClient.SqlConnection) As Decimal
        Dim qry As String = "select Sum(Amount) from CustomerLedgerEntry where TransactionDescription = 'Billing' and CustomerID = '" & custid & "' and TransactionDate <= '" & edate & "'"

        Dim perc As Decimal = 0

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "CustomerLedgerEntry")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) Then
                    perc = 0
                Else
                    perc = Decimal.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
                End If
            Else
                perc = 0
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getAllTotalBilling()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getAllTotalBilling()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getPendingAllocations(ByVal custid As String, ByVal edate As Date, ByVal con As SqlClient.SqlConnection) As Decimal
        Dim qry As String = "select SUM(Amount) from CustomerReceipts where CustomerID = '" & custid & "' and Status = 'New' and TransactionDate <= '" & edate & "'"

        Dim perc As Decimal = 0

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "CustomerLedgerEntry")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) Then
                    perc = 0
                Else
                    perc = Decimal.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
                End If
            Else
                perc = 0
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getPendingAllocations()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getPendingAllocations()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getAllPaidBilling(ByVal custid As String, ByVal edate As Date, ByVal con As SqlClient.SqlConnection) As Decimal
        Dim qry As String = "select Sum(Amount) from CustomerLedgerEntry where TransactionDescription = 'Billing' and CustomerID = '" & custid & "' and Status = 'Paid' and TransactionDate <= '" & edate & "'"

        Dim perc As Decimal = 0

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "CustomerLedgerEntry")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) Then
                    perc = 0
                Else
                    perc = Decimal.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
                End If
            Else
                perc = 0
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getAllPaidBilling()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getAllPaidBilling()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function checkCustomerID(ByVal custid As String, ByVal con As SqlClient.SqlConnection) As Integer
        Dim qry As String = "select * from CustomerMaster where  Account = '" & custid & "'"

        Dim perc As Integer = 0

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "CustomerMaster")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = ds.Tables(0).Rows.Count
            Else
                perc = 0
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "checkCustomerID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "checkCustomerID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getBalancesFromRM(ByVal customerid As String, ByVal sdate As Date, ByVal con As SqlClient.SqlConnection) As Decimal
        Dim Receipts, TotalRechargesAsReceipts, TotalContractualAmount, TotalBilling, PendingAllocations, TotalReceipts, PaidBills, balance As Decimal

        If checkCustomerID(customerid, con) = 1 Then
            Receipts = -getAllTotalReceipts(customerid, getLastDateOfMonth(sdate), con)
            TotalRechargesAsReceipts = getAllTotalRechargesAsReceipts(customerid, getLastDateOfMonth(sdate), con)
            TotalContractualAmount = getTotalContractualAmount(customerid, con)
            TotalBilling = getAllTotalBilling(customerid, getLastDateOfMonth(sdate), con)
            PendingAllocations = getPendingAllocations(customerid, getLastDateOfMonth(sdate), con)
            TotalReceipts = Receipts + PendingAllocations
            PaidBills = getAllPaidBilling(customerid, getLastDateOfMonth(sdate), con)
            TotalRechargesAsReceipts = TotalRechargesAsReceipts + PaidBills

            If TotalReceipts = TotalContractualAmount Then
                If TotalRechargesAsReceipts = TotalContractualAmount Then
                    If TotalReceipts = TotalRechargesAsReceipts Then
                        'RichTextBox1.Text = "1"
                        balance = TotalContractualAmount - TotalReceipts
                        'Refund = 0
                        If TotalBilling < TotalReceipts Then
                            If TotalRechargesAsReceipts < TotalReceipts Then
                                If TotalBilling < TotalRechargesAsReceipts Then
                                    'Refund = TotalReceipts - TotalRechargesAsReceipts
                                ElseIf TotalBilling > TotalRechargesAsReceipts Then
                                    'Refund = TotalReceipts - TotalBilling
                                End If
                            End If
                        End If
                    ElseIf TotalReceipts < TotalRechargesAsReceipts Then
                        'No Applicable
                        'RichTextBox1.Text = "2"
                        balance = TotalContractualAmount - TotalReceipts
                        'Refund = 0
                        If TotalBilling < TotalReceipts Then
                            If TotalRechargesAsReceipts < TotalReceipts Then
                                If TotalBilling < TotalRechargesAsReceipts Then
                                    ' Refund = TotalReceipts - TotalRechargesAsReceipts
                                ElseIf TotalBilling > TotalRechargesAsReceipts Then
                                    ' Refund = TotalReceipts - TotalBilling
                                End If
                            End If
                        End If
                    ElseIf TotalReceipts > TotalRechargesAsReceipts Then
                        'No Applicable
                        'RichTextBox1.Text = "3"
                        balance = TotalContractualAmount - TotalReceipts
                        'Refund = 0
                        If TotalBilling < TotalReceipts Then
                            If TotalRechargesAsReceipts < TotalReceipts Then
                                If TotalBilling < TotalRechargesAsReceipts Then
                                    'Refund = TotalReceipts - TotalRechargesAsReceipts
                                ElseIf TotalBilling > TotalRechargesAsReceipts Then
                                    'Refund = TotalReceipts - TotalBilling
                                End If
                            End If
                        End If
                    End If
                ElseIf TotalRechargesAsReceipts < TotalContractualAmount Then
                    If TotalReceipts = TotalRechargesAsReceipts Then
                        'Not Applicable
                        'RichTextBox1.Text = "4"
                        balance = TotalContractualAmount - TotalReceipts
                        'Refund = 0
                        If TotalBilling < TotalReceipts Then
                            If TotalRechargesAsReceipts < TotalReceipts Then
                                If TotalBilling < TotalRechargesAsReceipts Then
                                    'Refund = TotalReceipts - TotalRechargesAsReceipts
                                ElseIf TotalBilling > TotalRechargesAsReceipts Then
                                    ' Refund = TotalReceipts - TotalBilling
                                End If
                            End If
                        End If
                    ElseIf TotalReceipts < TotalRechargesAsReceipts Then
                        'Not Applicable
                        'RichTextBox1.Text = "5"
                        balance = TotalContractualAmount - TotalReceipts
                        'Refund = 0
                        If TotalBilling < TotalReceipts Then
                            If TotalRechargesAsReceipts < TotalReceipts Then
                                If TotalBilling < TotalRechargesAsReceipts Then
                                    'Refund = TotalReceipts - TotalRechargesAsReceipts
                                ElseIf TotalBilling > TotalRechargesAsReceipts Then
                                    'Refund = TotalReceipts - TotalBilling
                                End If
                            End If
                        End If
                    ElseIf TotalReceipts > TotalRechargesAsReceipts Then
                        'RichTextBox1.Text = "6. Uder-recharged customer"
                        'tbunderrecharge.Text = TotalReceipts - TotalRechargesAsReceipts
                        balance = TotalContractualAmount - TotalReceipts
                        'Refund = 0
                        If TotalBilling < TotalReceipts Then
                            If TotalRechargesAsReceipts < TotalReceipts Then
                                If TotalBilling < TotalRechargesAsReceipts Then
                                    'Refund = TotalReceipts - TotalRechargesAsReceipts
                                ElseIf TotalBilling > TotalRechargesAsReceipts Then
                                    ' Refund = TotalReceipts - TotalBilling
                                End If
                            End If
                        End If
                    End If
                ElseIf TotalRechargesAsReceipts > TotalContractualAmount Then
                    If TotalReceipts = TotalRechargesAsReceipts Then
                        'Not Applicable
                        'RichTextBox1.Text = "7"
                        balance = TotalContractualAmount - TotalReceipts
                        'Refund = 0
                        If TotalBilling < TotalReceipts Then
                            If TotalRechargesAsReceipts < TotalReceipts Then
                                If TotalBilling < TotalRechargesAsReceipts Then
                                    'Refund = TotalReceipts - TotalRechargesAsReceipts
                                ElseIf TotalBilling > TotalRechargesAsReceipts Then
                                    ' Refund = TotalReceipts - TotalBilling
                                End If
                            End If
                        End If
                    ElseIf TotalReceipts < TotalRechargesAsReceipts Then
                        'RichTextBox1.Text = "8. Over-recharged customer"
                        'tboverrecharge.Text = TotalRechargesAsReceipts - TotalReceipts
                        balance = TotalRechargesAsReceipts - TotalReceipts
                        'Refund = 0
                        If TotalBilling < TotalReceipts Then
                            If TotalRechargesAsReceipts < TotalReceipts Then
                                If TotalBilling < TotalRechargesAsReceipts Then
                                    'Refund = TotalReceipts - TotalRechargesAsReceipts
                                ElseIf TotalBilling > TotalRechargesAsReceipts Then
                                    'Refund = TotalReceipts - TotalBilling
                                End If
                            End If
                        End If
                    ElseIf TotalReceipts > TotalRechargesAsReceipts Then
                        'Not Applicable
                        'RichTextBox1.Text = "9"
                        balance = TotalContractualAmount - TotalReceipts
                        'Refund = 0
                        If TotalBilling < TotalReceipts Then
                            If TotalRechargesAsReceipts < TotalReceipts Then
                                If TotalBilling < TotalRechargesAsReceipts Then
                                    'Refund = TotalReceipts - TotalRechargesAsReceipts
                                ElseIf TotalBilling > TotalRechargesAsReceipts Then
                                    'Refund = TotalReceipts - TotalBilling
                                End If
                            End If
                        End If
                    End If
                End If
            ElseIf TotalReceipts < TotalContractualAmount Then
                If TotalRechargesAsReceipts = TotalContractualAmount Then
                    If TotalReceipts = TotalRechargesAsReceipts Then
                        'Not Applicable
                        'RichTextBox1.Text = "10"
                        balance = TotalContractualAmount - TotalReceipts
                        'Refund = 0
                        If TotalBilling < TotalReceipts Then
                            If TotalRechargesAsReceipts < TotalReceipts Then
                                If TotalBilling < TotalRechargesAsReceipts Then
                                    'Refund = TotalReceipts - TotalRechargesAsReceipts
                                ElseIf TotalBilling > TotalRechargesAsReceipts Then
                                    'Refund = TotalReceipts - TotalBilling
                                End If
                            End If
                        End If
                    ElseIf TotalReceipts < TotalRechargesAsReceipts Then
                        'RichTextBox1.Text = "11. Over-recharged customer"
                        'tboverrecharge.Text = TotalRechargesAsReceipts - TotalReceipts
                        balance = TotalContractualAmount - TotalReceipts
                        'Refund = 0
                        If TotalBilling < TotalReceipts Then
                            If TotalRechargesAsReceipts < TotalReceipts Then
                                If TotalBilling < TotalRechargesAsReceipts Then
                                    'Refund = TotalReceipts - TotalRechargesAsReceipts
                                ElseIf TotalBilling > TotalRechargesAsReceipts Then
                                    'Refund = TotalReceipts - TotalBilling
                                End If
                            End If
                        End If
                    ElseIf TotalReceipts > TotalRechargesAsReceipts Then
                        'Not Applicable
                        'RichTextBox1.Text = "12"
                        balance = TotalContractualAmount - TotalReceipts
                        'Refund = 0
                        If TotalBilling < TotalReceipts Then
                            If TotalRechargesAsReceipts < TotalReceipts Then
                                If TotalBilling < TotalRechargesAsReceipts Then
                                    'Refund = TotalReceipts - TotalRechargesAsReceipts
                                ElseIf TotalBilling > TotalRechargesAsReceipts Then
                                    ' Refund = TotalReceipts - TotalBilling
                                End If
                            End If
                        End If
                    End If
                ElseIf TotalRechargesAsReceipts < TotalContractualAmount Then
                    If TotalReceipts = TotalRechargesAsReceipts Then
                        'RichTextBox1.Text = "13"
                        balance = TotalContractualAmount - TotalReceipts
                        'Refund = 0
                        If TotalBilling < TotalReceipts Then
                            If TotalRechargesAsReceipts < TotalReceipts Then
                                If TotalBilling < TotalRechargesAsReceipts Then
                                    ' Refund = TotalReceipts - TotalRechargesAsReceipts
                                ElseIf TotalBilling > TotalRechargesAsReceipts Then
                                    ' Refund = TotalReceipts - TotalBilling
                                End If
                            End If
                        End If
                    ElseIf TotalReceipts < TotalRechargesAsReceipts Then
                        'RichTextBox1.Text = "14. Over-recharged customer"
                        'tboverrecharge.Text = TotalRechargesAsReceipts - TotalReceipts
                        balance = TotalContractualAmount - TotalReceipts
                        'Refund = 0
                        If TotalBilling < TotalReceipts Then
                            If TotalRechargesAsReceipts < TotalReceipts Then
                                If TotalBilling < TotalRechargesAsReceipts Then
                                    ' Refund = TotalReceipts - TotalRechargesAsReceipts
                                ElseIf TotalBilling > TotalRechargesAsReceipts Then
                                    ' Refund = TotalReceipts - TotalBilling
                                End If
                            End If
                        End If
                    ElseIf TotalReceipts > TotalRechargesAsReceipts Then
                        'RichTextBox1.Text = "15. Uder-recharged customer"
                        'tbunderrecharge.Text = TotalReceipts - TotalRechargesAsReceipts
                        balance = TotalContractualAmount - TotalReceipts
                        'Refund = 0
                        If TotalBilling < TotalReceipts Then
                            If TotalRechargesAsReceipts < TotalReceipts Then
                                If TotalBilling < TotalRechargesAsReceipts Then
                                    ' Refund = TotalReceipts - TotalRechargesAsReceipts
                                ElseIf TotalBilling > TotalRechargesAsReceipts Then
                                    ' Refund = TotalReceipts - TotalBilling
                                End If
                            End If
                        End If
                    End If
                ElseIf TotalRechargesAsReceipts > TotalContractualAmount Then
                    If TotalReceipts = TotalRechargesAsReceipts Then
                        'Not Applicable
                        'RichTextBox1.Text = "16"
                        balance = TotalContractualAmount - TotalReceipts
                        'Refund = 0
                        If TotalBilling < TotalReceipts Then
                            If TotalRechargesAsReceipts < TotalReceipts Then
                                If TotalBilling < TotalRechargesAsReceipts Then
                                    'Refund = TotalReceipts - TotalRechargesAsReceipts
                                ElseIf TotalBilling > TotalRechargesAsReceipts Then
                                    'Refund = TotalReceipts - TotalBilling
                                End If
                            End If
                        End If
                    ElseIf TotalReceipts < TotalRechargesAsReceipts Then
                        'RichTextBox1.Text = "17. Over-recharged customer"
                        'tboverrecharge.Text = TotalRechargesAsReceipts - TotalReceipts
                        balance = TotalRechargesAsReceipts - TotalReceipts
                        'Refund = 0
                        If TotalBilling < TotalReceipts Then
                            If TotalRechargesAsReceipts < TotalReceipts Then
                                If TotalBilling < TotalRechargesAsReceipts Then
                                    ' Refund = TotalReceipts - TotalRechargesAsReceipts
                                ElseIf TotalBilling > TotalRechargesAsReceipts Then
                                    ' Refund = TotalReceipts - TotalBilling
                                End If
                            End If
                        End If
                    ElseIf TotalReceipts > TotalRechargesAsReceipts Then
                        'No Applicable
                        'RichTextBox1.Text = "18"
                        balance = TotalContractualAmount - TotalReceipts
                        'Refund = 0
                        If TotalBilling < TotalReceipts Then
                            If TotalRechargesAsReceipts < TotalReceipts Then
                                If TotalBilling < TotalRechargesAsReceipts Then
                                    'Refund = TotalReceipts - TotalRechargesAsReceipts
                                ElseIf TotalBilling > TotalRechargesAsReceipts Then
                                    'Refund = TotalReceipts - TotalBilling
                                End If
                            End If
                        End If
                    End If
                End If
            ElseIf TotalReceipts > TotalContractualAmount Then
                If TotalRechargesAsReceipts = TotalContractualAmount Then
                    If TotalReceipts = TotalRechargesAsReceipts Then
                        'Not Applicable
                        'RichTextBox1.Text = "19"
                        balance = TotalContractualAmount - TotalReceipts
                        'Refund = 0
                        If TotalBilling < TotalReceipts Then
                            If TotalRechargesAsReceipts < TotalReceipts Then
                                If TotalBilling < TotalRechargesAsReceipts Then
                                    'Refund = TotalReceipts - TotalRechargesAsReceipts
                                ElseIf TotalBilling > TotalRechargesAsReceipts Then
                                    'Refund = TotalReceipts - TotalBilling
                                End If
                            End If
                        End If
                    ElseIf TotalReceipts < TotalRechargesAsReceipts Then
                        'Not Applicable
                        'RichTextBox1.Text = "20"
                        balance = TotalContractualAmount - TotalReceipts
                        'Refund = 0
                        If TotalBilling < TotalReceipts Then
                            If TotalRechargesAsReceipts < TotalReceipts Then
                                If TotalBilling < TotalRechargesAsReceipts Then
                                    ' Refund = TotalReceipts - TotalRechargesAsReceipts
                                ElseIf TotalBilling > TotalRechargesAsReceipts Then
                                    'Refund = TotalReceipts - TotalBilling
                                End If
                            End If
                        End If
                    ElseIf TotalReceipts > TotalRechargesAsReceipts Then
                        'RichTextBox1.Text = "21. Refund = " & TotalReceipts - TotalRechargesAsReceipts
                        balance = 0
                        'Refund = TotalReceipts - TotalRechargesAsReceipts
                    End If
                ElseIf TotalRechargesAsReceipts < TotalContractualAmount Then
                    If TotalReceipts = TotalRechargesAsReceipts Then
                        'Not Applicable
                        'RichTextBox1.Text = "22"
                        balance = TotalContractualAmount - TotalReceipts
                        'Refund = 0
                        If TotalBilling < TotalReceipts Then
                            If TotalRechargesAsReceipts < TotalReceipts Then
                                If TotalBilling < TotalRechargesAsReceipts Then
                                    'Refund = TotalReceipts - TotalRechargesAsReceipts
                                ElseIf TotalBilling > TotalRechargesAsReceipts Then
                                    'Refund = TotalReceipts - TotalBilling
                                End If
                            End If
                        End If
                    ElseIf TotalReceipts < TotalRechargesAsReceipts Then
                        'Not Applicable
                        'RichTextBox1.Text = "23"
                        balance = TotalContractualAmount - TotalReceipts
                        'Refund = 0
                        If TotalBilling < TotalReceipts Then
                            If TotalRechargesAsReceipts < TotalReceipts Then
                                If TotalBilling < TotalRechargesAsReceipts Then
                                    ' Refund = TotalReceipts - TotalRechargesAsReceipts
                                ElseIf TotalBilling > TotalRechargesAsReceipts Then
                                    ' Refund = TotalReceipts - TotalBilling
                                End If
                            End If
                        End If
                    ElseIf TotalReceipts > TotalRechargesAsReceipts Then
                        'RichTextBox1.Text = "24. Uder-recharged customer: Refund = " & TotalContractualAmount - TotalReceipts
                        'tbunderrecharge.Text = TotalContractualAmount - TotalRechargesAsReceipts
                        balance = 0
                        'Refund = TotalReceipts - TotalContractualAmount
                    End If
                ElseIf TotalRechargesAsReceipts > TotalContractualAmount Then
                    If TotalReceipts = TotalRechargesAsReceipts Then
                        'RichTextBox1.Text = "25"
                        balance = 0
                        'Refund = 0
                        If TotalBilling < TotalReceipts Then
                            If TotalRechargesAsReceipts < TotalReceipts Then
                                If TotalBilling < TotalRechargesAsReceipts Then
                                    'Refund = TotalReceipts - TotalRechargesAsReceipts
                                ElseIf TotalBilling > TotalRechargesAsReceipts Then
                                    'Refund = TotalReceipts - TotalBilling
                                End If
                            End If
                        End If
                    ElseIf TotalReceipts < TotalRechargesAsReceipts Then
                        'RichTextBox1.Text = "26. Over-recharged customer"
                        'tboverrecharge.Text = TotalRechargesAsReceipts - TotalReceipts
                        balance = TotalRechargesAsReceipts - TotalReceipts
                        'Refund = 0
                        If TotalBilling < TotalReceipts Then
                            If TotalRechargesAsReceipts < TotalReceipts Then
                                If TotalBilling < TotalRechargesAsReceipts Then
                                    'Refund = TotalReceipts - TotalRechargesAsReceipts
                                ElseIf TotalBilling > TotalRechargesAsReceipts Then
                                    'Refund = TotalReceipts - TotalBilling
                                End If
                            End If
                        End If
                    ElseIf TotalReceipts > TotalRechargesAsReceipts Then
                        'RichTextBox1.Text = "27. Refund = " & TotalReceipts - TotalRechargesAsReceipts
                        balance = 0
                        'Refund = TotalReceipts - TotalRechargesAsReceipts
                    End If
                End If
            End If
        ElseIf checkCustomerID(customerid, con) = 0 Then
            balance = 0
        End If
        Return balance
    End Function

    Public Function getCount(ByVal ds As DataSet) As Integer
        If ds.Tables(0).Rows.Count > 0 Then
            If ds.Tables(0).Rows(0).IsNull(0) = False Then
                Return ds.Tables(0).Rows.Count
            Else
                Return 0
            End If
        Else
            Return 0
        End If
    End Function

    Public Function checkCustomer(ByVal custid As String, ByVal con As SqlClient.SqlConnection) As DataSet
        Dim qry As String = "select * from CustomerDetails where ID_Number = '" & custid & "'"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim ctr As Integer = 0
            SQLAdapter.Fill(ds, "CustomerDetails")

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "checkCustomer()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return ds
        'ds.Clear()
    End Function

    Public Function getTakeOnDate(ByVal custid As String, ByVal loantype As String, ByVal con As SqlClient.SqlConnection) As Date
        Dim qry As String = "select TOP 1 TransactionDate from LedgerEntry where CustomerID = '" & custid & "' and LoanType = '" & loantype & "' and Reference in ('TOB')"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            SQLAdapter.Fill(ds, "LedgerEntry")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    Return Date.Now
                Else
                    Return Date.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
                End If
            Else
                Return Date.Now
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getTakeOnDate()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getTakeOnDate()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try

    End Function

    Public Function getTakeOnBalance(ByVal custid As String, ByVal loantype As String, ByVal con As SqlClient.SqlConnection) As Decimal
        Dim qry As String = "select Amount from LedgerEntry where CustomerID = '" & custid & "' and LoanType = '" & loantype & "' and Reference in ('TOB')"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            SQLAdapter.Fill(ds, "LedgerEntry")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    Return 0.0
                Else
                    Return Decimal.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
                End If
            Else
                Return 0
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getTakeOnBalance()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getTakeOnBalance()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Function

    Public Function getAllLoanIDsInLedger(ByVal customerid As String, ByVal LoanType As String, ByVal con As SqlClient.SqlConnection) As DataSet
        Dim qry As String = "select LoanID from LedgerEntry where CustomerID = '" & customerid & "' and Reference = 'P' and LoanType = '" & LoanType & "' order by LoanID"
        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim ctr As Integer = 0
            SQLAdapter.Fill(ds, "LoanDetails")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getAllLoanIDsInLedger()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getAllLoanIDsInLedger()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            'ds.Clear()
            con.Close()
        End Try
        Return ds
        'ds.Clear()
    End Function

    Public Function getLastActiveLoanID(ByVal custid As String, ByVal loantype As String, ByVal con As SqlClient.SqlConnection) As Integer
        Dim qry As String = "select TOP 1 AutoLoanID from LoanDetails where Customer_ID = '" & custid & "' and LoanDescription = '" & loantype & "' and LoanStatus in ('Active') order by AutoLoanID DESC"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim ctr As Integer = 0
            SQLAdapter.Fill(ds, "LoanDetails")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    Return 0
                Else
                    Return Integer.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
                End If
            Else
                Return 0
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getLastActiveLoanID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getLastActiveLoanID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            'ds.Clear()
            con.Close()
        End Try
    End Function

    Public Function getSecondActiveLoanID(ByVal custid As String, ByVal loantype As String, ByVal con As SqlClient.SqlConnection) As Integer
        Dim qry As String = "select Top 1 AutoLoanID from LoanDetails where Customer_ID = '" & custid & "' and LoanDescription = '" & loantype & "' and LoanStatus in ('Active') order by AutoLoanID DESC"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim ctr As Integer = 0
            SQLAdapter.Fill(ds, "LoanDetails")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 1 Then
                Return Integer.Parse(ds.Tables(0).Rows(1).Item(0).ToString)
            Else
                Return 0
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getSecondActiveLoanID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getSecondActiveLoanID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            'ds.Clear()
            con.Close()
        End Try
    End Function

    Public Function customerAvailable(ByVal custid As String, ByVal con As SqlClient.SqlConnection) As Boolean
        Dim qry As String = "select * from CustomerMaster where Account = '" & custid & "'"

        Dim perc As Boolean = False

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "CustomerMaster")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = True
            Else
                perc = False
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "customerAvailable()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "customerAvailable()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getCustomerNameRMLedger(ByVal custid As String, ByVal con As SqlClient.SqlConnection) As String
        Dim qry As String = "select FirstName + ' ' + Surname from CustomerMaster where Account = '" & custid & "'"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "CustomerMaster")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = ds.Tables(0).Rows(0).Item(0).ToString
            Else
                perc = ""
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getCustomerNameRMLedger()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getCustomerNameRMLedger()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getCompanyName(ByVal con As SqlClient.SqlConnection) As String
        Dim qry As String = "select CompanyName from CompanyDetails"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "DealTemplate")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) Then
                    perc = ""
                Else
                    perc = ds.Tables(0).Rows(0).Item(0).ToString
                End If
            Else
                perc = ""
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getCompanyName()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getCompanyName()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getOpeningBalanceDate(ByVal custid As String, ByVal con As SqlClient.SqlConnection) As Date
        Dim qry As String = "select Top 1 TransactionDate from CustomerLedgerEntry where CustomerID = '" & custid & "' order by TransactionDate ASC"

        Dim perc As Date

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) Then
                    perc = Date.Now
                Else
                    perc = Date.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
                End If
            Else
                perc = Date.Now
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getOpeningBalanceDate()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getOpeningBalanceDate()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getLatestLedgerDate(ByVal con As SqlClient.SqlConnection) As Date
        Dim qry As String = "select Top 1 TransactionDate from CustomerLedgerEntry order by TransactionDate ASC"

        Dim perc As Date

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "PaymentDates")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = Date.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
            Else
                perc = Date.Now
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getLatestLedgerDate()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getLatestLedgerDate()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getAllCustomerBalance(ByVal custid As String, ByVal con As SqlClient.SqlConnection) As Decimal
        Dim str As Decimal

        Dim qry As String = "select Sum(Amount) from LedgerEntry where CustomerID = '" & custid & "'"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim ctr As Integer = 0
            SQLAdapter.Fill(ds, "LedgerEntry")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    str = 0
                Else
                    str = Decimal.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
                End If
            Else
                str = 0
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getAllCustomerBalance()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getAllCustomerBalance()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            'ds.Clear()
            con.Close()
        End Try

        Return str
    End Function

    Public Sub saveStoppedRefunds(ByVal RefundID As String, ByVal Amount As Decimal, ByVal DepositDate As Date, ByVal Status As String, ByVal SystemName As String, ByVal CustomerID As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "Insert into StoppedRefunds values ('" & RefundID & "','" & Amount & "','" & DepositDate & "','" & Status & "','" & SystemName & "','" & CustomerID & "')"
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
            MessageBox.Show(ex.Message, "saveStoppedRefunds()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "saveStoppedRefunds()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Function checkStoppedRefund(ByVal refundid As String, ByVal amount As Decimal, ByVal systemname As String, ByVal customerid As String) As Boolean
        Dim qry As String = "select * from StoppedRefunds where RefundID = '" & refundid & "' and Amount = '" & amount & "' and Status = 'Stopped' and SystemName = '" & systemname & "' and CustomerID = '" & customerid & "'"

        Dim perc As Boolean = False

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = True
            Else
                perc = False
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "checkStoppedRefund()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "checkStoppedRefund()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Sub deleteCustomer(ByVal AutoID As Integer)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "delete from CustomerDetails where AutoCustID = '" & AutoID & "'"
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
            MessageBox.Show(ex.Message, "deleteCustomer()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "deleteCustomer()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Function getSuggestedDeduction(ByVal refundid As Integer, ByVal con As SqlClient.SqlConnection) As Decimal
        Dim str As Decimal

        Dim qry As String = "select SuggestedDeduction from CustomerRefundsStoped where RefundID = '" & refundid & "'"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim ctr As Integer = 0
            SQLAdapter.Fill(ds, "LedgerEntry")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    str = 0
                Else
                    str = Decimal.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
                End If
            Else
                str = 0
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getSuggestedDeduction()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getSuggestedDeduction()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            'ds.Clear()
            con.Close()
        End Try

        Return str
    End Function

    Public Function CheckIfCustomerPaysForPeopleInFPM(ByVal payerid As String, ByVal con As SqlClient.SqlConnection) As Boolean
        Dim qry As String = "select * from PolicyCovers where Status = 'Active' and PolicyNumber in (select PolicyNumber from Policy where Status = 'Active' and PayerID = '" & payerid & "' and PayerID not in (select PayerID from BenefitScheduleMessages where PayerID = '" & payerid & "'))"

        Dim perc As Boolean = False

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "SalesAllocation")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = False
                Else
                    perc = True
                End If
            Else
                perc = False
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "CheckIfCustomerPaysForPeopleInFPM()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "CheckIfCustomerPaysForPeopleInFPM()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            'ds.Clear()
            con.Close()
        End Try
        Return perc
    End Function

    Public Sub geleteCustomerDetails(ByVal AutoID As Integer, ByVal CustomerID As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try
            Dim qryinsert As String = "delete from CustomerDetails where AutoCustID = '" & AutoID & "' and CustomerID = '" & CustomerID & "'"
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
            MessageBox.Show(ex.Message, " geleteCustomerDetails()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, " geleteCustomerDetails()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Function getEmployerName(ByVal customerid As String) As String
        Dim qry As String = "select EmployerName from Employers where AutoID in (select EmployerID from CustomerDetails where CustomerID = '" & customerid & "')"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "DealTemplate")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) Then
                    perc = ""
                Else
                    perc = ds.Tables(0).Rows(0).Item(0).ToString
                End If
            Else
                perc = ""
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getEmployerName()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getEmployerName()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Sub saveUserType(ByVal RefundID As String, ByVal Amount As Decimal, ByVal DepositDate As Date, ByVal Status As String, ByVal SystemName As String, ByVal CustomerID As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "Insert into StoppedRefunds values ('" & RefundID & "','" & Amount & "','" & DepositDate & "','" & Status & "','" & SystemName & "','" & CustomerID & "')"
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
            MessageBox.Show(ex.Message, "saveUserType()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "saveUserType()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Function checkSettlementDetails() As Boolean
        Dim qry As String = "select TOP 1 * from SettlementDetails"

        Dim perc As Boolean = False

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "SalesAllocation")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = False
                Else
                    perc = True
                End If
            Else
                perc = False
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "checkSettlementDetails()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "checkSettlementDetails()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            'ds.Clear()
            con.Close()
        End Try
        Return perc
    End Function

    Public Sub saveSettlementLetterDesignDetails(ByVal ToWhom As String, ByVal Re As String, ByVal OfID As String, ByVal ThisServes As String, ByVal HasAccount As String, ByVal PleaseBe As String, ByVal IfYouIntend As String, ByVal ThitoAccount As String, ByVal ThereAfter As String, ByVal ForMoreInfo As String, ByVal Yours As String, ByVal SupervisorName As String, ByVal JobDescription As String, ByVal EmailAddress As String, ByVal PleaseBeInformed As String, ByVal CutOfDate As Integer, ByVal EarlyPayPercentage As Decimal, ByVal PleaseEnsure As String, ByVal ToAvoid As String, ByVal ValidityDate As Integer)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            transactioninsert = con.BeginTransaction()
            Dim qryinsert As String = "Insert into SettlementDetails values ('" & ToWhom & "','" & Re & "','" & OfID & "','" & ThisServes & "','" & HasAccount & "','" & PleaseBe & "','" & IfYouIntend & "','" & ThitoAccount & "','" & ThereAfter & "','" & ForMoreInfo & "','" & Yours & "','" & SupervisorName & "','" & JobDescription & "','" & EmailAddress & "','" & PleaseBeInformed & "','" & CutOfDate & "','" & EarlyPayPercentage & "','" & PleaseEnsure & "','" & ToAvoid & "','" & ValidityDate & "')"
            Dim cmdinsert As New SqlClient.SqlCommand(qryinsert, con)
            cmdinsert.Transaction = transactioninsert
            If cmdinsert.ExecuteNonQuery() Then
                transactioninsert.Commit()
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "saveSettlementLetterDesignDetails()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "saveSettlementLetterDesignDetails()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Function checkClearanceDetails() As Boolean
        Dim qry As String = "select TOP 1 * from ClearanceDetails"

        Dim perc As Boolean = False

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "SalesAllocation")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = False
                Else
                    perc = True
                End If
            Else
                perc = False
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "checkClearanceDetails()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "checkClearanceDetails()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            'ds.Clear()
            con.Close()
        End Try
        Return perc
    End Function

    Public Sub saveClearanceLetterDesignDetails(ByVal ToWhom As String, ByVal Re As String, ByVal OfID As String, ByVal ThisServes As String, ByVal HasFully As String, ByVal Therefore As String, ByVal ForMoreInfor As String, ByVal Yours As String, ByVal JobDescription As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            transactioninsert = con.BeginTransaction()
            Dim qryinsert As String = "Insert into ClearanceDetails values ('" & ToWhom & "','" & Re & "','" & OfID & "','" & ThisServes & "','" & HasFully & "','" & Therefore & "','" & ForMoreInfor & "','" & Yours & "','" & JobDescription & "')"
            Dim cmdinsert As New SqlClient.SqlCommand(qryinsert, con)
            cmdinsert.Transaction = transactioninsert
            If cmdinsert.ExecuteNonQuery() Then
                transactioninsert.Commit()
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "saveClearanceLetterDesignDetails()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "saveClearanceLetterDesignDetails()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Function checkEmployers() As Boolean
        Dim qry As String = "select TOP 1 * from Employers"

        Dim perc As Boolean = False

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "SalesAllocation")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = False
                Else
                    perc = True
                End If
            Else
                perc = False
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "checkEmployers()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "checkEmployers()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            'ds.Clear()
            con.Close()
        End Try
        Return perc
    End Function

    Public Function checkProducts() As Boolean
        Dim qry As String = "select TOP 1 * from Products"

        Dim perc As Boolean = False

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "SalesAllocation")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = False
                Else
                    perc = True
                End If
            Else
                perc = False
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "checkProducts()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "checkProducts()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            'ds.Clear()
            con.Close()
        End Try
        Return perc
    End Function

    Public Function checkSupervisors() As Boolean
        Dim qry As String = "select TOP 1 * from Supervisors"

        Dim perc As Boolean = False

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "SalesAllocation")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = False
                Else
                    perc = True
                End If
            Else
                perc = False
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "checkSupervisors()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "checkSupervisors()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            'ds.Clear()
            con.Close()
        End Try
        Return perc
    End Function

    Public Sub saveSuperviors(ByVal SupervisorName As String, ByVal Title As String, ByVal EmailAddress As String, ByVal Department As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            transactioninsert = con.BeginTransaction()
            Dim qryinsert As String = "Insert into Supervisors values ('" & SupervisorName & "','" & Title & "','" & EmailAddress & "','" & Department & "')"
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
            MessageBox.Show(ex.Message, "saveSuperviors()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "saveSuperviors()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Function getCustomerDetailsForRefil(ByVal con As SqlClient.SqlConnection) As DataSet
        Dim qry As String = "select * from CustomerDetails"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getCustomerDetailsForRefil()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getCustomerDetailsForRefil()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return ds
    End Function

    Public Function getSettlementsForRefil(ByVal customerid As String, ByVal con As SqlClient.SqlConnection) As DataSet
        Dim qry As String = "select * from SettlementBalances where CustomerID = '" & customerid & "'"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getSettlementsForRefil()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getSettlementsForRefil()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return ds
    End Function

    Public Function getTerminationsAndAmendmentsForRefil(ByVal customerid As String, ByVal con As SqlClient.SqlConnection) As DataSet
        Dim qry As String = "select * from TerminationsAndAmendments where CustomerID = '" & customerid & "'"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getTerminationsAndAmendmentsForRefil()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getTerminationsAndAmendmentsForRefil()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return ds
    End Function

    Public Function getExpectedDeductionsForRefil(ByVal customerid As String, ByVal con As SqlClient.SqlConnection) As DataSet
        Dim qry As String = "select * from ExpectedDeductions where CustomerID = '" & customerid & "'"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getExpectedDeductionsForRefil()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getExpectedDeductionsForRefil()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return ds
    End Function

    Public Function getExpectedDeductionsPerProductForRefil(ByVal customerid As String, ByVal con As SqlClient.SqlConnection) As DataSet
        Dim qry As String = "select * from ExpectedDeductionsPerProduct where CustomerID = '" & customerid & "'"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getExpectedDeductionsPerProductForRefil()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getExpectedDeductionsPerProductForRefil()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return ds
    End Function

    Public Function getActualDeductionForRefil(ByVal customerid As String, ByVal con As SqlClient.SqlConnection) As DataSet
        Dim qry As String = "select * from ActualDeductions where CustomerID = '" & customerid & "'"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getActualDeductionForRefil()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getActualDeductionForRefil()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return ds
    End Function

    Public Sub updateChequeStatus(ByVal status As String, ByVal autoid As Integer, ByVal paymentdate As Date)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update ProductsCheques set Status = '" & status & "',PaymentDate = '" & paymentdate & "' where AutoID = '" & autoid & "'"
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
            MessageBox.Show(ex.Message, "updateChequeStatus()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateChequeStatus()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Function checkCutOfDate() As Boolean
        Dim qry As String = "select * from CutOfDate"

        Dim perc As Boolean = False

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "SalesAllocation")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = False
                Else
                    perc = True
                End If
            Else
                perc = False
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "checkCutOfDate()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "checkCutOfDate()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            'ds.Clear()
            con.Close()
        End Try
        Return perc
    End Function

    Public Function checkIfMonthHasCutOfDate(ByVal sdate As Date, ByVal edate As Date) As Boolean
        Dim qry As String = "select * from CutOfDate where SelectedDate between '" & sdate & "' and '" & edate & "' and Status = 'Active'"

        Dim perc As Boolean = False

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "SalesAllocation")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = False
                Else
                    perc = True
                End If
            Else
                perc = False
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "checkIfMonthHasCutOfDate()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "checkIfMonthHasCutOfDate()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            'ds.Clear()
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getCutOffDate() As Date
        Dim qry As String = "select SelectedDate from CutOfDate where  Status = 'Active'"

        Dim perc As Date

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "PaymentDates")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = Date.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
            Else
                perc = Date.Now
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getCutOffDate()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getCutOffDate()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Sub updateCutOffDateToClosed()
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update CutOfDate set Status = 'Closed' where Status = 'Active'"
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
            MessageBox.Show(ex.Message, "updateCutOffDateToClosed()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateCutOffDateToClosed()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub updateCustOffDateInCompanyDetails(ByVal cutoffdate As Integer)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update CompanyDetails set CutOfDate = '" & cutoffdate & "'"
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
            MessageBox.Show(ex.Message, "updateCustOffDateInCompanyDetails()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateCustOffDateInCompanyDetails()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub saveCutOffDate(ByVal CutOfDate As Integer, ByVal SelectedDate As Date, ByVal ConfirmedBy As String, ByVal ConfirmedTime As Date, ByVal Status As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            transactioninsert = con.BeginTransaction()
            Dim qryinsert As String = "Insert into CutOfDate values ('" & CutOfDate & "','" & SelectedDate & "','" & ConfirmedBy & "','" & ConfirmedTime & "','" & Status & "')"
            Dim cmdinsert As New SqlClient.SqlCommand(qryinsert, con)
            cmdinsert.Transaction = transactioninsert
            If cmdinsert.ExecuteNonQuery() Then
                transactioninsert.Commit()
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "saveCutOffDate()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "saveCutOffDate()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Function getPayerBalance(ByVal payerid As String, ByVal con As SqlClient.SqlConnection) As Decimal
        Dim str As Decimal

        Dim qry As String = "select SUM(Amount) from CustomerLedger where PayerID = '" & payerid & "'"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim ctr As Integer = 0
            SQLAdapter.Fill(ds, "LedgerEntry")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    str = 0
                Else
                    str = Decimal.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
                End If
            Else
                str = 0
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getPayerBalance()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getPayerBalance()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            'ds.Clear()
            con.Close()
        End Try

        Return str
    End Function

    Public Sub saveReceipt(ByVal PayerID As String, ByVal Amount As Decimal, ByVal ModeID As Integer, ByVal TransactionDate As Date, ByVal CapturedBy As String, ByVal CapturedDate As Date, ByVal CapturedTime As Date, ByVal ConfirmedBy As String, ByVal ConfirmedDate As Date, ByVal ConfirmedTime As Date, ByVal BatchName As String, ByVal Status As String, ByVal Reference As String, ByVal BankingBatch As String, ByVal ReceiptNo As String, ByVal CompanyBranchID As Integer, ByVal BookID As Integer)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try
            Dim qryinsert As String = "Insert into Receipts values ('" & PayerID & "','" & Amount & "','" & ModeID & "','" & TransactionDate & "','" & CapturedBy & "','" & CapturedDate & "','" & CapturedTime & "','" & ConfirmedBy & "','" & ConfirmedDate & "','" & ConfirmedTime & "','" & BatchName & "','" & Status & "','" & Reference & "','" & BankingBatch & "','" & ReceiptNo & "','" & CompanyBranchID & "','" & BookID & "')"
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
            MessageBox.Show(ex.Message, "saveReceipt()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "saveReceipt()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub saveReceiptBreakdown(ByVal PayerID As String, ByVal Amount As Decimal, ByVal TransactionDate As Date, ByVal ProductID As Integer, ByVal BatchName As String, ByVal Reference As String, ByVal Status As String, ByVal BankingBatch As String, ByVal ReceiptType As String, ByVal ExportedBy As String, ByVal ExportedTime As Date, ByVal ReceiptNo As String, ByVal ModeID As Integer, ByVal CompanyBranchID As Integer, ByVal BookID As Integer, ByVal Comment As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try
            Dim qryinsert As String = "Insert into ReceiptsBreakdown values ('" & PayerID & "','" & Amount & "','" & TransactionDate & "','" & ProductID & "','" & BatchName & "','" & Reference & "','" & Status & "','" & BankingBatch & "','" & ReceiptType & "','" & ExportedBy & "','" & ExportedTime & "','" & ReceiptNo & "','" & ModeID & "','" & CompanyBranchID & "','" & BookID & "','" & Comment & "')"
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
            MessageBox.Show(ex.Message, "saveReceiptBreakdown()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "saveReceiptBreakdown()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub updateReceiptBankingDetails(ByVal batchname As String, ByVal status As String, ByVal confirmeddate As Date, ByVal confirmedby As String, ByVal BankingBatch As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update Receipts set Status = '" & status & "',ConfirmedDate = '" & confirmeddate & "',ConfirmedTime = '" & confirmeddate & "',ConfirmedBy = '" & confirmedby & "',BankingBatch = '" & BankingBatch & "' where BatchName = '" & batchname & "'"
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
            MessageBox.Show(ex.Message, "updateReceiptBankingDetails()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateReceiptBankingDetails()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub updateReceiptBreakdownBankingDetails(ByVal batchname As String, ByVal status As String, ByVal BankingBatch As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update ReceiptsBreakdown set Status = '" & status & "',BankingBatch = '" & BankingBatch & "' where BatchName = '" & batchname & "'"
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
            MessageBox.Show(ex.Message, "updateReceiptBreakdownBankingDetails()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateReceiptBreakdownBankingDetails()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub updateReceiptReference(ByVal reference As String, ByVal BankingBatch As String, ByVal TransactionDate As Date)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update Receipts set Reference = '" & reference & "',Status = 'Banked',TransactionDate = '" & TransactionDate & "' where BankingBatch = '" & BankingBatch & "'"
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
            MessageBox.Show(ex.Message, "updateReceiptReference()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateReceiptReference()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub updateReceiptBreakdownReference(ByVal reference As String, ByVal BankingBatch As String, ByVal TransactionDate As Date)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update ReceiptsBreakdown set Reference = '" & reference & "',Status = 'Banked',TransactionDate = '" & TransactionDate & "' where BankingBatch = '" & BankingBatch & "'"
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
            MessageBox.Show(ex.Message, "updateReceiptBreakdownReference()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateReceiptBreakdownReference()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub updateReceiptStatusByBankingBatch(ByVal status As String, ByVal BankingBatch As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update Receipts set Status = '" & status & "' where BankingBatch = '" & BankingBatch & "'"
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
            MessageBox.Show(ex.Message, "updateReceiptStatusByBankingBatch()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateReceiptStatusByBankingBatch()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub updateReceiptBreakdownStatusByBankingBatch(ByVal status As String, ByVal BankingBatch As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update ReceiptsBreakdown set Status = '" & status & "' where BankingBatch = '" & BankingBatch & "'"
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
            MessageBox.Show(ex.Message, "updateReceiptBreakdownStatusByBankingBatch()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateReceiptBreakdownStatusByBankingBatch()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub updateReceiptStatusByBatchName(ByVal status As String, ByVal BatchName As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update Receipts set Status = '" & status & "' where BatchName = '" & BatchName & "'"
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
            MessageBox.Show(ex.Message, "updateReceiptStatusByBatchName()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateReceiptStatusByBatchName()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub updateReceiptBreakdownStatusByBatchName(ByVal status As String, ByVal BatchName As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update ReceiptsBreakdown set Status = '" & status & "' where BatchName = '" & BatchName & "'"
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
            MessageBox.Show(ex.Message, "updateReceiptBreakdownStatusByBatchName()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateReceiptBreakdownStatusByBatchName()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub updateReceiptBreakdownExportDetailsByAutoID(ByVal ExportedBy As String, ByVal ExportedTime As Date, ByVal AutoID As Integer)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update ReceiptsBreakdown set ExportedBy = '" & ExportedBy & "',ExportedTime = '" & ExportedTime & "' where AutoID = '" & AutoID & "'"
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
            MessageBox.Show(ex.Message, "updateReceiptBreakdownExportDetailsByAutoID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateReceiptBreakdownExportDetailsByAutoID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Function checkIfToProceedToConfirm(ByVal modeid As Integer) As String
        Dim qry As String = "select ProceedToConfirm from ModeOfPayment where ModeID = '" & modeid & "'"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "SalesAllocation")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = ""
                Else
                    perc = ds.Tables(0).Rows(0).Item(0).ToString
                End If
            Else
                perc = ""
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "checkIfToProceedToConfirm()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "checkIfToProceedToConfirm()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            'ds.Clear()
            con.Close()
        End Try
        Return perc
    End Function

    Public Sub saveUser(ByVal Username As String, ByVal PasswordValue As String, ByVal UserType As String, ByVal Name As String, ByVal SuperviorsID As Integer, ByVal CompanyBranchID As Integer)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try
            Dim qryinsert As String = "Insert into Users values ('" & Username & "','" & PasswordValue & "','" & UserType & "','" & Name & "','" & SuperviorsID & "','" & CompanyBranchID & "')"
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
            MessageBox.Show(ex.Message, "saveUser()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "saveUser()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub updateUser(ByVal Username As String, ByVal PasswordValue As String, ByVal UserType As String, ByVal Name As String, ByVal SuperviorsID As Integer, ByVal CompanyBranchID As Integer)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "Update Users set PasswordValue = '" & PasswordValue & "',UserType = '" & UserType & "',Name = '" & Name & "',SuperviorsID = '" & SuperviorsID & "',CompanyBranchID = '" & CompanyBranchID & "' where Username = '" & Username & "'"
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
            MessageBox.Show(ex.Message, "updateUser()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateUser()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Function getModeOfPaymentByModeID(ByVal modeid As Integer) As String
        Dim qry As String = "select ModeName from ModeOfPayment where ModeID = '" & modeid & "'"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "SalesAllocation")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = ""
                Else
                    perc = ds.Tables(0).Rows(0).Item(0).ToString
                End If
            Else
                perc = ""
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getModeOfPaymentByModeID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getModeOfPaymentByModeID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            'ds.Clear()
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getProductNameByProductID(ByVal autoid As Integer) As String
        Dim qry As String = "select Products from Products where AutoID = '" & autoid & "'"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "SalesAllocation")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = ""
                Else
                    perc = ds.Tables(0).Rows(0).Item(0).ToString
                End If
            Else
                perc = ""
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getProductNameByProductID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getProductNameByProductID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            'ds.Clear()
            con.Close()
        End Try
        Return perc
    End Function

    Public Function checkIfSameDayBanking(ByVal modeid As Integer) As String
        Dim qry As String = "select SameDayBanking from ModeOfPayment where ModeID = '" & modeid & "'"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "SalesAllocation")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = ""
                Else
                    perc = ds.Tables(0).Rows(0).Item(0).ToString
                End If
            Else
                perc = ""
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "checkIfSameDayBanking()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "checkIfSameDayBanking()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            'ds.Clear()
            con.Close()
        End Try
        Return perc
    End Function

    Public Sub updateReceiptNumberForReceipts(ByVal receiptno As String, ByVal batchname As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update Receipts set ReceiptNo = '" & receiptno & "' where BatchName = '" & batchname & "'"
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
            MessageBox.Show(ex.Message, "updateReceiptNumberForReceipts()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateReceiptNumberForReceipts()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub updateReceiptNumberForReceiptsBreakdown(ByVal receiptno As String, ByVal batchname As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update ReceiptsBreakdown set ReceiptNo = '" & receiptno & "' where BatchName = '" & batchname & "'"
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
            MessageBox.Show(ex.Message, "updateReceiptNumberForReceiptsBreakdown()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateReceiptNumberForReceiptsBreakdown()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Function checkIfThereIsApprovedPointOfSalesForPreviousDate(ByVal modeid As Integer, ByVal transactiondate As Date) As Boolean
        Dim qry As String = "select * from Receipts where Status = 'Approved' and ModeID = '" & modeid & "' AND TransactionDate <= '" & transactiondate & "'"

        Dim perc As Boolean = False

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "SalesAllocation")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = False
                Else
                    perc = True
                End If
            Else
                perc = False
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "checkIfThereIsApprovedPointOfSalesForPreviousDate()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "checkIfThereIsApprovedPointOfSalesForPreviousDate()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            'ds.Clear()
            con.Close()
        End Try
        Return perc
    End Function

    Public Function checkIfConfirmationNotNeeded(ByVal modeid As Integer) As String
        Dim qry As String = "select NoBankingConfirmation from ModeOfPayment where ModeID = '" & modeid & "'"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "SalesAllocation")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = ""
                Else
                    perc = ds.Tables(0).Rows(0).Item(0).ToString
                End If
            Else
                perc = ""
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "checkIfConfirmationNotNeeded()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "checkIfConfirmationNotNeeded()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            'ds.Clear()
            con.Close()
        End Try
        Return perc
    End Function

    Public Sub updateReceiptReferenceForNoConfirmations(ByVal reference As String, ByVal batchname As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update Receipts set Reference = '" & reference & "' where BatchName = '" & batchname & "'"
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
            MessageBox.Show(ex.Message, "updateReceiptReferenceForNoConfirmations()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateReceiptReferenceForNoConfirmations()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub updateReceiptBreakdownReferenceForNoConfirmations(ByVal reference As String, ByVal batchname As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update ReceiptsBreakdown set Reference = '" & reference & "' where BatchName = '" & batchname & "'"
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
            MessageBox.Show(ex.Message, "updateReceiptBreakdownReferenceForNoConfirmations()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateReceiptBreakdownReferenceForNoConfirmations()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub updateReceiptCapturedByDetails(ByVal capturedby As String, ByVal captureddate As Date, ByVal batchname As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update Receipts set CapturedBy = '" & capturedby & "',CapturedDate = '" & captureddate & "',CapturedTime = '" & captureddate & "' where BatchName = '" & batchname & "'"
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
            MessageBox.Show(ex.Message, "updateReceiptCapturedByDetails()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateReceiptCapturedByDetails()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Function getCompanyBranch(ByVal branchid As Integer) As String
        Dim qry As String = "select BranchName from CompanyBranches where AutoID = '" & branchid & "'"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "DealTemplate")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) Then
                    perc = ""
                Else
                    perc = ds.Tables(0).Rows(0).Item(0).ToString
                End If
            Else
                perc = ""
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getCompanyBranch()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getCompanyBranch()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Sub updateReceiptPayerIDByBatchName(ByVal PayerID As String, ByVal BatchName As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update Receipts set PayerID = '" & PayerID & "' where BatchName = '" & BatchName & "'"
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
            MessageBox.Show(ex.Message, "updateReceiptPayerIDByBatchName()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateReceiptPayerIDByBatchName()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub updateReceiptBreakdownPayerIDByBatchName(ByVal PayerID As String, ByVal BatchName As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update ReceiptsBreakdown set PayerID = '" & PayerID & "' where BatchName = '" & BatchName & "'"
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
            MessageBox.Show(ex.Message, "updateReceiptBreakdownPayerIDByBatchName()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateReceiptBreakdownPayerIDByBatchName()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub updateQuerry(ByVal querry As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = querry
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
            MessageBox.Show(ex.Message, "updateQuerry()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateQuerry()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Function getExpectedDeductionsPerProductInDatabase(ByVal customerid As String, ByVal formonth As Date) As Decimal
        Dim qry As String = "select SUM(Amount) from ExpectedDeductionsPerProduct  where CustomerID = '" & customerid & "' and ForMonth = '" & formonth & "'"

        Dim perc As Decimal = 0

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = 0
                Else
                    perc = Decimal.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
                End If
            Else
                perc = 0
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getExpectedDeductionsPerProductInDatabase()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getExpectedDeductionsPerProductInDatabase()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Sub updateReceiptCompanyBookID(ByVal BookID As Integer, ByVal ReceiptID As Integer)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update Receipts set BookID = '" & BookID & "' where AutoID = '" & ReceiptID & "'"
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
            MessageBox.Show(ex.Message, "updateReceiptCompanyBookID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateReceiptCompanyBookID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub updateReceiptBreakdownCompanyBookID(ByVal BookID As Integer, ByVal BatchName As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update ReceiptsBreakdown set BookID = '" & BookID & "' where BatchName = '" & BatchName & "'"
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
            MessageBox.Show(ex.Message, "updateReceiptBreakdownCompanyBookID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateReceiptBreakdownCompanyBookID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub updateReceiptModeID(ByVal ModeID As Integer, ByVal ReceiptID As Integer)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update Receipts set ModeID = '" & ModeID & "' where AutoID = '" & ReceiptID & "'"
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
            MessageBox.Show(ex.Message, "updateReceiptModeID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateReceiptModeID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub updateReceiptBreakdownModeID(ByVal ModeID As Integer, ByVal BatchName As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update ReceiptsBreakdown set ModeID = '" & ModeID & "' where BatchName = '" & BatchName & "'"
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
            MessageBox.Show(ex.Message, "updateReceiptBreakdownModeID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateReceiptBreakdownModeID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub savePaymentDetails(ByVal ReceivedFrom As String, ByVal ReceivedFromIDNumber As String, ByVal ReceivedFromCellNumber As String, ByVal PayerName As String, ByVal PayerID As String, ByVal PayerCellNumber As String, ByVal ReceiptDate As Date, ByVal PaidAmount As Decimal, ByVal BatchName As String, ByVal Description As String, ByVal ReceiptNo As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try
            Dim qryinsert As String = "Insert into PaymentDetails values ('" & ReceivedFrom & "','" & ReceivedFromIDNumber & "','" & ReceivedFromCellNumber & "','" & PayerName & "','" & PayerID & "','" & PayerCellNumber & "','" & ReceiptDate & "','" & PaidAmount & "','" & BatchName & "','" & Description & "','" & ReceiptNo & "')"
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
            MessageBox.Show(ex.Message, "savePaymentDetails()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "savePaymentDetails()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Function getPostalAddress() As String
        Dim qry As String = "select AuthorizedAdress from CompanyDetails"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = ds.Tables(0).Rows(0).Item(0).ToString
            Else
                perc = ""
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getPostalAddress()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getPostalAddress()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getPhysicalAddress() As String
        Dim qry As String = "select PhysicalAddress from CompanyDetails"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = ds.Tables(0).Rows(0).Item(0).ToString
            Else
                perc = ""
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getPhysicalAddress()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getPhysicalAddress()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getCompanyContactNumber() As String
        Dim qry As String = "select ContactNumber from CompanyDetails"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = ds.Tables(0).Rows(0).Item(0).ToString
            Else
                perc = ""
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getCompanyContactNumber()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getCompanyContactNumber()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getCompanyEmailAdress() As String
        Dim qry As String = "select EmailAdress from CompanyDetails"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = ds.Tables(0).Rows(0).Item(0).ToString
            Else
                perc = ""
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getCompanyEmailAdress()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getCompanyEmailAdress()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getCompanyFaxNumber() As String
        Dim qry As String = "select FaxNumber from CompanyDetails"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = ds.Tables(0).Rows(0).Item(0).ToString
            Else
                perc = ""
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getCompanyFaxNumber()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getCompanyFaxNumber()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getNextReceipNo() As String
        Dim qry As String = "select Top 1 ReceiptNo from PaymentDetails Order by AutoID DESC"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                Dim int As Integer = 0
                int = Integer.Parse(ds.Tables(0).Rows(0).Item(0).ToString.Substring(3))
                perc = "REC" & int + 1
            Else
                perc = "REC1"
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getNextReceipNo()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getNextReceipNo()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getProductsForTotals(ByVal PaymentDate As Date, ByVal PayBy As String, ByVal Status As String) As DataSet
        Dim qry As String = "select DISTINCT(ProductID),Products from ProductsCheques inner join Products on ProductsCheques.ProductID = Products.AutoID where PaymentDate = '" & PaymentDate & "' and PayBy = '" & PayBy & "' and ChequeNumber not in ('') and ProductsCheques.Status = '" & Status & "'"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getProductsForTotals()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getProductsForTotals()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return ds
    End Function

    Public Function getTotalPerProductPerCustomerIDPerPaymentDate(ByVal PaymentDate As Date, ByVal PayBy As String, ByVal ProductID As Integer, ByVal CustomerID As String, ByVal Status As String) As Decimal
        Dim qry As String = "select SUM(Amount) from ProductsCheques where PaymentDate = '" & PaymentDate & "' and PayBy = '" & PayBy & "' and ChequeNumber not in ('')  and ProductID = '" & ProductID & "' and CustomerID = '" & CustomerID & "' and Status = '" & Status & "'"

        Dim perc As Decimal = 0

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = 0
                Else
                    perc = Decimal.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
                End If
            Else
                perc = 0
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getTotalPerProductPerCustomerIDPerPaymentDate()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getTotalPerProductPerCustomerIDPerPaymentDate()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Sub updateDailyTotalsExport(ByVal PaymentDate As Date, ByVal PayBy As String, ByVal CurrentStatus As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update ProductsCheques set Status = 'Exported' where PaymentDate = '" & PaymentDate & "' and PayBy = '" & PayBy & "' and ChequeNumber not in ('') and Status = '" & CurrentStatus & "'"
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
            MessageBox.Show(ex.Message, "updateDailyTotalsExport()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateDailyTotalsExport()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Function checkIfToModeIsDirectDeposit(ByVal ModeID As Integer) As Boolean
        Dim qry As String = "select * from ModeOfPayment where NoBankingConfirmation = 'YES' and ModeID = '" & ModeID & "'"

        Dim perc As Boolean = False

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "SalesAllocation")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = False
                Else
                    perc = True
                End If
            Else
                perc = False
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "checkIfToModeIsDirectDeposit()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "checkIfToModeIsDirectDeposit()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            'ds.Clear()
            con.Close()
        End Try
        Return perc
    End Function

    Public Sub saveCovidDetails(ByVal CustomerID As String, ByVal FirstNames As String, ByVal Surname As String, ByVal HomeAddress As String, ByVal Contacts As String, ByVal Gender As String, ByVal Temperature As Decimal, ByVal VisitedDate As Date, ByVal VisitedTime As Date, ByVal CapturedDate As Date, ByVal CapturedTime As Date, ByVal CapturedBy As String, ByVal Office As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try
            Dim qryinsert As String = "Insert into CovidRegister values ('" & CustomerID & "','" & FirstNames & "','" & Surname & "','" & HomeAddress & "','" & Contacts & "','" & Gender & "','" & Temperature & "','" & VisitedDate & "','" & VisitedTime & "','" & CapturedDate & "','" & CapturedTime & "','" & CapturedBy & "','" & Office & "')"
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
            MessageBox.Show(ex.Message, "saveCovidDetails()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "saveCovidDetails()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Function getCountForFilling(ByVal con As SqlClient.SqlConnection, ByVal str As String) As Integer
        Dim qry As String = str

        Dim perc As Integer = 0

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = 0
                Else
                    perc = Integer.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
                End If
            Else
                perc = 0
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getCountForFilling()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getCountForFilling()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Sub saveFillingConfirmation(ByVal Type As String, ByVal TypeID As Integer, ByVal Status As String, ByVal con As SqlClient.SqlConnection)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "Insert into FillingConfirmation values ('" & Type & "','" & TypeID & "','" & Status & "')"
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
            MessageBox.Show(ex.Message, "saveFillingConfirmation()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "saveFillingConfirmation()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Function getFillingData(ByVal con As SqlClient.SqlConnection, ByVal str As String) As DataSet
        Dim qry As String = str

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getFillingData()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getFillingData()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return ds
    End Function

    Public Function getPaidForAmountPerProductID(ByVal CustomerID As String, ByVal ProductID As Integer, ByVal PaidByCustomerID As String, ByVal EmployerID As Integer) As Decimal
        Dim qry As String = "select SUM(Amount) from PaidForCustomers where Status = 'Active' and CustomerID = '" & CustomerID & "' and ProductID = '" & ProductID & "' and PaidByCustomerID = '" & PaidByCustomerID & "' and EmployerID = '" & EmployerID & "'"

        Dim perc As Decimal = 0

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = 0
                Else
                    perc = Decimal.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
                End If
            Else
                perc = 0
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getPaidForAmountPerProductID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getPaidForAmountPerProductID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function checkProductForCustomerPaidFor(ByVal CustomerID As String, ByVal ProductID As Integer, ByVal PaidByCustomerID As String, ByVal EmployerID As Integer) As Boolean
        Dim qry As String = "select * from PaidForCustomers where CustomerID = '" & CustomerID & "' and ProductID = '" & ProductID & "' and PaidByCustomerID = '" & PaidByCustomerID & "' and EmployerID = '" & employerid & "'"

        Dim perc As Boolean = False

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = True
            Else
                perc = False
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "checkProductForCustomerPaidFor()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "checkProductForCustomerPaidFor()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Sub savePaidForCustomer(ByVal CustomerID As String, ByVal ProductID As Integer, ByVal Amount As Decimal, ByVal Status As String, ByVal PaidByCustomerID As String, ByVal EmployerID As Integer, ByVal con As SqlClient.SqlConnection)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "Insert into PaidForCustomers values ('" & CustomerID & "','" & ProductID & "','" & Amount & "','" & Status & "','" & PaidByCustomerID & "','" & EmployerID & "')"
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
            MessageBox.Show(ex.Message, "savePaidForCustomer()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "savePaidForCustomer()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub updatePaidForCustomer(ByVal Amount As Decimal, ByVal Status As String, ByVal CustomerID As String, ByVal ProductID As Integer, ByVal PaidByCustomerID As String, ByVal EmployerID As Integer)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update PaidForCustomers set Amount = '" & Amount & "',Status = '" & Status & "' where CustomerID = '" & CustomerID & "' and ProductID = '" & ProductID & "' and PaidByCustomerID = '" & PaidByCustomerID & "' and EmployerID = '" & EmployerID & "'"
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
            MessageBox.Show(ex.Message, "updatePaidForCustomer()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updatePaidForCustomer()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub updatePaidForCustomerStatus(ByVal Status As String, ByVal CustomerID As String, ByVal ProductID As Integer, ByVal PaidByCustomerID As String, ByVal EmployerID As Integer)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update PaidForCustomers set Status = '" & Status & "' where CustomerID = '" & CustomerID & "' and ProductID = '" & ProductID & "' and PaidByCustomerID = '" & PaidByCustomerID & "' and EmployerID = '" & EmployerID & "'"
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
            MessageBox.Show(ex.Message, "updatePaidForCustomerStatus()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updatePaidForCustomerStatus()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Function getProductName(ByVal AutoID As Integer) As String
        Dim qry As String = "select Products from Products where AutoID = '" & AutoID & "'"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = ds.Tables(0).Rows(0).Item(0).ToString
            Else
                perc = ""
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getProductName()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getProductName()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getProductIDForPaidForCustomers() As Integer
        Dim qry As String = "select ProductIDForPaidCustomers from CompanyDetails"

        Dim perc As Integer = 0

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    perc = 0
                Else
                    perc = Integer.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
                End If
            Else
                perc = 0
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getProductIDForPaidForCustomers()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getProductIDForPaidForCustomers()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Sub saveCustomer(ByVal Title As String, ByVal CountryID As String, ByVal IdentityTypeID As String, ByVal IdentityExpiryDate As Date, ByVal IDNumber As String, ByVal FirstNames As String, ByVal Surname As String, ByVal DateOfBirth As Date, ByVal GenderID As String, ByVal CellNumber As String, ByVal Email As String, ByVal PostalAddress As String, ByVal HomeTelephone As String, ByVal RegionID As String, ByVal TownVillageID As String, ByVal Area As String, ByVal Ward As String, ByVal Chief As String, ByVal HomePhysicalAddress As String, ByVal SchoolID As String, ByVal EmployerID As String, ByVal DepartmentID As String, ByVal Occupation As String, ByVal WorkPostalAddress As String, ByVal WorkPhysicalAddress As String, ByVal WorkEmailAddress As String, ByVal EmploymentType As String, ByVal EmployerContactPerson As String, ByVal DateOfEngagement As Date, ByVal EmpTelephone1 As String, ByVal EmpTelephone2 As String, ByVal EmpFax_1 As String, ByVal EmpFax_2 As String, ByVal NextOfKinName1 As String, ByVal NextOfKinID1 As String, ByVal NextOfKinContact1 As String, ByVal NextOfKinAddress1 As String, ByVal NextOfKinName2 As String, ByVal NextOfKinID2 As String, ByVal NextOfKinContact2 As String, ByVal NextOfKinAddress2 As String, ByVal NetSalary As Decimal, ByVal BankName As String, ByVal BranchName As String, ByVal AccountNumber As String, ByVal BranchID As String, ByVal CreationDate As Date, ByVal Status As String, ByVal DateOfDeath As Date, ByVal PlaceOfDeath As String, ByVal LastModified As Date, ByVal GroupID As String, ByVal PayrolNumber As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "Insert into Customers values ('" & Title & "','" & CountryID & "','" & IdentityTypeID & "','" & IdentityExpiryDate & "','" & IDNumber & "','" & FirstNames & "','" & Surname & "','" & DateOfBirth & "','" & GenderID & "','" & CellNumber & "','" & Email & "','" & PostalAddress & "','" & HomeTelephone & "','" & RegionID & "','" & TownVillageID & "','" & Area & "','" & Ward & "','" & Chief & "','" & HomePhysicalAddress & "','" & SchoolID & "','" & EmployerID & "','" & DepartmentID & "','" & Occupation & "','" & WorkPostalAddress & "','" & WorkPhysicalAddress & "','" & WorkEmailAddress & "','" & EmploymentType & "','" & EmployerContactPerson & "','" & DateOfEngagement & "','" & EmpTelephone1 & "','" & EmpTelephone2 & "','" & EmpFax_1 & "','" & EmpFax_2 & "','" & NextOfKinName1 & "','" & NextOfKinID1 & "','" & NextOfKinContact1 & "','" & NextOfKinAddress1 & "','" & NextOfKinName2 & "','" & NextOfKinID2 & "','" & NextOfKinContact2 & "','" & NextOfKinAddress2 & "','" & NetSalary & "','" & BankName & "','" & BranchName & "','" & AccountNumber & "','" & BranchID & "','" & CreationDate & "','" & Status & "','" & DateOfDeath & "','" & PlaceOfDeath & "','" & LastModified & "','" & GroupID & "','" & PayrolNumber & "')"
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
            MessageBox.Show(ex.Message, "saveCustomer()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "saveCustomer()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Function getCustomerDetailsFormCustomer(ByVal IDNumber As String) As DataSet
        Dim qry As String = "select * from Customers where IDNumber = '" & IDNumber & "'"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getCustomerDetailsFormCustomer()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getCustomerDetailsFormCustomer()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return ds
    End Function

    Public Function getValueString(ByVal str As String) As String
        Dim qry As String = str

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = ds.Tables(0).Rows(0).Item(0).ToString
            Else
                perc = ""
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getValueString()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getValueString()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getCustomerDetailsFromSystems(ByVal con As SqlClient.SqlConnection, ByVal str As String) As DataSet
        Dim qry As String = str

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getCustomerDetailsFromSystems()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getCustomerDetailsFromSystems()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return ds
    End Function

    Public Function checkCustomerInCustomer(ByVal IDNumber As String) As Boolean
        Dim qry As String = "select * from Customers where IDNumber = '" & IDNumber & "'"

        Dim perc As Boolean = False

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = True
            Else
                perc = False
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "checkCustomerInCustomer()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "checkCustomerInCustomer()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Sub saveEmployer(ByVal Employer As String, ByVal CollectionCommission As Decimal)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "Insert into Employer values ('" & Employer & "','" & CollectionCommission & "')"
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
            MessageBox.Show(ex.Message, "saveEmployer()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "saveEmployer()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Function getUsernameByName(ByVal Name As String) As String
        Dim qry As String = "select Username from Users where Name = '" & Name & "'"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = ds.Tables(0).Rows(0).Item(0).ToString
            Else
                perc = ""
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getUsernameByName()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getUsernameByName()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Sub savePrintingRequest(ByVal Username As String, ByVal SheetNumber As String, ByVal RequestTime As Date, ByVal RequestedBy As String, ByVal Status As String, ByVal PrintedDate As Date)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "Insert into PrintingRequests values ('" & Username & "','" & SheetNumber & "','" & RequestTime & "','" & RequestedBy & "','" & Status & "','" & PrintedDate & "')"
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
            MessageBox.Show(ex.Message, "savePrintingRequest()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "savePrintingRequest()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Function getSheetNumber(ByVal AutoID As Integer) As String
        Dim qry As String = "select SheetNumber from PrintingRequests where AutoID = '" & AutoID & "'"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "VCD")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                perc = ds.Tables(0).Rows(0).Item(0).ToString
            Else
                perc = ""
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getSheetNumber()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getSheetNumber()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Sub updatePrintingRequest(ByVal Username As String, ByVal SheetNumber As String, ByVal RequestTime As Date, ByVal RequestedBy As String, ByVal AutoID As Integer)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update PrintingRequests set Username = '" & Username & "',SheetNumber = '" & SheetNumber & "',RequestTime = '" & RequestTime & "',RequestedBy = '" & RequestedBy & "' where AutoID = '" & AutoID & "'"
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
            MessageBox.Show(ex.Message, "updatePrintingRequest()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updatePrintingRequest()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub updatePrintingStatusAndDate(ByVal PrintedDate As Date, ByVal Status As String, ByVal AutoID As Integer)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update PrintingRequests set PrintedDate = '" & PrintedDate & "',Status = '" & Status & "' where AutoID = '" & AutoID & "'"
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
            MessageBox.Show(ex.Message, "updatePrintingStatusAndDate()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updatePrintingStatusAndDate()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub updatePrintingStatusAlone(ByVal Status As String, ByVal AutoID As Integer)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update PrintingRequests set Status = '" & Status & "' where AutoID = '" & AutoID & "'"
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
            MessageBox.Show(ex.Message, "updatePrintingStatusAlone()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updatePrintingStatusAlone()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

End Module
