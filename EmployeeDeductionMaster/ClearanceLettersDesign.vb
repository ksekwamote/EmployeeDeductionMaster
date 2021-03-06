Public Class ClearanceLettersDesign



    Dim dttable As New DataTable

    Dim ToWhom, Re, OfID, ThisServes, HasFully, Therefore, ForMoreInfor, Yours, JobDescription As String
    Dim AutoID As Integer

    Dim EarlyPayPercentage As Decimal

    Private Sub Update_Settlement_Letters_Design_Details_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim dt As New DataTable
        DataGridView1.DataSource = Nothing

        dt.Columns.Add("Name").DataType = System.Type.GetType("System.String")
        dt.Columns.Add("Details").DataType = System.Type.GetType("System.String")

        loadDetails()

        For Each column As DataColumn In dttable.Columns
            addData(column.ColumnName, dttable.Rows(0).Item(column.ColumnName), dt)
        Next

        DataGridView1.DataSource = dt

        DataGridView1.Columns(1).Width = 1000

    End Sub

    Public Sub loadDetails()
        DataGridView1.DataSource = Nothing
        Dim qry As String = "select * from ClearanceDetails"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim ctr As Integer = 0
            SQLAdapter.Fill(ds, "ClearanceDetails")
            dttable = ds.Tables(0)

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "loadDetails()()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Public Sub addData(ByVal name As String, ByVal details As String, ByVal dt As DataTable)

        Dim newRow As DataRow
        newRow = dt.NewRow

        newRow("Name") = name
        newRow("Details") = details

        dt.Rows.Add(newRow)

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If DataGridView1.Rows.Count > 0 Then
            AutoID = Integer.Parse(DataGridView1.Rows(0).Cells(1).Value.ToString)
            ToWhom = DataGridView1.Rows(1).Cells(1).Value.ToString
            Re = DataGridView1.Rows(2).Cells(1).Value.ToString
            OfID = DataGridView1.Rows(3).Cells(1).Value.ToString
            HasFully = DataGridView1.Rows(4).Cells(1).Value.ToString
            ThisServes = DataGridView1.Rows(5).Cells(1).Value.ToString
            Therefore = DataGridView1.Rows(6).Cells(1).Value.ToString
            ForMoreInfor = DataGridView1.Rows(7).Cells(1).Value.ToString
            Yours = DataGridView1.Rows(8).Cells(1).Value.ToString
            JobDescription = DataGridView1.Rows(9).Cells(1).Value.ToString

            updateClearanceLetterDesign(ToWhom, Re, OfID, ThisServes, HasFully, Therefore, ForMoreInfor, Yours, JobDescription, AutoID)
            Close()
        End If
    End Sub
End Class