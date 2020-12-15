Public Class ReceiptAuthorization

    Public rc As ReceiptingCustomer
    Public urb As UploadReceiptBatch
    Public arb As ApproveReceiptBatch
    Dim useravailable As String
    Dim ds As New DataSet
    Dim authorizer As String
    Public whichside As String

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Then
            MessageBox.Show("Enter all the details", "Button1_Click()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            getUser(TextBox1.Text)
            If useravailable = "Yes" Then
                Dim cipherText As String = wrapper.EncryptData(TextBox2.Text.Trim)
                If cipherText = ds.Tables(0).Rows(0).Item(1).ToString Then
                    authorizer = Me.ds.Tables(0).Rows(0).Item(3).ToString
                    If whichside = "Normal" Then
                        rc.proceedToSaveReceipt(authorizer, Integer.Parse(Me.ds.Tables(0).Rows(0).Item(5).ToString), "New")
                    ElseIf whichside = "Upload" Then
                        urb.saveBulkReceipt(authorizer)
                    ElseIf whichside = "Batch" Then
                        arb.proceedToConfirm(authorizer)
                    End If
                    Me.Close()
                Else
                    MessageBox.Show("The password is not correct", "Button1_Click()", MessageBoxButtons.OK, MessageBoxIcon.Error)
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

    Private Sub ReceiptAuthorization_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class