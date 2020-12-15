Public Class DeleteCustomers

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If decision("Are you sure you want to proceed ?") = DialogResult.Yes Then
            If checkCustomerDetailsInDatabase(TextBox1.Text.Trim) = True Then
                If checkCustomerDetailsInDatabase(TextBox2.Text.Trim) = True Then
                    updateQuerry("update ActualDeductions set CustomerID = '" & TextBox2.Text.Trim & "' where CustomerID = '" & TextBox1.Text.Trim & "'")
                    updateQuerry("update Balances set CustomerID = '" & TextBox2.Text.Trim & "' where CustomerID = '" & TextBox1.Text.Trim & "'")
                    updateQuerry("update CancelledCheques set CustomerID = '" & TextBox2.Text.Trim & "' where CustomerID = '" & TextBox1.Text.Trim & "'")
                    updateQuerry("update ExpectedDeductions set CustomerID = '" & TextBox2.Text.Trim & "' where CustomerID = '" & TextBox1.Text.Trim & "'")
                    updateQuerry("update ExpectedDeductionsPerProduct set CustomerID = '" & TextBox2.Text.Trim & "' where CustomerID = '" & TextBox1.Text.Trim & "'")
                    updateQuerry("update Receipts set PayerID = '" & TextBox2.Text.Trim & "' where PayerID = '" & TextBox1.Text.Trim & "'")
                    updateQuerry("update ReceiptsBreakdown set PayerID = '" & TextBox2.Text.Trim & "' where PayerID = '" & TextBox1.Text.Trim & "'")
                    updateQuerry("update SettlementBalances set CustomerID = '" & TextBox2.Text.Trim & "' where CustomerID = '" & TextBox1.Text.Trim & "'")
                    updateQuerry("update StoppedRefunds set CustomerID = '" & TextBox2.Text.Trim & "' where CustomerID = '" & TextBox1.Text.Trim & "'")
                    updateQuerry("update TerminationsAndAmendments set CustomerID = '" & TextBox2.Text.Trim & "' where CustomerID = '" & TextBox1.Text.Trim & "'")
                    updateQuerry("update ProductsCheques set CustomerID = '" & TextBox2.Text.Trim & "' where CustomerID = '" & TextBox1.Text.Trim & "'")

                    updateQuerry("delete from CustomerDetails where CustomerID = '" & TextBox1.Text.Trim & "'")

                    MessageBox.Show("Done Updating and deleting the details", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show("Customer does not exist in the database (TO)", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            Else
                MessageBox.Show("Customer does not exist in the database (FROM)", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
    End Sub
End Class