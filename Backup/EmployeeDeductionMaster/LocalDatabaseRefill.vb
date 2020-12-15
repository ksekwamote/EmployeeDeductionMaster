Public Class LocalDatabaseRefill

    Dim dscustomers, dssettlement, dstermAmend, dsexpectpercustomer, dsexpectperproduct, dsactualdeduction As New DataSet

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If DatabaseLocation = "Local" Then

            dscustomers = getCustomerDetailsForRefil(conServer)
            ProgressBar1.Maximum = dscustomers.Tables(0).Rows.Count
            For custcount As Integer = 0 To dscustomers.Tables(0).Rows.Count - 1

                saveCustomerDetails(dscustomers.Tables(0).Rows(custcount).Item(1).ToString, dscustomers.Tables(0).Rows(custcount).Item(2).ToString, Integer.Parse(dscustomers.Tables(0).Rows(custcount).Item(3).ToString), dscustomers.Tables(0).Rows(custcount).Item(4).ToString)

                dssettlement.Clear()
                dssettlement = getSettlementsForRefil(dscustomers.Tables(0).Rows(custcount).Item(1).ToString, conServer)
                For settlecount As Integer = 0 To dssettlement.Tables(0).Rows.Count - 1
                    saveSettlementBalance(dssettlement.Tables(0).Rows(settlecount).Item(1).ToString, Integer.Parse(dssettlement.Tables(0).Rows(settlecount).Item(2).ToString), Decimal.Parse(dssettlement.Tables(0).Rows(settlecount).Item(3).ToString), Date.Parse(dssettlement.Tables(0).Rows(settlecount).Item(4).ToString), dssettlement.Tables(0).Rows(settlecount).Item(5).ToString, Date.Parse(dssettlement.Tables(0).Rows(settlecount).Item(6).ToString), Date.Parse(dssettlement.Tables(0).Rows(settlecount).Item(7).ToString), dssettlement.Tables(0).Rows(settlecount).Item(8).ToString, dssettlement.Tables(0).Rows(settlecount).Item(9).ToString)
                Next

                dstermAmend.Clear()
                dstermAmend = getTerminationsAndAmendmentsForRefil(dscustomers.Tables(0).Rows(custcount).Item(1).ToString, conServer)
                For termcount As Integer = 0 To dstermAmend.Tables(0).Rows.Count - 1
                    saveAmendments(dstermAmend.Tables(0).Rows(termcount).Item(1).ToString, Decimal.Parse(dstermAmend.Tables(0).Rows(termcount).Item(2).ToString), Date.Parse(dstermAmend.Tables(0).Rows(termcount).Item(3).ToString), Date.Parse(dstermAmend.Tables(0).Rows(termcount).Item(4).ToString), Integer.Parse(dstermAmend.Tables(0).Rows(termcount).Item(5).ToString), Integer.Parse(dstermAmend.Tables(0).Rows(termcount).Item(6).ToString))
                Next

                dsexpectpercustomer.Clear()
                dsexpectpercustomer = getExpectedDeductionsForRefil(dscustomers.Tables(0).Rows(custcount).Item(1).ToString, conServer)
                For expectcount As Integer = 0 To dsexpectpercustomer.Tables(0).Rows.Count - 1
                    saveExpectedDeduction(dsexpectpercustomer.Tables(0).Rows(expectcount).Item(1).ToString, Decimal.Parse(dsexpectpercustomer.Tables(0).Rows(expectcount).Item(2).ToString), Date.Parse(dsexpectpercustomer.Tables(0).Rows(expectcount).Item(3).ToString), Date.Parse(dsexpectpercustomer.Tables(0).Rows(expectcount).Item(4).ToString), Integer.Parse(dsexpectpercustomer.Tables(0).Rows(expectcount).Item(5).ToString))
                Next

                dsexpectperproduct.Clear()
                dsexpectperproduct = getExpectedDeductionsPerProductForRefil(dscustomers.Tables(0).Rows(custcount).Item(1).ToString, conServer)
                For perproductcount As Integer = 0 To dsexpectperproduct.Tables(0).Rows.Count - 1
                    saveExpectedDeductionsPerProduct(dsexpectperproduct.Tables(0).Rows(perproductcount).Item(1).ToString, Decimal.Parse(dsexpectperproduct.Tables(0).Rows(perproductcount).Item(2).ToString), Date.Parse(dsexpectperproduct.Tables(0).Rows(perproductcount).Item(3).ToString), Date.Parse(dsexpectperproduct.Tables(0).Rows(perproductcount).Item(4).ToString), Integer.Parse(dsexpectperproduct.Tables(0).Rows(perproductcount).Item(5).ToString), Integer.Parse(dsexpectperproduct.Tables(0).Rows(perproductcount).Item(6).ToString))
                Next

                dsactualdeduction.Clear()
                dsactualdeduction = getActualDeductionForRefil(dscustomers.Tables(0).Rows(custcount).Item(1).ToString, conServer)
                For actualcount As Integer = 0 To dsactualdeduction.Tables(0).Rows.Count - 1
                    saveActualDeduction(dsactualdeduction.Tables(0).Rows(actualcount).Item(1).ToString, Decimal.Parse(dsactualdeduction.Tables(0).Rows(actualcount).Item(2).ToString), Date.Parse(dsactualdeduction.Tables(0).Rows(actualcount).Item(3).ToString), Date.Parse(dsactualdeduction.Tables(0).Rows(actualcount).Item(4).ToString), Date.Parse(dsactualdeduction.Tables(0).Rows(actualcount).Item(5).ToString), Integer.Parse(dsactualdeduction.Tables(0).Rows(actualcount).Item(6).ToString))
                Next

                ProgressBar1.Value = custcount
            Next
            ProgressBar1.Value = 0

            MessageBox.Show("Customer Details Saved", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            MessageBox.Show("Cannot. You are logedin to the server", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub
End Class