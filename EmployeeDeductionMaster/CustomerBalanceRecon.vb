Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc

Public Class CustomerBalanceRecon

    Dim refund, billamount, rechargeamount, contractual, receipts, recharges, billings As Decimal
    Dim billcomment, rechargecomment As String

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        DataGridView1.AllowUserToAddRows = False
        Button1.Enabled = False
        Button2.Enabled = True

        If decision("Do you want to load the deductions ?") = DialogResult.Yes Then
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
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        For Each row As DataGridViewRow In DataGridView1.Rows

            contractual = Decimal.Parse(row.Cells(1).Value)
            receipts = Decimal.Parse(row.Cells(2).Value)
            recharges = Decimal.Parse(row.Cells(3).Value)
            billings = Decimal.Parse(row.Cells(4).Value)

            If receipts = contractual Then
                If recharges = contractual Then
                    If receipts = recharges Then
                        contractual = 0
                        receipts = 0
                        billings = 0
                        billcomment = ""
                        recharges = 0
                        rechargecomment = ""
                        If billings < receipts Then
                            If recharges < receipts Then
                                If billings < recharges Then
                                    refund = receipts - recharges
                                ElseIf billings > recharges Then
                                    refund = receipts - billings
                                End If
                            End If
                        End If
                    ElseIf receipts < recharges Then
                        contractual = 0
                        receipts = 0
                        billings = 0
                        billcomment = ""
                        recharges = 0
                        rechargecomment = ""
                        If billings < receipts Then
                            If recharges < receipts Then
                                If billings < recharges Then
                                    refund = receipts - recharges
                                ElseIf billings > recharges Then
                                    refund = receipts - billings
                                End If
                            End If
                        End If
                    ElseIf receipts > recharges Then
                        contractual = 0
                        receipts = 0
                        billings = 0
                        billcomment = ""
                        recharges = 0
                        rechargecomment = ""
                        If billings < receipts Then
                            If recharges < receipts Then
                                If billings < recharges Then
                                    refund = receipts - recharges
                                ElseIf billings > recharges Then
                                    refund = receipts - billings
                                End If
                            End If
                        End If
                    End If
                ElseIf recharges < contractual Then
                    If receipts = recharges Then
                        contractual = 0
                        receipts = 0
                        billings = 0
                        billcomment = ""
                        recharges = 0
                        rechargecomment = ""
                        If billings < receipts Then
                            If recharges < receipts Then
                                If billings < recharges Then
                                    refund = receipts - recharges
                                ElseIf billings > recharges Then
                                    refund = receipts - billings
                                End If
                            End If
                        End If
                    ElseIf receipts < recharges Then
                        contractual = 0
                        receipts = 0
                        billings = 0
                        billcomment = ""
                        recharges = 0
                        rechargecomment = ""
                        If billings < receipts Then
                            If recharges < receipts Then
                                If billings < recharges Then
                                    refund = receipts - recharges
                                ElseIf billings > recharges Then
                                    refund = receipts - billings
                                End If
                            End If
                        End If
                    ElseIf receipts > recharges Then
                        billcomment = "Uder-recharged customer"
                        recharges = receipts - recharges
                        contractual = 0
                        receipts = 0
                        billings = 0
                        billcomment = ""
                        If billings < receipts Then
                            If recharges < receipts Then
                                If billings < recharges Then
                                    refund = receipts - recharges
                                ElseIf billings > recharges Then
                                    refund = receipts - billings
                                End If
                            End If
                        End If
                    End If
                ElseIf recharges > contractual Then
                    If receipts = recharges Then
                        contractual = 0
                        receipts = 0
                        billings = 0
                        billcomment = ""
                        recharges = 0
                        rechargecomment = ""
                        If billings < receipts Then
                            If recharges < receipts Then
                                If billings < recharges Then
                                    refund = receipts - recharges
                                ElseIf billings > recharges Then
                                    refund = receipts - billings
                                End If
                            End If
                        End If
                    ElseIf receipts < recharges Then
                        recharges = "Over-recharged customer"
                        recharges = recharges - receipts
                        contractual = 0
                        receipts = 0
                        billings = 0
                        billcomment = ""
                        If billings < receipts Then
                            If recharges < receipts Then
                                If billings < recharges Then
                                    refund = receipts - recharges
                                ElseIf billings > recharges Then
                                    refund = receipts - billings
                                End If
                            End If
                        End If
                    ElseIf receipts > recharges Then
                        contractual = 0
                        receipts = 0
                        billings = 0
                        billcomment = ""
                        recharges = 0
                        rechargecomment = ""
                        If billings < receipts Then
                            If recharges < receipts Then
                                If billings < recharges Then
                                    refund = receipts - recharges
                                ElseIf billings > recharges Then
                                    refund = receipts - billings
                                End If
                            End If
                        End If
                    End If
                End If
            ElseIf receipts < contractual Then
                If recharges = contractual Then
                    If receipts = recharges Then
                        contractual = 0
                        receipts = 0
                        billings = 0
                        billcomment = ""
                        recharges = 0
                        rechargecomment = ""
                        If billings < receipts Then
                            If recharges < receipts Then
                                If billings < recharges Then
                                    refund = receipts - recharges
                                ElseIf billings > recharges Then
                                    refund = receipts - billings
                                End If
                            End If
                        End If
                    ElseIf receipts < recharges Then
                        rechargecomment = "Over-recharged customer"
                        recharges = recharges - receipts
                        contractual = 0
                        receipts = 0
                        billings = 0
                        billcomment = ""
                        If billings < receipts Then
                            If recharges < receipts Then
                                If billings < recharges Then
                                    refund = receipts - recharges
                                ElseIf billings > recharges Then
                                    refund = receipts - billings
                                End If
                            End If
                        End If
                    ElseIf receipts > recharges Then
                        contractual = 0
                        receipts = 0
                        billings = 0
                        billcomment = ""
                        recharges = 0
                        rechargecomment = ""
                        If billings < receipts Then
                            If recharges < receipts Then
                                If billings < recharges Then
                                    refund = receipts - recharges
                                ElseIf billings > recharges Then
                                    refund = receipts - billings
                                End If
                            End If
                        End If
                    End If
                ElseIf recharges < contractual Then
                    If receipts = recharges Then
                        contractual = 0
                        receipts = 0
                        billings = 0
                        billcomment = ""
                        recharges = 0
                        rechargecomment = ""
                        If billings < receipts Then
                            If recharges < receipts Then
                                If billings < recharges Then
                                    refund = receipts - recharges
                                ElseIf billings > recharges Then
                                    refund = receipts - billings
                                End If
                            End If
                        End If
                    ElseIf receipts < recharges Then
                        rechargecomment = "Over-recharged customer"
                        recharges = recharges - receipts
                        contractual = 0
                        receipts = 0
                        billings = 0
                        billcomment = ""
                        If billings < receipts Then
                            If recharges < receipts Then
                                If billings < recharges Then
                                    refund = receipts - recharges
                                ElseIf billings > recharges Then
                                    refund = receipts - billings
                                End If
                            End If
                        End If
                    ElseIf receipts > recharges Then
                        rechargecomment = "Uder-recharged customer"
                        recharges = receipts - recharges
                        contractual = 0
                        receipts = 0
                        billings = 0
                        billcomment = ""
                        If billings < receipts Then
                            If recharges < receipts Then
                                If billings < recharges Then
                                    refund = receipts - recharges
                                ElseIf billings > recharges Then
                                    refund = receipts - billings
                                End If
                            End If
                        End If
                    End If
                ElseIf recharges > contractual Then
                    If receipts = recharges Then
                        contractual = 0
                        receipts = 0
                        billings = 0
                        billcomment = ""
                        recharges = 0
                        rechargecomment = ""
                        If billings < receipts Then
                            If recharges < receipts Then
                                If billings < recharges Then
                                    refund = receipts - recharges

                                ElseIf billings > recharges Then
                                    refund = receipts - billings

                                End If
                            End If
                        End If
                    ElseIf receipts < recharges Then
                        rechargecomment = "Over-recharged customer"
                        recharges = recharges - receipts
                        contractual = 0
                        receipts = 0
                        billings = 0
                        billcomment = ""
                        If billings < receipts Then
                            If recharges < receipts Then
                                If billings < recharges Then
                                    refund = receipts - recharges

                                ElseIf billings > recharges Then
                                    refund = receipts - billings

                                End If
                            End If
                        End If
                    ElseIf receipts > recharges Then
                        contractual = 0
                        receipts = 0
                        billings = 0
                        billcomment = ""
                        recharges = 0
                        rechargecomment = ""
                        If billings < receipts Then
                            If recharges < receipts Then
                                If billings < recharges Then
                                    refund = receipts - recharges

                                ElseIf billings > recharges Then
                                    refund = receipts - billings

                                End If
                            End If
                        End If
                    End If
                End If
            ElseIf receipts > contractual Then
                If recharges = contractual Then
                    If receipts = recharges Then
                        contractual = 0
                        receipts = 0
                        billings = 0
                        billcomment = ""
                        recharges = 0
                        rechargecomment = ""
                        If billings < receipts Then
                            If recharges < receipts Then
                                If billings < recharges Then
                                    refund = receipts - recharges

                                ElseIf billings > recharges Then
                                    refund = receipts - billings

                                End If
                            End If
                        End If
                    ElseIf receipts < recharges Then
                        contractual = 0
                        receipts = 0
                        billings = 0
                        billcomment = ""
                        recharges = 0
                        rechargecomment = ""
                        If billings < receipts Then
                            If recharges < receipts Then
                                If billings < recharges Then
                                    refund = receipts - recharges

                                ElseIf billings > recharges Then
                                    refund = receipts - billings

                                End If
                            End If
                        End If
                    ElseIf receipts > recharges Then
                        refund = receipts - recharges
                        contractual = 0
                        receipts = 0
                        billings = 0
                        billcomment = ""
                        recharges = 0
                        rechargecomment = ""
                    End If
                ElseIf recharges < contractual Then
                    If receipts = recharges Then
                        contractual = 0
                        receipts = 0
                        billings = 0
                        billcomment = ""
                        recharges = 0
                        rechargecomment = ""
                        If billings < receipts Then
                            If recharges < receipts Then
                                If billings < recharges Then
                                    refund = receipts - recharges

                                ElseIf billings > recharges Then
                                    refund = receipts - billings

                                End If
                            End If
                        End If
                    ElseIf receipts < recharges Then
                        contractual = 0
                        receipts = 0
                        billings = 0
                        billcomment = ""
                        recharges = 0
                        rechargecomment = ""
                        If billings < receipts Then
                            If recharges < receipts Then
                                If billings < recharges Then
                                    refund = receipts - recharges

                                ElseIf billings > recharges Then
                                    refund = receipts - billings

                                End If
                            End If
                        End If
                    ElseIf receipts > recharges Then
                        rechargecomment = "Uder-recharged customer"
                        recharges = contractual - recharges
                        contractual = 0
                        receipts = 0
                        billings = 0
                        billcomment = ""
                        refund = contractual - receipts
                    End If
                ElseIf recharges > contractual Then
                    If receipts = recharges Then
                        contractual = 0
                        receipts = 0
                        billings = 0
                        billcomment = ""
                        recharges = 0
                        rechargecomment = ""
                        If billings < receipts Then
                            If recharges < receipts Then
                                If billings < recharges Then
                                    refund = receipts - recharges

                                ElseIf billings > recharges Then
                                    refund = receipts - billings

                                End If
                            End If
                        End If
                    ElseIf receipts < recharges Then
                        rechargecomment = "Over-recharged customer"
                        recharges = recharges - receipts
                        contractual = 0
                        receipts = 0
                        billings = 0
                        billcomment = ""
                        If billings < receipts Then
                            If recharges < receipts Then
                                If billings < recharges Then
                                    refund = receipts - recharges

                                ElseIf billings > recharges Then
                                    refund = receipts - billings

                                End If
                            End If
                        End If
                    ElseIf receipts > recharges Then
                        refund = receipts - recharges
                        contractual = 0
                        receipts = 0
                        billings = 0
                        billcomment = ""
                        recharges = 0
                        rechargecomment = ""
                    End If
                End If
            End If

            row.Cells(5).Value = refund
            row.Cells(6).Value = billcomment
            row.Cells(7).Value = billings
            row.Cells(8).Value = rechargecomment
            row.Cells(9).Value = recharges

        Next

    End Sub
End Class