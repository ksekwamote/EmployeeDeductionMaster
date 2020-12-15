Public Class EmployerAndGroupChanges
    Dim ds As New DataSet
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        dgv.Rows.Clear()
        ds.Clear()
        ds = getDataset("select WhatIsDone,WhoDidIt,WhenIsDone,CustomerID from UserActions where CustomerID = '" & tbcustomerid.Text.Trim & "' and WhatIsDone like 'Employer Changed%' or WhatIsDone like 'Group Changed%'", con)
        If ds.Tables(0).Rows.Count > 0 Then
            For count As Integer = 0 To ds.Tables(0).Rows.Count - 1
                dgv.Rows.Add()
                dgv.Rows(dgv.Rows.Count - 1).Cells(0).Value = ds.Tables(0).Rows(count).Item(0).ToString
                dgv.Rows(dgv.Rows.Count - 1).Cells(1).Value = ds.Tables(0).Rows(count).Item(1).ToString
                dgv.Rows(dgv.Rows.Count - 1).Cells(2).Value = ds.Tables(0).Rows(count).Item(2).ToString
                dgv.Rows(dgv.Rows.Count - 1).Cells(3).Value = ds.Tables(0).Rows(count).Item(3).ToString
                dgv.Rows(dgv.Rows.Count - 1).Cells(4).Value = "EDM"
            Next
        End If

        ds.Clear()
        ds = getDataset("select DoneWhat,Who,TransactionDate,CustomerID from AuditTrail where CustomerID = '" & tbcustomerid.Text.Trim & "' and DoneWhat like 'Employer Changed%' or DoneWhat like 'Group Changed%'", conFPM)
        If ds.Tables(0).Rows.Count > 0 Then
            For count As Integer = 0 To ds.Tables(0).Rows.Count - 1
                dgv.Rows.Add()
                dgv.Rows(dgv.Rows.Count - 1).Cells(0).Value = ds.Tables(0).Rows(count).Item(0).ToString
                dgv.Rows(dgv.Rows.Count - 1).Cells(1).Value = ds.Tables(0).Rows(count).Item(1).ToString
                dgv.Rows(dgv.Rows.Count - 1).Cells(2).Value = ds.Tables(0).Rows(count).Item(2).ToString
                dgv.Rows(dgv.Rows.Count - 1).Cells(3).Value = ds.Tables(0).Rows(count).Item(3).ToString
                dgv.Rows(dgv.Rows.Count - 1).Cells(4).Value = "FPM"
            Next
        End If

        ds.Clear()
        ds = getDataset("select DoneWhat,Who,TransactionDate,CustomerID from AuditTrail where CustomerID = '" & tbcustomerid.Text.Trim & "' and DoneWhat like 'Employer Changed%' or DoneWhat like 'Group Changed%'", conRMOrange)
        If ds.Tables(0).Rows.Count > 0 Then
            For count As Integer = 0 To ds.Tables(0).Rows.Count - 1
                dgv.Rows.Add()
                dgv.Rows(dgv.Rows.Count - 1).Cells(0).Value = ds.Tables(0).Rows(count).Item(0).ToString
                dgv.Rows(dgv.Rows.Count - 1).Cells(1).Value = ds.Tables(0).Rows(count).Item(1).ToString
                dgv.Rows(dgv.Rows.Count - 1).Cells(2).Value = ds.Tables(0).Rows(count).Item(2).ToString
                dgv.Rows(dgv.Rows.Count - 1).Cells(3).Value = ds.Tables(0).Rows(count).Item(3).ToString
                dgv.Rows(dgv.Rows.Count - 1).Cells(4).Value = "RM Orange"
            Next
        End If

        ds.Clear()
        ds = getDataset("select DoneWhat,Who,TransactionDate,CustomerID from AuditTrail where CustomerID = '" & tbcustomerid.Text.Trim & "' and DoneWhat like 'Employer Changed%' or DoneWhat like 'Group Changed%'", conRMBeMobile)
        If ds.Tables(0).Rows.Count > 0 Then
            For count As Integer = 0 To ds.Tables(0).Rows.Count - 1
                dgv.Rows.Add()
                dgv.Rows(dgv.Rows.Count - 1).Cells(0).Value = ds.Tables(0).Rows(count).Item(0).ToString
                dgv.Rows(dgv.Rows.Count - 1).Cells(1).Value = ds.Tables(0).Rows(count).Item(1).ToString
                dgv.Rows(dgv.Rows.Count - 1).Cells(2).Value = ds.Tables(0).Rows(count).Item(2).ToString
                dgv.Rows(dgv.Rows.Count - 1).Cells(3).Value = ds.Tables(0).Rows(count).Item(3).ToString
                dgv.Rows(dgv.Rows.Count - 1).Cells(4).Value = "RM Bemobile"
            Next
        End If

        ds.Clear()
        ds = getDataset("select DoneWhat,Who,TransactionDate,CustomerID from AuditTrail where CustomerID = '" & tbcustomerid.Text.Trim & "' and DoneWhat like 'Employer Changed%' or DoneWhat like 'Group Changed%'", conMLMAdvances)
        If ds.Tables(0).Rows.Count > 0 Then
            For count As Integer = 0 To ds.Tables(0).Rows.Count - 1
                dgv.Rows.Add()
                dgv.Rows(dgv.Rows.Count - 1).Cells(0).Value = ds.Tables(0).Rows(count).Item(0).ToString
                dgv.Rows(dgv.Rows.Count - 1).Cells(1).Value = ds.Tables(0).Rows(count).Item(1).ToString
                dgv.Rows(dgv.Rows.Count - 1).Cells(2).Value = ds.Tables(0).Rows(count).Item(2).ToString
                dgv.Rows(dgv.Rows.Count - 1).Cells(3).Value = ds.Tables(0).Rows(count).Item(3).ToString
                dgv.Rows(dgv.Rows.Count - 1).Cells(4).Value = "MLM Advances"
            Next
        End If

        ds.Clear()
        ds = getDataset("select DoneWhat,Who,TransactionDate,CustomerID from AuditTrail where CustomerID = '" & tbcustomerid.Text.Trim & "' and DoneWhat like 'Employer Changed%' or DoneWhat like 'Group Changed%'", conMLMEducational)
        If ds.Tables(0).Rows.Count > 0 Then
            For count As Integer = 0 To ds.Tables(0).Rows.Count - 1
                dgv.Rows.Add()
                dgv.Rows(dgv.Rows.Count - 1).Cells(0).Value = ds.Tables(0).Rows(count).Item(0).ToString
                dgv.Rows(dgv.Rows.Count - 1).Cells(1).Value = ds.Tables(0).Rows(count).Item(1).ToString
                dgv.Rows(dgv.Rows.Count - 1).Cells(2).Value = ds.Tables(0).Rows(count).Item(2).ToString
                dgv.Rows(dgv.Rows.Count - 1).Cells(3).Value = ds.Tables(0).Rows(count).Item(3).ToString
                dgv.Rows(dgv.Rows.Count - 1).Cells(4).Value = "MLM Educational"
            Next
        End If

    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        ExportToExcel(dgv, "")
    End Sub
End Class