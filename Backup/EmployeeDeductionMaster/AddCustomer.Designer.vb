<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AddCustomer
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.cbemployer = New System.Windows.Forms.ComboBox
        Me.tbemployername = New System.Windows.Forms.TextBox
        Me.tbemployerid = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.Button3 = New System.Windows.Forms.Button
        Me.tbcustomername = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.tbcustomerid = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.cbgender = New System.Windows.Forms.ComboBox
        Me.tbgender = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Button2 = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'cbemployer
        '
        Me.cbemployer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbemployer.FormattingEnabled = True
        Me.cbemployer.Location = New System.Drawing.Point(107, 74)
        Me.cbemployer.Name = "cbemployer"
        Me.cbemployer.Size = New System.Drawing.Size(158, 21)
        Me.cbemployer.TabIndex = 25
        '
        'tbemployername
        '
        Me.tbemployername.Enabled = False
        Me.tbemployername.Location = New System.Drawing.Point(271, 75)
        Me.tbemployername.Name = "tbemployername"
        Me.tbemployername.Size = New System.Drawing.Size(104, 20)
        Me.tbemployername.TabIndex = 24
        '
        'tbemployerid
        '
        Me.tbemployerid.Enabled = False
        Me.tbemployerid.Location = New System.Drawing.Point(381, 75)
        Me.tbemployerid.Name = "tbemployerid"
        Me.tbemployerid.Size = New System.Drawing.Size(48, 20)
        Me.tbemployerid.TabIndex = 23
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(18, 78)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(50, 13)
        Me.Label7.TabIndex = 22
        Me.Label7.Text = "Employer"
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(374, 139)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(96, 23)
        Me.Button3.TabIndex = 21
        Me.Button3.Text = "Save Customer"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'tbcustomername
        '
        Me.tbcustomername.Location = New System.Drawing.Point(107, 41)
        Me.tbcustomername.Name = "tbcustomername"
        Me.tbcustomername.Size = New System.Drawing.Size(220, 20)
        Me.tbcustomername.TabIndex = 18
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(18, 44)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(82, 13)
        Me.Label2.TabIndex = 17
        Me.Label2.Text = "Customer Name"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(252, 12)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 16
        Me.Button1.Text = "Get"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'tbcustomerid
        '
        Me.tbcustomerid.Location = New System.Drawing.Point(107, 15)
        Me.tbcustomerid.Name = "tbcustomerid"
        Me.tbcustomerid.Size = New System.Drawing.Size(100, 20)
        Me.tbcustomerid.TabIndex = 15
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(18, 19)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(65, 13)
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "Customer ID"
        '
        'cbgender
        '
        Me.cbgender.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbgender.FormattingEnabled = True
        Me.cbgender.Items.AddRange(New Object() {"Male", "Female"})
        Me.cbgender.Location = New System.Drawing.Point(107, 112)
        Me.cbgender.Name = "cbgender"
        Me.cbgender.Size = New System.Drawing.Size(158, 21)
        Me.cbgender.TabIndex = 28
        '
        'tbgender
        '
        Me.tbgender.Enabled = False
        Me.tbgender.Location = New System.Drawing.Point(271, 113)
        Me.tbgender.Name = "tbgender"
        Me.tbgender.Size = New System.Drawing.Size(104, 20)
        Me.tbgender.TabIndex = 27
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(18, 116)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(42, 13)
        Me.Label3.TabIndex = 26
        Me.Label3.Text = "Gender"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(111, 139)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(96, 23)
        Me.Button2.TabIndex = 29
        Me.Button2.Text = "Update"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'AddCustomer
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(482, 170)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.cbgender)
        Me.Controls.Add(Me.tbgender)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.cbemployer)
        Me.Controls.Add(Me.tbemployername)
        Me.Controls.Add(Me.tbemployerid)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.tbcustomername)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.tbcustomerid)
        Me.Controls.Add(Me.Label1)
        Me.Name = "AddCustomer"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "AddCustomer"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cbemployer As System.Windows.Forms.ComboBox
    Friend WithEvents tbemployername As System.Windows.Forms.TextBox
    Friend WithEvents tbemployerid As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents tbcustomername As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents tbcustomerid As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cbgender As System.Windows.Forms.ComboBox
    Friend WithEvents tbgender As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Button2 As System.Windows.Forms.Button
End Class
