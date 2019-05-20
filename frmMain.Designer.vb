<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.btnLoginInfo = New System.Windows.Forms.Button()
        Me.btnRegistration = New System.Windows.Forms.Button()
        Me.btnProducts = New System.Windows.Forms.Button()
        Me.btnRMReport = New System.Windows.Forms.Button()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.btnProductsReport = New System.Windows.Forms.Button()
        Me.btnRM = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btnLoginInfo)
        Me.GroupBox1.Controls.Add(Me.btnRegistration)
        Me.GroupBox1.Controls.Add(Me.btnProducts)
        Me.GroupBox1.Controls.Add(Me.btnRMReport)
        Me.GroupBox1.Controls.Add(Me.btnExit)
        Me.GroupBox1.Controls.Add(Me.btnProductsReport)
        Me.GroupBox1.Controls.Add(Me.btnRM)
        Me.GroupBox1.Location = New System.Drawing.Point(37, 29)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(239, 344)
        Me.GroupBox1.TabIndex = 23
        Me.GroupBox1.TabStop = False
        '
        'btnLoginInfo
        '
        Me.btnLoginInfo.BackColor = System.Drawing.Color.White
        Me.btnLoginInfo.Font = New System.Drawing.Font("Palatino Linotype", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnLoginInfo.ForeColor = System.Drawing.Color.Maroon
        Me.btnLoginInfo.Location = New System.Drawing.Point(31, 67)
        Me.btnLoginInfo.Name = "btnLoginInfo"
        Me.btnLoginInfo.Size = New System.Drawing.Size(186, 38)
        Me.btnLoginInfo.TabIndex = 30
        Me.btnLoginInfo.Text = "Login Info"
        Me.btnLoginInfo.UseVisualStyleBackColor = False
        '
        'btnRegistration
        '
        Me.btnRegistration.BackColor = System.Drawing.Color.White
        Me.btnRegistration.Font = New System.Drawing.Font("Palatino Linotype", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRegistration.ForeColor = System.Drawing.Color.Maroon
        Me.btnRegistration.Location = New System.Drawing.Point(31, 23)
        Me.btnRegistration.Name = "btnRegistration"
        Me.btnRegistration.Size = New System.Drawing.Size(186, 38)
        Me.btnRegistration.TabIndex = 29
        Me.btnRegistration.Text = "Registration"
        Me.btnRegistration.UseVisualStyleBackColor = False
        '
        'btnProducts
        '
        Me.btnProducts.BackColor = System.Drawing.Color.White
        Me.btnProducts.Font = New System.Drawing.Font("Palatino Linotype", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnProducts.ForeColor = System.Drawing.Color.Maroon
        Me.btnProducts.Location = New System.Drawing.Point(29, 155)
        Me.btnProducts.Name = "btnProducts"
        Me.btnProducts.Size = New System.Drawing.Size(186, 38)
        Me.btnProducts.TabIndex = 28
        Me.btnProducts.Text = "Products"
        Me.btnProducts.UseVisualStyleBackColor = False
        '
        'btnRMReport
        '
        Me.btnRMReport.BackColor = System.Drawing.Color.White
        Me.btnRMReport.Font = New System.Drawing.Font("Palatino Linotype", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRMReport.ForeColor = System.Drawing.Color.Maroon
        Me.btnRMReport.Location = New System.Drawing.Point(31, 199)
        Me.btnRMReport.Name = "btnRMReport"
        Me.btnRMReport.Size = New System.Drawing.Size(186, 38)
        Me.btnRMReport.TabIndex = 27
        Me.btnRMReport.Text = "Raw Materials Record"
        Me.btnRMReport.UseVisualStyleBackColor = False
        '
        'btnExit
        '
        Me.btnExit.BackColor = System.Drawing.Color.White
        Me.btnExit.Font = New System.Drawing.Font("Palatino Linotype", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnExit.ForeColor = System.Drawing.Color.Maroon
        Me.btnExit.Location = New System.Drawing.Point(31, 287)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(186, 38)
        Me.btnExit.TabIndex = 25
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = False
        '
        'btnProductsReport
        '
        Me.btnProductsReport.BackColor = System.Drawing.Color.White
        Me.btnProductsReport.Font = New System.Drawing.Font("Palatino Linotype", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnProductsReport.ForeColor = System.Drawing.Color.Maroon
        Me.btnProductsReport.Location = New System.Drawing.Point(33, 243)
        Me.btnProductsReport.Name = "btnProductsReport"
        Me.btnProductsReport.Size = New System.Drawing.Size(184, 38)
        Me.btnProductsReport.TabIndex = 24
        Me.btnProductsReport.Text = "Products Record"
        Me.btnProductsReport.UseVisualStyleBackColor = False
        '
        'btnRM
        '
        Me.btnRM.BackColor = System.Drawing.Color.White
        Me.btnRM.Font = New System.Drawing.Font("Palatino Linotype", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRM.ForeColor = System.Drawing.Color.Maroon
        Me.btnRM.Location = New System.Drawing.Point(29, 111)
        Me.btnRM.Name = "btnRM"
        Me.btnRM.Size = New System.Drawing.Size(186, 38)
        Me.btnRM.TabIndex = 23
        Me.btnRM.Text = "Raw Materials"
        Me.btnRM.UseVisualStyleBackColor = False
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.ClientSize = New System.Drawing.Size(311, 406)
        Me.Controls.Add(Me.GroupBox1)
        Me.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Main Menu [PM Soft]"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents btnProducts As System.Windows.Forms.Button
    Friend WithEvents btnRMReport As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnProductsReport As System.Windows.Forms.Button
    Friend WithEvents btnRM As System.Windows.Forms.Button
    Friend WithEvents btnLoginInfo As System.Windows.Forms.Button
    Friend WithEvents btnRegistration As System.Windows.Forms.Button

End Class
