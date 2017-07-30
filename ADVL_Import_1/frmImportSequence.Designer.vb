<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmImportSequence
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
        Me.btnExit = New System.Windows.Forms.Button()
        Me.btnRun = New System.Windows.Forms.Button()
        Me.btnStatusCheck = New System.Windows.Forms.Button()
        Me.btnStatements = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnOpen = New System.Windows.Forms.Button()
        Me.btnNew = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtDescription = New System.Windows.Forms.TextBox()
        Me.txtName = New System.Windows.Forms.TextBox()
        Me.rtbSequence = New System.Windows.Forms.RichTextBox()
        Me.chkRecordSteps = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'btnExit
        '
        Me.btnExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExit.Location = New System.Drawing.Point(573, 12)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(64, 22)
        Me.btnExit.TabIndex = 2
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnRun
        '
        Me.btnRun.Location = New System.Drawing.Point(407, 12)
        Me.btnRun.Name = "btnRun"
        Me.btnRun.Size = New System.Drawing.Size(52, 22)
        Me.btnRun.TabIndex = 49
        Me.btnRun.Text = "Run"
        Me.btnRun.UseVisualStyleBackColor = True
        '
        'btnStatusCheck
        '
        Me.btnStatusCheck.Location = New System.Drawing.Point(310, 12)
        Me.btnStatusCheck.Name = "btnStatusCheck"
        Me.btnStatusCheck.Size = New System.Drawing.Size(91, 22)
        Me.btnStatusCheck.TabIndex = 48
        Me.btnStatusCheck.Text = "Status Check"
        Me.btnStatusCheck.UseVisualStyleBackColor = True
        '
        'btnStatements
        '
        Me.btnStatements.Location = New System.Drawing.Point(222, 12)
        Me.btnStatements.Name = "btnStatements"
        Me.btnStatements.Size = New System.Drawing.Size(82, 22)
        Me.btnStatements.TabIndex = 47
        Me.btnStatements.Text = "Statements"
        Me.btnStatements.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(152, 12)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(64, 22)
        Me.btnSave.TabIndex = 46
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnOpen
        '
        Me.btnOpen.Location = New System.Drawing.Point(12, 12)
        Me.btnOpen.Name = "btnOpen"
        Me.btnOpen.Size = New System.Drawing.Size(64, 22)
        Me.btnOpen.TabIndex = 45
        Me.btnOpen.Text = "Open"
        Me.btnOpen.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Location = New System.Drawing.Point(82, 12)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(64, 22)
        Me.btnNew.TabIndex = 44
        Me.btnNew.Text = "New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(9, 92)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(63, 13)
        Me.Label2.TabIndex = 53
        Me.Label2.Text = "Description:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 66)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(38, 13)
        Me.Label1.TabIndex = 52
        Me.Label1.Text = "Name:"
        '
        'txtDescription
        '
        Me.txtDescription.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDescription.Location = New System.Drawing.Point(82, 89)
        Me.txtDescription.Multiline = True
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.Size = New System.Drawing.Size(555, 51)
        Me.txtDescription.TabIndex = 51
        '
        'txtName
        '
        Me.txtName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtName.Location = New System.Drawing.Point(82, 63)
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(555, 20)
        Me.txtName.TabIndex = 50
        '
        'rtbSequence
        '
        Me.rtbSequence.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.rtbSequence.Location = New System.Drawing.Point(12, 148)
        Me.rtbSequence.Name = "rtbSequence"
        Me.rtbSequence.Size = New System.Drawing.Size(625, 330)
        Me.rtbSequence.TabIndex = 54
        Me.rtbSequence.Text = ""
        '
        'chkRecordSteps
        '
        Me.chkRecordSteps.AutoSize = True
        Me.chkRecordSteps.Location = New System.Drawing.Point(12, 40)
        Me.chkRecordSteps.Name = "chkRecordSteps"
        Me.chkRecordSteps.Size = New System.Drawing.Size(146, 17)
        Me.chkRecordSteps.TabIndex = 55
        Me.chkRecordSteps.Text = "Record Processing Steps"
        Me.chkRecordSteps.UseVisualStyleBackColor = True
        '
        'frmImportSequence
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(649, 490)
        Me.Controls.Add(Me.chkRecordSteps)
        Me.Controls.Add(Me.rtbSequence)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtDescription)
        Me.Controls.Add(Me.txtName)
        Me.Controls.Add(Me.btnRun)
        Me.Controls.Add(Me.btnStatusCheck)
        Me.Controls.Add(Me.btnStatements)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnOpen)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.btnExit)
        Me.Name = "frmImportSequence"
        Me.Text = "Import Sequence"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnExit As Button
    Friend WithEvents btnRun As Button
    Friend WithEvents btnStatusCheck As Button
    Friend WithEvents btnStatements As Button
    Friend WithEvents btnSave As Button
    Friend WithEvents btnOpen As Button
    Friend WithEvents btnNew As Button
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents txtDescription As TextBox
    Friend WithEvents txtName As TextBox
    Friend WithEvents rtbSequence As RichTextBox
    Friend WithEvents chkRecordSteps As CheckBox
End Class
