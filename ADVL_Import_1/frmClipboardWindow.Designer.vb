<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmClipboardWindow
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
        Me.txtClipboard = New System.Windows.Forms.TextBox()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.btnClear = New System.Windows.Forms.Button()
        Me.btnPaste = New System.Windows.Forms.Button()
        Me.btnCharCodes = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'txtClipboard
        '
        Me.txtClipboard.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtClipboard.Location = New System.Drawing.Point(12, 40)
        Me.txtClipboard.Multiline = True
        Me.txtClipboard.Name = "txtClipboard"
        Me.txtClipboard.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtClipboard.Size = New System.Drawing.Size(729, 537)
        Me.txtClipboard.TabIndex = 0
        '
        'btnExit
        '
        Me.btnExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExit.Location = New System.Drawing.Point(677, 12)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(64, 22)
        Me.btnExit.TabIndex = 8
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnClear
        '
        Me.btnClear.Location = New System.Drawing.Point(12, 12)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(58, 22)
        Me.btnClear.TabIndex = 9
        Me.btnClear.Text = "Clear"
        Me.btnClear.UseVisualStyleBackColor = True
        '
        'btnPaste
        '
        Me.btnPaste.Location = New System.Drawing.Point(76, 12)
        Me.btnPaste.Name = "btnPaste"
        Me.btnPaste.Size = New System.Drawing.Size(125, 22)
        Me.btnPaste.TabIndex = 10
        Me.btnPaste.Text = "Paste from Clipboard"
        Me.btnPaste.UseVisualStyleBackColor = True
        '
        'btnCharCodes
        '
        Me.btnCharCodes.Location = New System.Drawing.Point(207, 12)
        Me.btnCharCodes.Name = "btnCharCodes"
        Me.btnCharCodes.Size = New System.Drawing.Size(134, 22)
        Me.btnCharCodes.TabIndex = 11
        Me.btnCharCodes.Text = "Show Character Codes"
        Me.btnCharCodes.UseVisualStyleBackColor = True
        '
        'frmClipboardWindow
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(753, 589)
        Me.Controls.Add(Me.btnCharCodes)
        Me.Controls.Add(Me.btnPaste)
        Me.Controls.Add(Me.btnClear)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.txtClipboard)
        Me.Name = "frmClipboardWindow"
        Me.Text = "Clipboard Window"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtClipboard As TextBox
    Friend WithEvents btnExit As Button
    Friend WithEvents btnClear As Button
    Friend WithEvents btnPaste As Button
    Friend WithEvents btnCharCodes As Button
End Class
