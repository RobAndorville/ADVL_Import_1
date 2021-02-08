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
        Me.components = New System.ComponentModel.Container()
        Dim XmlHtmDisplaySettings2 As ADVL_Utilities_Library_1.XmlHtmDisplaySettings = New ADVL_Utilities_Library_1.XmlHtmDisplaySettings()
        Dim TextSettings16 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim TextSettings17 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim TextSettings18 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim TextSettings19 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim TextSettings20 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim TextSettings21 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim TextSettings22 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim TextSettings23 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim TextSettings24 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim TextSettings25 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim TextSettings26 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim TextSettings27 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim TextSettings28 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim TextSettings29 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim TextSettings30 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Me.chkRecordSteps = New System.Windows.Forms.CheckBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtDescription = New System.Windows.Forms.TextBox()
        Me.txtName = New System.Windows.Forms.TextBox()
        Me.btnRun = New System.Windows.Forms.Button()
        Me.btnStatements = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnOpen = New System.Windows.Forms.Button()
        Me.btnNew = New System.Windows.Forms.Button()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.XmlHtmDisplay1 = New ADVL_Utilities_Library_1.XmlHtmDisplay(Me.components)
        Me.btnCancelImport = New System.Windows.Forms.Button()
        Me.btnFormat = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'chkRecordSteps
        '
        Me.chkRecordSteps.AutoSize = True
        Me.chkRecordSteps.Location = New System.Drawing.Point(12, 40)
        Me.chkRecordSteps.Name = "chkRecordSteps"
        Me.chkRecordSteps.Size = New System.Drawing.Size(146, 17)
        Me.chkRecordSteps.TabIndex = 67
        Me.chkRecordSteps.Text = "Record Processing Steps"
        Me.chkRecordSteps.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(9, 92)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(63, 13)
        Me.Label2.TabIndex = 66
        Me.Label2.Text = "Description:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 66)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(38, 13)
        Me.Label1.TabIndex = 65
        Me.Label1.Text = "Name:"
        '
        'txtDescription
        '
        Me.txtDescription.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDescription.Location = New System.Drawing.Point(82, 89)
        Me.txtDescription.Multiline = True
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.Size = New System.Drawing.Size(659, 51)
        Me.txtDescription.TabIndex = 64
        '
        'txtName
        '
        Me.txtName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtName.Location = New System.Drawing.Point(82, 63)
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(659, 20)
        Me.txtName.TabIndex = 63
        '
        'btnRun
        '
        Me.btnRun.Location = New System.Drawing.Point(316, 12)
        Me.btnRun.Name = "btnRun"
        Me.btnRun.Size = New System.Drawing.Size(52, 22)
        Me.btnRun.TabIndex = 62
        Me.btnRun.Text = "Run"
        Me.btnRun.UseVisualStyleBackColor = True
        '
        'btnStatements
        '
        Me.btnStatements.Location = New System.Drawing.Point(174, 12)
        Me.btnStatements.Name = "btnStatements"
        Me.btnStatements.Size = New System.Drawing.Size(82, 22)
        Me.btnStatements.TabIndex = 60
        Me.btnStatements.Text = "Statements"
        Me.btnStatements.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(120, 12)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(48, 22)
        Me.btnSave.TabIndex = 59
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnOpen
        '
        Me.btnOpen.Location = New System.Drawing.Point(12, 12)
        Me.btnOpen.Name = "btnOpen"
        Me.btnOpen.Size = New System.Drawing.Size(48, 22)
        Me.btnOpen.TabIndex = 58
        Me.btnOpen.Text = "Open"
        Me.btnOpen.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Location = New System.Drawing.Point(66, 12)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(48, 22)
        Me.btnNew.TabIndex = 57
        Me.btnNew.Text = "New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExit.Location = New System.Drawing.Point(677, 12)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(64, 22)
        Me.btnExit.TabIndex = 56
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'XmlHtmDisplay1
        '
        Me.XmlHtmDisplay1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.XmlHtmDisplay1.Font = New System.Drawing.Font("Arial", 10.0!)
        Me.XmlHtmDisplay1.Location = New System.Drawing.Point(12, 146)
        Me.XmlHtmDisplay1.Name = "XmlHtmDisplay1"
        TextSettings16.Bold = False
        TextSettings16.Color = System.Drawing.Color.Black
        TextSettings16.ColorIndex = 7
        TextSettings16.FontIndex = 7
        TextSettings16.FontName = "Arial"
        TextSettings16.HalfPointSize = 20
        TextSettings16.Italic = False
        TextSettings16.PointSize = 10.0!
        XmlHtmDisplaySettings2.DefaultText = TextSettings16
        TextSettings17.Bold = False
        TextSettings17.Color = System.Drawing.Color.Blue
        TextSettings17.ColorIndex = 4
        TextSettings17.FontIndex = 1
        TextSettings17.FontName = "Arial"
        TextSettings17.HalfPointSize = 20
        TextSettings17.Italic = False
        TextSettings17.PointSize = 10.0!
        XmlHtmDisplaySettings2.HAttribute = TextSettings17
        TextSettings18.Bold = False
        TextSettings18.Color = System.Drawing.Color.Gray
        TextSettings18.ColorIndex = 6
        TextSettings18.FontIndex = 1
        TextSettings18.FontName = "Arial"
        TextSettings18.HalfPointSize = 20
        TextSettings18.Italic = False
        TextSettings18.PointSize = 10.0!
        XmlHtmDisplaySettings2.HChar = TextSettings18
        TextSettings19.Bold = False
        TextSettings19.Color = System.Drawing.Color.Gray
        TextSettings19.ColorIndex = 6
        TextSettings19.FontIndex = 1
        TextSettings19.FontName = "Arial"
        TextSettings19.HalfPointSize = 20
        TextSettings19.Italic = False
        TextSettings19.PointSize = 10.0!
        XmlHtmDisplaySettings2.HComment = TextSettings19
        TextSettings20.Bold = False
        TextSettings20.Color = System.Drawing.Color.DarkRed
        TextSettings20.ColorIndex = 2
        TextSettings20.FontIndex = 1
        TextSettings20.FontName = "Arial"
        TextSettings20.HalfPointSize = 20
        TextSettings20.Italic = False
        TextSettings20.PointSize = 10.0!
        XmlHtmDisplaySettings2.HElement = TextSettings20
        TextSettings21.Bold = False
        TextSettings21.Color = System.Drawing.Color.Black
        TextSettings21.ColorIndex = 5
        TextSettings21.FontIndex = 1
        TextSettings21.FontName = "Arial"
        TextSettings21.HalfPointSize = 20
        TextSettings21.Italic = False
        TextSettings21.PointSize = 10.0!
        XmlHtmDisplaySettings2.HStyle = TextSettings21
        TextSettings22.Bold = False
        TextSettings22.Color = System.Drawing.Color.Black
        TextSettings22.ColorIndex = 7
        TextSettings22.FontIndex = 7
        TextSettings22.FontName = "Arial"
        TextSettings22.HalfPointSize = 20
        TextSettings22.Italic = False
        TextSettings22.PointSize = 10.0!
        XmlHtmDisplaySettings2.HText = TextSettings22
        TextSettings23.Bold = False
        TextSettings23.Color = System.Drawing.Color.Black
        TextSettings23.ColorIndex = 5
        TextSettings23.FontIndex = 1
        TextSettings23.FontName = "Arial"
        TextSettings23.HalfPointSize = 20
        TextSettings23.Italic = False
        TextSettings23.PointSize = 10.0!
        XmlHtmDisplaySettings2.HValue = TextSettings23
        TextSettings24.Bold = False
        TextSettings24.Color = System.Drawing.Color.Black
        TextSettings24.ColorIndex = 7
        TextSettings24.FontIndex = 7
        TextSettings24.FontName = "Arial"
        TextSettings24.HalfPointSize = 20
        TextSettings24.Italic = False
        TextSettings24.PointSize = 10.0!
        XmlHtmDisplaySettings2.PlainText = TextSettings24
        TextSettings25.Bold = False
        TextSettings25.Color = System.Drawing.Color.Red
        TextSettings25.ColorIndex = 3
        TextSettings25.FontIndex = 1
        TextSettings25.FontName = "Arial"
        TextSettings25.HalfPointSize = 20
        TextSettings25.Italic = False
        TextSettings25.PointSize = 10.0!
        XmlHtmDisplaySettings2.XAttributeKey = TextSettings25
        TextSettings26.Bold = False
        TextSettings26.Color = System.Drawing.Color.Blue
        TextSettings26.ColorIndex = 4
        TextSettings26.FontIndex = 1
        TextSettings26.FontName = "Arial"
        TextSettings26.HalfPointSize = 20
        TextSettings26.Italic = False
        TextSettings26.PointSize = 10.0!
        XmlHtmDisplaySettings2.XAttributeValue = TextSettings26
        TextSettings27.Bold = False
        TextSettings27.Color = System.Drawing.Color.Gray
        TextSettings27.ColorIndex = 6
        TextSettings27.FontIndex = 1
        TextSettings27.FontName = "Arial"
        TextSettings27.HalfPointSize = 20
        TextSettings27.Italic = False
        TextSettings27.PointSize = 10.0!
        XmlHtmDisplaySettings2.XComment = TextSettings27
        TextSettings28.Bold = False
        TextSettings28.Color = System.Drawing.Color.DarkRed
        TextSettings28.ColorIndex = 2
        TextSettings28.FontIndex = 1
        TextSettings28.FontName = "Arial"
        TextSettings28.HalfPointSize = 20
        TextSettings28.Italic = False
        TextSettings28.PointSize = 10.0!
        XmlHtmDisplaySettings2.XElement = TextSettings28
        XmlHtmDisplaySettings2.XIndentSpaces = 4
        XmlHtmDisplaySettings2.XmlLargeFileSizeLimit = 1000000
        TextSettings29.Bold = False
        TextSettings29.Color = System.Drawing.Color.Blue
        TextSettings29.ColorIndex = 1
        TextSettings29.FontIndex = 1
        TextSettings29.FontName = "Arial"
        TextSettings29.HalfPointSize = 20
        TextSettings29.Italic = False
        TextSettings29.PointSize = 10.0!
        XmlHtmDisplaySettings2.XTag = TextSettings29
        TextSettings30.Bold = False
        TextSettings30.Color = System.Drawing.Color.Black
        TextSettings30.ColorIndex = 5
        TextSettings30.FontIndex = 1
        TextSettings30.FontName = "Arial"
        TextSettings30.HalfPointSize = 20
        TextSettings30.Italic = False
        TextSettings30.PointSize = 10.0!
        XmlHtmDisplaySettings2.XValue = TextSettings30
        Me.XmlHtmDisplay1.Settings = XmlHtmDisplaySettings2
        Me.XmlHtmDisplay1.Size = New System.Drawing.Size(729, 438)
        Me.XmlHtmDisplay1.TabIndex = 277
        Me.XmlHtmDisplay1.Text = ""
        '
        'btnCancelImport
        '
        Me.btnCancelImport.Location = New System.Drawing.Point(374, 12)
        Me.btnCancelImport.Name = "btnCancelImport"
        Me.btnCancelImport.Size = New System.Drawing.Size(87, 22)
        Me.btnCancelImport.TabIndex = 278
        Me.btnCancelImport.Text = "Cancel Import"
        Me.btnCancelImport.UseVisualStyleBackColor = True
        '
        'btnFormat
        '
        Me.btnFormat.Location = New System.Drawing.Point(262, 12)
        Me.btnFormat.Name = "btnFormat"
        Me.btnFormat.Size = New System.Drawing.Size(48, 22)
        Me.btnFormat.TabIndex = 279
        Me.btnFormat.Text = "Format"
        Me.btnFormat.UseVisualStyleBackColor = True
        '
        'frmImportSequence
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(753, 596)
        Me.Controls.Add(Me.btnFormat)
        Me.Controls.Add(Me.btnCancelImport)
        Me.Controls.Add(Me.XmlHtmDisplay1)
        Me.Controls.Add(Me.chkRecordSteps)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtDescription)
        Me.Controls.Add(Me.txtName)
        Me.Controls.Add(Me.btnRun)
        Me.Controls.Add(Me.btnStatements)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnOpen)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.btnExit)
        Me.Name = "frmImportSequence"
        Me.Text = "Edit Import Sequence"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents chkRecordSteps As CheckBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents txtDescription As TextBox
    Friend WithEvents txtName As TextBox
    Friend WithEvents btnRun As Button
    Friend WithEvents btnStatements As Button
    Friend WithEvents btnSave As Button
    Friend WithEvents btnOpen As Button
    Friend WithEvents btnNew As Button
    Friend WithEvents btnExit As Button
    Friend WithEvents XmlHtmDisplay1 As ADVL_Utilities_Library_1.XmlHtmDisplay
    Friend WithEvents btnCancelImport As Button
    Friend WithEvents btnFormat As Button
End Class
