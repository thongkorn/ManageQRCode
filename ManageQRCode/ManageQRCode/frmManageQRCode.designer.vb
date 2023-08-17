<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmManageQRCode
    'Inherits System.Windows.Forms.Form
    Inherits MetroFramework.Forms.MetroForm

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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmManageQRCode))
        Me.label1 = New System.Windows.Forms.Label()
        Me.btnDecode = New System.Windows.Forms.Button()
        Me.btnSaveQRCode = New System.Windows.Forms.Button()
        Me.picBarcode = New System.Windows.Forms.PictureBox()
        Me.txtDescription = New System.Windows.Forms.TextBox()
        Me.btnCopyClipboard = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtURL = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnAdd = New System.Windows.Forms.Button()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.btnRefresh = New System.Windows.Forms.Button()
        Me.lblDesc = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.lblRecordCount = New System.Windows.Forms.Label()
        Me.txtSearch = New System.Windows.Forms.TextBox()
        Me.dgvData = New System.Windows.Forms.DataGridView()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.lblBrowseLogo = New System.Windows.Forms.LinkLabel()
        Me.chkLogo = New System.Windows.Forms.CheckBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.udMargin = New System.Windows.Forms.NumericUpDown()
        Me.dtpDateAdded = New System.Windows.Forms.DateTimePicker()
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel()
        CType(Me.picBarcode, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        CType(Me.dgvData, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        CType(Me.udMargin, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'label1
        '
        Me.label1.AutoSize = True
        Me.label1.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.label1.Location = New System.Drawing.Point(6, 20)
        Me.label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.label1.Name = "label1"
        Me.label1.Size = New System.Drawing.Size(100, 18)
        Me.label1.TabIndex = 9
        Me.label1.Text = "URL or Link:"
        '
        'btnDecode
        '
        Me.btnDecode.BackColor = System.Drawing.Color.YellowGreen
        Me.btnDecode.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnDecode.FlatAppearance.BorderSize = 0
        Me.btnDecode.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnDecode.ForeColor = System.Drawing.Color.Black
        Me.btnDecode.Location = New System.Drawing.Point(24, 582)
        Me.btnDecode.Margin = New System.Windows.Forms.Padding(4)
        Me.btnDecode.Name = "btnDecode"
        Me.btnDecode.Size = New System.Drawing.Size(247, 32)
        Me.btnDecode.TabIndex = 7
        Me.btnDecode.Text = "Load Image QR Code"
        Me.btnDecode.UseVisualStyleBackColor = False
        '
        'btnSaveQRCode
        '
        Me.btnSaveQRCode.BackColor = System.Drawing.Color.OrangeRed
        Me.btnSaveQRCode.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnSaveQRCode.FlatAppearance.BorderSize = 0
        Me.btnSaveQRCode.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnSaveQRCode.ForeColor = System.Drawing.Color.White
        Me.btnSaveQRCode.Location = New System.Drawing.Point(24, 545)
        Me.btnSaveQRCode.Margin = New System.Windows.Forms.Padding(4)
        Me.btnSaveQRCode.Name = "btnSaveQRCode"
        Me.btnSaveQRCode.Size = New System.Drawing.Size(247, 32)
        Me.btnSaveQRCode.TabIndex = 6
        Me.btnSaveQRCode.Text = "Save Image QR Code"
        Me.btnSaveQRCode.UseVisualStyleBackColor = False
        '
        'picBarcode
        '
        Me.picBarcode.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picBarcode.Location = New System.Drawing.Point(9, 232)
        Me.picBarcode.Margin = New System.Windows.Forms.Padding(4)
        Me.picBarcode.Name = "picBarcode"
        Me.picBarcode.Size = New System.Drawing.Size(247, 206)
        Me.picBarcode.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picBarcode.TabIndex = 6
        Me.picBarcode.TabStop = False
        '
        'txtDescription
        '
        Me.txtDescription.Location = New System.Drawing.Point(9, 148)
        Me.txtDescription.Margin = New System.Windows.Forms.Padding(4)
        Me.txtDescription.MaxLength = 255
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.Size = New System.Drawing.Size(247, 26)
        Me.txtDescription.TabIndex = 1
        '
        'btnCopyClipboard
        '
        Me.btnCopyClipboard.BackColor = System.Drawing.Color.SlateBlue
        Me.btnCopyClipboard.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnCopyClipboard.FlatAppearance.BorderSize = 0
        Me.btnCopyClipboard.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnCopyClipboard.ForeColor = System.Drawing.Color.White
        Me.btnCopyClipboard.Location = New System.Drawing.Point(24, 618)
        Me.btnCopyClipboard.Margin = New System.Windows.Forms.Padding(4)
        Me.btnCopyClipboard.Name = "btnCopyClipboard"
        Me.btnCopyClipboard.Size = New System.Drawing.Size(247, 32)
        Me.btnCopyClipboard.TabIndex = 8
        Me.btnCopyClipboard.Text = "Copy Image to Clipboard"
        Me.btnCopyClipboard.UseVisualStyleBackColor = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label2.Location = New System.Drawing.Point(6, 126)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(99, 18)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "Description:"
        '
        'txtURL
        '
        Me.txtURL.Location = New System.Drawing.Point(9, 42)
        Me.txtURL.Margin = New System.Windows.Forms.Padding(4)
        Me.txtURL.MaxLength = 255
        Me.txtURL.Multiline = True
        Me.txtURL.Name = "txtURL"
        Me.txtURL.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtURL.Size = New System.Drawing.Size(247, 80)
        Me.txtURL.TabIndex = 0
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label3.Location = New System.Drawing.Point(6, 178)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(48, 18)
        Me.Label3.TabIndex = 14
        Me.Label3.Text = "Date:"
        '
        'btnExit
        '
        Me.btnExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnExit.BackColor = System.Drawing.SystemColors.Control
        Me.btnExit.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnExit.FlatAppearance.BorderSize = 0
        Me.btnExit.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnExit.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.btnExit.ForeColor = System.Drawing.Color.Black
        Me.btnExit.Location = New System.Drawing.Point(716, 619)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(137, 32)
        Me.btnExit.TabIndex = 5
        Me.btnExit.Text = "E&xit"
        Me.btnExit.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnExit.UseVisualStyleBackColor = False
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnDelete.BackColor = System.Drawing.SystemColors.Control
        Me.btnDelete.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnDelete.FlatAppearance.BorderSize = 0
        Me.btnDelete.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnDelete.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.btnDelete.ForeColor = System.Drawing.Color.Black
        Me.btnDelete.Location = New System.Drawing.Point(573, 619)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(137, 32)
        Me.btnDelete.TabIndex = 4
        Me.btnDelete.Text = "Delete - F4"
        Me.btnDelete.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnDelete.UseVisualStyleBackColor = False
        '
        'btnSave
        '
        Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnSave.BackColor = System.Drawing.SystemColors.Control
        Me.btnSave.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnSave.FlatAppearance.BorderSize = 0
        Me.btnSave.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnSave.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.btnSave.ForeColor = System.Drawing.Color.Black
        Me.btnSave.Location = New System.Drawing.Point(430, 619)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(137, 32)
        Me.btnSave.TabIndex = 3
        Me.btnSave.Text = "Save - F3"
        Me.btnSave.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnSave.UseVisualStyleBackColor = False
        '
        'btnAdd
        '
        Me.btnAdd.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnAdd.BackColor = System.Drawing.SystemColors.Control
        Me.btnAdd.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnAdd.FlatAppearance.BorderSize = 0
        Me.btnAdd.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnAdd.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.btnAdd.ForeColor = System.Drawing.Color.Black
        Me.btnAdd.Location = New System.Drawing.Point(287, 619)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(137, 32)
        Me.btnAdd.TabIndex = 2
        Me.btnAdd.Text = "Add New - F2"
        Me.btnAdd.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnAdd.UseVisualStyleBackColor = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.btnRefresh)
        Me.GroupBox2.Controls.Add(Me.lblDesc)
        Me.GroupBox2.Controls.Add(Me.Label12)
        Me.GroupBox2.Controls.Add(Me.lblRecordCount)
        Me.GroupBox2.Controls.Add(Me.txtSearch)
        Me.GroupBox2.Controls.Add(Me.dgvData)
        Me.GroupBox2.Location = New System.Drawing.Point(283, 55)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(690, 561)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = " Listing QR Code "
        '
        'btnRefresh
        '
        Me.btnRefresh.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnRefresh.Image = CType(resources.GetObject("btnRefresh.Image"), System.Drawing.Image)
        Me.btnRefresh.Location = New System.Drawing.Point(237, 22)
        Me.btnRefresh.Name = "btnRefresh"
        Me.btnRefresh.Size = New System.Drawing.Size(24, 24)
        Me.btnRefresh.TabIndex = 1
        Me.btnRefresh.UseVisualStyleBackColor = True
        '
        'lblDesc
        '
        Me.lblDesc.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblDesc.AutoSize = True
        Me.lblDesc.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblDesc.ForeColor = System.Drawing.Color.Red
        Me.lblDesc.Location = New System.Drawing.Point(499, 28)
        Me.lblDesc.Name = "lblDesc"
        Me.lblDesc.Size = New System.Drawing.Size(187, 13)
        Me.lblDesc.TabIndex = 56
        Me.lblDesc.Text = "Double click the mouse to select items"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label12.ForeColor = System.Drawing.Color.Black
        Me.Label12.Location = New System.Drawing.Point(12, 24)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(62, 18)
        Me.Label12.TabIndex = 55
        Me.Label12.Text = "Search :"
        '
        'lblRecordCount
        '
        Me.lblRecordCount.AutoSize = True
        Me.lblRecordCount.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblRecordCount.Location = New System.Drawing.Point(261, 26)
        Me.lblRecordCount.Name = "lblRecordCount"
        Me.lblRecordCount.Size = New System.Drawing.Size(17, 14)
        Me.lblRecordCount.TabIndex = 54
        Me.lblRecordCount.Text = "[]"
        '
        'txtSearch
        '
        Me.txtSearch.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.txtSearch.Location = New System.Drawing.Point(80, 23)
        Me.txtSearch.Name = "txtSearch"
        Me.txtSearch.Size = New System.Drawing.Size(157, 22)
        Me.txtSearch.TabIndex = 0
        '
        'dgvData
        '
        Me.dgvData.AllowUserToAddRows = False
        Me.dgvData.AllowUserToDeleteRows = False
        Me.dgvData.AllowUserToOrderColumns = True
        Me.dgvData.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvData.Location = New System.Drawing.Point(4, 51)
        Me.dgvData.Name = "dgvData"
        Me.dgvData.ReadOnly = True
        Me.dgvData.RowHeadersVisible = False
        Me.dgvData.RowTemplate.Height = 28
        Me.dgvData.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvData.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvData.Size = New System.Drawing.Size(681, 504)
        Me.dgvData.TabIndex = 2
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.lblBrowseLogo)
        Me.GroupBox1.Controls.Add(Me.chkLogo)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.udMargin)
        Me.GroupBox1.Controls.Add(Me.picBarcode)
        Me.GroupBox1.Controls.Add(Me.txtDescription)
        Me.GroupBox1.Controls.Add(Me.label1)
        Me.GroupBox1.Controls.Add(Me.txtURL)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.dtpDateAdded)
        Me.GroupBox1.Location = New System.Drawing.Point(17, 55)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(264, 483)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Data "
        '
        'lblBrowseLogo
        '
        Me.lblBrowseLogo.AutoSize = True
        Me.lblBrowseLogo.Cursor = System.Windows.Forms.Cursors.Hand
        Me.lblBrowseLogo.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblBrowseLogo.Location = New System.Drawing.Point(173, 464)
        Me.lblBrowseLogo.Name = "lblBrowseLogo"
        Me.lblBrowseLogo.Size = New System.Drawing.Size(81, 16)
        Me.lblBrowseLogo.TabIndex = 60
        Me.lblBrowseLogo.TabStop = True
        Me.lblBrowseLogo.Text = "Browse Logo"
        '
        'chkLogo
        '
        Me.chkLogo.AutoSize = True
        Me.chkLogo.Location = New System.Drawing.Point(170, 446)
        Me.chkLogo.Name = "chkLogo"
        Me.chkLogo.Size = New System.Drawing.Size(88, 22)
        Me.chkLogo.TabIndex = 4
        Me.chkLogo.Text = "Add Logo"
        Me.chkLogo.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(10, 447)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(56, 18)
        Me.Label4.TabIndex = 59
        Me.Label4.Text = "Margin:"
        '
        'udMargin
        '
        Me.udMargin.Location = New System.Drawing.Point(68, 445)
        Me.udMargin.Name = "udMargin"
        Me.udMargin.Size = New System.Drawing.Size(53, 26)
        Me.udMargin.TabIndex = 3
        Me.udMargin.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.udMargin.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'dtpDateAdded
        '
        Me.dtpDateAdded.Location = New System.Drawing.Point(9, 199)
        Me.dtpDateAdded.Name = "dtpDateAdded"
        Me.dtpDateAdded.Size = New System.Drawing.Size(247, 26)
        Me.dtpDateAdded.TabIndex = 2
        '
        'LinkLabel1
        '
        Me.LinkLabel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Cursor = System.Windows.Forms.Cursors.Hand
        Me.LinkLabel1.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.LinkLabel1.LinkVisited = True
        Me.LinkLabel1.Location = New System.Drawing.Point(891, 626)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(78, 14)
        Me.LinkLabel1.TabIndex = 10
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "My Facebook"
        Me.LinkLabel1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'frmManageQRCode
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 18.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(988, 657)
        Me.Controls.Add(Me.LinkLabel1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnAdd)
        Me.Controls.Add(Me.btnCopyClipboard)
        Me.Controls.Add(Me.btnDecode)
        Me.Controls.Add(Me.btnSaveQRCode)
        Me.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MinimumSize = New System.Drawing.Size(988, 657)
        Me.Name = "frmManageQRCode"
        Me.ShadowType = MetroFramework.Forms.MetroFormShadowType.DropShadow
        Me.Style = MetroFramework.MetroColorStyle.Red
        Me.Text = "QR Code with ZXing.Net - coDe bY: Thongkorn Tubtimkrob"
        CType(Me.picBarcode, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.dgvData, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.udMargin, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Private WithEvents label1 As Label
    Private WithEvents btnDecode As Button
    Private WithEvents btnSaveQRCode As Button
    Private WithEvents picBarcode As PictureBox
    Private WithEvents txtDescription As TextBox
    Private WithEvents btnCopyClipboard As System.Windows.Forms.Button
    Private WithEvents Label2 As System.Windows.Forms.Label
    Private WithEvents txtURL As System.Windows.Forms.TextBox
    Private WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents btnRefresh As System.Windows.Forms.Button
    Friend WithEvents lblDesc As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents lblRecordCount As System.Windows.Forms.Label
    Friend WithEvents txtSearch As System.Windows.Forms.TextBox
    Friend WithEvents dgvData As System.Windows.Forms.DataGridView
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Private WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents udMargin As System.Windows.Forms.NumericUpDown
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents chkLogo As System.Windows.Forms.CheckBox
    Friend WithEvents dtpDateAdded As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblBrowseLogo As System.Windows.Forms.LinkLabel
End Class
