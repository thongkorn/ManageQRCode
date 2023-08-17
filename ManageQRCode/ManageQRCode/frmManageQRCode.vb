' / --------------------------------------------------------------------------------
' / Developer : Mr.Surapon Yodsanga (Thongkorn Tubtimkrob)
' / eMail : mailto:thongkorn@hotmail.com
' / URL: http://www.g2gnet.com (Khon Kaen - Thailand)
' / Facebook: https://www.facebook.com/g2gnet (For Thailand)
' / Facebook: https://www.facebook.com/commonindy (Worldwide)
' / More Info: http://www.g2gnet.com/webboard
' /
' / Purpose: Create QRCode (2D) with ZXing.Net and keep data into DataBase.
' / Microsoft Visual Basic .NET (2010) & MS Access (2010+)
' /
' / This is open source code under @CopyLeft by Thongkorn Tubtimkrob.
' / You can modify and/or distribute without to inform the developer.
' / --------------------------------------------------------------------------------

'// ZXing.Net release download.
'// https://github.com/micjahn/ZXing.Net/releases

Imports ZXing
Imports ZXing.Common
Imports ZXing.QrCode
Imports ZXing.QrCode.Internal
Imports ZXing.Rendering
Imports System.IO
Imports MetroFramework
Imports System.Drawing.Imaging
Imports System.Drawing.Drawing2D
Imports System.Data.OleDb

Public Class frmManageQRCode
    Dim PK As Long   '// Primary Key
    Dim NewData As Boolean = False  '// Add (True) or Edit (False) Mode.
    '//
    Dim ImageLogo As String = strPathImages & "egglogo.png"

    Private Sub frmManageQRCode_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        'Me.Dispose()
        'GC.SuppressFinalize(Me)
        'Application.Exit()
    End Sub

    Private Sub frmManageQRCode_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Result As Byte = MessageBox.Show("Are you sure you want to exit the program?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)
        If Result = DialogResult.Yes Then
            Me.Dispose()
            If Conn.State = ConnectionState.Open Then Conn.Close()
            GC.SuppressFinalize(Me)
            Application.Exit()
        Else
            e.Cancel = True
        End If
    End Sub

    Private Sub frmManageQRCode_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '// Must set KeyPreView = True on form.
        Select Case e.KeyCode
            Case Keys.F2
                Call btnAdd_Click(sender, e)
            Case Keys.F3
                Call btnSave_Click(sender, e)
            Case Keys.F4
                Call btnDelete_Click(sender, e)
            Case Keys.Escape
                Call btnDelete_Click(sender, e)
        End Select
    End Sub

    ' / --------------------------------------------------------------------------------
    ' / S T A R T ... H E R E
    ' / --------------------------------------------------------------------------------
    Private Sub frmManageQRCode_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '// Initial MetroFramework UI.
        Me.Style = MetroColorStyle.Red
        Me.ShadowType = Forms.MetroFormShadowType.DropShadow
        Me.TextAlign = Forms.MetroFormTextAlign.Center
        '//
        Me.KeyPreview = True
        'txtLink.Text = "www.g2gnet.com"
        lblRecordCount.Text = ""
        '//
        Call ConnectDataBase()
        Call SetupDataGridView(dgvData)
        Call RetrieveData()
        Call NewMode()
        txtSearch.Focus()
        '//
        Dim MyTip As New ToolTip
        MyTip.SetToolTip(btnRefresh, "Show all records.")
    End Sub

    ' / --------------------------------------------------------------------------------
    ' / Collect all searches and impressions. Come in the same place
    ' / blnSearch = True, Show that the search results.
    ' / blnSearch is set default to False, Show all records.
    ' / --------------------------------------------------------------------------------
    Private Sub RetrieveData(Optional ByVal blnSearch As Boolean = False)
        strSQL = _
            " SELECT PK, URL, Description, DateAdded, Logo, Margin " & _
            " FROM QRCode "
        '// blnSearch = True for Serach
        If blnSearch Then
            strSQL = strSQL & _
                " WHERE " & _
                " [URL] " & " Like '%" & txtSearch.Text & "%'" & " OR " & _
                " [Description] " & " Like '%" & txtSearch.Text & "%'" & _
                " ORDER BY PK "
        Else
            strSQL = strSQL & " ORDER BY PK "
        End If
        '//
        Try
            If Conn.State = ConnectionState.Closed Then Conn.Open()
            Cmd = New OleDbCommand
            With Cmd
                .Connection = Conn
                .CommandText = strSQL
            End With
            DR = Cmd.ExecuteReader
            Dim i As Long
            While DR.Read
                '// Load data into DataGridView.
                With dgvData
                    .Rows.Add(i)
                    .Rows(i).Cells(0).Value = DR.Item("PK").ToString.Trim
                    .Rows(i).Cells(1).Value = DR.Item("URL").ToString.Trim
                    .Rows(i).Cells(2).Value = DR.Item("Description").ToString.Trim
                    .Rows(i).Cells(3).Value = Format(CDate(DR.Item("DateAdded").ToString), "dd/MM/yyyy")
                    .Rows(i).Cells(4).Value = CBool(DR.Item("Logo").ToString)
                    .Rows(i).Cells(5).Value = CInt(DR.Item("Margin").ToString)
                End With
                i = i + 1
            End While
            lblRecordCount.Text = "[Total : " & dgvData.RowCount & " records]"
            DR.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        '//
        txtSearch.Clear()
    End Sub

    ' / --------------------------------------------------------------------------------
    ' / Double click to edit item.
    ' / --------------------------------------------------------------------------------
    Private Sub dgvData_DoubleClick(sender As Object, e As System.EventArgs) Handles dgvData.DoubleClick
        '// If you add / edit information should be reminded before.
        If btnDelete.Text = "Cancel - Esc" Then
            txtURL.Focus()
            Exit Sub
        End If
        '//
        If dgvData.RowCount <= 0 Then Return
        '// Read the value of the focus row.
        Dim iRow As Integer = dgvData.CurrentRow.Index
        PK = dgvData.Item(0, iRow).Value  '// Keep Primary Key
        '// If you share a file, you need to refresh the data.
        strSQL = _
            " SELECT PK, URL, Description, DateAdded, Logo, Margin " & _
            " FROM QRCode " & _
            " WHERE PK = " & PK
        If Conn.State = ConnectionState.Closed Then Conn.Open()
        DA = New OleDbDataAdapter(strSQL, Conn)
        DS = New DataSet
        DA.Fill(DS)
        ' / --------------------------------------------------------------------------------
        With DS.Tables(0)
            '// Using Double quote "" for trap error null value
            txtURL.Text = "" & .Rows(0)("URL").Trim.ToString()
            txtDescription.Text = "" & .Rows(0)("Description").ToString.Trim
            dtpDateAdded.Value = "" & .Rows(0)("DateAdded").ToString()
            If .Rows(0)("Logo").ToString() Then chkLogo.Checked = True
            udMargin.Value = .Rows(0)("Margin").ToString()
            '// Create QRCode.
            Call txtURL_TextChanged(sender, e)
        End With
        DS.Dispose()
        DA.Dispose()
        '// Change to Edit Mode.
        NewData = False
        Call EditMode()
    End Sub

    ' / --------------------------------------------------------------------------------
    '// Event when URL data is change.
    ' / --------------------------------------------------------------------------------
    Private Sub txtURL_TextChanged(sender As Object, e As EventArgs) Handles txtURL.TextChanged
        If String.IsNullOrWhiteSpace(txtURL.Text) Then
            picBarcode.Image = Nothing
            Return
        End If
        '//
        Dim options As EncodingOptions = New QrCodeEncodingOptions
        With options
            .Margin = udMargin.Value
            .NoPadding = True
            .Width = picBarcode.Width
            .Height = picBarcode.Height
            '// If have logo, ZXing not support UTF-8.
            If chkLogo.Checked Then
                .Hints.Add(EncodeHintType.ERROR_CORRECTION, ErrorCorrectionLevel.H)
            Else
                .Hints.Add(EncodeHintType.CHARACTER_SET, "UTF-8")
            End If
            .PureBarcode = False
        End With
        '//
        Dim objWriter As BarcodeWriter = New BarcodeWriter With {
                .Format = BarcodeFormat.QR_CODE,
                .Options = options,
                .Renderer = New BitmapRenderer
            }
        picBarcode.Image = New Bitmap(objWriter.Write(txtURL.Text))
        picBarcode.SizeMode = PictureBoxSizeMode.StretchImage
        '// Add Logo.
        If chkLogo.Checked Then
            Dim bitmap As Bitmap = picBarcode.Image '// objWriter.Write(txtURL.Text.Trim)
            bitmap.MakeTransparent()
            '// LOGO Path.
            Dim logo As Bitmap = New Bitmap(ImageLogo)
            Dim g As Graphics = Graphics.FromImage(bitmap)
            With g
                .SmoothingMode = SmoothingMode.AntiAlias
                .InterpolationMode = InterpolationMode.HighQualityBicubic
                .PixelOffsetMode = PixelOffsetMode.HighQuality
                .DrawImage(logo, New Point((bitmap.Width - logo.Width) \ 2, (bitmap.Height - logo.Height) \ 2))
                .Flush()
            End With
            '// 
            picBarcode.Image = New Bitmap(bitmap)
        End If
    End Sub

    ' / --------------------------------------------------------------------------------
    '// Save image of QR Code.
    ' / --------------------------------------------------------------------------------
    Private Sub btnSaveQRCode_Click(sender As Object, e As EventArgs) Handles btnSaveQRCode.Click
        Dim dlgSaveFile As New SaveFileDialog
        With dlgSaveFile
            .Title = "Select images"
            .Filter = "JPEG Image (.jpg)|*.jpg|Png Image (.png)|*.png|Bitmap Image (.bmp)|*.bmp;"
            If chkLogo.Checked Then
                .FilterIndex = 2
            Else
                .FilterIndex = 1
            End If
            .RestoreDirectory = True
            .InitialDirectory = strPathImages
        End With
        '//
        If dlgSaveFile.ShowDialog() = DialogResult.OK Then
            Try
                '// Saves the Image via a FileStream created by the OpenFile method.
                Dim fs = CType(dlgSaveFile.OpenFile, FileStream)
                '// Saves the Image in the appropriate ImageFormat based upon the
                '// file type selected in the dialog box.
                Select Case dlgSaveFile.FilterIndex
                    Case 1
                        picBarcode.Image.Save(fs, ImageFormat.Jpeg)
                    Case 2
                        picBarcode.Image.Save(fs, ImageFormat.Png)
                    Case 3
                        picBarcode.Image.Save(fs, ImageFormat.Bmp)
                End Select
                fs.Close()
                MessageBox.Show("QR Code image has been saved.", "Report Status", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If
    End Sub

    ' / --------------------------------------------------------------------------------
    '// Decode QR Code.
    ' / --------------------------------------------------------------------------------
    Private Sub btnDecode_Click(sender As Object, e As EventArgs) Handles btnDecode.Click
        Dim dlgImage As OpenFileDialog = New OpenFileDialog()
        ' / Open File Dialog
        With dlgImage
            .InitialDirectory = strPathImages
            .Title = "Select images"
            .Filter = "Image files (*.jpg,*.png,*bmp) | *.jpg; *.png; *.bmp"
            .FilterIndex = 1
            .RestoreDirectory = True
        End With
        Try
            '// Select OK after Browse ...
            If dlgImage.ShowDialog() = DialogResult.OK Then
                Using FS As IO.FileStream = File.Open(dlgImage.FileName, FileMode.Open)
                    Dim bitmap As Bitmap = New Bitmap(FS)
                    Dim CurrentPicture As Image = CType(bitmap, Image)
                    picBarcode.Image = CurrentPicture
                    '// Decode
                    Dim objReader As BarcodeReader = New BarcodeReader()
                    Dim objResult As Result = objReader.Decode(New Bitmap(picBarcode.Image))
                    'Dim objResult As Result = objReader.Decode(bitmap)
                    If objResult IsNot Nothing Then
                        txtURL.Text = objResult.Text
                        txtDescription.Text = ""
                        dtpDateAdded.Value = Now()
                        Call EditMode()
                        NewData = True
                    Else
                        MessageBox.Show("Cannot decode this image!")
                        Call NewMode()
                    End If
                End Using
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    ' / --------------------------------------------------------------------------------
    '// Copy to clipboard.
    ' / --------------------------------------------------------------------------------
    Private Sub btnCopyClipboard_Click(sender As System.Object, e As System.EventArgs) Handles btnCopyClipboard.Click
        If picBarcode.Image Is Nothing Then
            MessageBox.Show("There is no QR Code.", "Report Status", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return
        End If
        '/ Add it as an image
        Clipboard.SetImage(picBarcode.Image)
        '/ Create a JPG on disk and add the location to the clipboard
        Dim TempName As String = "TempName.jpg"
        Dim TempPath As String = System.IO.Path.Combine(My.Computer.FileSystem.SpecialDirectories.Temp, TempName)
        Using FS As New System.IO.FileStream(TempPath, IO.FileMode.Create, IO.FileAccess.Write, IO.FileShare.Read)
            picBarcode.Image.Save(FS, System.Drawing.Imaging.ImageFormat.Png)
        End Using
        Dim Paths As New System.Collections.Specialized.StringCollection()
        Paths.Add(TempPath)
        Clipboard.SetFileDropList(Paths)
    End Sub

    ' / --------------------------------------------------------------------------------
    '// Setup and Initialized DataGridView.
    ' / --------------------------------------------------------------------------------
    Public Sub SetupDataGridView(ByRef DGV As DataGridView)
        With DGV
            .RowHeadersVisible = True
            .AllowUserToAddRows = False
            .AllowUserToDeleteRows = False
            .AllowUserToResizeRows = False
            .MultiSelect = False
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .ReadOnly = True
        End With
        '// Columns Specified.
        With DGV.Columns
            .Add("PK", "Primary Key")           '// Index = 0
            .Add("URL", "URL or Link")
            .Add("Description", "Description")
            .Add("DateAdded", "Date")
            .Add("Logo", "Logo")                '// Index = 4
            .Add("Margin", "Margin")            '// Index = 5
        End With
        DGV.Columns(0).Visible = False
        DGV.Columns(4).Visible = False
        DGV.Columns(5).Visible = False
        '//
        With DGV
            .Font = New Font("Tahoma", 10)
            .RowTemplate.MinimumHeight = 28
            .RowTemplate.Height = 28
            '// Column Header
            .ColumnHeadersHeight = 30
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            '// Even-Odd Color
            .AlternatingRowsDefaultCellStyle.BackColor = Color.LightYellow
            With .ColumnHeadersDefaultCellStyle
                .BackColor = Color.Navy
                .ForeColor = Color.Black
                .Font = New Font("Tahoma", 10, FontStyle.Bold)
            End With
            '// Autosize Column
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            '// Change ForeColor of each Cell
            .DefaultCellStyle.ForeColor = Color.Black
            '// Change back color of each row
            .RowsDefaultCellStyle.BackColor = Color.AliceBlue
            '// Change GridLine Color
            .GridColor = Color.Blue
            '// Change Grid Border Style
            '.BorderStyle = BorderStyle.Fixed3D '// Can't use for MetroFramework UI.
        End With
    End Sub

    ' / --------------------------------------------------------------------------------
    '// Drag and Drop column.
    ' / --------------------------------------------------------------------------------
    Private Sub dgvData_DragDrop(sender As Object, e As System.Windows.Forms.DragEventArgs) Handles dgvData.DragDrop
        '// Just to Show a mouse icon to denote drop is allowed here.
        e.Effect = DragDropEffects.Move
    End Sub

    ' / --------------------------------------------------------------------------------
    ' / Add New Mode
    ' / --------------------------------------------------------------------------------
    Private Sub NewMode()
        txtURL.Text = "" : txtURL.Enabled = False
        txtDescription.Text = "" : txtDescription.Enabled = False
        dtpDateAdded.Enabled = False : dtpDateAdded.Value = Now()
        udMargin.Enabled = False
        udMargin.Value = 1
        chkLogo.Enabled = False
        chkLogo.Checked = False
        '//
        btnAdd.Enabled = True
        btnSave.Enabled = False
        btnSaveQRCode.Enabled = False
        btnCopyClipboard.Enabled = False
        btnDelete.Enabled = True
        btnDelete.Text = "Delete - F4"
        btnExit.Enabled = True
    End Sub

    ' / --------------------------------------------------------------------------------
    ' / Edit Data Mode
    ' / --------------------------------------------------------------------------------
    Private Sub EditMode()
        For Each c In GroupBox1.Controls
            If TypeOf c Is TextBox Then
                DirectCast(c, TextBox).Enabled = True
            End If
        Next
        dtpDateAdded.Enabled = True
        udMargin.Enabled = True
        chkLogo.Enabled = True
        '//
        btnAdd.Enabled = False
        btnSave.Enabled = True
        btnSaveQRCode.Enabled = True
        btnCopyClipboard.Enabled = True
        btnDelete.Enabled = True
        btnDelete.Text = "Cancel - Esc"
        btnExit.Enabled = False
        txtURL.Focus()
    End Sub

    ' / --------------------------------------------------------------------------------
    ' / Add Mode
    ' / --------------------------------------------------------------------------------
    Private Sub btnAdd_Click(sender As System.Object, e As System.EventArgs) Handles btnAdd.Click
        NewData = True  '// Add New Mode
        Call EditMode()
        txtURL.Focus()
    End Sub

    ' / --------------------------------------------------------------------------------
    Private Sub btnSave_Click(sender As System.Object, e As System.EventArgs) Handles btnSave.Click
        '// VALIDATE DATA.
        If txtURL.Text.Trim = "" Or IsNothing(txtURL.Text.Trim) Or txtURL.Text.Trim.Length = 0 Then
            MessageBox.Show("URL or Link cannot be empty.", "Report Status", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtURL.Focus()
            Exit Sub
        End If
        '// Call sub routine for UPDATE Record.
        Call SaveData()
    End Sub

    ' / --------------------------------------------------------------------------------
    ' / With both INSERT & UPDATE.
    ' / --------------------------------------------------------------------------------
    Private Sub SaveData()
        '// NewData = True, It's Add New Mode
        If NewData Then
            '// Call to Function "SetupNewPK" in the modDatabase.vb
            PK = SetupNewPK("SELECT MAX(QRCode.PK) AS MaxPK FROM QRCode")
            strSQL = _
                " INSERT INTO QRCode(" & _
                " PK, URL, Description, DateAdded, Logo, Margin) " & _
                " VALUES (" & _
                " @QPK, @QURL, @QDESC, @DADD, @LG, @MG " & _
                ")"
            '// EDIT MODE
        Else
            strSQL = _
                " UPDATE QRCode SET " & _
                " PK= @QPK, URL = QURL, Description = @QDESC, DateAdded = @DADD, " & _
                " Logo = @LG, Margin = @MG " & _
                " WHERE PK= @QPK"
        End If
        '// START
        Try
            If Conn.State = ConnectionState.Closed Then Conn.Open()
            Cmd = New OleDbCommand
            With Cmd.Parameters
                .AddWithValue("@QPK", PK)
                .AddWithValue("@QURL", txtURL.Text.Trim)
                .AddWithValue("@QDESC", txtDescription.Text.Trim)
                .AddWithValue("@DADD", Format(Cdate(dtpDateAdded.Value), "dd/MM/yyyy"))
                .AddWithValue("@LG", chkLogo.Checked)
                .AddWithValue("@MG", Val(udMargin.Value))
            End With
            '//
            With Cmd
                .Connection = Conn
                .CommandType = CommandType.Text
                .CommandText = strSQL
                .ExecuteNonQuery()
            End With
            '// Processing ...
            MessageBox.Show("Records Updated Completed.", "Update Status", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Cmd.Parameters.Clear()
            Cmd.Dispose()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        '//
        Call NewMode()
        dgvData.Rows.Clear()    '// Clear rows in DataGridView.
        Call RetrieveData()
        txtSearch.Focus()
    End Sub

    ' / --------------------------------------------------------------------------------
    '// Delete record.
    ' / --------------------------------------------------------------------------------
    Private Sub btnDelete_Click(sender As System.Object, e As System.EventArgs) Handles btnDelete.Click
        '// If Edit Data Mode
        If btnDelete.Text = "Cancel - Esc" Then
            btnAdd.Enabled = True
            btnSave.Enabled = True
            btnDelete.Enabled = True
            btnDelete.Text = "Delete - F4"
            btnExit.Enabled = True
            chkLogo.Checked = False
        Else
            If dgvData.RowCount = 0 Then Exit Sub
            '// Receive Primary Key value to confirm the deletion.
            Dim iRow As Long = dgvData.Item(0, dgvData.CurrentRow.Index).Value
            Dim URL As String = dgvData.Item(1, dgvData.CurrentRow.Index).Value
            Dim Result As Byte = MessageBox.Show("Are you sure you want to delete the data?" & vbCrLf & "URL: " & URL, "Confirm Deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)
            If Result = DialogResult.Yes Then
                '// iRow is the ContactPK or Primary key that is hidden.
                strSQL = " DELETE FROM QRCode WHERE PK = " & iRow
                If Conn.State = ConnectionState.Closed Then Conn.Open()
                '// UPDATE RECORD
                Cmd = New OleDbCommand
                With Cmd
                    .Connection = Conn
                    .CommandType = CommandType.Text
                    .CommandText = strSQL
                    .ExecuteNonQuery()
                    .Dispose()
                End With
            End If
        End If
        '//
        Call NewMode()
        dgvData.Rows.Clear()
        Call RetrieveData()
    End Sub

    Private Sub lblBrowseLogo_LinkClicked(sender As System.Object, e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lblBrowseLogo.LinkClicked
        Dim dlgImage As OpenFileDialog = New OpenFileDialog()
        ' / Open File Dialog
        With dlgImage
            .InitialDirectory = strPathImages
            .Title = "Select images"
            .Filter = "Image files (*.png,*gif) | *.png; *.gif"
            .FilterIndex = 1
            .RestoreDirectory = True
        End With
        Try
            ' Select OK after Browse ...
            If dlgImage.ShowDialog() = DialogResult.OK Then
                ImageLogo = dlgImage.FileName
                If chkLogo.Checked Then Call txtURL_TextChanged(sender, e)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub btnRefresh_Click(sender As System.Object, e As System.EventArgs) Handles btnRefresh.Click
        dgvData.Rows.Clear()
        Call RetrieveData()
    End Sub

    Private Sub udMargin_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles udMargin.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub

    ' / --------------------------------------------------------------------------------
    ' / Set Margin of PictureBox.
    ' / --------------------------------------------------------------------------------
    Private Sub udMargin_ValueChanged(sender As System.Object, e As System.EventArgs) Handles udMargin.ValueChanged
        '// Get new margin value and update QR Code.
        Call txtURL_TextChanged(sender, e)
    End Sub

    ' / --------------------------------------------------------------------------------
    ' / To check logo.
    ' / --------------------------------------------------------------------------------
    Private Sub chkLogo_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkLogo.CheckedChanged
        Call txtURL_TextChanged(sender, e)
    End Sub

    ' / --------------------------------------------------------------------------------
    ' / Event Press Enter to Search Data.
    ' / --------------------------------------------------------------------------------
    Private Sub txtSearch_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtSearch.KeyPress
        '// Undesirable characters for the database ex.  ', * or %
        txtSearch.Text = txtSearch.Text.Trim.Replace("'", "").Replace("%", "").Replace("*", "")
        If Trim(txtSearch.Text) = "" Or Len(Trim(txtSearch.Text)) = 0 Then Exit Sub
        If e.KeyChar = Chr(13) Then '// Press Enter
            '// No beep.
            e.Handled = True
            dgvData.Rows.Clear()
            '// RetrieveData(True) It means searching for information.
            Call RetrieveData(True)
        End If
    End Sub

    Private Sub btnExit_Click(sender As System.Object, e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub dgvData_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles dgvData.KeyDown
        If e.KeyCode = Keys.Enter Then
            Call dgvData_DoubleClick(sender, e)
            e.SuppressKeyPress = True
        End If
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As System.Object, e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Process.Start("https://www.facebook.com/g2gnet")
    End Sub

    Private Sub chkLogo_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles chkLogo.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub txtDescription_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles txtDescription.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub dtpDateAdded_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles dtpDateAdded.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub

End Class

