<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmUnitInputs
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.Evaporator = New System.Windows.Forms.TabPage
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.cmdEvapSave = New System.Windows.Forms.Button
        Me.cmdEvapDelete = New System.Windows.Forms.Button
        Me.cmdEvapUpdate = New System.Windows.Forms.Button
        Me.cmdEvapAdd = New System.Windows.Forms.Button
        Me.chkEvapDroptubes = New System.Windows.Forms.CheckBox
        Me.lblDroptubes = New System.Windows.Forms.Label
        Me.cboEvapSplit = New System.Windows.Forms.ComboBox
        Me.lblEvapSplits = New System.Windows.Forms.Label
        Me.txtEvapCKT = New System.Windows.Forms.TextBox
        Me.lblEvapCircuits = New System.Windows.Forms.Label
        Me.txtEvapCKT2 = New System.Windows.Forms.TextBox
        Me.lblEvapCKT2 = New System.Windows.Forms.Label
        Me.txtEvapCKT1 = New System.Windows.Forms.TextBox
        Me.lblEvapCKT1 = New System.Windows.Forms.Label
        Me.cboEvapWallThk = New System.Windows.Forms.ComboBox
        Me.lblWallThk = New System.Windows.Forms.Label
        Me.cboEvapFinPI = New System.Windows.Forms.ComboBox
        Me.lblFinPI = New System.Windows.Forms.Label
        Me.cboEvapFinMat = New System.Windows.Forms.ComboBox
        Me.lblEvapFinMat = New System.Windows.Forms.Label
        Me.cboEvapFinThk = New System.Windows.Forms.ComboBox
        Me.lblFinThk = New System.Windows.Forms.Label
        Me.txtEvapFL = New System.Windows.Forms.TextBox
        Me.lblFinLength = New System.Windows.Forms.Label
        Me.txtEvapFH = New System.Windows.Forms.TextBox
        Me.lblEvapFH = New System.Windows.Forms.Label
        Me.cboEvapRows = New System.Windows.Forms.ComboBox
        Me.lblRows = New System.Windows.Forms.Label
        Me.cboEvapCoilP = New System.Windows.Forms.ComboBox
        Me.lblCoilPattern = New System.Windows.Forms.Label
        Me.ddEvapModel = New System.Windows.Forms.ComboBox
        Me.lblEvapModel = New System.Windows.Forms.Label
        Me.Condenser = New System.Windows.Forms.TabPage
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.cmdCondDelete = New System.Windows.Forms.Button
        Me.cmdCondUpdate = New System.Windows.Forms.Button
        Me.cmdCondSave = New System.Windows.Forms.Button
        Me.cmdCondAdd = New System.Windows.Forms.Button
        Me.chkCondDroptubes = New System.Windows.Forms.CheckBox
        Me.lblCondDroptubes = New System.Windows.Forms.Label
        Me.cboCondSplit = New System.Windows.Forms.ComboBox
        Me.lblCondSplits = New System.Windows.Forms.Label
        Me.lblCondCircuits = New System.Windows.Forms.Label
        Me.txtCondCKT = New System.Windows.Forms.TextBox
        Me.txtCondCKT2 = New System.Windows.Forms.TextBox
        Me.lblCondCKT2 = New System.Windows.Forms.Label
        Me.txtCondCKT1 = New System.Windows.Forms.TextBox
        Me.lblCondCKT1 = New System.Windows.Forms.Label
        Me.cboCondWallThk = New System.Windows.Forms.ComboBox
        Me.lblCondWallThk = New System.Windows.Forms.Label
        Me.cboCondFinPI = New System.Windows.Forms.ComboBox
        Me.lblCondFinPI = New System.Windows.Forms.Label
        Me.cboCondFinMat = New System.Windows.Forms.ComboBox
        Me.lblCondFinMat = New System.Windows.Forms.Label
        Me.cboCondFinThk = New System.Windows.Forms.ComboBox
        Me.lblCondFinThk = New System.Windows.Forms.Label
        Me.txtCondFL = New System.Windows.Forms.TextBox
        Me.lblCondFL = New System.Windows.Forms.Label
        Me.txtCondFH = New System.Windows.Forms.TextBox
        Me.lblCondFH = New System.Windows.Forms.Label
        Me.cboCondRows = New System.Windows.Forms.ComboBox
        Me.lblCondRows = New System.Windows.Forms.Label
        Me.cboCondCoilP = New System.Windows.Forms.ComboBox
        Me.lblCondCoilP = New System.Windows.Forms.Label
        Me.ddCondModel = New System.Windows.Forms.ComboBox
        Me.lblCondModel = New System.Windows.Forms.Label
        Me.Compressor = New System.Windows.Forms.TabPage
        Me.CT_Max = New System.Windows.Forms.TextBox
        Me.lblCondMaxT = New System.Windows.Forms.Label
        Me.ET_Max = New System.Windows.Forms.TextBox
        Me.lblMaxEvapT = New System.Windows.Forms.Label
        Me.GroupBox6 = New System.Windows.Forms.GroupBox
        Me.optSales = New System.Windows.Forms.RadioButton
        Me.lblSalesReport = New System.Windows.Forms.Label
        Me.lblEngReport = New System.Windows.Forms.Label
        Me.OptEng = New System.Windows.Forms.RadioButton
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.cmdCompressorFile = New System.Windows.Forms.Button
        Me.txtSuperh = New System.Windows.Forms.TextBox
        Me.lblSuperh = New System.Windows.Forms.Label
        Me.txtSubcool = New System.Windows.Forms.TextBox
        Me.lblSubcool = New System.Windows.Forms.Label
        Me.cboAppCode = New System.Windows.Forms.ComboBox
        Me.lblCompAppCode = New System.Windows.Forms.Label
        Me.cboFreCode = New System.Windows.Forms.ComboBox
        Me.lblFreCode = New System.Windows.Forms.Label
        Me.cboRefCode = New System.Windows.Forms.ComboBox
        Me.lblRefCode = New System.Windows.Forms.Label
        Me.cboCompModel = New System.Windows.Forms.ComboBox
        Me.lblCompModel = New System.Windows.Forms.Label
        Me.Unit = New System.Windows.Forms.TabPage
        Me.GroupBox5 = New System.Windows.Forms.GroupBox
        Me.txtReheatLaDB = New System.Windows.Forms.TextBox
        Me.lblReheataDB = New System.Windows.Forms.Label
        Me.txtReheatCKT = New System.Windows.Forms.TextBox
        Me.lblReheatCKT = New System.Windows.Forms.Label
        Me.chkReheatDroptubes = New System.Windows.Forms.CheckBox
        Me.lblReDropTubes = New System.Windows.Forms.Label
        Me.txtReheatFH = New System.Windows.Forms.TextBox
        Me.lblReheatFH = New System.Windows.Forms.Label
        Me.cboReheatFinPI = New System.Windows.Forms.ComboBox
        Me.lblReheatFinPI = New System.Windows.Forms.Label
        Me.txtReheatFL = New System.Windows.Forms.TextBox
        Me.lblReheatFL = New System.Windows.Forms.Label
        Me.cboReheatWallThk = New System.Windows.Forms.ComboBox
        Me.lblReheatWallThk = New System.Windows.Forms.Label
        Me.cboReheatRows = New System.Windows.Forms.ComboBox
        Me.lblReheatRows = New System.Windows.Forms.Label
        Me.cboReheatFinThk = New System.Windows.Forms.ComboBox
        Me.lblReheatFinThk = New System.Windows.Forms.Label
        Me.cboReheatCoilP = New System.Windows.Forms.ComboBox
        Me.lblReheatCoilPat = New System.Windows.Forms.Label
        Me.cmdUnitDelete = New System.Windows.Forms.Button
        Me.cmdUnitUpdate = New System.Windows.Forms.Button
        Me.cmdUnitSave = New System.Windows.Forms.Button
        Me.cmdUnitAdd = New System.Windows.Forms.Button
        Me.cboUnitModel = New System.Windows.Forms.ComboBox
        Me.lblUnitModel = New System.Windows.Forms.Label
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.optHeatPump = New System.Windows.Forms.RadioButton
        Me.lblheatpump = New System.Windows.Forms.Label
        Me.lblaircond = New System.Windows.Forms.Label
        Me.optAirCond = New System.Windows.Forms.RadioButton
        Me.cmdCalculate = New System.Windows.Forms.Button
        Me.ddReheatOn_Off = New System.Windows.Forms.ComboBox
        Me.lblReheatCoil = New System.Windows.Forms.Label
        Me.cboCompQuantity = New System.Windows.Forms.ComboBox
        Me.lblCompQuantity = New System.Windows.Forms.Label
        Me.cboCondQuantity = New System.Windows.Forms.ComboBox
        Me.lblCondQuantity = New System.Windows.Forms.Label
        Me.cboEvapQuantity = New System.Windows.Forms.ComboBox
        Me.lblEvapQuantity = New System.Windows.Forms.Label
        Me.txtOutaCFM = New System.Windows.Forms.TextBox
        Me.lblOutAirCFM = New System.Windows.Forms.Label
        Me.txtAltitude = New System.Windows.Forms.TextBox
        Me.lblAltitude = New System.Windows.Forms.Label
        Me.cboRef = New System.Windows.Forms.ComboBox
        Me.lblRefrigerant = New System.Windows.Forms.Label
        Me.txtCoCFM = New System.Windows.Forms.TextBox
        Me.lblCoCFM = New System.Windows.Forms.Label
        Me.txtCoEaWB = New System.Windows.Forms.TextBox
        Me.lblCoEaWB = New System.Windows.Forms.Label
        Me.txtCoEaDB = New System.Windows.Forms.TextBox
        Me.lblCoEaDB = New System.Windows.Forms.Label
        Me.txtEvCFM = New System.Windows.Forms.TextBox
        Me.lblEvCFM = New System.Windows.Forms.Label
        Me.txtEvEaWB = New System.Windows.Forms.TextBox
        Me.lblEvEaWB = New System.Windows.Forms.Label
        Me.txtEvEaDB = New System.Windows.Forms.TextBox
        Me.lblEvapEaDB = New System.Windows.Forms.Label
        Me.FileDialog = New System.Windows.Forms.OpenFileDialog
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip
        Me.HelpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuFileHelp = New System.Windows.Forms.ToolStripMenuItem
        Me.TabControl1.SuspendLayout()
        Me.Evaporator.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.Condenser.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.Compressor.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.Unit.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.Evaporator)
        Me.TabControl1.Controls.Add(Me.Condenser)
        Me.TabControl1.Controls.Add(Me.Compressor)
        Me.TabControl1.Controls.Add(Me.Unit)
        Me.TabControl1.Location = New System.Drawing.Point(5, 26)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(868, 326)
        Me.TabControl1.TabIndex = 0
        '
        'Evaporator
        '
        Me.Evaporator.Controls.Add(Me.GroupBox1)
        Me.Evaporator.Location = New System.Drawing.Point(4, 22)
        Me.Evaporator.Name = "Evaporator"
        Me.Evaporator.Padding = New System.Windows.Forms.Padding(3)
        Me.Evaporator.Size = New System.Drawing.Size(860, 300)
        Me.Evaporator.TabIndex = 0
        Me.Evaporator.Text = "Evaporator"
        Me.Evaporator.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.cmdEvapSave)
        Me.GroupBox1.Controls.Add(Me.cmdEvapDelete)
        Me.GroupBox1.Controls.Add(Me.cmdEvapUpdate)
        Me.GroupBox1.Controls.Add(Me.cmdEvapAdd)
        Me.GroupBox1.Controls.Add(Me.chkEvapDroptubes)
        Me.GroupBox1.Controls.Add(Me.lblDroptubes)
        Me.GroupBox1.Controls.Add(Me.cboEvapSplit)
        Me.GroupBox1.Controls.Add(Me.lblEvapSplits)
        Me.GroupBox1.Controls.Add(Me.txtEvapCKT)
        Me.GroupBox1.Controls.Add(Me.lblEvapCircuits)
        Me.GroupBox1.Controls.Add(Me.txtEvapCKT2)
        Me.GroupBox1.Controls.Add(Me.lblEvapCKT2)
        Me.GroupBox1.Controls.Add(Me.txtEvapCKT1)
        Me.GroupBox1.Controls.Add(Me.lblEvapCKT1)
        Me.GroupBox1.Controls.Add(Me.cboEvapWallThk)
        Me.GroupBox1.Controls.Add(Me.lblWallThk)
        Me.GroupBox1.Controls.Add(Me.cboEvapFinPI)
        Me.GroupBox1.Controls.Add(Me.lblFinPI)
        Me.GroupBox1.Controls.Add(Me.cboEvapFinMat)
        Me.GroupBox1.Controls.Add(Me.lblEvapFinMat)
        Me.GroupBox1.Controls.Add(Me.cboEvapFinThk)
        Me.GroupBox1.Controls.Add(Me.lblFinThk)
        Me.GroupBox1.Controls.Add(Me.txtEvapFL)
        Me.GroupBox1.Controls.Add(Me.lblFinLength)
        Me.GroupBox1.Controls.Add(Me.txtEvapFH)
        Me.GroupBox1.Controls.Add(Me.lblEvapFH)
        Me.GroupBox1.Controls.Add(Me.cboEvapRows)
        Me.GroupBox1.Controls.Add(Me.lblRows)
        Me.GroupBox1.Controls.Add(Me.cboEvapCoilP)
        Me.GroupBox1.Controls.Add(Me.lblCoilPattern)
        Me.GroupBox1.Controls.Add(Me.ddEvapModel)
        Me.GroupBox1.Controls.Add(Me.lblEvapModel)
        Me.GroupBox1.Location = New System.Drawing.Point(7, 11)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(847, 229)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'cmdEvapSave
        '
        Me.cmdEvapSave.Location = New System.Drawing.Point(155, 187)
        Me.cmdEvapSave.Name = "cmdEvapSave"
        Me.cmdEvapSave.Size = New System.Drawing.Size(101, 24)
        Me.cmdEvapSave.TabIndex = 31
        Me.cmdEvapSave.Text = "Save New Evap"
        Me.cmdEvapSave.UseVisualStyleBackColor = True
        '
        'cmdEvapDelete
        '
        Me.cmdEvapDelete.Location = New System.Drawing.Point(412, 187)
        Me.cmdEvapDelete.Name = "cmdEvapDelete"
        Me.cmdEvapDelete.Size = New System.Drawing.Size(87, 24)
        Me.cmdEvapDelete.TabIndex = 30
        Me.cmdEvapDelete.Text = "Delete Evap"
        Me.cmdEvapDelete.UseVisualStyleBackColor = True
        '
        'cmdEvapUpdate
        '
        Me.cmdEvapUpdate.Location = New System.Drawing.Point(292, 187)
        Me.cmdEvapUpdate.Name = "cmdEvapUpdate"
        Me.cmdEvapUpdate.Size = New System.Drawing.Size(92, 24)
        Me.cmdEvapUpdate.TabIndex = 29
        Me.cmdEvapUpdate.Text = "Update Evap"
        Me.cmdEvapUpdate.UseVisualStyleBackColor = True
        '
        'cmdEvapAdd
        '
        Me.cmdEvapAdd.Location = New System.Drawing.Point(17, 187)
        Me.cmdEvapAdd.Name = "cmdEvapAdd"
        Me.cmdEvapAdd.Size = New System.Drawing.Size(94, 24)
        Me.cmdEvapAdd.TabIndex = 28
        Me.cmdEvapAdd.Text = "Add New Evap"
        Me.cmdEvapAdd.UseVisualStyleBackColor = True
        '
        'chkEvapDroptubes
        '
        Me.chkEvapDroptubes.AutoSize = True
        Me.chkEvapDroptubes.Location = New System.Drawing.Point(774, 86)
        Me.chkEvapDroptubes.Name = "chkEvapDroptubes"
        Me.chkEvapDroptubes.Size = New System.Drawing.Size(15, 14)
        Me.chkEvapDroptubes.TabIndex = 27
        Me.chkEvapDroptubes.UseVisualStyleBackColor = True
        '
        'lblDroptubes
        '
        Me.lblDroptubes.AutoSize = True
        Me.lblDroptubes.Location = New System.Drawing.Point(692, 84)
        Me.lblDroptubes.Name = "lblDroptubes"
        Me.lblDroptubes.Size = New System.Drawing.Size(69, 13)
        Me.lblDroptubes.TabIndex = 26
        Me.lblDroptubes.Text = "Drop Tubes?"
        '
        'cboEvapSplit
        '
        Me.cboEvapSplit.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboEvapSplit.FormattingEnabled = True
        Me.cboEvapSplit.Items.AddRange(New Object() {"IT", "FS", "NS"})
        Me.cboEvapSplit.Location = New System.Drawing.Point(774, 33)
        Me.cboEvapSplit.Name = "cboEvapSplit"
        Me.cboEvapSplit.Size = New System.Drawing.Size(58, 21)
        Me.cboEvapSplit.TabIndex = 25
        '
        'lblEvapSplits
        '
        Me.lblEvapSplits.AutoSize = True
        Me.lblEvapSplits.Location = New System.Drawing.Point(692, 35)
        Me.lblEvapSplits.Name = "lblEvapSplits"
        Me.lblEvapSplits.Size = New System.Drawing.Size(67, 13)
        Me.lblEvapSplits.TabIndex = 24
        Me.lblEvapSplits.Text = "lblEvapSplits"
        '
        'txtEvapCKT
        '
        Me.txtEvapCKT.Location = New System.Drawing.Point(621, 127)
        Me.txtEvapCKT.Name = "txtEvapCKT"
        Me.txtEvapCKT.Size = New System.Drawing.Size(53, 20)
        Me.txtEvapCKT.TabIndex = 23
        '
        'lblEvapCircuits
        '
        Me.lblEvapCircuits.AutoSize = True
        Me.lblEvapCircuits.Location = New System.Drawing.Point(517, 127)
        Me.lblEvapCircuits.Name = "lblEvapCircuits"
        Me.lblEvapCircuits.Size = New System.Drawing.Size(92, 13)
        Me.lblEvapCircuits.TabIndex = 22
        Me.lblEvapCircuits.Text = "Total No of Feeds"
        '
        'txtEvapCKT2
        '
        Me.txtEvapCKT2.Location = New System.Drawing.Point(619, 82)
        Me.txtEvapCKT2.Name = "txtEvapCKT2"
        Me.txtEvapCKT2.Size = New System.Drawing.Size(55, 20)
        Me.txtEvapCKT2.TabIndex = 21
        '
        'lblEvapCKT2
        '
        Me.lblEvapCKT2.AutoSize = True
        Me.lblEvapCKT2.Location = New System.Drawing.Point(517, 82)
        Me.lblEvapCKT2.Name = "lblEvapCKT2"
        Me.lblEvapCKT2.Size = New System.Drawing.Size(95, 13)
        Me.lblEvapCKT2.TabIndex = 20
        Me.lblEvapCKT2.Text = "CKT2 No of Feeds"
        '
        'txtEvapCKT1
        '
        Me.txtEvapCKT1.Location = New System.Drawing.Point(618, 33)
        Me.txtEvapCKT1.Name = "txtEvapCKT1"
        Me.txtEvapCKT1.Size = New System.Drawing.Size(56, 20)
        Me.txtEvapCKT1.TabIndex = 19
        '
        'lblEvapCKT1
        '
        Me.lblEvapCKT1.AutoSize = True
        Me.lblEvapCKT1.Location = New System.Drawing.Point(517, 33)
        Me.lblEvapCKT1.Name = "lblEvapCKT1"
        Me.lblEvapCKT1.Size = New System.Drawing.Size(95, 13)
        Me.lblEvapCKT1.TabIndex = 18
        Me.lblEvapCKT1.Text = "CKT1 No of Feeds"
        '
        'cboEvapWallThk
        '
        Me.cboEvapWallThk.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboEvapWallThk.FormattingEnabled = True
        Me.cboEvapWallThk.Items.AddRange(New Object() {"0.012", "0.014", "0.016", "0.017", "0.018", "0.025", "0.035", "0.049"})
        Me.cboEvapWallThk.Location = New System.Drawing.Point(444, 126)
        Me.cboEvapWallThk.Name = "cboEvapWallThk"
        Me.cboEvapWallThk.Size = New System.Drawing.Size(54, 21)
        Me.cboEvapWallThk.TabIndex = 17
        '
        'lblWallThk
        '
        Me.lblWallThk.AutoSize = True
        Me.lblWallThk.Location = New System.Drawing.Point(361, 126)
        Me.lblWallThk.Name = "lblWallThk"
        Me.lblWallThk.Size = New System.Drawing.Size(80, 13)
        Me.lblWallThk.TabIndex = 16
        Me.lblWallThk.Text = "Wall Thickness"
        '
        'cboEvapFinPI
        '
        Me.cboEvapFinPI.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboEvapFinPI.FormattingEnabled = True
        Me.cboEvapFinPI.Items.AddRange(New Object() {"6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20"})
        Me.cboEvapFinPI.Location = New System.Drawing.Point(443, 80)
        Me.cboEvapFinPI.Name = "cboEvapFinPI"
        Me.cboEvapFinPI.Size = New System.Drawing.Size(56, 21)
        Me.cboEvapFinPI.TabIndex = 15
        '
        'lblFinPI
        '
        Me.lblFinPI.AutoSize = True
        Me.lblFinPI.Location = New System.Drawing.Point(361, 81)
        Me.lblFinPI.Name = "lblFinPI"
        Me.lblFinPI.Size = New System.Drawing.Size(52, 13)
        Me.lblFinPI.TabIndex = 14
        Me.lblFinPI.Text = "Fins/Inch"
        '
        'cboEvapFinMat
        '
        Me.cboEvapFinMat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboEvapFinMat.FormattingEnabled = True
        Me.cboEvapFinMat.Items.AddRange(New Object() {"A", "C"})
        Me.cboEvapFinMat.Location = New System.Drawing.Point(444, 32)
        Me.cboEvapFinMat.Name = "cboEvapFinMat"
        Me.cboEvapFinMat.Size = New System.Drawing.Size(55, 21)
        Me.cboEvapFinMat.TabIndex = 13
        '
        'lblEvapFinMat
        '
        Me.lblEvapFinMat.AutoSize = True
        Me.lblEvapFinMat.Location = New System.Drawing.Point(361, 33)
        Me.lblEvapFinMat.Name = "lblEvapFinMat"
        Me.lblEvapFinMat.Size = New System.Drawing.Size(61, 13)
        Me.lblEvapFinMat.TabIndex = 12
        Me.lblEvapFinMat.Text = "Fin Material"
        '
        'cboEvapFinThk
        '
        Me.cboEvapFinThk.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboEvapFinThk.FormattingEnabled = True
        Me.cboEvapFinThk.Items.AddRange(New Object() {"0.0045", "0.0055", "0.0060", "0.0065", "0.0075", "0.0100"})
        Me.cboEvapFinThk.Location = New System.Drawing.Point(293, 123)
        Me.cboEvapFinThk.Name = "cboEvapFinThk"
        Me.cboEvapFinThk.Size = New System.Drawing.Size(62, 21)
        Me.cboEvapFinThk.TabIndex = 11
        '
        'lblFinThk
        '
        Me.lblFinThk.AutoSize = True
        Me.lblFinThk.Location = New System.Drawing.Point(214, 126)
        Me.lblFinThk.Name = "lblFinThk"
        Me.lblFinThk.Size = New System.Drawing.Size(73, 13)
        Me.lblFinThk.TabIndex = 10
        Me.lblFinThk.Text = "Fin Thickness"
        '
        'txtEvapFL
        '
        Me.txtEvapFL.Location = New System.Drawing.Point(292, 79)
        Me.txtEvapFL.Name = "txtEvapFL"
        Me.txtEvapFL.Size = New System.Drawing.Size(63, 20)
        Me.txtEvapFL.TabIndex = 9
        '
        'lblFinLength
        '
        Me.lblFinLength.AutoSize = True
        Me.lblFinLength.Location = New System.Drawing.Point(214, 79)
        Me.lblFinLength.Name = "lblFinLength"
        Me.lblFinLength.Size = New System.Drawing.Size(57, 13)
        Me.lblFinLength.TabIndex = 8
        Me.lblFinLength.Text = "Fin Length"
        '
        'txtEvapFH
        '
        Me.txtEvapFH.Location = New System.Drawing.Point(292, 30)
        Me.txtEvapFH.Name = "txtEvapFH"
        Me.txtEvapFH.Size = New System.Drawing.Size(63, 20)
        Me.txtEvapFH.TabIndex = 7
        '
        'lblEvapFH
        '
        Me.lblEvapFH.AutoSize = True
        Me.lblEvapFH.Location = New System.Drawing.Point(214, 30)
        Me.lblEvapFH.Name = "lblEvapFH"
        Me.lblEvapFH.Size = New System.Drawing.Size(55, 13)
        Me.lblEvapFH.TabIndex = 6
        Me.lblEvapFH.Text = "Fin Height"
        '
        'cboEvapRows
        '
        Me.cboEvapRows.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboEvapRows.FormattingEnabled = True
        Me.cboEvapRows.Items.AddRange(New Object() {"1", "2", "3", "4", "5", "6", "7", "8", "9", "10"})
        Me.cboEvapRows.Location = New System.Drawing.Point(112, 123)
        Me.cboEvapRows.Name = "cboEvapRows"
        Me.cboEvapRows.Size = New System.Drawing.Size(83, 21)
        Me.cboEvapRows.TabIndex = 5
        '
        'lblRows
        '
        Me.lblRows.AutoSize = True
        Me.lblRows.Location = New System.Drawing.Point(14, 126)
        Me.lblRows.Name = "lblRows"
        Me.lblRows.Size = New System.Drawing.Size(63, 13)
        Me.lblRows.TabIndex = 4
        Me.lblRows.Text = "No of Rows"
        '
        'cboEvapCoilP
        '
        Me.cboEvapCoilP.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboEvapCoilP.FormattingEnabled = True
        Me.cboEvapCoilP.Items.AddRange(New Object() {"7", "6", "3", "P", "5"})
        Me.cboEvapCoilP.Location = New System.Drawing.Point(112, 76)
        Me.cboEvapCoilP.Name = "cboEvapCoilP"
        Me.cboEvapCoilP.Size = New System.Drawing.Size(84, 21)
        Me.cboEvapCoilP.TabIndex = 3
        '
        'lblCoilPattern
        '
        Me.lblCoilPattern.AutoSize = True
        Me.lblCoilPattern.Location = New System.Drawing.Point(14, 81)
        Me.lblCoilPattern.Name = "lblCoilPattern"
        Me.lblCoilPattern.Size = New System.Drawing.Size(86, 13)
        Me.lblCoilPattern.TabIndex = 2
        Me.lblCoilPattern.Text = "EvapCoil Pattern"
        '
        'ddEvapModel
        '
        Me.ddEvapModel.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ddEvapModel.FormattingEnabled = True
        Me.ddEvapModel.Location = New System.Drawing.Point(111, 27)
        Me.ddEvapModel.Name = "ddEvapModel"
        Me.ddEvapModel.Size = New System.Drawing.Size(85, 21)
        Me.ddEvapModel.Sorted = True
        Me.ddEvapModel.TabIndex = 1
        '
        'lblEvapModel
        '
        Me.lblEvapModel.AutoSize = True
        Me.lblEvapModel.Location = New System.Drawing.Point(14, 30)
        Me.lblEvapModel.Name = "lblEvapModel"
        Me.lblEvapModel.Size = New System.Drawing.Size(91, 13)
        Me.lblEvapModel.TabIndex = 0
        Me.lblEvapModel.Text = "Evaporator Model"
        '
        'Condenser
        '
        Me.Condenser.Controls.Add(Me.GroupBox2)
        Me.Condenser.Location = New System.Drawing.Point(4, 22)
        Me.Condenser.Name = "Condenser"
        Me.Condenser.Padding = New System.Windows.Forms.Padding(3)
        Me.Condenser.Size = New System.Drawing.Size(860, 302)
        Me.Condenser.TabIndex = 1
        Me.Condenser.Text = "Condenser"
        Me.Condenser.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.cmdCondDelete)
        Me.GroupBox2.Controls.Add(Me.cmdCondUpdate)
        Me.GroupBox2.Controls.Add(Me.cmdCondSave)
        Me.GroupBox2.Controls.Add(Me.cmdCondAdd)
        Me.GroupBox2.Controls.Add(Me.chkCondDroptubes)
        Me.GroupBox2.Controls.Add(Me.lblCondDroptubes)
        Me.GroupBox2.Controls.Add(Me.cboCondSplit)
        Me.GroupBox2.Controls.Add(Me.lblCondSplits)
        Me.GroupBox2.Controls.Add(Me.lblCondCircuits)
        Me.GroupBox2.Controls.Add(Me.txtCondCKT)
        Me.GroupBox2.Controls.Add(Me.txtCondCKT2)
        Me.GroupBox2.Controls.Add(Me.lblCondCKT2)
        Me.GroupBox2.Controls.Add(Me.txtCondCKT1)
        Me.GroupBox2.Controls.Add(Me.lblCondCKT1)
        Me.GroupBox2.Controls.Add(Me.cboCondWallThk)
        Me.GroupBox2.Controls.Add(Me.lblCondWallThk)
        Me.GroupBox2.Controls.Add(Me.cboCondFinPI)
        Me.GroupBox2.Controls.Add(Me.lblCondFinPI)
        Me.GroupBox2.Controls.Add(Me.cboCondFinMat)
        Me.GroupBox2.Controls.Add(Me.lblCondFinMat)
        Me.GroupBox2.Controls.Add(Me.cboCondFinThk)
        Me.GroupBox2.Controls.Add(Me.lblCondFinThk)
        Me.GroupBox2.Controls.Add(Me.txtCondFL)
        Me.GroupBox2.Controls.Add(Me.lblCondFL)
        Me.GroupBox2.Controls.Add(Me.txtCondFH)
        Me.GroupBox2.Controls.Add(Me.lblCondFH)
        Me.GroupBox2.Controls.Add(Me.cboCondRows)
        Me.GroupBox2.Controls.Add(Me.lblCondRows)
        Me.GroupBox2.Controls.Add(Me.cboCondCoilP)
        Me.GroupBox2.Controls.Add(Me.lblCondCoilP)
        Me.GroupBox2.Controls.Add(Me.ddCondModel)
        Me.GroupBox2.Controls.Add(Me.lblCondModel)
        Me.GroupBox2.Location = New System.Drawing.Point(6, 8)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(846, 234)
        Me.GroupBox2.TabIndex = 0
        Me.GroupBox2.TabStop = False
        '
        'cmdCondDelete
        '
        Me.cmdCondDelete.Location = New System.Drawing.Point(426, 196)
        Me.cmdCondDelete.Name = "cmdCondDelete"
        Me.cmdCondDelete.Size = New System.Drawing.Size(90, 24)
        Me.cmdCondDelete.TabIndex = 31
        Me.cmdCondDelete.Text = "Delete Cond"
        Me.cmdCondDelete.UseVisualStyleBackColor = True
        '
        'cmdCondUpdate
        '
        Me.cmdCondUpdate.Location = New System.Drawing.Point(295, 196)
        Me.cmdCondUpdate.Name = "cmdCondUpdate"
        Me.cmdCondUpdate.Size = New System.Drawing.Size(96, 24)
        Me.cmdCondUpdate.TabIndex = 30
        Me.cmdCondUpdate.Text = "Update Cond"
        Me.cmdCondUpdate.UseVisualStyleBackColor = True
        '
        'cmdCondSave
        '
        Me.cmdCondSave.Location = New System.Drawing.Point(162, 196)
        Me.cmdCondSave.Name = "cmdCondSave"
        Me.cmdCondSave.Size = New System.Drawing.Size(96, 24)
        Me.cmdCondSave.TabIndex = 29
        Me.cmdCondSave.Text = "Save New Cond"
        Me.cmdCondSave.UseVisualStyleBackColor = True
        '
        'cmdCondAdd
        '
        Me.cmdCondAdd.Location = New System.Drawing.Point(21, 196)
        Me.cmdCondAdd.Name = "cmdCondAdd"
        Me.cmdCondAdd.Size = New System.Drawing.Size(98, 24)
        Me.cmdCondAdd.TabIndex = 28
        Me.cmdCondAdd.Text = "Add New Cond"
        Me.cmdCondAdd.UseVisualStyleBackColor = True
        '
        'chkCondDroptubes
        '
        Me.chkCondDroptubes.AutoSize = True
        Me.chkCondDroptubes.Location = New System.Drawing.Point(780, 88)
        Me.chkCondDroptubes.Name = "chkCondDroptubes"
        Me.chkCondDroptubes.Size = New System.Drawing.Size(15, 14)
        Me.chkCondDroptubes.TabIndex = 27
        Me.chkCondDroptubes.UseVisualStyleBackColor = True
        '
        'lblCondDroptubes
        '
        Me.lblCondDroptubes.AutoSize = True
        Me.lblCondDroptubes.Location = New System.Drawing.Point(705, 86)
        Me.lblCondDroptubes.Name = "lblCondDroptubes"
        Me.lblCondDroptubes.Size = New System.Drawing.Size(63, 13)
        Me.lblCondDroptubes.TabIndex = 26
        Me.lblCondDroptubes.Text = "Drop Tubes"
        '
        'cboCondSplit
        '
        Me.cboCondSplit.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCondSplit.FormattingEnabled = True
        Me.cboCondSplit.Items.AddRange(New Object() {"IT", "FS", "NS"})
        Me.cboCondSplit.Location = New System.Drawing.Point(781, 23)
        Me.cboCondSplit.Name = "cboCondSplit"
        Me.cboCondSplit.Size = New System.Drawing.Size(59, 21)
        Me.cboCondSplit.TabIndex = 25
        '
        'lblCondSplits
        '
        Me.lblCondSplits.AutoSize = True
        Me.lblCondSplits.Location = New System.Drawing.Point(705, 28)
        Me.lblCondSplits.Name = "lblCondSplits"
        Me.lblCondSplits.Size = New System.Drawing.Size(60, 13)
        Me.lblCondSplits.TabIndex = 24
        Me.lblCondSplits.Text = "Cond Splits"
        '
        'lblCondCircuits
        '
        Me.lblCondCircuits.AutoSize = True
        Me.lblCondCircuits.Location = New System.Drawing.Point(537, 142)
        Me.lblCondCircuits.Name = "lblCondCircuits"
        Me.lblCondCircuits.Size = New System.Drawing.Size(92, 13)
        Me.lblCondCircuits.TabIndex = 23
        Me.lblCondCircuits.Text = "Total No of Feeds"
        '
        'txtCondCKT
        '
        Me.txtCondCKT.Location = New System.Drawing.Point(638, 142)
        Me.txtCondCKT.Name = "txtCondCKT"
        Me.txtCondCKT.Size = New System.Drawing.Size(55, 20)
        Me.txtCondCKT.TabIndex = 22
        '
        'txtCondCKT2
        '
        Me.txtCondCKT2.Location = New System.Drawing.Point(638, 85)
        Me.txtCondCKT2.Name = "txtCondCKT2"
        Me.txtCondCKT2.Size = New System.Drawing.Size(55, 20)
        Me.txtCondCKT2.TabIndex = 21
        '
        'lblCondCKT2
        '
        Me.lblCondCKT2.AutoSize = True
        Me.lblCondCKT2.Location = New System.Drawing.Point(537, 86)
        Me.lblCondCKT2.Name = "lblCondCKT2"
        Me.lblCondCKT2.Size = New System.Drawing.Size(95, 13)
        Me.lblCondCKT2.TabIndex = 20
        Me.lblCondCKT2.Text = "CKT2 No of Feeds"
        '
        'txtCondCKT1
        '
        Me.txtCondCKT1.Location = New System.Drawing.Point(638, 23)
        Me.txtCondCKT1.Name = "txtCondCKT1"
        Me.txtCondCKT1.Size = New System.Drawing.Size(55, 20)
        Me.txtCondCKT1.TabIndex = 19
        '
        'lblCondCKT1
        '
        Me.lblCondCKT1.AutoSize = True
        Me.lblCondCKT1.Location = New System.Drawing.Point(537, 26)
        Me.lblCondCKT1.Name = "lblCondCKT1"
        Me.lblCondCKT1.Size = New System.Drawing.Size(95, 13)
        Me.lblCondCKT1.TabIndex = 18
        Me.lblCondCKT1.Text = "CKT1 No of Feeds"
        '
        'cboCondWallThk
        '
        Me.cboCondWallThk.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCondWallThk.FormattingEnabled = True
        Me.cboCondWallThk.Items.AddRange(New Object() {"0.012", "0.014", "0.016", "0.017", "0.018", "0.025", "0.035", "0.049"})
        Me.cboCondWallThk.Location = New System.Drawing.Point(460, 141)
        Me.cboCondWallThk.Name = "cboCondWallThk"
        Me.cboCondWallThk.Size = New System.Drawing.Size(56, 21)
        Me.cboCondWallThk.TabIndex = 17
        '
        'lblCondWallThk
        '
        Me.lblCondWallThk.AutoSize = True
        Me.lblCondWallThk.Location = New System.Drawing.Point(374, 142)
        Me.lblCondWallThk.Name = "lblCondWallThk"
        Me.lblCondWallThk.Size = New System.Drawing.Size(80, 13)
        Me.lblCondWallThk.TabIndex = 16
        Me.lblCondWallThk.Text = "Wall Thickness"
        '
        'cboCondFinPI
        '
        Me.cboCondFinPI.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCondFinPI.FormattingEnabled = True
        Me.cboCondFinPI.Items.AddRange(New Object() {"6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "20", "22", "24"})
        Me.cboCondFinPI.Location = New System.Drawing.Point(460, 85)
        Me.cboCondFinPI.Name = "cboCondFinPI"
        Me.cboCondFinPI.Size = New System.Drawing.Size(56, 21)
        Me.cboCondFinPI.TabIndex = 15
        '
        'lblCondFinPI
        '
        Me.lblCondFinPI.AutoSize = True
        Me.lblCondFinPI.Location = New System.Drawing.Point(374, 86)
        Me.lblCondFinPI.Name = "lblCondFinPI"
        Me.lblCondFinPI.Size = New System.Drawing.Size(52, 13)
        Me.lblCondFinPI.TabIndex = 14
        Me.lblCondFinPI.Text = "Fins/Inch"
        '
        'cboCondFinMat
        '
        Me.cboCondFinMat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCondFinMat.FormattingEnabled = True
        Me.cboCondFinMat.Items.AddRange(New Object() {"A", "C"})
        Me.cboCondFinMat.Location = New System.Drawing.Point(460, 24)
        Me.cboCondFinMat.Name = "cboCondFinMat"
        Me.cboCondFinMat.Size = New System.Drawing.Size(56, 21)
        Me.cboCondFinMat.TabIndex = 13
        '
        'lblCondFinMat
        '
        Me.lblCondFinMat.AutoSize = True
        Me.lblCondFinMat.Location = New System.Drawing.Point(374, 26)
        Me.lblCondFinMat.Name = "lblCondFinMat"
        Me.lblCondFinMat.Size = New System.Drawing.Size(60, 13)
        Me.lblCondFinMat.TabIndex = 12
        Me.lblCondFinMat.Text = "Fin material"
        '
        'cboCondFinThk
        '
        Me.cboCondFinThk.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCondFinThk.FormattingEnabled = True
        Me.cboCondFinThk.Items.AddRange(New Object() {"0.0045", "0.0055", "0.0060", "0.0065", "0.0075", "0.0100"})
        Me.cboCondFinThk.Location = New System.Drawing.Point(296, 138)
        Me.cboCondFinThk.Name = "cboCondFinThk"
        Me.cboCondFinThk.Size = New System.Drawing.Size(61, 21)
        Me.cboCondFinThk.TabIndex = 11
        '
        'lblCondFinThk
        '
        Me.lblCondFinThk.AutoSize = True
        Me.lblCondFinThk.Location = New System.Drawing.Point(213, 141)
        Me.lblCondFinThk.Name = "lblCondFinThk"
        Me.lblCondFinThk.Size = New System.Drawing.Size(73, 13)
        Me.lblCondFinThk.TabIndex = 10
        Me.lblCondFinThk.Text = "Fin Thickness"
        '
        'txtCondFL
        '
        Me.txtCondFL.Location = New System.Drawing.Point(294, 83)
        Me.txtCondFL.Name = "txtCondFL"
        Me.txtCondFL.Size = New System.Drawing.Size(62, 20)
        Me.txtCondFL.TabIndex = 9
        '
        'lblCondFL
        '
        Me.lblCondFL.AutoSize = True
        Me.lblCondFL.Location = New System.Drawing.Point(213, 85)
        Me.lblCondFL.Name = "lblCondFL"
        Me.lblCondFL.Size = New System.Drawing.Size(57, 13)
        Me.lblCondFL.TabIndex = 8
        Me.lblCondFL.Text = "Fin Length"
        '
        'txtCondFH
        '
        Me.txtCondFH.Location = New System.Drawing.Point(294, 27)
        Me.txtCondFH.Name = "txtCondFH"
        Me.txtCondFH.Size = New System.Drawing.Size(63, 20)
        Me.txtCondFH.TabIndex = 7
        '
        'lblCondFH
        '
        Me.lblCondFH.AutoSize = True
        Me.lblCondFH.Location = New System.Drawing.Point(215, 27)
        Me.lblCondFH.Name = "lblCondFH"
        Me.lblCondFH.Size = New System.Drawing.Size(55, 13)
        Me.lblCondFH.TabIndex = 6
        Me.lblCondFH.Text = "Fin Height"
        '
        'cboCondRows
        '
        Me.cboCondRows.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCondRows.FormattingEnabled = True
        Me.cboCondRows.Items.AddRange(New Object() {"1", "2", "3", "4", "5", "6"})
        Me.cboCondRows.Location = New System.Drawing.Point(120, 137)
        Me.cboCondRows.Name = "cboCondRows"
        Me.cboCondRows.Size = New System.Drawing.Size(80, 21)
        Me.cboCondRows.TabIndex = 5
        '
        'lblCondRows
        '
        Me.lblCondRows.AutoSize = True
        Me.lblCondRows.Location = New System.Drawing.Point(20, 140)
        Me.lblCondRows.Name = "lblCondRows"
        Me.lblCondRows.Size = New System.Drawing.Size(63, 13)
        Me.lblCondRows.TabIndex = 4
        Me.lblCondRows.Text = "No of Rows"
        '
        'cboCondCoilP
        '
        Me.cboCondCoilP.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCondCoilP.FormattingEnabled = True
        Me.cboCondCoilP.Items.AddRange(New Object() {"7", "6", "3", "P", "5"})
        Me.cboCondCoilP.Location = New System.Drawing.Point(119, 80)
        Me.cboCondCoilP.Name = "cboCondCoilP"
        Me.cboCondCoilP.Size = New System.Drawing.Size(81, 21)
        Me.cboCondCoilP.TabIndex = 3
        '
        'lblCondCoilP
        '
        Me.lblCondCoilP.AutoSize = True
        Me.lblCondCoilP.Location = New System.Drawing.Point(20, 83)
        Me.lblCondCoilP.Name = "lblCondCoilP"
        Me.lblCondCoilP.Size = New System.Drawing.Size(89, 13)
        Me.lblCondCoilP.TabIndex = 2
        Me.lblCondCoilP.Text = "Cond Coil Pattern"
        '
        'ddCondModel
        '
        Me.ddCondModel.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ddCondModel.FormattingEnabled = True
        Me.ddCondModel.Location = New System.Drawing.Point(119, 25)
        Me.ddCondModel.Name = "ddCondModel"
        Me.ddCondModel.Size = New System.Drawing.Size(81, 21)
        Me.ddCondModel.Sorted = True
        Me.ddCondModel.TabIndex = 1
        '
        'lblCondModel
        '
        Me.lblCondModel.AutoSize = True
        Me.lblCondModel.Location = New System.Drawing.Point(19, 28)
        Me.lblCondModel.Name = "lblCondModel"
        Me.lblCondModel.Size = New System.Drawing.Size(90, 13)
        Me.lblCondModel.TabIndex = 0
        Me.lblCondModel.Text = "Condenser Model"
        '
        'Compressor
        '
        Me.Compressor.Controls.Add(Me.CT_Max)
        Me.Compressor.Controls.Add(Me.lblCondMaxT)
        Me.Compressor.Controls.Add(Me.ET_Max)
        Me.Compressor.Controls.Add(Me.lblMaxEvapT)
        Me.Compressor.Controls.Add(Me.GroupBox6)
        Me.Compressor.Controls.Add(Me.GroupBox3)
        Me.Compressor.Location = New System.Drawing.Point(4, 22)
        Me.Compressor.Name = "Compressor"
        Me.Compressor.Size = New System.Drawing.Size(860, 302)
        Me.Compressor.TabIndex = 2
        Me.Compressor.Text = "Compressor"
        Me.Compressor.UseVisualStyleBackColor = True
        '
        'CT_Max
        '
        Me.CT_Max.Location = New System.Drawing.Point(413, 209)
        Me.CT_Max.Name = "CT_Max"
        Me.CT_Max.Size = New System.Drawing.Size(73, 20)
        Me.CT_Max.TabIndex = 5
        '
        'lblCondMaxT
        '
        Me.lblCondMaxT.AutoSize = True
        Me.lblCondMaxT.Location = New System.Drawing.Point(298, 210)
        Me.lblCondMaxT.Name = "lblCondMaxT"
        Me.lblCondMaxT.Size = New System.Drawing.Size(85, 13)
        Me.lblCondMaxT.TabIndex = 4
        Me.lblCondMaxT.Text = "Max Cond Temp"
        '
        'ET_Max
        '
        Me.ET_Max.Location = New System.Drawing.Point(150, 209)
        Me.ET_Max.Name = "ET_Max"
        Me.ET_Max.Size = New System.Drawing.Size(82, 20)
        Me.ET_Max.TabIndex = 3
        '
        'lblMaxEvapT
        '
        Me.lblMaxEvapT.AutoSize = True
        Me.lblMaxEvapT.Location = New System.Drawing.Point(33, 210)
        Me.lblMaxEvapT.Name = "lblMaxEvapT"
        Me.lblMaxEvapT.Size = New System.Drawing.Size(85, 13)
        Me.lblMaxEvapT.TabIndex = 2
        Me.lblMaxEvapT.Text = "Max Evap Temp"
        '
        'GroupBox6
        '
        Me.GroupBox6.Controls.Add(Me.optSales)
        Me.GroupBox6.Controls.Add(Me.lblSalesReport)
        Me.GroupBox6.Controls.Add(Me.lblEngReport)
        Me.GroupBox6.Controls.Add(Me.OptEng)
        Me.GroupBox6.Location = New System.Drawing.Point(12, 242)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(839, 37)
        Me.GroupBox6.TabIndex = 1
        Me.GroupBox6.TabStop = False
        '
        'optSales
        '
        Me.optSales.AutoSize = True
        Me.optSales.Location = New System.Drawing.Point(407, 16)
        Me.optSales.Name = "optSales"
        Me.optSales.Size = New System.Drawing.Size(14, 13)
        Me.optSales.TabIndex = 3
        Me.optSales.TabStop = True
        Me.optSales.UseVisualStyleBackColor = True
        '
        'lblSalesReport
        '
        Me.lblSalesReport.AutoSize = True
        Me.lblSalesReport.Location = New System.Drawing.Point(287, 16)
        Me.lblSalesReport.Name = "lblSalesReport"
        Me.lblSalesReport.Size = New System.Drawing.Size(68, 13)
        Me.lblSalesReport.TabIndex = 2
        Me.lblSalesReport.Text = "Sales Report"
        '
        'lblEngReport
        '
        Me.lblEngReport.AutoSize = True
        Me.lblEngReport.Location = New System.Drawing.Point(18, 16)
        Me.lblEngReport.Name = "lblEngReport"
        Me.lblEngReport.Size = New System.Drawing.Size(98, 13)
        Me.lblEngReport.TabIndex = 1
        Me.lblEngReport.Text = "Engineering Report"
        '
        'OptEng
        '
        Me.OptEng.AutoSize = True
        Me.OptEng.Checked = True
        Me.OptEng.Location = New System.Drawing.Point(135, 16)
        Me.OptEng.Name = "OptEng"
        Me.OptEng.Size = New System.Drawing.Size(14, 13)
        Me.OptEng.TabIndex = 0
        Me.OptEng.TabStop = True
        Me.OptEng.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.cmdCompressorFile)
        Me.GroupBox3.Controls.Add(Me.txtSuperh)
        Me.GroupBox3.Controls.Add(Me.lblSuperh)
        Me.GroupBox3.Controls.Add(Me.txtSubcool)
        Me.GroupBox3.Controls.Add(Me.lblSubcool)
        Me.GroupBox3.Controls.Add(Me.cboAppCode)
        Me.GroupBox3.Controls.Add(Me.lblCompAppCode)
        Me.GroupBox3.Controls.Add(Me.cboFreCode)
        Me.GroupBox3.Controls.Add(Me.lblFreCode)
        Me.GroupBox3.Controls.Add(Me.cboRefCode)
        Me.GroupBox3.Controls.Add(Me.lblRefCode)
        Me.GroupBox3.Controls.Add(Me.cboCompModel)
        Me.GroupBox3.Controls.Add(Me.lblCompModel)
        Me.GroupBox3.Location = New System.Drawing.Point(11, 13)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(841, 174)
        Me.GroupBox3.TabIndex = 0
        Me.GroupBox3.TabStop = False
        '
        'cmdCompressorFile
        '
        Me.cmdCompressorFile.Location = New System.Drawing.Point(661, 139)
        Me.cmdCompressorFile.Name = "cmdCompressorFile"
        Me.cmdCompressorFile.Size = New System.Drawing.Size(151, 23)
        Me.cmdCompressorFile.TabIndex = 12
        Me.cmdCompressorFile.Text = "Import Compressor File"
        Me.cmdCompressorFile.UseVisualStyleBackColor = True
        '
        'txtSuperh
        '
        Me.txtSuperh.Location = New System.Drawing.Point(405, 142)
        Me.txtSuperh.Name = "txtSuperh"
        Me.txtSuperh.Size = New System.Drawing.Size(70, 20)
        Me.txtSuperh.TabIndex = 11
        '
        'lblSuperh
        '
        Me.lblSuperh.AutoSize = True
        Me.lblSuperh.Location = New System.Drawing.Point(285, 145)
        Me.lblSuperh.Name = "lblSuperh"
        Me.lblSuperh.Size = New System.Drawing.Size(56, 13)
        Me.lblSuperh.TabIndex = 10
        Me.lblSuperh.Text = "Superheat"
        '
        'txtSubcool
        '
        Me.txtSubcool.Location = New System.Drawing.Point(137, 142)
        Me.txtSubcool.Name = "txtSubcool"
        Me.txtSubcool.Size = New System.Drawing.Size(84, 20)
        Me.txtSubcool.TabIndex = 9
        '
        'lblSubcool
        '
        Me.lblSubcool.AutoSize = True
        Me.lblSubcool.Location = New System.Drawing.Point(20, 149)
        Me.lblSubcool.Name = "lblSubcool"
        Me.lblSubcool.Size = New System.Drawing.Size(60, 13)
        Me.lblSubcool.TabIndex = 8
        Me.lblSubcool.Text = "Subcooling"
        '
        'cboAppCode
        '
        Me.cboAppCode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboAppCode.FormattingEnabled = True
        Me.cboAppCode.Location = New System.Drawing.Point(660, 84)
        Me.cboAppCode.Name = "cboAppCode"
        Me.cboAppCode.Size = New System.Drawing.Size(70, 21)
        Me.cboAppCode.TabIndex = 7
        '
        'lblCompAppCode
        '
        Me.lblCompAppCode.AutoSize = True
        Me.lblCompAppCode.Location = New System.Drawing.Point(534, 85)
        Me.lblCompAppCode.Name = "lblCompAppCode"
        Me.lblCompAppCode.Size = New System.Drawing.Size(87, 13)
        Me.lblCompAppCode.TabIndex = 6
        Me.lblCompAppCode.Text = "Application Code"
        '
        'cboFreCode
        '
        Me.cboFreCode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboFreCode.FormattingEnabled = True
        Me.cboFreCode.Location = New System.Drawing.Point(405, 84)
        Me.cboFreCode.Name = "cboFreCode"
        Me.cboFreCode.Size = New System.Drawing.Size(70, 21)
        Me.cboFreCode.TabIndex = 5
        '
        'lblFreCode
        '
        Me.lblFreCode.AutoSize = True
        Me.lblFreCode.Location = New System.Drawing.Point(285, 86)
        Me.lblFreCode.Name = "lblFreCode"
        Me.lblFreCode.Size = New System.Drawing.Size(85, 13)
        Me.lblFreCode.TabIndex = 4
        Me.lblFreCode.Text = "Frequency Code"
        '
        'cboRefCode
        '
        Me.cboRefCode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboRefCode.FormattingEnabled = True
        Me.cboRefCode.Location = New System.Drawing.Point(136, 83)
        Me.cboRefCode.Name = "cboRefCode"
        Me.cboRefCode.Size = New System.Drawing.Size(86, 21)
        Me.cboRefCode.Sorted = True
        Me.cboRefCode.TabIndex = 3
        '
        'lblRefCode
        '
        Me.lblRefCode.AutoSize = True
        Me.lblRefCode.Location = New System.Drawing.Point(20, 87)
        Me.lblRefCode.Name = "lblRefCode"
        Me.lblRefCode.Size = New System.Drawing.Size(87, 13)
        Me.lblRefCode.TabIndex = 2
        Me.lblRefCode.Text = "Refrigerant Code"
        '
        'cboCompModel
        '
        Me.cboCompModel.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCompModel.FormattingEnabled = True
        Me.cboCompModel.Location = New System.Drawing.Point(136, 34)
        Me.cboCompModel.Name = "cboCompModel"
        Me.cboCompModel.Size = New System.Drawing.Size(144, 21)
        Me.cboCompModel.Sorted = True
        Me.cboCompModel.TabIndex = 1
        '
        'lblCompModel
        '
        Me.lblCompModel.AutoSize = True
        Me.lblCompModel.Location = New System.Drawing.Point(15, 34)
        Me.lblCompModel.Name = "lblCompModel"
        Me.lblCompModel.Size = New System.Drawing.Size(94, 13)
        Me.lblCompModel.TabIndex = 0
        Me.lblCompModel.Text = "Compressor Model"
        '
        'Unit
        '
        Me.Unit.Controls.Add(Me.GroupBox5)
        Me.Unit.Controls.Add(Me.cmdUnitDelete)
        Me.Unit.Controls.Add(Me.cmdUnitUpdate)
        Me.Unit.Controls.Add(Me.cmdUnitSave)
        Me.Unit.Controls.Add(Me.cmdUnitAdd)
        Me.Unit.Controls.Add(Me.cboUnitModel)
        Me.Unit.Controls.Add(Me.lblUnitModel)
        Me.Unit.Controls.Add(Me.GroupBox4)
        Me.Unit.Location = New System.Drawing.Point(4, 22)
        Me.Unit.Name = "Unit"
        Me.Unit.Size = New System.Drawing.Size(860, 302)
        Me.Unit.TabIndex = 3
        Me.Unit.Text = "System"
        Me.Unit.UseVisualStyleBackColor = True
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.txtReheatLaDB)
        Me.GroupBox5.Controls.Add(Me.lblReheataDB)
        Me.GroupBox5.Controls.Add(Me.txtReheatCKT)
        Me.GroupBox5.Controls.Add(Me.lblReheatCKT)
        Me.GroupBox5.Controls.Add(Me.chkReheatDroptubes)
        Me.GroupBox5.Controls.Add(Me.lblReDropTubes)
        Me.GroupBox5.Controls.Add(Me.txtReheatFH)
        Me.GroupBox5.Controls.Add(Me.lblReheatFH)
        Me.GroupBox5.Controls.Add(Me.cboReheatFinPI)
        Me.GroupBox5.Controls.Add(Me.lblReheatFinPI)
        Me.GroupBox5.Controls.Add(Me.txtReheatFL)
        Me.GroupBox5.Controls.Add(Me.lblReheatFL)
        Me.GroupBox5.Controls.Add(Me.cboReheatWallThk)
        Me.GroupBox5.Controls.Add(Me.lblReheatWallThk)
        Me.GroupBox5.Controls.Add(Me.cboReheatRows)
        Me.GroupBox5.Controls.Add(Me.lblReheatRows)
        Me.GroupBox5.Controls.Add(Me.cboReheatFinThk)
        Me.GroupBox5.Controls.Add(Me.lblReheatFinThk)
        Me.GroupBox5.Controls.Add(Me.cboReheatCoilP)
        Me.GroupBox5.Controls.Add(Me.lblReheatCoilPat)
        Me.GroupBox5.Location = New System.Drawing.Point(7, 195)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(845, 96)
        Me.GroupBox5.TabIndex = 7
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Reheat Coil"
        '
        'txtReheatLaDB
        '
        Me.txtReheatLaDB.Location = New System.Drawing.Point(759, 63)
        Me.txtReheatLaDB.Name = "txtReheatLaDB"
        Me.txtReheatLaDB.Size = New System.Drawing.Size(72, 20)
        Me.txtReheatLaDB.TabIndex = 19
        '
        'lblReheataDB
        '
        Me.lblReheataDB.AutoSize = True
        Me.lblReheataDB.Location = New System.Drawing.Point(641, 63)
        Me.lblReheataDB.Name = "lblReheataDB"
        Me.lblReheataDB.Size = New System.Drawing.Size(72, 13)
        Me.lblReheataDB.TabIndex = 18
        Me.lblReheataDB.Text = "Reheat Temp"
        '
        'txtReheatCKT
        '
        Me.txtReheatCKT.Location = New System.Drawing.Point(759, 19)
        Me.txtReheatCKT.Name = "txtReheatCKT"
        Me.txtReheatCKT.Size = New System.Drawing.Size(73, 20)
        Me.txtReheatCKT.TabIndex = 17
        '
        'lblReheatCKT
        '
        Me.lblReheatCKT.AutoSize = True
        Me.lblReheatCKT.Location = New System.Drawing.Point(641, 22)
        Me.lblReheatCKT.Name = "lblReheatCKT"
        Me.lblReheatCKT.Size = New System.Drawing.Size(36, 13)
        Me.lblReheatCKT.TabIndex = 16
        Me.lblReheatCKT.Text = "Feeds"
        '
        'chkReheatDroptubes
        '
        Me.chkReheatDroptubes.AutoSize = True
        Me.chkReheatDroptubes.Location = New System.Drawing.Point(570, 63)
        Me.chkReheatDroptubes.Name = "chkReheatDroptubes"
        Me.chkReheatDroptubes.Size = New System.Drawing.Size(15, 14)
        Me.chkReheatDroptubes.TabIndex = 15
        Me.chkReheatDroptubes.UseVisualStyleBackColor = True
        '
        'lblReDropTubes
        '
        Me.lblReDropTubes.AutoSize = True
        Me.lblReDropTubes.Location = New System.Drawing.Point(480, 63)
        Me.lblReDropTubes.Name = "lblReDropTubes"
        Me.lblReDropTubes.Size = New System.Drawing.Size(66, 13)
        Me.lblReDropTubes.TabIndex = 14
        Me.lblReDropTubes.Text = "DropTubes?"
        '
        'txtReheatFH
        '
        Me.txtReheatFH.Location = New System.Drawing.Point(570, 22)
        Me.txtReheatFH.Name = "txtReheatFH"
        Me.txtReheatFH.Size = New System.Drawing.Size(52, 20)
        Me.txtReheatFH.TabIndex = 13
        '
        'lblReheatFH
        '
        Me.lblReheatFH.AutoSize = True
        Me.lblReheatFH.Location = New System.Drawing.Point(480, 21)
        Me.lblReheatFH.Name = "lblReheatFH"
        Me.lblReheatFH.Size = New System.Drawing.Size(55, 13)
        Me.lblReheatFH.TabIndex = 12
        Me.lblReheatFH.Text = "Fin Height"
        '
        'cboReheatFinPI
        '
        Me.cboReheatFinPI.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboReheatFinPI.FormattingEnabled = True
        Me.cboReheatFinPI.Items.AddRange(New Object() {"6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22"})
        Me.cboReheatFinPI.Location = New System.Drawing.Point(404, 60)
        Me.cboReheatFinPI.Name = "cboReheatFinPI"
        Me.cboReheatFinPI.Size = New System.Drawing.Size(57, 21)
        Me.cboReheatFinPI.TabIndex = 11
        '
        'lblReheatFinPI
        '
        Me.lblReheatFinPI.AutoSize = True
        Me.lblReheatFinPI.Location = New System.Drawing.Point(322, 63)
        Me.lblReheatFinPI.Name = "lblReheatFinPI"
        Me.lblReheatFinPI.Size = New System.Drawing.Size(47, 13)
        Me.lblReheatFinPI.TabIndex = 10
        Me.lblReheatFinPI.Text = "Fin/Inch"
        '
        'txtReheatFL
        '
        Me.txtReheatFL.Location = New System.Drawing.Point(405, 18)
        Me.txtReheatFL.Name = "txtReheatFL"
        Me.txtReheatFL.Size = New System.Drawing.Size(57, 20)
        Me.txtReheatFL.TabIndex = 9
        '
        'lblReheatFL
        '
        Me.lblReheatFL.AutoSize = True
        Me.lblReheatFL.Location = New System.Drawing.Point(324, 22)
        Me.lblReheatFL.Name = "lblReheatFL"
        Me.lblReheatFL.Size = New System.Drawing.Size(57, 13)
        Me.lblReheatFL.TabIndex = 8
        Me.lblReheatFL.Text = "Fin Length"
        '
        'cboReheatWallThk
        '
        Me.cboReheatWallThk.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboReheatWallThk.FormattingEnabled = True
        Me.cboReheatWallThk.Items.AddRange(New Object() {"0.012", "0.014", "0.016", "0.017", "0.018", "0.025", "0.035", "0.049"})
        Me.cboReheatWallThk.Location = New System.Drawing.Point(255, 60)
        Me.cboReheatWallThk.Name = "cboReheatWallThk"
        Me.cboReheatWallThk.Size = New System.Drawing.Size(57, 21)
        Me.cboReheatWallThk.TabIndex = 7
        '
        'lblReheatWallThk
        '
        Me.lblReheatWallThk.AutoSize = True
        Me.lblReheatWallThk.Location = New System.Drawing.Point(165, 63)
        Me.lblReheatWallThk.Name = "lblReheatWallThk"
        Me.lblReheatWallThk.Size = New System.Drawing.Size(80, 13)
        Me.lblReheatWallThk.TabIndex = 6
        Me.lblReheatWallThk.Text = "Wall Thickness"
        '
        'cboReheatRows
        '
        Me.cboReheatRows.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboReheatRows.FormattingEnabled = True
        Me.cboReheatRows.Items.AddRange(New Object() {"1", "2", "3", "4"})
        Me.cboReheatRows.Location = New System.Drawing.Point(103, 58)
        Me.cboReheatRows.Name = "cboReheatRows"
        Me.cboReheatRows.Size = New System.Drawing.Size(54, 21)
        Me.cboReheatRows.TabIndex = 5
        '
        'lblReheatRows
        '
        Me.lblReheatRows.AutoSize = True
        Me.lblReheatRows.Location = New System.Drawing.Point(12, 60)
        Me.lblReheatRows.Name = "lblReheatRows"
        Me.lblReheatRows.Size = New System.Drawing.Size(63, 13)
        Me.lblReheatRows.TabIndex = 4
        Me.lblReheatRows.Text = "No of Rows"
        '
        'cboReheatFinThk
        '
        Me.cboReheatFinThk.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboReheatFinThk.FormattingEnabled = True
        Me.cboReheatFinThk.Items.AddRange(New Object() {"0.0045", "0.0055", "0.0060", "0.0065", "0.0075", "0.0100"})
        Me.cboReheatFinThk.Location = New System.Drawing.Point(255, 18)
        Me.cboReheatFinThk.Name = "cboReheatFinThk"
        Me.cboReheatFinThk.Size = New System.Drawing.Size(57, 21)
        Me.cboReheatFinThk.TabIndex = 3
        '
        'lblReheatFinThk
        '
        Me.lblReheatFinThk.AutoSize = True
        Me.lblReheatFinThk.Location = New System.Drawing.Point(165, 26)
        Me.lblReheatFinThk.Name = "lblReheatFinThk"
        Me.lblReheatFinThk.Size = New System.Drawing.Size(73, 13)
        Me.lblReheatFinThk.TabIndex = 2
        Me.lblReheatFinThk.Text = "Fin Thickness"
        '
        'cboReheatCoilP
        '
        Me.cboReheatCoilP.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboReheatCoilP.FormattingEnabled = True
        Me.cboReheatCoilP.Items.AddRange(New Object() {"7", "6", "3", "P", "5"})
        Me.cboReheatCoilP.Location = New System.Drawing.Point(103, 18)
        Me.cboReheatCoilP.Name = "cboReheatCoilP"
        Me.cboReheatCoilP.Size = New System.Drawing.Size(54, 21)
        Me.cboReheatCoilP.TabIndex = 1
        '
        'lblReheatCoilPat
        '
        Me.lblReheatCoilPat.AutoSize = True
        Me.lblReheatCoilPat.Location = New System.Drawing.Point(12, 21)
        Me.lblReheatCoilPat.Name = "lblReheatCoilPat"
        Me.lblReheatCoilPat.Size = New System.Drawing.Size(61, 13)
        Me.lblReheatCoilPat.TabIndex = 0
        Me.lblReheatCoilPat.Text = "Coil Pattern"
        '
        'cmdUnitDelete
        '
        Me.cmdUnitDelete.Location = New System.Drawing.Point(766, 16)
        Me.cmdUnitDelete.Name = "cmdUnitDelete"
        Me.cmdUnitDelete.Size = New System.Drawing.Size(83, 23)
        Me.cmdUnitDelete.TabIndex = 6
        Me.cmdUnitDelete.Text = "Delete Unit"
        Me.cmdUnitDelete.UseVisualStyleBackColor = True
        '
        'cmdUnitUpdate
        '
        Me.cmdUnitUpdate.Location = New System.Drawing.Point(651, 16)
        Me.cmdUnitUpdate.Name = "cmdUnitUpdate"
        Me.cmdUnitUpdate.Size = New System.Drawing.Size(83, 23)
        Me.cmdUnitUpdate.TabIndex = 5
        Me.cmdUnitUpdate.Text = "Unit Update"
        Me.cmdUnitUpdate.UseVisualStyleBackColor = True
        '
        'cmdUnitSave
        '
        Me.cmdUnitSave.Location = New System.Drawing.Point(528, 16)
        Me.cmdUnitSave.Name = "cmdUnitSave"
        Me.cmdUnitSave.Size = New System.Drawing.Size(90, 23)
        Me.cmdUnitSave.TabIndex = 4
        Me.cmdUnitSave.Text = "Save New Unit"
        Me.cmdUnitSave.UseVisualStyleBackColor = True
        '
        'cmdUnitAdd
        '
        Me.cmdUnitAdd.Location = New System.Drawing.Point(412, 16)
        Me.cmdUnitAdd.Name = "cmdUnitAdd"
        Me.cmdUnitAdd.Size = New System.Drawing.Size(84, 23)
        Me.cmdUnitAdd.TabIndex = 3
        Me.cmdUnitAdd.Text = "Add New Unit"
        Me.cmdUnitAdd.UseVisualStyleBackColor = True
        '
        'cboUnitModel
        '
        Me.cboUnitModel.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboUnitModel.FormattingEnabled = True
        Me.cboUnitModel.Location = New System.Drawing.Point(78, 16)
        Me.cboUnitModel.Name = "cboUnitModel"
        Me.cboUnitModel.Size = New System.Drawing.Size(305, 21)
        Me.cboUnitModel.Sorted = True
        Me.cboUnitModel.TabIndex = 2
        '
        'lblUnitModel
        '
        Me.lblUnitModel.AutoSize = True
        Me.lblUnitModel.Location = New System.Drawing.Point(14, 19)
        Me.lblUnitModel.Name = "lblUnitModel"
        Me.lblUnitModel.Size = New System.Drawing.Size(55, 13)
        Me.lblUnitModel.TabIndex = 1
        Me.lblUnitModel.Text = "UnitModel"
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.optHeatPump)
        Me.GroupBox4.Controls.Add(Me.lblheatpump)
        Me.GroupBox4.Controls.Add(Me.lblaircond)
        Me.GroupBox4.Controls.Add(Me.optAirCond)
        Me.GroupBox4.Controls.Add(Me.cmdCalculate)
        Me.GroupBox4.Controls.Add(Me.ddReheatOn_Off)
        Me.GroupBox4.Controls.Add(Me.lblReheatCoil)
        Me.GroupBox4.Controls.Add(Me.cboCompQuantity)
        Me.GroupBox4.Controls.Add(Me.lblCompQuantity)
        Me.GroupBox4.Controls.Add(Me.cboCondQuantity)
        Me.GroupBox4.Controls.Add(Me.lblCondQuantity)
        Me.GroupBox4.Controls.Add(Me.cboEvapQuantity)
        Me.GroupBox4.Controls.Add(Me.lblEvapQuantity)
        Me.GroupBox4.Controls.Add(Me.txtOutaCFM)
        Me.GroupBox4.Controls.Add(Me.lblOutAirCFM)
        Me.GroupBox4.Controls.Add(Me.txtAltitude)
        Me.GroupBox4.Controls.Add(Me.lblAltitude)
        Me.GroupBox4.Controls.Add(Me.cboRef)
        Me.GroupBox4.Controls.Add(Me.lblRefrigerant)
        Me.GroupBox4.Controls.Add(Me.txtCoCFM)
        Me.GroupBox4.Controls.Add(Me.lblCoCFM)
        Me.GroupBox4.Controls.Add(Me.txtCoEaWB)
        Me.GroupBox4.Controls.Add(Me.lblCoEaWB)
        Me.GroupBox4.Controls.Add(Me.txtCoEaDB)
        Me.GroupBox4.Controls.Add(Me.lblCoEaDB)
        Me.GroupBox4.Controls.Add(Me.txtEvCFM)
        Me.GroupBox4.Controls.Add(Me.lblEvCFM)
        Me.GroupBox4.Controls.Add(Me.txtEvEaWB)
        Me.GroupBox4.Controls.Add(Me.lblEvEaWB)
        Me.GroupBox4.Controls.Add(Me.txtEvEaDB)
        Me.GroupBox4.Controls.Add(Me.lblEvapEaDB)
        Me.GroupBox4.Location = New System.Drawing.Point(8, 48)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(844, 138)
        Me.GroupBox4.TabIndex = 0
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "System Properties"
        '
        'optHeatPump
        '
        Me.optHeatPump.AutoSize = True
        Me.optHeatPump.Location = New System.Drawing.Point(726, 110)
        Me.optHeatPump.Name = "optHeatPump"
        Me.optHeatPump.Size = New System.Drawing.Size(14, 13)
        Me.optHeatPump.TabIndex = 30
        Me.optHeatPump.UseVisualStyleBackColor = True
        '
        'lblheatpump
        '
        Me.lblheatpump.AutoSize = True
        Me.lblheatpump.Location = New System.Drawing.Point(640, 109)
        Me.lblheatpump.Name = "lblheatpump"
        Me.lblheatpump.Size = New System.Drawing.Size(60, 13)
        Me.lblheatpump.TabIndex = 29
        Me.lblheatpump.Text = "Heat Pump"
        '
        'lblaircond
        '
        Me.lblaircond.AutoSize = True
        Me.lblaircond.Location = New System.Drawing.Point(640, 67)
        Me.lblaircond.Name = "lblaircond"
        Me.lblaircond.Size = New System.Drawing.Size(80, 13)
        Me.lblaircond.TabIndex = 28
        Me.lblaircond.Text = "Air Conditioning"
        '
        'optAirCond
        '
        Me.optAirCond.AutoSize = True
        Me.optAirCond.Checked = True
        Me.optAirCond.Location = New System.Drawing.Point(726, 67)
        Me.optAirCond.Name = "optAirCond"
        Me.optAirCond.Size = New System.Drawing.Size(14, 13)
        Me.optAirCond.TabIndex = 27
        Me.optAirCond.TabStop = True
        Me.optAirCond.UseVisualStyleBackColor = True
        '
        'cmdCalculate
        '
        Me.cmdCalculate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCalculate.Location = New System.Drawing.Point(752, 103)
        Me.cmdCalculate.Name = "cmdCalculate"
        Me.cmdCalculate.Size = New System.Drawing.Size(86, 26)
        Me.cmdCalculate.TabIndex = 26
        Me.cmdCalculate.Text = "Calculate"
        Me.cmdCalculate.UseVisualStyleBackColor = True
        '
        'ddReheatOn_Off
        '
        Me.ddReheatOn_Off.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ddReheatOn_Off.FormattingEnabled = True
        Me.ddReheatOn_Off.Items.AddRange(New Object() {"None", "Liquid", "VaporParallel", "VaporSeries"})
        Me.ddReheatOn_Off.Location = New System.Drawing.Point(750, 23)
        Me.ddReheatOn_Off.Name = "ddReheatOn_Off"
        Me.ddReheatOn_Off.Size = New System.Drawing.Size(86, 21)
        Me.ddReheatOn_Off.TabIndex = 25
        '
        'lblReheatCoil
        '
        Me.lblReheatCoil.AutoSize = True
        Me.lblReheatCoil.Location = New System.Drawing.Point(640, 30)
        Me.lblReheatCoil.Name = "lblReheatCoil"
        Me.lblReheatCoil.Size = New System.Drawing.Size(68, 13)
        Me.lblReheatCoil.TabIndex = 24
        Me.lblReheatCoil.Text = "Reheat Coil?"
        '
        'cboCompQuantity
        '
        Me.cboCompQuantity.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCompQuantity.FormattingEnabled = True
        Me.cboCompQuantity.Items.AddRange(New Object() {"1", "2"})
        Me.cboCompQuantity.Location = New System.Drawing.Point(569, 108)
        Me.cboCompQuantity.Name = "cboCompQuantity"
        Me.cboCompQuantity.Size = New System.Drawing.Size(52, 21)
        Me.cboCompQuantity.TabIndex = 23
        '
        'lblCompQuantity
        '
        Me.lblCompQuantity.AutoSize = True
        Me.lblCompQuantity.Location = New System.Drawing.Point(476, 111)
        Me.lblCompQuantity.Name = "lblCompQuantity"
        Me.lblCompQuantity.Size = New System.Drawing.Size(76, 13)
        Me.lblCompQuantity.TabIndex = 22
        Me.lblCompQuantity.Text = "Comp Quantity"
        '
        'cboCondQuantity
        '
        Me.cboCondQuantity.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCondQuantity.FormattingEnabled = True
        Me.cboCondQuantity.Items.AddRange(New Object() {"1", "2"})
        Me.cboCondQuantity.Location = New System.Drawing.Point(569, 67)
        Me.cboCondQuantity.Name = "cboCondQuantity"
        Me.cboCondQuantity.Size = New System.Drawing.Size(53, 21)
        Me.cboCondQuantity.TabIndex = 21
        '
        'lblCondQuantity
        '
        Me.lblCondQuantity.AutoSize = True
        Me.lblCondQuantity.Location = New System.Drawing.Point(478, 70)
        Me.lblCondQuantity.Name = "lblCondQuantity"
        Me.lblCondQuantity.Size = New System.Drawing.Size(74, 13)
        Me.lblCondQuantity.TabIndex = 20
        Me.lblCondQuantity.Text = "Cond Quantity"
        '
        'cboEvapQuantity
        '
        Me.cboEvapQuantity.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboEvapQuantity.FormattingEnabled = True
        Me.cboEvapQuantity.Items.AddRange(New Object() {"1", "2"})
        Me.cboEvapQuantity.Location = New System.Drawing.Point(569, 28)
        Me.cboEvapQuantity.Name = "cboEvapQuantity"
        Me.cboEvapQuantity.Size = New System.Drawing.Size(54, 21)
        Me.cboEvapQuantity.TabIndex = 19
        '
        'lblEvapQuantity
        '
        Me.lblEvapQuantity.AutoSize = True
        Me.lblEvapQuantity.Location = New System.Drawing.Point(478, 28)
        Me.lblEvapQuantity.Name = "lblEvapQuantity"
        Me.lblEvapQuantity.Size = New System.Drawing.Size(74, 13)
        Me.lblEvapQuantity.TabIndex = 18
        Me.lblEvapQuantity.Text = "Evap Quantity"
        '
        'txtOutaCFM
        '
        Me.txtOutaCFM.Location = New System.Drawing.Point(404, 109)
        Me.txtOutaCFM.Name = "txtOutaCFM"
        Me.txtOutaCFM.Size = New System.Drawing.Size(57, 20)
        Me.txtOutaCFM.TabIndex = 17
        '
        'lblOutAirCFM
        '
        Me.lblOutAirCFM.AutoSize = True
        Me.lblOutAirCFM.Location = New System.Drawing.Point(321, 110)
        Me.lblOutAirCFM.Name = "lblOutAirCFM"
        Me.lblOutAirCFM.Size = New System.Drawing.Size(83, 13)
        Me.lblOutAirCFM.TabIndex = 16
        Me.lblOutAirCFM.Text = "Outside Air CFM"
        '
        'txtAltitude
        '
        Me.txtAltitude.Location = New System.Drawing.Point(404, 67)
        Me.txtAltitude.Name = "txtAltitude"
        Me.txtAltitude.Size = New System.Drawing.Size(57, 20)
        Me.txtAltitude.TabIndex = 15
        '
        'lblAltitude
        '
        Me.lblAltitude.AutoSize = True
        Me.lblAltitude.Location = New System.Drawing.Point(321, 71)
        Me.lblAltitude.Name = "lblAltitude"
        Me.lblAltitude.Size = New System.Drawing.Size(42, 13)
        Me.lblAltitude.TabIndex = 14
        Me.lblAltitude.Text = "Altitude"
        '
        'cboRef
        '
        Me.cboRef.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboRef.FormattingEnabled = True
        Me.cboRef.Items.AddRange(New Object() {"134A", "22", "404A", "407C Dew Pt", "407C Mid Pt", "410A"})
        Me.cboRef.Location = New System.Drawing.Point(404, 27)
        Me.cboRef.Name = "cboRef"
        Me.cboRef.Size = New System.Drawing.Size(57, 21)
        Me.cboRef.TabIndex = 13
        '
        'lblRefrigerant
        '
        Me.lblRefrigerant.AutoSize = True
        Me.lblRefrigerant.Location = New System.Drawing.Point(321, 27)
        Me.lblRefrigerant.Name = "lblRefrigerant"
        Me.lblRefrigerant.Size = New System.Drawing.Size(59, 13)
        Me.lblRefrigerant.TabIndex = 12
        Me.lblRefrigerant.Text = "Refrigerant"
        '
        'txtCoCFM
        '
        Me.txtCoCFM.Location = New System.Drawing.Point(254, 107)
        Me.txtCoCFM.Name = "txtCoCFM"
        Me.txtCoCFM.Size = New System.Drawing.Size(57, 20)
        Me.txtCoCFM.TabIndex = 11
        '
        'lblCoCFM
        '
        Me.lblCoCFM.AutoSize = True
        Me.lblCoCFM.Location = New System.Drawing.Point(164, 109)
        Me.lblCoCFM.Name = "lblCoCFM"
        Me.lblCoCFM.Size = New System.Drawing.Size(79, 13)
        Me.lblCoCFM.TabIndex = 10
        Me.lblCoCFM.Text = "Cond/Ret CFM"
        '
        'txtCoEaWB
        '
        Me.txtCoEaWB.Location = New System.Drawing.Point(254, 63)
        Me.txtCoEaWB.Name = "txtCoEaWB"
        Me.txtCoEaWB.Size = New System.Drawing.Size(57, 20)
        Me.txtCoEaWB.TabIndex = 9
        '
        'lblCoEaWB
        '
        Me.lblCoEaWB.AutoSize = True
        Me.lblCoEaWB.Location = New System.Drawing.Point(164, 67)
        Me.lblCoEaWB.Name = "lblCoEaWB"
        Me.lblCoEaWB.Size = New System.Drawing.Size(90, 13)
        Me.lblCoEaWB.TabIndex = 8
        Me.lblCoEaWB.Text = "Cond/Ret Air WB"
        '
        'txtCoEaDB
        '
        Me.txtCoEaDB.Location = New System.Drawing.Point(257, 24)
        Me.txtCoEaDB.Name = "txtCoEaDB"
        Me.txtCoEaDB.Size = New System.Drawing.Size(54, 20)
        Me.txtCoEaDB.TabIndex = 7
        '
        'lblCoEaDB
        '
        Me.lblCoEaDB.AutoSize = True
        Me.lblCoEaDB.Location = New System.Drawing.Point(164, 24)
        Me.lblCoEaDB.Name = "lblCoEaDB"
        Me.lblCoEaDB.Size = New System.Drawing.Size(87, 13)
        Me.lblCoEaDB.TabIndex = 6
        Me.lblCoEaDB.Text = "Cond/Ret Air DB"
        '
        'txtEvCFM
        '
        Me.txtEvCFM.Location = New System.Drawing.Point(103, 107)
        Me.txtEvCFM.Name = "txtEvCFM"
        Me.txtEvCFM.Size = New System.Drawing.Size(54, 20)
        Me.txtEvCFM.TabIndex = 5
        '
        'lblEvCFM
        '
        Me.lblEvCFM.AutoSize = True
        Me.lblEvCFM.Location = New System.Drawing.Point(6, 109)
        Me.lblEvCFM.Name = "lblEvCFM"
        Me.lblEvCFM.Size = New System.Drawing.Size(79, 13)
        Me.lblEvCFM.TabIndex = 4
        Me.lblEvCFM.Text = "Evap/Ret CFM"
        '
        'txtEvEaWB
        '
        Me.txtEvEaWB.Location = New System.Drawing.Point(103, 64)
        Me.txtEvEaWB.Name = "txtEvEaWB"
        Me.txtEvEaWB.Size = New System.Drawing.Size(54, 20)
        Me.txtEvEaWB.TabIndex = 3
        '
        'lblEvEaWB
        '
        Me.lblEvEaWB.AutoSize = True
        Me.lblEvEaWB.Location = New System.Drawing.Point(6, 67)
        Me.lblEvEaWB.Name = "lblEvEaWB"
        Me.lblEvEaWB.Size = New System.Drawing.Size(90, 13)
        Me.lblEvEaWB.TabIndex = 2
        Me.lblEvEaWB.Text = "Evap/Ret Air WB"
        '
        'txtEvEaDB
        '
        Me.txtEvEaDB.Location = New System.Drawing.Point(102, 24)
        Me.txtEvEaDB.Name = "txtEvEaDB"
        Me.txtEvEaDB.Size = New System.Drawing.Size(54, 20)
        Me.txtEvEaDB.TabIndex = 1
        '
        'lblEvapEaDB
        '
        Me.lblEvapEaDB.AutoSize = True
        Me.lblEvapEaDB.Location = New System.Drawing.Point(9, 27)
        Me.lblEvapEaDB.Name = "lblEvapEaDB"
        Me.lblEvapEaDB.Size = New System.Drawing.Size(87, 13)
        Me.lblEvapEaDB.TabIndex = 0
        Me.lblEvapEaDB.Text = "Evap/Ret Air DB"
        '
        'FileDialog
        '
        Me.FileDialog.ReadOnlyChecked = True
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.HelpToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(877, 24)
        Me.MenuStrip1.TabIndex = 1
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'HelpToolStripMenuItem
        '
        Me.HelpToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFileHelp})
        Me.HelpToolStripMenuItem.Name = "HelpToolStripMenuItem"
        Me.HelpToolStripMenuItem.Size = New System.Drawing.Size(40, 20)
        Me.HelpToolStripMenuItem.Text = "Help"
        '
        'mnuFileHelp
        '
        Me.mnuFileHelp.Name = "mnuFileHelp"
        Me.mnuFileHelp.Size = New System.Drawing.Size(154, 22)
        Me.mnuFileHelp.Text = "Open Help File"
        '
        'frmUnitInputs
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(877, 356)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "frmUnitInputs"
        Me.Text = "Unit Balance"
        Me.TabControl1.ResumeLayout(False)
        Me.Evaporator.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.Condenser.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.Compressor.ResumeLayout(False)
        Me.Compressor.PerformLayout()
        Me.GroupBox6.ResumeLayout(False)
        Me.GroupBox6.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.Unit.ResumeLayout(False)
        Me.Unit.PerformLayout()
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents Evaporator As System.Windows.Forms.TabPage
    Friend WithEvents Condenser As System.Windows.Forms.TabPage
    Friend WithEvents Compressor As System.Windows.Forms.TabPage
    Friend WithEvents Unit As System.Windows.Forms.TabPage
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cboEvapCoilP As System.Windows.Forms.ComboBox
    Friend WithEvents lblCoilPattern As System.Windows.Forms.Label
    Friend WithEvents ddEvapModel As System.Windows.Forms.ComboBox
    Friend WithEvents lblEvapModel As System.Windows.Forms.Label
    Friend WithEvents cboEvapRows As System.Windows.Forms.ComboBox
    Friend WithEvents lblRows As System.Windows.Forms.Label
    Friend WithEvents lblFinLength As System.Windows.Forms.Label
    Friend WithEvents txtEvapFH As System.Windows.Forms.TextBox
    Friend WithEvents lblEvapFH As System.Windows.Forms.Label
    Friend WithEvents cboEvapFinMat As System.Windows.Forms.ComboBox
    Friend WithEvents lblEvapFinMat As System.Windows.Forms.Label
    Friend WithEvents cboEvapFinThk As System.Windows.Forms.ComboBox
    Friend WithEvents lblFinThk As System.Windows.Forms.Label
    Friend WithEvents txtEvapFL As System.Windows.Forms.TextBox
    Friend WithEvents cboEvapFinPI As System.Windows.Forms.ComboBox
    Friend WithEvents lblFinPI As System.Windows.Forms.Label
    Friend WithEvents cboEvapWallThk As System.Windows.Forms.ComboBox
    Friend WithEvents lblWallThk As System.Windows.Forms.Label
    Friend WithEvents lblEvapCircuits As System.Windows.Forms.Label
    Friend WithEvents txtEvapCKT2 As System.Windows.Forms.TextBox
    Friend WithEvents lblEvapCKT2 As System.Windows.Forms.Label
    Friend WithEvents txtEvapCKT1 As System.Windows.Forms.TextBox
    Friend WithEvents lblEvapCKT1 As System.Windows.Forms.Label
    Friend WithEvents txtEvapCKT As System.Windows.Forms.TextBox
    Friend WithEvents chkEvapDroptubes As System.Windows.Forms.CheckBox
    Friend WithEvents lblDroptubes As System.Windows.Forms.Label
    Friend WithEvents cboEvapSplit As System.Windows.Forms.ComboBox
    Friend WithEvents lblEvapSplits As System.Windows.Forms.Label
    Friend WithEvents cmdEvapAdd As System.Windows.Forms.Button
    Friend WithEvents cmdEvapDelete As System.Windows.Forms.Button
    Friend WithEvents cmdEvapUpdate As System.Windows.Forms.Button
    Friend WithEvents cmdEvapSave As System.Windows.Forms.Button
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents lblCondModel As System.Windows.Forms.Label
    Friend WithEvents lblCondRows As System.Windows.Forms.Label
    Friend WithEvents cboCondCoilP As System.Windows.Forms.ComboBox
    Friend WithEvents lblCondCoilP As System.Windows.Forms.Label
    Friend WithEvents ddCondModel As System.Windows.Forms.ComboBox
    Friend WithEvents txtCondFH As System.Windows.Forms.TextBox
    Friend WithEvents lblCondFH As System.Windows.Forms.Label
    Friend WithEvents cboCondRows As System.Windows.Forms.ComboBox
    Friend WithEvents txtCondFL As System.Windows.Forms.TextBox
    Friend WithEvents lblCondFL As System.Windows.Forms.Label
    Friend WithEvents cboCondFinThk As System.Windows.Forms.ComboBox
    Friend WithEvents lblCondFinThk As System.Windows.Forms.Label
    Friend WithEvents lblCondFinPI As System.Windows.Forms.Label
    Friend WithEvents cboCondFinMat As System.Windows.Forms.ComboBox
    Friend WithEvents lblCondFinMat As System.Windows.Forms.Label
    Friend WithEvents cboCondFinPI As System.Windows.Forms.ComboBox
    Friend WithEvents lblCondCKT1 As System.Windows.Forms.Label
    Friend WithEvents cboCondWallThk As System.Windows.Forms.ComboBox
    Friend WithEvents lblCondWallThk As System.Windows.Forms.Label
    Friend WithEvents txtCondCKT2 As System.Windows.Forms.TextBox
    Friend WithEvents lblCondCKT2 As System.Windows.Forms.Label
    Friend WithEvents txtCondCKT1 As System.Windows.Forms.TextBox
    Friend WithEvents lblCondCircuits As System.Windows.Forms.Label
    Friend WithEvents txtCondCKT As System.Windows.Forms.TextBox
    Friend WithEvents chkCondDroptubes As System.Windows.Forms.CheckBox
    Friend WithEvents lblCondDroptubes As System.Windows.Forms.Label
    Friend WithEvents cboCondSplit As System.Windows.Forms.ComboBox
    Friend WithEvents lblCondSplits As System.Windows.Forms.Label
    Friend WithEvents cmdCondAdd As System.Windows.Forms.Button
    Friend WithEvents cmdCondDelete As System.Windows.Forms.Button
    Friend WithEvents cmdCondUpdate As System.Windows.Forms.Button
    Friend WithEvents cmdCondSave As System.Windows.Forms.Button
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents cboCompModel As System.Windows.Forms.ComboBox
    Friend WithEvents lblCompModel As System.Windows.Forms.Label
    Friend WithEvents lblFreCode As System.Windows.Forms.Label
    Friend WithEvents cboRefCode As System.Windows.Forms.ComboBox
    Friend WithEvents lblRefCode As System.Windows.Forms.Label
    Friend WithEvents cboAppCode As System.Windows.Forms.ComboBox
    Friend WithEvents lblCompAppCode As System.Windows.Forms.Label
    Friend WithEvents cboFreCode As System.Windows.Forms.ComboBox
    Friend WithEvents cboUnitModel As System.Windows.Forms.ComboBox
    Friend WithEvents lblUnitModel As System.Windows.Forms.Label
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdUnitUpdate As System.Windows.Forms.Button
    Friend WithEvents cmdUnitSave As System.Windows.Forms.Button
    Friend WithEvents cmdUnitAdd As System.Windows.Forms.Button
    Friend WithEvents cmdUnitDelete As System.Windows.Forms.Button
    Friend WithEvents lblEvapEaDB As System.Windows.Forms.Label
    Friend WithEvents txtEvCFM As System.Windows.Forms.TextBox
    Friend WithEvents lblEvCFM As System.Windows.Forms.Label
    Friend WithEvents txtEvEaWB As System.Windows.Forms.TextBox
    Friend WithEvents lblEvEaWB As System.Windows.Forms.Label
    Friend WithEvents txtEvEaDB As System.Windows.Forms.TextBox
    Friend WithEvents lblCoCFM As System.Windows.Forms.Label
    Friend WithEvents txtCoEaWB As System.Windows.Forms.TextBox
    Friend WithEvents lblCoEaWB As System.Windows.Forms.Label
    Friend WithEvents txtCoEaDB As System.Windows.Forms.TextBox
    Friend WithEvents lblCoEaDB As System.Windows.Forms.Label
    Friend WithEvents cboRef As System.Windows.Forms.ComboBox
    Friend WithEvents lblRefrigerant As System.Windows.Forms.Label
    Friend WithEvents txtCoCFM As System.Windows.Forms.TextBox
    Friend WithEvents lblOutAirCFM As System.Windows.Forms.Label
    Friend WithEvents txtAltitude As System.Windows.Forms.TextBox
    Friend WithEvents lblAltitude As System.Windows.Forms.Label
    Friend WithEvents txtOutaCFM As System.Windows.Forms.TextBox
    Friend WithEvents cboCondQuantity As System.Windows.Forms.ComboBox
    Friend WithEvents lblCondQuantity As System.Windows.Forms.Label
    Friend WithEvents cboEvapQuantity As System.Windows.Forms.ComboBox
    Friend WithEvents lblEvapQuantity As System.Windows.Forms.Label
    Friend WithEvents ddReheatOn_Off As System.Windows.Forms.ComboBox
    Friend WithEvents lblReheatCoil As System.Windows.Forms.Label
    Friend WithEvents cboCompQuantity As System.Windows.Forms.ComboBox
    Friend WithEvents lblCompQuantity As System.Windows.Forms.Label
    Friend WithEvents cmdCalculate As System.Windows.Forms.Button
    Friend WithEvents lblaircond As System.Windows.Forms.Label
    Friend WithEvents optAirCond As System.Windows.Forms.RadioButton
    Friend WithEvents optHeatPump As System.Windows.Forms.RadioButton
    Friend WithEvents lblheatpump As System.Windows.Forms.Label
    Friend WithEvents lblSubcool As System.Windows.Forms.Label
    Friend WithEvents txtSuperh As System.Windows.Forms.TextBox
    Friend WithEvents lblSuperh As System.Windows.Forms.Label
    Friend WithEvents txtSubcool As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents cboReheatCoilP As System.Windows.Forms.ComboBox
    Friend WithEvents lblReheatCoilPat As System.Windows.Forms.Label
    Friend WithEvents cboReheatRows As System.Windows.Forms.ComboBox
    Friend WithEvents lblReheatRows As System.Windows.Forms.Label
    Friend WithEvents cboReheatFinThk As System.Windows.Forms.ComboBox
    Friend WithEvents lblReheatFinThk As System.Windows.Forms.Label
    Friend WithEvents cboReheatFinPI As System.Windows.Forms.ComboBox
    Friend WithEvents lblReheatFinPI As System.Windows.Forms.Label
    Friend WithEvents txtReheatFL As System.Windows.Forms.TextBox
    Friend WithEvents lblReheatFL As System.Windows.Forms.Label
    Friend WithEvents cboReheatWallThk As System.Windows.Forms.ComboBox
    Friend WithEvents lblReheatWallThk As System.Windows.Forms.Label
    Friend WithEvents txtReheatFH As System.Windows.Forms.TextBox
    Friend WithEvents lblReheatFH As System.Windows.Forms.Label
    Friend WithEvents txtReheatLaDB As System.Windows.Forms.TextBox
    Friend WithEvents lblReheataDB As System.Windows.Forms.Label
    Friend WithEvents txtReheatCKT As System.Windows.Forms.TextBox
    Friend WithEvents lblReheatCKT As System.Windows.Forms.Label
    Friend WithEvents chkReheatDroptubes As System.Windows.Forms.CheckBox
    Friend WithEvents lblReDropTubes As System.Windows.Forms.Label
    Friend WithEvents GroupBox6 As System.Windows.Forms.GroupBox
    Friend WithEvents lblEngReport As System.Windows.Forms.Label
    Friend WithEvents OptEng As System.Windows.Forms.RadioButton
    Friend WithEvents optSales As System.Windows.Forms.RadioButton
    Friend WithEvents lblSalesReport As System.Windows.Forms.Label
    Friend WithEvents CT_Max As System.Windows.Forms.TextBox
    Friend WithEvents lblCondMaxT As System.Windows.Forms.Label
    Friend WithEvents ET_Max As System.Windows.Forms.TextBox
    Friend WithEvents lblMaxEvapT As System.Windows.Forms.Label
    Friend WithEvents cmdCompressorFile As System.Windows.Forms.Button
    Friend WithEvents FileDialog As System.Windows.Forms.OpenFileDialog
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents HelpToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuFileHelp As System.Windows.Forms.ToolStripMenuItem

End Class
