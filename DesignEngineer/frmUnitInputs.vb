Imports System.Data
Imports System.Data.OleDb
Imports System.IO
Imports System.Windows.Forms

Public Class frmUnitInputs
    Dim das1 As New DataSet()
    Dim dapEvap As New OleDbDataAdapter()
    Dim dapCond As New OleDbDataAdapter()
    Dim dapComp As New OleDbDataAdapter()
    Dim dapTempComp As New OleDbDataAdapter()
    Dim dapUnit As New OleDbDataAdapter()
    Dim MyConnection As OleDbConnection = ConnectToUnitData()
    Dim bNewRow As Boolean

    Private Sub frmUnitInputs_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        dapEvap.SelectCommand = New OleDbCommand("SELECT * FROM Evaporator Order By Model", MyConnection)
        dapCond.SelectCommand = New OleDbCommand("SELECT * FROM Condenser Order By Model", MyConnection)
        dapTempComp.SelectCommand = New OleDbCommand("SELECT Model, RefCode, FreCode, AppCode FROM Compressor Order By Model", MyConnection)
        dapUnit.SelectCommand = New OleDbCommand("SELECT * FROM Unit", MyConnection)

        dapEvap.MissingSchemaAction = MissingSchemaAction.AddWithKey
        dapCond.MissingSchemaAction = MissingSchemaAction.AddWithKey
        dapUnit.MissingSchemaAction = MissingSchemaAction.AddWithKey

        dapEvap.InsertCommand = New OleDbCommand("INSERT INTO Evaporator " & _
        "(Model, FH, FL, FinPI, Passes, FinThk, CoilP, WallThk, CKT, Split, CKT1, CKT2, FinMat, Droptubes) " & _
        "VALUES (@Model, @FH, @FL, @FinPI, @Passes, @FinThk, @CoilP, @WallThk, @CKT, @Split, @CKT1, @CKT2, " & _
        "@FinMat, @Droptubes)", MyConnection)

        dapEvap.UpdateCommand = New OleDbCommand("Update Evaporator SET Model = ?, " & _
                                "FH = ?, FL  = ?, FinPI = ?, Passes = ?, FinThk = ?, " & _
                                "CoilP = ?, WallThk = ?, CKT = ?, Split = ?, CKT1 = ?, " & _
                                "CKT2 = ?, FinMat = ?, Droptubes = ? WHERE Model =?", MyConnection)

        dapEvap.DeleteCommand = New OleDbCommand("DELETE FROM Evaporator WHERE Model = ?", MyConnection)

        With dapEvap.InsertCommand.Parameters
            .Add("@Model", OleDbType.VarChar, 40, "Model")
            .Add("@FH", OleDbType.Char, 12, "FH")
            .Add("@FL", OleDbType.Char, 12, "FL")
            .Add("@FinPI", OleDbType.Char, 8, "FinPI")
            .Add("@Passes", OleDbType.Char, 8, "Passes")
            .Add("@FinThk", OleDbType.Char, 8, "FinThk")
            .Add("@CoilP", OleDbType.Char, 8, "CoilP")
            .Add("@WallThk", OleDbType.Char, 8, "WallThk")
            .Add("@CKT", OleDbType.Char, 8, "CKT")
            .Add("@Split", OleDbType.Char, 8, "Split")
            .Add("@CKT1", OleDbType.Char, 8, "CKT1")
            .Add("@CKT2", OleDbType.Char, 8, "CKT2")
            .Add("@FinMat", OleDbType.Char, 8, "FinMat")
            .Add("@Droptubes", OleDbType.Boolean, 2, "Droptubes")
        End With

        With dapEvap.UpdateCommand.Parameters
            .Add("@Model", OleDbType.VarChar, 40, "Model")
            .Add("@FH", OleDbType.Char, 12, "FH")
            .Add("@FL", OleDbType.Char, 12, "FL")
            .Add("@FinPI", OleDbType.Char, 8, "FinPI")
            .Add("@Passes", OleDbType.Char, 8, "Passes")
            .Add("@FinThk", OleDbType.Char, 8, "FinThk")
            .Add("@CoilP", OleDbType.Char, 8, "CoilP")
            .Add("@WallThk", OleDbType.Char, 8, "WallThk")
            .Add("@CKT", OleDbType.Char, 8, "CKT")
            .Add("@Split", OleDbType.Char, 8, "Split")
            .Add("@CKT1", OleDbType.Char, 8, "CKT1")
            .Add("@CKT2", OleDbType.Char, 8, "CKT2")
            .Add("@FinMat", OleDbType.Char, 8, "FinMat")
            .Add("@Droptubes", OleDbType.Boolean, 2, "Droptubes")
            .Add("@oldModel", OleDbType.VarChar, 40, "Model").SourceVersion = DataRowVersion.Original
        End With

        dapEvap.DeleteCommand.Parameters.Add("@oldModel", OleDbType.VarChar, 40, "Model").SourceVersion = DataRowVersion.Original

        dapEvap.Fill(das1, "Evaporator")

        Dim drwEvap As DataRow

        For Each drwEvap In das1.Tables("Evaporator").Rows
            ddEvapModel.Items.Add(drwEvap("Model"))
        Next

        ddEvapModel.SelectedIndex = 0

        dapCond.InsertCommand = New OleDbCommand("INSERT INTO Condenser " & _
        "(Model, FH, FL, FinPI, Passes, FinThk, CoilP, WallThk, CKT, Split, CKT1, CKT2, FinMat, Droptubes) " & _
        "VALUES (@Model, @FH, @FL, @FinPI, @Passes, @FinThk, @CoilP, @WallThk, @CKT, @Split, @CKT1, @CKT2, " & _
        "@FinMat, @Droptubes)", MyConnection)

        dapCond.UpdateCommand = New OleDbCommand("Update Condenser SET Model = ?, " & _
                                "FH = ?, FL  = ?, FinPI = ?, Passes = ?, FinThk = ?, " & _
                                "CoilP = ?, WallThk = ?, CKT = ?, Split = ?, CKT1 = ?, " & _
                                "CKT2 = ?, FinMat = ?, Droptubes = ? WHERE Model =?", MyConnection)

        dapCond.DeleteCommand = New OleDbCommand("DELETE FROM Condenser WHERE Model = ?", MyConnection)

        With dapCond.InsertCommand.Parameters
            .Add("@Model", OleDbType.VarChar, 40, "Model")
            .Add("@FH", OleDbType.Char, 12, "FH")
            .Add("@FL", OleDbType.Char, 12, "FL")
            .Add("@FinPI", OleDbType.Char, 8, "FinPI")
            .Add("@Passes", OleDbType.Char, 8, "Passes")
            .Add("@FinThk", OleDbType.Char, 8, "FinThk")
            .Add("@CoilP", OleDbType.Char, 8, "CoilP")
            .Add("@WallThk", OleDbType.Char, 8, "WallThk")
            .Add("@CKT", OleDbType.Char, 8, "CKT")
            .Add("@Split", OleDbType.Char, 8, "Split")
            .Add("@CKT1", OleDbType.Char, 8, "CKT1")
            .Add("@CKT2", OleDbType.Char, 8, "CKT2")
            .Add("@FinMat", OleDbType.Char, 8, "FinMat")
            .Add("@Droptubes", OleDbType.Boolean, 2, "Droptubes")
        End With

        With dapCond.UpdateCommand.Parameters
            .Add("@Model", OleDbType.VarChar, 40, "Model")
            .Add("@FH", OleDbType.Char, 12, "FH")
            .Add("@FL", OleDbType.Char, 12, "FL")
            .Add("@FinPI", OleDbType.Char, 8, "FinPI")
            .Add("@Passes", OleDbType.Char, 8, "Passes")
            .Add("@FinThk", OleDbType.Char, 8, "FinThk")
            .Add("@CoilP", OleDbType.Char, 8, "CoilP")
            .Add("@WallThk", OleDbType.Char, 8, "WallThk")
            .Add("@CKT", OleDbType.Char, 8, "CKT")
            .Add("@Split", OleDbType.Char, 8, "Split")
            .Add("@CKT1", OleDbType.Char, 8, "CKT1")
            .Add("@CKT2", OleDbType.Char, 8, "CKT2")
            .Add("@FinMat", OleDbType.Char, 8, "FinMat")
            .Add("@Droptubes", OleDbType.Boolean, 2, "Droptubes")
            .Add("@oldModel", OleDbType.VarChar, 40, "Model").SourceVersion = DataRowVersion.Original
        End With

        dapCond.DeleteCommand.Parameters.Add("@oldModel", OleDbType.VarChar, 40, "Model").SourceVersion = DataRowVersion.Original

        dapCond.Fill(das1, "Condenser")

        Dim drwCond As DataRow

        For Each drwCond In das1.Tables("Condenser").Rows
            ddCondModel.Items.Add(drwCond("Model"))
        Next

        ddCondModel.SelectedIndex = 0

        dapTempComp.Fill(das1, "TempCompressor")

        Dim drwComp As DataRow

        For Each drwComp In das1.Tables("TempCompressor").Rows
            cboCompModel.Items.Add(drwComp("Model"))
        Next

        cboCompModel.SelectedIndex = 0

        dapUnit.InsertCommand = New OleDbCommand("INSERT INTO Unit " & _
        "(UnitModel, CompressorModel, CompressorRef, CompressorFre, CompressorApp, Com2Model, Com2Ref, Com2Fre, Com2App, " & _
        " EvaporatorModel, CondenserModel, EvapEADB, EvapEAWB, EvapCFM, CondEADB, CondEAWB, CondCFM, OACFM, " & _
        " CompQty, EvapQty, CondQty, Refrigerant, Altitude, Reheat, ReheatCoilP, ReheatRows, ReheatFinThk, " & _
        " ReheatWallThk, ReheatFPI, ReheatCKT, ReheatTemp, ReheatFL, ReheatFH, Droptubes, Subcool, Superh, ETMax, CTMax) " & _
        "VALUES (@UnitModel, @CompressorModel, @CompresorRef, @CompressorFre, @CompressorApp, @Com2Model, @Com2Ref, " & _
        " @Com2Fre, @Com2App, @EvaporatorModel, @CondenserModel, @EvapEADB, @EVEAWB, @EvapCFM, @CondEADB, @CondEAWB, " & _
        " @CondCFM, @OACFM, @CompQty, @EvapQty, @CondQty, @Refrigerant, @Altitude, @Reheat, @ReheatCoilP, @ReheatRows, " & _
        " @ReheatFinThk, @ReheatWallThk, @ReheatFPI, @ReheatCKT, @ReheatTemp, @ReheatFL, @ReheatFH, @Droptubes, @Subcool, " & _
        " @Superh, @ETMax, @CTMax)", MyConnection)

        dapUnit.UpdateCommand = New OleDbCommand("Update Unit SET UnitModel = ?, " & _
        " CompressorModel= ?, CompressorRef  = ?, CompressorFre = ?, CompressorApp = ?, Com2Model = ?, " & _
        " Com2Ref = ?, Com2Fre = ?, Com2App = ?, EvaporatorModel = ?, CondenserModel = ?, EvapEADB = ?, " & _
        " EvapEAWB = ?, EvapCFM = ?, CondEADB = ?, CondEAWB = ?, CondCFM = ?, OACFM = ?, CompQty = ?, " & _
        " EvapQty = ?, CondQty = ?, Refrigerant = ?, Altitude = ?, Reheat = ?, ReheatCoilP = ?, " & _
        " ReheatRows = ?, ReheatFinThk = ?, ReheatWallThk = ?, ReheatFPI = ?, ReheatCKT = ?, ReheatTemp = ?, " & _
        " ReheatFL = ?, ReheatFH = ?, Droptubes = ?, Subcool = ?, Superh = ?, ETMax = ?, CTMax =? WHERE UnitModel = ?", MyConnection)

        dapUnit.DeleteCommand = New OleDbCommand("DELETE FROM Unit WHERE UnitModel = ?", MyConnection)

        With dapUnit.InsertCommand.Parameters
            .Add("@UnitModel", OleDbType.VarChar, 50, "UnitModel")
            .Add("@CompressorModel", OleDbType.VarChar, 40, "CompressorModel")
            .Add("@CompressorRef", OleDbType.Char, 12, "CompressorRef")
            .Add("@CompressorFre", OleDbType.Char, 8, "CompressorFre")
            .Add("@CompressorApp", OleDbType.Char, 8, "CompressorApp")
            .Add("@Com2Model", OleDbType.VarChar, 40, "Com2Model")
            .Add("@Com2Ref", OleDbType.Char, 12, "Com2Ref")
            .Add("@Com2Fre", OleDbType.Char, 8, "Com2Fre")
            .Add("@Com2App", OleDbType.Char, 8, "Com2App")
            .Add("@EvaporatorModel", OleDbType.VarChar, 40, "EvaporatorModel")
            .Add("@CondenserModel", OleDbType.VarChar, 40, "CondenserModel")
            .Add("@EvapEADB", OleDbType.Char, 8, "EvapEADB")
            .Add("@EvapEAWB", OleDbType.Char, 8, "EvapEAWB")
            .Add("@EvapCFM", OleDbType.Char, 8, "EvapCFM")
            .Add("@CondEADB", OleDbType.Char, 8, "CondEADB")
            .Add("@CondEAWB", OleDbType.Char, 8, "CondEAWB")
            .Add("@CondCFM", OleDbType.Char, 8, "CondCFM")
            .Add("@OACFM", OleDbType.Char, 8, "OACFM")
            .Add("@CompQty", OleDbType.Char, 8, "CompQty")
            .Add("@EvapQty", OleDbType.Char, 8, "EvapQty")
            .Add("@CondQty", OleDbType.Char, 8, "CondQty")
            .Add("@Refrigerant", OleDbType.Char, 12, "Refrigerant")
            .Add("@Altitude", OleDbType.Char, 8, "Altitude")
            .Add("@Reheat", OleDbType.VarChar, 40, "Reheat")
            .Add("@ReheatCoilP", OleDbType.Char, 8, "ReheatCoilP")
            .Add("@ReheatRows", OleDbType.Char, 8, "ReheatRows")
            .Add("@ReheatFinThk", OleDbType.Char, 8, "ReheatFinThk")
            .Add("@ReheatWallThk", OleDbType.Char, 8, "ReheatWallThk")
            .Add("@ReheatFPI", OleDbType.Char, 8, "ReheatFPI")
            .Add("@ReheatCKT", OleDbType.Char, 8, "ReheatCKT")
            .Add("@ReheatTemp", OleDbType.Char, 8, "ReheatTemp")
            .Add("@ReheatFL", OleDbType.Char, 8, "ReheatFL")
            .Add("@ReheatFH", OleDbType.Char, 8, "ReheatFH")
            .Add("@Droptubes", OleDbType.Boolean, 2, "Droptubes")
            .Add("@Subcool", OleDbType.Char, 8, "Subcool")
            .Add("@Superh", OleDbType.Char, 8, "Superh")
            .Add("@ETMax", OleDbType.Char, 8, "ETMax")
            .Add("@CTMax", OleDbType.Char, 8, "CTMax")
        End With

        With dapUnit.UpdateCommand.Parameters
            .Add("@UnitModel", OleDbType.VarChar, 50, "UnitModel")
            .Add("@CompressorModel", OleDbType.VarChar, 40, "CompressorModel")
            .Add("@CompressorRef", OleDbType.Char, 12, "CompressorRef")
            .Add("@CompressorFre", OleDbType.Char, 8, "CompressorFre")
            .Add("@CompressorApp", OleDbType.Char, 8, "CompressorApp")
            .Add("@Com2Model", OleDbType.VarChar, 40, "Com2Model")
            .Add("@Com2Ref", OleDbType.Char, 12, "Com2Ref")
            .Add("@Com2Fre", OleDbType.Char, 8, "Com2Fre")
            .Add("@Com2App", OleDbType.Char, 8, "Com2App")
            .Add("@EvaporatorModel", OleDbType.VarChar, 40, "EvaporatorModel")
            .Add("@CondenserModel", OleDbType.VarChar, 40, "CondenserModel")
            .Add("@EvapEADB", OleDbType.Char, 8, "EvapEADB")
            .Add("@EvapEAWB", OleDbType.Char, 8, "EvapEAWB")
            .Add("@EvapCFM", OleDbType.Char, 8, "EvapCFM")
            .Add("@CondEADB", OleDbType.Char, 8, "CondEADB")
            .Add("@CondEAWB", OleDbType.Char, 8, "CondEAWB")
            .Add("@CondCFM", OleDbType.Char, 8, "CondCFM")
            .Add("@OACFM", OleDbType.Char, 8, "OACFM")
            .Add("@CompQty", OleDbType.Char, 8, "CompQty")
            .Add("@EvapQty", OleDbType.Char, 8, "EvapQty")
            .Add("@CondQty", OleDbType.Char, 8, "CondQty")
            .Add("@Refrigerant", OleDbType.Char, 12, "Refrigerant")
            .Add("@Altitude", OleDbType.Char, 8, "Altitude")
            .Add("@Reheat", OleDbType.VarChar, 40, "Reheat")
            .Add("@ReheatCoilP", OleDbType.Char, 8, "ReheatCoilP")
            .Add("@ReheatRows", OleDbType.Char, 8, "ReheatRows")
            .Add("@ReheatFinThk", OleDbType.Char, 8, "ReheatFinThk")
            .Add("@ReheatWallThk", OleDbType.Char, 8, "ReheatWallThk")
            .Add("@ReheatFPI", OleDbType.Char, 8, "ReheatFPI")
            .Add("@ReheatCKT", OleDbType.Char, 8, "ReheatCKT")
            .Add("@ReheatTemp", OleDbType.Char, 8, "ReheatTemp")
            .Add("@ReheatFL", OleDbType.Char, 8, "ReheatFL")
            .Add("@ReheatFH", OleDbType.Char, 8, "ReheatFH")
            .Add("@Droptubes", OleDbType.Boolean, 2, "Droptubes")
            .Add("@Subcool", OleDbType.Char, 8, "Subcool")
            .Add("@Superh", OleDbType.Char, 8, "Superh")
            .Add("@ETMax", OleDbType.Char, 8, "ETMax")
            .Add("@CTMax", OleDbType.Char, 8, "CTMax")
            .Add("@oldModel", OleDbType.VarChar, 50, "UnitModel").SourceVersion = DataRowVersion.Original
        End With

        dapUnit.DeleteCommand.Parameters.Add("@oldModel", OleDbType.VarChar, 50, "UnitModel").SourceVersion = DataRowVersion.Original

        dapUnit.Fill(das1, "Unit")

        Dim drwUnit As DataRow

        For Each drwUnit In das1.Tables("Unit").Rows
            cboUnitModel.Items.Add(drwUnit("UnitModel"))
        Next

        cboUnitModel.SelectedIndex = 0

    End Sub

    Private Sub ddEvapModel_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) _
                                                   Handles ddEvapModel.SelectedIndexChanged
        If bNewRow = False Then
            PopulateEvaporator(ddEvapModel.SelectedItem)
        Else
            cboEvapCoilP.Text = 3
            cboEvapRows.Text = 1
            txtEvapFH.Text = ""
            txtEvapFL.Text = ""
            cboEvapFinThk.Text = 0.006
            cboEvapFinMat.Text = "A"
            cboEvapFinPI.Text = 10
            cboEvapWallThk.Text = 0.012
            txtEvapCKT1.Text = ""
            txtEvapCKT2.Text = ""
            txtEvapCKT.Text = ""
            cboEvapSplit.Text = "NS"
            chkEvapDroptubes.Checked = True
        End If
    End Sub

    Sub PopulateEvaporator(ByVal EvapModel As String)

        Dim drwEvap As DataRow = das1.Tables("Evaporator").Rows.Find(EvapModel)
        cboEvapCoilP.Text = drwEvap("CoilP")
        cboEvapRows.Text = drwEvap("Passes")
        txtEvapFH.Text = drwEvap("FH")
        txtEvapFL.Text = drwEvap("FL")
        cboEvapFinThk.Text = drwEvap("FinThk")
        cboEvapFinMat.Text = drwEvap("FinMat")
        cboEvapFinPI.Text = drwEvap("FinPI")
        cboEvapWallThk.Text = drwEvap("WallThk")
        txtEvapCKT1.Text = drwEvap("CKT1")
        txtEvapCKT2.Text = drwEvap("CKT2")
        txtEvapCKT.Text = drwEvap("CKT")
        cboEvapSplit.Text = drwEvap("Split")
        chkEvapDroptubes.Checked = drwEvap("Droptubes")

    End Sub
    Function ConnectToUnitData() As OleDbConnection
        Dim DBPATH As String
        Dim CNString As String
        Dim cnn2 As OleDbConnection
        DBPATH = AppPath() & "\UnitData.mdb"
        CNString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source =" & DBPATH & ""
        cnn2 = New OleDbConnection()
        cnn2.ConnectionString = CNString
        Return cnn2
    End Function
    Public Function AppPath() As String
        Return Application.StartupPath
    End Function

    Private Sub cmdEvapAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEvapAdd.Click

        If cmdEvapAdd.Text = "Add New Evap" Then
            Dim NewEvapModel As String
            Dim Mysentence As String

            cmdEvapDelete.Enabled = False
            cmdEvapSave.Enabled = True
            cmdEvapAdd.Text = "Cancel Evap"
            Mysentence = InputBox("Give New Evap model a name:", "New Evaporator Model", "")

            If Mysentence <> "" Then

                Dim Seperator() As Char = {" "c, "."c, ","c}
                Dim Words() As String = EnhancedSplit(Mysentence, Seperator)

                Dim Word As String
                NewEvapModel = ""
                For Each Word In Words
                    NewEvapModel = NewEvapModel & Word
                Next

                Dim MyRow As DataRow = das1.Tables("Evaporator").Rows.Find(NewEvapModel)

                If MyRow Is Nothing Then
                    ddEvapModel.Items.Add(NewEvapModel)
                    bNewRow = True
                    ddEvapModel.SelectedIndex = ddEvapModel.Items.IndexOf(NewEvapModel)
                    MsgBox("Complete data fields for model " & NewEvapModel & " and click Save", vbOKOnly, "Complete data and save")
                Else
                    cmdEvapDelete.Enabled = True
                    cmdEvapSave.Enabled = False
                    cmdEvapAdd.Text = "Add New Evap"
                    ddEvapModel.SelectedIndex = 0
                    MsgBox("Evaporator Model: " & NewEvapModel & " already exists", vbOKOnly, "Model Already Exists")
                End If

            Else
                cmdEvapDelete.Enabled = True
                cmdEvapSave.Enabled = False
                cmdEvapAdd.Text = "Add New Evap"
                ddEvapModel.SelectedIndex = 0
            End If

        Else
            cmdEvapDelete.Enabled = True
            cmdEvapSave.Enabled = False
            cmdEvapAdd.Text = "Add New Evap"
            ddEvapModel.SelectedIndex = 0
        End If

    End Sub

    Private Sub cmdEvapSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEvapSave.Click

        Dim drwEvap As DataRow = das1.Tables("Evaporator").NewRow
        drwEvap("Model") = ddEvapModel.SelectedItem
        drwEvap("CoilP") = cboEvapCoilP.Text
        drwEvap("Passes") = cboEvapRows.Text
        drwEvap("FH") = txtEvapFH.Text
        drwEvap("FL") = txtEvapFL.Text
        drwEvap("FinThk") = cboEvapFinThk.Text
        drwEvap("FinMat") = cboEvapFinMat.Text
        drwEvap("FinPI") = cboEvapFinPI.Text
        drwEvap("WallThk") = cboEvapWallThk.Text
        drwEvap("CKT1") = txtEvapCKT1.Text
        drwEvap("CKT2") = txtEvapCKT2.Text
        drwEvap("CKT") = txtEvapCKT.Text
        drwEvap("Split") = cboEvapSplit.Text
        drwEvap("Droptubes") = chkEvapDroptubes.CheckState

        das1.Tables("Evaporator").Rows.Add(drwEvap)
        dapEvap.Update(das1, "Evaporator")

        cmdEvapDelete.Enabled = True
        cmdEvapSave.Enabled = False
        cmdEvapAdd.Text = "Add New Evap"
        bNewRow = False

    End Sub

    Private Sub cmdEvapDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEvapDelete.Click

        Dim MyRow As DataRow = das1.Tables("Evaporator").Rows.Find(ddEvapModel.SelectedItem)

        If Not (MyRow Is Nothing) Then
            das1.Tables("Evaporator").Rows.Find(ddEvapModel.SelectedItem).Delete()
        End If

        dapEvap.Update(das1, "Evaporator")
        ddEvapModel.Items.Remove(ddEvapModel.SelectedItem)
        ddEvapModel.SelectedIndex = 0

        MsgBox("Record Deleted", vbInformation, "Delete Confirm")
    End Sub

    Private Sub cmdEvapUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEvapUpdate.Click

        Dim drwEvap As DataRow = das1.Tables("Evaporator").Rows.Find(ddEvapModel.SelectedItem)

        drwEvap("CoilP") = cboEvapCoilP.Text
        drwEvap("Passes") = cboEvapRows.Text
        drwEvap("FH") = txtEvapFH.Text
        drwEvap("FL") = txtEvapFL.Text
        drwEvap("FinThk") = cboEvapFinThk.Text
        drwEvap("FinMat") = cboEvapFinMat.Text
        drwEvap("FinPI") = cboEvapFinPI.Text
        drwEvap("WallThk") = cboEvapWallThk.Text
        drwEvap("CKT1") = txtEvapCKT1.Text
        drwEvap("CKT2") = txtEvapCKT2.Text
        drwEvap("CKT") = txtEvapCKT.Text
        drwEvap("Split") = cboEvapSplit.Text
        drwEvap("Droptubes") = chkEvapDroptubes.CheckState

        dapEvap.Update(das1, "Evaporator")

    End Sub
    Private Sub ddCondModel_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) _
                                                   Handles ddCondModel.SelectedIndexChanged
        If bNewRow = False Then
            PopulateCondenser(ddCondModel.SelectedItem)
        Else
            cboCondCoilP.Text = 3
            cboCondRows.Text = 1
            txtCondFH.Text = ""
            txtCondFL.Text = ""
            cboCondFinThk.Text = 0.006
            cboCondFinMat.Text = "A"
            cboCondFinPI.Text = 10
            cboCondWallThk.Text = 0.012
            txtCondCKT1.Text = ""
            txtCondCKT2.Text = ""
            txtCondCKT.Text = ""
            cboCondSplit.Text = "NS"
            chkCondDroptubes.Checked = True
        End If
    End Sub
    Sub PopulateCondenser(ByVal CondModel As String)

        Dim drwCond As DataRow = das1.Tables("Condenser").Rows.Find(CondModel)
        cboCondCoilP.Text = drwCond("CoilP")
        cboCondRows.Text = drwCond("Passes")
        txtCondFH.Text = drwCond("FH")
        txtCondFL.Text = drwCond("FL")
        cboCondFinThk.Text = drwCond("FinThk")
        cboCondFinMat.Text = drwCond("FinMat")
        cboCondFinPI.Text = drwCond("FinPI")
        cboCondWallThk.Text = drwCond("WallThk")
        txtCondCKT1.Text = drwCond("CKT1")
        txtCondCKT2.Text = drwCond("CKT2")
        txtCondCKT.Text = drwCond("CKT")
        cboCondSplit.Text = drwCond("Split")
        chkCondDroptubes.Checked = drwCond("Droptubes")

    End Sub
    Private Sub cmdCondAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCondAdd.Click

        If cmdCondAdd.Text = "Add New Cond" Then
            Dim NewCondModel As String
            Dim Mysentence As String

            cmdCondDelete.Enabled = False
            cmdCondSave.Enabled = True
            cmdCondAdd.Text = "Cancel Cond"
            Mysentence = InputBox("Give New Cond model a name:", "New Condenser Model", "")

            If Mysentence <> "" Then

                Dim Seperator() As Char = {" "c, "."c, ","c}
                Dim Words() As String = EnhancedSplit(Mysentence, Seperator)

                Dim Word As String
                NewCondModel = ""
                For Each Word In Words
                    NewCondModel = NewCondModel & Word
                Next

                Dim MyRow As DataRow = das1.Tables("Condenser").Rows.Find(NewCondModel)

                If MyRow Is Nothing Then
                    ddCondModel.Items.Add(NewCondModel)
                    bNewRow = True
                    ddCondModel.SelectedIndex = ddCondModel.Items.IndexOf(NewCondModel)
                    MsgBox("Complete data fields for model " & NewCondModel & " and click Save", vbOKOnly, "Complete data and save")
                Else
                    cmdCondDelete.Enabled = True
                    cmdCondSave.Enabled = False
                    cmdCondAdd.Text = "Add New Cond"
                    ddCondModel.SelectedIndex = 0
                    MsgBox("Condenser Model: " & NewCondModel & " already exists", vbOKOnly, "Model Already Exists")
                End If

            Else
                cmdCondDelete.Enabled = True
                cmdCondSave.Enabled = False
                cmdCondAdd.Text = "Add New Cond"
                ddCondModel.SelectedIndex = 0
            End If

        Else
            cmdCondDelete.Enabled = True
            cmdCondSave.Enabled = False
            cmdCondAdd.Text = "Add New Cond"
            ddCondModel.SelectedIndex = 0
        End If

    End Sub
    Private Sub cmdCondSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCondSave.Click

        Dim drwCond As DataRow = das1.Tables("Condenser").NewRow
        drwCond("Model") = ddCondModel.SelectedItem
        drwCond("CoilP") = cboCondCoilP.Text
        drwCond("Passes") = cboCondRows.Text
        drwCond("FH") = txtCondFH.Text
        drwCond("FL") = txtCondFL.Text
        drwCond("FinThk") = cboCondFinThk.Text
        drwCond("FinMat") = cboCondFinMat.Text
        drwCond("FinPI") = cboCondFinPI.Text
        drwCond("WallThk") = cboCondWallThk.Text
        drwCond("CKT1") = txtCondCKT1.Text
        drwCond("CKT2") = txtCondCKT2.Text
        drwCond("CKT") = txtCondCKT.Text
        drwCond("Split") = cboCondSplit.Text
        drwCond("Droptubes") = chkCondDroptubes.CheckState

        das1.Tables("Condenser").Rows.Add(drwCond)
        dapCond.Update(das1, "Condenser")

        cmdCondDelete.Enabled = True
        cmdCondSave.Enabled = False
        cmdCondAdd.Text = "Add New Cond"
        bNewRow = False

    End Sub
    Private Sub cmdCondDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCondDelete.Click

        Dim MyRow As DataRow = das1.Tables("Condenser").Rows.Find(ddCondModel.SelectedItem)

        If Not (MyRow Is Nothing) Then
            das1.Tables("Condenser").Rows.Find(ddCondModel.SelectedItem).Delete()
        End If

        dapCond.Update(das1, "Condenser")
        ddCondModel.Items.Remove(ddCondModel.SelectedItem)
        ddCondModel.SelectedIndex = 0

        MsgBox("Record Deleted", vbInformation, "Delete Confirm")
    End Sub
    Private Sub cmdCondUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCondUpdate.Click

        Dim drwCond As DataRow = das1.Tables("Condenser").Rows.Find(ddCondModel.SelectedItem)

            drwCond("CoilP") = cboCondCoilP.Text
            drwCond("Passes") = cboCondRows.Text
            drwCond("FH") = txtCondFH.Text
            drwCond("FL") = txtCondFL.Text
            drwCond("FinThk") = cboCondFinThk.Text
            drwCond("FinMat") = cboCondFinMat.Text
            drwCond("FinPI") = cboCondFinPI.Text
            drwCond("WallThk") = cboCondWallThk.Text
            drwCond("CKT1") = txtCondCKT1.Text
            drwCond("CKT2") = txtCondCKT2.Text
            drwCond("CKT") = txtCondCKT.Text
            drwCond("Split") = cboCondSplit.Text
            drwCond("Droptubes") = chkCondDroptubes.CheckState

            dapCond.Update(das1, "Condenser")

    End Sub

    Private Sub cboCompModel_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboCompModel.SelectedIndexChanged

        LoadCompRefDD()
        LoadCompFreDD()
        LoadCompAppDD()

    End Sub

    Private Sub cboRefCode_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboRefCode.SelectedIndexChanged

        LoadCompFreDD()
        LoadCompAppDD()

    End Sub
    Private Sub cboFreCode_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboFreCode.SelectedIndexChanged

        LoadCompAppDD()

    End Sub
    Private Sub cboAppCode_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboAppCode.SelectedIndexChanged

        If Not das1.Tables("Compressor") Is Nothing Then
            das1.Tables("Compressor").Clear()
        End If

        dapComp.SelectCommand = New OleDbCommand("SELECT * FROM Compressor WHERE Model = '" & cboCompModel.SelectedItem & "' " & _
        "AND RefCode = '" & cboRefCode.SelectedItem & "' AND FreCode = '" & cboFreCode.SelectedItem & "'" & _
        "AND AppCode = '" & cboAppCode.SelectedItem & "'", MyConnection)

        dapComp.Fill(das1, "Compressor")
        'DataGridView1.DataSource = das1.Tables("Compressor")
    End Sub

    Private Sub LoadCompRefDD()

        cboRefCode.Items.Clear()

        Dim dvComp As New DataView(das1.Tables("TempCompressor"))
        dvComp.RowFilter = "Model = '" & cboCompModel.SelectedItem & "'"
        Dim drvComp As DataRowView

        For Each drvComp In dvComp
            cboRefCode.Items.Add(drvComp("RefCode"))
        Next

        cboRefCode.SelectedIndex = 0

    End Sub

    Private Sub LoadCompFreDD()

        cboFreCode.Items.Clear()

        Dim dvComp As New DataView(das1.Tables("TempCompressor"))
        dvComp.RowFilter = "Model = '" & cboCompModel.SelectedItem & "'"
        dvComp.Sort = "RefCode"
        Dim drvComp As DataRowView

        For Each drvComp In dvComp.FindRows(cboRefCode.SelectedItem)
            cboFreCode.Items.Add(drvComp("FreCode"))
        Next

        cboFreCode.SelectedIndex = 0

    End Sub
    Private Sub LoadCompAppDD()

        cboAppCode.Items.Clear()

        Dim dvComp As New DataView(das1.Tables("TempCompressor"))
        dvComp.RowFilter = "Model = '" & cboCompModel.SelectedItem & "'"
        dvComp.Sort = "RefCode,FreCode"
        Dim drvComp As DataRowView
        Dim Key(1) As Object

        Key(0) = cboRefCode.SelectedItem
        Key(1) = cboFreCode.SelectedItem

        For Each drvComp In dvComp.FindRows(Key)
            cboAppCode.Items.Add(drvComp("AppCode"))
        Next

        cboAppCode.SelectedIndex = 0

    End Sub

    Private Sub cboUnitModel_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboUnitModel.SelectedIndexChanged

        If bNewRow = False Then
            PopulateUnit(cboUnitModel.SelectedItem)
        Else
            txtEvEaDB.Text = 80
            txtEvEaWB.Text = 67
            txtEvCFM.Text = 0
            txtCoEaDB.Text = 95
            txtCoEaWB.Text = 75
            txtCoCFM.Text = 0
            txtOutaCFM.Text = 0
            cboEvapQuantity.Text = 1
            cboCondQuantity.Text = 1
            cboCompQuantity.Text = 1
            cboRef.Text = "22"
            txtAltitude.Text = 0
            ddReheatOn_Off.Text = "None"
            cboReheatCoilP.Text = 3
            cboReheatFinThk.Text = 0.006
            cboReheatRows.Text = 1
            cboReheatWallThk.Text = 0.012
            txtReheatFL.Text = ""
            cboReheatFinPI.Text = 10
            txtReheatFH.Text = ""
            txtReheatCKT.Text = ""
            txtReheatLaDB.Text = 70
            chkReheatDroptubes.Checked = True
            txtSubcool.Text = 15
            txtSuperh.Text = 20
        End If
    End Sub

    Sub PopulateUnit(ByVal UnitModel As String)

        Dim I As Integer


        Dim drwUnit As DataRow = das1.Tables("Unit").Rows.Find(UnitModel)

        For I = 0 To ddEvapModel.Items.Count - 1
            If ddEvapModel.Items(I) = drwUnit("EvaporatorModel") Then
                ddEvapModel.SelectedIndex = I
                Exit For
            End If

        Next

        For I = 0 To ddCondModel.Items.Count - 1
            If ddCondModel.Items(I) = drwUnit("CondenserModel") Then
                ddCondModel.SelectedIndex = I
                Exit For
            End If

        Next

        For I = 0 To cboCompModel.Items.Count - 1
            If cboCompModel.Items(I) = drwUnit("CompressorModel") Then
                cboCompModel.SelectedIndex = I
                Exit For
            End If

        Next

        For I = 0 To cboRefCode.Items.Count - 1
            If cboRefCode.Items(I) = drwUnit("CompressorRef") Then
                cboRefCode.SelectedIndex = I
                Exit For
            End If

        Next

        For I = 0 To cboFreCode.Items.Count - 1
            If cboFreCode.Items(I) = drwUnit("CompressorFre") Then
                cboFreCode.SelectedIndex = I
                Exit For
            End If

        Next

        For I = 0 To cboAppCode.Items.Count - 1
            If cboAppCode.Items(I) = drwUnit("CompressorApp") Then
                cboAppCode.SelectedIndex = I
                Exit For
            End If

        Next

        txtEvEaDB.Text = drwUnit("EvapEADB")
        txtEvEaWB.Text = drwUnit("EvapEAWB")
        txtEvCFM.Text = drwUnit("EvapCFM")
        txtCoEaDB.Text = drwUnit("CondEADB")
        txtCoEaWB.Text = drwUnit("CondEAWB")
        txtCoCFM.Text = drwUnit("CondCFM")
        txtOutaCFM.Text = drwUnit("OACFM")
        cboEvapQuantity.Text = drwUnit("EvapQty")
        cboCondQuantity.Text = drwUnit("CondQty")
        cboCompQuantity.Text = drwUnit("CompQty")
        cboRef.Text = drwUnit("Refrigerant")
        txtAltitude.Text = drwUnit("Altitude")
        txtSubcool.Text = drwUnit("Subcool")
        txtSuperh.Text = drwUnit("Superh")
        ddReheatOn_Off.Text = drwUnit("Reheat")
        ET_Max.Text = drwUnit("ETMax")
        CT_Max.Text = drwUnit("CTMax")

        If drwUnit("Reheat") <> "None" Then
            GroupBox5.Visible = True
            cboReheatCoilP.Text = drwUnit("ReheatCoilP")
            cboReheatFinThk.Text = drwUnit("ReheatFinThk")
            cboReheatRows.Text = drwUnit("ReheatRows")
            cboReheatWallThk.Text = drwUnit("ReheatWallThk")
            txtReheatFL.Text = drwUnit("ReheatFL")
            cboReheatFinPI.Text = drwUnit("ReheatFPI")
            txtReheatFH.Text = drwUnit("ReheatFH")
            txtReheatCKT.Text = drwUnit("ReheatCKT")
            txtReheatLaDB.Text = drwUnit("ReheatTemp")
            chkReheatDroptubes.Checked = drwUnit("Droptubes")
        Else
            GroupBox5.Visible = False
        End If

    End Sub

    Private Sub ddReheatOn_Off_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ddReheatOn_Off.SelectedIndexChanged

        If ddReheatOn_Off.SelectedItem <> "None" Then
            GroupBox5.Visible = True
        Else
            GroupBox5.Visible = False
        End If

    End Sub

    Private Sub cmdUnitAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUnitAdd.Click

        If cmdUnitAdd.Text = "Add New Unit" Then
            Dim NewUnitModel As String
            Dim Mysentence As String

            cmdUnitDelete.Enabled = False
            cmdUnitSave.Enabled = True
            cmdUnitAdd.Text = "Cancel Unit"
            Mysentence = InputBox("Give New Unit model a name:", "New Unit Model", "")

            If Mysentence <> "" Then

                Dim Seperator() As Char = {" "c, "."c, ","c}
                Dim Words() As String = EnhancedSplit(Mysentence, Seperator)

                Dim Word As String
                NewUnitModel = ""
                For Each Word In Words
                    NewUnitModel = NewUnitModel & Word
                Next

                Dim MyRow As DataRow = das1.Tables("Unit").Rows.Find(NewUnitModel)

                If MyRow Is Nothing Then
                    cboUnitModel.Items.Add(NewUnitModel)
                    bNewRow = True
                    cboUnitModel.SelectedIndex = cboUnitModel.Items.IndexOf(NewUnitModel)
                    MsgBox("Complete data fields for model " & NewUnitModel & " and click Save", vbOKOnly, "Complete data and save")
                Else
                    cmdUnitDelete.Enabled = True
                    cmdUnitSave.Enabled = False
                    cmdUnitAdd.Text = "Add New Unit"
                    cboUnitModel.SelectedIndex = 0
                    MsgBox("Unit Model: " & NewUnitModel & " already exists", vbOKOnly, "Model Already Exists")
                End If

            Else
                cmdUnitDelete.Enabled = True
                cmdUnitSave.Enabled = False
                cmdUnitAdd.Text = "Add New Unit"
                cboUnitModel.SelectedIndex = 0
            End If

        Else
            cmdUnitDelete.Enabled = True
            cmdUnitSave.Enabled = False
            cmdUnitAdd.Text = "Add New Unit"
            cboUnitModel.SelectedIndex = 0
        End If

    End Sub

    Private Sub cmdUnitSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUnitSave.Click

        Dim drwUnit As DataRow = das1.Tables("Unit").NewRow
        drwUnit("UnitModel") = cboUnitModel.SelectedItem
        drwUnit("EvaporatorModel") = ddEvapModel.SelectedItem
        drwUnit("CondenserModel") = ddCondModel.SelectedItem
        drwUnit("CompressorModel") = cboCompModel.SelectedItem
        drwUnit("CompressorRef") = cboRefCode.SelectedItem
        drwUnit("CompressorFre") = cboFreCode.SelectedItem
        drwUnit("CompressorApp") = cboAppCode.SelectedItem
        drwUnit("EvapEADB") = txtEvEaDB.Text
        drwUnit("EvapEAWB") = txtEvEaWB.Text
        drwUnit("EvapCFM") = txtEvCFM.Text
        drwUnit("CondEADB") = txtCoEaDB.Text
        drwUnit("CondEAWB") = txtCoEaWB.Text
        drwUnit("CondCFM") = txtCoCFM.Text
        drwUnit("OACFM") = txtOutaCFM.Text
        drwUnit("EvapQty") = cboEvapQuantity.Text
        drwUnit("CondQty") = cboCondQuantity.Text
        drwUnit("CompQty") = cboCompQuantity.Text
        drwUnit("Refrigerant") = cboRef.Text
        drwUnit("Altitude") = txtAltitude.Text
        drwUnit("Subcool") = txtSubcool.Text
        drwUnit("Superh") = txtSuperh.Text
        drwUnit("Reheat") = ddReheatOn_Off.Text
        drwUnit("ETMax") = ET_Max.Text
        drwUnit("CTMax") = CT_Max.Text

        drwUnit("ReheatCoilP") = cboReheatCoilP.Text
        drwUnit("ReheatFinThk") = cboReheatFinThk.Text
        drwUnit("ReheatRows") = cboReheatRows.Text
        drwUnit("ReheatWallThk") = cboReheatWallThk.Text
        drwUnit("ReheatFL") = txtReheatFL.Text
        drwUnit("ReheatFPI") = cboReheatFinPI.Text
        drwUnit("ReheatFH") = txtReheatFH.Text
        drwUnit("ReheatCKT") = txtReheatCKT.Text
        drwUnit("ReheatTemp") = txtReheatLaDB.Text
        drwUnit("Droptubes") = chkReheatDroptubes.CheckState

        das1.Tables("Unit").Rows.Add(drwUnit)
        dapUnit.Update(das1, "Unit")

        cmdUnitDelete.Enabled = True
        cmdUnitSave.Enabled = False
        cmdUnitAdd.Text = "Add New Unit"
        bNewRow = False

    End Sub

    Private Sub cmdUnitDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUnitDelete.Click

        Dim MyRow As DataRow = das1.Tables("Unit").Rows.Find(cboUnitModel.SelectedItem)

        If Not (MyRow Is Nothing) Then
            das1.Tables("Unit").Rows.Find(cboUnitModel.SelectedItem).Delete()
        End If

        dapUnit.Update(das1, "Unit")
        cboUnitModel.Items.Remove(cboUnitModel.SelectedItem)
        cboUnitModel.SelectedIndex = 0
        MsgBox("Record Deleted", vbInformation, "Delete Confirm")
    End Sub

    Private Sub cmdUnitUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUnitUpdate.Click

        Dim drwUnit As DataRow = das1.Tables("Unit").Rows.Find(cboUnitModel.SelectedItem)

        drwUnit("EvaporatorModel") = ddEvapModel.SelectedItem
        drwUnit("CondenserModel") = ddCondModel.SelectedItem
        drwUnit("CompressorModel") = cboCompModel.SelectedItem
        drwUnit("CompressorRef") = cboRefCode.SelectedItem
        drwUnit("CompressorFre") = cboFreCode.SelectedItem
        drwUnit("CompressorApp") = cboAppCode.SelectedItem
        drwUnit("EvapEADB") = txtEvEaDB.Text
        drwUnit("EvapEAWB") = txtEvEaWB.Text
        drwUnit("EvapCFM") = txtEvCFM.Text
        drwUnit("CondEADB") = txtCoEaDB.Text
        drwUnit("CondEAWB") = txtCoEaWB.Text
        drwUnit("CondCFM") = txtCoCFM.Text
        drwUnit("OACFM") = txtOutaCFM.Text
        drwUnit("EvapQty") = cboEvapQuantity.Text
        drwUnit("CondQty") = cboCondQuantity.Text
        drwUnit("CompQty") = cboCompQuantity.Text
        drwUnit("Refrigerant") = cboRef.Text
        drwUnit("Altitude") = txtAltitude.Text
        drwUnit("Subcool") = txtSubcool.Text
        drwUnit("Superh") = txtSuperh.Text
        drwUnit("Reheat") = ddReheatOn_Off.Text

        drwUnit("ReheatCoilP") = cboReheatCoilP.Text
        drwUnit("ReheatFinThk") = cboReheatFinThk.Text
        drwUnit("ReheatRows") = cboReheatRows.Text
        drwUnit("ReheatWallThk") = cboReheatWallThk.Text
        drwUnit("ReheatFL") = txtReheatFL.Text
        drwUnit("ReheatFPI") = cboReheatFinPI.Text
        drwUnit("ReheatFH") = txtReheatFH.Text
        drwUnit("ReheatCKT") = txtReheatCKT.Text
        drwUnit("ReheatTemp") = txtReheatLaDB.Text
        drwUnit("Droptubes") = chkReheatDroptubes.CheckState

        dapUnit.Update(das1, "Unit")

    End Sub

    Public Sub cmdCalculate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCalculate.Click

        If frmUnitOutput.Visible = True Then
            frmUnitOutput.Close()
        End If

        mEvaPClass = New pjEvap.EvapClass
        mConPClass = New pjCond.CondClass
        mComPClass = New pjComp.CompClass
        mRecPClass = New pjReheat.ReheatClass
        'mRefClass = New pjRefProp.RefProp

        modMain.dsItems.Clear()
        modMain.dtItems.Clear()

        If Not ValidateInputs() Then
            Exit Sub
        End If

        Call AssignEvapProperties()
        Call AssignCondProperties()
        Call AssignCompProperties()

        mEvaPClass.EvapProperties = struEvap
        mConPClass.CondProperties = struCond
        mComPClass.CompProperties = struComp

        'mEvaPClass.RefClassReference = mRefClass
        'mConPClass.RefClassReference = mRefClass

        If ddReheatOn_Off.Text <> "None" Then

            Call AssignReheatProperties()
            mRecPClass.ReheatProperties = struReheat
            'mRecPClass.RefClassReference = mRefClass
        End If

        If AssignProperties() Then

            If optAirCond.Checked = True Then
                If mSysP.i_CompQuantity = 2 And mSysP.i_CondQuantity = 2 And mEvaPClass.EvapProperties.s_Split = "IT" Then
                    Call StartSys_Balance_2Cond_AC()
                ElseIf mSysP.i_CompQuantity = 2 And mConPClass.CondProperties.s_Split = "FS" And mEvaPClass.EvapProperties.s_Split = "IT" Then
                    Call StartSys_Balance_2Cond_AC()
                ElseIf mSysP.i_CompQuantity = 2 And mEvaPClass.EvapProperties.s_Split = "IT" And mConPClass.CondProperties.s_Split = "IT" Then
                    Call StartSys_Balance_ITEITC_AC()
                ElseIf mSysP.i_CompQuantity = 2 And mEvaPClass.EvapProperties.s_Split = "FS" And mConPClass.CondProperties.s_Split = "IT" Then
                    Call StartSys_Balance_FSEITC_AC()
                Else
                    Call StartSys_Balance_AC()
                End If

            ElseIf optHeatPump.Checked = True Then

                If mSysP.i_CompQuantity = 2 And mSysP.i_EvapQuantity = 2 And mConPClass.CondProperties.s_Split = "IT" Then
                    Call StartSys_Balance_2Evap_HP()
                ElseIf mSysP.i_CompQuantity = 2 And mEvaPClass.EvapProperties.s_Split = "FS" And mConPClass.CondProperties.s_Split = "IT" Then
                    Call StartSys_Balance_2Evap_HP()
                ElseIf mSysP.i_CompQuantity = 2 And mEvaPClass.EvapProperties.s_Split = "IT" And mConPClass.CondProperties.s_Split = "IT" Then
                    Call StartSys_Balance_ITEITC_HP()
                ElseIf mSysP.i_CompQuantity = 2 And mEvaPClass.EvapProperties.s_Split = "IT" And mConPClass.CondProperties.s_Split = "FS" Then
                    Call StartSys_Balance_ITEFSC_HP()
                Else
                    Call StartSys_Balance_HP()
                End If

            End If

            If optSales.Checked = True Then
                Dim loopVar As Integer

                For loopVar = 0 To modMain.dsItems.Tables(0).Rows.Count - 1
                    Dim drw1 As DataRow = modMain.dsItems.Tables(0).Rows.Find(loopVar)
                    If Not (drw1 Is Nothing) Then
                        modMain.dsItems.Tables(0).Rows.Find(loopVar).Delete()
                    End If
                Next
            End If

            Call ShowReport(modMain.dsItems)

        End If

    End Sub

    Private Sub AssignEvapProperties()

        struEvap.s_Model = ddEvapModel.Text
        struEvap.n_FH = txtEvapFH.Text
        struEvap.n_FL = txtEvapFL.Text
        struEvap.i_Rows = cboEvapRows.Text
        struEvap.i_FPI = cboEvapFinPI.Text
        struEvap.n_FinThk = cboEvapFinThk.Text
        struEvap.s_CoilPat = cboEvapCoilP.Text
        struEvap.n_WallThk = cboEvapWallThk.Text
        struEvap.s_Split = cboEvapSplit.Text
        struEvap.i_CKT = txtEvapCKT.Text
        struEvap.i_CKT1 = txtEvapCKT1.Text
        struEvap.i_CKT2 = txtEvapCKT2.Text
        struEvap.s_Droptubes = chkEvapDroptubes.CheckState
        struEvap.s_FinMat = cboEvapFinMat.Text
        struEvap.i_EvapQuantity = cboEvapQuantity.Text
        struEvap.i_CompQuantity = cboCompQuantity.Text
        '    struEvap.s_CompAPPCode = cboAppCode.Text
        struEvap.s_Refrigerant = cboRef.Text
        struEvap.n_Altitude = txtAltitude.Text

    End Sub

    Private Sub AssignCondProperties()

        struCond.s_Model = ddCondModel.Text
        struCond.n_FH = txtCondFH.Text
        struCond.n_FL = txtCondFL.Text
        struCond.i_FPI = cboCondFinPI.Text
        struCond.i_Rows = cboCondRows.Text
        struCond.n_FinThk = cboCondFinThk.Text
        struCond.s_CoilPat = cboCondCoilP.Text
        struCond.n_WallThk = cboCondWallThk.Text
        struCond.s_Split = cboCondSplit.Text
        struCond.s_FinMat = cboCondFinMat.Text
        struCond.s_Droptubes = chkCondDroptubes.CheckState
        struCond.i_CKT = txtCondCKT.Text
        struCond.i_CKT1 = txtCondCKT1.Text
        struCond.i_CKT2 = txtCondCKT2.Text
        struCond.i_CondQuantity = cboCondQuantity.Text
        struCond.i_CompQuantity = cboCompQuantity.Text
        '    struCond.s_CompAPPCode = cboAppCode.Text
        struCond.s_Refrigerant = cboRef.Text
        struCond.s_ReheatType = ddReheatOn_Off.Text
        struCond.n_Altitude = txtAltitude.Text
    End Sub

    Private Sub AssignCompProperties()

        Dim I_X As Integer
        Dim drwComp As DataRow
        Dim Strname(40) As String
        Dim Dblvalue(40) As Double

        drwComp = das1.Tables("Compressor").Rows(0)

        struComp.s_Model = drwComp("Model")
        struComp.s_RefCode = drwComp("RefCode")
        struComp.s_FreCode = drwComp("FreCode")
        struComp.s_AppCode = drwComp("AppCode")

        ReDim struComp.d_CF(40)

        For I_X = 0 To 39

            Strname.SetValue("CF" & I_X.ToString(), I_X)
            Dblvalue.SetValue(drwComp(Strname(I_X)), I_X)
            struComp.d_CF(I_X) = Dblvalue(I_X)

        Next

        struComp.n_Subcool = txtSubcool.Text
        struComp.n_Superh = txtSuperh.Text

    End Sub

    Private Sub AssignReheatProperties()

        struReheat.n_LaDB = txtReheatLaDB.Text
        struReheat.n_FH = txtReheatFH.Text
        struReheat.n_FL = txtReheatFL.Text
        struReheat.i_FPI = cboReheatFinPI.Text
        struReheat.i_Rows = cboReheatRows.Text
        struReheat.n_FinThk = cboReheatFinThk.Text
        struReheat.s_CoilP = cboReheatCoilP.Text
        struReheat.n_WallThk = cboReheatWallThk.Text
        struReheat.s_Droptubes = chkReheatDroptubes.CheckState
        struReheat.i_CKT = txtReheatCKT.Text
        struReheat.i_CompQuantity = cboCompQuantity.Text
        struReheat.s_ReheatType = ddReheatOn_Off.Text
        struReheat.s_Refrigerant = cboRef.Text
        struReheat.n_MassFraction = 0.5

    End Sub

    Public Function AssignProperties() As Boolean

        Dim Compquantity As Integer

        mSysP.n_EvCFM = Val(txtEvCFM.Text)
        mSysP.n_CoCFM = Val(txtCoCFM.Text)
        mSysP.n_OaCFM = Val(txtOutaCFM.Text)
        mSysP.n_EvEaDB = Val(txtEvEaDB.Text)
        mSysP.n_CoEaDB = Val(txtCoEaDB.Text)
        mSysP.n_EvEaWB = Val(txtEvEaWB.Text)
        mSysP.n_CoEaWB = Val(txtCoEaWB.Text)
        mSysP.n_Subcool = Val(txtSubcool.Text)
        mSysP.n_Superh = Val(txtSuperh.Text)
        mSysP.n_Altitude = Val(txtAltitude.Text)
        mSysP.n_Refrigerant = cboRef.Text
        mSysP.i_CondQuantity = CInt(cboCondQuantity.Text)
        mSysP.i_EvapQuantity = CInt(cboEvapQuantity.Text)
        mSysP.i_CompQuantity = CInt(cboCompQuantity.Text)
        mSysP.s_CompAppCode = cboAppCode.Text
        mSysP.n_ReheatYes_No = ddReheatOn_Off.Text
        mSysP.s_UnitModel = cboUnitModel.Text
        mSysP.n_ETMax = Val(ET_Max.Text)
        mSysP.n_CTMax = Val(CT_Max.Text)

        If ddReheatOn_Off.Text <> "None" Then

            Call AssignReheatProperties()

        End If

        Compquantity = mSysP.i_CompQuantity

        If ddReheatOn_Off.Text <> "None" Then

            If mConPClass.CondCirCuitryCheck() And mEvaPClass.EvapCircuitryCheck(Compquantity) And mRecPClass.ReheCircuitryCheck() Then
                AssignProperties = True
            End If
        Else

            If mConPClass.CondCirCuitryCheck() And mEvaPClass.EvapCircuitryCheck(Compquantity) Then
                AssignProperties = True
            End If

        End If

    End Function ' AssignProperties

    Private Function ValidateInputs() As Boolean

        Dim output As String = ""
        Dim dblValue As Double
        Dim intValue As Integer
        Dim EvapFPM As Double
        Dim CondFPM As Double

        If Not Double.TryParse(txtCondFH.Text, dblValue) Then
            Output = Output & "Condenser Fin Height must be numeric." & vbCrLf
        ElseIf dblValue <= 0 Then
            output = output & "Condenser Fin Height must be greater than zero." & vbCrLf
        End If

        If Not Double.TryParse(txtCondFL.Text, dblValue) Then
            output = output & "Condenser Fin Length must be numeric." & vbCrLf
        ElseIf dblValue <= 0 Then
            output = output & "Condenser Fin Length must be greater than zero." & vbCrLf
        End If

        If Not Integer.TryParse(txtCondCKT.Text, intValue) Then
            output = output & "Condenser Feeds must be an integer." & vbCrLf
        ElseIf intValue <= 0 Then
            output = output & "Condenser Feeds must be greater than zero." & vbCrLf
        End If

        If Not Integer.TryParse(txtCondCKT1.Text, intValue) Then
            output = output & "Condenser Feeds must be an integer." & vbCrLf
        ElseIf intValue < 0 Then
            output = output & "Condenser circuit 1 Feeds must be a positive integer or zero." & vbCrLf
        End If

        If Not Integer.TryParse(txtCondCKT2.Text, intValue) Then
            output = output & "Condenser Feeds must be an integer." & vbCrLf
        ElseIf intValue < 0 Then
            output = output & "Condenser Circuit 2 Feeds must be a positive integer or zero." & vbCrLf
        End If

        If Not Double.TryParse(txtEvapFH.Text, dblValue) Then
            output = output & "Evaporator Fin Height must be numeric." & vbCrLf
        ElseIf dblValue <= 0 Then
            output = output & "Evaporator Fin Height must be greater than zero." & vbCrLf
        End If

        If Not Double.TryParse(txtEvapFL.Text, dblValue) Then
            output = output & "Evaporator Fin Length must be numeric." & vbCrLf
        ElseIf dblValue <= 0 Then
            output = output & "Evaporator Fin Length must be greater than zero." & vbCrLf
        End If

        If Not Integer.TryParse(txtEvapCKT.Text, intValue) Then
            output = output & "Evaporator Feeds must be numeric." & vbCrLf
        ElseIf intValue <= 0 Then
            output = output & "Evaporator Feeds must be greater than zero." & vbCrLf
        End If

        If Not Integer.TryParse(txtEvapCKT1.Text, intValue) Then
            output = output & "Evaporator Circuit 1 Feeds must be numeric." & vbCrLf
        ElseIf intValue < 0 Then
            output = output & "Evaporator Circuit 1 Feeds must be positive number." & vbCrLf
        End If

        If Not Integer.TryParse(txtEvapCKT2.Text, intValue) Then
            output = output & "Evaporator Circuit 2 Feeds must be numeric." & vbCrLf
        ElseIf intValue < 0 Then
            output = output & "Evaporator Circuit 2 Feeds must be positive number." & vbCrLf
        End If

        If Not Double.TryParse(txtOutaCFM.Text, dblValue) Then
            output = output & "Outside Air CFM must be numeric" & vbCrLf
        ElseIf dblValue < 0 Then
            output = output & "Outside Air CFM must be positive number or zero" & vbCrLf
        End If

        If Not Double.TryParse(txtEvCFM.Text, dblValue) Then
            output = output & "Evaporator Return CFM must be numeric" & vbCrLf
        ElseIf dblValue < 0 Then
            output = output & "Evaporator Return CFM must be positive number or zero" & vbCrLf
        End If

        If Not Double.TryParse(txtCoCFM.Text, dblValue) Then
            output = output & "Condenser CFM must be numeric" & vbCrLf
        ElseIf dblValue <= 0 Then
            output = output & "Condenser CFM must be greater than 0" & vbCrLf
        End If

        If Not Double.TryParse(txtEvEaDB.Text, dblValue) Then
            output = output & "Return Air DB must be numeric" & vbCrLf
        ElseIf dblValue <= 14.9 Or dblValue > 125 Then
            output = output & "Return Air DB must be greater than 15 and lower than 125" & vbCrLf
        End If

        If Not Double.TryParse(txtCoEaDB.Text, dblValue) Then
            output = output & "Outside Air DB must be numeric" & vbCrLf
        ElseIf dblValue <= 14.9 Or dblValue > 125 Then
            output = output & "Outside Air DB must be greater than 15 and lower than 125" & vbCrLf
        End If

        If Not Double.TryParse(txtEvEaWB.Text, dblValue) Then
            output = output & "Return Air WB must be numeric" & vbCrLf
        ElseIf dblValue <= 11.9 Or dblValue > CDbl(txtEvEaDB.Text) Then
            output = output & "Return Air WB must be greater than 12 and less than DB" & vbCrLf
        End If

        If Not Double.TryParse(txtCoEaWB.Text, dblValue) Then
            output = output & "Outside Air WB must be numeric" & vbCrLf
        ElseIf dblValue <= 0 Or dblValue > CDbl(txtCoEaDB.Text) Then
            output = output & "Outside Air WB must be greater than zero and less than DB" & vbCrLf
        End If

        If Not Double.TryParse(txtSubcool.Text, dblValue) Then
            output = output & "Subcooling Temperature must be numeric" & vbCrLf
        ElseIf dblValue < 10 Then
            output = output & "Subcooling Temperature must be greater than 10" & vbCrLf
        End If

        If Not Double.TryParse(txtSuperh.Text, dblValue) Then
            output = output & "Superheat Temperature must be numeric" & vbCrLf
        ElseIf dblValue < 10 Then
            output = output & "Superheat Temperature must be greater than 10" & vbCrLf
        End If

        If Not Double.TryParse(txtAltitude.Text, dblValue) Then
            output = output & "Altitude must be numeric" & vbCrLf
        ElseIf dblValue < 0 Then
            output = output & "Altitude must be positive number or zero" & vbCrLf
        End If

        If cboRef.Text <> cboRefCode.Text Then
            output = output & "Match System and Compressor Refrigerant" & vbCrLf
        End If

        If Val(txtCoCFM.Text) > (Val(txtOutaCFM.Text) + Val(txtEvCFM.Text)) And optHeatPump.Checked = True Then
            output = output & "Please select Air Conditioning option" & vbCrLf
        End If

        If Val(txtCoCFM.Text) <= (Val(txtOutaCFM.Text) + Val(txtEvCFM.Text)) And optHeatPump.Checked = False Then
            output = output & "Please select Heat Pump option" & vbCrLf
        End If

        If optAirCond.Checked = True Then

            If (Val(txtEvCFM.Text) + Val(txtOutaCFM.Text)) = 0 Then
                output = output & "Sum of Return & Outside Air CFM must be greater than 0" & vbCrLf
            End If

            If CInt(cboCompQuantity.Text) = 2 Then
                If cboEvapSplit.Text = "NS" And CInt(cboCondQuantity.Text) = 2 And CInt(cboEvapQuantity.Text) = 1 Then
                    output = output & "Check the Evaporator Split" & vbCrLf
                End If
                If cboEvapSplit.Text = "NS" And CInt(cboCondQuantity.Text) = 1 And CInt(cboEvapQuantity.Text) = 1 Then
                    output = output & "Check the Compressor Quantity" & vbCrLf
                End If
                If cboEvapSplit.Text = "NS" And CInt(cboCondQuantity.Text) = 1 And CInt(cboEvapQuantity.Text) = 2 Then
                    output = output & "Treat this combination as a FS evaporator" & vbCrLf
                End If
                If cboEvapSplit.Text = "IT" And CInt(cboCondQuantity.Text) = 1 And cboCondSplit.Text = "NS" Then
                    output = output & "Check the Condenser Quantity" & vbCrLf
                End If
                If cboEvapSplit.Text = "IT" And CInt(cboEvapQuantity.Text) = 2 Then
                    output = output & "Check the Evaporator Quantity" & vbCrLf
                End If
                If cboEvapSplit.Text = "FS" And CInt(cboCondQuantity.Text) = 1 And cboCondSplit.Text = "NS" Then
                    output = output & "Check the Condenser Quantity" & vbCrLf
                End If
                If cboEvapSplit.Text = "FS" And CInt(cboEvapQuantity.Text) = 2 Then
                    output = output & "Check the Evaporator Quantity" & vbCrLf
                End If
                If cboCondSplit.Text = "IT" And CInt(cboCondQuantity.Text) = 2 Then
                    output = output & "Check the Condenser Quantity" & vbCrLf
                End If
                If cboCondSplit.Text = "FS" And CInt(cboCondQuantity.Text) = 2 Then
                    output = output & "Check the Condenser Quantity" & vbCrLf
                End If

            ElseIf CInt(cboCompQuantity.Text) = 1 Then
                If cboEvapSplit.Text = "NS" And CInt(cboEvapQuantity.Text) = 2 Then
                    output = output & "Check the Evaporator Quantity" & vbCrLf
                End If
                If cboEvapSplit.Text = "NS" And CInt(cboCondQuantity.Text) = 2 Then
                    output = output & "Check the Condenser Quantity" & vbCrLf
                End If

            End If

            If CInt(cboEvapQuantity.Text) = 1 Then
                EvapFPM = (Val(txtOutaCFM.Text) + Val(txtEvCFM.Text)) / (Val(txtEvapFH.Text) * Val(txtEvapFL.Text))
                EvapFPM = EvapFPM * 144
            Else
                EvapFPM = (Val(txtOutaCFM.Text) + Val(txtEvCFM.Text)) / (Val(txtEvapFH.Text) * Val(txtEvapFL.Text) * 2)
                EvapFPM = EvapFPM * 144
            End If

            If CInt(cboCondQuantity.Text) = 1 Then
                CondFPM = Val(txtCoCFM.Text) / (Val(txtCondFH.Text) * Val(txtCondFL.Text) / 144)
            Else
                CondFPM = Val(txtCoCFM.Text) / (Val(txtCondFH.Text) * Val(txtCondFL.Text) * 2 / 144)
            End If

            If EvapFPM > 650 Then
                output = output & "Air Flow velocity over the Evap must be below 650 FPM" & vbCrLf
            End If

            If CondFPM > 800 Then
                output = output & "Air Flow velocity over the Cond must be below 800 FPM" & vbCrLf
            End If

        ElseIf optHeatPump.Checked = True Then

            If (Val(txtCoCFM.Text) + Val(txtOutaCFM.Text)) = 0 Then
                output = output & "Sum of Return & Outside Air CFM must be greater than 0" & vbCrLf
            End If

            If CInt(cboCompQuantity.Text) = 2 Then
                If cboCondSplit.Text = "NS" And CInt(cboEvapQuantity.Text) = 2 And CInt(cboCondQuantity.Text) = 1 Then
                    output = output & "Check the Condenser Split" & vbCrLf
                End If
                If cboCondSplit.Text = "NS" And CInt(cboEvapQuantity.Text) = 1 And CInt(cboCondQuantity.Text) = 1 Then
                    output = output & "Check the Compressor Quantity" & vbCrLf
                End If
                If cboCondSplit.Text = "NS" And CInt(cboEvapQuantity.Text) = 1 And CInt(cboCondQuantity.Text) = 2 Then
                    output = output & "Treat this combination as a FS condenser" & vbCrLf
                End If
                If cboCondSplit.Text = "IT" And CInt(cboEvapQuantity.Text) = 1 And cboEvapSplit.Text = "NS" Then
                    output = output & "Check the evaporatorr Quantity" & vbCrLf
                End If
                If cboCondSplit.Text = "IT" And CInt(cboCondQuantity.Text) = 2 Then
                    output = output & "Check the Condenser Quantity" & vbCrLf
                End If
                If cboCondSplit.Text = "FS" And CInt(cboEvapQuantity.Text) = 1 And cboEvapSplit.Text = "NS" Then
                    output = output & "Check the Evaporator Quantity" & vbCrLf
                End If
                If cboCondSplit.Text = "FS" And CInt(cboCondQuantity.Text) = 2 Then
                    output = output & "Check the Condenser Quantity" & vbCrLf
                End If

            ElseIf CInt(cboCompQuantity.Text) = 1 Then
                If cboCondSplit.Text = "NS" And CInt(cboCondQuantity.Text) = 2 Then
                    output = output & "Check the Condenser Quantity" & vbCrLf
                End If
                If cboCondSplit.Text = "NS" And CInt(cboEvapQuantity.Text) = 2 Then
                    output = output & "Check the Evaporator Quantity" & vbCrLf
                End If

            End If

            If CInt(cboEvapQuantity.Text) = 1 Then
                EvapFPM = Val(txtEvCFM.Text) / (Val(txtEvapFH.Text) * Val(txtEvapFL.Text))
                EvapFPM = EvapFPM * 144
            Else
                EvapFPM = Val(txtEvCFM.Text) / (Val(txtEvapFH.Text) * Val(txtEvapFL.Text) * 2)
                EvapFPM = EvapFPM * 144
            End If

            If CInt(cboCondQuantity.Text) = 1 Then
                CondFPM = (Val(txtOutaCFM.Text) + Val(txtCoCFM.Text)) / (Val(txtCondFH.Text) * Val(txtCondFL.Text) / 144)
            Else
                CondFPM = (Val(txtOutaCFM.Text) + Val(txtCoCFM.Text)) / (Val(txtCondFH.Text) * Val(txtCondFL.Text) * 2 / 144)
            End If

            If EvapFPM > 800 Then
                output = output & "Air Flow velocity over the Evap must be below 800 FPM" & vbCrLf
            End If

            If CondFPM > 650 Then
                output = output & "Air Flow velocity over the Cond must be below 650 FPM" & vbCrLf
            End If

        End If

        If optHeatPump.Checked = True And ddReheatOn_Off.Text <> "None" Then
            output = output & "Please select Reheat None option" & vbCrLf
        Else
            If cboAppCode.Text <> "AC" And ddReheatOn_Off.Text <> "None" Then
                output = output & "Please select Reheat None option" & vbCrLf
            ElseIf CInt(cboCondQuantity.Text) = 2 And CInt(cboCompQuantity.Text) = 1 And ddReheatOn_Off.Text <> "None" Then
                output = output & "Please select Reheat None option" & vbCrLf
            ElseIf cboCondSplit.Text = "IT" And CInt(cboCompQuantity.Text) = 1 And ddReheatOn_Off.Text <> "None" Then
                output = output & "Please select Reheat None option" & vbCrLf
            ElseIf cboCondSplit.Text = "FS" And CInt(cboCompQuantity.Text) = 1 And ddReheatOn_Off.Text <> "None" Then
                output = output & "Please select Reheat None option" & vbCrLf
            End If
        End If

        If ddReheatOn_Off.Text <> "None" Then
            If Not Integer.TryParse(txtReheatCKT.Text, intValue) Then
                output = output & "Reheat Feeds must be numeric." & vbCrLf
            ElseIf intValue <= 0 Then
                output = output & "Reheat Feeds must be greater than zero." & vbCrLf
            End If

            If Not Double.TryParse(txtReheatFL.Text, dblValue) Then
                output = output & "Reheat Fin Length must be numeric." & vbCrLf
            ElseIf dblValue <= 0 Then
                output = output & "Reheat Fin Length must be greater than zero." & vbCrLf
            End If

            If Not Double.TryParse(txtReheatFH.Text, dblValue) Then
                output = output & "Reheat Fin Height must be numeric." & vbCrLf
            ElseIf dblValue <= 0 Then
                output = output & "Reheat Fin Height must be greater than zero." & vbCrLf
            End If

            If Not Double.TryParse(txtReheatLaDB.Text, dblValue) Then
                output = output & " Desired leaving air DB must be numeric" & vbCrLf
            ElseIf dblValue < 0 Then
                output = output & " Desired leaving air DB must be positive number" & vbCrLf
            End If
        End If


        If output <> "" Then
            MessageBox.Show(output, "Invalid Inputs", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

            ValidateInputs = False
        Else
            ValidateInputs = True
        End If

    End Function

    Friend Sub optAirCond_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optAirCond.CheckedChanged

        Dim NewUnitModel As String
        Dim NewEvapmodel As String
        Dim NewCondmodel As String
        Dim NewCompmodel As String
        Dim NewRefCode As String
        Dim NewFreCode As String
        Dim NewAppCode As String
        Dim OldEvapQty As Integer
        Dim OldCondQty As Integer
        Dim OldCompQty As Integer
        Dim OldEvapCFM As Double
        Dim OldCondCFM As Double
        Dim OldOutaCFM As Double
        Dim OldSubcool As Double
        Dim OldSuperh As Double
        Dim OldAltitude As Double
        Dim Str1 As String
        Dim Str2 As String
        Dim Pos As Integer

        NewEvapmodel = ddCondModel.Text
        NewCondmodel = ddEvapModel.Text
        NewCompmodel = cboCompModel.Text
        NewRefCode = cboRefCode.Text
        NewFreCode = cboFreCode.Text
        NewAppCode = cboAppCode.Text

        Str1 = cboUnitModel.Text
        Str2 = "HeatPump"
        Pos = InStr(Str1, Str2)

        If Pos = 0 Then
            Exit Sub
        End If

        OldEvapCFM = Val(txtEvCFM.Text)
        OldCondCFM = Val(txtCoCFM.Text)
        OldOutaCFM = Val(txtOutaCFM.Text)
        OldEvapQty = Val(cboEvapQuantity.Text)
        OldCompQty = Val(cboCompQuantity.Text)
        OldCondQty = Val(cboCondQuantity.Text)
        OldSubcool = Val(txtSubcool.Text)
        OldSuperh = Val(txtSuperh.Text)
        OldAltitude = Val(txtAltitude.Text)

        NewUnitModel = Mid(cboUnitModel.Text, 1, Pos - 1)

        Dim MyRow As DataRow = das1.Tables("Unit").Rows.Find(NewUnitModel)

        If MyRow Is Nothing Then

            cboUnitModel.Text = (NewUnitModel)
            cboUnitModel.Items.Add(NewUnitModel)

            Call AddSaveACUnit(OldEvapCFM, OldCondCFM, OldOutaCFM, OldEvapQty, OldCondQty, OldCompQty, _
                               OldSubcool, OldSuperh, OldAltitude)
            Call AddSaveEvap(NewEvapmodel)
            Call AddSaveCond(NewCondmodel)
            cboUnitModel.SelectedIndex = cboUnitModel.Items.IndexOf(NewUnitModel)
        Else

            Call AddSaveEvap(NewEvapmodel)
            Call AddSaveCond(NewCondmodel)
            Call UpdateACUnit(NewUnitModel, OldEvapCFM, OldCondCFM, OldOutaCFM, OldSubcool, OldSuperh, OldAltitude, _
                                NewCompmodel, NewRefCode, NewFreCode, NewAppCode)
            cboUnitModel.SelectedIndex = cboUnitModel.Items.IndexOf(NewUnitModel)

        End If

    End Sub

    Friend Sub optHeatPump_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optHeatPump.CheckedChanged

        Dim NewUnitModel As String
        Dim NewEvapmodel As String
        Dim NewCondmodel As String
        Dim NewCompmodel As String
        Dim NewRefCode As String
        Dim NewFreCode As String
        Dim NewAppCode As String
        Dim OldEvapQty As Integer
        Dim OldCondQty As Integer
        Dim OldCompQty As Integer
        Dim OldEvapCFM As Double
        Dim OldCondCFM As Double
        Dim OldOutaCFM As Double
        Dim OldSubcool As Double
        Dim OldSuperh As Double
        Dim OldAltitude As Double
        Dim Str1 As String
        Dim Str2 As String
        Dim Pos As Integer

        NewEvapmodel = ddCondModel.Text
        NewCondmodel = ddEvapModel.Text
        NewCompmodel = cboCompModel.Text
        NewRefCode = cboRefCode.Text
        NewFreCode = cboFreCode.Text
        NewAppCode = cboAppCode.Text

        Str1 = cboUnitModel.Text
        Str2 = "HeatPump"
        Pos = InStr(Str1, Str2)

        If Pos <> 0 Then
            Exit Sub
        End If

        OldEvapCFM = Val(txtEvCFM.Text)
        OldCondCFM = Val(txtCoCFM.Text)
        OldOutaCFM = Val(txtOutaCFM.Text)
        OldEvapQty = Val(cboEvapQuantity.Text)
        OldCompQty = Val(cboCompQuantity.Text)
        OldCondQty = Val(cboCondQuantity.Text)
        OldSubcool = Val(txtSubcool.Text)
        OldSuperh = Val(txtSuperh.Text)
        OldAltitude = Val(txtAltitude.Text)

        NewUnitModel = cboUnitModel.Text & "HeatPump"

        Dim MyRow As DataRow = das1.Tables("Unit").Rows.Find(NewUnitModel)

        If MyRow Is Nothing Then

            cboUnitModel.Text = NewUnitModel
            cboUnitModel.Items.Add(NewUnitModel)

            Call AddSaveHPUnit(OldEvapCFM, OldCondCFM, OldOutaCFM, OldEvapQty, OldCondQty, OldCompQty, _
                               OldSubcool, OldSuperh, OldAltitude)
            Call AddSaveEvap(NewEvapmodel)
            Call AddSaveCond(NewCondmodel)
            cboUnitModel.SelectedIndex = cboUnitModel.Items.IndexOf(NewUnitModel)

        Else

            Call AddSaveEvap(NewEvapmodel)
            Call AddSaveCond(NewCondmodel)
            Call UpdateHPUnit(NewUnitModel, OldEvapCFM, OldCondCFM, OldOutaCFM, OldSubcool, OldSuperh, OldAltitude, _
                                NewCompmodel, NewRefCode, NewFreCode, NewAppCode)
            cboUnitModel.SelectedIndex = cboUnitModel.Items.IndexOf(NewUnitModel)

        End If

    End Sub

    Private Sub AddSaveACUnit(ByVal OldEvapCFM As Double, ByVal OldCondCFM As Double, ByVal OldOutaCFM As Double, _
    ByVal OldEvapQty As Integer, ByVal OldCondQty As Integer, ByVal OldCompQty As Integer, _
    ByVal OldSubcool As Double, ByVal OldSuperh As Double, ByVal OldAltitude As Double)

        Dim drwUnit As DataRow = das1.Tables("Unit").NewRow

        drwUnit("UnitModel") = cboUnitModel.Text
        drwUnit("EvaporatorModel") = ddCondModel.SelectedItem
        drwUnit("CondenserModel") = ddEvapModel.SelectedItem
        drwUnit("CompressorModel") = cboCompModel.SelectedItem
        drwUnit("CompressorRef") = cboRefCode.SelectedItem
        drwUnit("CompressorFre") = cboFreCode.SelectedItem
        drwUnit("CompressorApp") = cboAppCode.SelectedItem
        drwUnit("EvapEADB") = 80
        drwUnit("EvapEAWB") = 67
        drwUnit("EvapCFM") = OldCondCFM
        drwUnit("CondEADB") = 95
        drwUnit("CondEAWB") = 75
        drwUnit("CondCFM") = OldEvapCFM
        drwUnit("OACFM") = OldOutaCFM
        drwUnit("EvapQty") = OldCondQty
        drwUnit("CondQty") = OldEvapQty
        drwUnit("CompQty") = OldCompQty
        drwUnit("Refrigerant") = cboRef.Text
        drwUnit("Reheat") = "None"
        drwUnit("Subcool") = OldSubcool
        drwUnit("Superh") = OldSuperh
        drwUnit("Altitude") = OldAltitude
        drwUnit("ETMax") = 55
        drwUnit("CTMax") = 150

        das1.Tables("Unit").Rows.Add(drwUnit)
        dapUnit.Update(das1, "Unit")

    End Sub

    Private Sub AddSaveHPUnit(ByVal OldEvapCFM As Double, ByVal OldCondCFM As Double, ByVal OldOutaCFM As Double, _
    ByVal OldEvapQty As Integer, ByVal OldCondQty As Integer, ByVal OldCompQty As Integer, _
    ByVal OldSubcool As Double, ByVal OldSuperh As Double, ByVal OldAltitude As Double)

        Dim drwUnit As DataRow = das1.Tables("Unit").NewRow

        drwUnit("UnitModel") = cboUnitModel.Text
        drwUnit("EvaporatorModel") = ddCondModel.SelectedItem
        drwUnit("CondenserModel") = ddEvapModel.SelectedItem
        drwUnit("CompressorModel") = cboCompModel.SelectedItem
        drwUnit("CompressorRef") = cboRefCode.SelectedItem
        drwUnit("CompressorFre") = cboFreCode.SelectedItem
        drwUnit("CompressorApp") = cboAppCode.SelectedItem
        drwUnit("EvapEADB") = 47
        drwUnit("EvapEAWB") = 43
        drwUnit("EvapCFM") = OldCondCFM
        drwUnit("CondEADB") = 70
        drwUnit("CondEAWB") = 57
        drwUnit("CondCFM") = OldEvapCFM
        drwUnit("OACFM") = OldOutaCFM
        drwUnit("EvapQty") = OldCondQty
        drwUnit("CondQty") = OldEvapQty
        drwUnit("CompQty") = OldCompQty
        drwUnit("Refrigerant") = cboRefCode.Text
        drwUnit("Reheat") = "None"
        drwUnit("Subcool") = OldSubcool
        drwUnit("Superh") = OldSuperh
        drwUnit("Altitude") = OldAltitude
        drwUnit("ETMax") = 55
        drwUnit("CTMax") = 150

        das1.Tables("Unit").Rows.Add(drwUnit)
        dapUnit.Update(das1, "Unit")

    End Sub

    Private Sub AddSaveEvap(ByVal NewEvapmodel As String)

        Dim MyRow As DataRow = das1.Tables("Evaporator").Rows.Find(NewEvapmodel)

        If MyRow Is Nothing Then

            ddEvapModel.Items.Add(NewEvapmodel)

            Dim drwEvap As DataRow = das1.Tables("Evaporator").NewRow

            drwEvap("Model") = NewEvapmodel
            drwEvap("FH") = txtCondFH.Text
            drwEvap("FL") = txtCondFL.Text
            drwEvap("FinPI") = cboCondFinPI.Text
            drwEvap("Passes") = cboCondRows.Text
            drwEvap("FinThk") = cboCondFinThk.Text
            drwEvap("CoilP") = cboCondCoilP.Text
            drwEvap("WallThk") = cboCondWallThk.Text
            drwEvap("FinMat") = cboCondFinMat.Text
            drwEvap("Droptubes") = chkCondDroptubes.CheckState
            drwEvap("CKT") = txtCondCKT.Text
            drwEvap("CKT1") = txtCondCKT1.Text
            drwEvap("CKT2") = txtCondCKT2.Text
            drwEvap("Split") = cboCondSplit.Text

            das1.Tables("Evaporator").Rows.Add(drwEvap)
            dapEvap.Update(das1, "Evaporator")

        Else
            
            MyRow("FH") = txtCondFH.Text
            MyRow("FL") = txtCondFL.Text
            MyRow("FinPI") = cboCondFinPI.Text
            MyRow("Passes") = cboCondRows.Text
            MyRow("FinThk") = cboCondFinThk.Text
            MyRow("CoilP") = cboCondCoilP.Text
            MyRow("WallThk") = cboCondWallThk.Text
            MyRow("FinMat") = cboCondFinMat.Text
            MyRow("Droptubes") = chkCondDroptubes.CheckState
            MyRow("CKT") = txtCondCKT.Text
            MyRow("CKT1") = txtCondCKT1.Text
            MyRow("CKT2") = txtCondCKT2.Text
            MyRow("Split") = cboCondSplit.Text

            dapEvap.Update(das1, "Evaporator")

        End If

    End Sub

    Private Sub AddSaveCond(ByVal NewCondmodel As String)

        Dim strEvapFH As String, strEvapFinPI As String, strEvapSplit As String
        Dim strEvapRows As String, strEvapCoilP As String, strEvapFL As String
        Dim strEvapFinThk As String, strEvapWallThk As String, strEvapCKT As String
        Dim strEvapCKT1 As String, strEvapCKT2 As String, strEvapFinMat As String, vntEvapDroptubes As Object
        
        Dim drwEvap As DataRow = das1.Tables("Evaporator").Rows.Find(NewCondmodel)

        strEvapFH = drwEvap("FH")
        strEvapFinPI = drwEvap("FinPI")
        strEvapRows = drwEvap("Passes")
        strEvapCoilP = drwEvap("CoilP")
        strEvapFL = drwEvap("FL")
        strEvapFinThk = drwEvap("FinThk")
        strEvapWallThk = drwEvap("WallThk")
        strEvapCKT = drwEvap("CKT")
        strEvapCKT1 = drwEvap("CKT1")
        strEvapCKT2 = drwEvap("CKT2")
        strEvapFinMat = drwEvap("FinMat")
        strEvapSplit = drwEvap("Split")
        vntEvapDroptubes = drwEvap("Droptubes")

        Dim MyRow As DataRow = das1.Tables("Condenser").Rows.Find(NewCondmodel)

        If MyRow Is Nothing Then
            ddCondModel.Items.Add(NewCondmodel)

            Dim drwCond As DataRow = das1.Tables("Condenser").NewRow

            drwCond("Model") = NewCondmodel
            drwCond("FH") = strEvapFH
            drwCond("FL") = strEvapFL
            drwCond("FinPI") = strEvapFinPI
            drwCond("Passes") = strEvapRows
            drwCond("FinThk") = strEvapFinThk
            drwCond("CoilP") = strEvapCoilP
            drwCond("WallThk") = strEvapWallThk
            drwCond("CKT") = strEvapCKT
            drwCond("CKT1") = strEvapCKT1
            drwCond("CKT2") = strEvapCKT2
            drwCond("FinMat") = strEvapFinMat
            drwCond("Split") = strEvapSplit
            drwCond("Droptubes") = vntEvapDroptubes

            das1.Tables("Condenser").Rows.Add(drwCond)
            dapCond.Update(das1, "Condenser")

        Else
            
            MyRow("FH") = strEvapFH
            MyRow("FL") = strEvapFL
            MyRow("FinPI") = strEvapFinPI
            MyRow("Passes") = strEvapRows
            MyRow("FinThk") = strEvapFinThk
            MyRow("CoilP") = strEvapCoilP
            MyRow("WallThk") = strEvapWallThk
            MyRow("CKT") = strEvapCKT
            MyRow("CKT1") = strEvapCKT1
            MyRow("CKT2") = strEvapCKT2
            MyRow("FinMat") = strEvapFinMat
            MyRow("Split") = strEvapSplit
            MyRow("Droptubes") = vntEvapDroptubes

            dapCond.Update(das1, "Condenser")

        End If

    End Sub

    Private Sub UpdateACUnit(ByVal NewUnitModel As String, ByVal OldEvapCFM As Double, ByVal OldCondCFM As Double, ByVal OldOutaCFM As Double, _
    ByVal OldSubcool As Double, ByVal OldSuperh As Double, ByVal OldAltitude As Double, ByVal NewCompmodel As String, _
    ByVal NewRefCode As String, ByVal NewFreCode As String, ByVal NewAppCode As String)

        Dim MyRow As DataRow = das1.Tables("Unit").Rows.Find(NewUnitModel)

        MyRow("CompressorModel") = NewCompmodel
        MyRow("CompressorRef") = NewRefCode
        MyRow("CompressorFre") = NewFreCode
        MyRow("CompressorApp") = NewAppCode
        MyRow("EvapCFM") = OldCondCFM
        MyRow("CondCFM") = OldEvapCFM
        MyRow("OACFM") = OldOutaCFM
        MyRow("Subcool") = OldSubcool
        MyRow("Superh") = OldSuperh
        MyRow("Reheat") = "None"
        MyRow("Altitude") = OldAltitude

        dapUnit.Update(das1, "Unit")

    End Sub

    Private Sub UpdateHPUnit(ByVal NewUnitModel As String, ByVal OldEvapCFM As Double, ByVal OldCondCFM As Double, ByVal OldOutaCFM As Double, _
    ByVal OldSubcool As Double, ByVal OldSuperh As Double, ByVal OldAltitude As Double, ByVal NewCompmodel As String, _
    ByVal NewRefCode As String, ByVal NewFreCode As String, ByVal NewAppCode As String)

        Dim MyRow As DataRow = das1.Tables("Unit").Rows.Find(NewUnitModel)

        MyRow("CompressorModel") = NewCompmodel
        MyRow("CompressorRef") = NewRefCode
        MyRow("CompressorFre") = NewFreCode
        MyRow("CompressorApp") = NewAppCode
        MyRow("EvapCFM") = OldCondCFM
        MyRow("CondCFM") = OldEvapCFM
        MyRow("OACFM") = OldOutaCFM
        MyRow("Subcool") = OldSubcool
        MyRow("Superh") = OldSuperh
        MyRow("Reheat") = "None"
        MyRow("Altitude") = OldAltitude

        dapUnit.Update(das1, "Unit")

    End Sub

    Private Sub cmdCompressorFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCompressorFile.Click
        Dim StrFileName As String

        If FileDialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
            StrFileName = FileDialog.FileName

            If StrFileName <> "" Then
                LoadCompFile(StrFileName)
                MsgBox("Import Complete", vbOKOnly, "Import Complete")
            End If
        End If
 
    End Sub

    Private Sub LoadCompFile(ByVal strFileName As String)

        Dim textLine As String
        Dim strParm As String
        Dim vals() As String
        Dim firstCoef As Integer
        Dim modelNo As Integer
        Dim RefcodeNo As Integer
        Dim FreCodeNo As Integer
        Dim AppCodeNo As Integer
        Dim i As Object
        Dim j As Object
        Dim K As Object

        'Screen.MousePointer = vbHourglass

        ' Open the coefficients file
        Dim objReader As StreamReader = New StreamReader(strFileName)

        ' Read in the first line which will contain the column headers
        textLine = objReader.ReadLine()

        vals = Split(textLine, ",")
        j = UBound(vals)

        For i = 0 To UBound(vals)
            ' Search for the first coefficient column
            If vals(i) = "RT_C0" Then
                firstCoef = i
            End If
            ' Search for the Model column
            If vals(i) = "Model" Then
                modelNo = i
            End If
            'Search for the Applucation no Column
            If vals(i) = "Appl" Then
                AppCodeNo = i
            End If
            ' Search for the Refrigerant type column
            If vals(i) = "Refr" Then
                RefcodeNo = i
            End If
            ' Search for the Frequency column
            If vals(i) = "Hertz" Then
                FreCodeNo = i
            End If
        Next

        Dim cmdInsertComp1 As New OleDbCommand()
        Dim CompCon As OleDbConnection = ConnectToUnitData()
        cmdInsertComp1.Connection = CompCon
        
        cmdInsertComp1.CommandText = "INSERT INTO Compressor " & _
        "(Model, RefCode, FreCode, AppCode, CF0, CF1, CF2, CF3, CF4, CF5, CF6, CF7, CF8, CF9, CF10, " & _
        " CF11, CF12, CF13, CF14, CF15, CF16, CF17, CF18, CF19, CF20, CF21, CF22, CF23, CF24, CF25, " & _
        " CF26, CF27, CF28, CF29, CF30, CF31, CF32, CF33, CF34, CF35, CF36, CF37, CF38, CF39) " & _
        "VALUES (@Model, @RefCode, @FreCode, @AppCode, @CF0, @CF1, @CF2, @CF3, @CF4, @CF5, @CF6, @CF7, @CF8, @CF9, @CF10, " & _
        "@CF11, @CF12, @CF13, @CF14, @CF15, @CF16, @CF17, @CF18, @CF19, @CF20, @CF21, @CF22, @CF23, @CF24, @CF25, " & _
        "@CF26, @CF27, @CF28, @CF29, @CF30, @CF31, @CF32, @CF33, @CF34, @CF35, @CF36, @CF37, @CF38, @CF39)"

        With cmdInsertComp1.Parameters
            .Add("@Model", OleDbType.VarChar, 50, "Model")
            .Add("@RefCode", OleDbType.Char, 12, "RefCode")
            .Add("@FreCode", OleDbType.Char, 12, "FreCode")
            .Add("@AppCode", OleDbType.Char, 12, "AppCode")
            .Add("@CF0", OleDbType.Double, 8, "CF0")
            .Add("@CF1", OleDbType.Double, 8, "CF1")
            .Add("@CF2", OleDbType.Double, 8, "CF2")
            .Add("@CF3", OleDbType.Double, 8, "CF3")
            .Add("@CF4", OleDbType.Double, 8, "CF4")
            .Add("@CF5", OleDbType.Double, 8, "CF5")
            .Add("@CF6", OleDbType.Double, 8, "CF6")
            .Add("@CF7", OleDbType.Double, 8, "CF7")
            .Add("@CF8", OleDbType.Double, 8, "CF8")
            .Add("@CF9", OleDbType.Double, 8, "CF9")
            .Add("@CF10", OleDbType.Double, 8, "CF10")
            .Add("@CF11", OleDbType.Double, 8, "CF11")
            .Add("@CF12", OleDbType.Double, 8, "CF12")
            .Add("@CF13", OleDbType.Double, 8, "CF13")
            .Add("@CF14", OleDbType.Double, 8, "CF14")
            .Add("@CF15", OleDbType.Double, 8, "CF15")
            .Add("@CF16", OleDbType.Double, 8, "CF16")
            .Add("@CF17", OleDbType.Double, 8, "CF17")
            .Add("@CF18", OleDbType.Double, 8, "CF18")
            .Add("@CF19", OleDbType.Double, 8, "CF19")
            .Add("@CF20", OleDbType.Double, 8, "CF20")
            .Add("@CF21", OleDbType.Double, 8, "CF21")
            .Add("@CF22", OleDbType.Double, 8, "CF22")
            .Add("@CF23", OleDbType.Double, 8, "CF23")
            .Add("@CF24", OleDbType.Double, 8, "CF24")
            .Add("@CF25", OleDbType.Double, 8, "CF25")
            .Add("@CF26", OleDbType.Double, 8, "CF26")
            .Add("@CF27", OleDbType.Double, 8, "CF27")
            .Add("@CF28", OleDbType.Double, 8, "CF28")
            .Add("@CF29", OleDbType.Double, 8, "CF29")
            .Add("@CF30", OleDbType.Double, 8, "CF30")
            .Add("@CF31", OleDbType.Double, 8, "CF31")
            .Add("@CF32", OleDbType.Double, 8, "CF32")
            .Add("@CF33", OleDbType.Double, 8, "CF33")
            .Add("@CF34", OleDbType.Double, 8, "CF34")
            .Add("@CF35", OleDbType.Double, 8, "CF35")
            .Add("@CF36", OleDbType.Double, 8, "CF36")
            .Add("@CF37", OleDbType.Double, 8, "CF37")
            .Add("@CF38", OleDbType.Double, 8, "CF38")
            .Add("@CF39", OleDbType.Double, 8, "CF39")
        End With

        cmdInsertComp1.Connection.Open()

        Do While objReader.Peek() >= 0

            textLine = objReader.ReadLine()

            vals = Split(textLine, ",")
            K = UBound(vals)

            Dim MyRows As DataRow() = das1.Tables("TempCompressor").Select("Model = '" & Trim(vals(modelNo)) & "'" & _
            "and RefCode = '" & Trim(vals(RefcodeNo + K - j)) & "' and FreCode = '" & Trim(vals(FreCodeNo + K - j)) & "' and AppCode = '" & Trim(vals(AppCodeNo)) & "'")

            If MyRows.Length = 0 Then

                With cmdInsertComp1
                    .Parameters("@Model").Value = Trim(vals(modelNo))
                    .Parameters("@RefCode").Value = Trim(vals(RefcodeNo + K - j))
                    .Parameters("@FreCode").Value = Trim(vals(FreCodeNo + K - j))
                    .Parameters("@AppCode").Value = Trim(vals(AppCodeNo))

                    cboCompModel.Items.Add(Trim(vals(modelNo)))
                    cboRefCode.Items.Add(Trim(vals(RefcodeNo + K - j)))
                    cboFreCode.Items.Add(Trim(vals(FreCodeNo + K - j)))
                    cboAppCode.Items.Add(Trim(vals(AppCodeNo)))

                    For i = 0 To 39
                        strParm = ("@CF" & i)
                        .Parameters(strParm).Value = Trim(vals(i + firstCoef + K - j))
                    Next

                    cboCompModel.Refresh()
                    cboRefCode.Refresh()
                    cboFreCode.Refresh()
                    cboAppCode.Refresh()
                End With

                
                cmdInsertComp1.ExecuteNonQuery()

            Else

                MsgBox("Compressor Model: " & Trim(vals(modelNo)) & " w/Refrigerant, Frequency & Application already exists", vbOKOnly, "Model Already Exists")

            End If

        Loop

        cmdInsertComp1.Connection.Close()

        dapTempComp.Fill(das1, "TempCompressor")

        Dim drwComp As DataRow

        For Each drwComp In das1.Tables("TempCompressor").Rows
            cboCompModel.Items.Add(drwComp("Model"))
        Next

        cboCompModel.SelectedIndex = 0

        'Screen.MousePointer = vbDefault

    End Sub

    Private Function EnhancedSplit(ByVal stringToSplit As String, ByVal delimiters() As Char) As String()
        Dim Words() As String

        Words = stringToSplit.Split(delimiters)
        Dim FilterWords As New ArrayList()
        Dim Word As String
        For Each Word In Words
            If Word <> String.Empty Then
                FilterWords.Add(Word)
            End If
        Next

        Return CType(FilterWords.ToArray(GetType(String)), String())

    End Function

    Private Sub txtEvapFH_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) _
                                Handles txtEvapFH.KeyPress
        If Char.IsLetter(e.KeyChar) = True Then
            e.Handled = True
            MsgBox("Please no letters in this field", vbOKOnly, "Evap Fin height must be numeric")
        End If
    End Sub

    Private Sub txtEvapFL_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) _
                                Handles txtEvapFL.KeyPress
        If Char.IsLetter(e.KeyChar) = True Then
            e.Handled = True
            MsgBox("Please no letters in this field", vbOKOnly, "Evap Fin length must be numeric")
        End If
    End Sub

    Private Sub txtEvapCKT1_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) _
                                Handles txtEvapCKT1.KeyPress
        If Not Char.IsDigit(e.KeyChar) And Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
            MsgBox("Please no decimals or letters in this field", vbOKOnly, "Evap circuit 1 feeds must be an integer")
        End If
    End Sub

    Private Sub txtEvapCKT2_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) _
                                Handles txtEvapCKT2.KeyPress
        If Not Char.IsDigit(e.KeyChar) And Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
            MsgBox("Please no decimals or letters in this field", vbOKOnly, "Evap circuit 2 feeds must be an integer")
        End If
    End Sub

    Private Sub txtEvapCKT_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) _
                                Handles txtEvapCKT.KeyPress
        If Not Char.IsDigit(e.KeyChar) And Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
            MsgBox("Please no decimals or letters in this field", vbOKOnly, "Evap total feeds must be an integer")
        End If
    End Sub

    Private Sub txtCondFH_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) _
                                Handles txtCondFH.KeyPress
        If Char.IsLetter(e.KeyChar) = True Then
            e.Handled = True
            MsgBox("Please no letters in this field", vbOKOnly, "Cond Fin height must be numeric")
        End If
    End Sub

    Private Sub txtCondFL_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) _
                                Handles txtCondFL.KeyPress
        If Char.IsLetter(e.KeyChar) = True Then
            e.Handled = True
            MsgBox("Please no letters in this field", vbOKOnly, "Cond Fin length must be numeric")
        End If
    End Sub

    Private Sub txtCondCKT1_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) _
                                Handles txtCondCKT1.KeyPress
        If Not Char.IsDigit(e.KeyChar) And Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
            MsgBox("Please no decimals or letters in this field", vbOKOnly, "Cond circuit 1 feeds must be an integer")
        End If
    End Sub

    Private Sub txtCondCKT2_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) _
                                Handles txtCondCKT2.KeyPress
        If Not Char.IsDigit(e.KeyChar) And Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
            MsgBox("Please no decimals or letters in this field", vbOKOnly, "Cond circuit 2 feeds must be an integer")
        End If
    End Sub

    Private Sub txtCondCKT_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) _
                                Handles txtCondCKT.KeyPress
        If Not Char.IsDigit(e.KeyChar) And Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
            MsgBox("Please no decimals or letters in this field", vbOKOnly, "Cond total feeds must be an integer")
        End If
    End Sub

    Private Sub txtEvEaDB_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) _
                                Handles txtEvEaDB.KeyPress
        If Char.IsLetter(e.KeyChar) = True Then
            e.Handled = True
            MsgBox("Please no letters in this field", vbOKOnly, "Evap entering air DB must be numeric")
        End If
    End Sub

    Private Sub txtEvEaWB_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) _
                                Handles txtEvEaWB.KeyPress
        If Char.IsLetter(e.KeyChar) = True Then
            e.Handled = True
            MsgBox("Please no letters in this field", vbOKOnly, "Evap entering air WB must be numeric")
        End If
    End Sub

    Private Sub txtEvCFM_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) _
                                Handles txtEvCFM.KeyPress
        If Char.IsLetter(e.KeyChar) = True Then
            e.Handled = True
            MsgBox("Please no letters in this field", vbOKOnly, "Evap CFM must be numeric")
        End If
    End Sub

    Private Sub txtCoEaDB_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) _
                                Handles txtCoEaDB.KeyPress
        If Char.IsLetter(e.KeyChar) = True Then
            e.Handled = True
            MsgBox("Please no letters in this field", vbOKOnly, "Cond entering air DB must be numeric")
        End If
    End Sub

    Private Sub txtCoEaWB_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) _
                                Handles txtCoEaWB.KeyPress
        If Char.IsLetter(e.KeyChar) = True Then
            e.Handled = True
            MsgBox("Please no letters in this field", vbOKOnly, "Cond entering air WB must be numeric")
        End If
    End Sub

    Private Sub txtCoCFM_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) _
                                Handles txtCoCFM.KeyPress
        If Char.IsLetter(e.KeyChar) = True Then
            e.Handled = True
            MsgBox("Please no letters in this field", vbOKOnly, "Cond CFM must be numeric")
        End If
    End Sub

    Private Sub txtOutaCFM_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) _
                                Handles txtOutaCFM.KeyPress
        If Char.IsLetter(e.KeyChar) = True Then
            e.Handled = True
            MsgBox("Please no letters in this field", vbOKOnly, "Outside air CFM must be numeric")
        End If
    End Sub

    Private Sub txtAltitude_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) _
                                Handles txtAltitude.KeyPress
        If Char.IsLetter(e.KeyChar) = True Then
            e.Handled = True
            MsgBox("Please no letters in this field", vbOKOnly, "Altitude must be numeric")
        End If
    End Sub

    Private Sub txtReheatFL_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) _
                                Handles txtReheatFL.KeyPress
        If Char.IsLetter(e.KeyChar) = True Then
            e.Handled = True
            MsgBox("Please no letters in this field", vbOKOnly, "Reheat coil fin length must be numeric")
        End If
    End Sub

    Private Sub txtReheatFH_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) _
                                Handles txtReheatFH.KeyPress
        If Char.IsLetter(e.KeyChar) = True Then
            e.Handled = True
            MsgBox("Please no letters in this field", vbOKOnly, "Reheat coil fin heigth must be numeric")
        End If
    End Sub

    Private Sub txtReheatLaDB_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) _
                                Handles txtReheatLaDB.KeyPress
        If Char.IsLetter(e.KeyChar) = True Then
            e.Handled = True
            MsgBox("Please no letters in this field", vbOKOnly, "Reheat coil leaving air DB must be numeric")
        End If
    End Sub

    Private Sub txtRehaetCKT_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) _
                                Handles txtReheatCKT.KeyPress
        If Not Char.IsDigit(e.KeyChar) And Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
            MsgBox("Please no decimals or letters in this field", vbOKOnly, "Reheat coil feeds must be an integer")
        End If
    End Sub

    Private Sub mnuFileHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileHelp.Click
        System.Diagnostics.Process.Start(AppPath() & "\UnitHelp.doc")
    End Sub
End Class
