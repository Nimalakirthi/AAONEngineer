Imports System.Math

Module modMain

    Public mEvaPClass As pjEvap.EvapClass
    Public mConPClass As pjCond.CondClass
    Public mComPClass As pjComp.CompClass
    Public mRecPClass As pjReheat.ReheatClass
    'Public mRefClass As pjRefProp.RefProp

    Public struEvap As pjEvap.Evap
    Public struCond As pjCond.Cond
    Public struComp As pjComp.Comp
    Public struReheat As pjReheat.Reheat
    Public mSysP As Syst

    Public dsItems As New DataSet()
    Public dtItems As New DataTable("Items")

    Structure Syst

        Public n_EvCFM As Double           'Entering Air CFM
        Public n_OaCFM As Double           'Out side air CFM
        Public n_CoCFM As Double           'Condenser CFM
        Public n_Refrigerant As String           'Refrigerant
        Public n_EvEaDB As Double           'Entering air Dry Bulb
        Public n_EvEaWB As Double           'Entering Air wet bulb
        Public n_CoEaDB As Double           'Condenser Entering Air Dry Bulb
        Public n_CoEaWB As Double           'Condenser Entering wet Bulb
        Public n_Altitude As Double
        Public n_EvACFM As Double
        Public n_CoACFM As Double
        Public n_OaACFM As Double
        Public n_ReheatYes_No As String
        Public i_CondQuantity As Integer   'Condenser Quantity
        Public i_EvapQuantity As Integer
        Public i_CompQuantity As Integer   'Compressor Quantity
        Public n_MixedDB As Double
        Public n_MixedWB As Double
        Public n_Subcool As Double
        Public n_Superh As Double
        Public s_UnitModel As String
        Public s_CompAppCode As String
        Public s_Hit As String
        Public n_ETMax As Double
        Public n_CTMax As Double

    End Structure

    Private Function StartCalculations(ByVal Compquantity As Integer) As Boolean

        StartCalculations = False

        Call mEvaPClass.StartEvap(Compquantity)

        Call mConPClass.StartCond(Compquantity)

        StartCalculations = True

        Exit Function
    End Function

    Public Sub StartSys_Balance_AC()
        Dim mn_ET As Double
        Dim mn_CT As Double
        Dim mn_MassFlow As Double
        Dim mn_CompBtuh As Double
        Dim MassFraction As Double
        Dim n_CompWattsBtuh As Double
        Dim b_EvapHit As Boolean
        Dim b_CondHit As Boolean
        Dim b_ReheatHit As Boolean
        Dim i_DifficulityCounter As Integer
        Dim i_LoopCounter As Integer          'number of times thru loop
        Dim s_Hit As String
        Dim n_TolerancePercent As Double           'tolerance (input)
        Dim n_Percent1 As Double           'evap-comp percentage
        Dim n_Percent2 As Double           'cond-comp+watts percentage
        Dim n_EvapStep As Double           'increment evap temp by (Input by user)
        Dim n_CondStep As Double           'increment evap temp by (Input by user)
        Dim ReheatYes_No As Boolean
        Dim HeatRej As Double
        Dim Subcool As Double
        Dim Superh As Double
        Dim ReheatDelta As Double
        Dim EvACFM As Double
        Dim EvEaDB As Double
        Dim EvEaWB As Double
        Dim CoACFM As Double
        Dim EvLaDB As Double
        Dim Compquantity As Integer

        Dim nActDensity As Double
        Dim n_EWBlblbSR As Double
        Dim n_IntV1 As Double
        Dim n_ElblbR As Double
        Dim mn_PSIA As Double
        Dim nStdDensity As Double
        Dim mn_EatWB As Double
        Dim mn_EatDB As Double
        Dim Frej As Double
        Const mn_a1 As Double = 0.00080246
        Const mn_a2 As Double = 0.000024525
        Const mn_a3 As Double = 2.542 * (10 ^ -6)
        Const mn_a4 As Double = -2.5855 * (10 ^ -8)
        Const mn_a5 As Double = 4.038 * (10 ^ -10)

        Compquantity = mSysP.i_CompQuantity
        mn_PSIA = maPSIA(mSysP.n_Altitude)
        Frej = 1.33
        ReheatDelta = 0.04

        If mSysP.n_Altitude > 1999 Then

            mn_EatWB = mSysP.n_EvEaWB
            mn_EatDB = mSysP.n_EvEaDB
            n_EWBlblbSR = mn_a1 + (mn_a2 * mn_EatWB) + (mn_a3 * mn_EatWB ^ 2) + (mn_a4 * mn_EatWB ^ 3) + (mn_a5 * mn_EatWB ^ 4)
            n_IntV1 = ((1093 - 0.556 * mn_EatWB) * n_EWBlblbSR) - (0.24 * (mn_EatDB - mn_EatWB))
            n_ElblbR = n_IntV1 / (1093 + 0.444 * mn_EatDB - mn_EatWB)
            nActDensity = maDensity(mn_EatDB, n_ElblbR, mn_PSIA)
            nStdDensity = maDensity(70, 0, 14.696)
            mSysP.n_EvACFM = mSysP.n_EvCFM / (nActDensity / nStdDensity)

            mn_EatWB = mSysP.n_CoEaWB
            mn_EatDB = mSysP.n_CoEaDB
            n_EWBlblbSR = mn_a1 + (mn_a2 * mn_EatWB) + (mn_a3 * mn_EatWB ^ 2) + (mn_a4 * mn_EatWB ^ 3) + (mn_a5 * mn_EatWB ^ 4)
            n_IntV1 = ((1093 - 0.556 * mn_EatWB) * n_EWBlblbSR) - (0.24 * (mn_EatDB - mn_EatWB))
            n_ElblbR = n_IntV1 / (1093 + 0.444 * mn_EatDB - mn_EatWB)
            nActDensity = maDensity(mn_EatDB, n_ElblbR, mn_PSIA)
            nStdDensity = maDensity(70, 0, 14.696)
            mSysP.n_OaACFM = mSysP.n_OaCFM / (nActDensity / nStdDensity)
            mSysP.n_CoACFM = mSysP.n_CoCFM / (nActDensity / nStdDensity)

        Else

            mSysP.n_EvACFM = mSysP.n_EvCFM
            mSysP.n_OaACFM = mSysP.n_OaCFM
            mSysP.n_CoACFM = mSysP.n_CoCFM

        End If

        If mSysP.n_OaACFM = 0 Then
            EvEaDB = mSysP.n_EvEaDB
            EvEaWB = mSysP.n_EvEaWB
            EvACFM = mSysP.n_EvACFM
        Else
            EvEaDB = mSysP.n_EvEaDB
            EvEaWB = mSysP.n_EvEaWB
            EvACFM = mSysP.n_EvACFM + mSysP.n_OaACFM

            Call MixedAir_AC(EvEaDB, EvEaWB)
            EvEaDB = EvEaDB
            EvEaWB = EvEaWB
        End If

        mSysP.n_MixedDB = EvEaDB
        mSysP.n_MixedWB = EvEaWB
        CoACFM = mSysP.n_CoACFM

        If StartCalculations(Compquantity) Then
            mn_ET = 55
            mn_CT = mSysP.n_CoEaDB + 25

            i_LoopCounter = 0
            s_Hit = ""
            n_TolerancePercent = 0.01

            Subcool = mSysP.n_Subcool
            Superh = mSysP.n_Superh

            Call mEvaPClass.PassMixAirInfo(EvEaDB, EvEaWB, EvACFM, mSysP.s_CompAppCode)
            Call mConPClass.PassAirCorrection(mSysP.n_CoEaDB, mSysP.n_CoEaWB, CoACFM, mSysP.s_CompAppCode)
            Call mRecPClass.PassMixAirCFM(EvACFM)

            If mSysP.n_ReheatYes_No <> "None" Then
                ReheatYes_No = True
            Else
                ReheatYes_No = False
            End If

            Do

                Call mComPClass.CompCalc(mn_ET, mn_CT)
                mn_MassFlow = mComPClass.CompProperties.n_MassFlow
                mn_CompBtuh = mComPClass.CompProperties.n_BTUH
                Call mEvaPClass.EvapCalc(mn_CompBtuh, mn_MassFlow, mn_ET, mn_CT)

                EvLaDB = mEvaPClass.EvapProperties.n_LaDB

                If ReheatYes_No = True Then
                    Call mRecPClass.ReNT(mn_MassFlow, mn_CT, EvLaDB)
                    MassFraction = mRecPClass.ReheatProperties.n_MassFraction
                Else
                    MassFraction = 0
                End If

                Call mConPClass.CondCalc(mn_MassFlow, MassFraction, mn_ET, mn_CT)

                n_CompWattsBtuh = mEvaPClass.EvapProperties.n_BTUH + (mComPClass.CompProperties.n_WattsBTUH * 3.415)

                If mEvaPClass.EvapProperties.n_BTUH >= mComPClass.CompProperties.n_BTUH Then
                    n_Percent1 = Abs(mEvaPClass.EvapProperties.n_BTUH - mComPClass.CompProperties.n_BTUH) / mEvaPClass.EvapProperties.n_BTUH
                Else
                    n_Percent1 = Abs(mEvaPClass.EvapProperties.n_BTUH - mComPClass.CompProperties.n_BTUH) / mComPClass.CompProperties.n_BTUH
                End If

                If (mConPClass.CondProperties.n_BTUH + mRecPClass.ReheatProperties.n_BTUH) >= n_CompWattsBtuh Then
                    n_Percent2 = Abs(mConPClass.CondProperties.n_BTUH + mRecPClass.ReheatProperties.n_BTUH - n_CompWattsBtuh) / (mConPClass.CondProperties.n_BTUH + mRecPClass.ReheatProperties.n_BTUH)
                Else
                    n_Percent2 = Abs(mConPClass.CondProperties.n_BTUH + mRecPClass.ReheatProperties.n_BTUH - n_CompWattsBtuh) / (n_CompWattsBtuh)
                End If

                If mn_ET = 55 Then
                    Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                End If

                i_LoopCounter = i_LoopCounter + 1
                If n_Percent1 <= n_TolerancePercent Then
                    b_EvapHit = True
                ElseIf mEvaPClass.EvapProperties.n_BTUH < mComPClass.CompProperties.n_BTUH Then
                    n_EvapStep = (n_Percent1 / 0.07)
                    mn_ET = mn_ET - n_EvapStep
                    If mn_ET < 10 Then
                        mn_ET = 10
                    End If
                    b_EvapHit = False
                Else
                    n_EvapStep = (n_Percent1 / 0.07)
                    mn_ET = mn_ET + n_EvapStep
                    b_EvapHit = False
                End If

                If n_Percent2 <= (n_TolerancePercent + 0.01) Then
                    b_CondHit = True
                ElseIf (mConPClass.CondProperties.n_BTUH + mRecPClass.ReheatProperties.n_BTUH) < n_CompWattsBtuh Then
                    n_CondStep = (n_Percent2 / 0.06)
                    mn_CT = mn_CT + n_CondStep
                    If mn_CT > 155 Then
                        mn_CT = 155
                    End If
                    b_CondHit = False
                Else
                    n_CondStep = (n_Percent2 / 0.06)
                    mn_CT = mn_CT - n_CondStep
                    If mn_CT < 75 Then
                        mn_CT = 75
                    End If
                    b_CondHit = False
                End If

                If mSysP.n_ReheatYes_No <> "None" Then

                    If mRecPClass.ReheatProperties.n_DesiredBTUH <= 0 Then
                        MsgBox("Reheat temperature must be greater than LADB", vbOKOnly, "Check Reheat temperature")
                        s_Hit = " Reheat temperature must be greater than LADB"
                        Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                        Exit Do
                    End If

                    If ((mRecPClass.ReheatProperties.n_BTUH - mRecPClass.ReheatProperties.n_DesiredBTUH) / mRecPClass.ReheatProperties.n_DesiredBTUH) > ReheatDelta Then
                        MassFraction = MassFraction - 0.02
                        b_ReheatHit = False
                    ElseIf ((mRecPClass.ReheatProperties.n_BTUH - mRecPClass.ReheatProperties.n_DesiredBTUH) / mRecPClass.ReheatProperties.n_DesiredBTUH) < (-ReheatDelta) Then
                        MassFraction = MassFraction + 0.02
                        b_ReheatHit = False
                    Else
                        Call mRecPClass.VariationMassFraction(MassFraction)
                        b_ReheatHit = True
                    End If

                    If MassFraction <= 0 Or MassFraction >= 1 Then
                        If MassFraction <= 0 Then
                            MassFraction = 0.005
                            MsgBox("Difference between Reheat temperature and evap leaving temperature must be > 10 degrees", vbOKOnly, "Check Reheat temperature")
                            s_Hit = " (Reheat temperature - LADB) > 10? If Yes Change Fin/Inch, If No Correct it"
                            Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                            Exit Do
                        Else
                            MassFraction = 0.99
                            MsgBox("Difference between Reheat temperature and evap leaving temperature must be < 20 degrees", vbOKOnly, "Check Reheat temperature")
                            s_Hit = " (Reheat temperature - LADB) < 20? If Yes Change Fin/Inch, If No Correct it"
                            Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                            Exit Do
                        End If
                    End If

                    Call mRecPClass.VariationMassFraction(MassFraction)

                    HeatRej = (mConPClass.CondProperties.n_BTUH + mRecPClass.ReheatProperties.n_BTUH) / mEvaPClass.EvapProperties.n_BTUH

                    Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)

                    If b_CondHit And b_EvapHit And b_ReheatHit And (mn_ET < mSysP.n_ETMax) And (mn_CT < mSysP.n_CTMax) And (HeatRej < Frej) Then

                        If (mn_ET < 52.5) And (mn_CT < 130) Then
                            s_Hit = " Balance Unit!"
                        Else
                            s_Hit = " Warning! System Balance at high evap/cond Temperature"
                        End If

                        Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                        Exit Do
                    End If

                Else
                    Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                    HeatRej = mConPClass.CondProperties.n_BTUH / mEvaPClass.EvapProperties.n_BTUH

                    If b_CondHit And b_EvapHit And (mn_ET < mSysP.n_ETMax) And (mn_CT < mSysP.n_CTMax) And (HeatRej < Frej) Then

                        If (mn_ET < 52.5) And (mn_CT < 130) Then
                            s_Hit = " Balance Unit!"
                        Else
                            s_Hit = " Warning! System Balance at high evap/cond Temperature"
                        End If

                        Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                        Exit Do
                    End If
                End If

                If i_LoopCounter > 33 Then
                    i_DifficulityCounter = i_DifficulityCounter + 1
                    ReheatDelta = 0.05
                    If i_DifficulityCounter = 5 Then
                        s_Hit = " Could Not Balance!"
                        ReheatDelta = 0.04
                        Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                        Exit Do
                    End If

                End If

            Loop

        End If
    End Sub

    Public Sub StartSys_Balance_HP()
        Dim mn_ET As Double
        Dim mn_CT As Double
        Dim mn_MassFlow As Double
        Dim mn_CompBtuh As Double
        Dim MassFraction As Double
        Dim n_CompWattsBtuh As Double
        Dim b_EvapHit As Boolean
        Dim b_CondHit As Boolean
        Dim i_DifficulityCounter As Integer
        Dim i_LoopCounter As Integer          'number of times thru loop
        Dim s_Hit As String
        Dim n_TolerancePercent As Double           'tolerance (input)
        Dim n_Percent1 As Double           'evap-comp percentage
        Dim n_Percent2 As Double           'cond-comp+watts percentage
        Dim n_EvapStep As Double           'increment evap temp by (Input by user)
        Dim n_CondStep As Double           'increment evap temp by (Input by user)
        Dim HeatRej As Double
        Dim Subcool As Double
        Dim Superh As Double

        Dim EvACFM As Double
        Dim CoEaDB As Double
        Dim CoEaWB As Double
        Dim CoACFM As Double
        Dim Compquantity As Integer
        Dim EvStepDiv As Double
        Dim CoStepDiv As Double
        Dim nActDensity As Double
        Dim n_EWBlblbSR As Double
        Dim n_IntV1 As Double
        Dim n_ElblbR As Double
        Dim mn_PSIA As Double
        Dim nStdDensity As Double
        Dim mn_EatWB As Double
        Dim mn_EatDB As Double
        Dim Frej As Double
        Const mn_a1 As Double = 0.00080246
        Const mn_a2 As Double = 0.000024525
        Const mn_a3 As Double = 2.542 * (10 ^ -6)
        Const mn_a4 As Double = -2.5855 * (10 ^ -8)
        Const mn_a5 As Double = 4.038 * (10 ^ -10)

        Compquantity = mSysP.i_CompQuantity
        mn_PSIA = maPSIA(mSysP.n_Altitude)
        Frej = 1.43

        If mSysP.n_Altitude > 1999 Then

            mn_EatWB = mSysP.n_EvEaWB
            mn_EatDB = mSysP.n_EvEaDB
            n_EWBlblbSR = mn_a1 + (mn_a2 * mn_EatWB) + (mn_a3 * mn_EatWB ^ 2) + (mn_a4 * mn_EatWB ^ 3) + (mn_a5 * mn_EatWB ^ 4)
            n_IntV1 = ((1093 - 0.556 * mn_EatWB) * n_EWBlblbSR) - (0.24 * (mn_EatDB - mn_EatWB))
            n_ElblbR = n_IntV1 / (1093 + 0.444 * mn_EatDB - mn_EatWB)
            nActDensity = maDensity(mn_EatDB, n_ElblbR, mn_PSIA)
            nStdDensity = maDensity(70, 0, 14.696)
            mSysP.n_EvACFM = mSysP.n_EvCFM / (nActDensity / nStdDensity)
            mSysP.n_OaACFM = mSysP.n_OaCFM / (nActDensity / nStdDensity)

            mn_EatWB = mSysP.n_CoEaWB
            mn_EatDB = mSysP.n_CoEaDB
            n_EWBlblbSR = mn_a1 + (mn_a2 * mn_EatWB) + (mn_a3 * mn_EatWB ^ 2) + (mn_a4 * mn_EatWB ^ 3) + (mn_a5 * mn_EatWB ^ 4)
            n_IntV1 = ((1093 - 0.556 * mn_EatWB) * n_EWBlblbSR) - (0.24 * (mn_EatDB - mn_EatWB))
            n_ElblbR = n_IntV1 / (1093 + 0.444 * mn_EatDB - mn_EatWB)
            nActDensity = maDensity(mn_EatDB, n_ElblbR, mn_PSIA)
            nStdDensity = maDensity(70, 0, 14.696)
            mSysP.n_CoACFM = mSysP.n_CoCFM / (nActDensity / nStdDensity)

        Else

            mSysP.n_EvACFM = mSysP.n_EvCFM
            mSysP.n_OaACFM = mSysP.n_OaCFM
            mSysP.n_CoACFM = mSysP.n_CoCFM

        End If

        If mSysP.n_OaACFM = 0 Then
            CoEaDB = mSysP.n_CoEaDB
            CoEaWB = mSysP.n_CoEaWB
            CoACFM = mSysP.n_CoACFM
        Else
            CoEaDB = mSysP.n_CoEaDB
            CoEaWB = mSysP.n_CoEaWB
            CoACFM = mSysP.n_CoACFM + mSysP.n_OaACFM

            Call MixedAir_HP(CoEaDB, CoEaWB)
            CoEaDB = CoEaDB
            CoEaWB = CoEaWB
        End If

        mSysP.n_MixedDB = CoEaDB
        mSysP.n_MixedWB = CoEaWB
        EvACFM = mSysP.n_EvACFM

        If mSysP.n_EvEaDB >= 40 Then
            EvStepDiv = 0.07 * 3
        ElseIf mSysP.n_EvEaDB >= 30 Then
            EvStepDiv = 0.07 * 3.5
        ElseIf mSysP.n_EvEaDB >= 15 Then
            EvStepDiv = 0.07 * 4
        Else
            EvStepDiv = 0.07 * 5
        End If

        If mSysP.n_EvEaDB >= 30 Then
            CoStepDiv = 0.06
        ElseIf mSysP.n_EvEaDB >= 15 Then
            CoStepDiv = 0.06 * 2
        Else
            CoStepDiv = 0.06 * 3
        End If

        If StartCalculations(Compquantity) Then
            mn_ET = 35
            mn_CT = mSysP.n_MixedDB + 25

            i_LoopCounter = 0
            s_Hit = ""
            n_TolerancePercent = 0.01

            Subcool = mSysP.n_Subcool
            Superh = mSysP.n_Superh

            Call mEvaPClass.PassAirCorrection(mSysP.n_EvEaDB, mSysP.n_EvEaWB, EvACFM, mSysP.s_CompAppCode)
            Call mConPClass.PassMixAirInfo(CoEaDB, CoEaWB, CoACFM, mSysP.s_CompAppCode)

            Do

                Call mComPClass.CompCalc(mn_ET, mn_CT)
                mn_MassFlow = mComPClass.CompProperties.n_MassFlow
                mn_CompBtuh = mComPClass.CompProperties.n_BTUH
                Call mEvaPClass.EvapCalc(mn_CompBtuh, mn_MassFlow, mn_ET, mn_CT)

                MassFraction = 0

                Call mConPClass.CondCalc(mn_MassFlow, MassFraction, mn_ET, mn_CT)

                n_CompWattsBtuh = mEvaPClass.EvapProperties.n_BTUH + (mComPClass.CompProperties.n_WattsBTUH * 3.415)

                If mEvaPClass.EvapProperties.n_BTUH >= mComPClass.CompProperties.n_BTUH Then
                    n_Percent1 = Abs(mEvaPClass.EvapProperties.n_BTUH - mComPClass.CompProperties.n_BTUH) / mEvaPClass.EvapProperties.n_BTUH
                Else
                    n_Percent1 = Abs(mEvaPClass.EvapProperties.n_BTUH - mComPClass.CompProperties.n_BTUH) / mComPClass.CompProperties.n_BTUH
                End If

                If (mConPClass.CondProperties.n_BTUH) >= n_CompWattsBtuh Then
                    n_Percent2 = Abs(mConPClass.CondProperties.n_BTUH - n_CompWattsBtuh) / (mConPClass.CondProperties.n_BTUH)
                Else
                    n_Percent2 = Abs(mConPClass.CondProperties.n_BTUH - n_CompWattsBtuh) / (n_CompWattsBtuh)
                End If

                If mn_ET = 35 Then
                    Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                End If

                i_LoopCounter = i_LoopCounter + 1
                If n_Percent1 <= n_TolerancePercent Then
                    b_EvapHit = True
                ElseIf mEvaPClass.EvapProperties.n_BTUH < mComPClass.CompProperties.n_BTUH Then
                    n_EvapStep = (n_Percent1 / EvStepDiv)
                    mn_ET = mn_ET - n_EvapStep
                    If mn_ET < 2 Then
                        mn_ET = 2
                    End If
                    b_EvapHit = False
                Else
                    n_EvapStep = (n_Percent1 / EvStepDiv)
                    mn_ET = mn_ET + n_EvapStep
                    b_EvapHit = False
                End If

                If n_Percent2 <= (n_TolerancePercent + 0.01) Then
                    b_CondHit = True
                ElseIf (mConPClass.CondProperties.n_BTUH) < n_CompWattsBtuh Then
                    n_CondStep = (n_Percent2 / CoStepDiv)
                    mn_CT = mn_CT + n_CondStep
                    If mn_CT > 155 Then
                        mn_CT = 155
                    End If
                    b_CondHit = False
                Else
                    n_CondStep = (n_Percent2 / CoStepDiv)
                    mn_CT = mn_CT - n_CondStep
                    If mn_CT < 75 Then
                        mn_CT = 75
                    End If
                    b_CondHit = False
                End If

                Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                HeatRej = mConPClass.CondProperties.n_BTUH / mEvaPClass.EvapProperties.n_BTUH

                If b_CondHit And b_EvapHit And (mn_ET < mSysP.n_ETMax) And (mn_CT < mSysP.n_CTMax) And (HeatRej < Frej) Then

                    If (mn_ET > 30) And (mn_ET < 50) And (mn_CT < 130) Then
                        s_Hit = " Balance Unit!"
                    ElseIf (mn_ET > 30) Then
                        s_Hit = " Warning! System Balance at high evap or cond Temperature"
                    ElseIf (mn_ET < 30) Then
                        s_Hit = " Warning! System will require longer defrosting time"
                    End If

                    Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                    Exit Do
                End If

                If i_LoopCounter > 33 Then
                    i_DifficulityCounter = i_DifficulityCounter + 1
                    If i_DifficulityCounter = 5 Then
                        s_Hit = " Could Not Balance!"
                        Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                        Exit Do
                    End If

                End If

            Loop

        End If
    End Sub

    Public Sub StartSys_Balance_2Cond_AC()
        Dim mn_ET As Double
        Dim mn_CT As Double
        Dim mn_MassFlow As Double
        Dim mn_CompBtuh As Double
        Dim MassFraction As Double
        Dim n_CompWattsBtuh As Double
        Dim b_EvapHit As Boolean
        Dim b_CondHit As Boolean
        Dim b_ReheatHit As Boolean
        Dim i_DifficulityCounter As Integer
        Dim i_LoopCounter As Integer          'number of times thru loop
        Dim s_Hit As String
        Dim n_TolerancePercent As Double           'tolerance (input)
        Dim n_Percent1 As Double           'evap-comp percentage
        Dim n_Percent2 As Double           'cond-comp+watts percentage
        Dim n_EvapStep As Double           'increment evap temp by (Input by user)
        Dim n_CondStep As Double           'increment evap temp by (Input by user)
        Dim ReheatYes_No As Boolean
        Dim HeatRej As Double
        Dim Subcool As Double
        Dim Superh As Double
        Dim ReheatDelta As Double
        Dim EvACFM As Double
        Dim EvEaDB As Double
        Dim EvEaWB As Double
        Dim CoACFM As Double
        Dim EvLaDB As Double
        Dim Compquantity As Integer

        Dim nActDensity As Double
        Dim n_EWBlblbSR As Double
        Dim n_IntV1 As Double
        Dim n_ElblbR As Double
        Dim mn_PSIA As Double
        Dim nStdDensity As Double
        Dim mn_EatWB As Double
        Dim mn_EatDB As Double
        Dim Frej As Double
        Const mn_a1 As Double = 0.00080246
        Const mn_a2 As Double = 0.000024525
        Const mn_a3 As Double = 2.542 * (10 ^ -6)
        Const mn_a4 As Double = -2.5855 * (10 ^ -8)
        Const mn_a5 As Double = 4.038 * (10 ^ -10)

        Compquantity = mSysP.i_CompQuantity
        mn_PSIA = maPSIA(mSysP.n_Altitude)
        Frej = 1.33
        ReheatDelta = 0.04

        If mSysP.n_Altitude > 1999 Then

            mn_EatWB = mSysP.n_EvEaWB
            mn_EatDB = mSysP.n_EvEaDB
            n_EWBlblbSR = mn_a1 + (mn_a2 * mn_EatWB) + (mn_a3 * mn_EatWB ^ 2) + (mn_a4 * mn_EatWB ^ 3) + (mn_a5 * mn_EatWB ^ 4)
            n_IntV1 = ((1093 - 0.556 * mn_EatWB) * n_EWBlblbSR) - (0.24 * (mn_EatDB - mn_EatWB))
            n_ElblbR = n_IntV1 / (1093 + 0.444 * mn_EatDB - mn_EatWB)
            nActDensity = maDensity(mn_EatDB, n_ElblbR, mn_PSIA)
            nStdDensity = maDensity(70, 0, 14.696)
            mSysP.n_EvACFM = mSysP.n_EvCFM / (nActDensity / nStdDensity)

            mn_EatWB = mSysP.n_CoEaWB
            mn_EatDB = mSysP.n_CoEaDB
            n_EWBlblbSR = mn_a1 + (mn_a2 * mn_EatWB) + (mn_a3 * mn_EatWB ^ 2) + (mn_a4 * mn_EatWB ^ 3) + (mn_a5 * mn_EatWB ^ 4)
            n_IntV1 = ((1093 - 0.556 * mn_EatWB) * n_EWBlblbSR) - (0.24 * (mn_EatDB - mn_EatWB))
            n_ElblbR = n_IntV1 / (1093 + 0.444 * mn_EatDB - mn_EatWB)
            nActDensity = maDensity(mn_EatDB, n_ElblbR, mn_PSIA)
            nStdDensity = maDensity(70, 0, 14.696)
            mSysP.n_OaACFM = mSysP.n_OaCFM / (nActDensity / nStdDensity)
            mSysP.n_CoACFM = mSysP.n_CoCFM / (nActDensity / nStdDensity)

        Else
            mSysP.n_EvACFM = mSysP.n_EvCFM
            mSysP.n_OaACFM = mSysP.n_OaCFM
            mSysP.n_CoACFM = mSysP.n_CoCFM
        End If

        If mSysP.n_OaACFM = 0 Then
            EvEaDB = mSysP.n_EvEaDB
            EvEaWB = mSysP.n_EvEaWB
            EvACFM = mSysP.n_EvACFM
        Else
            EvEaDB = mSysP.n_EvEaDB
            EvEaWB = mSysP.n_EvEaWB
            EvACFM = mSysP.n_EvACFM + mSysP.n_OaACFM

            Call MixedAir_AC(EvEaDB, EvEaWB)
            EvEaDB = EvEaDB
            EvEaWB = EvEaWB
        End If

        mSysP.n_MixedDB = EvEaDB
        mSysP.n_MixedWB = EvEaWB
        CoACFM = mSysP.n_CoACFM

        If StartCalculations(Compquantity) Then
            mn_ET = 55
            mn_CT = mSysP.n_CoEaDB + 25

            i_LoopCounter = 0
            s_Hit = ""
            n_TolerancePercent = 0.01


            Subcool = mSysP.n_Subcool
            Superh = mSysP.n_Superh

            Call mEvaPClass.PassMixAirInfo(EvEaDB, EvEaWB, EvACFM, mSysP.s_CompAppCode)
            Call mConPClass.PassAirCorrection(mSysP.n_CoEaDB, mSysP.n_CoEaWB, CoACFM, mSysP.s_CompAppCode)
            Call mRecPClass.PassMixAirCFM(EvACFM)

            If mSysP.n_ReheatYes_No <> "None" Then
                ReheatYes_No = True
            Else
                ReheatYes_No = False
            End If

            Do

                Call mComPClass.CompCalc(mn_ET, mn_CT)
                mn_MassFlow = mComPClass.CompProperties.n_MassFlow
                mn_CompBtuh = mComPClass.CompProperties.n_BTUH
                Call mEvaPClass.EvapCalc(mn_CompBtuh, mn_MassFlow, mn_ET, mn_CT)

                EvLaDB = mEvaPClass.EvapProperties.n_LaDB

                If ReheatYes_No = True Then
                    Call mRecPClass.ReNT(mn_MassFlow, mn_CT, EvLaDB)
                    MassFraction = mRecPClass.ReheatProperties.n_MassFraction
                Else
                    MassFraction = 0
                End If

                Call mConPClass.CondCalc(mn_MassFlow, MassFraction, mn_ET, mn_CT)

                n_CompWattsBtuh = mEvaPClass.EvapProperties.n_BTUH + (mComPClass.CompProperties.n_WattsBTUH * 2 * 3.415)

                If mEvaPClass.EvapProperties.n_BTUH >= (mComPClass.CompProperties.n_BTUH * 2) Then
                    n_Percent1 = Abs(mEvaPClass.EvapProperties.n_BTUH - (mComPClass.CompProperties.n_BTUH * 2)) / mEvaPClass.EvapProperties.n_BTUH
                Else
                    n_Percent1 = Abs(mEvaPClass.EvapProperties.n_BTUH - (mComPClass.CompProperties.n_BTUH * 2)) / (mComPClass.CompProperties.n_BTUH * 2)
                End If

                If ((mConPClass.CondProperties.n_BTUH + mRecPClass.ReheatProperties.n_BTUH) * 2) >= n_CompWattsBtuh Then
                    n_Percent2 = Abs((mConPClass.CondProperties.n_BTUH + mRecPClass.ReheatProperties.n_BTUH) * 2 - n_CompWattsBtuh) / ((mConPClass.CondProperties.n_BTUH + mRecPClass.ReheatProperties.n_BTUH) * 2)
                Else
                    n_Percent2 = Abs((mConPClass.CondProperties.n_BTUH + mRecPClass.ReheatProperties.n_BTUH) * 2 - n_CompWattsBtuh) / (n_CompWattsBtuh)
                End If

                If mn_ET = 55 Then
                    Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                End If

                i_LoopCounter = i_LoopCounter + 1
                If n_Percent1 <= n_TolerancePercent Then
                    b_EvapHit = True
                ElseIf mEvaPClass.EvapProperties.n_BTUH < (mComPClass.CompProperties.n_BTUH * 2) Then
                    n_EvapStep = (n_Percent1 / 0.07)
                    mn_ET = mn_ET - n_EvapStep
                    If mn_ET < 10 Then
                        mn_ET = 10
                    End If
                    b_EvapHit = False
                Else
                    n_EvapStep = (n_Percent1 / 0.07)
                    mn_ET = mn_ET + n_EvapStep
                    b_EvapHit = False
                End If

                If n_Percent2 <= (n_TolerancePercent + 0.01) Then
                    b_CondHit = True
                ElseIf ((mConPClass.CondProperties.n_BTUH + mRecPClass.ReheatProperties.n_BTUH) * 2) < n_CompWattsBtuh Then
                    n_CondStep = (n_Percent2 / 0.06)
                    mn_CT = mn_CT + n_CondStep
                    If mn_CT > 155 Then
                        mn_CT = 155
                    End If
                    b_CondHit = False
                Else
                    n_CondStep = (n_Percent2 / 0.06)
                    mn_CT = mn_CT - n_CondStep
                    If mn_CT < 75 Then
                        mn_CT = 75
                    End If
                    b_CondHit = False
                End If

                If mSysP.n_ReheatYes_No <> "None" Then

                    If mRecPClass.ReheatProperties.n_DesiredBTUH <= 0 Then
                        MsgBox("Reheat temperature must be greater than LADB", vbOKOnly, "Check Reheat temperature")
                        s_Hit = " Reheat temperature must be greater than LADB"
                        Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                        Exit Do
                    End If

                    If ((mRecPClass.ReheatProperties.n_BTUH - mRecPClass.ReheatProperties.n_DesiredBTUH) / mRecPClass.ReheatProperties.n_DesiredBTUH) > ReheatDelta Then
                        MassFraction = MassFraction - 0.02
                        b_ReheatHit = False
                    ElseIf ((mRecPClass.ReheatProperties.n_BTUH - mRecPClass.ReheatProperties.n_DesiredBTUH) / mRecPClass.ReheatProperties.n_DesiredBTUH) < (-ReheatDelta) Then
                        MassFraction = MassFraction + 0.02
                        b_ReheatHit = False
                    Else

                        Call mRecPClass.VariationMassFraction(MassFraction)
                        b_ReheatHit = True
                    End If

                    If MassFraction <= 0 Or MassFraction >= 1 Then
                        If MassFraction <= 0 Then
                            MassFraction = 0.005
                            MsgBox("Difference between Reheat temperature and evap leaving temperature must be > 10 degrees", vbOKOnly, "Check Reheat temperature")
                            s_Hit = " (Reheat temperature - LADB) > 10? If Yes Change Fin/Inch, If No Correct it"
                            Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                            Exit Do
                        Else
                            MassFraction = 0.99
                            MsgBox("Difference between Reheat temperature and evap leaving temperature must be < 20 degrees", vbOKOnly, "Check Reheat temperature")
                            s_Hit = " (Reheat temperature - LADB) < 20? If Yes Change Fin/Inch, If No Correct it"
                            Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                            Exit Do
                        End If
                    End If

                    Call mRecPClass.VariationMassFraction(MassFraction)

                    If (MassFraction * 2) > 1 Then
                        MsgBox(" Reheat coil Mass Fraction exceeded 100% of one compressor mass flow")
                        s_Hit = " Reheat Mass Fraction exceeded one compressor mass flow rate"
                        Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                        Exit Do
                    End If

                    HeatRej = ((mConPClass.CondProperties.n_BTUH + mRecPClass.ReheatProperties.n_BTUH) * 2) / mEvaPClass.EvapProperties.n_BTUH

                    Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)

                    If b_CondHit And b_EvapHit And b_ReheatHit And (mn_ET < mSysP.n_ETMax) And (mn_CT < mSysP.n_CTMax) And (HeatRej < Frej) Then

                        If (mn_ET < 52.5) And (mn_CT < 130) Then
                            s_Hit = " Balance Unit!"
                        Else
                            s_Hit = " Warning! System Balance at high evap/cond Temperature"
                        End If

                        Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                        Exit Do
                    End If

                Else
                    Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                    HeatRej = (mConPClass.CondProperties.n_BTUH * 2) / mEvaPClass.EvapProperties.n_BTUH

                    If b_CondHit And b_EvapHit And (mn_ET < mSysP.n_ETMax) And (mn_CT < mSysP.n_CTMax) And (HeatRej < Frej) Then

                        If (mn_ET < 52.5) And (mn_CT < 130) Then
                            s_Hit = " Balance Unit!"
                        Else
                            s_Hit = " Warning! System Balance at high evap/Cond Temperature"
                        End If

                        Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                        Exit Do
                    End If
                End If

                If i_LoopCounter > 33 Then
                    i_DifficulityCounter = i_DifficulityCounter + 1
                    ReheatDelta = 0.05
                    If i_DifficulityCounter = 5 Then
                        s_Hit = " Could Not Balance!"
                        ReheatDelta = 0.04
                        Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                        Exit Do
                    End If

                End If

            Loop

        End If

    End Sub

    Public Sub StartSys_Balance_2Evap_HP()
        Dim mn_ET As Double
        Dim mn_CT As Double
        Dim mn_MassFlow As Double
        Dim mn_CompBtuh As Double
        Dim MassFraction As Double
        Dim n_CompWattsBtuh As Double
        Dim b_EvapHit As Boolean
        Dim b_CondHit As Boolean
        Dim i_DifficulityCounter As Integer
        Dim i_LoopCounter As Integer          'number of times thru loop
        Dim s_Hit As String
        Dim n_TolerancePercent As Double           'tolerance (input)
        Dim n_Percent1 As Double           'evap-comp percentage
        Dim n_Percent2 As Double           'cond-comp+watts percentage
        Dim n_EvapStep As Double           'increment evap temp by (Input by user)
        Dim n_CondStep As Double           'increment evap temp by (Input by user)
        Dim HeatRej As Double
        Dim Subcool As Double
        Dim Superh As Double

        Dim EvACFM As Double
        Dim CoEaDB As Double
        Dim CoEaWB As Double
        Dim CoACFM As Double
        Dim Compquantity As Integer
        Dim EvStepDiv As Double
        Dim CoStepDiv As Double
        Dim nActDensity As Double
        Dim n_EWBlblbSR As Double
        Dim n_IntV1 As Double
        Dim n_ElblbR As Double
        Dim mn_PSIA As Double
        Dim nStdDensity As Double
        Dim mn_EatWB As Double
        Dim mn_EatDB As Double
        Dim Frej As Double
        Const mn_a1 As Double = 0.00080246
        Const mn_a2 As Double = 0.000024525
        Const mn_a3 As Double = 2.542 * (10 ^ -6)
        Const mn_a4 As Double = -2.5855 * (10 ^ -8)
        Const mn_a5 As Double = 4.038 * (10 ^ -10)

        Compquantity = mSysP.i_CompQuantity
        mn_PSIA = maPSIA(mSysP.n_Altitude)
        Frej = 1.43

        If mSysP.n_Altitude > 1999 Then

            mn_EatWB = mSysP.n_EvEaWB
            mn_EatDB = mSysP.n_EvEaDB
            n_EWBlblbSR = mn_a1 + (mn_a2 * mn_EatWB) + (mn_a3 * mn_EatWB ^ 2) + (mn_a4 * mn_EatWB ^ 3) + (mn_a5 * mn_EatWB ^ 4)
            n_IntV1 = ((1093 - 0.556 * mn_EatWB) * n_EWBlblbSR) - (0.24 * (mn_EatDB - mn_EatWB))
            n_ElblbR = n_IntV1 / (1093 + 0.444 * mn_EatDB - mn_EatWB)
            nActDensity = maDensity(mn_EatDB, n_ElblbR, mn_PSIA)
            nStdDensity = maDensity(70, 0, 14.696)
            mSysP.n_EvACFM = mSysP.n_EvCFM / (nActDensity / nStdDensity)
            mSysP.n_OaACFM = mSysP.n_OaCFM / (nActDensity / nStdDensity)

            mn_EatWB = mSysP.n_CoEaWB
            mn_EatDB = mSysP.n_CoEaDB
            n_EWBlblbSR = mn_a1 + (mn_a2 * mn_EatWB) + (mn_a3 * mn_EatWB ^ 2) + (mn_a4 * mn_EatWB ^ 3) + (mn_a5 * mn_EatWB ^ 4)
            n_IntV1 = ((1093 - 0.556 * mn_EatWB) * n_EWBlblbSR) - (0.24 * (mn_EatDB - mn_EatWB))
            n_ElblbR = n_IntV1 / (1093 + 0.444 * mn_EatDB - mn_EatWB)
            nActDensity = maDensity(mn_EatDB, n_ElblbR, mn_PSIA)
            nStdDensity = maDensity(70, 0, 14.696)
            mSysP.n_CoACFM = mSysP.n_CoCFM / (nActDensity / nStdDensity)

        Else

            mSysP.n_EvACFM = mSysP.n_EvCFM
            mSysP.n_OaACFM = mSysP.n_OaCFM
            mSysP.n_CoACFM = mSysP.n_CoCFM

        End If

        If mSysP.n_OaACFM = 0 Then
            CoEaDB = mSysP.n_CoEaDB
            CoEaWB = mSysP.n_CoEaWB
            CoACFM = mSysP.n_CoACFM
        Else
            CoEaDB = mSysP.n_CoEaDB
            CoEaWB = mSysP.n_CoEaWB
            CoACFM = mSysP.n_CoACFM + mSysP.n_OaACFM

            Call MixedAir_HP(CoEaDB, CoEaWB)
            CoEaDB = CoEaDB
            CoEaWB = CoEaWB
        End If

        mSysP.n_MixedDB = CoEaDB
        mSysP.n_MixedWB = CoEaWB
        EvACFM = mSysP.n_EvACFM

        If mSysP.n_EvEaDB >= 40 Then
            EvStepDiv = 0.07 * 3
        ElseIf mSysP.n_EvEaDB >= 30 Then
            EvStepDiv = 0.07 * 3.5
        ElseIf mSysP.n_EvEaDB >= 15 Then
            EvStepDiv = 0.07 * 4
        Else
            EvStepDiv = 0.07 * 5
        End If

        If mSysP.n_EvEaDB >= 30 Then
            CoStepDiv = 0.06
        ElseIf mSysP.n_EvEaDB >= 15 Then
            CoStepDiv = 0.06 * 2
        Else
            CoStepDiv = 0.06 * 3
        End If

        If StartCalculations(Compquantity) Then
            mn_ET = 35
            mn_CT = mSysP.n_MixedDB + 25

            i_LoopCounter = 0
            s_Hit = ""
            n_TolerancePercent = 0.01

            Subcool = mSysP.n_Subcool
            Superh = mSysP.n_Superh

            Call mEvaPClass.PassAirCorrection(mSysP.n_EvEaDB, mSysP.n_EvEaWB, EvACFM, mSysP.s_CompAppCode)
            Call mConPClass.PassMixAirInfo(CoEaDB, CoEaWB, CoACFM, mSysP.s_CompAppCode)

            Do

                Call mComPClass.CompCalc(mn_ET, mn_CT)
                mn_MassFlow = mComPClass.CompProperties.n_MassFlow
                mn_CompBtuh = mComPClass.CompProperties.n_BTUH
                Call mEvaPClass.EvapCalc(mn_CompBtuh, mn_MassFlow, mn_ET, mn_CT)

                MassFraction = 0

                Call mConPClass.CondCalc(mn_MassFlow, MassFraction, mn_ET, mn_CT)

                n_CompWattsBtuh = (mEvaPClass.EvapProperties.n_BTUH * 2) + (mComPClass.CompProperties.n_WattsBTUH * 2 * 3.415)

                If (mEvaPClass.EvapProperties.n_BTUH * 2) >= (mComPClass.CompProperties.n_BTUH * 2) Then
                    n_Percent1 = Abs((mEvaPClass.EvapProperties.n_BTUH * 2) - (mComPClass.CompProperties.n_BTUH * 2)) / (mEvaPClass.EvapProperties.n_BTUH * 2)
                Else
                    n_Percent1 = Abs((mEvaPClass.EvapProperties.n_BTUH * 2) - (mComPClass.CompProperties.n_BTUH * 2)) / (mComPClass.CompProperties.n_BTUH * 2)
                End If

                If (mConPClass.CondProperties.n_BTUH) >= n_CompWattsBtuh Then
                    n_Percent2 = Abs(mConPClass.CondProperties.n_BTUH - n_CompWattsBtuh) / (mConPClass.CondProperties.n_BTUH)
                Else
                    n_Percent2 = Abs(mConPClass.CondProperties.n_BTUH - n_CompWattsBtuh) / (n_CompWattsBtuh)
                End If

                If mn_ET = 35 Then
                    Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                End If

                i_LoopCounter = i_LoopCounter + 1
                If n_Percent1 <= n_TolerancePercent Then
                    b_EvapHit = True
                ElseIf (mEvaPClass.EvapProperties.n_BTUH * 2) < (mComPClass.CompProperties.n_BTUH * 2) Then
                    n_EvapStep = (n_Percent1 / EvStepDiv)
                    mn_ET = mn_ET - n_EvapStep
                    If mn_ET < 2 Then
                        mn_ET = 2
                    End If
                    b_EvapHit = False
                Else
                    n_EvapStep = (n_Percent1 / EvStepDiv)
                    mn_ET = mn_ET + n_EvapStep
                    b_EvapHit = False
                End If

                If n_Percent2 <= (n_TolerancePercent + 0.01) Then
                    b_CondHit = True
                ElseIf (mConPClass.CondProperties.n_BTUH) < n_CompWattsBtuh Then
                    n_CondStep = (n_Percent2 / CoStepDiv)
                    mn_CT = mn_CT + n_CondStep
                    If mn_CT > 155 Then
                        mn_CT = 155
                    End If
                    b_CondHit = False
                Else
                    n_CondStep = (n_Percent2 / CoStepDiv)
                    mn_CT = mn_CT - n_CondStep
                    If mn_CT < 75 Then
                        mn_CT = 75
                    End If
                    b_CondHit = False
                End If

                Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                HeatRej = (mConPClass.CondProperties.n_BTUH) / (mEvaPClass.EvapProperties.n_BTUH * 2)

                If b_CondHit And b_EvapHit And (mn_ET < mSysP.n_ETMax) And (mn_CT < mSysP.n_CTMax) And (HeatRej < Frej) Then

                    If (mn_ET > 30) And (mn_ET < 50) And (mn_CT < 130) Then
                        s_Hit = " Balance Unit!"
                    ElseIf (mn_ET > 30) Then
                        s_Hit = " Warning! System Balance at high evap or cond Temperature"
                    ElseIf (mn_ET < 30) Then
                        s_Hit = " Warning! System will require longer defrosting time"
                    End If

                    Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                    Exit Do
                End If

                If i_LoopCounter > 33 Then
                    i_DifficulityCounter = i_DifficulityCounter + 1
                    If i_DifficulityCounter = 5 Then
                        s_Hit = " Could Not Balance!"
                        Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                        Exit Do
                    End If

                End If

            Loop

        End If
    End Sub

    Public Sub StartSys_Balance_ITEITC_AC()
        Dim mn_ET As Double
        Dim mn_CT As Double
        Dim mn_MassFlow As Double
        Dim mn_CompBtuh As Double
        Dim MassFraction As Double
        Dim n_CompWattsBtuh As Double
        Dim b_EvapHit As Boolean
        Dim b_CondHit As Boolean
        Dim b_ReheatHit As Boolean
        Dim i_DifficulityCounter As Integer
        Dim i_LoopCounter As Integer          'number of times thru loop
        Dim s_Hit As String
        Dim n_TolerancePercent As Double           'tolerance (input)
        Dim n_Percent1 As Double           'evap-comp percentage
        Dim n_Percent2 As Double           'cond-comp+watts percentage
        Dim n_EvapStep As Double           'increment evap temp by (Input by user)
        Dim n_CondStep As Double           'increment evap temp by (Input by user)
        Dim ReheatYes_No As Boolean
        Dim HeatRej As Double
        Dim Subcool As Double
        Dim Superh As Double
        Dim ReheatDelta As Double
        Dim EvACFM As Double
        Dim EvEaDB As Double
        Dim EvEaWB As Double
        Dim CoACFM As Double
        Dim EvLaDB As Double
        Dim Compquantity As Integer

        Dim nActDensity As Double
        Dim n_EWBlblbSR As Double
        Dim n_IntV1 As Double
        Dim n_ElblbR As Double
        Dim mn_PSIA As Double
        Dim nStdDensity As Double
        Dim mn_EatWB As Double
        Dim mn_EatDB As Double
        Dim Frej As Double
        Const mn_a1 As Double = 0.00080246
        Const mn_a2 As Double = 0.000024525
        Const mn_a3 As Double = 2.542 * (10 ^ -6)
        Const mn_a4 As Double = -2.5855 * (10 ^ -8)
        Const mn_a5 As Double = 4.038 * (10 ^ -10)

        Compquantity = mSysP.i_CompQuantity
        mn_PSIA = maPSIA(mSysP.n_Altitude)
        Frej = 1.33
        ReheatDelta = 0.04

        If mSysP.n_Altitude > 1999 Then

            mn_EatWB = mSysP.n_EvEaWB
            mn_EatDB = mSysP.n_EvEaDB
            n_EWBlblbSR = mn_a1 + (mn_a2 * mn_EatWB) + (mn_a3 * mn_EatWB ^ 2) + (mn_a4 * mn_EatWB ^ 3) + (mn_a5 * mn_EatWB ^ 4)
            n_IntV1 = ((1093 - 0.556 * mn_EatWB) * n_EWBlblbSR) - (0.24 * (mn_EatDB - mn_EatWB))
            n_ElblbR = n_IntV1 / (1093 + 0.444 * mn_EatDB - mn_EatWB)
            nActDensity = maDensity(mn_EatDB, n_ElblbR, mn_PSIA)
            nStdDensity = maDensity(70, 0, 14.696)
            mSysP.n_EvACFM = mSysP.n_EvCFM / (nActDensity / nStdDensity)

            mn_EatWB = mSysP.n_CoEaWB
            mn_EatDB = mSysP.n_CoEaDB
            n_EWBlblbSR = mn_a1 + (mn_a2 * mn_EatWB) + (mn_a3 * mn_EatWB ^ 2) + (mn_a4 * mn_EatWB ^ 3) + (mn_a5 * mn_EatWB ^ 4)
            n_IntV1 = ((1093 - 0.556 * mn_EatWB) * n_EWBlblbSR) - (0.24 * (mn_EatDB - mn_EatWB))
            n_ElblbR = n_IntV1 / (1093 + 0.444 * mn_EatDB - mn_EatWB)
            nActDensity = maDensity(mn_EatDB, n_ElblbR, mn_PSIA)
            nStdDensity = maDensity(70, 0, 14.696)
            mSysP.n_OaACFM = mSysP.n_OaCFM / (nActDensity / nStdDensity)
            mSysP.n_CoACFM = mSysP.n_CoCFM / (nActDensity / nStdDensity)

        Else
            mSysP.n_EvACFM = mSysP.n_EvCFM
            mSysP.n_OaACFM = mSysP.n_OaCFM
            mSysP.n_CoACFM = mSysP.n_CoCFM
        End If

        If mSysP.n_OaACFM = 0 Then
            EvEaDB = mSysP.n_EvEaDB
            EvEaWB = mSysP.n_EvEaWB
            EvACFM = mSysP.n_EvACFM
        Else
            EvEaDB = mSysP.n_EvEaDB
            EvEaWB = mSysP.n_EvEaWB
            EvACFM = mSysP.n_EvACFM + mSysP.n_OaACFM

            Call MixedAir_AC(EvEaDB, EvEaWB)
            EvEaDB = EvEaDB
            EvEaWB = EvEaWB
        End If

        mSysP.n_MixedDB = EvEaDB
        mSysP.n_MixedWB = EvEaWB
        CoACFM = mSysP.n_CoACFM

        If StartCalculations(Compquantity) Then
            mn_ET = 55
            mn_CT = mSysP.n_CoEaDB + 25

            i_LoopCounter = 0
            s_Hit = ""
            n_TolerancePercent = 0.01


            Subcool = mSysP.n_Subcool
            Superh = mSysP.n_Superh

            Call mEvaPClass.PassMixAirInfo(EvEaDB, EvEaWB, EvACFM, mSysP.s_CompAppCode)
            Call mConPClass.PassAirCorrection(mSysP.n_CoEaDB, mSysP.n_CoEaWB, CoACFM, mSysP.s_CompAppCode)
            Call mRecPClass.PassMixAirCFM(EvACFM)

            If mSysP.n_ReheatYes_No <> "None" Then
                ReheatYes_No = True
            Else
                ReheatYes_No = False
            End If

            Do

                Call mComPClass.CompCalc(mn_ET, mn_CT)
                mn_MassFlow = mComPClass.CompProperties.n_MassFlow
                mn_CompBtuh = mComPClass.CompProperties.n_BTUH
                Call mEvaPClass.EvapCalc(mn_CompBtuh, mn_MassFlow, mn_ET, mn_CT)

                EvLaDB = mEvaPClass.EvapProperties.n_LaDB

                If ReheatYes_No = True Then
                    Call mRecPClass.ReNT(mn_MassFlow, mn_CT, EvLaDB)
                    MassFraction = mRecPClass.ReheatProperties.n_MassFraction
                Else
                    MassFraction = 0
                End If

                Call mConPClass.CondCalc(mn_MassFlow, MassFraction, mn_ET, mn_CT)

                n_CompWattsBtuh = mEvaPClass.EvapProperties.n_BTUH + (mComPClass.CompProperties.n_WattsBTUH * 2 * 3.415)

                If mEvaPClass.EvapProperties.n_BTUH >= (mComPClass.CompProperties.n_BTUH * 2) Then
                    n_Percent1 = Abs(mEvaPClass.EvapProperties.n_BTUH - (mComPClass.CompProperties.n_BTUH * 2)) / mEvaPClass.EvapProperties.n_BTUH
                Else
                    n_Percent1 = Abs(mEvaPClass.EvapProperties.n_BTUH - (mComPClass.CompProperties.n_BTUH * 2)) / (mComPClass.CompProperties.n_BTUH * 2)
                End If

                If (mConPClass.CondProperties.n_BTUH + (mRecPClass.ReheatProperties.n_BTUH * 2)) >= n_CompWattsBtuh Then
                    n_Percent2 = Abs(mConPClass.CondProperties.n_BTUH + (mRecPClass.ReheatProperties.n_BTUH * 2) - n_CompWattsBtuh) / (mConPClass.CondProperties.n_BTUH + (mRecPClass.ReheatProperties.n_BTUH * 2))
                Else
                    n_Percent2 = Abs(mConPClass.CondProperties.n_BTUH + (mRecPClass.ReheatProperties.n_BTUH * 2) - n_CompWattsBtuh) / (n_CompWattsBtuh)
                End If

                If mn_ET = 55 Then
                    Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                End If

                i_LoopCounter = i_LoopCounter + 1
                If n_Percent1 <= n_TolerancePercent Then
                    b_EvapHit = True
                ElseIf mEvaPClass.EvapProperties.n_BTUH < (mComPClass.CompProperties.n_BTUH * 2) Then
                    n_EvapStep = (n_Percent1 / 0.07)
                    mn_ET = mn_ET - n_EvapStep
                    If mn_ET < 10 Then
                        mn_ET = 10
                    End If
                    b_EvapHit = False
                Else
                    n_EvapStep = (n_Percent1 / 0.07)
                    mn_ET = mn_ET + n_EvapStep
                    b_EvapHit = False
                End If

                If n_Percent2 <= (n_TolerancePercent + 0.01) Then
                    b_CondHit = True
                ElseIf (mConPClass.CondProperties.n_BTUH + (mRecPClass.ReheatProperties.n_BTUH * 2)) < n_CompWattsBtuh Then
                    n_CondStep = (n_Percent2 / 0.06)
                    mn_CT = mn_CT + n_CondStep
                    If mn_CT > 155 Then
                        mn_CT = 155
                    End If
                    b_CondHit = False
                Else
                    n_CondStep = (n_Percent2 / 0.06)
                    mn_CT = mn_CT - n_CondStep
                    If mn_CT < 75 Then
                        mn_CT = 75
                    End If
                    b_CondHit = False
                End If

                If mSysP.n_ReheatYes_No <> "None" Then

                    If mRecPClass.ReheatProperties.n_DesiredBTUH <= 0 Then
                        MsgBox("Reheat temperature must be greater than LADB", vbOKOnly, "Check Reheat temperature")
                        s_Hit = " Reheat temperature must be greater than LADB"
                        Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                        Exit Do
                    End If

                    If ((mRecPClass.ReheatProperties.n_BTUH - mRecPClass.ReheatProperties.n_DesiredBTUH) / mRecPClass.ReheatProperties.n_DesiredBTUH) > ReheatDelta Then
                        MassFraction = MassFraction - 0.02
                        b_ReheatHit = False
                    ElseIf ((mRecPClass.ReheatProperties.n_BTUH - mRecPClass.ReheatProperties.n_DesiredBTUH) / mRecPClass.ReheatProperties.n_DesiredBTUH) < (-ReheatDelta) Then
                        MassFraction = MassFraction + 0.02
                        b_ReheatHit = False
                    Else

                        Call mRecPClass.VariationMassFraction(MassFraction)
                        b_ReheatHit = True
                    End If

                    If MassFraction <= 0 Or MassFraction >= 1 Then
                        If MassFraction <= 0 Then
                            MassFraction = 0.005
                            MsgBox("Difference between Reheat temperature and evap leaving temperature must be > 10 degrees", vbOKOnly, "Check Reheat temperature")
                            s_Hit = " (Reheat temperature - LADB) > 10? If Yes Change Fin/Inch, If No Correct it"
                            Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                            Exit Do
                        Else
                            MassFraction = 0.99
                            MsgBox("Difference between Reheat temperature and evap leaving temperature must be < 20 degrees", vbOKOnly, "Check Reheat temperature")
                            s_Hit = " (Reheat temperature - LADB) < 20? If Yes Change Fin/Inch, If No Correct it"
                            Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                            Exit Do
                        End If
                    End If

                    Call mRecPClass.VariationMassFraction(MassFraction)

                    If (MassFraction * 2) > 1 Then
                        MsgBox(" Reheat coil Mass Fraction exceeded 100% of one compressor mass flow")
                        s_Hit = " Reheat Mass Fraction exceeded one compressor mass flow rate"
                        Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                        Exit Do
                    End If

                    HeatRej = (mConPClass.CondProperties.n_BTUH + (mRecPClass.ReheatProperties.n_BTUH * 2)) / mEvaPClass.EvapProperties.n_BTUH

                    Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)

                    If b_CondHit And b_EvapHit And b_ReheatHit And (mn_ET < mSysP.n_ETMax) And (mn_CT < mSysP.n_CTMax) And (HeatRej < Frej) Then

                        If (mn_ET < 52.5) And (mn_CT < 130) Then
                            s_Hit = " Balance Unit!"
                        Else
                            s_Hit = " Warning! System Balance at high evap/cond Temperature"
                        End If

                        Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                        Exit Do
                    End If

                Else
                    Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                    HeatRej = (mConPClass.CondProperties.n_BTUH) / mEvaPClass.EvapProperties.n_BTUH

                    If b_CondHit And b_EvapHit And (mn_ET < mSysP.n_ETMax) And (mn_CT < mSysP.n_CTMax) And (HeatRej < Frej) Then

                        If (mn_ET < 52.5) And (mn_CT < 130) Then
                            s_Hit = " Balance Unit!"
                        Else
                            s_Hit = " Warning! System Balance at high evap/Cond Temperature"
                        End If

                        Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                        Exit Do
                    End If
                End If

                If i_LoopCounter > 33 Then
                    i_DifficulityCounter = i_DifficulityCounter + 1
                    ReheatDelta = 0.05
                    If i_DifficulityCounter = 5 Then
                        s_Hit = " Could Not Balance!"
                        ReheatDelta = 0.04
                        Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                        Exit Do
                    End If

                End If

            Loop

        End If

    End Sub

    Public Sub StartSys_Balance_ITEITC_HP()
        Dim mn_ET As Double
        Dim mn_CT As Double
        Dim mn_MassFlow As Double
        Dim mn_CompBtuh As Double
        Dim MassFraction As Double
        Dim n_CompWattsBtuh As Double
        Dim b_EvapHit As Boolean
        Dim b_CondHit As Boolean
        Dim i_DifficulityCounter As Integer
        Dim i_LoopCounter As Integer          'number of times thru loop
        Dim s_Hit As String
        Dim n_TolerancePercent As Double           'tolerance (input)
        Dim n_Percent1 As Double           'evap-comp percentage
        Dim n_Percent2 As Double           'cond-comp+watts percentage
        Dim n_EvapStep As Double           'increment evap temp by (Input by user)
        Dim n_CondStep As Double           'increment evap temp by (Input by user)
        Dim HeatRej As Double
        Dim Subcool As Double
        Dim Superh As Double

        Dim EvACFM As Double
        Dim CoEaDB As Double
        Dim CoEaWB As Double
        Dim CoACFM As Double
        Dim Compquantity As Integer
        Dim EvStepDiv As Double
        Dim CoStepDiv As Double
        Dim nActDensity As Double
        Dim n_EWBlblbSR As Double
        Dim n_IntV1 As Double
        Dim n_ElblbR As Double
        Dim mn_PSIA As Double
        Dim nStdDensity As Double
        Dim mn_EatWB As Double
        Dim mn_EatDB As Double
        Dim Frej As Double
        Const mn_a1 As Double = 0.00080246
        Const mn_a2 As Double = 0.000024525
        Const mn_a3 As Double = 2.542 * (10 ^ -6)
        Const mn_a4 As Double = -2.5855 * (10 ^ -8)
        Const mn_a5 As Double = 4.038 * (10 ^ -10)

        Compquantity = mSysP.i_CompQuantity
        mn_PSIA = maPSIA(mSysP.n_Altitude)
        Frej = 1.43

        If mSysP.n_Altitude > 1999 Then

            mn_EatWB = mSysP.n_EvEaWB
            mn_EatDB = mSysP.n_EvEaDB
            n_EWBlblbSR = mn_a1 + (mn_a2 * mn_EatWB) + (mn_a3 * mn_EatWB ^ 2) + (mn_a4 * mn_EatWB ^ 3) + (mn_a5 * mn_EatWB ^ 4)
            n_IntV1 = ((1093 - 0.556 * mn_EatWB) * n_EWBlblbSR) - (0.24 * (mn_EatDB - mn_EatWB))
            n_ElblbR = n_IntV1 / (1093 + 0.444 * mn_EatDB - mn_EatWB)
            nActDensity = maDensity(mn_EatDB, n_ElblbR, mn_PSIA)
            nStdDensity = maDensity(70, 0, 14.696)
            mSysP.n_EvACFM = mSysP.n_EvCFM / (nActDensity / nStdDensity)
            mSysP.n_OaACFM = mSysP.n_OaCFM / (nActDensity / nStdDensity)

            mn_EatWB = mSysP.n_CoEaWB
            mn_EatDB = mSysP.n_CoEaDB
            n_EWBlblbSR = mn_a1 + (mn_a2 * mn_EatWB) + (mn_a3 * mn_EatWB ^ 2) + (mn_a4 * mn_EatWB ^ 3) + (mn_a5 * mn_EatWB ^ 4)
            n_IntV1 = ((1093 - 0.556 * mn_EatWB) * n_EWBlblbSR) - (0.24 * (mn_EatDB - mn_EatWB))
            n_ElblbR = n_IntV1 / (1093 + 0.444 * mn_EatDB - mn_EatWB)
            nActDensity = maDensity(mn_EatDB, n_ElblbR, mn_PSIA)
            nStdDensity = maDensity(70, 0, 14.696)
            mSysP.n_CoACFM = mSysP.n_CoCFM / (nActDensity / nStdDensity)

        Else

            mSysP.n_EvACFM = mSysP.n_EvCFM
            mSysP.n_OaACFM = mSysP.n_OaCFM
            mSysP.n_CoACFM = mSysP.n_CoCFM

        End If

        If mSysP.n_OaACFM = 0 Then
            CoEaDB = mSysP.n_CoEaDB
            CoEaWB = mSysP.n_CoEaWB
            CoACFM = mSysP.n_CoACFM
        Else
            CoEaDB = mSysP.n_CoEaDB
            CoEaWB = mSysP.n_CoEaWB
            CoACFM = mSysP.n_CoACFM + mSysP.n_OaACFM

            Call MixedAir_HP(CoEaDB, CoEaWB)
            CoEaDB = CoEaDB
            CoEaWB = CoEaWB
        End If

        mSysP.n_MixedDB = CoEaDB
        mSysP.n_MixedWB = CoEaWB
        EvACFM = mSysP.n_EvACFM

        If mSysP.n_EvEaDB >= 40 Then
            EvStepDiv = 0.07 * 3
        ElseIf mSysP.n_EvEaDB >= 30 Then
            EvStepDiv = 0.07 * 3.5
        ElseIf mSysP.n_EvEaDB >= 15 Then
            EvStepDiv = 0.07 * 4
        Else
            EvStepDiv = 0.07 * 5
        End If

        If mSysP.n_EvEaDB >= 30 Then
            CoStepDiv = 0.06
        ElseIf mSysP.n_EvEaDB >= 15 Then
            CoStepDiv = 0.06 * 2
        Else
            CoStepDiv = 0.06 * 3
        End If

        If StartCalculations(Compquantity) Then
            mn_ET = 35
            mn_CT = mSysP.n_MixedDB + 25

            i_LoopCounter = 0
            s_Hit = ""
            n_TolerancePercent = 0.01

            Subcool = mSysP.n_Subcool
            Superh = mSysP.n_Superh

            Call mEvaPClass.PassAirCorrection(mSysP.n_EvEaDB, mSysP.n_EvEaWB, EvACFM, mSysP.s_CompAppCode)
            Call mConPClass.PassMixAirInfo(CoEaDB, CoEaWB, CoACFM, mSysP.s_CompAppCode)

            Do

                Call mComPClass.CompCalc(mn_ET, mn_CT)
                mn_MassFlow = mComPClass.CompProperties.n_MassFlow
                mn_CompBtuh = mComPClass.CompProperties.n_BTUH
                Call mEvaPClass.EvapCalc(mn_CompBtuh, mn_MassFlow, mn_ET, mn_CT)

                MassFraction = 0

                Call mConPClass.CondCalc(mn_MassFlow, MassFraction, mn_ET, mn_CT)

                n_CompWattsBtuh = (mEvaPClass.EvapProperties.n_BTUH) + (mComPClass.CompProperties.n_WattsBTUH * 2 * 3.415)

                If (mEvaPClass.EvapProperties.n_BTUH) >= (mComPClass.CompProperties.n_BTUH * 2) Then
                    n_Percent1 = Abs((mEvaPClass.EvapProperties.n_BTUH) - (mComPClass.CompProperties.n_BTUH * 2)) / (mEvaPClass.EvapProperties.n_BTUH)
                Else
                    n_Percent1 = Abs((mEvaPClass.EvapProperties.n_BTUH) - (mComPClass.CompProperties.n_BTUH * 2)) / (mComPClass.CompProperties.n_BTUH * 2)
                End If

                If (mConPClass.CondProperties.n_BTUH) >= n_CompWattsBtuh Then
                    n_Percent2 = Abs(mConPClass.CondProperties.n_BTUH - n_CompWattsBtuh) / (mConPClass.CondProperties.n_BTUH)
                Else
                    n_Percent2 = Abs(mConPClass.CondProperties.n_BTUH - n_CompWattsBtuh) / (n_CompWattsBtuh)
                End If

                If mn_ET = 35 Then
                    Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                End If

                i_LoopCounter = i_LoopCounter + 1
                If n_Percent1 <= n_TolerancePercent Then
                    b_EvapHit = True
                ElseIf (mEvaPClass.EvapProperties.n_BTUH) < (mComPClass.CompProperties.n_BTUH * 2) Then
                    n_EvapStep = (n_Percent1 / EvStepDiv)
                    mn_ET = mn_ET - n_EvapStep
                    If mn_ET < 2 Then
                        mn_ET = 2
                    End If
                    b_EvapHit = False
                Else
                    n_EvapStep = (n_Percent1 / EvStepDiv)
                    mn_ET = mn_ET + n_EvapStep
                    b_EvapHit = False
                End If

                If n_Percent2 <= (n_TolerancePercent + 0.01) Then
                    b_CondHit = True
                ElseIf (mConPClass.CondProperties.n_BTUH) < n_CompWattsBtuh Then
                    n_CondStep = (n_Percent2 / CoStepDiv)
                    mn_CT = mn_CT + n_CondStep
                    If mn_CT > 155 Then
                        mn_CT = 155
                    End If
                    b_CondHit = False
                Else
                    n_CondStep = (n_Percent2 / CoStepDiv)
                    mn_CT = mn_CT - n_CondStep
                    If mn_CT < 75 Then
                        mn_CT = 75
                    End If
                    b_CondHit = False
                End If

                Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                HeatRej = (mConPClass.CondProperties.n_BTUH) / (mEvaPClass.EvapProperties.n_BTUH)

                If b_CondHit And b_EvapHit And (mn_ET < mSysP.n_ETMax) And (mn_CT < mSysP.n_CTMax) And (HeatRej < Frej) Then

                    If (mn_ET > 30) And (mn_ET < 50) And (mn_CT < 130) Then
                        s_Hit = " Balance Unit!"
                    ElseIf (mn_ET > 30) Then
                        s_Hit = " Warning! System Balance at high evap or cond Temperature"
                    ElseIf (mn_ET < 30) Then
                        s_Hit = " Warning! System will require longer defrosting time"
                    End If

                    Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                    Exit Do
                End If

                If i_LoopCounter > 33 Then
                    i_DifficulityCounter = i_DifficulityCounter + 1
                    If i_DifficulityCounter = 5 Then
                        s_Hit = " Could Not Balance!"
                        Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                        Exit Do
                    End If

                End If

            Loop

        End If
    End Sub

    Public Sub StartSys_Balance_FSEITC_AC()
        Dim mn_ET As Double
        Dim mn_CT As Double
        Dim mn_MassFlow As Double
        Dim mn_CompBtuh As Double
        Dim MassFraction As Double
        Dim n_CompWattsBtuh As Double
        Dim b_EvapHit As Boolean
        Dim b_CondHit As Boolean
        Dim b_ReheatHit As Boolean
        Dim i_DifficulityCounter As Integer
        Dim i_LoopCounter As Integer          'number of times thru loop
        Dim s_Hit As String
        Dim n_TolerancePercent As Double           'tolerance (input)
        Dim n_Percent1 As Double           'evap-comp percentage
        Dim n_Percent2 As Double           'cond-comp+watts percentage
        Dim n_EvapStep As Double           'increment evap temp by (Input by user)
        Dim n_CondStep As Double           'increment evap temp by (Input by user)
        Dim ReheatYes_No As Boolean
        Dim HeatRej As Double
        Dim Subcool As Double
        Dim Superh As Double
        Dim ReheatDelta As Double
        Dim EvACFM As Double
        Dim EvEaDB As Double
        Dim EvEaWB As Double
        Dim CoACFM As Double
        Dim EvLaDB As Double
        Dim Compquantity As Integer

        Dim nActDensity As Double
        Dim n_EWBlblbSR As Double
        Dim n_IntV1 As Double
        Dim n_ElblbR As Double
        Dim mn_PSIA As Double
        Dim nStdDensity As Double
        Dim mn_EatWB As Double
        Dim mn_EatDB As Double
        Dim Frej As Double
        Const mn_a1 As Double = 0.00080246
        Const mn_a2 As Double = 0.000024525
        Const mn_a3 As Double = 2.542 * (10 ^ -6)
        Const mn_a4 As Double = -2.5855 * (10 ^ -8)
        Const mn_a5 As Double = 4.038 * (10 ^ -10)

        Compquantity = mSysP.i_CompQuantity
        mn_PSIA = maPSIA(mSysP.n_Altitude)
        Frej = 1.33
        ReheatDelta = 0.04

        If mSysP.n_Altitude > 1999 Then

            mn_EatWB = mSysP.n_EvEaWB
            mn_EatDB = mSysP.n_EvEaDB
            n_EWBlblbSR = mn_a1 + (mn_a2 * mn_EatWB) + (mn_a3 * mn_EatWB ^ 2) + (mn_a4 * mn_EatWB ^ 3) + (mn_a5 * mn_EatWB ^ 4)
            n_IntV1 = ((1093 - 0.556 * mn_EatWB) * n_EWBlblbSR) - (0.24 * (mn_EatDB - mn_EatWB))
            n_ElblbR = n_IntV1 / (1093 + 0.444 * mn_EatDB - mn_EatWB)
            nActDensity = maDensity(mn_EatDB, n_ElblbR, mn_PSIA)
            nStdDensity = maDensity(70, 0, 14.696)
            mSysP.n_EvACFM = mSysP.n_EvCFM / (nActDensity / nStdDensity)

            mn_EatWB = mSysP.n_CoEaWB
            mn_EatDB = mSysP.n_CoEaDB
            n_EWBlblbSR = mn_a1 + (mn_a2 * mn_EatWB) + (mn_a3 * mn_EatWB ^ 2) + (mn_a4 * mn_EatWB ^ 3) + (mn_a5 * mn_EatWB ^ 4)
            n_IntV1 = ((1093 - 0.556 * mn_EatWB) * n_EWBlblbSR) - (0.24 * (mn_EatDB - mn_EatWB))
            n_ElblbR = n_IntV1 / (1093 + 0.444 * mn_EatDB - mn_EatWB)
            nActDensity = maDensity(mn_EatDB, n_ElblbR, mn_PSIA)
            nStdDensity = maDensity(70, 0, 14.696)
            mSysP.n_OaACFM = mSysP.n_OaCFM / (nActDensity / nStdDensity)
            mSysP.n_CoACFM = mSysP.n_CoCFM / (nActDensity / nStdDensity)

        Else
            mSysP.n_EvACFM = mSysP.n_EvCFM
            mSysP.n_OaACFM = mSysP.n_OaCFM
            mSysP.n_CoACFM = mSysP.n_CoCFM
        End If

        If mSysP.n_OaACFM = 0 Then
            EvEaDB = mSysP.n_EvEaDB
            EvEaWB = mSysP.n_EvEaWB
            EvACFM = mSysP.n_EvACFM
        Else
            EvEaDB = mSysP.n_EvEaDB
            EvEaWB = mSysP.n_EvEaWB
            EvACFM = mSysP.n_EvACFM + mSysP.n_OaACFM

            Call MixedAir_AC(EvEaDB, EvEaWB)
            EvEaDB = EvEaDB
            EvEaWB = EvEaWB
        End If

        mSysP.n_MixedDB = EvEaDB
        mSysP.n_MixedWB = EvEaWB
        CoACFM = mSysP.n_CoACFM

        If StartCalculations(Compquantity) Then
            mn_ET = 55
            mn_CT = mSysP.n_CoEaDB + 25

            i_LoopCounter = 0
            s_Hit = ""
            n_TolerancePercent = 0.01


            Subcool = mSysP.n_Subcool
            Superh = mSysP.n_Superh

            Call mEvaPClass.PassMixAirInfo(EvEaDB, EvEaWB, EvACFM, mSysP.s_CompAppCode)
            Call mConPClass.PassAirCorrection(mSysP.n_CoEaDB, mSysP.n_CoEaWB, CoACFM, mSysP.s_CompAppCode)
            Call mRecPClass.PassMixAirCFM(EvACFM)

            If mSysP.n_ReheatYes_No <> "None" Then
                ReheatYes_No = True
            Else
                ReheatYes_No = False
            End If

            Do

                Call mComPClass.CompCalc(mn_ET, mn_CT)
                mn_MassFlow = mComPClass.CompProperties.n_MassFlow
                mn_CompBtuh = mComPClass.CompProperties.n_BTUH
                Call mEvaPClass.EvapCalc(mn_CompBtuh, mn_MassFlow, mn_ET, mn_CT)

                EvLaDB = mEvaPClass.EvapProperties.n_LaDB

                If ReheatYes_No = True Then
                    Call mRecPClass.ReNT(mn_MassFlow, mn_CT, EvLaDB)
                    MassFraction = mRecPClass.ReheatProperties.n_MassFraction
                Else
                    MassFraction = 0
                End If

                Call mConPClass.CondCalc(mn_MassFlow, MassFraction, mn_ET, mn_CT)

                n_CompWattsBtuh = (mEvaPClass.EvapProperties.n_BTUH * 2) + (mComPClass.CompProperties.n_WattsBTUH * 2 * 3.415)

                If (mEvaPClass.EvapProperties.n_BTUH * 2) >= (mComPClass.CompProperties.n_BTUH * 2) Then
                    n_Percent1 = Abs((mEvaPClass.EvapProperties.n_BTUH * 2) - (mComPClass.CompProperties.n_BTUH * 2)) / (mEvaPClass.EvapProperties.n_BTUH * 2)
                Else
                    n_Percent1 = Abs((mEvaPClass.EvapProperties.n_BTUH * 2) - (mComPClass.CompProperties.n_BTUH * 2)) / (mComPClass.CompProperties.n_BTUH * 2)
                End If

                If (mConPClass.CondProperties.n_BTUH + (mRecPClass.ReheatProperties.n_BTUH * 2)) >= n_CompWattsBtuh Then
                    n_Percent2 = Abs(mConPClass.CondProperties.n_BTUH + (mRecPClass.ReheatProperties.n_BTUH * 2) - n_CompWattsBtuh) / (mConPClass.CondProperties.n_BTUH + (mRecPClass.ReheatProperties.n_BTUH * 2))
                Else
                    n_Percent2 = Abs(mConPClass.CondProperties.n_BTUH + (mRecPClass.ReheatProperties.n_BTUH * 2) - n_CompWattsBtuh) / (n_CompWattsBtuh)
                End If

                If mn_ET = 55 Then
                    Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                End If

                i_LoopCounter = i_LoopCounter + 1
                If n_Percent1 <= n_TolerancePercent Then
                    b_EvapHit = True
                ElseIf (mEvaPClass.EvapProperties.n_BTUH * 2) < (mComPClass.CompProperties.n_BTUH * 2) Then
                    n_EvapStep = (n_Percent1 / 0.07)
                    mn_ET = mn_ET - n_EvapStep
                    If mn_ET < 10 Then
                        mn_ET = 10
                    End If
                    b_EvapHit = False
                Else
                    n_EvapStep = (n_Percent1 / 0.07)
                    mn_ET = mn_ET + n_EvapStep
                    b_EvapHit = False
                End If

                If n_Percent2 <= (n_TolerancePercent + 0.01) Then
                    b_CondHit = True
                ElseIf (mConPClass.CondProperties.n_BTUH + (mRecPClass.ReheatProperties.n_BTUH * 2)) < n_CompWattsBtuh Then
                    n_CondStep = (n_Percent2 / 0.06)
                    mn_CT = mn_CT + n_CondStep
                    If mn_CT > 155 Then
                        mn_CT = 155
                    End If
                    b_CondHit = False
                Else
                    n_CondStep = (n_Percent2 / 0.06)
                    mn_CT = mn_CT - n_CondStep
                    If mn_CT < 75 Then
                        mn_CT = 75
                    End If
                    b_CondHit = False
                End If

                If mSysP.n_ReheatYes_No <> "None" Then

                    If mRecPClass.ReheatProperties.n_DesiredBTUH <= 0 Then
                        MsgBox("Reheat temperature must be greater than LADB", vbOKOnly, "Check Reheat temperature")
                        s_Hit = " Reheat temperature must be greater than LADB"
                        Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                        Exit Do
                    End If

                    If ((mRecPClass.ReheatProperties.n_BTUH - mRecPClass.ReheatProperties.n_DesiredBTUH) / mRecPClass.ReheatProperties.n_DesiredBTUH) > ReheatDelta Then
                        MassFraction = MassFraction - 0.02
                        b_ReheatHit = False
                    ElseIf ((mRecPClass.ReheatProperties.n_BTUH - mRecPClass.ReheatProperties.n_DesiredBTUH) / mRecPClass.ReheatProperties.n_DesiredBTUH) < (-ReheatDelta) Then
                        MassFraction = MassFraction + 0.02
                        b_ReheatHit = False
                    Else

                        Call mRecPClass.VariationMassFraction(MassFraction)
                        b_ReheatHit = True
                    End If

                    If MassFraction <= 0 Or MassFraction >= 1 Then
                        If MassFraction <= 0 Then
                            MassFraction = 0.005
                            MsgBox("Difference between Reheat temperature and evap leaving temperature must be > 10 degrees", vbOKOnly, "Check Reheat temperature")
                            s_Hit = " (Reheat temperature - LADB) > 10? If Yes Change Fin/Inch, If No Correct it"
                            Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                            Exit Do
                        Else
                            MassFraction = 0.99
                            MsgBox("Difference between Reheat temperature and evap leaving temperature must be < 20 degrees", vbOKOnly, "Check Reheat temperature")
                            s_Hit = " (Reheat temperature - LADB) < 20? If Yes Change Fin/Inch, If No Correct it"
                            Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                            Exit Do
                        End If
                    End If

                    Call mRecPClass.VariationMassFraction(MassFraction)

                    If (MassFraction * 2) > 1 Then
                        MsgBox(" Reheat coil Mass Fraction exceeded 100% of one compressor mass flow")
                        s_Hit = " Reheat Mass Fraction exceeded one compressor mass flow rate"
                        Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                        Exit Do
                    End If

                    HeatRej = (mConPClass.CondProperties.n_BTUH + (mRecPClass.ReheatProperties.n_BTUH * 2)) / (mEvaPClass.EvapProperties.n_BTUH * 2)

                    Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)

                    If b_CondHit And b_EvapHit And b_ReheatHit And (mn_ET < mSysP.n_ETMax) And (mn_CT < mSysP.n_CTMax) And (HeatRej < Frej) Then

                        If (mn_ET < 52.5) And (mn_CT < 130) Then
                            s_Hit = " Balance Unit!"
                        Else
                            s_Hit = " Warning! System Balance at high evap/cond Temperature"
                        End If

                        Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                        Exit Do
                    End If

                Else
                    Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                    HeatRej = (mConPClass.CondProperties.n_BTUH) / (mEvaPClass.EvapProperties.n_BTUH * 2)

                    If b_CondHit And b_EvapHit And (mn_ET < mSysP.n_ETMax) And (mn_CT < mSysP.n_CTMax) And (HeatRej < Frej) Then

                        If (mn_ET < 52.5) And (mn_CT < 130) Then
                            s_Hit = " Balance Unit!"
                        Else
                            s_Hit = " Warning! System Balance at high evap/Cond Temperature"
                        End If

                        Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                        Exit Do
                    End If
                End If

                If i_LoopCounter > 33 Then
                    i_DifficulityCounter = i_DifficulityCounter + 1
                    ReheatDelta = 0.05
                    If i_DifficulityCounter = 5 Then
                        s_Hit = " Could Not Balance!"
                        ReheatDelta = 0.04
                        Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                        Exit Do
                    End If

                End If

            Loop

        End If

    End Sub

    Public Sub StartSys_Balance_ITEFSC_HP()
        Dim mn_ET As Double
        Dim mn_CT As Double
        Dim mn_MassFlow As Double
        Dim mn_CompBtuh As Double
        Dim MassFraction As Double
        Dim n_CompWattsBtuh As Double
        Dim b_EvapHit As Boolean
        Dim b_CondHit As Boolean
        Dim i_DifficulityCounter As Integer
        Dim i_LoopCounter As Integer          'number of times thru loop
        Dim s_Hit As String
        Dim n_TolerancePercent As Double           'tolerance (input)
        Dim n_Percent1 As Double           'evap-comp percentage
        Dim n_Percent2 As Double           'cond-comp+watts percentage
        Dim n_EvapStep As Double           'increment evap temp by (Input by user)
        Dim n_CondStep As Double           'increment evap temp by (Input by user)
        Dim HeatRej As Double
        Dim Subcool As Double
        Dim Superh As Double

        Dim EvACFM As Double
        Dim CoEaDB As Double
        Dim CoEaWB As Double
        Dim CoACFM As Double
        Dim Compquantity As Integer
        Dim EvStepDiv As Double
        Dim CoStepDiv As Double
        Dim nActDensity As Double
        Dim n_EWBlblbSR As Double
        Dim n_IntV1 As Double
        Dim n_ElblbR As Double
        Dim mn_PSIA As Double
        Dim nStdDensity As Double
        Dim mn_EatWB As Double
        Dim mn_EatDB As Double
        Dim Frej As Double
        Const mn_a1 As Double = 0.00080246
        Const mn_a2 As Double = 0.000024525
        Const mn_a3 As Double = 2.542 * (10 ^ -6)
        Const mn_a4 As Double = -2.5855 * (10 ^ -8)
        Const mn_a5 As Double = 4.038 * (10 ^ -10)

        Compquantity = mSysP.i_CompQuantity
        mn_PSIA = maPSIA(mSysP.n_Altitude)
        Frej = 1.43

        If mSysP.n_Altitude > 1999 Then

            mn_EatWB = mSysP.n_EvEaWB
            mn_EatDB = mSysP.n_EvEaDB
            n_EWBlblbSR = mn_a1 + (mn_a2 * mn_EatWB) + (mn_a3 * mn_EatWB ^ 2) + (mn_a4 * mn_EatWB ^ 3) + (mn_a5 * mn_EatWB ^ 4)
            n_IntV1 = ((1093 - 0.556 * mn_EatWB) * n_EWBlblbSR) - (0.24 * (mn_EatDB - mn_EatWB))
            n_ElblbR = n_IntV1 / (1093 + 0.444 * mn_EatDB - mn_EatWB)
            nActDensity = maDensity(mn_EatDB, n_ElblbR, mn_PSIA)
            nStdDensity = maDensity(70, 0, 14.696)
            mSysP.n_EvACFM = mSysP.n_EvCFM / (nActDensity / nStdDensity)
            mSysP.n_OaACFM = mSysP.n_OaCFM / (nActDensity / nStdDensity)

            mn_EatWB = mSysP.n_CoEaWB
            mn_EatDB = mSysP.n_CoEaDB
            n_EWBlblbSR = mn_a1 + (mn_a2 * mn_EatWB) + (mn_a3 * mn_EatWB ^ 2) + (mn_a4 * mn_EatWB ^ 3) + (mn_a5 * mn_EatWB ^ 4)
            n_IntV1 = ((1093 - 0.556 * mn_EatWB) * n_EWBlblbSR) - (0.24 * (mn_EatDB - mn_EatWB))
            n_ElblbR = n_IntV1 / (1093 + 0.444 * mn_EatDB - mn_EatWB)
            nActDensity = maDensity(mn_EatDB, n_ElblbR, mn_PSIA)
            nStdDensity = maDensity(70, 0, 14.696)
            mSysP.n_CoACFM = mSysP.n_CoCFM / (nActDensity / nStdDensity)

        Else

            mSysP.n_EvACFM = mSysP.n_EvCFM
            mSysP.n_OaACFM = mSysP.n_OaCFM
            mSysP.n_CoACFM = mSysP.n_CoCFM

        End If

        If mSysP.n_OaACFM = 0 Then
            CoEaDB = mSysP.n_CoEaDB
            CoEaWB = mSysP.n_CoEaWB
            CoACFM = mSysP.n_CoACFM
        Else
            CoEaDB = mSysP.n_CoEaDB
            CoEaWB = mSysP.n_CoEaWB
            CoACFM = mSysP.n_CoACFM + mSysP.n_OaACFM

            Call MixedAir_HP(CoEaDB, CoEaWB)
            CoEaDB = CoEaDB
            CoEaWB = CoEaWB
        End If

        mSysP.n_MixedDB = CoEaDB
        mSysP.n_MixedWB = CoEaWB
        EvACFM = mSysP.n_EvACFM

        If mSysP.n_EvEaDB >= 40 Then
            EvStepDiv = 0.07 * 3
        ElseIf mSysP.n_EvEaDB >= 30 Then
            EvStepDiv = 0.07 * 3.5
        ElseIf mSysP.n_EvEaDB >= 15 Then
            EvStepDiv = 0.07 * 4
        Else
            EvStepDiv = 0.07 * 5
        End If

        If mSysP.n_EvEaDB >= 30 Then
            CoStepDiv = 0.06
        ElseIf mSysP.n_EvEaDB >= 15 Then
            CoStepDiv = 0.06 * 2
        Else
            CoStepDiv = 0.06 * 3
        End If

        If StartCalculations(Compquantity) Then
            mn_ET = 35
            mn_CT = mSysP.n_MixedDB + 25

            i_LoopCounter = 0
            s_Hit = ""
            n_TolerancePercent = 0.01

            Subcool = mSysP.n_Subcool
            Superh = mSysP.n_Superh

            Call mEvaPClass.PassAirCorrection(mSysP.n_EvEaDB, mSysP.n_EvEaWB, EvACFM, mSysP.s_CompAppCode)
            Call mConPClass.PassMixAirInfo(CoEaDB, CoEaWB, CoACFM, mSysP.s_CompAppCode)

            Do

                Call mComPClass.CompCalc(mn_ET, mn_CT)
                mn_MassFlow = mComPClass.CompProperties.n_MassFlow
                mn_CompBtuh = mComPClass.CompProperties.n_BTUH
                Call mEvaPClass.EvapCalc(mn_CompBtuh, mn_MassFlow, mn_ET, mn_CT)

                MassFraction = 0

                Call mConPClass.CondCalc(mn_MassFlow, MassFraction, mn_ET, mn_CT)

                n_CompWattsBtuh = (mEvaPClass.EvapProperties.n_BTUH) + (mComPClass.CompProperties.n_WattsBTUH * 2 * 3.415)

                If (mEvaPClass.EvapProperties.n_BTUH) >= (mComPClass.CompProperties.n_BTUH * 2) Then
                    n_Percent1 = Abs((mEvaPClass.EvapProperties.n_BTUH) - (mComPClass.CompProperties.n_BTUH * 2)) / (mEvaPClass.EvapProperties.n_BTUH)
                Else
                    n_Percent1 = Abs((mEvaPClass.EvapProperties.n_BTUH) - (mComPClass.CompProperties.n_BTUH * 2)) / (mComPClass.CompProperties.n_BTUH * 2)
                End If

                If (mConPClass.CondProperties.n_BTUH * 2) >= n_CompWattsBtuh Then
                    n_Percent2 = Abs((mConPClass.CondProperties.n_BTUH * 2) - n_CompWattsBtuh) / (mConPClass.CondProperties.n_BTUH * 2)
                Else
                    n_Percent2 = Abs((mConPClass.CondProperties.n_BTUH * 2) - n_CompWattsBtuh) / (n_CompWattsBtuh)
                End If

                If mn_ET = 35 Then
                    Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                End If

                i_LoopCounter = i_LoopCounter + 1
                If n_Percent1 <= n_TolerancePercent Then
                    b_EvapHit = True
                ElseIf (mEvaPClass.EvapProperties.n_BTUH) < (mComPClass.CompProperties.n_BTUH * 2) Then
                    n_EvapStep = (n_Percent1 / EvStepDiv)
                    mn_ET = mn_ET - n_EvapStep
                    If mn_ET < 2 Then
                        mn_ET = 2
                    End If
                    b_EvapHit = False
                Else
                    n_EvapStep = (n_Percent1 / EvStepDiv)
                    mn_ET = mn_ET + n_EvapStep
                    b_EvapHit = False
                End If

                If n_Percent2 <= (n_TolerancePercent + 0.01) Then
                    b_CondHit = True
                ElseIf (mConPClass.CondProperties.n_BTUH * 2) < n_CompWattsBtuh Then
                    n_CondStep = (n_Percent2 / CoStepDiv)
                    mn_CT = mn_CT + n_CondStep
                    If mn_CT > 155 Then
                        mn_CT = 155
                    End If
                    b_CondHit = False
                Else
                    n_CondStep = (n_Percent2 / CoStepDiv)
                    mn_CT = mn_CT - n_CondStep
                    If mn_CT < 75 Then
                        mn_CT = 75
                    End If
                    b_CondHit = False
                End If

                Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                HeatRej = (mConPClass.CondProperties.n_BTUH * 2) / (mEvaPClass.EvapProperties.n_BTUH)

                If b_CondHit And b_EvapHit And (mn_ET < mSysP.n_ETMax) And (mn_CT < mSysP.n_CTMax) And (HeatRej < Frej) Then

                    If (mn_ET > 30) And (mn_ET < 50) And (mn_CT < 130) Then
                        s_Hit = " Balance Unit!"
                    ElseIf (mn_ET > 30) Then
                        s_Hit = " Warning! System Balance at high evap or cond Temperature"
                    ElseIf (mn_ET < 30) Then
                        s_Hit = " Warning! System will require longer defrosting time"
                    End If

                    Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                    Exit Do
                End If

                If i_LoopCounter > 33 Then
                    i_DifficulityCounter = i_DifficulityCounter + 1
                    If i_DifficulityCounter = 5 Then
                        s_Hit = " Could Not Balance!"
                        Call Output(mn_ET, mn_CT, s_Hit, i_LoopCounter)
                        Exit Do
                    End If

                End If

            Loop

        End If
    End Sub

    Public Sub MixedAir_AC(ByRef EvEaDB As Double, ByRef EvEaWB As Double)

        Dim n_OutAirlblbR As Double      'Out Door Air Air lb/lb moisture ratio
        Dim n_OutAirWBlblbSR As Double      'Out Door Air lb/lb saturation moisture ratio at wet bulb
        Dim n_ReturnAirlblbR As Double      'Return air lb/lb moisture ratio
        Dim n_ReturnAirWBlblbSR As Double      'Return air lb/lb saturation moisture ratio at wet bulb
        Dim n_MixedAirlblbR As Double      'Out Door Air Air lb/lb moisture ratio
        Dim n_LaIva As Double       'Intermediate variable
        Dim n_RAirDB As Double      'Return Air DB
        Dim n_RAirWB As Double      'Return Air WB
        Dim n_OAirDB As Double      'Outdoor Air DB
        Dim n_OAirWB As Double      'Outdoor Air WB
        Dim n_AverageWB As Double      'Average Wet Bulb use to calculate lb/lb ratio at wet bulb
        Dim n_AverageWBlblbSR As Double      'Average lb/lb saturation moisture ratio at average WB
        Dim n_IntV1 As Double      'Intermidiate Variable
        Dim n_OutAirEnthalphy As Double      'Outdoor Air Enthalphy
        Dim n_RetAirEnthalphy As Double      'Return Air Enthalphy
        Dim n_MixAirEnthalphy As Double      'Mix Air Enthalphy
        Const mn_a1 As Double = 0.00080246
        Const mn_a2 As Double = 0.000024525
        Const mn_a3 As Double = 2.542 * (10 ^ -6)
        Const mn_a4 As Double = -2.5855 * (10 ^ -8)
        Const mn_a5 As Double = 4.038 * (10 ^ -10)
        'Const mn_Pat As Double = 14.697

        n_RAirDB = mSysP.n_EvEaDB
        n_RAirWB = mSysP.n_EvEaWB
        n_OAirDB = mSysP.n_CoEaDB
        n_OAirWB = mSysP.n_CoEaWB

        n_ReturnAirWBlblbSR = mn_a1 + (mn_a2 * n_RAirWB) + (mn_a3 * n_RAirWB ^ 2) + (mn_a4 * n_RAirWB ^ 3) + (mn_a5 * n_RAirWB ^ 4)
        n_IntV1 = ((1093 - 0.556 * n_RAirWB) * n_ReturnAirWBlblbSR) - (0.24 * (n_RAirDB - n_RAirWB))
        n_ReturnAirlblbR = n_IntV1 / (1093 + (0.444 * n_RAirDB) - n_RAirWB)

        n_OutAirWBlblbSR = mn_a1 + (mn_a2 * n_OAirWB) + (mn_a3 * n_OAirWB ^ 2) + (mn_a4 * n_OAirWB ^ 3) + (mn_a5 * n_OAirWB ^ 4)
        n_IntV1 = ((1093 - 0.556 * n_OAirWB) * n_OutAirWBlblbSR) - (0.24 * (n_OAirDB - n_OAirWB))
        n_OutAirlblbR = n_IntV1 / (1093 + (0.444 * n_OAirDB) - n_OAirWB)

        n_MixedAirlblbR = (n_OutAirlblbR * mSysP.n_OaACFM + n_ReturnAirlblbR * mSysP.n_EvACFM) / (mSysP.n_OaACFM + mSysP.n_EvACFM)
        n_RetAirEnthalphy = (0.24 * mSysP.n_EvEaDB) + (1061 + 0.444 * mSysP.n_EvEaDB) * n_ReturnAirlblbR
        n_OutAirEnthalphy = (0.24 * mSysP.n_CoEaDB) + (1061 + 0.444 * mSysP.n_CoEaDB) * n_OutAirlblbR
        n_MixAirEnthalphy = (n_RetAirEnthalphy * mSysP.n_EvACFM + n_OutAirEnthalphy * mSysP.n_OaACFM) / (mSysP.n_EvACFM + mSysP.n_OaACFM)

        EvEaDB = (n_MixAirEnthalphy - 1061 * n_MixedAirlblbR) / (0.24 + 0.444 * n_MixedAirlblbR)

        n_AverageWB = (mSysP.n_EvEaWB * mSysP.n_EvACFM + mSysP.n_CoEaWB * mSysP.n_OaACFM) / (mSysP.n_EvACFM + mSysP.n_OaACFM)
        n_AverageWBlblbSR = mn_a1 + (mn_a2 * n_AverageWB) + (mn_a3 * n_AverageWB ^ 2) + (mn_a4 * n_AverageWB ^ 3) + (mn_a5 * n_AverageWB ^ 4)

        n_LaIva = (0.444 * EvEaDB + 1093) * n_MixedAirlblbR
        EvEaWB = (n_LaIva + (0.24 * EvEaDB) - (1093 * n_AverageWBlblbSR)) / (0.24 - (0.556 * n_AverageWBlblbSR) + n_MixedAirlblbR)

    End Sub

    Public Sub MixedAir_HP(ByRef CoEaDB As Double, ByRef CoEaWB As Double)

        Dim n_OutAirlblbR As Double      'Out Door Air Air lb/lb moisture ratio
        Dim n_OutAirWBlblbSR As Double      'Out Door Air lb/lb saturation moisture ratio at wet bulb
        Dim n_ReturnAirlblbR As Double      'Return air lb/lb moisture ratio
        Dim n_ReturnAirWBlblbSR As Double      'Return air lb/lb saturation moisture ratio at wet bulb
        Dim n_MixedAirlblbR As Double      'Out Door Air Air lb/lb moisture ratio
        Dim n_LaIva As Double       'Intermediate variable
        Dim n_RAirDB As Double      'Return Air DB
        Dim n_RAirWB As Double      'Return Air WB
        Dim n_OAirDB As Double      'Outdoor Air DB
        Dim n_OAirWB As Double      'Outdoor Air WB
        Dim n_AverageWB As Double      'Average Wet Bulb use to calculate lb/lb ratio at wet bulb
        Dim n_AverageWBlblbSR As Double      'Average lb/lb saturation moisture ratio at average WB
        Dim n_IntV1 As Double      'Intermidiate Variable
        Dim n_OutAirEnthalphy As Double      'Outdoor Air Enthalphy
        Dim n_RetAirEnthalphy As Double      'Return Air Enthalphy
        Dim n_MixAirEnthalphy As Double      'Mix Air Enthalphy
        Const mn_a1 As Double = 0.00080246
        Const mn_a2 As Double = 0.000024525
        Const mn_a3 As Double = 2.542 * (10 ^ -6)
        Const mn_a4 As Double = -2.5855 * (10 ^ -8)
        Const mn_a5 As Double = 4.038 * (10 ^ -10)
        'Const mn_Pat As Double = 14.697

        n_RAirDB = mSysP.n_CoEaDB
        n_RAirWB = mSysP.n_CoEaWB
        n_OAirDB = mSysP.n_EvEaDB
        n_OAirWB = mSysP.n_EvEaWB

        n_ReturnAirWBlblbSR = mn_a1 + (mn_a2 * n_RAirWB) + (mn_a3 * n_RAirWB ^ 2) + (mn_a4 * n_RAirWB ^ 3) + (mn_a5 * n_RAirWB ^ 4)
        n_IntV1 = ((1093 - 0.556 * n_RAirWB) * n_ReturnAirWBlblbSR) - (0.24 * (n_RAirDB - n_RAirWB))
        n_ReturnAirlblbR = n_IntV1 / (1093 + (0.444 * n_RAirDB) - n_RAirWB)

        n_OutAirWBlblbSR = mn_a1 + (mn_a2 * n_OAirWB) + (mn_a3 * n_OAirWB ^ 2) + (mn_a4 * n_OAirWB ^ 3) + (mn_a5 * n_OAirWB ^ 4)
        n_IntV1 = ((1093 - 0.556 * n_OAirWB) * n_OutAirWBlblbSR) - (0.24 * (n_OAirDB - n_OAirWB))
        n_OutAirlblbR = n_IntV1 / (1093 + (0.444 * n_OAirDB) - n_OAirWB)

        n_MixedAirlblbR = (n_OutAirlblbR * mSysP.n_OaACFM + n_ReturnAirlblbR * mSysP.n_CoACFM) / (mSysP.n_OaACFM + mSysP.n_CoACFM)
        n_RetAirEnthalphy = (0.24 * mSysP.n_CoEaDB) + (1061 + 0.444 * mSysP.n_CoEaDB) * n_ReturnAirlblbR
        n_OutAirEnthalphy = (0.24 * mSysP.n_EvEaDB) + (1061 + 0.444 * mSysP.n_EvEaDB) * n_OutAirlblbR
        n_MixAirEnthalphy = (n_RetAirEnthalphy * mSysP.n_CoACFM + n_OutAirEnthalphy * mSysP.n_OaACFM) / (mSysP.n_CoACFM + mSysP.n_OaACFM)

        CoEaDB = (n_MixAirEnthalphy - 1061 * n_MixedAirlblbR) / (0.24 + 0.444 * n_MixedAirlblbR)

        n_AverageWB = (mSysP.n_CoEaWB * mSysP.n_CoACFM + mSysP.n_CoEaWB * mSysP.n_OaACFM) / (mSysP.n_CoACFM + mSysP.n_OaACFM)
        n_AverageWBlblbSR = mn_a1 + (mn_a2 * n_AverageWB) + (mn_a3 * n_AverageWB ^ 2) + (mn_a4 * n_AverageWB ^ 3) + (mn_a5 * n_AverageWB ^ 4)

        n_LaIva = (0.444 * CoEaDB + 1093) * n_MixedAirlblbR
        CoEaWB = (n_LaIva + (0.24 * CoEaDB) - (1093 * n_AverageWBlblbSR)) / (0.24 - (0.556 * n_AverageWBlblbSR) + n_MixedAirlblbR)

    End Sub

    Private Function maDensity(ByVal mn_EatDB As Double, ByVal n_ElblbR As Double, ByVal mn_PSIA As Double) As Double

        Dim mn_PinHg As Double
        Dim SpeVolume As Double

        mn_PinHg = mn_PSIA * 29.92 / 14.696

        SpeVolume = 0.7543 * (mn_EatDB + 459.67) * (1 + 1.6078 * n_ElblbR) / mn_PinHg

        maDensity = 1 / SpeVolume
    End Function
    Private Function maEnthalpy(ByVal mn_EatDB As Double, ByVal n_ElblbR As Double) As Double
        maEnthalpy = (0.24 * mn_EatDB) + (n_ElblbR * (1061 + 0.444 * mn_EatDB))
    End Function
    Private Function maDewPoint(ByVal n_EWVPre As Double) As Double
        maDewPoint = 100.45 + (33.193 * Log(n_EWVPre)) + (2.319 * (Log(n_EWVPre)) ^ 2) + (0.1704 * (Log(n_EWVPre)) ^ 3) + (1.2063 * (n_EWVPre) ^ 0.1984)
    End Function
    Private Function maSatHumidityR(ByVal mn_EatDB As Double) As Double
        Const mn_a1 As Double = 0.00080246
        Const mn_a2 As Double = 0.000024525
        Const mn_a3 As Double = 2.542 * (10 ^ -6)
        Const mn_a4 As Double = -2.5855 * (10 ^ -8)
        Const mn_a5 As Double = 4.038 * (10 ^ -10)

        maSatHumidityR = mn_a1 + (mn_a2 * mn_EatDB) + (mn_a3 * mn_EatDB ^ 2) + (mn_a4 * mn_EatDB ^ 3) + (mn_a5 * mn_EatDB ^ 4)
    End Function

    Private Function maPSIA(ByVal Altitude As Double) As Double

        Dim IntV1 As Double
        Dim PA As Double

        Altitude = Altitude * 0.3048

        IntV1 = 9.80665 * 0.0289644 / (8.31432 * 0.0065)

        PA = (1 - (0.0065 * Altitude / 288.15)) ^ IntV1

        maPSIA = 14.696 * PA


    End Function

    Public Sub LA_MixedAir_AC(ByRef EvLaDB As Double, ByRef EvLaWB As Double)

        Dim n_EAirlblbR As Double      'Out Door Air Air lb/lb moisture ratio
        Dim n_EAirWBlblbSR As Double      'Out Door Air lb/lb saturation moisture ratio at wet bulb
        Dim n_LAirlblbR As Double      'Return air lb/lb moisture ratio
        Dim n_LAirWBlblbSR As Double      'Return air lb/lb saturation moisture ratio at wet bulb
        Dim n_MixedAirlblbR As Double      'Out Door Air Air lb/lb moisture ratio
        Dim n_LaIva As Double       'Intermediate variable
        Dim n_LAirDB As Double      'Return Air DB
        Dim n_LAirWB As Double      'Return Air WB
        Dim n_EAirDB As Double      'Outdoor Air DB
        Dim n_EAirWB As Double      'Outdoor Air WB
        Dim n_AverageWB As Double      'Average Wet Bulb use to calculate lb/lb ratio at wet bulb
        Dim n_AverageWBlblbSR As Double      'Average lb/lb saturation moisture ratio at average WB
        Dim n_IntV1 As Double      'Intermidiate Variable
        Dim n_EAirEnthalphy As Double      'Outdoor Air Enthalphy
        Dim n_LAirEnthalphy As Double      'Return Air Enthalphy
        Dim n_MixAirEnthalphy As Double      'Mix Air Enthalphy
        Dim n_ACFM As Double
        Const mn_a1 As Double = 0.00080246
        Const mn_a2 As Double = 0.000024525
        Const mn_a3 As Double = 2.542 * (10 ^ -6)
        Const mn_a4 As Double = -2.5855 * (10 ^ -8)
        Const mn_a5 As Double = 4.038 * (10 ^ -10)
        'Const mn_Pat As Double = 14.697

        n_LAirDB = EvLaDB
        n_LAirWB = EvLaWB
        n_EAirDB = mSysP.n_EvEaDB
        n_EAirWB = mSysP.n_EvEaWB

        If mSysP.i_CompQuantity = 2 Then
            If mComPClass.CompProperties.s_AppCode = "AH" Or mComPClass.CompProperties.s_AppCode = "A5" Then
                n_ACFM = mSysP.n_EvACFM * 0.5
            ElseIf mComPClass.CompProperties.s_AppCode = "A6" Then
                n_ACFM = mSysP.n_EvACFM * 0.66
            ElseIf mComPClass.CompProperties.s_AppCode = "AL" Then
                n_ACFM = mSysP.n_EvACFM * 0.55
            ElseIf mComPClass.CompProperties.s_AppCode = "AS" Then
                n_ACFM = mSysP.n_EvACFM * 0.45
            End If
        Else
            n_ACFM = mSysP.n_EvACFM * 0.5
        End If

        n_LAirWBlblbSR = mn_a1 + (mn_a2 * n_LAirWB) + (mn_a3 * n_LAirWB ^ 2) + (mn_a4 * n_LAirWB ^ 3) + (mn_a5 * n_LAirWB ^ 4)
        n_IntV1 = ((1093 - 0.556 * n_LAirWB) * n_LAirWBlblbSR) - (0.24 * (n_LAirDB - n_LAirWB))
        n_LAirlblbR = n_IntV1 / (1093 + (0.444 * n_LAirDB) - n_LAirWB)

        n_EAirWBlblbSR = mn_a1 + (mn_a2 * n_EAirWB) + (mn_a3 * n_EAirWB ^ 2) + (mn_a4 * n_EAirWB ^ 3) + (mn_a5 * n_EAirWB ^ 4)
        n_IntV1 = ((1093 - 0.556 * n_EAirWB) * n_EAirWBlblbSR) - (0.24 * (n_EAirDB - n_EAirWB))
        n_EAirlblbR = n_IntV1 / (1093 + (0.444 * n_EAirDB) - n_EAirWB)

        n_MixedAirlblbR = (n_EAirlblbR * n_ACFM + n_LAirlblbR * n_ACFM) / (mSysP.n_EvACFM)
        n_LAirEnthalphy = (0.24 * n_LAirDB) + (1061 + 0.444 * n_LAirDB) * n_LAirlblbR
        n_EAirEnthalphy = (0.24 * n_EAirDB) + (1061 + 0.444 * n_EAirDB) * n_EAirlblbR
        n_MixAirEnthalphy = (n_LAirEnthalphy * n_ACFM + n_EAirEnthalphy * n_ACFM) / (mSysP.n_EvACFM)

        EvLaDB = (n_MixAirEnthalphy - 1061 * n_MixedAirlblbR) / (0.24 + 0.444 * n_MixedAirlblbR)

        n_AverageWB = (n_LAirWB * n_ACFM + n_EAirWB * n_ACFM) / (mSysP.n_EvACFM)
        n_AverageWBlblbSR = mn_a1 + (mn_a2 * n_AverageWB) + (mn_a3 * n_AverageWB ^ 2) + (mn_a4 * n_AverageWB ^ 3) + (mn_a5 * n_AverageWB ^ 4)

        n_LaIva = (0.444 * EvLaDB + 1093) * n_MixedAirlblbR
        EvLaWB = (n_LaIva + (0.24 * EvLaDB) - (1093 * n_AverageWBlblbSR)) / (0.24 - (0.556 * n_AverageWBlblbSR) + n_MixedAirlblbR)

    End Sub

    Public Sub LA_MixedAir_HP(ByRef CoLaDB As Double, ByRef CoLaWB As Double)

        Dim n_EAirlblbR As Double      'Out Door Air Air lb/lb moisture ratio
        Dim n_EAirWBlblbSR As Double      'Out Door Air lb/lb saturation moisture ratio at wet bulb
        Dim n_LAirlblbR As Double      'Return air lb/lb moisture ratio
        Dim n_LAirDBlblbSR As Double      'Return air lb/lb saturation moisture ratio at wet bulb
        Dim n_LSatWVPre As Double
        Dim n_LWVPre As Double
        Dim n_MixedAirlblbR As Double      'Out Door Air Air lb/lb moisture ratio
        'Dim n_LaIva As Double       'Intermediate variable
        Dim n_LAirDB As Double      'Return Air DB
        Dim n_LAirHumidity As Double      'Return Air WB
        Dim n_EAirDB As Double      'Outdoor Air DB
        Dim n_EAirWB As Double      'Outdoor Air WB
        '    Dim n_AverageDB                     As Double
        Dim n_MixDBlblbSR As Double      'Average lb/lb saturation moisture ratio at average DB
        Dim n_IntV1 As Double      'Intermidiate Variable
        Dim n_EAirEnthalphy As Double      'Outdoor Air Enthalphy
        Dim n_LAirEnthalphy As Double      'Return Air Enthalphy
        Dim n_MixAirEnthalphy As Double      'Mix Air Enthalphy
        Dim n_MixSatWVPre As Double
        Dim n_MixSaturation As Double
        Dim n_ACFM As Double
        Const mn_a1 As Double = 0.00080246
        Const mn_a2 As Double = 0.000024525
        Const mn_a3 As Double = 2.542 * (10 ^ -6)
        Const mn_a4 As Double = -2.5855 * (10 ^ -8)
        Const mn_a5 As Double = 4.038 * (10 ^ -10)
        Const mn_Pat As Double = 14.697

        n_LAirDB = CoLaDB
        n_LAirHumidity = CoLaWB
        n_EAirDB = mSysP.n_CoEaDB
        n_EAirWB = mSysP.n_CoEaWB

        If mSysP.i_CompQuantity = 2 Then
            If mComPClass.CompProperties.s_AppCode = "AH" Or mComPClass.CompProperties.s_AppCode = "A5" Then
                n_ACFM = mSysP.n_CoACFM * 0.5
            ElseIf mComPClass.CompProperties.s_AppCode = "A6" Then
                n_ACFM = mSysP.n_CoACFM * 0.66
            ElseIf mComPClass.CompProperties.s_AppCode = "AL" Then
                n_ACFM = mSysP.n_CoACFM * 0.55
            ElseIf mComPClass.CompProperties.s_AppCode = "AS" Then
                n_ACFM = mSysP.n_CoACFM * 0.45
            End If
        Else
            n_ACFM = mSysP.n_CoACFM * 0.5
        End If

        n_LAirDBlblbSR = mn_a1 + (mn_a2 * n_LAirDB) + (mn_a3 * n_LAirDB ^ 2) + (mn_a4 * n_LAirDB ^ 3) + (mn_a5 * n_LAirDB ^ 4)
        n_LSatWVPre = (mn_Pat * n_LAirDBlblbSR) / (n_LAirDBlblbSR + 0.62198)
        n_LWVPre = n_LSatWVPre * n_LAirHumidity
        n_LAirlblbR = (0.62198 * n_LWVPre) / (mn_Pat - n_LWVPre)

        n_EAirWBlblbSR = mn_a1 + (mn_a2 * n_EAirWB) + (mn_a3 * n_EAirWB ^ 2) + (mn_a4 * n_EAirWB ^ 3) + (mn_a5 * n_EAirWB ^ 4)
        n_IntV1 = ((1093 - 0.556 * n_EAirWB) * n_EAirWBlblbSR) - (0.24 * (n_EAirDB - n_EAirWB))
        n_EAirlblbR = n_IntV1 / (1093 + (0.444 * n_EAirDB) - n_EAirWB)

        n_MixedAirlblbR = (n_EAirlblbR * n_ACFM + n_LAirlblbR * n_ACFM) / (mSysP.n_CoACFM)
        n_LAirEnthalphy = (0.24 * n_LAirDB) + (1061 + 0.444 * n_LAirDB) * n_LAirlblbR
        n_EAirEnthalphy = (0.24 * n_EAirDB) + (1061 + 0.444 * n_EAirDB) * n_EAirlblbR
        n_MixAirEnthalphy = (n_LAirEnthalphy * n_ACFM + n_EAirEnthalphy * n_ACFM) / (mSysP.n_CoACFM)

        CoLaDB = (n_MixAirEnthalphy - 1061 * n_MixedAirlblbR) / (0.24 + 0.444 * n_MixedAirlblbR)

        '    n_AverageDB = (n_LAirDB * n_ACFM + n_EAirDB * n_ACFM) / (mSysP.n_CoACFM)
        n_MixDBlblbSR = mn_a1 + (mn_a2 * CoLaDB) + (mn_a3 * CoLaDB ^ 2) + (mn_a4 * CoLaDB ^ 3) + (mn_a5 * CoLaDB ^ 4)
        n_MixSatWVPre = (mn_Pat * n_MixDBlblbSR) / (n_MixDBlblbSR + 0.62198)
        n_MixSaturation = n_MixedAirlblbR / n_MixDBlblbSR

        CoLaWB = n_MixSaturation / (1 - (1 - n_MixSaturation) * (n_MixSatWVPre / mn_Pat))

    End Sub

    Public Sub Output(ByVal mn_ET As Double, ByVal mn_CT As Double, ByVal s_Hit As String, ByVal i_LoopCounter As Integer)


        If s_Hit = "" Then
            Call CreateRecords(mn_ET, mn_CT, i_LoopCounter)
        Else
            mSysP.s_Hit = Left(s_Hit, 85)
        End If

    End Sub

    Public Sub CreateRecords(ByVal mn_ET As Double, ByVal mn_CT As Double, ByVal i_LoopCounter As Integer)

        Dim EvLaDB As Double
        Dim EvLaWB As Double
        Dim CoLaDB As Double
        Dim CoLaWB As Double
        Dim PartTandem As Boolean

        PartTandem = False

        If mComPClass.CompProperties.s_AppCode = "AH" Or mComPClass.CompProperties.s_AppCode = "A5" Or _
           mComPClass.CompProperties.s_AppCode = "A6" Or mComPClass.CompProperties.s_AppCode = "AL" Or _
           mComPClass.CompProperties.s_AppCode = "AS" Then

            PartTandem = True
        End If

        If i_LoopCounter = 0 Then

            'dtItems.Columns.Clear()
            If dtItems.Columns.Count = 0 Then

                With dtItems.Columns
                    .Add("Iter", GetType(Integer))
                    .Add("EVT", GetType(Single))
                    .Add("COT", GetType(Single))
                    .Add("TotalMB", GetType(Single))
                    .Add("SensMB", GetType(Single))
                    .Add("CondMB", GetType(Single))
                    .Add("CompMB", GetType(Single))
                    .Add("MassF", GetType(Single))
                    .Add("LADB", GetType(Single))
                    .Add("LAWB", GetType(Object))
                    .Add("EvPD", GetType(Single))
                    .Add("CoPD", GetType(Single))
                    .Add("RHReq", GetType(Single))
                    .Add("RHOut", GetType(Single))
                    .Add("RHMF%", GetType(Object))
                End With

            End If

            If dsItems.Tables.Count = 0 Then
                dsItems.Tables.Add(dtItems)
                dsItems.Tables(0).PrimaryKey = New DataColumn() {dsItems.Tables(0).Columns("Iter")}
                Call WriteSchemaToFile(dsItems)
            End If

        End If

        If i_LoopCounter <> 0 Then

            Dim drItems As DataRow = dtItems.NewRow()

            If mSysP.i_CompQuantity = 2 And mSysP.i_CondQuantity = 2 And mEvaPClass.EvapProperties.s_Split = "IT" Then
                drItems("Iter") = i_LoopCounter
                drItems("EVT") = Round(mn_ET, 2)
                drItems("COT") = Round(mn_CT, 2)
                drItems("TotalMB") = Round(mEvaPClass.EvapProperties.n_BTUH / 1000, 2)
                drItems("SensMB") = Round(mEvaPClass.EvapProperties.n_Sensible / 1000, 2)
                drItems("CondMB") = Round((2 * (mConPClass.CondProperties.n_BTUH + mRecPClass.ReheatProperties.n_BTUH)) / 1000, 2)
                drItems("CompMB") = Round(mComPClass.CompProperties.n_BTUH * 2 / 1000, 2)
                drItems("MassF") = Round(mComPClass.CompProperties.n_MassFlow * 2, 1)
                drItems("EvPD") = Round(mEvaPClass.EvapProperties.n_EvapPD, 2)
                drItems("CoPD") = Round(mConPClass.CondProperties.n_CondPD, 2)
                drItems("RHReq") = Round(2 * mRecPClass.ReheatProperties.n_DesiredBTUH / 1000, 2)
                drItems("RHOut") = Round(2 * mRecPClass.ReheatProperties.n_BTUH / 1000, 2)
                drItems("RHMF%") = FormatPercent(mRecPClass.ReheatProperties.n_MassFraction, 1)

            ElseIf mSysP.i_CompQuantity = 2 And mConPClass.CondProperties.s_Split = "IT" And mEvaPClass.EvapProperties.s_Split = "IT" Then
                drItems("Iter") = i_LoopCounter
                drItems("EVT") = Round(mn_ET, 2)
                drItems("COT") = Round(mn_CT, 2)
                drItems("TotalMB") = Round(mEvaPClass.EvapProperties.n_BTUH / 1000, 2)
                drItems("SensMB") = Round(mEvaPClass.EvapProperties.n_Sensible / 1000, 2)
                drItems("CondMB") = Round((mConPClass.CondProperties.n_BTUH + (mRecPClass.ReheatProperties.n_BTUH * 2)) / 1000, 2)
                drItems("CompMB") = Round(mComPClass.CompProperties.n_BTUH * 2 / 1000, 2)
                drItems("MassF") = Round(mComPClass.CompProperties.n_MassFlow * 2, 1)
                drItems("EvPD") = Round(mEvaPClass.EvapProperties.n_EvapPD, 2)
                drItems("CoPD") = Round(mConPClass.CondProperties.n_CondPD, 2)
                drItems("RHReq") = Round(2 * mRecPClass.ReheatProperties.n_DesiredBTUH / 1000, 2)
                drItems("RHOut") = Round(2 * mRecPClass.ReheatProperties.n_BTUH / 1000, 2)
                drItems("RHMF%") = FormatPercent(mRecPClass.ReheatProperties.n_MassFraction, 1)

            ElseIf mSysP.i_CompQuantity = 2 And mEvaPClass.EvapProperties.s_Split = "FS" And mConPClass.CondProperties.s_Split <> "IT" Then
                drItems("Iter") = i_LoopCounter
                drItems("EVT") = Round(mn_ET, 2)
                drItems("COT") = Round(mn_CT, 2)
                drItems("TotalMB") = Round(2 * mEvaPClass.EvapProperties.n_BTUH / 1000, 2)
                drItems("SensMB") = Round(2 * mEvaPClass.EvapProperties.n_Sensible / 1000, 2)
                drItems("CondMB") = Round((2 * (mConPClass.CondProperties.n_BTUH + mRecPClass.ReheatProperties.n_BTUH)) / 1000, 2)
                drItems("CompMB") = Round(2 * mComPClass.CompProperties.n_BTUH / 1000, 2)
                drItems("MassF") = Round(mComPClass.CompProperties.n_MassFlow * 2, 1)
                drItems("EvPD") = Round(mEvaPClass.EvapProperties.n_EvapPD, 2)
                drItems("CoPD") = Round(mConPClass.CondProperties.n_CondPD, 2)
                drItems("RHReq") = Round(2 * mRecPClass.ReheatProperties.n_DesiredBTUH / 1000, 2)
                drItems("RHOut") = Round(2 * mRecPClass.ReheatProperties.n_BTUH / 1000, 2)
                drItems("RHMF%") = FormatPercent(mRecPClass.ReheatProperties.n_MassFraction, 1)

            ElseIf mSysP.i_CompQuantity = 2 And mEvaPClass.EvapProperties.s_Split = "FS" And mConPClass.CondProperties.s_Split = "IT" Then
                drItems("Iter") = i_LoopCounter
                drItems("EVT") = Round(mn_ET, 2)
                drItems("COT") = Round(mn_CT, 2)
                drItems("TotalMB") = Round(2 * mEvaPClass.EvapProperties.n_BTUH / 1000, 2)
                drItems("SensMB") = Round(2 * mEvaPClass.EvapProperties.n_Sensible / 1000, 2)
                drItems("CondMB") = Round((mConPClass.CondProperties.n_BTUH + (2 * mRecPClass.ReheatProperties.n_BTUH)) / 1000, 2)
                drItems("CompMB") = Round(2 * mComPClass.CompProperties.n_BTUH / 1000, 2)
                drItems("MassF") = Round(mComPClass.CompProperties.n_MassFlow * 2, 1)
                drItems("EvPD") = Round(mEvaPClass.EvapProperties.n_EvapPD, 2)
                drItems("CoPD") = Round(mConPClass.CondProperties.n_CondPD, 2)
                drItems("RHReq") = Round(2 * mRecPClass.ReheatProperties.n_DesiredBTUH / 1000, 2)
                drItems("RHOut") = Round(2 * mRecPClass.ReheatProperties.n_BTUH / 1000, 2)
                drItems("RHMF%") = FormatPercent(mRecPClass.ReheatProperties.n_MassFraction, 1)

            ElseIf mSysP.i_CompQuantity = 2 And mSysP.i_EvapQuantity = 2 And mConPClass.CondProperties.s_Split = "IT" Then
                drItems("Iter") = i_LoopCounter
                drItems("EVT") = Round(mn_ET, 2)
                drItems("COT") = Round(mn_CT, 2)
                drItems("TotalMB") = Round(mEvaPClass.EvapProperties.n_BTUH * 2 / 1000, 2)
                drItems("SensMB") = Round(mEvaPClass.EvapProperties.n_Sensible * 2 / 1000, 2)
                drItems("CondMB") = Round((mConPClass.CondProperties.n_BTUH + (2 * mRecPClass.ReheatProperties.n_BTUH)) / 1000, 2)
                drItems("CompMB") = Round(mComPClass.CompProperties.n_BTUH * 2 / 1000, 2)
                drItems("MassF") = Round(mComPClass.CompProperties.n_MassFlow * 2, 1)
                drItems("EvPD") = Round(mEvaPClass.EvapProperties.n_EvapPD, 2)
                drItems("CoPD") = Round(mConPClass.CondProperties.n_CondPD, 2)
                drItems("RHReq") = Round(2 * mRecPClass.ReheatProperties.n_DesiredBTUH / 1000, 2)
                drItems("RHOut") = Round(2 * mRecPClass.ReheatProperties.n_BTUH / 1000, 2)
                drItems("RHMF%") = FormatPercent(mRecPClass.ReheatProperties.n_MassFraction, 1)

            ElseIf mSysP.i_CompQuantity = 2 And mConPClass.CondProperties.s_Split = "FS" And mEvaPClass.EvapProperties.s_Split <> "IT" Then
                drItems("Iter") = i_LoopCounter
                drItems("EVT") = Round(mn_ET, 2)
                drItems("COT") = Round(mn_CT, 2)
                drItems("TotalMB") = Round(2 * mEvaPClass.EvapProperties.n_BTUH / 1000, 2)
                drItems("SensMB") = Round(2 * mEvaPClass.EvapProperties.n_Sensible / 1000, 2)
                drItems("CondMB") = Round((2 * (mConPClass.CondProperties.n_BTUH + mRecPClass.ReheatProperties.n_BTUH)) / 1000, 2)
                drItems("CompMB") = Round(2 * mComPClass.CompProperties.n_BTUH / 1000, 2)
                drItems("MassF") = Round(mComPClass.CompProperties.n_MassFlow * 2, 1)
                drItems("EvPD") = Round(mEvaPClass.EvapProperties.n_EvapPD, 2)
                drItems("CoPD") = Round(mConPClass.CondProperties.n_CondPD, 2)
                drItems("RHReq") = Round(2 * mRecPClass.ReheatProperties.n_DesiredBTUH / 1000, 2)
                drItems("RHOut") = Round(2 * mRecPClass.ReheatProperties.n_BTUH / 1000, 2)
                drItems("RHMF%") = FormatPercent(mRecPClass.ReheatProperties.n_MassFraction, 1)

            ElseIf mSysP.i_CompQuantity = 2 And mConPClass.CondProperties.s_Split = "FS" And mEvaPClass.EvapProperties.s_Split = "IT" Then
                drItems("Iter") = i_LoopCounter
                drItems("EVT") = Round(mn_ET, 2)
                drItems("COT") = Round(mn_CT, 2)
                drItems("TotalMB") = Round(mEvaPClass.EvapProperties.n_BTUH / 1000, 2)
                drItems("SensMB") = Round(mEvaPClass.EvapProperties.n_Sensible / 1000, 2)
                drItems("CondMB") = Round((2 * (mConPClass.CondProperties.n_BTUH + mRecPClass.ReheatProperties.n_BTUH)) / 1000, 2)
                drItems("CompMB") = Round(2 * mComPClass.CompProperties.n_BTUH / 1000, 2)
                drItems("MassF") = Round(mComPClass.CompProperties.n_MassFlow * 2, 1)
                drItems("EvPD") = Round(mEvaPClass.EvapProperties.n_EvapPD, 2)
                drItems("CoPD") = Round(mConPClass.CondProperties.n_CondPD, 2)
                drItems("RHReq") = Round(2 * mRecPClass.ReheatProperties.n_DesiredBTUH / 1000, 2)
                drItems("RHOut") = Round(2 * mRecPClass.ReheatProperties.n_BTUH / 1000, 2)
                drItems("RHMF%") = FormatPercent(mRecPClass.ReheatProperties.n_MassFraction, 1)

            ElseIf mSysP.i_CompQuantity = 2 And mConPClass.CondProperties.s_Split = "NS" And mEvaPClass.EvapProperties.s_Split = "NS" Then
                drItems("Iter") = i_LoopCounter
                drItems("EVT") = Round(mn_ET, 2)
                drItems("COT") = Round(mn_CT, 2)
                drItems("TotalMB") = Round(2 * mEvaPClass.EvapProperties.n_BTUH / 1000, 2)
                drItems("SensMB") = Round(2 * mEvaPClass.EvapProperties.n_Sensible / 1000, 2)
                drItems("CondMB") = Round((2 * (mConPClass.CondProperties.n_BTUH + mRecPClass.ReheatProperties.n_BTUH)) / 1000, 2)
                drItems("CompMB") = Round(2 * mComPClass.CompProperties.n_BTUH / 1000, 2)
                drItems("MassF") = Round(2 * mComPClass.CompProperties.n_MassFlow, 1)
                drItems("EvPD") = Round(mEvaPClass.EvapProperties.n_EvapPD, 2)
                drItems("CoPD") = Round(mConPClass.CondProperties.n_CondPD, 2)
                drItems("RHReq") = Round((2 * mRecPClass.ReheatProperties.n_DesiredBTUH) / 1000, 2)
                drItems("RHOut") = Round((2 * mRecPClass.ReheatProperties.n_BTUH) / 1000, 2)
                drItems("RHMF%") = FormatPercent((2 * mRecPClass.ReheatProperties.n_MassFraction), 1)

            ElseIf mSysP.i_CompQuantity = 1 Then
                drItems("Iter") = i_LoopCounter
                drItems("EVT") = Round(mn_ET, 2)
                drItems("COT") = Round(mn_CT, 2)
                drItems("TotalMB") = Round(mEvaPClass.EvapProperties.n_BTUH / 1000, 2)
                drItems("SensMB") = Round(mEvaPClass.EvapProperties.n_Sensible / 1000, 2)
                drItems("CondMB") = Round((mConPClass.CondProperties.n_BTUH + mRecPClass.ReheatProperties.n_BTUH) / 1000, 2)
                drItems("CompMB") = Round(mComPClass.CompProperties.n_BTUH / 1000, 1)
                drItems("MassF") = Round(mComPClass.CompProperties.n_MassFlow, 1)
                drItems("EvPD") = Round(mEvaPClass.EvapProperties.n_EvapPD, 2)
                drItems("CoPD") = Round(mConPClass.CondProperties.n_CondPD, 2)
                drItems("RHReq") = Round(mRecPClass.ReheatProperties.n_DesiredBTUH / 1000, 2)
                drItems("RHOut") = Round(mRecPClass.ReheatProperties.n_BTUH / 1000, 2)
                drItems("RHMF%") = FormatPercent(mRecPClass.ReheatProperties.n_MassFraction, 1)

            End If

            If mSysP.i_CompQuantity = 2 Then

                If frmUnitInputs.optHeatPump.Checked = True Then
                    If PartTandem = True Then
                        CoLaDB = mConPClass.CondProperties.n_LaDB
                        CoLaWB = mConPClass.CondProperties.n_LaWB
                        Call LA_MixedAir_HP(CoLaDB, CoLaWB)
                        drItems("LADB") = Round(CoLaDB, 2)
                        drItems("LAWB") = FormatPercent(CoLaWB, 1)
                    Else
                        drItems("LADB") = Round(mConPClass.CondProperties.n_LaDB, 2)
                        drItems("LAWB") = FormatPercent(mConPClass.CondProperties.n_LaWB, 1)
                    End If
                Else
                    If PartTandem = True Then
                        EvLaDB = mEvaPClass.EvapProperties.n_LaDB
                        EvLaWB = mEvaPClass.EvapProperties.n_LaWB
                        Call LA_MixedAir_AC(EvLaDB, EvLaWB)
                        drItems("LADB") = Round(EvLaDB, 2)
                        drItems("LAWB") = Round(EvLaWB, 2)
                    Else
                        drItems("LADB") = Round(mEvaPClass.EvapProperties.n_LaDB, 2)
                        drItems("LAWB") = Round(mEvaPClass.EvapProperties.n_LaWB, 2)
                    End If
                End If

            ElseIf mSysP.i_CompQuantity = 1 Then

                If frmUnitInputs.optHeatPump.Checked = True Then
                    If mConPClass.CondProperties.s_Split = "IT" Or mConPClass.CondProperties.s_Split = "FS" Then
                        CoLaDB = mConPClass.CondProperties.n_LaDB
                        CoLaWB = mConPClass.CondProperties.n_LaWB
                        Call LA_MixedAir_HP(CoLaDB, CoLaWB)
                        drItems("LADB") = Round(CoLaDB, 2)
                        drItems("LAWB") = FormatPercent(CoLaWB, 1)
                    Else
                        drItems("LADB") = Round(mConPClass.CondProperties.n_LaDB, 2)
                        drItems("LAWB") = FormatPercent(mConPClass.CondProperties.n_LaWB, 1)
                    End If
                Else
                    If mEvaPClass.EvapProperties.s_Split = "IT" Or mEvaPClass.EvapProperties.s_Split = "FS" Then
                        EvLaDB = mEvaPClass.EvapProperties.n_LaDB
                        EvLaWB = mEvaPClass.EvapProperties.n_LaWB
                        Call LA_MixedAir_AC(EvLaDB, EvLaWB)
                        drItems("LADB") = Round(EvLaDB, 2)
                        drItems("LAWB") = Round(EvLaWB, 2)
                    Else
                        drItems("LADB") = Round(mEvaPClass.EvapProperties.n_LaDB, 2)
                        drItems("LAWB") = Round(mEvaPClass.EvapProperties.n_LaWB, 2)
                    End If
                End If

            End If

            dtItems.Rows.Add(drItems)

        End If

    End Sub

    Private Sub WriteSchemaToFile(ByVal dsItems As DataSet)

        dsItems.DataSetName = "dsItems"
        Dim filename As String = frmUnitInputs.AppPath() & "\dsItems.xsd"
        dsItems.Tables(0).WriteXmlSchema(filename)
    End Sub

    Public Sub ShowReport(ByVal dsItems As DataSet)

        Dim MyReport As New UnitOut()

        Try
            MyReport.SetDataSource(dsItems)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Exit Sub
        End Try

        frmUnitOutput.UnitReportViewer.ReportSource = MyReport

        With MyReport.DataDefinition.FormulaFields
            .Item("strUnitModel").Text = "'" & Left(mSysP.s_UnitModel, 80) & "'"
            .Item("strEvapModel").Text = "'" & Left(mEvaPClass.EvapProperties.s_Model, 10) & "'"
            .Item("strCondModel").Text = "'" & Left(mConPClass.CondProperties.s_Model, 10) & "'"
            .Item("strCompModel").Text = "'" & Left(mComPClass.CompProperties.s_Model, 28) & "'"
            .Item("strAppCode").Text = "'" & mComPClass.CompProperties.s_AppCode & "'"
            .Item("strRefrigerant").Text = "'" & mComPClass.CompProperties.s_RefCode & "'"
            .Item("nuEvapQty").Text = mSysP.i_EvapQuantity
            .Item("nuCondQty").Text = mSysP.i_CondQuantity
            .Item("nuCompQty").Text = mSysP.i_CompQuantity
            .Item("nuCompHZ").Text = mComPClass.CompProperties.s_FreCode
            .Item("nuSuperh").Text = mSysP.n_Superh
            .Item("nuSubcool").Text = mSysP.n_Subcool
            .Item("nuEvapCFM").Text = mSysP.n_EvCFM
            .Item("nuEvapDB").Text = mSysP.n_EvEaDB
            .Item("nuEvapWB").Text = mSysP.n_EvEaWB
            .Item("nuCondCFM").Text = mSysP.n_CoCFM
            .Item("nuCondDB").Text = mSysP.n_CoEaDB
            .Item("nuCondWB").Text = mSysP.n_CoEaWB
            .Item("nuOutacfm").Text = mSysP.n_OaCFM
            .Item("strEvapCoilPat").Text = "'" & Left(mEvaPClass.EvapProperties.s_CoilPat, 1) & "'"
            .Item("strCondCoilPat").Text = "'" & Left(mConPClass.CondProperties.s_CoilPat, 1) & "'"
            .Item("strReheatCoilPat").Text = "'" & Left(mRecPClass.ReheatProperties.s_CoilP, 1) & "'"
            .Item("nuEvapFH").Text = Round(mEvaPClass.EvapProperties.n_FH, 3)
            .Item("nuCondFH").Text = Round(mConPClass.CondProperties.n_FH, 3)
            .Item("nuReheatFH").Text = Round(mRecPClass.ReheatProperties.n_FH, 3)
            .Item("nuEvapFL").Text = Round(mEvaPClass.EvapProperties.n_FL, 2)
            .Item("nuCondFL").Text = Round(mConPClass.CondProperties.n_FL, 2)
            .Item("nuReheatFl").Text = Round(mRecPClass.ReheatProperties.n_FL, 2)
            .Item("nuEvapRD").Text = mEvaPClass.EvapProperties.i_Rows
            .Item("nuCondRD").Text = mConPClass.CondProperties.i_Rows
            .Item("nuReheatRd").Text = mRecPClass.ReheatProperties.i_Rows
            .Item("nuEvapFPI").Text = mEvaPClass.EvapProperties.i_FPI
            .Item("nuCondFPI").Text = mConPClass.CondProperties.i_FPI
            .Item("nuReheatFPI").Text = mRecPClass.ReheatProperties.i_FPI
            .Item("strEvapSplit").Text = "'" & Left(mEvaPClass.EvapProperties.s_Split, 2) & "'"
            .Item("strCondSplit").Text = "'" & Left(mConPClass.CondProperties.s_Split, 2) & "'"
            .Item("strReheatSplit").Text = "'" & Left("NS", 2) & "'"
            .Item("nuEvapTotalF").Text = mEvaPClass.EvapProperties.i_CKT
            .Item("nuCondTotalF").Text = mConPClass.CondProperties.i_CKT
            .Item("nuReheatTotalF").Text = mRecPClass.ReheatProperties.i_CKT
            .Item("nuEvapCKT1F").Text = mEvaPClass.EvapProperties.i_CKT1
            .Item("nuCondCKT1F").Text = mConPClass.CondProperties.i_CKT1
            .Item("nuEvapCKT2F").Text = mEvaPClass.EvapProperties.i_CKT2
            .Item("nuCondCKT2F").Text = mConPClass.CondProperties.i_CKT2
            .Item("nuEvapAirPD").Text = Round(mEvaPClass.EvapProperties.n_AirPD, 2)
            .Item("nuCondAirPD").Text = Round(mConPClass.CondProperties.n_AirPD, 2)
            .Item("nuEvapHeadOD").Text = Round(mEvaPClass.EvapProperties.n_HeaderId, 3)
            .Item("nuCondHeadOD").Text = Round(mConPClass.CondProperties.n_HeaderId, 3)
            .Item("nuReheatHeadOD").Text = Round(mConPClass.CondProperties.n_HeaderId, 3)
            .Item("nuEvapWThk").Text = Round(mEvaPClass.EvapProperties.n_WallThk, 3)
            .Item("nuCondWThk").Text = Round(mConPClass.CondProperties.n_WallThk, 3)
            .Item("nuReheatWThk").Text = Round(mRecPClass.ReheatProperties.n_WallThk, 3)
            .Item("nuEvapFThk").Text = Round(mEvaPClass.EvapProperties.n_FinThk, 3)
            .Item("nuCondFThk").Text = Round(mConPClass.CondProperties.n_FinThk, 3)
            .Item("nuReheatfThk").Text = Round(mRecPClass.ReheatProperties.n_FinThk, 3)
            .Item("strOutput").Text = "'" & Left(mSysP.s_Hit, 85) & "'"

            If frmUnitInputs.optHeatPump.Checked = True Then
                .Item("strAppMode").Text = "'" & Left("Heat Pump", 10) & "'"
                .Item("strLAWB").Text = "'" & Left("RH%", 5) & "'"
            Else
                .Item("strAppMode").Text = "'" & Left("Air Cond", 10) & "'"
                .Item("strLAWB").Text = "'" & Left("Deg F", 5) & "'"
            End If

        End With

        frmUnitOutput.Show()

    End Sub

End Module
