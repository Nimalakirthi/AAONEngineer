Imports System.Math

Public Structure Cond
    Public i_FPI As Integer    'fins per inch
    Public i_Rows As Integer    'rows deep
    Public i_CKT As Integer    'No of feeds
    Public i_CKT1 As Integer    'Feeds on Circuit 1
    Public i_CKT2 As Integer    'Feeds on Circuit 2

    Public n_FH As Double     'fin height
    Public n_FinThk As Double     'fin thickness
    Public s_FinMat As String     'fin material
    Public n_FL As Double     'fined length
    Public n_WallThk As Double     'tube wall thickness

    Public s_CoilPat As String
    Public s_Model As String     'coil model
    Public s_Droptubes As String
    Public s_Split As String
    Public n_BTUH As Double
    Public n_AsubS As Double
    Public n_AsubP As Double
    Public n_AsubO As Double
    Public n_AsubI As Double
    Public n_CondPD As Double
    Public n_HeaderId As Double
    Public n_AirPD As Double
    Public n_RefCharge As Double
    Public n_CoEaDB As Double
    Public n_CoEaWB As Double
    Public n_LaDB As Double
    Public n_LaWB As String
    Public n_CoACFM As Double
    Public s_Refrigerant As String
    Public s_ReheatType As String
    Public i_CompQuantity As Integer
    Public i_CondQuantity As Integer
    Public s_CompAppCode As String
    Public n_Altitude As Double

End Structure

Public Class CondClass

    Dim mRefObject As pjRefProp.RefProp

    Private mConP As Cond

    Public Property CondProperties() As Cond
        Get
            Return mConP
        End Get

        Set(ByVal NewCond As Cond)
            mConP = NewCond
        End Set
    End Property

    Public Sub PassAirCorrection(ByVal CoEaDB As Double, ByVal CoEaWB As Double, ByVal CoACFM As Double, ByVal AppCode As String)
        mConP.n_CoEaDB = CoEaDB
        mConP.n_CoEaWB = CoEaWB
        mConP.n_CoACFM = CoACFM
        mConP.s_CompAppCode = AppCode
    End Sub
    Public Sub PassMixAirInfo(ByVal CoEaDB As Double, ByVal CoEaWB As Double, ByVal CoACFM As Double, ByVal AppCode As String)

        mConP.n_CoEaDB = CoEaDB
        mConP.n_CoEaWB = CoEaWB
        mConP.n_CoACFM = CoACFM
        mConP.s_CompAppCode = AppCode
    End Sub
    Public Sub CondCalc(ByVal mn_MassFlow As Double, ByVal MassFraction As Double, ByVal mn_ET As Double, ByVal mn_CT As Double)

        Dim mn_EatDB As Double
        Dim mn_EatWB As Double
        Dim mn_AsubS As Double
        Dim mn_AsubP As Double
        Dim mn_AsubO As Double
        Dim mn_AsubI As Double
        Dim n_Am As Double
        Dim mn_FinThk As Double
        Dim mn_ACFM As Double
        Dim mn_id As Double
        Dim n_CoOD As Double
        Dim n_CoSt As Double
        Dim n_CoSl As Double
        Dim n_Cont As Double
        Dim n_EWBlblbSR As Double
        Dim n_IntV1 As Double
        Dim n_ElblbR As Double
        Dim n_LlblbR As Double
        Dim n_LDBlblbSR As Double
        Dim n_LSatWVPre As Double
        Dim n_LWVPre As Double
        Dim n_LHumidity As Double
        Dim n_AirDensity As Double
        Dim n_cpair As Double
        Dim n_PRA As Double
        Dim n_FaceArea As Double
        Dim mn_AirVelocity As Double
        Dim nMaxAirvelocity1 As Double
        Dim nReOD As Double
        Dim nReSl As Double      'Renyolds number based on logtidudinal tube spacing
        Dim nJP As Double      'Parameter JP in McQuinston Correlation
        Dim nGc As Double      'Air Mass velocity at minimum flow area
        Dim nho As Double      'Heat transfer coefficient of the outer suf of coil
        Dim nRf As Double      'R Value of Fins (to be evaluated explicitly in future)
        Dim nfinef As Double      'Fin efficiency
        Dim nfinsef As Double      'Fin Surface efficiency
        Dim nhi As Double
        Dim nxx As Double      'intermediate variable
        Dim nyy As Double      'intermediate variable
        Dim nvv As Double      'intermediate variable
        Dim nzz As Double
        Dim nUA As Double      'Overall heat Trnsfer Coefficient * Outside Area
        Dim nCa As Double      'Heat Capacity Rate of entering Air
        Dim nNTU As Double      'Number of Transfer Units
        Dim nEveff As Double      'Effectiveness of Condenserr coil
        Dim n_ReRefV As Double
        Dim n_ReLiq As Double
        Dim n_PrLiq As Double
        Dim nM As Double       'Intermediate Variable
        Dim n_DenV As Double
        Dim n_DenL As Double
        Dim n_VisV As Double
        Dim n_VisL As Double
        Dim n_kL As Double
        Dim n_CpL As Double
        Dim n_fink As Double
        Dim Reyeq As Double
        Dim Densityratio As Double
        Dim n_NoCKT As Integer      'Number of Circuits
        Dim n_RefMassFlRate As Double
        Dim n_RefLatent As Double
        Dim s_CoilType As String
        Dim n_TubesPerCKT As Double
        Dim n_CKTLength As Double
        Dim n_TubeHi As Double
        Dim n_HeaderId As Double
        Dim n_HeaderL As Double
        Dim nAmin As Double        ' minimum flow area
        Dim ParameterA As Double         ' Transverse tube spacing/tube OD
        Dim nRows As Double
        Dim VaporBTUH As Double        'Capacity due to vapor fraction
        Dim LiqBTUH As Double
        Dim n_ReRefL As Double        'Reynolds number for Liquid portion of condenser for series Reheat
        Dim nhV As Double        'Inside heat Transfer coefficient for vapor portion for series reheat
        Dim nhL As Double        'Inside heat Transfer coefficient for liquid portion for series reheat
        Dim CondMassFraction As Double
        Dim Si As Single
        Dim beta As Single
        Dim eqradius As Single
        Dim SurParameter As Single
        Dim resnumber As Single
        Dim tanheff As Single
        Dim tempfine As Single
        Dim n_NTUAo As Single
        Dim n_Twi As Single
        Dim n_TFin As Single
        Dim mn_LatDB As Single
        Dim n_Hae As Single
        Dim n_Hal As Single
        Dim C As Single
        Dim nNTUC As Single
        Dim nEveff1 As Single
        Dim nEveff2 As Single
        Dim nEveff3 As Single
        Dim nEveffCross As Single
        Dim OD_Btuh As Single
        Dim Refrigerant As String
        Dim CompQuantity As Integer
        Dim CondQuantity As Integer
        Dim s_ReheatType As String
        Dim Altitude As Single
        Const n_cp As Single = 0.24
        Const n_TuCond As Single = 220
        Const vaporquality As Single = 0.75

        Const mn_a1 As Double = 0.00080246
        Const mn_a2 As Double = 0.000024525
        Const mn_a3 As Double = 2.542 * (10 ^ -6)
        Const mn_a4 As Double = -2.5855 * (10 ^ -8)
        Const mn_a5 As Double = 4.038 * (10 ^ -10)
        Dim mn_PSIA As Double

        Refrigerant = mConP.s_Refrigerant
        CompQuantity = mConP.i_CompQuantity
        CondQuantity = mConP.i_CondQuantity
        s_ReheatType = mConP.s_ReheatType
        Altitude = mConP.n_Altitude

        mn_PSIA = maPSIA(Altitude)
        mRefObject = New pjRefProp.RefProp

        If Refrigerant = "22" Then
            n_DenV = mRefObject.R22_DenV(mn_CT)
            n_DenL = mRefObject.R22_DenL(mn_CT)
            n_VisV = mRefObject.R22_VisV(mn_CT)
            n_VisL = mRefObject.R22_VisL(mn_CT)
            n_kL = mRefObject.R22_kL(mn_CT)
            n_CpL = mRefObject.R22_CpL(mn_CT)

        ElseIf Refrigerant = "410A" Then
            n_DenV = mRefObject.R410A_DenV(mn_CT)
            n_DenL = mRefObject.R410A_DenL(mn_CT)
            n_VisV = mRefObject.R410A_VisV(mn_CT)
            n_VisL = mRefObject.R410A_VisL(mn_CT)
            n_kL = mRefObject.R410A_kL(mn_CT)
            n_CpL = mRefObject.R410A_CpL(mn_CT)

        ElseIf Refrigerant = "404A" Then
            n_DenV = mRefObject.R404A_DenV(mn_CT)
            n_DenL = mRefObject.R404A_DenL(mn_CT)
            n_VisV = mRefObject.R404A_VisV(mn_CT)
            n_VisL = mRefObject.R404A_VisL(mn_CT)
            n_kL = mRefObject.R404A_kL(mn_CT)
            n_CpL = mRefObject.R404A_CpL(mn_CT)

        ElseIf Refrigerant = "407C Mid Pt" Or Refrigerant = "407C Dew Pt" Then
            n_CpL = mRefObject.R407C_CpL(mn_CT)
            n_DenV = mRefObject.R407C_DenV(mn_CT)
            n_DenL = mRefObject.R407C_DenL(mn_CT)
            n_VisV = mRefObject.R407C_VisV(mn_CT)
            n_VisL = mRefObject.R407C_VisL(mn_CT)
            n_kL = mRefObject.R407C_kL(mn_CT)

        ElseIf Refrigerant = "134A" Then
            n_kL = mRefObject.R134A_kL(mn_CT)
            n_CpL = mRefObject.R134A_CpL(mn_CT)
            n_DenV = mRefObject.R134A_DenV(mn_CT)
            n_DenL = mRefObject.R134A_DenL(mn_CT)
            n_VisV = mRefObject.R134A_VisV(mn_CT)
            n_VisL = mRefObject.R134A_VisL(mn_CT)

        ElseIf Refrigerant = "507" Then
            n_kL = mRefObject.R507_kL(mn_CT)
            n_CpL = mRefObject.R507_CpL(mn_CT)
            n_DenV = mRefObject.R507_DenV(mn_CT)
            n_DenL = mRefObject.R507_DenL(mn_CT)
            n_VisV = mRefObject.R507_VisV(mn_CT)
            n_VisL = mRefObject.R507_VisL(mn_CT)
        End If

        If mConP.s_Split = "IT" And CompQuantity = 1 Then

            If mConP.s_CoilPat = "6" Then
                n_CoOD = 0.331
                n_CoSt = 1 * 2
                n_CoSl = 0.625 * 2

            ElseIf mConP.s_CoilPat = "3" Then
                n_CoOD = 0.395
                n_CoSt = 1 * 2
                n_CoSl = 0.866 * 2

            ElseIf mConP.s_CoilPat = "P" Then
                n_CoOD = 0.52
                n_CoSt = 1.25 * 2
                n_CoSl = 1.0825 * 2

            ElseIf mConP.s_CoilPat = "5" Then
                n_CoOD = 0.645
                n_CoSt = 1.5 * 2
                n_CoSl = 1.299 * 2

            ElseIf mConP.s_CoilPat = "7" Then
                n_CoOD = 0.282
                n_CoSl = 0.472 * 2  'tube stagger
                n_CoSt = 0.827 * 2

            End If

        ElseIf mConP.s_Split = "IT" And CompQuantity = 0 Then

            If mConP.s_CoilPat = "6" Then
                n_CoOD = 0.331
                n_CoSt = 1 * 2
                n_CoSl = 0.625 * 2

            ElseIf mConP.s_CoilPat = "3" Then
                n_CoOD = 0.395
                n_CoSt = 1 * 2
                n_CoSl = 0.866 * 2

            ElseIf mConP.s_CoilPat = "P" Then
                n_CoOD = 0.52
                n_CoSt = 1.25 * 2
                n_CoSl = 1.0825 * 2

            ElseIf mConP.s_CoilPat = "5" Then
                n_CoOD = 0.645
                n_CoSt = 1.5 * 2
                n_CoSl = 1.299 * 2

            ElseIf mConP.s_CoilPat = "7" Then
                n_CoOD = 0.282
                n_CoSl = 0.472 * 2  'tube stagger
                n_CoSt = 0.827 * 2

            End If

        Else

            If mConP.s_CoilPat = "6" Then
                n_CoOD = 0.331
                n_CoSt = 1
                n_CoSl = 0.625

            ElseIf mConP.s_CoilPat = "3" Then
                n_CoOD = 0.395
                n_CoSt = 1
                n_CoSl = 0.866

            ElseIf mConP.s_CoilPat = "P" Then
                n_CoOD = 0.52
                n_CoSt = 1.25
                n_CoSl = 1.0825

            ElseIf mConP.s_CoilPat = "5" Then
                n_CoOD = 0.645
                n_CoSt = 1.5
                n_CoSl = 1.299

            ElseIf mConP.s_CoilPat = "1" Then
                n_CoOD = 1
                n_CoSt = 3
                n_CoSl = 1.5

            ElseIf mConP.s_CoilPat = "7" Then
                n_CoOD = 0.282
                n_CoSl = 0.472    'tube stagger
                n_CoSt = 0.827

            End If

        End If

        If mConP.s_FinMat = "A" Then
            n_fink = 127      'AL(1000)
        Else
            n_fink = 226     'CU(12200)
        End If

        mn_EatDB = mConP.n_CoEaDB 'n_CoEaDB
        mn_EatWB = mConP.n_CoEaWB 'n_CoEaWB
        mn_AsubS = mConP.n_AsubS 'Class.CondProperties.n_AsubS
        mn_AsubP = mConP.n_AsubP 'Class.CondProperties.n_AsubP
        mn_AsubO = mConP.n_AsubO 'Class.CondProperties.n_AsubO
        mn_AsubI = mConP.n_AsubI 'Class.CondProperties.n_AsubI
        mn_FinThk = mConP.n_FinThk 'Class.CondProperties.n_FinThk
        mn_ACFM = mConP.n_CoACFM  'n_CoACFM
        mn_id = n_CoOD - 2 * mConP.n_WallThk 'Class.CondProperties.n_WallThk
        n_Cont = mConP.n_FH * mConP.i_Rows / n_CoSt
        n_Am = 3.14 * n_CoOD * n_Cont * mConP.n_FL

        n_FaceArea = (mConP.n_FL * mConP.n_FH) / 144
        nRows = mConP.i_Rows

        If mConP.s_CompAppCode = "AH" Or mConP.s_CompAppCode = "A5" Then
            mn_ACFM = mn_ACFM * 0.5
        ElseIf mConP.s_CompAppCode = "A6" Then
            mn_ACFM = mn_ACFM * 0.66
        ElseIf mConP.s_CompAppCode = "AL" Then
            mn_ACFM = mn_ACFM * 0.55
        ElseIf mConP.s_CompAppCode = "AS" Then
            mn_ACFM = mn_ACFM * 0.45
        End If

        If CondQuantity = 1 Then

            If mConP.s_Split = "IT" And CompQuantity = 2 Then
                mn_MassFlow = 2 * mn_MassFlow
                n_NoCKT = mConP.i_CKT
                mn_AirVelocity = mn_ACFM / n_FaceArea

            ElseIf mConP.s_Split = "IT" And CompQuantity = 1 Then
                mn_AirVelocity = mn_ACFM / n_FaceArea
                mn_ACFM = mn_ACFM * 0.5
                If mConP.i_CKT1 > mConP.i_CKT2 Then
                    n_NoCKT = mConP.i_CKT1
                Else
                    n_NoCKT = mConP.i_CKT2
                End If

            ElseIf mConP.s_Split = "IT" And CompQuantity = 0 Then
                mn_AirVelocity = mn_ACFM / n_FaceArea
                mn_ACFM = mn_ACFM * 0.5
                If mConP.i_CKT1 > mConP.i_CKT2 Then
                    n_NoCKT = mConP.i_CKT1
                Else
                    n_NoCKT = mConP.i_CKT2
                End If

            ElseIf mConP.s_Split = "FS" Then
                mn_ACFM = mn_ACFM * 0.5
                n_NoCKT = mConP.i_CKT * 0.5
                mn_AirVelocity = mn_ACFM / n_FaceArea

            ElseIf mConP.s_Split = "NS" Then
                mn_AirVelocity = mn_ACFM / n_FaceArea
                n_NoCKT = mConP.i_CKT
            End If

        Else

            n_NoCKT = mConP.i_CKT
            mn_ACFM = mn_ACFM * 0.5
            mn_AirVelocity = mn_ACFM / n_FaceArea

        End If

        If Refrigerant = "410A" Then
            n_RefLatent = mRefObject.Sysbal_R410A_HFG(mn_CT)
        ElseIf Refrigerant = "22" Then
            n_RefLatent = mRefObject.Sysbal_R22_HG(mn_CT) - mRefObject.Sysbal_R22_HL(mn_CT)
        ElseIf Refrigerant = "407C Mid Pt" Or Refrigerant = "407C Dew Pt" Then
            n_RefLatent = mRefObject.Sysbal_R407C_HFG(mn_CT)
        ElseIf Refrigerant = "404A" Then
            n_RefLatent = mRefObject.Sysbal_R404A_HFG(mn_CT)
        ElseIf Refrigerant = "134A" Then
            n_RefLatent = mRefObject.Sysbal_R134A_HFG(mn_CT)
        ElseIf Refrigerant = "507" Then
            n_RefLatent = mRefObject.Sysbal_R404A_HFG(mn_CT)
        End If

        If mn_EatDB >= 90 Then
            n_PRA = 0.703                                    'Prandalt's number of Air
        ElseIf mn_EatDB >= 70 Then
            n_PRA = 0.708
        ElseIf mn_EatDB >= 50 Then
            n_PRA = 0.711
        Else
            n_PRA = 0.716
        End If

        n_EWBlblbSR = mn_a1 + (mn_a2 * mn_EatWB) + (mn_a3 * mn_EatWB ^ 2) + (mn_a4 * mn_EatWB ^ 3) + (mn_a5 * mn_EatWB ^ 4)
        n_IntV1 = ((1093 - 0.556 * mn_EatWB) * n_EWBlblbSR) - (0.24 * (mn_EatDB - mn_EatWB))
        n_ElblbR = n_IntV1 / (1093 + 0.444 * mn_EatDB - mn_EatWB)

        n_AirDensity = maDensity(mn_EatDB, n_ElblbR, mn_PSIA)
        n_cpair = 0.24 + 0.444 * n_ElblbR

        Select Case mConP.s_CoilPat
            Case Is = "P"
                nMaxAirvelocity1 = mn_AirVelocity * 1.25 / (1.25 - n_CoOD)     'ft/Min
                nReOD = n_AirDensity * (nMaxAirvelocity1 * 60) * (n_CoOD / 12) / 0.0437      'Air Dynamic viscosity=0.0437lb/ft-hr
                nAmin = n_FaceArea * (1.25 - n_CoOD) / 1.25                                   'nReOD is Air Side Reynolds number based on OD
                ParameterA = 1.25 / n_CoOD
                nReSl = n_AirDensity * (nMaxAirvelocity1 * 60) * (1.08 / 12) / 0.0437

            Case Is = "5"
                nMaxAirvelocity1 = mn_AirVelocity * 1.5 / (1.5 - n_CoOD)     'ft/Min
                nReOD = n_AirDensity * (nMaxAirvelocity1 * 60) * (n_CoOD / 12) / 0.0437      'Air Dynamic viscosity=0.0437lb/ft-hr
                nAmin = n_FaceArea * (1.5 - n_CoOD) / 1.5
                ParameterA = 1.5 / n_CoOD
                nReSl = n_AirDensity * (nMaxAirvelocity1 * 60) * (1.299 / 12) / 0.0437

            Case Is = "3"
                nMaxAirvelocity1 = mn_AirVelocity * 1 / (1 - n_CoOD)     'ft/Min
                nReOD = n_AirDensity * (nMaxAirvelocity1 * 60) * (n_CoOD / 12) / 0.0437      'Air Dynamic viscosity=0.0437lb/ft-hr
                nAmin = n_FaceArea * (1 - n_CoOD)
                ParameterA = 1 / n_CoOD
                nReSl = n_AirDensity * (nMaxAirvelocity1 * 60) * (0.866 / 12) / 0.0437

            Case Is = "6"
                nMaxAirvelocity1 = mn_AirVelocity * 1 / (1 - n_CoOD)     'ft/Min
                nReOD = n_AirDensity * (nMaxAirvelocity1 * 60) * (n_CoOD / 12) / 0.0437     'Air Dynamic viscosity=0.0437lb/ft-hr
                nAmin = n_FaceArea * (1 - n_CoOD)
                ParameterA = 1 / n_CoOD
                nReSl = n_AirDensity * (nMaxAirvelocity1 * 60) * (0.625 / 12) / 0.0437

            Case Is = "7"
                nMaxAirvelocity1 = mn_AirVelocity * 0.827 / (0.827 - n_CoOD)     'ft/Min
                nReOD = n_AirDensity * (nMaxAirvelocity1 * 60) * (n_CoOD / 12) / 0.0437      'Air Dynamic viscosity=0.0437lb/ft-hr
                nAmin = n_FaceArea * (0.827 - n_CoOD) / 0.827
                ParameterA = 0.827 / n_CoOD
                nReSl = n_AirDensity * (nMaxAirvelocity1 * 60) * (0.472 / 12) / 0.0437
        End Select

        nJP = (nReOD ^ -0.4) * ((mn_AsubO / mn_AsubP) ^ -0.15)   'Parameter JP in Chilton-Colburn j-factors

        nGc = n_AirDensity * (nMaxAirvelocity1 * 60)    'Air Mass Velocity at Minimum Flow Area

        nho = (nGc * n_cpair) * ((0.00125 + 0.275 * nJP) / (n_PRA ^ 0.666))   'Outer surface area heat transfer coefficient

        Si = n_CoSt / n_CoOD
        beta = ((n_CoSt * 0.5) ^ 2) + ((n_CoSl) ^ 2)
        beta = (beta ^ 0.5) / n_CoSt
        eqradius = 1.27 * Si * ((beta - 0.3) ^ 0.5)
        SurParameter = (2 * nho / (n_fink * mConP.n_FinThk * 0.0833)) ^ 0.5
        resnumber = (eqradius - 1) * (1 + 0.35 * Log(eqradius))
        tempfine = (Exp(SurParameter * n_CoOD * 0.0416 * resnumber) - Exp(-SurParameter * n_CoOD * 0.0416 * resnumber)) / (Exp(SurParameter * n_CoOD * 0.0416 * resnumber) + Exp(-SurParameter * n_CoOD * 0.0416 * resnumber))
        tanheff = tempfine / (SurParameter * n_CoOD * 0.0416 * resnumber)
        nfinef = tanheff

        nRf = ((1 / nfinef) - 1) / nho

        nfinsef = (1 - (mn_AsubS / mn_AsubO) * (1 - nfinef)) 'Surface Efficiency

        nxx = nfinsef * nho * mn_AsubO  'Heat Transfer term realted to outer surface area

        nCa = 60 * mn_ACFM * n_AirDensity * n_cpair 'Moist Air heat cpacity rate

        '    If CompQuantity = 2 Or CompQuantity = 0 Then
        '        n_ReRefV = (mn_MassFlow * (1 - (MassFraction * 0.5)) * 4 * 12) / (3.14 * mConP.i_CKT * n_VisL * mn_id)
        '    Else
        n_ReRefV = (mn_MassFlow * (1 - MassFraction) * 4 * 12) / (3.14 * mConP.i_CKT * n_VisL * mn_id) 'Refrigerant side Reynolds numbeer ReLiq
        '    End If

        Densityratio = (n_DenL / n_DenV) ^ 0.5
        Reyeq = n_ReRefV * ((1 - vaporquality) + vaporquality * Densityratio * (n_VisV / n_VisL))

        n_PrLiq = (n_VisL * n_CpL / n_kL) ''Refrigerant side Prandalt's number

        'nM = ((n_DenL / n_DenV) ^ 0.375) * ((n_VisV / n_VisL) ^ 0.075)

        nhV = 0.05 * mConP.i_CKT * (n_kL * 12 / mn_id) * (Reyeq ^ 0.8) * (n_PrLiq ^ 0.33) 'Average refrigirant side heat transfer coefficient
        'Dittus-Boelter Equation

        If s_ReheatType <> "VaporSeries" Then
            If nhV <> 0 Then
                nhi = nhV
                nyy = nhi * mn_AsubI

                nvv = mConP.n_WallThk / (24 * n_TuCond) 'TBE: Thermal conductivity of Tube wall 220
                nvv = nvv / n_Am

                nzz = (1 / nxx) + nvv + (1 / nyy)   'Intermediate variable

                nUA = 1 / nzz       'Overall Heat transfer Coeficient

                nNTU = nUA / nCa    'Number of Transfer Units

                nEveff = (1 - Exp(-nNTU)) 'Condenser Effectiveness

                VaporBTUH = nEveff * nCa * (mn_CT - mn_EatDB)

                n_NTUAo = nxx / nCa

                n_Twi = mn_CT - (VaporBTUH / (nhi * mn_AsubI))
                n_TFin = n_Twi - (VaporBTUH * 12 * n_CoSt * Log(n_CoOD / mn_id) / (mConP.n_FH * mConP.n_FL * mConP.i_Rows * 2 * 3.14 * n_TuCond))
                VaporBTUH = (1 - Exp(-n_NTUAo)) * (n_TFin - mn_EatDB) * nCa

            Else
                VaporBTUH = 0
            End If
        Else

            CondMassFraction = MassFraction

            If nhV <> 0 Then
                nhi = nhV
                nyy = nhi * mn_AsubI

                nvv = mConP.n_WallThk / (24 * 220) 'TBE: Thermal conductivity of Tube wall 220
                nvv = nvv / n_Am

                nzz = (1 / nxx) + nvv + (1 / nyy)   'Intermediate variable

                nUA = 1 / nzz       'Overall Heat transfer Coeficient

                nNTU = nUA / (nCa * (1 - CondMassFraction)) 'Number of Transfer Units

                nEveff = (1 - Exp(-nNTU)) 'Condenser Effectiveness

                VaporBTUH = nEveff * nCa * (mn_CT - mn_EatDB)

                n_NTUAo = nxx / (nCa * (1 - CondMassFraction))

                n_Twi = mn_CT - (VaporBTUH / (nhi * mn_AsubI))
                n_TFin = n_Twi - (VaporBTUH * 12 * n_CoSt * Log(n_CoOD / mn_id) / (mConP.n_FH * mConP.n_FL * mConP.i_Rows * 2 * 3.14 * n_TuCond))
                VaporBTUH = (1 - Exp(-n_NTUAo)) * (n_TFin - mn_EatDB) * nCa * (1 - CondMassFraction)

            End If

            n_ReRefL = (mn_MassFlow * CondMassFraction * 4 * 12) / (3.14 * mConP.i_CKT * n_VisL * mn_id)
            nhL = 0.023 * mConP.i_CKT * (n_kL * 12 / mn_id) * (n_ReRefL ^ 0.8) * (n_PrLiq ^ 0.4)

            If nhL <> 0 Then
                nhi = nhL
                nyy = nhi * mn_AsubI

                nvv = mConP.n_WallThk / (24 * 220) 'TBE: Thermal conductivity of Tube wall 220
                nvv = nvv / n_Am

                nzz = (1 / nxx) + nvv + (1 / nyy)   'Intermediate variable

                nUA = 1 / nzz       'Overall Heat transfer Coeficient

                nNTU = nUA / (nCa * CondMassFraction) 'Number of Transfer Units

                C = (n_AirDensity * mn_ACFM * n_cpair) / (mn_MassFlow * n_CpL)
                nNTUC = nNTU * (1 - C)

                If mConP.i_Rows = 1 Then
                    nEveff = (1 - Exp(-nNTUC)) / (1 - C * Exp(-nNTUC)) ' coil Effectivenesss based on overall heat transfer
                ElseIf mConP.i_Rows = 2 Then
                    nEveff1 = 1 - (Exp(-nNTU / 2))
                    nEveff2 = 1 + C * (nEveff1 ^ 2)
                    nEveffCross = 1 - ((Exp(-2 * nEveff1 * C)) * nEveff2)
                    nEveff = nEveffCross / C
                ElseIf mConP.i_Rows = 3 Then
                    nEveff1 = 1 - (Exp(-nNTU / 3))
                    nEveff2 = 1 + (C * (nEveff1 ^ 2) * (3 - nEveff1)) + (1.5 * (C ^ 2) * (nEveff1 ^ 4))
                    nEveffCross = 1 - ((Exp(-3 * nEveff1 * C)) * nEveff2)
                    nEveff = nEveffCross / C
                ElseIf mConP.i_Rows = 4 Then
                    nEveff1 = 1 - (Exp(-nNTU / 4))
                    nEveff2 = 1 + (C * (nEveff1 ^ 2) * (6 - (4 * nEveff1) + (nEveff1 ^ 2)))
                    nEveff3 = nEveff2 + (2.666 * (C ^ 3) * (nEveff1 ^ 6)) + (4 * (C ^ 2) * (nEveff1 ^ 4) * (2 - nEveff1))
                    nEveffCross = 1 - ((Exp(-4 * nEveff1 * C)) * nEveff3)
                    nEveff = nEveffCross / C
                ElseIf mConP.i_Rows > 4 Then
                    nEveff1 = ((Exp(-C * (nNTU ^ 0.78))) - 1) / C
                    nEveffCross = 1 - (Exp((nNTU ^ 0.22) * nEveff1))
                    nEveff = nEveffCross
                End If

                '         nEveff = (1 - Exp(-nNTU)) 'Condenser Effectiveness

                LiqBTUH = nEveff * nCa * (mn_CT - mn_EatDB) * CondMassFraction
            End If
        End If

        mConP.n_BTUH = VaporBTUH + LiqBTUH

        mn_LatDB = (n_TFin + ((mn_EatDB - n_TFin) * Exp(-n_NTUAo)))
        mConP.n_LaDB = mn_LatDB
        n_Hae = n_cp * mn_EatDB + n_ElblbR * (1061 + 0.444 * mn_EatDB)
        n_Hal = (mConP.n_BTUH / (mn_ACFM * 60 * n_AirDensity)) + n_Hae
        n_LlblbR = (n_Hal - n_cp * mn_LatDB) / (1061 + 0.444 * mn_LatDB)
        n_LDBlblbSR = mn_a1 + (mn_a2 * mn_LatDB) + (mn_a3 * mn_LatDB ^ 2) + (mn_a4 * mn_LatDB ^ 3) + (mn_a5 * mn_LatDB ^ 4)
        n_LSatWVPre = (mn_PSIA * n_LDBlblbSR) / (n_LDBlblbSR + 0.62198)
        n_LWVPre = (mn_PSIA * n_LlblbR) / (n_LlblbR + 0.62198)
        n_LHumidity = n_LWVPre / n_LSatWVPre
        '        mConP.n_LaWB = n_LHumidity
        '        mConP.n_LaWB = Round(n_LHumidity, 1) & "%"
        mConP.n_LaWB = Round(n_LHumidity, 4)

        n_TubeHi = mConP.n_FH / n_CoSt
        n_TubesPerCKT = n_TubeHi * mConP.i_Rows / n_NoCKT
        n_CKTLength = n_TubesPerCKT * mConP.n_FL
        s_CoilType = "Cond"
        n_RefMassFlRate = mn_MassFlow * (1 - MassFraction)
        OD_Btuh = mConP.n_BTUH
        mConP.n_HeaderId = HeaderId(OD_Btuh)
        n_HeaderId = mConP.n_HeaderId
        n_HeaderL = mConP.n_FH
        n_ReLiq = n_ReRefV + n_ReRefL
        mConP.n_CondPD = CondFluidPD(mn_id, n_RefMassFlRate, n_NoCKT, n_HeaderL, n_CKTLength, n_ReLiq, _
                                        n_RefLatent, n_DenV, n_DenL, n_VisV, n_VisL, n_HeaderId)

        mConP.n_AirPD = CondAirPD(mn_AsubS, nAmin, nGc, nReOD, nReSl, ParameterA, nRows)
        mConP.n_RefCharge = CondRefCharge(n_HeaderId, n_TubesPerCKT)

    End Sub

    Public Sub StartCond(ByVal CompQuantity)

        Dim n_CoOD As Double
        Dim n_CoID As Double
        Dim n_CoSt As Double
        Dim n_CoSl As Double
        Dim n_Cold As Double
        Dim n_Colc As Double

        Dim n_Conf As Double
        Dim n_Conh As Double
        Dim n_Cont As Double

        If mConP.s_Split = "IT" And CompQuantity = 1 Then
            If mConP.s_CoilPat = "6" Then
                n_CoOD = 0.331
                n_CoSt = 1 * 2
                n_CoSl = 0.625 * 2

            ElseIf mConP.s_CoilPat = "3" Then
                n_CoOD = 0.395
                n_CoSt = 1 * 2
                n_CoSl = 0.866 * 2

            ElseIf mConP.s_CoilPat = "P" Then
                n_CoOD = 0.52
                n_CoSt = 1.25 * 2
                n_CoSl = 1.0825 * 2

            ElseIf mConP.s_CoilPat = "5" Then
                n_CoOD = 0.645
                n_CoSt = 1.5 * 2
                n_CoSl = 1.299 * 2

            ElseIf mConP.s_CoilPat = "7" Then
                n_CoOD = 0.282
                n_CoSl = 0.472 * 2  'tube stagger
                n_CoSt = 0.827 * 2

            End If

        ElseIf mConP.s_Split = "IT" And CompQuantity = 0 Then

            If mConP.s_CoilPat = "6" Then
                n_CoOD = 0.331
                n_CoSt = 1 * 2
                n_CoSl = 0.625 * 2

            ElseIf mConP.s_CoilPat = "3" Then
                n_CoOD = 0.395
                n_CoSt = 1 * 2
                n_CoSl = 0.866 * 2

            ElseIf mConP.s_CoilPat = "P" Then
                n_CoOD = 0.52
                n_CoSt = 1.25 * 2
                n_CoSl = 1.0825 * 2

            ElseIf mConP.s_CoilPat = "5" Then
                n_CoOD = 0.645
                n_CoSt = 1.5 * 2
                n_CoSl = 1.299 * 2

            ElseIf mConP.s_CoilPat = "7" Then
                n_CoOD = 0.282
                n_CoSl = 0.472 * 2  'tube stagger
                n_CoSt = 0.827 * 2

            End If

        Else

            If mConP.s_CoilPat = "6" Then
                n_CoOD = 0.331
                n_CoSt = 1
                n_CoSl = 0.625

            ElseIf mConP.s_CoilPat = "3" Then
                n_CoOD = 0.395
                n_CoSt = 1
                n_CoSl = 0.866

            ElseIf mConP.s_CoilPat = "P" Then
                n_CoOD = 0.52
                n_CoSt = 1.25
                n_CoSl = 1.0825

            ElseIf mConP.s_CoilPat = "5" Then
                n_CoOD = 0.645
                n_CoSt = 1.5
                n_CoSl = 1.299

            ElseIf mConP.s_CoilPat = "7" Then
                n_CoOD = 0.282
                n_CoSl = 0.472    'tube stagger
                n_CoSt = 0.827

            End If

        End If

        If mConP.s_Split = "FS" Then
            mConP.n_FH = mConP.n_FH * 0.5
        End If

        n_Cold = mConP.i_Rows * n_CoSl
        n_CoID = n_CoOD - (2 * mConP.n_WallThk)
        n_Conh = mConP.n_FH * mConP.i_Rows / n_CoSt
        n_Cont = n_Conh
        'mn_Ldrat = mn_FL / mn_id

        n_Conf = mConP.n_FL * mConP.i_FPI
        n_Colc = (mConP.n_FL - (n_Conf * mConP.n_FinThk)) / n_Conf

        mConP.n_AsubS = n_Conf * (((mConP.n_FH * n_Cold) / 72) - ((n_Conh * ((n_CoOD + 2 * mConP.n_FinThk) ^ 2)) / 91.68) + _
        (((n_CoOD + 2 * mConP.n_FinThk) * (n_Conh - n_Cont) * n_Colc) / 45.84))
        mConP.n_AsubP = ((n_Cont * n_CoOD * mConP.n_FL) - (n_Cont * n_Conf * mConP.n_FinThk) * _
        (n_CoOD - 2 * n_Colc)) / 45.84

        mConP.n_AsubO = mConP.n_AsubS + mConP.n_AsubP

        mConP.n_AsubI = (n_CoID * n_Cont * mConP.n_FL) / 45.84

    End Sub

    Public Function CondCirCuitryCheck() As Boolean

        Dim Cond_TubeHi As Double
        Dim Cond_Passess As Double
        Dim n_CoSt As Double
        Dim CondOpp As Boolean
        Dim Output As String
        Dim Cond_PassRound As Integer
        Dim Cond_TubeHiRound As Integer
        Dim Cond_DropTubes As Integer
        Dim TotalTubes As Double

        Output = ""

        If mConP.s_CoilPat = "6" Then
            n_CoSt = 1
        ElseIf mConP.s_CoilPat = "3" Then
            n_CoSt = 1
        ElseIf mConP.s_CoilPat = "P" Then
            n_CoSt = 1.25
        ElseIf mConP.s_CoilPat = "5" Then
            n_CoSt = 1.5
        ElseIf mConP.s_CoilPat = "7" Then
            n_CoSt = 0.827
        End If

        Cond_TubeHi = mConP.n_FH / n_CoSt

        Cond_TubeHiRound = Round(Cond_TubeHi)

        If Round(Cond_TubeHi, 4) <> Cond_TubeHiRound Then
            Output = Output & " Condenser Fin Height must be an integeral multiple of " + Str(n_CoSt) + "" & vbCrLf
        End If
        TotalTubes = mConP.i_Rows * Round(Cond_TubeHi, 4)
        Cond_Passess = TotalTubes / mConP.i_CKT
        Cond_PassRound = Int(Cond_Passess)

        Cond_DropTubes = (TotalTubes) - (Cond_PassRound * mConP.i_CKT)

        If mConP.i_Rows = mConP.i_CKT Then
            If (Cond_Passess / 2) <> Int(Cond_Passess / 2) Then
                Cond_DropTubes = mConP.i_Rows
            End If
        End If

        If Cond_DropTubes <> 0 Then

            If (Cond_DropTubes / TotalTubes) < 0.1 Then

                If mConP.s_Droptubes = 0 Then
                    Output = Output & "This configuration will drop " + Str(Cond_DropTubes) + " Condenser tubes." & _
                                     " To accept check the condenser drop tubes box " & _
                                     " To select different configuration Change Total No of feeds" & vbCrLf
                Else
                    Cond_Passess = (TotalTubes - Cond_DropTubes) / mConP.i_CKT
                    Cond_PassRound = Int(Cond_Passess)
                    Output = Output & "This configuration will drop " + Str(Cond_DropTubes) + " Condenser tubes." & vbCrLf
                End If

            Else
                Output = Output & "WARNING: This Condenser Configuration require dropping more than 10% of tubes!" & _
                                 " If you chose dropping tubes capacity is not accurate! " & _
                                 " Recommendation: Change No of feeds" & vbCrLf
            End If
        End If

        If (Cond_Passess / 2) <> Int(Cond_PassRound / 2) Then
            CondOpp = True
            Call CondFeeds(TotalTubes, Output)
            '                    Output = Output & "WARNING: This Condenser Configuration uses Opposite end connection!" & _
            '                              " Please change No of feeds" & vbCrLf
        End If

        If Not CondOpp Then
            If Cond_DropTubes <> 0 Then
                If mConP.s_Droptubes = 1 Then
                    ' Calculate
                    CondCirCuitryCheck = True
                Else
                    ' Return to form
                    CondCirCuitryCheck = False
                End If
            Else
                ' Calculate
                CondCirCuitryCheck = True
            End If

        Else
            CondCirCuitryCheck = False

        End If

        If Output <> "" Then
            MsgBox(Output, vbOKOnly, "Change  Condenser Circuitry")
        End If

    End Function

    Private Sub CondFeeds(ByVal TotalTubes, ByRef Output)
        Dim Passess As Double
        Dim Tubes As Integer
        Dim PassRound As Integer
        Dim I As Integer

        Output = Output & "WARNING: This Condenser Configuration either uses Opposite end connection or require drop tubes!" & _
                                  " If you chose to change Total No of feeds " & _
                                  " Please select one of the following: " & vbCrLf

        TotalTubes = Int(TotalTubes)

        If (TotalTubes / 2) = Int(TotalTubes / 2) Then
            Tubes = TotalTubes
        Else
            Tubes = TotalTubes - 1
        End If

        For I = 1 To Tubes
            Passess = Tubes / I
            PassRound = Int(Passess)
            If (Passess / 2) = Int(PassRound / 2) Then
                Output = Output & Str(I) & ", "
            End If
        Next I

        Output = Left(Output, Len(Output) - 2)

    End Sub

    Function CondFluidPD(ByVal mn_id, ByVal n_RefMassFlRate, ByVal n_NoCKT, ByVal n_HeaderL, ByVal n_CKTLength, ByVal n_ReLiq, ByVal n_RefLatent, _
    ByVal n_DenV, ByVal n_DenL, ByVal n_VisV, ByVal n_VisL, ByVal n_HeaderId)

        Dim n_BoilingNumber As Double
        Dim n_Friction As Double
        Dim n_MassVel As Double
        Dim inletspevolume As Double
        Dim outletspevolume As Double
        Dim n_FrictionPD As Double
        Dim n_CircuitPD As Double
        Dim outletvaporquality As Double
        Dim inletvaporquality As Double
        Dim n_HeaderVel As Double
        Dim n_ReHeader As Double
        Dim n_HeaderFri As Double
        Dim n_HeaderFriPD As Double
        Dim n_HeaderbendPD As Double
        Dim n_CondLHeaderId As Double
        Dim n_CondLHeaderFri As Double
        Dim n_CondLHeaderFriPD As Double

        outletvaporquality = 0.01
        inletvaporquality = 0.99
        If n_HeaderId >= (2 + 5 / 8) Then
            n_CondLHeaderId = 1 + 5 / 8
        ElseIf n_HeaderId >= (2 + 1 / 8) Then
            n_CondLHeaderId = 1 + 3 / 8
        ElseIf n_HeaderId >= (1 + 5 / 8) Then
            n_CondLHeaderId = 1 + 1 / 8
        ElseIf n_HeaderId >= 1 + 3 / 8 Then
            n_CondLHeaderId = 7 / 8
        ElseIf n_HeaderId >= 1 + 1 / 8 Then
            n_CondLHeaderId = 5 / 8
        ElseIf n_HeaderId >= 5 / 8 Then
            n_CondLHeaderId = 1 / 2
            '            ElseIf n_HedaerId >= 5 / 8 Then
            '                    HeaderId = 5 / 8
        Else
            n_CondLHeaderId = 3 / 8
        End If

        n_BoilingNumber = (inletvaporquality - outletvaporquality) * n_RefLatent * 778 / (n_CKTLength / 12)

        'n_Friction = 0.00228 * (n_ReLiq ^ -0.062) * (n_BoilingNumber ^ 0.211)
        n_Friction = 0.00506 * (n_ReLiq ^ -0.0951) * (n_BoilingNumber ^ 0.1554)
        n_MassVel = n_RefMassFlRate / (0.00545 * n_NoCKT * (mn_id ^ 2) * 3600)
        inletspevolume = (inletvaporquality / n_DenV) + (1 - inletvaporquality) / n_DenL
        outletspevolume = (outletvaporquality / n_DenV) + (1 - outletvaporquality) / n_DenL

        n_FrictionPD = (n_Friction * n_CKTLength * (outletspevolume + inletspevolume)) / mn_id
        n_FrictionPD = n_FrictionPD * (n_MassVel ^ 2) * 0.0069 / 32.2
        If n_FrictionPD >= 15 Then
            n_CircuitPD = n_FrictionPD + ((outletspevolume - inletspevolume) * (n_MassVel ^ 2) * 0.0069 / 32.2)
        ElseIf n_FrictionPD >= 10 Then
            n_CircuitPD = n_FrictionPD + ((outletspevolume - inletspevolume) * (n_MassVel ^ 2) * 0.0069 * 0.5 / 32.2)
        ElseIf n_FrictionPD >= 5 Then
            n_CircuitPD = n_FrictionPD + ((outletspevolume - inletspevolume) * (n_MassVel ^ 2) * 0.0069 * 0.25 / 32.2)
        ElseIf n_FrictionPD >= 1 Then
            n_CircuitPD = n_FrictionPD
        Else
            n_CircuitPD = n_FrictionPD - ((outletspevolume - inletspevolume) * (n_MassVel ^ 2) * 0.0069 / 32.2)
        End If

        n_HeaderVel = (n_RefMassFlRate * 4 * 144) / (3.14 * n_DenV * 3600 * n_HeaderId * n_HeaderId)
        n_ReHeader = (n_RefMassFlRate * 4 * 12) / (3.14 * n_VisV * n_HeaderId)

        If n_ReHeader < 3000 Then
            n_HeaderFri = 64 / n_ReHeader
        Else
            n_HeaderFri = 0.3164 / (n_ReHeader ^ 0.25)
        End If

        n_HeaderFriPD = (n_DenV * 0.0069) * (n_HeaderVel ^ 2) * (n_HeaderL * 1.2 / n_HeaderId) * (n_HeaderFri / 64.4)
        n_HeaderbendPD = 0.2 * (n_DenV * 0.0069) * (n_HeaderVel ^ 2) / 64.4

        n_HeaderVel = (n_RefMassFlRate * 4 * 144) / (3.14 * n_DenL * 3600 * n_CondLHeaderId * n_CondLHeaderId)
        n_ReHeader = (n_RefMassFlRate * 4 * 12) / (3.14 * n_VisL * n_CondLHeaderId)

        If n_ReHeader < 3000 Then
            n_CondLHeaderFri = 64 / n_ReHeader
        Else
            n_CondLHeaderFri = 0.3164 / (n_ReHeader ^ 0.25)
        End If

        n_CondLHeaderFriPD = (n_DenL * 0.0069) * (n_HeaderVel ^ 2) * (n_HeaderL * 1.2 / n_CondLHeaderId) * (n_CondLHeaderFri / 64.4)
        CondFluidPD = n_CircuitPD + n_HeaderFriPD + n_HeaderbendPD + n_CondLHeaderFriPD

    End Function

    Function CondAirPD(ByVal mn_AsubS, ByVal nAmin, ByVal nGc, ByVal nReOD, ByVal nReSl, ByVal ParameterA, ByVal nRows)
        Dim finFriction As Double
        Dim finPD As Double
        Dim q As Double
        Dim r As Double
        Dim S As Double
        Dim T As Double
        Dim u As Double
        Dim Eu As Double
        Dim tubePD As Double

        finFriction = 1.7 * (nReSl ^ -0.5) ' Rich equation taken from Stewarts Thesis
        nGc = nGc / 3600
        finPD = finFriction * (nGc ^ 2) * mn_AsubS / (2 * nAmin * 0.0735)

        If ParameterA >= 2.5 Then
            q = 0.33
            r = 98.9
            S = -14800
            T = 1920000
            u = 86200000
        ElseIf ParameterA >= 2 Then
            q = 0.343
            r = 303
            S = -71700
            T = 8800000
            u = -380000000
        ElseIf ParameterA >= 1.5 Then
            q = 0.203
            r = 2480
            S = -7580000
            T = 10400000000.0#
            u = -4820000000000.0#
        Else
            q = 0.245
            r = 3339
            S = -9840000
            T = 13200000000.0#
            u = -5990000000000.0#
        End If

        Eu = q + (r / nReOD) + (S / (nReOD ^ 2)) + (T / (nReOD ^ 3)) + (u / (nReOD ^ 4))
        tubePD = Eu * (nGc ^ 2) * nRows / (2 * 0.0735)
        CondAirPD = (finPD + tubePD) * 0.0069 * 2.307 * 12 / 32 'inches of water abs
    End Function

    Function HeaderId(ByVal Btuh)

        If mConP.s_Split = "IT" And mConP.i_CompQuantity = 2 Then
            Btuh = Btuh * 0.5
        End If

        If Btuh >= 900000 Then
            HeaderId = 2 + 5 / 8
        ElseIf Btuh >= 540000 Then
            HeaderId = 2 + 1 / 8
        ElseIf Btuh >= 264000 Then
            HeaderId = 1 + 5 / 8
        ElseIf Btuh >= 144000 Then
            HeaderId = 1 + 3 / 8
        ElseIf Btuh >= 96000 Then
            HeaderId = 1 + 1 / 8
        ElseIf Btuh >= 60000 Then
            HeaderId = 7 / 8
        ElseIf Btuh >= 36000 Then
            HeaderId = 5 / 8
        Else
            HeaderId = 1 / 2
        End If

    End Function

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

    Function CondRefCharge(ByVal n_HeaderId, ByVal n_TubesPerCKT)

        Dim n_CKTLength As Double
        Dim mn_id As Double
        Dim n_TubeVolume As Double
        Dim n_HeaderVolume As Double
        Dim n_CoilVolume As Double
        Dim n_CoOD As Double
        Dim n_DenV As Double
        Dim n_DenL As Double
        Dim mn_CT As Double
        Dim Refrigerant As String

        If mConP.s_CoilPat = "6" Then
            n_CoOD = 0.331

        ElseIf mConP.s_CoilPat = "3" Then
            n_CoOD = 0.395

        ElseIf mConP.s_CoilPat = "P" Then
            n_CoOD = 0.52

        ElseIf mConP.s_CoilPat = "5" Then
            n_CoOD = 0.645

        ElseIf mConP.s_CoilPat = "7" Then
            n_CoOD = 0.282

        End If

        Refrigerant = mConP.s_Refrigerant
        mn_CT = 120

        If Refrigerant = "22" Then
            n_DenV = mRefObject.R22_DenV(mn_CT)
            n_DenL = mRefObject.R22_DenL(mn_CT)

        ElseIf Refrigerant = "410A" Then
            n_DenV = mRefObject.R410A_DenV(mn_CT)
            n_DenL = mRefObject.R410A_DenL(mn_CT)

        ElseIf Refrigerant = "404A" Then
            n_DenV = mRefObject.R404A_DenV(mn_CT)
            n_DenL = mRefObject.R404A_DenL(mn_CT)

        ElseIf Refrigerant = "407C Mid Pt" Or Refrigerant = "407C Dew Pt" Then
            n_DenV = mRefObject.R407C_DenV(mn_CT)
            n_DenL = mRefObject.R407C_DenL(mn_CT)

        ElseIf Refrigerant = "134A" Then
            n_DenV = mRefObject.R134A_DenV(mn_CT)
            n_DenL = mRefObject.R134A_DenL(mn_CT)

        ElseIf Refrigerant = "507" Then
            n_DenV = mRefObject.R507_DenV(mn_CT)
            n_DenL = mRefObject.R507_DenL(mn_CT)

        End If

        mn_id = n_CoOD - 2 * mConP.n_WallThk
        n_CKTLength = (n_TubesPerCKT * mConP.n_FL / 12) ' in ft
        n_TubeVolume = 3.14 * ((mn_id / 24) ^ 2) * n_CKTLength * mConP.i_CKT
        n_HeaderVolume = 2 * 3.14 * ((n_HeaderId / 24) ^ 2) * (mConP.n_FH / 12)
        n_CoilVolume = n_TubeVolume + n_HeaderVolume
        CondRefCharge = n_CoilVolume * 0.5 * (n_DenV + n_DenL)

    End Function

End Class
