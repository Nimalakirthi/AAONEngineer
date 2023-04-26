Imports System.Math

Public Structure Evap
    Public i_FPI As Integer    'fins per inch
    Public i_Rows As Integer    'rows deep

    Public n_FH As Double     'fin height
    Public n_FinThk As Double     'fin thickness
    Public s_FinMat As String     'fin material
    Public n_FL As Double     'fined length
    Public n_WallThk As Double     'tube wall thickness

    Public s_CoilPat As String     'Coil pattern based on OD
    Public s_Model As String     'coil model number
    Public s_Split As String     'IT, NS FS etc
    Public s_Droptubes As String
    Public i_CKT As Integer    'No of Feeds
    Public i_CKT1 As Integer    'Feeds on Circuit 1
    Public i_CKT2 As Integer    'Feeds on Circuit 2

    Public n_AsubS As Double
    Public n_AsubP As Double
    Public n_AsubO As Double
    Public n_AsubI As Double

    Public n_BTUH As Double      'Total Heat capacity
    Public n_Latent As Double      'Latent Heat Capacity
    Public n_Sensible As Double      'Sensible Heat Capacity
    Public n_LaDB As Double      'Leaving Air Dry Bulb
    Public n_LaWB As Double      'Leaving Air Wet Bulb
    Public n_EvapPD As Double      'Evaporator Pressure Drop
    Public n_HeaderId As Double
    Public n_AirPD As Double      'Evaporator Air PD

    Public s_Refrigerant As String
    Public n_EvEaDB As Double
    Public n_EvEaWB As Double
    Public n_EvACFM As Double
    Public i_CompQuantity As Integer
    Public i_EvapQuantity As Integer
    Public s_CompAppCode As String
    Public n_Altitude As Double

End Structure

Public Class EvapClass

    Private mEvaP As Evap

    Public Property EvapProperties() As Evap
        Get
            Return mEvaP
        End Get

        Set(ByVal NewEvap As Evap)
            mEvaP = NewEvap
        End Set

    End Property

    Public Sub PassMixAirInfo(ByVal EvEaDB As Double, ByVal EvEaWB As Double, ByVal EvACFM As Double, ByVal AppCode As String)

        mEvaP.n_EvEaDB = EvEaDB
        mEvaP.n_EvEaWB = EvEaWB
        mEvaP.n_EvACFM = EvACFM
        mEvaP.s_CompAppCode = AppCode

    End Sub
    Public Sub PassAirCorrection(ByVal EvEaDB As Double, ByVal EvEaWB As Double, ByVal EvACFM As Double, ByVal AppCode As String)

        mEvaP.n_EvEaDB = EvEaDB
        mEvaP.n_EvEaWB = EvEaWB
        mEvaP.n_EvACFM = EvACFM
        mEvaP.s_CompAppCode = AppCode

    End Sub

    Public Sub EvapCalc(ByVal mn_CompBtuh As Double, ByVal mn_MassFlow As Double, ByVal mn_ET As Double, ByVal mn_CT As Double)

        Dim mRefObject As pjRefProp.RefProp
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
        Dim mn_LatDB As Double         'Leaving Air Dry Bulb Temperature
        Dim mn_LatWB As Double         'Leaving Air Wet Bulb Temperature
        Dim n_FaceArea As Double         'Coil face Area
        Dim mn_AirVelocity As Double
        Dim n_EvSt As Double         'Evaporator Transverse Tube Spacing
        Dim n_EvSl As Double         'Evaporator Longitudinal Tube Spacing
        Dim n_Evnt As Double
        Dim n_cpair As Double         'Heat Capacity of humid air
        Dim n_Haewb As Double         'Entering air enthalphy
        Dim n_EvOD As Double         'Evaporator Tube Outer Diameter
        Dim n_AirDensity As Double
        Dim nActDensity As Double
        Dim nMaxAirvelocity1 As Double
        Dim nReOD As Double       'Air Side Reynolds number based on OD
        Dim nReSl As Double        'Air Side Reynolds number based on longitudinal tube spacing
        Dim nJP As Double      'Parameter JP in McQuinston Correlation
        Dim nGc As Double      'Air Mass velocity at minimum flow area
        Dim nho As Double      'Heat transfer coefficient of the outer suf of coil
        Dim nRf As Double      'R Value of Fins (to be evaluated explicitly in future)
        Dim nfinef As Double      'Fin efficiency
        Dim nfinsef As Double      'Fin Surface efficiency
        Dim nxx As Double      'intermediate variable
        Dim nyy As Double      'intermediate variable
        Dim nvv As Double      'intermediate variable
        Dim nzz As Double      'intermediate variable
        Dim nUA As Double      'Overall heat Trnsfer Coefficient
        Dim nCa As Double      'Heat Capacity Rate of entering Air
        Dim nNTU As Double      'Number of Transfer Units
        Dim nEveff As Double      'Effectiveness of Evaporator coil
        Dim Ny As Double      'Equation 13 of ASHRAE 21.11
        Dim n_RefLatent As Double      'Refrigerant latent heat Capacity
        Dim nC As Double      'Coil Characteristic Parameter
        Dim d_temp As Double      'Entering air Dew Point
        Dim nh1dew As Double      'Entering Air Enthalpy at the Dew point
        Dim nHab As Double      'Enthalphy of Air at Dry/Wet boundry
        Dim nSatlblbR As Double      'Saturated grains at dew point
        Dim nSatlblbRtrb As Double      'Saturated Grains at Trb
        Dim nSatlblbRts2 As Double      'Saturated Grains at Ts2
        Dim nHsb As Double      'Enthalpy of Air at fin surface dry/wet boundry
        Dim nTrb As Double      'Temperature at refrigerant side dry/wet boundry
        Dim nHrb As Double      'Enthalpy of Air corrosponding to nTrb,nGrainsTrb
        Dim nts2 As Double      'Fin Surface temperature
        Dim nhs2 As Double      'Air Enthalphy corrosponding to fin Surface conditions
        Dim nHal As Double      'Leaving Air Enthalphy
        Dim nTal As Double      'Leaving Air Temperature
        Dim n_EvapBtuh As Double      'Evaporator BTUH Capacity
        Dim n_Delta1 As Double      'Difference between Compressor and vaporatoe Capacity
        Dim n_Delta2 As Double      'Step addition to evaporating temperature
        Dim i_LoopCount As Integer     'Count for Refrigerant side Gradient Steps
        Dim i_LoopCount4 As Integer     'Count for the leaving air wet bulb
        Dim n_Percent3 As Double      'Percent Deviation of Evap & Comp Capacities
        Dim n_Percent4 As Double      'Percent deviation for wet bulb calculation
        Dim nTolerancePercent3 As Double      'Tolerance Percentage
        Dim n_ET2 As Double      'Tr2
        Dim n_LlblbR As Double      'Leaving Air lb/lb moisture ratio
        Dim n_LWBlblbSR As Double      'Leaving Air lb/lb saturation moisture ratio at wet bulb
        Dim n_LDBlblbSR As Double      'Leaving Air lb/lb saturation moisture ratio at dry bulb
        Dim n_ElblbR As Double      'Entering air lb/lb moisture ratio
        Dim n_EWBlblbSR As Double      'Entering air lb/lb saturation moisture ratio at wet bulb
        Dim n_EDBlblbSR As Double      'Entering air lb/lb saturation moisture ratio at dry bulb
        Dim n_LaIva As Double      'Intermediate variable
        Dim n_EHumidity As Double      'Entering Air Humidity
        Dim n_IntV1 As Double      'Intermidiate Variable
        Dim n_ESatWVPre As Double      'Entering Air Saturated Water Vapor Presure
        Dim n_EWVPre As Double      'Enterinng Air Water Vapor Pressure
        Dim n_LHumidity As Double      'Leaving Air Humidity
        Dim n_LsatWVPre As Double      'Leaving Air Saturated Water Vapor Pressure
        Dim n_LWVPre As Double      'Leaving Air Water vapor Pressure
        Dim n_LairEn As Double      'Leaving air Enthalphy for calculation of wet bulb
        Dim nTab As Double      'Air temperature at dry/Wet boundry
        Dim n_Sensible As Double
        Dim n_SenHRatio As Double      'Sensible Heat ratio
        Dim nhi As Double      'Inside Area Heat Transfer Coeficient
        Dim n_DenV As Double      'Refrigerant Vapor Density
        Dim n_DenL As Double      'Refrigerant Liquid Density
        Dim n_VisV As Double      'Refrigerant vapor Viscosity
        Dim n_VisL As Double      'Refrigerant Liquid viscosity
        Dim n_kL As Double      'Refrigerant liquid Conductivity
        Dim n_CpL As Double      'Refrigerant liquid Heat Capacity
        Const n_cp As Double = 0.24
        Dim n_ReLiq As Double       'Liquid Reynolds number
        Dim n_PrLiq As Double       'Liquid Prandalt's number
        Dim nM As Double       'Intermediate Variable
        Dim nNTUAo As Double       ' Air Side Number of Transfer units
        Dim nEAo As Double       ' Air Side Effectiveness
        Dim b_IsWetCoil As Boolean
        Dim b_IsDryCoil As Boolean
        Dim b_IsDryWetCoil As Boolean
        Dim Altitude As Double
        Dim n_ActDensity As Double
        Dim n_AvgAirMoisture As Double
        Dim n_AirHeatCapacity As Double
        Dim n_PRA As Double       'Prandalt's number of Air
        Dim n_fink As Double       'Fin Conductivity
        Dim n_NoCKT As Integer      'Number of Circuits
        Dim n_RefMassFlRate As Double
        Dim n_TubesPerCKT As Double
        Dim n_CKTLength As Double
        Dim n_TubeHi As Double
        Dim mnSrb As Double
        Dim n_Raw As Double       'Wet Surface Resitance
        Dim nRt As Double       'Metal Thermal resistance
        Dim n_HeaderId As Double
        Dim n_HeaderL As Double
        Dim Si As Double
        Dim beta As Double
        Dim eqradius As Double
        Dim SurParameter As Double
        Dim resnumber As Double
        Dim tanheff As Double
        Dim tempfine As Double
        Dim nhDew_1 As Double
        Dim nhDew_0 As Double
        Dim nSatSlope As Double
        Dim nAmin As Double        'Minimum flow area
        Dim ParameterA As Double
        Dim OD_Btuh As Double
        Dim nRows As Integer
        Dim Refrigerant As String
        Dim CompQuantity As Integer
        Dim EvapQuantity As Integer

        Const inletvaporquality As Double = 0.2
        Const outletvaporquality As Double = 0.99

        Const mn_a1 As Double = 0.00080246
        Const mn_a2 As Double = 0.000024525
        Const mn_a3 As Double = 2.542 * (10 ^ -6)
        Const mn_a4 As Double = -2.5855 * (10 ^ -8)
        Const mn_a5 As Double = 4.038 * (10 ^ -10)
        Dim mn_PSIA As Double

        mRefObject = New pjRefProp.RefProp

        Refrigerant = mEvaP.s_Refrigerant
        mn_EatDB = mEvaP.n_EvEaDB
        mn_EatWB = mEvaP.n_EvEaWB
        mn_ACFM = mEvaP.n_EvACFM
        CompQuantity = mEvaP.i_CompQuantity
        EvapQuantity = mEvaP.i_EvapQuantity
        Altitude = mEvaP.n_Altitude

        If Refrigerant = "22" Then
            n_DenV = mRefObject.R22_DenV(mn_ET)
            n_DenL = mRefObject.R22_DenL(mn_ET)
            n_VisV = mRefObject.R22_VisV(mn_ET)
            n_VisL = mRefObject.R22_VisL(mn_ET)
            n_kL = mRefObject.R22_kL(mn_ET)
            n_CpL = mRefObject.R22_CpL(mn_ET)

        ElseIf Refrigerant = "410A" Then
            n_DenV = mRefObject.R410A_DenV(mn_ET)
            n_DenL = mRefObject.R410A_DenL(mn_ET)
            n_VisV = mRefObject.R410A_VisV(mn_ET)
            n_VisL = mRefObject.R410A_VisL(mn_ET)
            n_kL = mRefObject.R410A_kL(mn_ET)
            n_CpL = mRefObject.R410A_CpL(mn_ET)

        ElseIf Refrigerant = "407C Mid Pt" Or Refrigerant = "407C Dew Pt" Then
            n_CpL = mRefObject.R407C_CpL(mn_ET)
            n_DenV = mRefObject.R407C_DenV(mn_ET)
            n_DenL = mRefObject.R407C_DenL(mn_ET)
            n_VisV = mRefObject.R407C_VisV(mn_ET)
            n_VisL = mRefObject.R407C_VisL(mn_ET)
            n_kL = mRefObject.R407C_kL(mn_ET)

        ElseIf Refrigerant = "404A" Then
            n_DenV = mRefObject.R404A_DenV(mn_ET)
            n_DenL = mRefObject.R404A_DenL(mn_ET)
            n_VisV = mRefObject.R404A_VisV(mn_ET)
            n_VisL = mRefObject.R404A_VisL(mn_ET)
            n_kL = mRefObject.R404A_kL(mn_ET)
            n_CpL = mRefObject.R404A_CpL(mn_ET)

        ElseIf Refrigerant = "134A" Then
            n_kL = mRefObject.R134A_kL(mn_ET)
            n_CpL = mRefObject.R134A_CpL(mn_ET)
            n_DenV = mRefObject.R134A_DenV(mn_ET)
            n_DenL = mRefObject.R134A_DenL(mn_ET)
            n_VisV = mRefObject.R134A_VisV(mn_ET)
            n_VisL = mRefObject.R134A_VisL(mn_ET)

        ElseIf Refrigerant = "507" Then
            n_kL = mRefObject.R507_kL(mn_ET)
            n_CpL = mRefObject.R507_CpL(mn_ET)
            n_DenV = mRefObject.R507_DenV(mn_ET)
            n_DenL = mRefObject.R507_DenL(mn_ET)
            n_VisV = mRefObject.R507_VisV(mn_ET)
            n_VisL = mRefObject.R507_VisL(mn_ET)

        End If


        If mEvaP.s_Split = "IT" And CompQuantity = 1 Then

            If mEvaP.s_CoilPat = "6" Then
                n_EvOD = 0.331
                n_EvSt = 1 * 2
                n_EvSl = 0.625 * 2

            ElseIf mEvaP.s_CoilPat = "3" Then
                n_EvOD = 0.395
                n_EvSt = 1 * 2
                n_EvSl = 0.866 * 2

            ElseIf mEvaP.s_CoilPat = "P" Then
                n_EvOD = 0.52
                n_EvSt = 1.25 * 2
                n_EvSl = 1.0825 * 2

            ElseIf mEvaP.s_CoilPat = "5" Then
                n_EvOD = 0.645
                n_EvSt = 1.5 * 2
                n_EvSl = 1.299 * 2

            ElseIf mEvaP.s_CoilPat = "1" Then
                n_EvOD = 1
                n_EvSt = 3 * 2
                n_EvSl = 1.5 * 2

            ElseIf mEvaP.s_CoilPat = "7" Then
                n_EvOD = 0.282
                n_EvSl = 0.472 * 2  'tube stagger
                n_EvSt = 0.827 * 2

            End If

        ElseIf mEvaP.s_Split = "IT" And CompQuantity = 0 Then

            If mEvaP.s_CoilPat = "6" Then
                n_EvOD = 0.331
                n_EvSt = 1 * 2
                n_EvSl = 0.625 * 2

            ElseIf mEvaP.s_CoilPat = "3" Then
                n_EvOD = 0.395
                n_EvSt = 1 * 2
                n_EvSl = 0.866 * 2

            ElseIf mEvaP.s_CoilPat = "P" Then
                n_EvOD = 0.52
                n_EvSt = 1.25 * 2
                n_EvSl = 1.0825 * 2

            ElseIf mEvaP.s_CoilPat = "5" Then
                n_EvOD = 0.645
                n_EvSt = 1.5 * 2
                n_EvSl = 1.299 * 2

            ElseIf mEvaP.s_CoilPat = "1" Then
                n_EvOD = 1
                n_EvSt = 3 * 2
                n_EvSl = 1.5 * 2

            ElseIf mEvaP.s_CoilPat = "7" Then
                n_EvOD = 0.282
                n_EvSl = 0.472 * 2  'tube stagger
                n_EvSt = 0.827 * 2

            End If

        Else

            If mEvaP.s_CoilPat = "6" Then
                n_EvOD = 0.331
                n_EvSt = 1
                n_EvSl = 0.625

            ElseIf mEvaP.s_CoilPat = "3" Then
                n_EvOD = 0.395
                n_EvSt = 1
                n_EvSl = 0.866

            ElseIf mEvaP.s_CoilPat = "P" Then
                n_EvOD = 0.52
                n_EvSt = 1.25
                n_EvSl = 1.0825

            ElseIf mEvaP.s_CoilPat = "5" Then
                n_EvOD = 0.645
                n_EvSt = 1.5
                n_EvSl = 1.299

            ElseIf mEvaP.s_CoilPat = "1" Then
                n_EvOD = 1
                n_EvSt = 3
                n_EvSl = 1.5

            ElseIf mEvaP.s_CoilPat = "7" Then
                n_EvOD = 0.282
                n_EvSl = 0.472    'tube stagger
                n_EvSt = 0.827

            End If

        End If

        nTolerancePercent3 = 0.01

        If mEvaP.s_FinMat = "A" Then
            n_fink = 127      'AL(1000)
        Else
            n_fink = 226     'CU(12200)
        End If

        n_Evnt = mEvaP.n_FH * mEvaP.i_Rows / n_EvSt
        n_Am = 3.14 * n_EvOD * n_Evnt * mEvaP.n_FL
        n_FaceArea = (mEvaP.n_FL * mEvaP.n_FH) / 144
        nRows = mEvaP.i_Rows
        '     mn_AirVelocity = mn_ACFM / n_FaceArea

        '     If mEvaP.s_Split = "FS" Then
        '        mn_AirVelocity = mn_AirVelocity * 0.5
        '     End If

        If mEvaP.s_CompAppCode = "AH" Or mEvaP.s_CompAppCode = "A5" Then
            mn_ACFM = mn_ACFM * 0.5
        ElseIf mEvaP.s_CompAppCode = "A6" Then
            mn_ACFM = mn_ACFM * 0.66
        ElseIf mEvaP.s_CompAppCode = "AL" Then
            mn_ACFM = mn_ACFM * 0.55
        ElseIf mEvaP.s_CompAppCode = "AS" Then
            mn_ACFM = mn_ACFM * 0.45
        End If

        If EvapQuantity = 1 Then

            If mEvaP.s_Split = "IT" And CompQuantity = 2 Then
                n_RefMassFlRate = 2 * mn_MassFlow
                mn_CompBtuh = 2 * mn_CompBtuh
                n_NoCKT = mEvaP.i_CKT
                mn_AirVelocity = mn_ACFM / n_FaceArea

            ElseIf mEvaP.s_Split = "IT" And CompQuantity = 1 Then
                mn_AirVelocity = mn_ACFM / n_FaceArea
                n_RefMassFlRate = mn_MassFlow
                mn_ACFM = mn_ACFM * 0.5
                If mEvaP.i_CKT1 > mEvaP.i_CKT2 Then
                    n_NoCKT = mEvaP.i_CKT1
                Else
                    n_NoCKT = mEvaP.i_CKT2
                End If

            ElseIf mEvaP.s_Split = "IT" And CompQuantity = 0 Then
                mn_AirVelocity = mn_ACFM / n_FaceArea
                n_RefMassFlRate = mn_MassFlow
                mn_ACFM = mn_ACFM * 0.5
                If mEvaP.i_CKT1 > mEvaP.i_CKT2 Then
                    n_NoCKT = mEvaP.i_CKT1
                Else
                    n_NoCKT = mEvaP.i_CKT2
                End If

            ElseIf mEvaP.s_Split = "FS" Then
                n_RefMassFlRate = mn_MassFlow
                mn_ACFM = mn_ACFM * 0.5
                n_NoCKT = mEvaP.i_CKT * 0.5
                mn_AirVelocity = mn_ACFM / n_FaceArea

            ElseIf mEvaP.s_Split = "NS" Then
                mn_AirVelocity = mn_ACFM / n_FaceArea
                n_RefMassFlRate = mn_MassFlow
                n_NoCKT = mEvaP.i_CKT
            End If

        Else

            n_RefMassFlRate = mn_MassFlow
            n_NoCKT = mEvaP.i_CKT
            mn_ACFM = mn_ACFM * 0.5
            mn_AirVelocity = mn_ACFM / n_FaceArea

        End If


        mn_AsubS = mEvaP.n_AsubS
        mn_AsubP = mEvaP.n_AsubP
        mn_AsubO = mEvaP.n_AsubO
        mn_AsubI = mEvaP.n_AsubI
        mn_FinThk = mEvaP.n_FinThk
        mn_id = n_EvOD - 2 * mEvaP.n_WallThk
        mn_PSIA = maPSIA(Altitude)

        If Refrigerant = "410A" Then
            n_RefLatent = mRefObject.Sysbal_R410A_HFG(mn_ET)
        ElseIf Refrigerant = "22" Then
            n_RefLatent = mRefObject.Sysbal_R22_HG(mn_ET) - mRefObject.Sysbal_R22_HL(mn_ET)
        ElseIf Refrigerant = "407C Mid Pt" Or Refrigerant = "407C Dew Pt" Then
            n_RefLatent = mRefObject.Sysbal_R407C_HFG(mn_ET)
        ElseIf Refrigerant = "404A" Then
            n_RefLatent = mRefObject.Sysbal_R404A_HFG(mn_ET)
        ElseIf Refrigerant = "134A" Then
            n_RefLatent = mRefObject.Sysbal_R134A_HFG(mn_ET)
        ElseIf Refrigerant = "507" Then
            n_RefLatent = mRefObject.Sysbal_R404A_HFG(mn_ET)
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
        n_EDBlblbSR = mn_a1 + (mn_a2 * mn_EatDB) + (mn_a3 * mn_EatDB ^ 2) + (mn_a4 * mn_EatDB ^ 3) + (mn_a5 * mn_EatDB ^ 4)
        n_ESatWVPre = (mn_PSIA * n_EDBlblbSR) / (n_EDBlblbSR + 0.62198)
        n_EWVPre = (mn_PSIA * n_ElblbR) / (n_ElblbR + 0.62198)
        n_EHumidity = n_EWVPre / n_ESatWVPre

        n_AirDensity = maDensity(mn_EatDB, n_ElblbR, mn_PSIA)
        n_cpair = n_cp + 0.444 * n_ElblbR

        Select Case mEvaP.s_CoilPat
            Case Is = "P"
                nMaxAirvelocity1 = mn_AirVelocity * 1.25 / (1.25 - n_EvOD)     'ft/Min
                nReOD = n_AirDensity * (nMaxAirvelocity1 * 60) * (n_EvOD / 12) / 0.0437      'Air Dynamic viscosity=0.0437lb/ft-hr
                nAmin = n_FaceArea * (1.25 - n_EvOD) / 1.25                                                                       'nReOD is Air Side Reynolds number
                ParameterA = 1.25 / n_EvOD
                nReSl = n_AirDensity * (nMaxAirvelocity1 * 60) * (1.08 / 12) / 0.0437

            Case Is = "5"
                nMaxAirvelocity1 = mn_AirVelocity * 1.5 / (1.5 - n_EvOD)     'ft/Min
                nReOD = n_AirDensity * (nMaxAirvelocity1 * 60) * (n_EvOD / 12) / 0.0437      'Air Dynamic viscosity=0.0437lb/ft-hr
                nAmin = n_FaceArea * (1.5 - n_EvOD) / 1.5
                ParameterA = 1.5 / n_EvOD
                nReSl = n_AirDensity * (nMaxAirvelocity1 * 60) * (1.299 / 12) / 0.0437

            Case Is = "3"
                nMaxAirvelocity1 = mn_AirVelocity * 1 / (1 - n_EvOD)     'ft/Min
                nReOD = n_AirDensity * (nMaxAirvelocity1 * 60) * (n_EvOD / 12) / 0.0437      'Air Dynamic viscosity=0.0437lb/ft-hr
                nAmin = n_FaceArea * (1 - n_EvOD)
                ParameterA = 1 / n_EvOD
                nReSl = n_AirDensity * (nMaxAirvelocity1 * 60) * (0.866 / 12) / 0.0437

            Case Is = "6"
                nMaxAirvelocity1 = mn_AirVelocity * 1 / (1 - n_EvOD)     'ft/Min
                nReOD = nActDensity * (nMaxAirvelocity1 * 60) * (n_EvOD / 12) / 0.0437     'Air Dynamic viscosity=0.0437lb/ft-hr
                nAmin = n_FaceArea * (1 - n_EvOD)
                ParameterA = 1 / n_EvOD
                nReSl = n_AirDensity * (nMaxAirvelocity1 * 60) * (0.625 / 12) / 0.0437

            Case Is = "7"
                nMaxAirvelocity1 = mn_AirVelocity * 0.827 / (0.827 - n_EvOD)     'ft/Min
                nReOD = n_AirDensity * (nMaxAirvelocity1 * 60) * (n_EvOD / 12) / 0.0437      'Air Dynamic viscosity=0.0437lb/ft-hr
                nAmin = n_FaceArea * (0.827 - n_EvOD) / 0.827
                ParameterA = 0.827 / n_EvOD
                nReSl = n_AirDensity * (nMaxAirvelocity1 * 60) * (0.472 / 12) / 0.0437
        End Select

        nJP = (nReOD ^ -0.4) * ((mn_AsubO / mn_AsubP) ^ -0.15) 'Parameter JP in Chilton-Colburn j-factors
        'nJP = 0.4 * (nReOD ^ (-0.468 + 0.04076 * nRows)) * ((mn_AsubO / mn_AsubP) ^ 0.159) * (nRows ^ -1.261) 'Parameter JP in Chilton-Colburn j-factors, Wang 1997

        nGc = n_AirDensity * (nMaxAirvelocity1 * 60)          'Air Mass Velocity at Minimum Flow Area

        nho = (nGc * n_cpair) * ((0.00125 + 0.275 * nJP) / (n_PRA ^ 0.666)) 'Outer surface area heat transfer coefficient
        'nho = (nGc * n_cpair) * (nJP / (n_PRA ^ 0.666)) 'Outer surface area heat transfer coefficient, Wang 1997

        Si = n_EvSt / n_EvOD
        beta = ((n_EvSt * 0.5) ^ 2) + ((n_EvSl) ^ 2)
        beta = (beta ^ 0.5) / n_EvSt
        eqradius = 1.27 * Si * ((beta - 0.3) ^ 0.5)
        SurParameter = (2 * nho / (n_fink * mEvaP.n_FinThk * 0.0833)) ^ 0.5
        resnumber = (eqradius - 1) * (1 + 0.35 * Log(eqradius))
        tempfine = (Exp(SurParameter * n_EvOD * 0.0416 * resnumber) - Exp(-SurParameter * n_EvOD * 0.0416 * resnumber)) / (Exp(SurParameter * n_EvOD * 0.0416 * resnumber) + Exp(-SurParameter * n_EvOD * 0.0416 * resnumber))
        tanheff = tempfine / (SurParameter * n_EvOD * 0.0416 * resnumber)
        nfinef = tanheff

        nfinsef = (1 - (mn_AsubS / mn_AsubO) * (1 - nfinef)) 'Surface Efficiency

        nxx = nfinsef * nho * mn_AsubO 'Heat Transfer term realted to outer surface area

        nCa = 60 * mn_ACFM * n_AirDensity * n_cpair 'Moist Air heat cpacity rate

        n_ReLiq = (n_RefMassFlRate * 4 * 12 / (3.14 * n_NoCKT * n_VisL * mn_id)) 'Refrigerant side Reynolds numbeer ^ 0.8

        n_PrLiq = (n_VisL * n_CpL / n_kL) 'Refrigerant side Prandalt's number ^ 0.4

        nM = ((n_DenL / n_DenV) ^ 0.375) * ((n_VisV / n_VisL) ^ 0.075) 'Intermediate variable to evaluate refrigerant side heat transfer coefficient

        nhi = 0.0186875 * (n_kL * 12 / mn_id) * (n_ReLiq ^ 0.8) * (n_PrLiq ^ 0.4) * nM 'Average evaporative heat transfer coefficient
        nhi = nhi * n_NoCKT * (outletvaporquality - inletvaporquality) / ((outletvaporquality ^ 0.325) - (inletvaporquality ^ 0.325))

        nyy = nhi * mn_AsubI

        nvv = mEvaP.n_WallThk / (24 * 220) 'TBE: Thermal conductivity of Tube wall 220
        nvv = nvv / n_Am

        nzz = (1 / nxx) + nvv + (1 / nyy) 'Intermediate variable

        nUA = 1 / nzz 'Overall Heat transfer Coefficient

        nNTU = nUA / nCa 'Number of Transfer Units

        nEveff = (1 - Exp(-nNTU)) ' Wet coil Effectivenesss based on overall heat transfer

        nNTUAo = nxx / nCa ' Number of transfer units based on outer surface area

        nEAo = (1 - Exp(-nNTUAo)) 'Coil effectiveness based on outer surface area

        Ny = (60 * mn_ACFM * n_AirDensity) / (n_RefMassFlRate * n_RefLatent) ' Equation 13 of ASHRAE Handbook Systems & Equipment 21.11

        mnSrb = mn_AsubO / mn_AsubI
        nRt = mnSrb * mn_id * Log(n_EvOD / mn_id) / (24 * 226)

        n_Haewb = maEnthalpy(mn_EatDB, n_ElblbR)

        d_temp = maDewPoint(n_EWVPre)

        nSatlblbR = maSatHumidityR(d_temp + 1)
        nhDew_1 = maEnthalpy(d_temp + 1, nSatlblbR)
        nSatlblbR = maSatHumidityR(d_temp - 1)
        nhDew_0 = maEnthalpy(d_temp - 1, nSatlblbR)
        nSatSlope = (nhDew_1 - nhDew_0) / 2

        nRf = ((1 / nfinef) - 1) / nho
        n_Raw = nSatSlope / (nho * n_cpair)
        nC = ((mnSrb / nhi) + nRf + nRt) / (n_cpair * n_Raw) ' Coil Characteristic parameter, Eq 23 of ASHRAE handbook Systems 21.11

        nh1dew = maEnthalpy(d_temp, n_ElblbR)

        i_LoopCount = 0
        n_Delta2 = 0 'Step variable along the Refrigerant side temperature Gradient. Will beincremented stepwise

        n_ET2 = mn_ET ' Refrigerant side Dry Surface temperature is Initial Guesss Evaporation Temperature

        Do While i_LoopCount < 20

            b_IsWetCoil = False
            b_IsDryCoil = False
            b_IsDryWetCoil = False

            n_ET2 = n_ET2 + (n_Delta2 * 0.1)

            nHab = ((nC * nh1dew) + d_temp - n_ET2 + (Ny * n_Haewb)) / (nC + Ny) ' Air Enthalphy at Dry-WEt Boundry Surface. Equation 14 of ASHRAE handbook 21.11
            nTab = mn_EatDB - ((n_Haewb - nHab) / n_cpair) 'Air Temperature at Dry-Wet Boundary

            nSatlblbR = maSatHumidityR(d_temp)
            nHsb = maEnthalpy(d_temp, nSatlblbR) 'air Enthalphy On the Surface at Dew Point

            nTrb = (d_temp - nC * (nHab - nHsb)) ' Refrigerant temperatture at Dry-Wet Boundary

            nSatlblbRtrb = maSatHumidityR(nTrb)
            nHrb = maEnthalpy(nTrb, nSatlblbRtrb) 'Air enthalphy corrosponding to the Refrigerant temperature at Dry-Wet boundry

            'nts2 = (mn_ET + (nC * (1 - nEveff) * (n_Haewb - nHrb))) 
            nts2 = Newton(mn_ET + 1.5, nC, nEveff, n_Haewb)

            nSatlblbRts2 = maSatHumidityR(nts2)
            nhs2 = maEnthalpy(nts2, nSatlblbRts2) 'Enthalphy of Ait at Fin surface temperature

            nHal = (n_Haewb - (n_Haewb - nhs2) * nEveff) ' Leaving Air Enthalphy

            'nTal = (nts2 + ((mn_EatDB - nts2) * Exp(-nNTU)))

            nTal = (nts2 + ((mn_EatDB - nts2) * Exp(-nNTUAo))) 'Leaving Air Temperature, Dry Bulb

            n_SenHRatio = n_cpair * (mn_EatDB - nTal) / (n_Haewb - nHal) ' Sensible Heat Ratio

            mn_LatDB = nTal

            If n_Haewb <= nHab Then
                b_IsWetCoil = True
            ElseIf nHal < nHab And nHab < n_Haewb Then
                b_IsDryWetCoil = True
            Else
                b_IsDryCoil = True
            End If

            If b_IsWetCoil = True Then ' If Surface is fully wetted

                n_EvapBtuh = 60 * n_AirDensity * mn_ACFM * nEveff * (n_Haewb - nhs2)

            ElseIf b_IsDryWetCoil = True Then 'If surface is partially Dry

                n_EvapBtuh = 60 * n_AirDensity * mn_ACFM * (nEveff * (nHab - nhs2) + n_cpair * (mn_EatDB - nTab))
                ' Evaporator BTUH capacity if the surface is partially dry

            Else
                n_EvapBtuh = 60 * n_AirDensity * mn_ACFM * n_cpair * nEAo * (mn_EatDB - nts2)
                'Evaporator Capacity if the surface is completely dry
            End If

            n_Delta1 = mn_CompBtuh - n_EvapBtuh
            '    n_Delta1 = mComPClass.CompProperties.n_BTUH - n_EvapBtuh

            n_Percent3 = (Abs(n_Delta1) / n_EvapBtuh)

            If n_Percent3 <= nTolerancePercent3 Then
                mEvaP.n_BTUH = n_EvapBtuh
                i_LoopCount = 20
            Else
                i_LoopCount = (i_LoopCount + 1)
                n_Delta2 = (n_Delta2 + 0.2)
            End If

        Loop

        n_Sensible = 60 * n_AirDensity * n_cpair * mn_ACFM * (mn_EatDB - mn_LatDB) 'Sensible Heat Capacity

        'evaluate leaving air wet bulb for fully wetted or completely dry surface situations
        n_LlblbR = (nHal - n_cp * mn_LatDB) / (1061 + 0.444 * mn_LatDB)
        n_LDBlblbSR = mn_a1 + (mn_a2 * mn_LatDB) + (mn_a3 * mn_LatDB ^ 2) + (mn_a4 * mn_LatDB ^ 3) + (mn_a5 * mn_LatDB ^ 4)
        n_LaIva = (0.444 * mn_LatDB + 1093) * n_LlblbR
        mn_LatWB = (n_LaIva + (n_cp * mn_LatDB) - (1093 * n_LDBlblbSR)) / (n_cp - (0.556 * n_LDBlblbSR) + n_LlblbR)

        i_LoopCount4 = 0

        Do While i_LoopCount4 < 50

            n_LWBlblbSR = mn_a1 + (mn_a2 * mn_LatWB) + (mn_a3 * mn_LatWB ^ 2) + (mn_a4 * mn_LatWB ^ 3) + (mn_a5 * mn_LatWB ^ 4)
            n_LaIva = ((1093 - 0.556 * mn_LatWB) * n_LWBlblbSR) - (n_cp * (mn_LatDB - mn_LatWB))
            n_LlblbR = n_LaIva / (1093 + 0.444 * mn_LatDB - mn_LatWB)

            n_LsatWVPre = (mn_PSIA * n_LDBlblbSR) / (n_LDBlblbSR + 0.62198)
            n_LWVPre = (mn_PSIA * n_LlblbR) / (n_LlblbR + 0.62198)

            n_LairEn = (n_cp * mn_LatDB) + (n_LlblbR * (1061 + 0.444 * mn_EatDB))
            n_LHumidity = n_LWVPre / n_LsatWVPre

            n_Percent4 = (Abs(n_LairEn - nHal) / nHal)


            If n_Percent4 <= nTolerancePercent3 And mn_LatDB > mn_LatWB Then
                i_LoopCount4 = 51

            ElseIf Abs(mn_LatDB - mn_LatWB) <= 2 And mn_LatDB > mn_LatWB Then
                i_LoopCount4 = 51

            ElseIf (mn_LatDB - mn_LatWB) <= 0 Then
                mn_LatWB = mn_LatWB - 0.1
                i_LoopCount4 = i_LoopCount4 + 1
            ElseIf (mn_LatDB - mn_LatWB) > 2 And (mn_LatDB - mn_LatWB) < 4 Then
                mn_LatWB = mn_LatWB + 0.035
                i_LoopCount4 = i_LoopCount4 + 1
            ElseIf (mn_LatDB - mn_LatWB) > 4 And (mn_LatDB - mn_LatWB) < 6 Then
                mn_LatWB = mn_LatWB + 0.079
                i_LoopCount4 = i_LoopCount4 + 1
            ElseIf (mn_LatDB - mn_LatWB) > 6 And (mn_LatDB - mn_LatWB) < 8 Then
                mn_LatWB = mn_LatWB + 0.119
                i_LoopCount4 = i_LoopCount4 + 1
            ElseIf (mn_LatDB - mn_LatWB) > 8 And (mn_LatDB - mn_LatWB) < 10 Then
                mn_LatWB = mn_LatWB + 0.159
                i_LoopCount4 = i_LoopCount4 + 1
            ElseIf (mn_LatDB - mn_LatWB) > 10 Then
                mn_LatWB = mn_LatWB + 0.195
                i_LoopCount4 = i_LoopCount4 + 1
            End If

        Loop

        mEvaP.n_LaDB = mn_LatDB ' Evaporator leaving Air Dry Bulb
        mEvaP.n_LaWB = mn_LatWB 'Evapaorator leaving air Wet Bulb

        If mn_LatWB > mn_LatDB Then
            mn_LatWB = mn_LatDB
        Else
            mn_LatWB = mn_LatWB
        End If

        n_ActDensity = n_AirDensity
        n_AvgAirMoisture = (n_ElblbR + n_LlblbR) / 2
        n_AirHeatCapacity = n_cp + 0.444 * n_AvgAirMoisture ' Altitude Correction to Heat Capcaity

        If b_IsDryCoil = True Then

            mEvaP.n_Sensible = n_EvapBtuh
        ElseIf b_IsWetCoil = True Then

            mEvaP.n_Sensible = 60 * n_ActDensity * n_AirHeatCapacity * mn_ACFM * (mn_EatDB - mn_LatDB)
        Else
            'mEvaP.n_Sensible = n_EvapBtuh - n_HLatent
            mEvaP.n_Sensible = 60 * n_ActDensity * n_AirHeatCapacity * mn_ACFM * (mn_EatDB - mn_LatDB)
        End If

        mEvaP.n_BTUH = n_EvapBtuh
        mEvaP.n_Latent = mEvaP.n_BTUH - mEvaP.n_Sensible
        OD_Btuh = n_EvapBtuh
        n_TubeHi = mEvaP.n_FH / n_EvSt
        n_TubesPerCKT = n_TubeHi * mEvaP.i_Rows / n_NoCKT
        n_CKTLength = n_TubesPerCKT * mEvaP.n_FL
        mEvaP.n_HeaderId = HeaderId(OD_Btuh)
        n_HeaderId = mEvaP.n_HeaderId
        n_HeaderL = mEvaP.n_FH

        mEvaP.n_EvapPD = EvapFluidPD(mn_id, n_RefMassFlRate, n_NoCKT, n_HeaderL, n_CKTLength, _
                                        n_ReLiq, n_RefLatent, n_DenV, n_DenL, n_VisV, n_VisL, n_HeaderId)

        mEvaP.n_AirPD = EvapAirPD(mn_AsubS, nAmin, nGc, nReOD, nReSl, ParameterA, nRows)

Error28:
        'ms_ErrorMes = "Error Code: " & Err & " - " & Error(Err) & vbCrLf & _
        ' "Location: " & ms_ProcedureLocation
        Exit Sub

    End Sub

    Public Sub StartEvap(ByVal CompQuantity As Integer)

        Dim n_EvOD As Double         'Evaporator Tube OD
        Dim n_EvID As Double
        Dim n_EvSt As Double
        Dim n_EvSl As Double
        Dim n_Evld As Double
        Dim n_Evlc As Double

        Dim n_Evnf As Double
        Dim n_Evnh As Double
        Dim n_Evnt As Double

        If mEvaP.s_Split = "IT" And CompQuantity = 1 Then

            If mEvaP.s_CoilPat = "6" Then
                n_EvOD = 0.331
                n_EvSt = 1 * 2
                n_EvSl = 0.625 * 2

            ElseIf mEvaP.s_CoilPat = "3" Then
                n_EvOD = 0.395
                n_EvSt = 1 * 2
                n_EvSl = 0.866 * 2

            ElseIf mEvaP.s_CoilPat = "P" Then
                n_EvOD = 0.52
                n_EvSt = 1.25 * 2
                n_EvSl = 1.0825 * 2

            ElseIf mEvaP.s_CoilPat = "5" Then
                n_EvOD = 0.645
                n_EvSt = 1.5 * 2
                n_EvSl = 1.299 * 2

            ElseIf mEvaP.s_CoilPat = "7" Then
                n_EvOD = 0.282
                n_EvSl = 0.472 * 2  'tube stagger
                n_EvSt = 0.827 * 2

            End If

        ElseIf mEvaP.s_Split = "IT" And CompQuantity = 0 Then

            If mEvaP.s_CoilPat = "6" Then
                n_EvOD = 0.331
                n_EvSt = 1 * 2
                n_EvSl = 0.625 * 2

            ElseIf mEvaP.s_CoilPat = "3" Then
                n_EvOD = 0.395
                n_EvSt = 1 * 2
                n_EvSl = 0.866 * 2

            ElseIf mEvaP.s_CoilPat = "P" Then
                n_EvOD = 0.52
                n_EvSt = 1.25 * 2
                n_EvSl = 1.0825 * 2

            ElseIf mEvaP.s_CoilPat = "5" Then
                n_EvOD = 0.645
                n_EvSt = 1.5 * 2
                n_EvSl = 1.299 * 2

            ElseIf mEvaP.s_CoilPat = "7" Then
                n_EvOD = 0.282
                n_EvSl = 0.472 * 2  'tube stagger
                n_EvSt = 0.827 * 2

            End If

        Else

            If mEvaP.s_CoilPat = "6" Then
                n_EvOD = 0.331
                n_EvSt = 1
                n_EvSl = 0.625

            ElseIf mEvaP.s_CoilPat = "3" Then
                n_EvOD = 0.395
                n_EvSt = 1
                n_EvSl = 0.866

            ElseIf mEvaP.s_CoilPat = "P" Then
                n_EvOD = 0.52
                n_EvSt = 1.25
                n_EvSl = 1.0825

            ElseIf mEvaP.s_CoilPat = "5" Then
                n_EvOD = 0.645
                n_EvSt = 1.5
                n_EvSl = 1.299

            ElseIf mEvaP.s_CoilPat = "7" Then
                n_EvOD = 0.282
                n_EvSl = 0.472    'tube stagger
                n_EvSt = 0.827

            End If

        End If

        If mEvaP.s_Split = "FS" Then
            mEvaP.n_FH = mEvaP.n_FH * 0.5
        End If

        n_Evld = mEvaP.i_Rows * n_EvSl
        n_EvID = n_EvOD - (2 * mEvaP.n_WallThk)
        n_Evnh = mEvaP.n_FH * mEvaP.i_Rows / n_EvSt
        n_Evnt = n_Evnh

        n_Evnf = mEvaP.n_FL * mEvaP.i_FPI
        n_Evlc = (mEvaP.n_FL - (n_Evnf * mEvaP.n_FinThk)) / n_Evnf

        mEvaP.n_AsubS = n_Evnf * (((mEvaP.n_FH * n_Evld) / 72) - ((n_Evnh * ((n_EvOD + 2 * mEvaP.n_FinThk) ^ 2)) / 91.68) + _
        (((n_EvOD + 2 * mEvaP.n_FinThk) * (n_Evnh - n_Evnt) * n_Evlc) / 45.84))
        mEvaP.n_AsubP = ((n_Evnt * n_EvOD * mEvaP.n_FL) - (n_Evnt * n_Evnf * mEvaP.n_FinThk) * _
        (n_EvOD - 2 * n_Evlc)) / 45.84

        mEvaP.n_AsubO = mEvaP.n_AsubS + mEvaP.n_AsubP

        mEvaP.n_AsubI = (n_EvID * n_Evnt * mEvaP.n_FL) / 45.84

    End Sub

    Public Function EvapCircuitryCheck(ByVal CompQuantity As Integer) As Boolean

        Dim Evap_TubeHi As Double
        'Dim Evap_CKT1          As Double
        'Dim Evap_CKT2          As Double
        Dim Evap_Passess As Double
        Dim n_EvSt As Double
        Dim TotalTubes As Double
        Dim EvapOpp As Boolean
        Dim Output As String
        Dim Evap_PassRound As Integer
        Dim Evap_TubeHiRound As Integer
        Dim Evap_DropTubes As Integer
        Dim IT_2Comp As Boolean
        Dim FS_2Comp As Boolean

        EvapOpp = False
        IT_2Comp = False
        FS_2Comp = False
        Output = ""

        If mEvaP.s_CoilPat = "6" Then
            n_EvSt = 1
        ElseIf mEvaP.s_CoilPat = "3" Then
            n_EvSt = 1
        ElseIf mEvaP.s_CoilPat = "P" Then
            n_EvSt = 1.25
        ElseIf mEvaP.s_CoilPat = "5" Then
            n_EvSt = 1.5
        ElseIf mEvaP.s_CoilPat = "7" Then
            n_EvSt = 0.827
        End If

        Evap_TubeHi = mEvaP.n_FH / n_EvSt

        Evap_TubeHiRound = Round(Evap_TubeHi)

        If Round(Evap_TubeHi, 4) <> Evap_TubeHiRound Then
            Output = Output & " Evaporator Fin Height must be an integeral multiple of " + Str(n_EvSt) + "" & vbCrLf
        End If

        If mEvaP.s_Split = "IT" And CompQuantity = 2 Then
            IT_2Comp = True
        End If

        If mEvaP.s_Split = "FS" And CompQuantity = 2 Then
            FS_2Comp = True
        End If

        If IT_2Comp Or FS_2Comp Then
            If mEvaP.i_CKT <> (mEvaP.i_CKT1 + mEvaP.i_CKT2) Then
                Output = Output & "Sum of Citcuit 1&2 feeds must be = " + Str(mEvaP.i_CKT) + "" & vbCrLf
            End If
        ElseIf mEvaP.s_Split = "IT" And CompQuantity = 1 Then
            If mEvaP.i_CKT1 <> 0 Then
                Output = Output & "Inactive Citcuit 2 feeds must be = 0 " & vbCrLf
            End If
            If mEvaP.i_CKT2 <> 0 Then
                Output = Output & "Inactive Citcuit 1 feeds must be = 0 " & vbCrLf
            End If

        ElseIf mEvaP.s_Split = "NS" And CompQuantity = 1 Then
            If mEvaP.i_CKT1 <> 0 Then
                If mEvaP.i_CKT <> mEvaP.i_CKT1 Then
                    Output = Output & "Citcuit 1 feeds must be = " + Str(mEvaP.i_CKT) + "" & vbCrLf
                End If
            End If
            If mEvaP.i_CKT2 <> 0 Then
                If mEvaP.i_CKT <> mEvaP.i_CKT2 Then
                    Output = Output & "Citcuit 2 feeds must be = " + Str(mEvaP.i_CKT) + "" & vbCrLf
                End If
            End If
        End If

        TotalTubes = mEvaP.i_Rows * Round(Evap_TubeHi, 4)
        Evap_Passess = TotalTubes / mEvaP.i_CKT
        Evap_PassRound = Int(Evap_Passess)

        Evap_DropTubes = (TotalTubes) - (Evap_PassRound * mEvaP.i_CKT)

        If mEvaP.i_Rows = mEvaP.i_CKT Then
            If (Evap_Passess / 2) <> Int(Evap_Passess / 2) Then
                Evap_DropTubes = mEvaP.i_Rows
            End If
        End If

        If Evap_DropTubes <> 0 Then
            If (Evap_DropTubes / TotalTubes) < 0.1 Then

                If mEvaP.s_Droptubes = 0 Then
                    Output = Output & "This configuration will drop " + Str(Evap_DropTubes) + " Evaporator tubes." & _
                                     " To accept check the Evaporator drop tubes box. " & _
                                     " To select different configuration Change Total No of feeds" & vbCrLf
                Else
                    Evap_Passess = (TotalTubes - Evap_DropTubes) / mEvaP.i_CKT
                    Evap_PassRound = Int(Evap_Passess)
                    Output = Output & "This configuration will drop " + Str(Evap_DropTubes) + " Evaporator tubes." & vbCrLf
                End If

            Else
                Output = Output & "WARNING: This Evaporator Configuration require dropping more than 10% of tubes!" & _
                                     " If you chose dropping tubes capacity is not accurate!" & _
                                     " Recommendation: change Total No of feeds " & vbCrLf
            End If
        End If

        If (Evap_Passess / 2) <> Int(Evap_PassRound / 2) Then
            EvapOpp = True
            Call EvapFeeds(TotalTubes, Output)
        End If

        If Not EvapOpp Then
            If Evap_DropTubes <> 0 Then
                If mEvaP.s_Droptubes = 1 Then
                    ' Calculate
                    EvapCircuitryCheck = True
                Else
                    ' Return to form
                    EvapCircuitryCheck = False
                End If
            Else
                ' Calculate
                EvapCircuitryCheck = True
            End If
        Else
            EvapCircuitryCheck = False
        End If

        If Output <> "" Then
            MsgBox(Output, vbOKOnly, "Change  Evaporator Circuitry")
        End If

    End Function

    Private Sub EvapFeeds(ByVal TotalTubes, ByRef Output)
        Dim Passess As Double
        Dim Tubes As Integer
        Dim PassRound As Integer
        Dim I As Integer

        Output = Output & "WARNING: This Evaporator Configuration either uses Opposite end connection or require drop tubes!" & _
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

    Function EvapFluidPD(ByVal mn_id, ByVal n_RefMassFlRate, ByVal n_NoCKT, ByVal n_HeaderL, ByVal n_CKTLength, ByVal n_ReLiq, ByVal n_RefLatent, _
    ByVal n_DenV, ByVal n_DenL, ByVal n_VisV, ByVal n_VisL, ByVal n_HeaderId)

        Dim n_BoilingNumber As Double
        Dim n_Friction As Double
        Dim n_MassVel As Double
        Dim inletspevolume As Double
        Dim outletspevolume As Double
        Dim n_CircuitPD As Double
        Dim outletvaporquality As Double
        Dim inletvaporquality As Double
        Dim n_HeaderVel As Double
        Dim n_ReHeader As Double
        Dim n_HeaderFri As Double
        Dim n_HeaderFriPD As Double
        Dim n_HeaderbendPD As Double

        outletvaporquality = 0.99
        inletvaporquality = 0.2

        n_BoilingNumber = (outletvaporquality - inletvaporquality) * n_RefLatent * 778 / (n_CKTLength / 12) 'NIST with correct units

        'n_Friction = 0.00228 * (n_ReLiq ^ -0.062) * (n_BoilingNumber ^ 0.211)
        n_Friction = 0.00506 * (n_ReLiq ^ -0.0951) * (n_BoilingNumber ^ 0.1554) ' NIST Correlation
        n_MassVel = n_RefMassFlRate / (0.00545 * n_NoCKT * (mn_id ^ 2) * 3600)
        inletspevolume = (inletvaporquality / n_DenV) + (1 - inletvaporquality) / n_DenL
        outletspevolume = (outletvaporquality / n_DenV) + (1 - outletvaporquality) / n_DenL

        n_CircuitPD = (n_Friction * n_CKTLength * (outletspevolume + inletspevolume)) / mn_id
        n_CircuitPD = (n_CircuitPD + (outletspevolume - inletspevolume)) * ((n_MassVel) ^ 2) * 0.0069 / 32.2 ' NIST PD

        n_HeaderVel = (n_RefMassFlRate * 4 * 144) / (3.14 * n_DenV * 3600 * n_HeaderId * n_HeaderId)
        n_ReHeader = (n_RefMassFlRate * 4 * 12) / (3.14 * n_VisV * n_HeaderId)

        If n_ReHeader < 3000 Then
            n_HeaderFri = 64 / n_ReHeader
        Else
            n_HeaderFri = 0.3164 / (n_ReHeader ^ 0.25)
        End If

        n_HeaderFriPD = (n_DenV * 0.0069) * (n_HeaderVel ^ 2) * (n_HeaderL * 1.2 / n_HeaderId) * (n_HeaderFri / 64.4)
        n_HeaderbendPD = 0.2 * (n_DenV * 0.0069) * (n_HeaderVel ^ 2) / 64.4

        EvapFluidPD = n_CircuitPD + n_HeaderFriPD + n_HeaderbendPD

    End Function

    Function EvapAirPD(ByVal mn_AsubS, ByVal nAmin, ByVal nGc, ByVal nReOD, ByVal nReSl, ByVal ParameterA, ByVal nRows)
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
        EvapAirPD = (finPD + tubePD) * 0.0069 * 2.307 * 12 / 32 'inches of water abs
    End Function

    Function HeaderId(ByVal Btuh)

        If mEvaP.s_Split = "IT" And mEvaP.i_CompQuantity = 2 Then
            Btuh = Btuh * 0.5
        End If

        If Btuh >= 900000 Then
            HeaderId = 3 + 1 / 8
        ElseIf Btuh >= 540000 Then
            HeaderId = 2 + 5 / 8
        ElseIf Btuh >= 264000 Then
            HeaderId = 2 + 1 / 8
        ElseIf Btuh >= 144000 Then
            HeaderId = 1 + 5 / 8
        ElseIf Btuh >= 96000 Then
            HeaderId = 1 + 3 / 8
        ElseIf Btuh >= 60000 Then
            HeaderId = 1 + 1 / 8
        ElseIf Btuh >= 36000 Then
            HeaderId = 7 / 8
        Else
            HeaderId = 5 / 8
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

    Private Function Newton(ByVal mn_ET As Double, ByVal nC As Double, ByVal nEveff As Double, ByVal n_Haewb As Double)
        Const mn_a1 As Double = 0.00080246
        Const mn_a2 As Double = 0.000024525
        Const mn_a3 As Double = 2.542 * (10 ^ -6)
        Const mn_a4 As Double = -2.5855 * (10 ^ -8)
        Const mn_a5 As Double = 4.038 * (10 ^ -10)
        Dim Ts As Double, Tsf As Double
        Dim fTs As Double, dfTs As Double
        Dim Int1 As Double, Int2 As Double
        Dim i_LoopCount As Integer
        Dim Error1 As Double
        Const Percent As Double = 0.0001

        Ts = 50

        i_LoopCount = 0

        Do While i_LoopCount < 100

            Int1 = (1061 * mn_a1) + ((1061 * mn_a2 + 0.444 * mn_a1 + 0.24) * Ts) _
                    + ((1061 * mn_a3 + 0.444 * mn_a2) * (Ts ^ 2)) + ((1061 * mn_a4 + 0.444 * mn_a3) * (Ts ^ 3)) _
                    + ((1061 * mn_a5 + 0.444 * mn_a4) * (Ts ^ 4)) + (0.444 * mn_a5 * (Ts ^ 5))

            Int2 = (1061 * mn_a2 + 0.444 * mn_a1 + 0.24) + ((1061 * mn_a3 + 0.44 * mn_a2) * 2 * Ts) _
                    + ((1061 * mn_a4 + 0.444 * mn_a3) * 3 * (Ts ^ 2)) + ((1061 * mn_a5 + 0.444 * mn_a4) * 4 * (Ts ^ 3)) _
                    + (0.444 * mn_a5 * 5 * (Ts ^ 4))

            fTs = -Ts + mn_ET + (nC * (1 - nEveff) * (n_Haewb - Int1))
            dfTs = -1 - (nC * (1 - nEveff) * Int2)

            Tsf = Ts - (fTs / dfTs)

            Error1 = Abs((Tsf - Ts) / Tsf) * 100

            If Error1 > Percent Then

                Ts = Tsf
                i_LoopCount = i_LoopCount + 1

            Else
                Exit Do
            End If

        Loop

        Newton = Tsf

    End Function

End Class
