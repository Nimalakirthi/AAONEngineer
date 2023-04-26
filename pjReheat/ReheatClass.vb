Imports System.Math

Public Structure Reheat
    Public n_FL As Double           'Fin Length
    Public n_FH As Double
    Public i_FPI As Integer           'Fins Per Inch
    Public s_CoilP As String           'Coil Pattern code
    Public n_FinThk As Double           'Fin Thickness
    Public n_WallThk As Double
    Public i_Rows As Integer          'No of Rows
    Public n_LaDB As Double           'Desired output temperature
    Public n_BTUH As Double
    Public n_DesiredBTUH As Double
    Public n_MassFraction As Double
    Public i_CKT As Integer
    Public s_Droptubes As String
    Public i_CompQuantity As Integer
    Public n_EvACFM As Double
    Public s_Refrigerant As String
    Public s_ReheatType As String
End Structure
Public Class ReheatClass

    Private mRecP As Reheat

    Public Property ReheatProperties() As Reheat
        Get
            Return mRecP
        End Get

        Set(ByVal NewReheat As Reheat)
            mRecP = NewReheat
        End Set
    End Property

    Public Sub PassMixAirCFM(ByVal EvACFM As Double)
        mRecP.n_EvACFM = EvACFM
    End Sub

    Public Sub ReNT(ByVal mn_MassFlow As Double, ByVal mn_CT As Double, ByVal EvLaDB As Double)

        Dim mRefObject As pjRefProp.RefProp
        Dim n_Fs As Double       'Coil Core Surface Area per unit
        Dim n_af As Double       'Fin Surface area per unit
        Dim n_TAa As Double      'Total Surface area of the coil
        Dim n_ai As Double       'inner surface area of the tubes (per unit)
        Dim n_ap As Double       'Primary Surface Area - Per unit
        Dim n_FPM As Double       'Face Velocity
        Dim n_TAo As Double       'Total Outside Surface Area of the coil
        Dim n_TAi As Double       'Total inner surface area of the tubes
        Dim n_TAp As Double       'Total Primary area of the coil
        Dim n_Gc As Double       'Air mass  velocity at minimum Flow area
        Dim n_Rfmax As Double       'Rf_Max, Maximum R factor as ASHRAE
        Dim n_ReAir As Double       'Reynolds number of Air Side
        Dim n_JP As Double       'Parameter JP in Chilton-Colburn j-Factors
        Dim n_ho As Double       'Outside area heat Transfer coefficient
        Dim n_nf As Double       'Fin efficiency
        Dim n_ns As Double       'Fin Surface efficiency
        Dim n_x As Double       'intermediate variable
        Dim n_P As Double       'intermediate varaible
        Dim n_y As Double
        Dim n_z As Double
        Dim n_UA As Double       'Overall Heat Transfer Coefficient
        Dim n_OD As Double
        Dim n_ID As Double       'Tube ID
        Dim n_TrTuSp As Double       'Transverse Tube Spacing
        Dim n_LoTuSp As Double       'Longitudinal Tube Spacing
        Dim n_MaxFPM As Double       'Air Velocoty at minimum flow area
        Dim n_ReRefV As Double       'Refrigerant Liquid Reynold Number with vapor fraction
        Dim n_ReLiq As Double       'Ref liquid Reynolds number with liquid fraction
        Dim n_PRL As Double       'Prandalts number of Liquid Refrigerant
        Dim n_hi As Double       'Inside Area Heat Transfer Coefficient with vapor
        Dim n_hL As Double       'inside area heat transfer coefficient with liquid
        Dim n_hV As Double       'inside heat Transfer coefficient with vapor
        Dim n_AirCa As Double       'Air Heat Capacity rate
        Dim n_NTU As Double       'Number of Transfer Units
        Dim n_NTUAo As Double       'Number of Air Side transfer units
        Dim n_Qs As Double       'Sensible Heat Capacity of the Reheat Coil
        Dim n_Qt As Double       'Total Heat Capacity required
        Dim n_MaxBTUH As Double       'Max Heat Capacity that coil can produce
        Dim n_Tolerance As Double       'Capacity Tolerance
        Dim n_Tmix As Double       'Temperature of the mixture of liquid & Vapor
        Dim n_Multi As Double       'Shah Correlation Multiplier
        Dim C As Double       'Intermediate variable
        Dim nEveff As Double       'Intermediate variable
        Dim nNTUC As Double
        Dim n_TFin As Double       'Fin Surface Temperature
        Dim n_DenV As Double       'Vapor Density
        Dim n_DenL As Double       'Liquid Density
        Dim n_Twi As Double
        Const n_FinCond As Double = 127
        Const n_TuCond As Double = 220
        Const n_Density As Double = 0.075
        Const n_cpa As Double = 0.243
        Dim n_VisL As Double       'Dynamic Viscocity of liquid R22 at 100 F, TBC
        Dim n_VisV As Double       'DYnamic Viscosity of Vapor
        Dim n_CpL As Double       'Heat Capacity of liquid R22 at 100 F
        Dim n_kL As Double       'Thermal Conductivity of Liquid R22 at 100 F
        Dim nEveff1 As Double
        Dim nEveff2 As Double
        Dim Refrigerant As String
        Dim ReheatType As String
        Dim n_FinH As Double
        Dim nEveffCross As Double
        Dim n_QsV As Double
        Dim n_QsL As Double
        Dim MassFraction As Double
        Dim n_RecACFM As Double
        Dim Reyeq As Double
        Dim Densityratio As Double
        Dim vaporquality As Double

        Refrigerant = mRecP.s_Refrigerant
        ReheatType = mRecP.s_ReheatType
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

        ElseIf Refrigerant = "407C Mid Pt" Or Refrigerant = "407C Dew Pt" Then
            n_CpL = mRefObject.R407C_CpL(mn_CT)
            n_DenV = mRefObject.R407C_DenV(mn_CT)
            n_DenL = mRefObject.R407C_DenL(mn_CT)
            n_VisV = mRefObject.R407C_VisV(mn_CT)
            n_VisL = mRefObject.R407C_VisL(mn_CT)
            n_kL = mRefObject.R407C_kL(mn_CT)

        ElseIf Refrigerant = "404A" Then
            n_DenV = mRefObject.R404A_DenV(mn_CT)
            n_DenL = mRefObject.R404A_DenL(mn_CT)
            n_VisV = mRefObject.R404A_VisV(mn_CT)
            n_VisL = mRefObject.R404A_VisL(mn_CT)
            n_kL = mRefObject.R404A_kL(mn_CT)
            n_CpL = mRefObject.R404A_CpL(mn_CT)

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

        If mRecP.s_CoilP = "6" Then
            n_OD = 0.331
            n_ID = 0.331 - 2 * mRecP.n_WallThk
            n_TrTuSp = 1
            n_LoTuSp = 0.625
        ElseIf mRecP.s_CoilP = "3" Then
            n_OD = 0.395
            n_ID = 0.395 - 2 * mRecP.n_WallThk
            n_TrTuSp = 1
            n_LoTuSp = 0.866
        ElseIf mRecP.s_CoilP = "5" Then
            n_OD = 0.645
            n_ID = 0.645 - 2 * mRecP.n_WallThk
            n_TrTuSp = 1.5
            n_LoTuSp = 1.299
        ElseIf mRecP.s_CoilP = "7" Then
            n_OD = 0.282
            n_ID = 0.282 - 2 * mRecP.n_WallThk
            n_TrTuSp = 0.827
            n_LoTuSp = 0.472
        ElseIf mRecP.s_CoilP = "P" Then
            n_OD = 0.52
            n_ID = 0.52 - 2 * mRecP.n_WallThk
            n_TrTuSp = 1.25
            n_LoTuSp = 1.0825
        End If

        If mRecP.s_ReheatType = "VaporParallel" Then
            vaporquality = 0.66
        Else
            vaporquality = 0.33
        End If

        If mRecP.i_CompQuantity = 2 Then
            n_RecACFM = mRecP.n_EvACFM * 0.5
            n_FinH = mRecP.n_FH * 0.5
        Else
            n_RecACFM = mRecP.n_EvACFM
            n_FinH = mRecP.n_FH
        End If

        n_ap = 3.14 * n_OD / n_TrTuSp
        n_ai = 3.14 * n_ID / n_TrTuSp
        n_P = n_LoTuSp - (3.14 * ((0.5 * n_OD) ^ 2) / n_TrTuSp)
        n_af = mRecP.i_FPI * 2 * n_P                  'Multiplication by 2 for both side?
        n_Fs = n_ap + n_af
        n_TAa = n_FinH * mRecP.n_FL / 144
        n_FPM = n_RecACFM / n_TAa
        n_TAo = n_Fs * n_TAa * mRecP.i_Rows
        n_TAi = n_ai * n_TAa * mRecP.i_Rows
        n_TAp = n_ap * n_TAa * mRecP.i_Rows

        n_AirCa = 60 * n_cpa * n_Density * n_RecACFM
        n_Qt = n_AirCa * (mRecP.n_LaDB - EvLaDB)

        If mRecP.s_CoilP = "P" Then
            n_MaxFPM = n_FPM * 1.25 / (1.25 - n_OD)
        ElseIf mRecP.s_CoilP = "3" Then
            n_MaxFPM = n_FPM * 1 / (1 - n_OD)
        ElseIf mRecP.s_CoilP = "5" Then
            n_MaxFPM = n_FPM * 1.5 / (1.5 - n_OD)
        ElseIf mRecP.s_CoilP = "7" Then
            n_MaxFPM = n_FPM * 0.827 / (0.827 - n_OD)
        ElseIf mRecP.s_CoilP = "6" Then
            n_MaxFPM = n_FPM * 1 / (1 - n_OD)
        End If

        n_Gc = n_Density * n_MaxFPM * 60    'Air Mass Velocity at Minimum Flow Area

        n_ReAir = n_Gc * n_OD / (12 * 0.0437)   'nReOD is Air Side Reynolds number

        n_JP = (n_ReAir ^ -0.4) * ((n_Fs * 1.1 / n_ap) ^ -0.15) 'Parameter JP in Chilton-Colburn j-factors

        n_ho = (n_Gc * n_cpa) * ((0.00125 + 0.275 * n_JP) / 0.711 ^ 0.666)  'Outer surface area heat transfer coefficient

        n_Rfmax = 0.5 * ((n_TrTuSp * 0.5) ^ 2) / (mRecP.n_FinThk * n_FinCond * 12)  'R-value related to Fin Efficiency

        n_nf = 1 / (1 + n_ho * n_Rfmax) 'Fin efficiency

        n_ns = 1 - ((n_af / n_Fs) * (1 - n_nf)) 'Surface Efficiency

        n_x = n_ns * n_ho * n_TAo   'Heat Transfer term realted to outer surface area

        MassFraction = mRecP.n_MassFraction

        n_PRL = n_VisL * n_CpL / n_kL 'Prandlt's number of liquid

        n_ReRefV = mn_MassFlow * MassFraction * 4 * 12 / (3.14 * n_ID * n_VisL * mRecP.i_CKT)
        Densityratio = (n_DenL / n_DenV) ^ 0.5
        Reyeq = n_ReRefV * ((1 - vaporquality) + vaporquality * Densityratio * (n_VisV / n_VisL))

        n_hV = 0.05 * mRecP.i_CKT * (Reyeq ^ 0.8) * (n_PRL ^ 0.33) * n_kL * 12 / n_ID

        If ReheatType = "VaporParallel" Or ReheatType = "VaporSeries" Then
            n_hi = n_hV
            n_y = n_hi * n_TAi
            n_z = (1 / n_x) + (1 / n_y) + ((n_OD - n_ID) / (24 * n_TuCond))
            n_UA = 1 / n_z
            n_NTU = n_UA / n_AirCa
            n_NTUAo = n_x / n_AirCa
            n_MaxBTUH = (1 - Exp(-n_NTU)) * n_AirCa * (mn_CT - EvLaDB)
            n_Twi = mn_CT - (n_MaxBTUH / (n_hi * n_TAi))
            n_TFin = n_Twi - (n_MaxBTUH * 12 * n_TrTuSp * Log(n_OD / n_ID) / (n_FinH * mRecP.n_FL * mRecP.i_Rows * 2 * 3.14 * n_TuCond))
            'n_Qs = (1 - Exp(-n_NTUAo)) * (n_TFin - EvLaDB) * n_AirCa
            n_Qs = (1 - Exp(-n_NTU)) * (n_Twi - EvLaDB) * n_AirCa

        Else
            If n_hV <> 0 Then
                n_hi = n_hV
                n_y = n_hi * n_TAi
                n_z = (1 / n_x) + (1 / n_y) + ((n_OD - n_ID) / (24 * n_TuCond))
                n_UA = 1 / n_z
                n_NTU = n_UA / (n_AirCa * MassFraction)
                n_NTUAo = n_x / (n_AirCa * MassFraction)
                n_MaxBTUH = (1 - Exp(-n_NTU)) * n_AirCa * (mn_CT - EvLaDB)
                n_Twi = mn_CT - (n_MaxBTUH / (n_hi * n_TAi))
                n_TFin = n_Twi - (n_MaxBTUH * 12 * n_TrTuSp * Log(n_OD / n_ID) / (n_FinH * mRecP.n_FL * mRecP.i_Rows * 2 * 3.14 * n_TuCond))
                n_QsV = (1 - Exp(-n_NTUAo)) * (n_TFin - EvLaDB) * n_AirCa * MassFraction

            End If

            n_ReLiq = mn_MassFlow * (1 - MassFraction) * 4 * 12 / (3.14 * n_ID * n_VisL * mRecP.i_CKT)
            n_hL = 0.023 * (n_ReLiq ^ 0.8) * (n_PRL ^ 0.4) * n_kL * 12 * (mRecP.i_CKT) / n_ID

            If n_hL <> 0 Then
                n_hi = n_hL
                n_y = n_hi * n_TAi
                n_z = (1 / n_x) + (1 / n_y) + ((n_OD - n_ID) / (24 * n_TuCond))
                n_UA = 1 / n_z
                n_NTU = n_UA / (n_AirCa * (1 - MassFraction))
                n_MaxBTUH = (1 - Exp(-n_NTU)) * n_AirCa * (mn_CT - EvLaDB)
                n_Twi = mn_CT - (n_MaxBTUH / (n_hi * n_TAi))
                n_TFin = n_Twi - (n_MaxBTUH * 12 * n_TrTuSp * Log(n_OD / n_ID) / (n_FinH * mRecP.n_FL * mRecP.i_Rows * 2 * 3.14 * n_TuCond))

                C = n_AirCa / (mn_MassFlow * n_CpL * 60)

                If mRecP.i_Rows >= 2 Then
                    nEveff1 = 1 - (Exp(-n_NTU / 2))
                    nEveff2 = 1 + C * (nEveff1 ^ 2)
                    nEveffCross = 1 - ((Exp(-2 * nEveff1 * C)) * nEveff2)
                    nEveff = nEveffCross / C
                Else
                    nNTUC = n_NTU * (1 - C)
                    nEveff = (1 - Exp(-nNTUC)) / (1 - C * Exp(-nNTUC))
                End If

                n_QsL = nEveff * (n_TFin - EvLaDB) * n_AirCa * (1 - MassFraction)
            End If

            n_Qs = n_QsL + n_QsV

        End If

        n_Tolerance = Abs(n_Qs - n_Qt) / n_Qt

        mRecP.n_DesiredBTUH = n_Qt
        mRecP.n_BTUH = n_Qs
        'mOut.n_Tolerance = n_Tolerance * 100

    End Sub

    Public Function ReheCircuitryCheck() As Boolean
        Dim Rehe_TubeHi As Double
        Dim Rehe_Passess As Double
        Dim n_CoSt As Double
        Dim ReheOpp As Boolean
        Dim Output As String
        Dim Rehe_PassRound As Integer
        Dim Rehe_TubeHiRound As Integer
        Dim Rehe_DropTubes As Integer
        Dim TotalTubes As Double

        Output = ""

        If mRecP.s_CoilP = "6" Then
            n_CoSt = 1
        ElseIf mRecP.s_CoilP = "3" Then
            n_CoSt = 1
        ElseIf mRecP.s_CoilP = "P" Then
            n_CoSt = 1.25
        ElseIf mRecP.s_CoilP = "5" Then
            n_CoSt = 1.5
        ElseIf mRecP.s_CoilP = "7" Then
            n_CoSt = 0.827
        End If

        Rehe_TubeHi = mRecP.n_FH / n_CoSt

        Rehe_TubeHiRound = Round(Rehe_TubeHi)

        If Round(Rehe_TubeHi, 4) <> Rehe_TubeHiRound Then
            Output = Output & " Reheat Fin Height must be an integeral multiple of " + Str(n_CoSt) + "" & vbCrLf
        End If
        TotalTubes = mRecP.i_Rows * Round(Rehe_TubeHi, 4)
        Rehe_Passess = TotalTubes / mRecP.i_CKT
        Rehe_PassRound = Int(Rehe_Passess)

        Rehe_DropTubes = (TotalTubes) - (Rehe_PassRound * mRecP.i_CKT)

        If mRecP.i_Rows = mRecP.i_CKT Then
            If (Rehe_Passess / 2) <> Int(Rehe_Passess / 2) Then
                Rehe_DropTubes = mRecP.i_Rows
            End If
        End If

        If Rehe_DropTubes <> 0 Then

            If (Rehe_DropTubes / TotalTubes) < 0.1 Then

                If mRecP.s_Droptubes = 0 Then
                    Output = Output & "This configuration will drop " + Str(Rehe_DropTubes) + " Reheat tubes." & _
                                     " To accept check the Reheat drop tubes box " & _
                                     " To select different configuration Change Total No of feeds" & vbCrLf
                Else
                    Rehe_Passess = (TotalTubes - Rehe_DropTubes) / mRecP.i_CKT
                    Rehe_PassRound = Int(Rehe_Passess)
                    Output = Output & "This configuration will drop " + Str(Rehe_DropTubes) + " Reheat tubes." & vbCrLf
                End If

            Else
                Output = Output & "WARNING: This Reheat Configuration require dropping more than 10% of tubes!" & _
                                 " If you chose dropping tubes capacity is not accurate! " & _
                                 " Recommendation: Change No of feeds" & vbCrLf
            End If
        End If

        If (Rehe_Passess / 2) <> Int(Rehe_PassRound / 2) Then
            ReheOpp = True
            Call ReheFeeds(TotalTubes, Output)
            '                    Output = Output & "WARNING: This Reheat Configuration uses Opposite end connection!" & _
            '                              " Please change No of feeds" & vbCrLf
        End If

        If Not ReheOpp Then
            If Rehe_DropTubes <> 0 Then
                If mRecP.s_Droptubes = 1 Then
                    ' Calculate
                    ReheCircuitryCheck = True
                Else
                    ' Return to form
                    ReheCircuitryCheck = False
                End If
            Else
                ' Calculate
                ReheCircuitryCheck = True
            End If

        Else
            ReheCircuitryCheck = False

        End If

        If Output <> "" Then
            MsgBox(Output, vbOKOnly, "Change  Reheat Circuitry")
        End If
    End Function

    Private Sub ReheFeeds(ByVal TotalTubes, ByRef Output)
        Dim Passess As Double
        Dim Tubes As Integer
        Dim PassRound As Integer
        Dim I As Integer

        Output = Output & "WARNING: This Reheat Configuration either uses Opposite end connection or require drop tubes!" & _
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

    Public Sub VariationMassFraction(ByVal VariationMassF As Double)
        mRecP.n_MassFraction = VariationMassF
    End Sub

End Class
