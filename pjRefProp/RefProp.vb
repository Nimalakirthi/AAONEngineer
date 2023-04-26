Imports System.Math

Public Class RefProp

    Public Function Sysbal_R410A_HL(ByVal Tsat)

        Const A As Double = 34.42467
        Const B As Double = -221.5447
        Const C As Double = -271.7314
        Const D As Double = -112.8898
        Const E As Double = 452.8092
        Const F As Double = 686.6152
        Const Xo As Double = 0.5541498
        Const Tc As Double = 621.5
        Dim X As Double
        Dim T As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        T = Tsat + 459.67

        X = ((1 - (T / Tc)) ^ 0.333) - Xo

        Sysbal_R410A_HL = A + B * X + C * (X ^ 2) + D * (X ^ 3) + E * (X ^ 4) + F * (X ^ 5)

    End Function
    Public Function Sysbal_R407C_HL(ByVal Tsat)

        Const A As Double = 131.7716
        Const B As Double = 1224.77
        Const C As Double = -4212.446
        Const D As Double = 5045.1332
        Const E As Double = -2528.57
        Const Tc As Double = 359.2
        Dim X As Double
        Dim T As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        T = 273 + ((Tsat - 32) / 1.8)

        X = (1 - (T / Tc)) ^ 0.333

        Sysbal_R407C_HL = (A + B * X + C * (X ^ 2) + D * (X ^ 3) + E * (X ^ 4)) / 2.326

    End Function
    Public Function Sysbal_R404A_HL(ByVal Tsat)

        Const A As Double = 270.9661
        Const B As Double = -170.275
        Const C As Double = -37.7576
        Const D As Double = -272.734
        Const Tc As Double = 344.7
        Dim X As Double
        Dim T As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        T = 273 + ((Tsat - 32) / 1.8)

        X = (1 - (T / Tc)) ^ 0.333

        Sysbal_R404A_HL = (A + B * X + C * (X ^ 2) + D * (X ^ 3)) / 2.326

    End Function
    Public Function Sysbal_R402A_HL(ByVal Tsat)

        Const A As Double = 87.394
        Const B As Double = -63.532
        Const C As Double = 71.991
        Const D As Double = -114.633
        Const E As Double = 29.941
        Const Tc As Double = 167.89
        Dim X As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        X = (1 - (Tsat / Tc)) ^ 0.333

        Sysbal_R402A_HL = A + B * X + C * (X ^ 2) + D * (X ^ 3) + E * (X ^ 4)

    End Function

    Public Function Sysbal_R134A_HL(ByVal Tsat)

        Const A As Double = 249.0896
        Const B As Double = 189.8021
        Const C As Double = -753.47
        Const D As Double = 261.1633
        Const E As Double = -157.687
        Const Tc As Double = 374
        Dim X As Double
        Dim T As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        T = 273 + ((Tsat - 32) / 1.8)

        X = (1 - (T / Tc)) ^ 0.333

        Sysbal_R134A_HL = (A + B * X + C * (X ^ 2) + D * (X ^ 3) + E * (X ^ 4)) / 2.326

    End Function

    Public Function Sysbal_R22_HG(ByVal Tsat)

        Const sv_b1 As Double = 104.465
        Const sv_b2 As Double = 0.098445
        Const sv_b3 As Double = -0.0001226
        Const sv_b4 As Double = -0.0000009861

        If Tsat > 155 Then
            Tsat = 155
        End If

        Sysbal_R22_HG = sv_b1 + (sv_b2 * Tsat) + (sv_b3 * (Tsat ^ 2)) + (sv_b4 * (Tsat ^ 3))

    End Function
    Public Function Sysbal_R22_HL(ByVal Tsat)

        Const sl_a1 As Double = 10.409
        Const sl_a2 As Double = 0.26851
        Const sl_a3 As Double = 0.00014794
        Const sl_a4 As Double = 0.00000053429

        If Tsat > 155 Then
            Tsat = 155
        End If

        Sysbal_R22_HL = sl_a1 + (sl_a2 * Tsat) + (sl_a3 * (Tsat ^ 2)) + (sl_a4 * (Tsat ^ 3))

    End Function

    Public Function Sysbal_R410A_HFG(ByVal Tsat)

        Const B As Double = 314.9286
        Const C As Double = -86.3852
        Const D As Double = 612.07716
        Const E As Double = -515.30242
        Const Tc As Double = 344
        Dim X As Double
        Dim T As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        T = 273 + ((Tsat - 32) / 1.8)

        X = (1 - (T / Tc)) ^ 0.333

        Sysbal_R410A_HFG = (B * X + C * (X ^ 2) + D * (X ^ 3) + E * (X ^ 4)) / 2.326

    End Function

    Public Function Sysbal_R404A_HFG(ByVal Tsat)
        Const A As Double = 67.117635
        Const B As Double = -492.87541
        Const C As Double = 2639.8565
        Const D As Double = -3589.4277
        Const E As Double = 1733.943
        Const Tc As Double = 344.7
        Dim X As Double
        Dim T As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        T = 273 + ((Tsat - 32) / 1.8)

        X = (1 - (T / Tc)) ^ 0.333

        Sysbal_R404A_HFG = (A + B * X + C * (X ^ 2) + D * (X ^ 3) + E * (X ^ 4)) / 2.326

    End Function
    Public Function Sysbal_R402A_HFG(ByVal Tsat)
        Const A As Double = 14.131
        Const B As Double = -56.59
        Const C As Double = 265.021
        Const D As Double = -195.811
        Const E As Double = 48.164
        Const Tc As Double = 167.89
        Dim X As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        X = (1 - (Tsat / Tc)) ^ 0.333

        Sysbal_R402A_HFG = A + B * X + C * (X ^ 2) + D * (X ^ 3) + E * (X ^ 4)

    End Function
    Public Function Sysbal_R407C_HFG(ByVal Tsat)
        Const A As Double = 218.8494
        Const B As Double = -1683.224
        Const C As Double = 6325.3877
        Const D As Double = -8231.939
        Const E As Double = 3884.46
        Const Tc As Double = 359.2
        Dim X As Double
        Dim T As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        T = 273 + ((Tsat - 32) / 1.8)

        X = (1 - (T / Tc)) ^ 0.333

        Sysbal_R407C_HFG = (A + B * X + C * (X ^ 2) + D * (X ^ 3) + E * (X ^ 4)) / 2.326

    End Function

    Public Function Sysbal_R134A_HFG(ByVal Tsat)

        Const B As Double = 163.7313
        Const C As Double = 460.3925
        Const D As Double = -510.952
        Const E As Double = 219.1886
        Const Tc As Double = 374
        Dim X As Double
        Dim T As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        T = 273 + ((Tsat - 32) / 1.8)

        X = (1 - (T / Tc)) ^ 0.333

        Sysbal_R134A_HFG = (B * X + C * (X ^ 2) + D * (X ^ 3) + E * (X ^ 4)) / 2.326

    End Function

    Public Function R22_VisL(ByVal Tsat)
        Const A As Double = 0.64
        Const B As Double = -0.003853
        Const C As Double = 0.00001174
        Const D As Double = 0.00000002564
        Dim T As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        T = Tsat

        R22_VisL = A + B * T + C * (T ^ 2) + D * (T ^ 3)

    End Function
    Public Function R22_VisV(ByVal Tsat)
        Const A As Double = 0.026
        Const B As Double = 0.00007255
        Const C As Double = -0.0000002391
        Const D As Double = 0.000000001839
        Dim T As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        T = Tsat

        R22_VisV = A + B * T + C * (T ^ 2) + D * (T ^ 3)

    End Function
    Public Function R22_kL(ByVal Tsat)
        Const A As Double = 0.06
        Const B As Double = -0.0001753
        Const C As Double = 0.0000005157
        Const D As Double = -0.000000002383
        Dim T As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        T = Tsat

        R22_kL = A + B * T + C * (T ^ 2) + D * (T ^ 3)

    End Function
    Public Function R22_CpL(ByVal Tsat)
        Const A As Double = 0.26
        Const B As Double = 0.0007688
        Const C As Double = -0.000007233
        Const D As Double = 0.00000005142
        Dim T As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        T = Tsat

        R22_CpL = A + B * T + C * (T ^ 2) + D * (T ^ 3)

    End Function
    Public Function R22_DenL(ByVal Tsat)
        Const A As Double = 28.171
        Const B As Double = 60.36
        Const C As Double = -26.882
        Const D As Double = 27.417
        Const E As Double = -5.435
        Const Tc As Double = 204.81
        Dim X As Double
        Dim T As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        T = Tsat

        X = (1 - (T / Tc)) ^ 0.333

        R22_DenL = A + B * X + C * (X ^ 2) + D * (X ^ 3) + E * (X ^ 4)

    End Function
    Public Function R22_DenV(ByVal Tsat)
        Const A As Double = 25.157
        Const B As Double = 0.284
        Const C As Double = -85.014
        Const D As Double = 80.243
        Const E As Double = -19.937
        Const Tc As Double = 204.81
        Dim X As Double
        Dim T As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        T = Tsat

        X = (1 - (T / Tc)) ^ 0.333

        R22_DenV = A + B * X + C * (X ^ 2) + D * (X ^ 3) + E * (X ^ 4)

    End Function
    Public Function R410A_DenL(ByVal Tsat)
        Const A As Double = 493
        Const B As Double = 930.2971
        Const C As Double = 416.4226
        Const D As Double = -86.8832
        Const Tc As Double = 344
        Dim X As Double
        Dim T As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        T = 273 + ((Tsat - 32) / 1.8)

        X = (1 - (T / Tc)) ^ 0.333

        R410A_DenL = (A + B * X + C * (X ^ 2) + D * (X ^ 3)) / 16

    End Function
    Public Function R410A_DenV(ByVal Tsat)
        Const A As Double = 0.36359063
        Const B As Double = 2.6775079
        Const C As Double = -5.9330937
        Const D As Double = 9.5070291
        Const E As Double = -9.1236153
        Const F As Double = 3.5429932
        Const Tc As Double = 344
        Const Rho As Double = 493
        Dim Y As Double
        Dim X As Double
        Dim T As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        T = 273 + ((Tsat - 32) / 1.8)

        X = (1 - (T / Tc)) ^ 0.333

        Y = A + B * X + C * (X ^ 2) + D * (X ^ 3) + E * (X ^ 4) + F * (X ^ 5)

        R410A_DenV = (1 - (Y ^ 3)) * Rho / 16

    End Function
    Public Function R410A_VisV(ByVal Tsat)
        Const A As Double = 12.1
        Const B As Double = 0.0833
        Const C As Double = 0.000147
        Const D As Double = -0.0000467
        Const E As Double = 0.00000108
        Dim T As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        T = (Tsat - 32) / 1.8

        R410A_VisV = (A + B * T + C * (T ^ 2) + D * (T ^ 3) + E * (T ^ 4)) * 2.42 * 0.001

    End Function
    Public Function R410A_VisL(ByVal Tsat)
        Const A As Double = 166
        Const B As Double = -2.25
        Const C As Double = 0.0181
        Const D As Double = -0.000092
        Dim T As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        T = (Tsat - 32) / 1.8

        R410A_VisL = (A + B * T + C * (T ^ 2) + D * (T ^ 3)) * 2.42 * 0.001

    End Function
    Public Function R410A_kL(ByVal Tsat)
        Const A As Double = 0.8148462
        Const B As Double = -0.00597897
        Const C As Double = 0.000017297
        Const D As Double = -0.0000000182606
        Dim T As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        T = 273 + ((Tsat - 32) / 1.8)

        R410A_kL = (A + B * T + C * (T ^ 2) + D * (T ^ 3)) / 1.730735

    End Function
    Public Function R410A_CpL(ByVal Tsat)
        Const A As Double = 1.603
        Const B As Double = 0.005727
        Const C As Double = 0.00009903
        Const D As Double = 0.000001855
        Dim T As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        T = (Tsat - 32) / 1.8

        R410A_CpL = (A + B * T + C * (T ^ 2) + D * (T ^ 3)) / 4.184

    End Function
    Public Function R404A_DenL(ByVal Tsat)
        Const A As Double = 380.3969
        Const B As Double = 1360.097
        Const C As Double = -420.077
        Const D As Double = 515.1893
        Const Tc As Double = 344.7
        Dim X As Double
        Dim T As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        T = 273 + ((Tsat - 32) / 1.8)

        X = (1 - (T / Tc)) ^ 0.333

        R404A_DenL = (A + B * X + C * (X ^ 2) + D * (X ^ 3)) / 16

    End Function
    Public Function R404A_DenV(ByVal Tsat)
        Const A As Double = 317.9835
        Const B As Double = 225.24813
        Const C As Double = -3304.8511
        Const D As Double = 4657.4605
        Const E As Double = -1868.115
        Const Tc As Double = 344.7
        Dim X As Double
        Dim T As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        T = 273 + ((Tsat - 32) / 1.8)

        X = (1 - (T / Tc)) ^ 0.333

        R404A_DenV = (A + B * X + C * (X ^ 2) + D * (X ^ 3) + E * (X ^ 4)) / 16

    End Function
    Public Function R404A_VisV(ByVal Tsat)
        Const A As Double = 0.010527
        Const C As Double = 0.0000000537
        Const E As Double = -0.799031
        Dim T As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        T = 273 + (Tsat - 32) / 1.8

        R404A_VisV = (A + C * (T ^ 2) + (E / T)) * 2.42

    End Function
    Public Function R404A_VisL(ByVal Tsat)
        Const A As Double = 10.6798
        Const B As Double = -0.0615909
        Const C As Double = 0.000148
        Const D As Double = -0.00000012971
        Const E As Double = -567.2323
        Dim T As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        T = 273 + (Tsat - 32) / 1.8

        R404A_VisL = (A + B * T + C * (T ^ 2) + D * (T ^ 3) + (E / T)) * 2.42

    End Function
    Public Function R404A_CpL(ByVal Tsat)
        Const A As Double = 0.174
        Const B As Double = 0.007694
        Const C As Double = -0.0001117
        Const D As Double = 0.0000005684
        Dim T As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        T = Tsat

        R404A_CpL = A + B * T + C * (T ^ 2) + D * (T ^ 3)

    End Function
    Public Function R404A_kL(ByVal Tsat)
        Const A As Double = 0.220049
        Const B As Double = -0.00065352
        Const C As Double = 0.0000011827
        Const D As Double = -0.0000000022864
        Dim T As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        T = 273 + ((Tsat - 32) / 1.8)

        R404A_kL = (A + B * T + C * (T ^ 2) + D * (T ^ 3)) / 1.730735

    End Function

    Public Function R407C_DenL(ByVal Tsat)
        Const A As Double = -650.4582
        Const B As Double = 7861.656
        Const C As Double = -15312.36
        Const D As Double = 15705.71
        Const E As Double = -5734.191
        Const Tc As Double = 359.2
        Dim X As Double
        Dim T As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        T = 273 + ((Tsat - 32) / 1.8)

        X = (1 - (T / Tc)) ^ 0.333

        R407C_DenL = (A + B * X + C * (X ^ 2) + D * (X ^ 3) + E * (X ^ 4)) / 16

    End Function
    Public Function R407C_DenV(ByVal Tsat)
        Const A As Double = 315.4088
        Const B As Double = 343.5035
        Const C As Double = -3883.4765
        Const D As Double = 5587.518
        Const E As Double = -2353.113
        Const Tc As Double = 359.2
        Dim X As Double
        Dim T As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        T = 273 + ((Tsat - 32) / 1.8)

        X = (1 - (T / Tc)) ^ 0.333

        R407C_DenV = (A + B * X + C * (X ^ 2) + D * (X ^ 3) + E * (X ^ 4)) / 16

    End Function
    Public Function R407C_VisL(ByVal Tsat)
        Const A As Double = 15.66442
        Const B As Double = -1283.053
        Const C As Double = -0.061504
        Const D As Double = 0.0000581907
        Dim T As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        T = 273 + ((Tsat - 32) / 1.8)

        R407C_VisL = 2.42 * Exp(A + (B / T) + C * T + D * (T ^ 2))

    End Function
    Public Function R407C_VisV(ByVal Tsat)
        Const A As Double = 0.375554
        Const B As Double = -0.00136789
        Const C As Double = 0.00000177845
        Const D As Double = -33.3749
        Dim T As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        T = 273 + ((Tsat - 32) / 1.8)

        R407C_VisV = (A + B * T + C * (T ^ 2) + (D / T)) * 2.42

    End Function
    Public Function R407C_CpL(ByVal Tsat)
        Const A As Double = 0.309
        Const B As Double = 0.001067
        Const C As Double = -0.00001134
        Const D As Double = 0.00000008174
        Dim T As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        T = Tsat

        R407C_CpL = A + B * T + C * (T ^ 2) + D * (T ^ 3)

    End Function
    Public Function R407C_kL(ByVal Tsat)
        Const A As Double = 0.11898
        Const C As Double = -0.000000671955
        Const D As Double = 9.039943
        Dim T As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        T = 273 + ((Tsat - 32) / 1.8)

        R407C_kL = (A + C * (T ^ 2) + (D / T)) / 1.730735

    End Function
    Public Function R134A_DenL(ByVal Tsat)
        Const A As Double = 508
        Const B As Double = 967.57693
        Const C As Double = 298.02172
        Const D As Double = 79.877831
        Const E As Double = 89.838713
        Const Tc As Double = 374.3
        Dim X As Double
        Dim T As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        T = 273 + ((Tsat - 32) / 1.8)

        X = (1 - (T / Tc)) ^ 0.333

        R134A_DenL = (A + B * X + C * (X ^ 2) + D * (X ^ 3) + E * (X ^ 4)) / 16

    End Function
    Public Function R134A_DenV(ByVal Tsat)
        Const A As Double = 388.752
        Const B As Double = 84.07428
        Const C As Double = -3500.71
        Const D As Double = 5252.284
        Const E As Double = -2202.55
        Const Tc As Double = 374.3
        Dim X As Double
        Dim T As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        T = 273 + ((Tsat - 32) / 1.8)

        X = (1 - (T / Tc)) ^ 0.333

        R134A_DenV = (A + B * X + C * (X ^ 2) + D * (X ^ 3) + E * (X ^ 4)) / 16

    End Function
    Public Function R134A_VisL(ByVal Tsat)
        Const A As Double = -9.707292
        Const B As Double = 1140.7291
        Const C As Double = 0.0282451
        Const D As Double = -0.00004672
        Dim T As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        T = 273 + ((Tsat - 32) / 1.8)

        R134A_VisL = 2.42 * Exp(A + (B / T) + C * T + D * (T ^ 2))

    End Function
    Public Function R134A_VisV(ByVal Tsat)
        Const A As Double = -0.32671694
        Const B As Double = 0.003456914
        Const C As Double = -0.000011836
        Const D As Double = 0.000000013599
        Dim T As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        T = 273 + ((Tsat - 32) / 1.8)

        R134A_VisV = (A + B * T + C * (T ^ 2) + D * (T ^ 3)) * 2.42

    End Function
    Public Function R134A_kL(ByVal Tsat)
        Const A As Double = 0.058
        Const B As Double = -0.0001456
        Const C As Double = 0.00000006882
        Const D As Double = -0.0000000002072
        Dim T As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        T = Tsat

        R134A_kL = A + B * T + C * (T ^ 2) + D * (T ^ 3)

    End Function
    Public Function R134A_CpL(ByVal Tsat)
        Const A As Double = 0.301
        Const B As Double = 0.0007941
        Const C As Double = -0.000007068
        Const D As Double = 0.00000004481
        Dim T As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        T = Tsat

        R134A_CpL = A + B * T + C * (T ^ 2) + D * (T ^ 3)

    End Function
    Public Function R507_DenL(ByVal Tsat)
        Const A As Double = 32.278
        Const B As Double = 24.19
        Const C As Double = 34.381
        Const D As Double = -23.077
        Const E As Double = 8.487
        Const Tc As Double = 159.11
        Dim X As Double
        Dim T As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        T = Tsat

        X = (1 - (T / Tc)) ^ 0.333

        R507_DenL = A + B * X + C * (X ^ 2) + D * (X ^ 3) + E * (X ^ 4)

    End Function
    Public Function R507_DenV(ByVal Tsat)
        Const A As Double = 28.729
        Const B As Double = -23.282
        Const C As Double = -34.095
        Const D As Double = 37.871
        Const E As Double = -8.11
        Const Tc As Double = 159.11
        Dim X As Double
        Dim T As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        T = Tsat

        X = (1 - (T / Tc)) ^ 0.333

        R507_DenV = A + B * X + C * (X ^ 2) + D * (X ^ 3) + E * (X ^ 4)

    End Function
    Public Function R507_VisV(ByVal Tsat)
        Const A As Double = 0.024
        Const B As Double = 0.0001028
        Const C As Double = -0.0000009602
        Const D As Double = 0.000000007394
        Dim T As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        T = Tsat

        R507_VisV = A + B * T + C * (T ^ 2) + D * (T ^ 3)

    End Function
    Public Function R507_VisL(ByVal Tsat)
        Const A As Double = 0.536
        Const B As Double = -0.003866
        Const C As Double = 0.00001448
        Const D As Double = -0.00000004144
        Dim T As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        T = Tsat

        R507_VisL = A + B * T + C * (T ^ 2) + D * (T ^ 3)

    End Function
    Public Function R507_kL(ByVal Tsat)
        Const A As Double = 0.049
        Const B As Double = -0.0001491
        Const C As Double = 0.0000002388
        Const D As Double = -0.0000000009713
        Dim T As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        T = Tsat

        R507_kL = A + B * T + C * (T ^ 2) + D * (T ^ 3)

    End Function
    Public Function R507_CpL(ByVal Tsat)
        Const A As Double = 0.285
        Const B As Double = 0.002409
        Const C As Double = -0.00003697
        Const D As Double = 0.0000002483
        Dim T As Double

        If Tsat > 155 Then
            Tsat = 155
        End If

        T = Tsat

        R507_CpL = A + B * T + C * (T ^ 2) + D * (T ^ 3)

    End Function
    Public Function H2O_DenL(ByVal Tsat)
        Const A As Double = 62.068
        Const B As Double = 0.007673
        Const C As Double = -0.0001049
        Const D As Double = 0.00000008903

        H2O_DenL = A + B * Tsat + C * (Tsat ^ 2) + D * (Tsat ^ 3)

    End Function
    Public Function H2O_DenV(ByVal Tsat)
        Const A As Double = -0.095
        Const B As Double = 0.001589
        Const C As Double = -0.000009388
        Const D As Double = 0.00000002288

        H2O_DenV = A + B * Tsat + C * (Tsat ^ 2) + D * (Tsat ^ 3)

    End Function
    Public Function H2O_VisV(ByVal Tsat)
        Const A As Double = 0.021
        Const B As Double = 0.00003034
        Const C As Double = 0.00000006236
        Const D As Double = -0.00000000008094

        H2O_VisV = A + B * Tsat + C * (Tsat ^ 2) + D * (Tsat ^ 3)

    End Function
    Public Function H2O_VisL(ByVal Tsat)
        Const A As Double = 3.089
        Const B As Double = -0.022
        Const C As Double = 0.00006412
        Const D As Double = -0.0000000691

        H2O_VisL = A + B * Tsat + C * (Tsat ^ 2) + D * (Tsat ^ 3)

    End Function
    Public Function H2O_kL(ByVal Tsat)
        Const A As Double = 0.312
        Const B As Double = 0.0007101
        Const C As Double = -0.000001828
        Const D As Double = 0.000000001234

        H2O_kL = A + B * Tsat + C * (Tsat ^ 2) + D * (Tsat ^ 3)

    End Function
    Public Function H2O_CpL(ByVal Tsat)
        Const A As Double = 0.948
        Const B As Double = 0.000684
        Const C As Double = -0.000003185
        Const D As Double = 0.00000000607

        H2O_CpL = A + B * Tsat + C * (Tsat ^ 2) + D * (Tsat ^ 3)

    End Function


End Class
