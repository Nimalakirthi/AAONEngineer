
Public Structure Comp
    Public s_Model As String     'model of compressor
    Public s_RefCode As String     'Ref code
    Public s_FreCode As String     'Fre code
    Public s_AppCode As String     'App code
    Public d_CF() As Double     'coefficients C/W/A/M
    Public n_MassFlow As Double
    Public n_BTUH As Double
    Public n_WattsBTUH As Double
    Public n_Amps As Double

    Public s_ReheatType As String
    Public n_Subcool As Double
    Public n_Superh As Double

End Structure

Public Class CompClass

    Private mComp As Comp

    Public Property CompProperties() As Comp
        Get
            Return mComp
        End Get

        Set(ByVal NewComp As Comp)
            mComp = NewComp
        End Set
    End Property

    Public Sub CompCalc(ByVal mn_ET As Double, ByVal mn_CT As Double)

        With mComp
            .n_BTUH = CompressorProp(0, mn_ET, mn_CT)
            .n_MassFlow = CompressorProp(30, mn_ET, mn_CT)
            .n_WattsBTUH = CompressorProp(10, mn_ET, mn_CT)
            .n_Amps = CompressorProp(20, mn_ET, mn_CT)
        End With
    End Sub

    Public Function CompressorProp(ByVal i_X As Integer, ByVal mn_ET As Double, ByVal mn_CT As Double) As Double

        ' i_x will determine the starting location of the coeficients in the array
        ' The Compressor capacity uses a different formula

        ' Capacity      Table: C0-C9    Array: 0  -  9
        ' Watts         Table: W0-W9    Array: 10 - 19
        ' Amps          Table: A0-A9    Array: 20 - 29
        ' Mass Flow     Table: M0-M9    Array: 30 - 39

        'Dim i_Compressor As Integer
        Dim n_ET As Double
        Dim n_CT As Double
        Dim SuctionLoss As Double
        Dim DischargeLoss As Double
        Dim DefaultCapacity As Double
        Dim Superh As Double
        Dim Subcool As Double

        SuctionLoss = 1.5

        If mComp.s_ReheatType <> "None" Then
            DischargeLoss = 4.5
        Else
            DischargeLoss = 2.5
        End If

        n_ET = mn_ET - SuctionLoss
        n_CT = mn_CT + DischargeLoss
        Subcool = mComp.n_Subcool
        Superh = mComp.n_Superh


        With mComp

            If i_X = 0 Then   'Compressor Capactiy (C0-C9)

                DefaultCapacity = .d_CF(0) + .d_CF(1) * n_ET + .d_CF(2) * n_CT + .d_CF(3) * n_ET * n_ET + .d_CF(4) * _
                n_CT * n_ET + .d_CF(5) * n_CT * n_CT + .d_CF(6) * n_ET * n_ET * n_ET + .d_CF(7) * _
                n_CT * (n_ET * n_ET) + .d_CF(8) * n_ET * (n_CT * n_CT) + .d_CF(9) * n_CT * n_CT * n_CT

                CompressorProp = DefaultCapacity * (1 - 0.00125 * (Superh - 20) + 0.0045 * (Subcool - 15))

            Else

                CompressorProp = .d_CF(i_X + 0) + .d_CF(i_X + 1) * n_ET + .d_CF(i_X + 2) * n_CT + .d_CF(i_X + 3) _
                * n_ET * n_ET + .d_CF(i_X + 4) * n_CT * n_ET + .d_CF(i_X + 5) * n_CT * n_CT + .d_CF(i_X + 6) _
                * n_ET * n_ET * n_ET + .d_CF(i_X + 7) * n_CT * (n_ET * n_ET) + .d_CF(i_X + 8) * n_ET _
                * (n_CT * n_CT) + .d_CF(i_X + 9) * n_CT * n_CT * n_CT


            End If

        End With

    End Function

End Class
