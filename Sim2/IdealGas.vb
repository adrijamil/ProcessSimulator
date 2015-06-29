Imports System
Imports Math = System.Math


Public Class IdealGas
    Private Const R = 8.314
    Private Const Tref = 273.15
    ' ------------------------------------------------------------------------------------------------------
    ' ClassName   : IdealGas
    ' Version     : 1.0
    ' Author      : AJS
    ' Purpose     : This is a property package type object. It implements flash routines on streams based on Ideal Gas law and other simplified equations
    ' Remarks     : This class serves mainly as a basis to develop more advanced packages
    ' Improvements: 1) To add other flash routines (T-S, Vf-T etc)
    '               2) Standardise the way NComps is obtained (several ways to do this).
    '               3) Can be used as a basis to implement Aqueous phase.
    ' ------------------------------------------------------------------------------------------------------
    '
    ' ------------------------------------------------------------------------------------------------------



    Private myParent As FlowSheet 'Property Packages are attached to flowsheet. That way only property package object needs to be instantiated. To figure out how to associate PropPacks to streams
    Private TargStream As Stream ' This will change when stack is being solved.
    Private myName As String 'does not really need a name
    Private myComponents As New System.Collections.Generic.SortedList(Of String, Component)
    Private myNComps As Integer
    Private myKs() As Double 'K values (yi/xi)
    Private myCompDB As String

    WriteOnly Property CompDB As String 'Component database. It is a .dat file
        Set(daName As String)
            myCompDB = daName
        End Set
    End Property
    ReadOnly Property nComps As Integer
        Get
            nComps = myNComps
        End Get
    End Property

    WriteOnly Property Parent As FlowSheet
        Set(daParent As FlowSheet)
            myParent = daParent
        End Set
    End Property
    Property Name As String
        'put some logic to check for conflicts
        Get
            Name = myName
        End Get

        Set(daName As String)
            myName = daName
        End Set
    End Property

    ReadOnly Property Component(ByVal i As Integer) As Component

        Get
            ' to access Component objects
            Component = myComponents.Values(i)
        End Get

    End Property

    Public Function CreateComponent(daName As String) As Component
        Dim tempComp As New Component
        tempComp.MakeComponent(daName, myCompDB)
        CreateComponent = tempComp
    End Function

    Sub AddComponent(ByRef daComp As Component)
        'check for conflicts
        myNComps = myNComps + 1
        myComponents.Add(daComp.Name, daComp)
    End Sub


    Public Function BubTemp() As Double

        'Bubble point temperature calculation at specified P
        'To make it more robust by handling exceptions


        Dim i As Integer
        Dim tol As Double
        Dim Told As Double
        Dim t As Double
        Dim Fold As Double
        Dim F As Double
        Dim tempDb As Double
        Dim j As Integer

        'initial guesses for temperature 'a "smarter" guess should be implemented

        Told = 273
        t = 323
        tol = 0.001

startagain:  'this line is not implemented but may be required to start again

        'Calculate K values at given T and Told
        'from K values, calculate F. Where F=0 is function to be solved
        'F=sum[yi/Ki] -1

        EstimateKs(TargStream.Pressure.Val, Told)
        Fold = 0
        For i = 0 To myNComps - 1
            Fold = Fold + TargStream.Composition.Val(i) * myKs(i)
        Next
        Fold = Fold - 1


        EstimateKs(TargStream.Pressure.Val, t)
        F = 0
        For i = 0 To myNComps - 1
            F = F + TargStream.Composition.Val(i) * myKs(i) '
        Next
        F = F - 1


        'do a secant method

        Do
            'Secant new guess
            tempDb = t
            tempDb = t - F * (t - Told) / (F - Fold)

            'limit movement of T when you are far away from solution
            ' this avoids large fluctuations in T guess
            If Math.Abs(t - tempDb) > 50 Then
                If tempDb > t Then
                    tempDb = t + 50
                Else
                    tempDb = t - 50
                End If
            End If

            Told = t
            Fold = F
            t = tempDb

            'evaluate F again
            EstimateKs(TargStream.Pressure.Val, t)
            F = 0
            For i = 0 To myNComps - 1
                F = F + TargStream.Composition.Val(i) * myKs(i) '
            Next
            F = F - 1

            j = j + 1
        Loop Until Math.Abs(F) < tol

        BubTemp = t

    End Function
    Public Function DewTemp() As Double
        'Dew point temperature calculation at specified P
        'To make it more robust by handling exceptions

        'dim tempSum As Double
        Dim i As Integer
        Dim tol As Double
        Dim Told As Double
        Dim t As Double
        Dim Fold As Double
        Dim F As Double
        'Dim T3 As Double
        Dim tempDb As Double
        Dim j As Integer

        'Smarter initial guess needed
        tol = 0.001
        Told = 273
        t = 323

startagain:

        'Calculate K values at given T and Told
        'from K values, calculate F. Where F=0 is function to be solved
        'F=sum[xi*Ki] -1

        EstimateKs(TargStream.Pressure.Val, Told)
        Fold = 0
        For i = 0 To myNComps - 1
            ' Fold = Fold + TargStream.Composition.Val(i) * myKs(i) ' this is for bubblepoint
            Fold = Fold + TargStream.Composition.Val(i) / myKs(i) '
        Next
        Fold = Fold - 1

        EstimateKs(TargStream.Pressure.Val, t)
        F = 0
        For i = 0 To myNComps - 1
            F = F + TargStream.Composition.Val(i) / myKs(i) '
        Next
        F = F - 1


        'do a secant method

        Do
            'Secant new guess
            tempDb = t
            tempDb = t - F * (t - Told) / (F - Fold)

            'limit movement of T when you are far away from solution
            If Math.Abs(t - tempDb) > 50 Then
                If tempDb > t Then
                    tempDb = t + 50
                Else
                    tempDb = t - 50
                End If
            End If

            'not sure if this solves anything
            'If tempDb < 0 Then
            '    Told = T
            '    T = T + 20
            '    GoTo startagain
            'End If


            Told = t
            Fold = F
            t = tempDb

            'Evaluate F again
            EstimateKs(TargStream.Pressure.Val, t)
            F = 0
            For i = 0 To myNComps - 1
                F = F + TargStream.Composition.Val(i) / myKs(i) '
            Next
            F = F - 1

            j = j + 1
            'Debug.Print Abs(F)
        Loop Until Math.Abs(F) < tol
        'Debug.Print j

        DewTemp = t

    End Function

    Public Sub PH_Flash(daStream As Stream)
        'This sub can be basis to use the Boston-Britt Inside-Out for rigourous PH flash (if fugacities are calculated then must use inside out)
        'solve ideal "easy" equation sets
        'use solution rigourous "hard" system with fugacities
        'repeat until converge

        'see here:
        'http://www.bvucoepune.edu.in/pdf%27s/Research%20and%20Publication/Research%20Publications_2008-09/National%20Conference_2008-09/Mathematical%20Tools%20Mrs%20V%20A%20Shinde.pdf


        'stupidest way:
        'guess T= 300 K, 350 K
        'calc H
        'start secant method to find T

        'smarter abit: 'this is implemented here
        'do bub and dew temp
        'calc enth for dew and bub
        'interpolate to get guess for T
        'iterate

        'to implement max iterations. Make more robust


        Dim Tdew As Double
        Dim Tbub As Double
        Dim HDew As Double
        Dim HBub As Double
        Dim Tguess As Double
        Dim Told As Double
        Dim TargEnth As Double
        Dim tol As Double
        Dim tempSatLiq As New Stream
        Dim tempSatVap As New Stream
        Dim tempPhase As Phase
        Dim j As Integer
        Dim F As Double
        Dim Fold As Double
        Dim tempDb As Double


        TargStream = daStream

        tol = 10 'large tolerance is required here. Because we are using J/kg, numbers tend to be large and cannot converge to smaller tol


        TargEnth = daStream.MassEnthalpy.Val

        Tdew = DewTemp()
        Tbub = BubTemp()

        'Create saturated liquid phase to calculate enthalpy
        tempSatLiq.Composition.Val = TargStream.Phases(1).Composition.Val
        tempSatLiq.SetParent(TargStream.myParent)

        tempPhase = TargStream.Phases(1).Clone
        tempPhase.Name = "Liquid"
        tempSatLiq.Phases.Add(tempPhase.Name, tempPhase)

        tempSatLiq.Pressure.Val = TargStream.Pressure.Val
        tempSatLiq.Temperature.Val = Tbub
        CalcEnthalpy(tempSatLiq) 'after this the overall stream enthalpy will not be calculated ' make a stream copy function
        HBub = tempSatLiq.Phases(2).MassEnthalpy.Val

        'Create saturated vapour phase to calculate enthalpy
        tempSatVap.Composition.Val = TargStream.Phases(1).Composition.Val
        tempSatVap.SetParent(TargStream.myParent)

        tempPhase = TargStream.Phases(1).Clone
        tempPhase.Name = "Vapour"
        tempSatVap.Phases.Add(tempPhase.Name, tempPhase)
        tempSatVap.Pressure.Val = TargStream.Pressure.Val
        tempSatVap.Temperature.Val = Tdew
        CalcEnthalpy(tempSatVap)
        HDew = tempSatVap.Phases(2).MassEnthalpy.Val


        'now intepolate between the two temperatures and enthalpies
        'for close boiling mixtures. dT/dH will be close to 0. T guess may be thrown far away
        If TargEnth > HBub And TargEnth < HDew Then
            Tguess = Tbub + (Tdew - Tbub) / (HDew - HBub) * (TargStream.MassEnthalpy.Val - HBub)
        ElseIf TargEnth < HBub Then ' subcooled
            Tguess = Tbub - 30
        ElseIf TargEnth > HDew Then 'sub
            Tguess = Tdew + 30
        End If

        'start with intital guess and secant to find T
        Told = Tguess
        Tguess = Tguess + 10 '2nd point to start secant

        ' Evaluate F for Told and Tguess
        'F = CalcEnthalpy - Target Enthalpy
        'Use PT_Flash to get enthalpies
        daStream.Temperature.Val = Told
        PT_Flash(daStream)
        Fold = daStream.MassEnthalpy.Val - TargEnth

        daStream.Temperature.Val = Tguess
        PT_Flash(daStream)
        F = daStream.MassEnthalpy.Val - TargEnth

        Do
            'secant next guess
            tempDb = Tguess
            tempDb = Tguess - F * (Tguess - Told) / (F - Fold)

            Told = Tguess
            Fold = F
            Tguess = tempDb

            daStream.Temperature.Val = Tguess
            PT_Flash(daStream)
            F = daStream.MassEnthalpy.Val - TargEnth

            j = j + 1
            'Debug.Print Abs(F) & " " & Tguess
            'Debug.Print Tguess
        Loop Until Math.Abs(F) < tol



    End Sub
    Public Sub PT_Flash(daStream As Stream)
        'Dim daMw As Double
        Dim i As Integer
        'Dim P As Double
        'Dim t As Double
        'Dim daDens As Double

        Dim Kmax As Double
        Dim Kmin As Double
        Dim Vfrac As Double
        Dim tempLiq As Phase
        Dim tempVap As Phase

        Dim tempComps() As Double

        ReDim tempComps(0 To myNComps)

        ' calculate K values from Temperature
        TargStream = daStream
        EstimateKs()

        Kmax = 1
        Kmin = 1

        For i = 0 To myNComps - 1
            If myKs(i) > Kmax Then Kmax = myKs(i)
            If myKs(i) < Kmin Then Kmin = myKs(i)
        Next

        If Kmax > 1 And Kmin < 1 Then
            Vfrac = SolveRR(1 / (1 - Kmax), 1 / (1 - Kmin)) ' 1 / (1 - Kmax), 1 / (1 - Kmin) are supposed to be upper and lower bound for Vfrac but they are always < 0 and > 1  respectively [CHECK]

        Else
            If Kmax > 1 Then 'if all Pvaps>P then must be gas otherwise it is liquid [LOGIC CHECK]
                Vfrac = 1
            Else
                Vfrac = 0
            End If
        End If

        With TargStream

            If Vfrac < 1 Then
                '    If .Phases("Liquid").MoleFrac.Val = -1 Then 'Add liquid phase if there is none
                '        Set tempLiq = New Phase
                '        tempLiq.Name = "Liquid"
                '        tempLiq.SetParent TargStream
                '        .Phases.Add tempLiq, tempLiq.Name
                '    Else
                tempLiq = .Phases("Liquid")
                '    End If

                If Vfrac = 0 Then 'if liquid only then just take compositions from overall phase
                    For i = 0 To myNComps - 1
                        tempComps(i) = TargStream.Composition.Val(i)
                    Next
                Else 'otherwise do component balance
                    For i = 0 To myNComps - 1
                        tempComps(i) = TargStream.Composition.Val(i) / (1 + Vfrac * (myKs(i) - 1))
                    Next
                    tempLiq.Composition.Val = tempComps
                End If
                tempLiq.MoleFrac.Val = 1 - Vfrac
            Else
                .Phases("Liquid").MoleFrac.Val = -1
            End If

            If Vfrac > 0 Then
                '    If .PhasePresent("Vapour") = False Then 'Add Vapour phase if there is none
                '        Set tempVap = New Phase
                '        tempVap.Name = "Vapour"
                '        tempVap.SetParent TargStream
                '        .Phases.Add tempVap, tempVap.Name
                'Else
                tempVap = .Phases("Vapour")
                '    End If
                tempVap.MoleFrac.Val = Vfrac

                tempComps = Nothing
                ReDim tempComps(0 To myNComps)
                If Vfrac = 1 Then
                    For i = 0 To myNComps - 1
                        tempComps(i) = TargStream.Composition.Val(i)
                    Next
                Else
                    tempLiq = .Phases("Liquid")
                    For i = 0 To myNComps - 1

                        tempComps(i) = myKs(i) * tempLiq.Composition.Val(i)
                    Next
                End If
                tempVap.Composition.Val = tempComps
            Else
                .Phases("Vapour").MoleFrac.Val = -1
            End If

            'calculate Enthalpy and density (note that these 2 subs will be referring to TargStream)

            CalcEnthalpy()
            CalDensity()

        End With


    End Sub

    Sub CalcEnthalpy(Optional altStream As Stream = Nothing)


        Dim tempEnth As Double
        Dim tempCp As Double
        Dim j As Integer
        Dim tempTotEnth
        Dim t As Double
        Dim i As Integer
        On Error GoTo errorhandler

        If altStream Is Nothing Then altStream = TargStream ' can use this sub on another stream (see PH_Flash)

        With altStream
            t = .Temperature.Val
            tempTotEnth = 0

            Dim tempPhase As Phase
            For i = 2 To .Phases.Count ' calculate enthalpy of each phase
                tempPhase = .Phases(i)
                tempCp = 0
                tempEnth = 0
                tempPhase.CalcMw()

                'Calculate Cp using simple linear relationship. No temperature dependancy
                For j = 1 To myParent.PropPack.nComps
                    tempCp = tempCp + myComponents.Values(j).CpMass * myComponents.Values(j).MolWt * tempPhase.Composition.Val(j)
                Next

                tempCp = tempCp / tempPhase.Mw.Val
                tempPhase.MassCp.Val = tempCp
                tempEnth = tempPhase.MassCp.Val * (t - Tref)

                If tempPhase.Name = "Vapour" Then
                    For j = 1 To myParent.PropPack.nComps
                        If myComponents.Values(j).RefState = "L" Then
                            tempEnth = tempEnth + myComponents.Values(j).Hvap
                        End If
                    Next
                End If

                If tempPhase.Name = "Liquid" Then
                    For j = 1 To myParent.PropPack.nComps
                        If myComponents.Values(j).RefState = "G" Then
                            tempEnth = tempEnth - myComponents.Values(j).Hvap
                        End If
                    Next
                End If

                tempPhase.MassEnthalpy.Val = tempEnth

                If Not tempPhase.MoleFrac.Val = -1 Then
                    tempTotEnth = tempTotEnth + tempEnth * tempPhase.Mw.Val * tempPhase.MoleFrac.Val 'sum in J/mol
                End If

            Next

            tempPhase = .Phases(1)

            tempPhase.CalcMw() 'probably a better place to put this line ' but not at every call for Mw because it will be expensive


            tempPhase.MassEnthalpy.Val = tempTotEnth / tempPhase.Mw.Val 'times overall Mw to get J/kg

        End With
errorhandler:
        If Err.Number <> 0 Then
            Debug.Print(Err.Number & " " & Err.Description)
            Err.Clear()
            Resume Next
        End If

    End Sub
    Sub CalDensity()
        Dim tempDens As Double
        Dim i As Integer
        Dim tempEnth As Double
        Dim tempTotMassDens As Double
        Dim tempCp As Double
        On Error GoTo errorhandler
        tempCp = 0
        Dim tempPhase As Phase
        Dim j As Integer

        With TargStream
            For j = 2 To TargStream.Phases.Count

                tempPhase = TargStream.Phases(j)

                If tempPhase.Name = "Vapour" Then 'calc vapour density based on ideal gas
                    tempPhase.MassDensity.Val = .Pressure.Val * tempPhase.Mw.Val / R / .Temperature.Val / 1000
                ElseIf tempPhase.Name = "Liquid" Then 'calc liq density based on StdIdealLiq density
                    For i = 1 To myParent.PropPack.nComps
                        tempDens = tempDens + myComponents.Values(i).StdLiqDens / myComponents.Values(i).MolWt * tempPhase.Composition.Val(i) 'sum in kmol/m3
                    Next
                    tempPhase.MassDensity.Val = tempDens * tempPhase.Mw.Val 'back to kg/m3
                End If

                'prepare sum for total mass
                If Not tempPhase.MoleFrac.Val = -1 Then tempTotMassDens = tempTotMassDens + tempPhase.MassDensity.Val / tempPhase.Mw.Val * tempPhase.MoleFrac.Val 'sum in kmol/m3
            Next
            'calc overall density
            .Phases(1).MassDensity.Val = tempTotMassDens * .Mw.Val 'back to kg/m3

        End With
errorhandler:
        If Err.Number <> 0 Then
            Debug.Print(Err.Number & " " & Err.Description)
            Err.Clear()
            Resume Next
        End If
    End Sub

    Sub CalcTemp()
        'not implemented anymore

        Dim t As Double
        With TargStream
            If .MassCp.IsKnown = False Then
                CalcCp()
            End If

            If .MassEnthalpy.IsKnown Then
                t = .MassEnthalpy.Val / .MassCp.Val + Tref
                .Temperature.Val = t
            Else
                Exit Sub
            End If
        End With

    End Sub
    Sub CalcCp()
        'calculates mixture Cp assuming linear relationship

        Dim tempCp As Double
        Dim i As Integer
        'Dim tempEnth As Double
        tempCp = 0
        With TargStream


            For i = 1 To myParent.PropPack.nComps

                tempCp = tempCp + myComponents.Values(i).CpMass * myComponents.Values(i).MolWt * .Composition.Val(i)
            Next

            tempCp = tempCp / .Mw.Val

            .MassCp.Val = tempCp
        End With

    End Sub

    Private Function SolveRR(Vfmin As Double, Vfmax As Double) As Double
        'solving Rachford-Rice equation using bisection
        'basis for alot of other routines

        ' Dim daSum As Double
        Dim j As Integer
        Dim tol As Double
        Dim maxiter As Integer
        Dim i As Integer
        Dim Fx As Double
        Dim Fmax As Double
        Dim Fmin As Double
        Dim Vf As Double

        maxiter = 1000
        tol = 0.0001
        Fmax = 0
        Fmin = 0
        Fx = 0


        'make sure Vfmax and Vfmin are in range
        If Vfmin < 0 Then Vfmin = 0
        If Vfmax > 1 Then Vfmax = 1

        'calc Fmax and Fmin
        For i = 0 To myNComps - 1
            Fmax = Fmax + TargStream.Composition.Val(i) * (myKs(i) - 1) / (1 + Vfmax * (myKs(i) - 1))
        Next

        For i = 0 To myNComps - 1
            Fmin = Fmin + TargStream.Composition.Val(i) * (myKs(i) - 1) / (1 + Vfmin * (myKs(i) - 1))
        Next


        'bisection algorithm
        j = 0
        Do
            Vf = (Vfmax + Vfmin) / 2

            Fx = 0
            For i = 0 To myNComps - 1
                Fx = Fx + TargStream.Composition.Val(i) * (myKs(i) - 1) / (1 + Vf * (myKs(i) - 1))
            Next

            If Fx * Fmax > 0 Then 'same sign
                Vfmax = Vf
                Fmax = Fx
            Else
                Vfmin = Vf
                Fmin = Fx
            End If

            If j = maxiter Then Exit Do

            j = j + 1

        Loop Until Math.Abs(Fx) < tol

        SolveRR = Vf

    End Function
    Sub EstimateKs(Optional P As Double = -32767, Optional t As Double = -32767)


        ReDim myKs(0 To myNComps)
        If P = -32767 Then P = TargStream.Pressure.Val
        If t = -32767 Then t = TargStream.Temperature.Val

        'Wilson’s empirical correlation.
        'https://www.e-education.psu.edu/png520/m13_p4.html
        Dim i As Integer
        For i = 0 To myNComps - 1
            myKs(i) = 1 / (P / myComponents.Values(i).Pc) * Math.Exp(5.37 * (1 + myComponents.Values(i).AcFact) * (1 - 1 / (t / myComponents.Values(i).Tc)))
        Next


    End Sub

    Public Sub New()
        myCompDB = "C:\Users\User\Desktop\TheSimulator\ComponentDatabase.xml"
    End Sub



End Class
