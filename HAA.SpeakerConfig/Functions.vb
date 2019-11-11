Public Class Functions
    Public Function CrossProduct(
    ByVal X1 As Single, ByVal Y1 As Single, ByVal Z1 As Single,
    ByVal X2 As Single, ByVal Y2 As Single, ByVal Z2 As Single,
    ByVal X3 As Single, ByVal Y3 As Single, ByVal Z3 As Single) As Object 'input three points to create two vectors sharing a common; point Point 2
        Dim V1X As Single, V1Y As Single, V1Z As Single, V2X As Single, V2Y As Single, V2Z As Single, CrossOutput() As Object

        ReDim CrossOutput(2)

        'Vector 1
        V1X = X1 - X2
        V1Y = Y1 - Y2
        V1Z = Z1 - Z2
        'Vector 2
        V2X = X3 - X2
        V2Y = Y3 - Y2
        V2Z = Z3 - Z2

        'Get Cross Product vector: <y1z2 - z1y2, z1x2 - x1z2, x1y2 - y1x2>
        'S1 = Y1Z2 - Z1Y2
        'S2 = Z1X2 - X1Z2
        'S3 = X1Y2 - Y1X2
        CrossOutput(0) = (V1Y * V2Z) - (V1Z * V2Y) 'Y1Z2 - Z1Y2
        CrossOutput(1) = (V1Z * V2X) - (V1X * V2Z) 'Z1X2 - X1Z2
        CrossOutput(2) = (V1X * V2Y) - (V1Y * V2X) 'X1Y2 - Y1X2

        CrossProduct = CrossOutput
    End Function
    Public Function DotProduct(
    ByVal X1 As Single, ByVal Y1 As Single, ByVal Z1 As Single,
    ByVal X2 As Single, ByVal Y2 As Single, ByVal Z2 As Single,
    ByVal X3 As Single, ByVal Y3 As Single, ByVal Z3 As Single) As Single 'input three points to create two vectors sharing a common; point Point 2
        Dim V1X As Single, V1Y As Single, V1Z As Single, V2X As Single, V2Y As Single, V2Z As Single

        'Vector 1
        V1X = X1 - X2
        V1Y = Y1 - Y2
        V1Z = Z1 - Z2
        'Vector 2
        V2X = X3 - X2
        V2Y = Y3 - Y2
        V2Z = Z3 - Z2

        ' Calculate the dot product.
        DotProduct = V1X * V2X + V1Y * V2Y + V1Z * V2Z
    End Function
    Public Function StandardDeviation(ByRef NumericArray As Object) As Single
        'Function takes an array with numeric elements as a parameter
        'and calculates the standard deviation
        Dim dblSum As Single, dblSumSqdDevs As Single, dblMean As Single
        Dim lngCount As Long, dblAnswer As Single
        Dim vElement As Object
        Dim lngStartPoint As Long, lngEndPoint As Long, lngCtr As Long

        'if NumericArray is not an array, this statement will
        'raise an error in the errorhandler

        lngCount = UBound(NumericArray)

        On Error Resume Next
        lngCount = 0

        'the check below will allow
        'for 0 or 1 based arrays.

        vElement = NumericArray(0)

        lngStartPoint = IIf(Err.Number = 0, 0, 1)
        lngEndPoint = UBound(NumericArray)

        'get sum and sample size
        For lngCtr = lngStartPoint To lngEndPoint
            vElement = NumericArray(lngCtr)
            If IsNumeric(vElement) Then
                lngCount = lngCount + 1
                dblSum = dblSum + CDbl(vElement)
            End If
        Next

        'get mean
        If lngCount > 1 Then
            dblMean = dblSum / lngCount

            'get sum of squared deviations
            For lngCtr = lngStartPoint To lngEndPoint
                vElement = NumericArray(lngCtr)

                If IsNumeric(vElement) Then
                    dblSumSqdDevs = dblSumSqdDevs +
            ((vElement - dblMean) ^ 2)
                End If
            Next

            'divide result by sample size - 1 and get square root.
            'this function calculates standard deviation of a sample.
            'If your  set of values represents the population, use sample
            'size not sample size - 1

            If lngCount > 1 Then
                lngCount = lngCount - 1 'eliminate for population values
                dblAnswer = Math.Sqrt(dblSumSqdDevs / lngCount)
            End If

        End If

        Return dblAnswer

    End Function
    Public Function Radians(Entry As Single) 'converts degrees to radians
        Radians = Entry * (Math.PI / 180)
    End Function
    Public Function Degrees(Entry As Single) ' converts radians to degrees
        Degrees = Entry * (180 / Math.PI)
    End Function
    Public Function TriangleArea(ByRef P1 As Object, ByRef P2 As Object, ByRef P3 As Object)
        Dim ResultVar As Single, Element1 As Single, Element2 As Single, Element3 As Single
        'used to determine if three point are on a line, if area =0 then on a line

        'ResultVar = (((X1 - X2) * (Y2 - Y3)) - ((X2 - X3) * (Y1 - Y2))) / 2
        Element1 = P2(0) * P1(1) - P3(0) * P1(1) - P1(0) * P2(1) + P3(0) * P2(1) + P1(0) * P3(1) - P2(0) * P3(1)
        Element2 = P2(0) * P1(2) - P3(0) * P1(2) - P1(0) * P2(2) + P3(0) * P2(2) + P1(0) * P3(2) - P2(0) * P3(2)
        Element3 = P2(1) * P1(2) - P3(1) * P1(2) - P1(1) * P2(2) + P3(1) * P2(2) + P1(1) * P3(2) - P2(1) * P3(2)

        Element1 = Element1 * Element1
        Element2 = Element2 * Element2
        Element3 = Element3 * Element3

        ResultVar = Math.Sqrt(Element1 + Element2 + Element3) / 2

        TriangleArea = ResultVar
    End Function
    Public Function PythagDist(X1 As Single, X2 As Single, Y1 As Single, Y2 As Single) As Single
        Dim Hypoyvar As Single, XSqr As Single, YSqr As Single

        XSqr = Math.Abs(X1 - X2) * Math.Abs(X1 - X2)
        YSqr = Math.Abs(Y1 - Y2) * Math.Abs(Y1 - Y2)
        Hypoyvar = Math.Sqrt(XSqr + YSqr)
        PythagDist = Hypoyvar
    End Function
    Public Function Intersection(ByRef P1() As Single, ByRef P2() As Single, ByRef Corner1() As Single, ByRef Corner2() _
As Single, ByRef Corner3() As Single, ByRef Corner4() As Single) As Object
        Dim EPS As Single, P As Object
        Dim Intersect As Boolean, Cnt As Integer
        Dim PA() As Single, pb() As Single, pc() As Single
        Dim D As Single, Delta As Single
        Dim A1 As Single, A2 As Single, A3 As Single
        Dim Total As Single, denom As Single, mu As Single
        Dim n(2) As Single, pa1(2) As Single, pa2(2) As Single, pa3(2) As Single, TempArr As Object

        EPS = 0.1

        ReDim P(2)

        Intersect = True
        PA = Corner1
        pb = Corner2
        pc = Corner3

        For Cnt = 1 To 2
            If Cnt = 2 And Intersect = True Then
                Intersection = P
                Exit Function
            ElseIf Cnt = 2 Then
                PA = Corner1
                pb = Corner4
                pc = Corner3
            End If

            '/* Calculate the parameters for the plane */
            n(0) = (pb(1) - PA(1)) * (pc(2) - PA(2)) - (pb(2) - PA(2)) * (pc(1) - PA(1))
            n(1) = (pb(2) - PA(2)) * (pc(0) - PA(0)) - (pb(0) - PA(0)) * (pc(2) - PA(2))
            n(2) = (pb(0) - PA(0)) * (pc(1) - PA(1)) - (pb(1) - PA(1)) * (pc(0) - PA(0))
            TempArr = Normalize(n(0), n(1), n(2))
            n(0) = TempArr(0)
            n(1) = TempArr(1)
            n(2) = TempArr(2)
            D = -n(0) * PA(0) - n(1) * PA(1) - n(2) * PA(2)
            '/* Calculate the position on the line that intersects the plane */
            denom = n(0) * (P2(0) - P1(0)) + n(1) * (P2(1) - P1(1)) + n(2) * (P2(2) - P1(2))
            If denom = 0 Then denom = 0.00001
            If (Math.Abs(denom) < EPS) Then
                Intersect = False
            Else
                Intersect = True
            End If
            mu = -(D + n(0) * P1(0) + n(1) * P1(1) + n(2) * P1(2)) / denom
            P(0) = P1(0) + mu * (P2(0) - P1(0))
            P(1) = P1(1) + mu * (P2(1) - P1(1))
            P(2) = P1(2) + mu * (P2(2) - P1(2))
            If mu < 0 Or mu > 1 Then
                Intersect = False
            Else
                Intersect = True
            End If
            If Intersect Then
                '/* Determine whether or not the intersection point is bounded by pa,pb,pc */
                pa1(0) = PA(0) - P(0)
                pa1(1) = PA(1) - P(1)
                pa1(2) = PA(2) - P(2)
                TempArr = Normalize(pa1(0), pa1(1), pa1(2))
                pa1(0) = TempArr(0)
                pa1(1) = TempArr(1)
                pa1(2) = TempArr(2)
                pa2(0) = pb(0) - P(0)
                pa2(1) = pb(1) - P(1)
                pa2(2) = pb(2) - P(2)
                TempArr = Normalize(pa2(0), pa2(1), pa2(2))
                pa2(0) = TempArr(0)
                pa2(1) = TempArr(1)
                pa2(2) = TempArr(2)
                pa3(0) = pc(0) - P(0)
                pa3(1) = pc(1) - P(1)
                pa3(2) = pc(2) - P(2)
                TempArr = Normalize(pa3(0), pa3(1), pa3(2))
                pa3(0) = TempArr(0)
                pa3(1) = TempArr(1)
                pa3(2) = TempArr(2)
                A1 = pa1(0) * pa2(0) + pa1(1) * pa2(1) + pa1(2) * pa2(2)
                A2 = pa2(0) * pa3(0) + pa2(1) * pa3(1) + pa2(2) * pa3(2)
                A3 = pa3(0) * pa1(0) + pa3(1) * pa1(1) + pa3(2) * pa1(2)
                If A1 = 1 Then A1 = 0.999999
                If A2 > 0.999999 Then
                    A2 = 0.999999
                End If
                If A3 = 1 Then A3 = 0.999999
                Total = (Math.Acos(A1) + Math.Acos(A2) + Math.Acos(A3)) * (180 / Math.PI)
                Delta = Math.Abs(Total - 360)
                If Delta > EPS Then
                    Intersect = False
                Else
                    Intersect = True
                End If
            End If
            ''''
            ''If Not IntersectPP Then ' added so drawLR boxes will work
            'If Intersect = True Then
            '    If P1(0) < P(0) And P(0) < P2(0) Then
            '         Intersect = True
            '    ElseIf P2(0) < P(0) And P(0) < P1(0) Then
            '         Intersect = True
            '    Else
            '         Intersect = False
            '    End If
            'End If
            '
            ' If Intersect = True Then
            '     If P1(2) < P(2) And P(2) < P2(2) Then
            '          Intersect = True
            '     ElseIf P2(2) < P(2) And P(2) < P1(2) Then
            '          Intersect = True
            '     Else
            '          Intersect = False
            '     End If
            ' End If
            'End If
            ''''
            If Intersect = False Then
                P(0) = 0
                P(1) = 0
                P(2) = 0
            Else
                P(0) = P(0)
            End If

        Next Cnt
        Intersection = P
    End Function
    Private Function Normalize(ByVal a As Single, ByVal b As Single, c As Single) As Object
        Dim Mag As Single, OutAns(2) As Object, AO As Single, BO As Single, CO As Single

        Mag = Math.Sqrt(a * a + b * b + c * c)
        If Mag <> 0 Then
            AO = a / Mag
            BO = b / Mag
            CO = c / Mag
        End If
        OutAns(0) = AO
        OutAns(1) = BO
        OutAns(2) = CO
        Return OutAns
    End Function
End Class
