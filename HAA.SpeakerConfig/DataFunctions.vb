Option Explicit On

Public Class DataFunctions
    ReadOnly Functions = New Functions()
    Public SurroundHeight As Single
    Public BaseHeight As Object
    Public SpeakerToe As Object
    Public LeftHeight As Single
    Public ForwardLength As Single
    Public BackwardLength As Single

    Public Function DetermineWallType(ByRef Coordin As Object, ByRef SpeakerLabel As String) As Object
        Dim WallNormal As Object, ProjectNormal As Object, CosVar As Object, XAng As Double, YAng As Double
        Dim MagPN As Double, MagWN As Double, DotVar As Double, ZAng As Double, TempWall As Object

        ReDim ProjectNormal(2)
        ReDim TempWall(3)
        If SpeakerLabel = "Rw" Then
            XAng = XAng
        End If

        'deterime XAng
        ProjectNormal(0) = 1 'x
        ProjectNormal(1) = 0 'y
        ProjectNormal(2) = 0 'z
        'determine normal to wall plane
        WallNormal = Functions.functions.CrossProduct(Coordin(0), Coordin(1), Coordin(2), Coordin(3), Coordin(4), Coordin(5), Coordin(6), Coordin(7), Coordin(8))
        'determine the cosine of the angle between unit vectors
        DotVar = Functions.functions.DotProduct(ProjectNormal(0), ProjectNormal(1), ProjectNormal(2), 0, 0, 0, WallNormal(0), WallNormal(1), WallNormal(2))
        MagPN = (ProjectNormal(0) * ProjectNormal(0)) + (ProjectNormal(1) * ProjectNormal(1)) + (ProjectNormal(2) * ProjectNormal(2))
        MagWN = (WallNormal(0) * WallNormal(0)) + (WallNormal(1) * WallNormal(1)) + (WallNormal(2) * WallNormal(2))
        CosVar = DotVar / (Math.Sqrt(MagPN) * Math.Sqrt(MagWN))
        'determine angle
        If CosVar > 1 Then CosVar = 1
        If CosVar < -1 Then CosVar = -1
        XAng = Functions.functions.Degrees(Math.Acos(CosVar))

        'deterime YAng
        ProjectNormal(0) = 0 'x
        ProjectNormal(1) = 1 'y
        ProjectNormal(2) = 0 'z
        'determine the cosine of the angle between unit vectors
        DotVar = Functions.functions.DotProduct(ProjectNormal(0), ProjectNormal(1), ProjectNormal(2), 0, 0, 0, WallNormal(0), WallNormal(1), WallNormal(2))
        MagPN = (ProjectNormal(0) * ProjectNormal(0)) + (ProjectNormal(1) * ProjectNormal(1)) + (ProjectNormal(2) * ProjectNormal(2))
        CosVar = DotVar / (Math.Sqrt(MagPN) * Math.Sqrt(MagWN))
        'determine angle
        If CosVar > 1 Then CosVar = 1
        If CosVar < -1 Then CosVar = -1
        YAng = Functions.functions.Degrees(Math.Acos(CosVar))

        'deterime ZAng
        ProjectNormal(0) = 0 'x
        ProjectNormal(1) = 0 'y
        ProjectNormal(2) = 1 'z
        'determine the cosine of the angle between unit vectors
        DotVar = Functions.functions.DotProduct(ProjectNormal(0), ProjectNormal(1), ProjectNormal(2), 0, 0, 0, WallNormal(0), WallNormal(1), WallNormal(2))
        MagPN = (ProjectNormal(0) * ProjectNormal(0)) + (ProjectNormal(1) * ProjectNormal(1)) + (ProjectNormal(2) * ProjectNormal(2))
        CosVar = DotVar / (Math.Sqrt(MagPN) * Math.Sqrt(MagWN))
        'determine angle
        If CosVar > 1 Then CosVar = 1
        If CosVar < -1 Then CosVar = -1
        ZAng = Functions.functions.Degrees(Math.Acos(CosVar))

        SpeakerLabel = SpeakerLabel 'Front:Xang=46 & Yang=90 Zang=43,Left:Xang=90 & Yang=27 & Zang=63,Right: Xang=90 & Yang=158 & Zang=68,Rear: Xang=180 & Yang=90 & Zang=90
        'If ZAng < 70 Then ZAng = ZAng + 180
        If Mid(SpeakerLabel, 2, 1) = "t" Then
            'If ZAng >= 135 Then
            TempWall(0) = "Ceiling"
            TempWall(1) = 0 'toe-in
            TempWall(2) = 90 - XAng + 90 'toe-up
            TempWall(3) = 90 + YAng  'roll
            'Else
            '    TempWall(0) = "Ceiling"
            '    TempWall(1) = 0 'toe-in
            '    TempWall(2) = 90 - XAng + 90 'toe-up
            '    TempWall(3) = 360 - ZAng '90 - YAng  'roll
            'End If
            '    ElseIf YAng <= 135 And YAng >= 45 Then 'Front or Rear
            '        If XAng < 90 Then 'Front:Xang=46 & Yang=90 Zang=43
            '            TempWall(0) = "Front"
            '            TempWall(1) = 90 - YAng 'toe-in
            '            TempWall(2) = 90 - ZAng 'toe-up
            '           TempWall(3) = 0 'roll
            '       ElseIf XAng >= 135 And XAng <= 225 Then 'Rear: Xang=180 & Yang=90 & Zang=90
            '            TempWall(0) = "Rear"
            '            TempWall(1) = 90 + YAng  'toe-in
            '            TempWall(2) = ZAng - 90 'toe-up
            '            TempWall(3) = 0 'roll
            '        End If
            '    ElseIf XAng <= 135 And XAng >= 45 Then 'Left or Right
            '        If YAng < 90 Then 'Left:  XAng = 90 & YAng = 27 & ZAng = 63
            '            TempWall(0) = "Left"
            '            TempWall(1) = 0 'XAng 'toe-in
            '            TempWall(2) = XAng '90 'toe-up
            '            TempWall(3) = 360 - ZAng '90 - ZAng 'roll
            '        ElseIf YAng >= 135 And YAng <= 225 Then 'Right: Xang=90 & Yang=158 & Zang=68
            '            TempWall(0) = "Right"
            '            TempWall(1) = 360 - XAng 'toe-in
            '            TempWall(2) = 0 'toe-up
            '            TempWall(3) = ZAng - 90 'roll
            '        End If
            '    End If
        Else
            If ZAng >= 135 Then
                TempWall(0) = "Ceiling"
                TempWall(1) = 0 'toe-in
                TempWall(2) = 90 - XAng + 90 'toe-up
                TempWall(3) = 90 + YAng  'roll
            ElseIf YAng <= 135 And YAng >= 45 Then 'Front or Rear
                If XAng < 90 Then 'Front:Xang=46 & Yang=90 Zang=43
                    TempWall(0) = "Rear"
                    TempWall(1) = 90 + YAng '90 - YAng  'toe-in
                    TempWall(2) = ZAng - 90 'toe-up
                    TempWall(3) = 180 'roll
                ElseIf XAng >= 135 And XAng <= 225 Then 'Rear: Xang=180 & Yang=90 & Zang=90
                    TempWall(0) = "Front"
                    TempWall(1) = 90 - YAng '90 + YAng 'toe-in
                    TempWall(2) = 90 - ZAng 'toe-up
                    TempWall(3) = 180 'roll
                End If
            ElseIf XAng <= 135 And XAng >= 45 Then 'Left or Right
                If YAng < 90 Then '
                    TempWall(0) = "Left"
                    TempWall(1) = 90 + (90 - XAng) 'toe-in
                    TempWall(2) = 0 'toe-up
                    TempWall(3) = (180 - ZAng) + 90 'roll
                ElseIf YAng >= 135 And YAng <= 225 Then '
                    TempWall(0) = "Right"
                    TempWall(1) = 360 - (180 - XAng) 'toe-in
                    TempWall(2) = 0 'toe-up
                    TempWall(3) = ZAng + 90 'roll
                End If
            End If
        End If
        Return TempWall
    End Function
    Public Function DetermineWallTypebak(ByRef Coordin As Object, ByRef SpeakerLabel As String) As Object
        Dim WallNormal As Object, ProjectNormal As Object, CosVar As Object, XAng As Double, YAng As Double
        Dim MagPN As Double, MagWN As Double, DotVar As Double, ZAng As Double, TempWall As Object

        ReDim ProjectNormal(2)
        ReDim TempWall(3)
        If SpeakerLabel = "Rw" Then
            XAng = XAng
        End If

        'deterime XAng
        ProjectNormal(0) = 1 'x
        ProjectNormal(1) = 0 'y
        ProjectNormal(2) = 0 'z
        'determine normal to wall plane
        WallNormal = Functions.CrossProduct(Coordin(0), Coordin(1), Coordin(2), Coordin(3), Coordin(4), Coordin(5), Coordin(6), Coordin(7), Coordin(8))
        'determine the cosine of the angle between unit vectors
        DotVar = Functions.DotProduct(ProjectNormal(0), ProjectNormal(1), ProjectNormal(2), 0, 0, 0, WallNormal(0), WallNormal(1), WallNormal(2))
        MagPN = (ProjectNormal(0) * ProjectNormal(0)) + (ProjectNormal(1) * ProjectNormal(1)) + (ProjectNormal(2) * ProjectNormal(2))
        MagWN = (WallNormal(0) * WallNormal(0)) + (WallNormal(1) * WallNormal(1)) + (WallNormal(2) * WallNormal(2))
        CosVar = DotVar / (Math.Sqrt(MagPN) * Math.Sqrt(MagWN))
        'determine angle
        If CosVar > 1 Then CosVar = 1
        If CosVar < -1 Then CosVar = -1
        XAng = Functions.Degrees(Math.Acos(CosVar))

        'deterime YAng
        ProjectNormal(0) = 0 'x
        ProjectNormal(1) = 1 'y
        ProjectNormal(2) = 0 'z
        'determine the cosine of the angle between unit vectors
        DotVar = Functions.DotProduct(ProjectNormal(0), ProjectNormal(1), ProjectNormal(2), 0, 0, 0, WallNormal(0), WallNormal(1), WallNormal(2))
        MagPN = (ProjectNormal(0) * ProjectNormal(0)) + (ProjectNormal(1) * ProjectNormal(1)) + (ProjectNormal(2) * ProjectNormal(2))
        CosVar = DotVar / (Math.Sqrt(MagPN) * Math.Sqrt(MagWN))
        'determine angle
        If CosVar > 1 Then CosVar = 1
        If CosVar < -1 Then CosVar = -1
        YAng = Functions.Degrees(Math.Acos(CosVar))

        'deterime ZAng
        ProjectNormal(0) = 0 'x
        ProjectNormal(1) = 0 'y
        ProjectNormal(2) = 1 'z
        'determine the cosine of the angle between unit vectors
        DotVar = Functions.DotProduct(ProjectNormal(0), ProjectNormal(1), ProjectNormal(2), 0, 0, 0, WallNormal(0), WallNormal(1), WallNormal(2))
        MagPN = (ProjectNormal(0) * ProjectNormal(0)) + (ProjectNormal(1) * ProjectNormal(1)) + (ProjectNormal(2) * ProjectNormal(2))
        CosVar = DotVar / (Math.Sqrt(MagPN) * Math.Sqrt(MagWN))
        'determine angle
        If CosVar > 1 Then CosVar = 1
        If CosVar < -1 Then CosVar = -1
        ZAng = Functions.Degrees(Math.Acos(CosVar))

        SpeakerLabel = SpeakerLabel 'Front:Xang=46 & Yang=90 Zang=43,Left:Xang=90 & Yang=27 & Zang=63,Right: Xang=90 & Yang=158 & Zang=68,Rear: Xang=180 & Yang=90 & Zang=90
        'If ZAng < 70 Then ZAng = ZAng + 180
        If Mid(SpeakerLabel, 2, 1) = "t" Then
            'If ZAng >= 135 Then
            TempWall(0) = "Ceiling"
            TempWall(1) = 0 'toe-in
            TempWall(2) = 90 - XAng + 90 'toe-up
            TempWall(3) = 90 - YAng  'roll
            'Else
            '    TempWall(0) = "Ceiling"
            '    TempWall(1) = 0 'toe-in
            '    TempWall(2) = 90 - XAng + 90 'toe-up
            '    TempWall(3) = 360 - ZAng '90 - YAng  'roll
            'End If
            '    ElseIf YAng <= 135 And YAng >= 45 Then 'Front or Rear
            '        If XAng < 90 Then 'Front:Xang=46 & Yang=90 Zang=43
            '            TempWall(0) = "Front"
            '            TempWall(1) = 90 - YAng 'toe-in
            '            TempWall(2) = 90 - ZAng 'toe-up
            '           TempWall(3) = 0 'roll
            '       ElseIf XAng >= 135 And XAng <= 225 Then 'Rear: Xang=180 & Yang=90 & Zang=90
            '            TempWall(0) = "Rear"
            '            TempWall(1) = 90 + YAng  'toe-in
            '            TempWall(2) = ZAng - 90 'toe-up
            '            TempWall(3) = 0 'roll
            '        End If
            '    ElseIf XAng <= 135 And XAng >= 45 Then 'Left or Right
            '        If YAng < 90 Then 'Left:  XAng = 90 & YAng = 27 & ZAng = 63
            '            TempWall(0) = "Left"
            '            TempWall(1) = 0 'XAng 'toe-in
            '            TempWall(2) = XAng '90 'toe-up
            '            TempWall(3) = 360 - ZAng '90 - ZAng 'roll
            '        ElseIf YAng >= 135 And YAng <= 225 Then 'Right: Xang=90 & Yang=158 & Zang=68
            '            TempWall(0) = "Right"
            '            TempWall(1) = 360 - XAng 'toe-in
            '            TempWall(2) = 0 'toe-up
            '            TempWall(3) = ZAng - 90 'roll
            '        End If
            '    End If
        Else
            If ZAng >= 135 Then
                TempWall(0) = "Ceiling"
                TempWall(1) = 0 'toe-in
                TempWall(2) = 90 - XAng + 90 'toe-up
                TempWall(3) = 90 - YAng  'roll
            ElseIf YAng <= 135 And YAng >= 45 Then 'Front or Rear
                If XAng < 90 Then 'Front:Xang=46 & Yang=90 Zang=43
                    TempWall(0) = "Front"
                    TempWall(1) = 90 - YAng 'toe-in
                    TempWall(2) = 90 - ZAng 'toe-up
                    TempWall(3) = 0 'roll
                ElseIf XAng >= 135 And XAng <= 225 Then 'Rear: Xang=180 & Yang=90 & Zang=90
                    TempWall(0) = "Rear"
                    TempWall(1) = 90 + YAng  'toe-in
                    TempWall(2) = ZAng - 90 'toe-up
                    TempWall(3) = 0 'roll
                End If
            ElseIf XAng <= 135 And XAng >= 45 Then 'Left or Right
                If YAng < 90 Then 'Left:  XAng = 90 & YAng = 27 & ZAng = 63
                    TempWall(0) = "Left"
                    TempWall(1) = XAng 'toe-in
                    TempWall(2) = 0 'toe-up
                    TempWall(3) = 90 - ZAng 'roll
                ElseIf YAng >= 135 And YAng <= 225 Then 'Right: Xang=90 & Yang=158 & Zang=68
                    TempWall(0) = "Right"
                    TempWall(1) = 360 - XAng 'toe-in
                    TempWall(2) = 0 'toe-up
                    TempWall(3) = ZAng - 90 'roll
                End If
            End If
        End If
        Return TempWall
    End Function

    Public Function DetermineViewbak(ByRef IndexNum As Integer, ByRef TempArray As Object) As String
        Dim FRSD As Single, RLSD As Single, CPSD As Single, NumArr(3) As Object
        Dim n As Integer, z As Integer, XMax As Single, YMax As Single, ZMax As Single
        Dim XMin As Single, YMin As Single, ZMin As Single
        Dim WallView As String = ""

        'Temparray(0, n) = dbRecordsetProject!WallIndex
        'Temparray(1, n) = dbRecordsetProject!ProjectNumber
        'Temparray(2, n) = dbRecordsetProject!Wallname
        'Temparray(3, n) = dbRecordsetProject!DrawRefIndex
        'Temparray(4, n) = dbRecordsetProject!MaterialIndex
        'Temparray(5, n) = dbRecordsetProject!X1
        'Temparray(6, n) = dbRecordsetProject!Y1
        'Temparray(7, n) = dbRecordsetProject!Z1
        'Temparray(8, n) = dbRecordsetProject!X2
        'Temparray(9, n) = dbRecordsetProject!Y2
        'Temparray(10, n) = dbRecordsetProject!Z2
        'Temparray(11, n) = dbRecordsetProject!X3
        'Temparray(12, n) = dbRecordsetProject!Y3
        'Temparray(13, n) = dbRecordsetProject!Z3
        'Temparray(14, n) = dbRecordsetProject!X4
        'Temparray(15, n) = dbRecordsetProject!Y4
        'Temparray(16, n) = dbRecordsetProject!Z4
        'Temparray(17, n) = dbRecordsetProject!Thickness

        XMax = 0
        YMax = 0
        ZMax = 0
        XMin = 999
        YMin = 999
        ZMin = 999

        'Determine if wall is (left or right), (front or back), (ceiling or floor)
        For n = 0 To UBound(TempArray, 2)
            'determine wall in question index
            'Determine X max and Y Min
            For z = 5 To 14 Step 3
                If TempArray(z, n) >= XMax Then XMax = TempArray(z, n)
                If TempArray(z, n) <= XMin Then XMin = TempArray(z, n)
            Next z
            For z = 6 To 15 Step 3
                If TempArray(z, n) >= YMax Then YMax = TempArray(z, n)
                If TempArray(z, n) <= YMin Then YMin = TempArray(z, n)
            Next z
            For z = 7 To 16 Step 3
                If TempArray(z, n) >= ZMax Then ZMax = TempArray(z, n)
                If TempArray(z, n) <= ZMin Then ZMin = TempArray(z, n)
            Next z
        Next n
        'Determine if wall is (left or right), (front or back), (Ceiling or Floor)
        NumArr(0) = TempArray(5, IndexNum)
        NumArr(1) = TempArray(8, IndexNum)
        NumArr(2) = TempArray(11, IndexNum)
        NumArr(3) = TempArray(14, IndexNum)
        FRSD = Functions.StandardDeviation(NumArr)

        NumArr(0) = TempArray(6, IndexNum)
        NumArr(1) = TempArray(9, IndexNum)
        NumArr(2) = TempArray(12, IndexNum)
        NumArr(3) = TempArray(15, IndexNum)
        RLSD = Functions.StandardDeviation(NumArr)

        NumArr(0) = TempArray(7, IndexNum)
        NumArr(1) = TempArray(10, IndexNum)
        NumArr(2) = TempArray(13, IndexNum)
        NumArr(3) = TempArray(16, IndexNum)
        CPSD = Functions.StandardDeviation(NumArr)

        If FRSD < RLSD And FRSD < CPSD Then
            If TempArray(5, IndexNum) < (XMin + XMax) / 2 Then
                WallView = "Front" 'If X1> XMin then Rear wall
            Else
                WallView = "Rear" 'If X1< XMax then Front wall
            End If
        End If
        If RLSD < FRSD And RLSD < CPSD Then
            If TempArray(6, IndexNum) < (YMin + YMax) / 2 Then
                WallView = "Right" 'If Y1< YMax then Right wall
            Else
                WallView = "Left" 'If Y1> YMin then Left wall
            End If
        End If
        If CPSD < FRSD And CPSD < RLSD Then
            If TempArray(7, IndexNum) < (ZMin + ZMax) / 2 Then
                WallView = "Ceiling" 'If Z1< ZMax then Ceiling wall
            Else
                WallView = "Plan" 'If Z1> ZMin then Plan
            End If
        End If

        Return WallView

    End Function
    Public Function CalcImagePoint(ByRef MainLP As Object, ByRef SpeakerArr As Object, ByRef SpkrPick As Integer, ByRef TopPnt As Object) As Object
        Dim AngleNum As Single, P1(2) As Single, P2(2) As Single, SpkrUse As String, ImageArr As Object, n As Integer, CeilingAdd As Single

        ReDim ImageArr(9)

        P1(0) = MainLP(0)
        P1(1) = MainLP(1)
        P1(2) = MainLP(2)

        For n = 0 To 2
            AngleNum = SpeakerArr(n, SpkrPick)
            ImageArr(9) = SpeakerArr(3, SpkrPick) 'speaker label
            CeilingAdd = MainLP(1)
            If Mid(SpeakerArr(3, SpkrPick), 1, 2) = "Lt" Then
                CeilingAdd = TopPnt(2, 0)
            ElseIf Mid(SpeakerArr(3, SpkrPick), 1, 2) = "Rt" Then
                CeilingAdd = TopPnt(2, 1)
            End If
            SpkrUse = SpeakerArr(4, SpkrPick)
            If AngleNum = 90 Then AngleNum = 90.1
            If AngleNum = 270 Then AngleNum = 270.1
            If AngleNum = 180 Then AngleNum = 180.1
            Select Case Left(SpkrUse, 4)
                Case "Ceil"
                    Select Case AngleNum 'RA
                        Case Is <= 90
                            AngleNum = 90 - AngleNum
                            P2(0) = P1(1) - Math.Tan((AngleNum * (Math.PI / 180))) * 50 'x
                            P2(1) = CeilingAdd 'MainLP(1)
                            P2(2) = P1(2) + 50 'z
                        Case 90 To 180
                            AngleNum = AngleNum - 90
                            P2(0) = P1(0) + Math.Tan((AngleNum * (Math.PI / 180))) * 50 '(P1(0) - P2(0)) 'x
                            P2(1) = CeilingAdd 'MainLP(1)
                            P2(2) = P1(2) + 50 'z
                    End Select
                Case Else
                    Select Case AngleNum 'azimuth
                        Case Is <= 90
                            P2(0) = P1(0) - 50 'x
                            P2(1) = P1(1) - Math.Tan((AngleNum * (Math.PI / 180))) * (P1(0) - P2(0)) 'y
                            P2(2) = P1(2) 'z
                        Case 90 To 180
                            AngleNum = AngleNum - 90
                            P2(0) = P1(0) + Math.Tan((AngleNum * (Math.PI / 180))) * 50 '(P1(0) - P2(0)) 'x
                            P2(1) = P1(1) - 50  'y
                            P2(2) = P1(2) 'z
                        Case 180 To 270
                            AngleNum = 270 - AngleNum
                            P2(0) = P1(0) + Math.Tan((AngleNum * (Math.PI / 180))) * 50  'x
                            P2(1) = P1(1) + 50  'y
                            P2(2) = P1(2) 'z
                        Case Is >= 270
                            P2(0) = P1(0) - 50  'x
                            P2(1) = P1(1) + Math.Tan(((360 - AngleNum) * (Math.PI / 180))) * (P1(0) - P2(0)) 'y
                            P2(2) = P1(2) 'z
                    End Select
            End Select
            If n = 0 Then
                ImageArr(0) = P2(0) 'x
                ImageArr(1) = P2(1) 'y
                ImageArr(2) = P2(2) 'z
            ElseIf n = 1 Then
                ImageArr(3) = P2(0) 'x
                ImageArr(4) = P2(1) 'y
                ImageArr(5) = P2(2) 'z
            ElseIf n = 2 Then
                ImageArr(6) = P2(0) 'x
                ImageArr(7) = P2(1) 'y
                ImageArr(8) = P2(2) 'z
            End If

        Next n

        Return ImageArr

    End Function
    Public Function CalcMirrorPnts(ByRef MainLP As Object, ByRef WallArr As Object, ByRef SpkrImage As Object _
, ByRef SpkrConfig As String, ByRef TopWidth As Object, ByRef SpeakerCnt As Integer, ByRef SpeakerInfo As Object) As Object
        Dim n As Integer, MPnt As Object, TempArr As Object, P1(2) As Single, P2(2) As Single, TopRight As Single, TopLeft As Single
        Dim Corn1(2) As Single, Corn2(2) As Single, Corn3(2) As Single, Corn4(2) As Single
        Dim BoxHeight As Single, HeightVar As Single, Cnt As Integer, IntResult As Object, WidthCor As Single
        Dim SurfIndex(2) As Long, WallCnt(2) As Integer, Coord0(8) As Single, Coord1(8) As Single, Coord2(8) As Single, MaxDim(3, 2) As Single
        Dim OutArr As Object, DistanceVar As Single, TempDis As Single, GoVar As Boolean, MaxDimX(3, 2) As Single
        Dim IsCeiling As Integer, MaxCeiling As Single, CeilingCnt As Integer, WallName As String
        Dim ModInitX As Single, ModInitY As Single, ModInitZ As Single, YMPnt(15) As Object
        Dim IncrementX As Single, IncrementY As Single, PM15 As Integer, MPntTU(15) As Object, MPntTD(15) As Object
        Dim SpkrHgt As Single, SpkrWth As Single, m As Integer, t As Integer, FactorX As Single, FactorY As Single, LengthCor As Single
        Dim BoxWidth As Single, FloorHeight As Single

        ReDim MPnt(15)

        ReDim IntResult(2, 2)

        LengthCor = 0
        WidthCor = 0
        IsCeiling = 0
        MaxCeiling = 999
        FloorHeight = 0
        If SpeakerInfo(4) = "Floorstanding" Then FloorHeight = MainLP(2) - 2
        SpkrHgt = SpeakerInfo(1)
        SpkrWth = SpeakerInfo(2)
        'the size of the box representing the placement tolerance of the speaker
        BoxHeight = SpkrHgt
        BoxWidth = SpkrWth
        If SpkrConfig = "Rs" Then
            n = n
        End If
        If SpkrConfig = "R" Then HeightVar = 0
        If SpkrConfig = "C" Then HeightVar = 0
        If SpkrConfig = "L" Then
            HeightVar = 0
        End If
        If SpkrConfig = "Rs" Then HeightVar = 0.5
        If SpkrConfig = "Ls" Then HeightVar = 0.5
        If SpkrConfig = "Rtr" Or SpkrConfig = "Rtf" Then
            HeightVar = 0
            WidthCor = TopWidth(2, 0) - MainLP(1)
            If SpkrConfig = "Rtr" Then LengthCor = 1
            IsCeiling = 0 '1
        End If
        If SpkrConfig = "Ltr" Or SpkrConfig = "Ltf" Then
            HeightVar = 0
            WidthCor = TopWidth(2, 1) - MainLP(1)
            If SpkrConfig = "Ltr" Then LengthCor = 1
            IsCeiling = 0 '1
        End If
        If SpkrConfig = "Lrs" Then
            HeightVar = 0.5
        End If
        If SpkrConfig = "Rrs" Then
            HeightVar = 0.5
        End If
        P1(0) = MainLP(0) + LengthCor
        P1(1) = MainLP(1) + WidthCor
        P1(2) = MainLP(2)
        For CeilingCnt = 0 To 1
            For Cnt = 0 To 2
                If Cnt = 0 Then
                    P2(0) = SpkrImage(0)
                    P2(1) = SpkrImage(1)
                    P2(2) = SpkrImage(2)
                ElseIf Cnt = 1 Then
                    P2(0) = SpkrImage(3)
                    P2(1) = SpkrImage(4)
                    P2(2) = SpkrImage(5)
                Else
                    P2(0) = SpkrImage(6)
                    P2(1) = SpkrImage(7)
                    P2(2) = SpkrImage(8)
                End If
                'If FloorHeight <> 0 And Mid(SpeakerInfo(0), 2, 1) <> "t" Then
                'P2(2) = FloorHeight
                'P1(2) = FloorHeight
                'End If
                If SpkrConfig = "Rtf" Or SpkrConfig = "Ltf" Then
                    DistanceVar = -999999
                Else
                    DistanceVar = 999999
                End If

                For n = 0 To UBound(WallArr, 2) + IsCeiling
                    If IsCeiling = 1 And n = UBound(WallArr, 2) + IsCeiling Then
                        WallName = "Y" 'ficticious ceiling added to show full range of angle intersections
                    Else
                        If WallArr(4, n) = "00000" Then
                            WallName = "N" 'if the wall is imaginary
                        Else
                            WallName = "Y"
                        End If
                    End If
                    If WallName = "Y" Then
                        If IsCeiling = 1 And n = UBound(WallArr, 2) + IsCeiling Then 'ficticious ceiling added to show full range of angle intersections when intersection is not possible on existing ceiling
                            Corn1(0) = -50
                            Corn1(1) = -50
                            Corn1(2) = MaxCeiling
                            Corn2(0) = 100
                            Corn2(1) = -50
                            Corn2(2) = MaxCeiling
                            Corn3(0) = 100
                            Corn3(1) = 100
                            Corn3(2) = MaxCeiling
                            Corn4(0) = -50
                            Corn4(1) = 100
                            Corn4(2) = MaxCeiling
                        Else
                            For m = 5 To 16
                                If WallArr(m, n) = 0 Then ' avoids math error caused by using 0 as a coordinate
                                    WallArr(m, n) = 0.00005
                                End If
                            Next m
                            Corn1(0) = WallArr(5, n)
                            Corn1(1) = WallArr(6, n)
                            Corn1(2) = WallArr(7, n)
                            Corn2(0) = WallArr(8, n)
                            Corn2(1) = WallArr(9, n)
                            Corn2(2) = WallArr(10, n)
                            Corn3(0) = WallArr(11, n)
                            Corn3(1) = WallArr(12, n)
                            Corn3(2) = WallArr(13, n)
                            Corn4(0) = WallArr(14, n)
                            Corn4(1) = WallArr(15, n)
                            Corn4(2) = WallArr(16, n)
                            For t = 7 To 16 Step 3
                                If WallArr(t, n) > MaxCeiling Then
                                    MaxCeiling = WallArr(t, n)
                                End If
                            Next t
                        End If
                        TempArr = Functions.Intersection(P1, P2, Corn1, Corn2, Corn3, Corn4)
                        If TempArr(0) <> 0 And TempArr(1) <> 0 And TempArr(2) <> 0 Then
                            TempDis = Functions.PythagDist(P1(0), CSng(TempArr(0)), P1(1), CSng(TempArr(1)))
                            GoVar = False
                            If SpkrConfig = "Rtf" Or SpkrConfig = "Ltf" Then
                                If TempDis >= DistanceVar Then
                                    GoVar = True
                                End If
                            Else
                                If TempDis < DistanceVar Then
                                    GoVar = True
                                End If
                            End If
                            If IsCeiling = 1 And n = UBound(WallArr, 2) + IsCeiling Then
                                GoVar = True
                            End If
                            If GoVar Then
                                IntResult(0, Cnt) = TempArr(0)
                                IntResult(1, Cnt) = TempArr(1)
                                IntResult(2, Cnt) = TempArr(2)
                                DistanceVar = TempDis
                            End If
                        End If
                    End If
                Next n
            Next Cnt
            If IntResult(0, 2) <> 0 And IntResult(1, 2) <> 0 And IntResult(2, 2) <> 0 Then
                If IntResult(2, 2) + 3 > MainLP(2) Then
                    IsCeiling = 1 'add ficticious ceiling
                    n = UBound(WallArr, 2) + IsCeiling 'begin n range at ficticious ceiling
                Else
                    CeilingCnt = 1 'avoids use of ficticious ceiling if intersection is found
                End If
            Else
                IsCeiling = 1 'add ficticious ceiling
                n = UBound(WallArr, 2) + IsCeiling 'begin n range at ficticious ceiling
            End If
        Next CeilingCnt

        'Process resulting intersections per speaker type and wall
        Select Case SpkrConfig
            Case "R"
                MPnt(0) = IntResult(0, 2) 'preferred position
                MPnt(1) = IntResult(1, 2)
                MPnt(2) = IntResult(2, 2)
                ModInitX = IntResult(0, 2)
                ModInitY = IntResult(1, 2)
                ModInitZ = IntResult(2, 2)
                OutArr = FindWallIndex("FrontR", CSng(MainLP(0)), CSng(ModInitY), CSng(ModInitZ), WallArr, 120, BoxWidth)
                MPnt(0) = OutArr(0) '
                MPnt(1) = OutArr(1) '
                MPnt(2) = OutArr(2) '
                MPnt(15) = OutArr(3) '
                'Check if too far right
                OutArr = FindWallIndex("Right", CSng(MainLP(0) - 6), CSng(MainLP(1)), CSng(MainLP(2)), WallArr, 1, BoxWidth)
                YMPnt(0) = OutArr(0) '
                YMPnt(1) = OutArr(1) '
                YMPnt(2) = OutArr(2) '
                YMPnt(15) = OutArr(3) '
                If YMPnt(1) > MPnt(1) Then
                    OutArr = FindWallIndex("FrontR", CSng(MainLP(0)), CSng(YMPnt(1)), CSng(ModInitZ), WallArr, 120, BoxWidth)
                    MPnt(0) = OutArr(0) '
                    MPnt(1) = OutArr(1) '
                    MPnt(2) = OutArr(2) '
                    MPnt(15) = OutArr(3) '
                End If

                'Check height spacing
                If FloorHeight = 0 Then 'if not floorstanding speakers
                    If MPnt(15) = 99 Then MPnt(1) = ModInitY
                    OutArr = FindWallIndex("HeightFD", CSng(MainLP(0)), CSng(MPnt(1)), CSng(LeftHeight), WallArr, 50, BoxHeight) 'use left speaker height previously calculated to match height
                    MPnt(0) = OutArr(0) '
                    MPnt(1) = OutArr(1) '
                    MPnt(2) = OutArr(2) '
                    MPnt(15) = OutArr(3) '
                    If MPnt(2) < MainLP(2) + 0.5 Or MPnt(15) = 99 Then
                        'If MPnt(15) = 99 Then MPnt(1) = ModInitY
                        OutArr = FindWallIndex("HeightFU", CSng(MainLP(0)), CSng(MPnt(1)), CSng(LeftHeight), WallArr, 50, BoxHeight)
                        MPnt(0) = OutArr(0) '
                        MPnt(1) = OutArr(1) '
                        MPnt(2) = OutArr(2) '
                        MPnt(15) = OutArr(3) '
                    End If
                Else
                    'Determine BaseHeight; height above floor for floorstanding speakers
                    OutArr = FindWallIndex("BaseHeightF", CSng(MPnt(0) + 0.4), CSng(MPnt(1)), CSng(FloorHeight), WallArr, 120, BoxHeight)
                    BaseHeight(SpeakerCnt) = OutArr(2) '
                    FloorHeight = BaseHeight(SpeakerCnt) + BoxHeight / 2
                    MPnt(2) = FloorHeight
                    IntResult(2, 0) = FloorHeight
                    IntResult(2, 1) = FloorHeight
                End If

                MPnt(3) = IntResult(0, 0) 'x top left point 1
                MPnt(4) = IntResult(1, 0) 'y
                MPnt(5) = IntResult(2, 0) - BoxHeight / 2 'z

                MPnt(6) = IntResult(0, 1) 'x top right point 2
                MPnt(7) = IntResult(1, 1) 'y
                MPnt(8) = IntResult(2, 1) - BoxHeight / 2 'z

                MPnt(9) = IntResult(0, 1) 'x bottom right point 3
                MPnt(10) = IntResult(1, 1) 'y
                MPnt(11) = IntResult(2, 1) + BoxHeight / 2 'z

                MPnt(12) = IntResult(0, 0) 'x bottom left point 4
                MPnt(13) = IntResult(1, 0) 'y
                MPnt(14) = IntResult(2, 0) + BoxHeight / 2 'z

            Case "C"
                MPnt(0) = IntResult(0, 2) 'x center point 0
                MPnt(1) = (TopWidth(2, 0) + TopWidth(2, 1)) / 2 'IntResult(1, 2) 'y
                MPnt(2) = LeftHeight 'use left speaker height previously calculated to match height z
                ModInitX = IntResult(0, 2)
                ModInitY = (TopWidth(2, 0) + TopWidth(2, 1)) / 2 'IntResult(1, 2)
                'Check height spacing
                OutArr = FindWallIndex("HeightFD", CSng(MainLP(0)), CSng(ModInitY), CSng(LeftHeight), WallArr, 50, BoxWidth)
                MPnt(0) = OutArr(0) '
                MPnt(1) = OutArr(1) '
                MPnt(2) = OutArr(2) '
                MPnt(15) = OutArr(3) '
                'if too low
                If MPnt(2) < MainLP(2) + 0.5 Or MPnt(15) = 99 Then
                    OutArr = FindWallIndex("HeightFU", CSng(MainLP(0)), CSng(ModInitY), CSng(LeftHeight), WallArr, 50, BoxWidth)
                    MPnt(0) = OutArr(0) '
                    MPnt(1) = OutArr(1) '
                    MPnt(2) = OutArr(2) '
                    MPnt(15) = OutArr(3) '
                End If

                If FloorHeight <> 0 Then  'use for floorstanding speakers
                    'Determine BaseHeight; height above floor
                    OutArr = FindWallIndex("BaseHeightF", CSng(MPnt(0) + 0.4), CSng(MPnt(1)), CSng(FloorHeight), WallArr, 120, BoxHeight)
                    BaseHeight(SpeakerCnt) = OutArr(2) '
                    FloorHeight = BaseHeight(SpeakerCnt) + BoxHeight / 2
                    MPnt(2) = FloorHeight
                    IntResult(2, 0) = FloorHeight
                    IntResult(2, 1) = FloorHeight
                End If

                MPnt(3) = IntResult(0, 0) 'x top left point 1
                MPnt(4) = IntResult(1, 0) 'y
                MPnt(5) = IntResult(2, 0) - BoxHeight / 2 'z

                MPnt(6) = IntResult(0, 1) 'x top right point 2
                MPnt(7) = IntResult(1, 1) 'y
                MPnt(8) = IntResult(2, 1) - BoxHeight / 2 'z

                MPnt(9) = IntResult(0, 1) 'x bottom right point 3
                MPnt(10) = IntResult(1, 1) 'y
                MPnt(11) = IntResult(2, 1) + BoxHeight / 2 'z

                MPnt(12) = IntResult(0, 0) 'x bottom left point 4
                MPnt(13) = IntResult(1, 0) 'y
                MPnt(14) = IntResult(2, 0) + BoxHeight / 2 'z

            Case "L"
                MPnt(0) = IntResult(0, 2) 'preferred position
                MPnt(1) = IntResult(1, 2)
                MPnt(2) = IntResult(2, 2)
                ModInitX = IntResult(0, 2)
                ModInitY = IntResult(1, 2)
                ModInitZ = IntResult(2, 2)
                'move position inward until a fit
                OutArr = FindWallIndex("FrontL", CSng(MainLP(0)), CSng(ModInitY), CSng(ModInitZ), WallArr, 120, BoxWidth)
                MPnt(0) = OutArr(0) '
                MPnt(1) = OutArr(1) '
                MPnt(2) = OutArr(2) '
                MPnt(15) = OutArr(3) '
                'Check if too far left
                OutArr = FindWallIndex("Left", CSng(MainLP(0) - 6), CSng(MainLP(1)), CSng(MainLP(2)), WallArr, 1, BoxWidth)
                YMPnt(0) = OutArr(0) '
                YMPnt(1) = OutArr(1) '
                YMPnt(2) = OutArr(2) '
                YMPnt(15) = OutArr(3) '
                If YMPnt(1) < MPnt(1) Then
                    OutArr = FindWallIndex("FrontL", CSng(MainLP(0)), CSng(YMPnt(1)), CSng(ModInitZ), WallArr, 120, BoxWidth)
                    MPnt(0) = OutArr(0) '
                    MPnt(1) = OutArr(1) '
                    MPnt(2) = OutArr(2) '
                    MPnt(15) = OutArr(3) '
                End If

                'Check height spacing
                If FloorHeight = 0 Then 'if not floorstanding speakers
                    If MPnt(15) = 99 Then MPnt(1) = ModInitY
                    OutArr = FindWallIndex("HeightFD", CSng(MainLP(0)), CSng(MPnt(1)), CSng(MainLP(2)), WallArr, 50, BoxHeight)
                    MPnt(0) = OutArr(0) '
                    MPnt(1) = OutArr(1) '
                    MPnt(2) = OutArr(2) '
                    MPnt(15) = OutArr(3) '
                    'if too low
                    If MPnt(2) < MainLP(2) + 0.5 Or MPnt(15) = 99 Then
                        'If MPnt(15) = 99 Then MPnt(1) = ModInitY
                        OutArr = FindWallIndex("HeightFU", CSng(MainLP(0)), CSng(MPnt(1)), CSng(MainLP(2)), WallArr, 50, BoxHeight)
                        MPnt(0) = OutArr(0) '
                        MPnt(1) = OutArr(1) '
                        MPnt(2) = OutArr(2) '
                        MPnt(15) = OutArr(3) '
                    End If
                Else
                    'Determine BaseHeight; height above floor for floorstanding speakers
                    OutArr = FindWallIndex("BaseHeightF", CSng(MPnt(0) + 0.4), CSng(MPnt(1)), CSng(FloorHeight), WallArr, 120, BoxHeight)
                    BaseHeight(SpeakerCnt) = OutArr(2) '
                    FloorHeight = BaseHeight(SpeakerCnt) + BoxHeight / 2
                    MPnt(2) = FloorHeight
                    IntResult(2, 0) = FloorHeight
                    IntResult(2, 1) = FloorHeight
                End If

                MPnt(3) = IntResult(0, 0) 'x top left point 1
                MPnt(4) = IntResult(1, 0) 'y
                MPnt(5) = IntResult(2, 0) - BoxHeight / 2 'z

                MPnt(6) = IntResult(0, 1) 'x top right point 2
                MPnt(7) = IntResult(1, 1) 'y
                MPnt(8) = IntResult(2, 1) - BoxHeight / 2 'z

                MPnt(9) = IntResult(0, 1) 'x bottom right point 3
                MPnt(10) = IntResult(1, 1) 'y
                MPnt(11) = IntResult(2, 1) + BoxHeight / 2 'z

                MPnt(12) = IntResult(0, 0) 'x bottom left point 4
                MPnt(13) = IntResult(1, 0) 'y
                MPnt(14) = IntResult(2, 0) + BoxHeight / 2 'z

                LeftHeight = MPnt(2)


            Case "Lrs"
                'TopWidth(2, 1) Left speaker position
                MPnt(0) = IntResult(0, 0) 'closest position to front
                MPnt(1) = IntResult(1, 0)
                MPnt(2) = SurroundHeight 'use surround height of Ls to match surround height
                ModInitX = IntResult(0, 0)
                ModInitY = IntResult(1, 0)
                ModInitZ = IntResult(2, 2)
                'move position inward until a fit
                OutArr = FindWallIndex("RearL", CSng(MainLP(0)), CSng(ModInitY), CSng(ModInitZ), WallArr, 120, BoxWidth)
                MPnt(0) = OutArr(0) '
                MPnt(1) = OutArr(1) '
                MPnt(2) = OutArr(2) '
                MPnt(15) = OutArr(3) '
                'Check if too far left
                OutArr = FindWallIndex("Left", CSng(MainLP(0)), CSng(MainLP(1)), CSng(MainLP(2)), WallArr, 1, BoxWidth)
                YMPnt(0) = OutArr(0) '
                YMPnt(1) = OutArr(1) '
                YMPnt(2) = OutArr(2) '
                YMPnt(15) = OutArr(3) '
                If YMPnt(1) + 1 > MPnt(1) Then
                    OutArr = FindWallIndex("RearL", CSng(MainLP(0)), CSng(YMPnt(1) + 1), CSng(ModInitZ), WallArr, 120, BoxWidth)
                    MPnt(0) = OutArr(0) '
                    MPnt(1) = OutArr(1) '
                    MPnt(2) = OutArr(2) '
                    MPnt(15) = OutArr(3) '
                End If

                'Check height spacing
                If FloorHeight = 0 Then 'not floorstanding speakers
                    'move position down until a fit
                    If MPnt(15) = 99 Then MPnt(1) = ModInitY
                    OutArr = FindWallIndex("HeightRrD", CSng(MainLP(0)), CSng(MPnt(1)), CSng(SurroundHeight), WallArr, 50, BoxHeight)
                    MPnt(0) = OutArr(0) '
                    MPnt(1) = OutArr(1) '
                    MPnt(2) = OutArr(2) '
                    MPnt(15) = OutArr(3) '
                    'if location is too low
                    If MPnt(2) < MainLP(2) Then
                        If MPnt(15) = 99 Then MPnt(1) = ModInitY
                        OutArr = FindWallIndex("HeightRrU", CSng(MainLP(0)), CSng(MPnt(1)), CSng(SurroundHeight), WallArr, 50, BoxHeight)
                        MPnt(0) = OutArr(0) '
                        MPnt(1) = OutArr(1) '
                        MPnt(2) = OutArr(2) '
                        MPnt(15) = OutArr(3) '
                    End If
                Else
                    'Determine BaseHeight; height above floor for floorstanding speakers
                    OutArr = FindWallIndex("BaseHeightR", CSng(MPnt(0) - 0.4), CSng(MPnt(1)), CSng(FloorHeight), WallArr, 120, BoxHeight)
                    BaseHeight(SpeakerCnt) = OutArr(2) '
                    FloorHeight = BaseHeight(SpeakerCnt) + BoxHeight / 2
                    MPnt(2) = FloorHeight
                    IntResult(2, 0) = FloorHeight
                    IntResult(2, 1) = FloorHeight
                End If

                MPnt(3) = IntResult(0, 0) 'x top left point 1
                MPnt(4) = IntResult(1, 0) 'y
                MPnt(5) = IntResult(2, 0) + BoxHeight + HeightVar  'z'c

                MPnt(6) = IntResult(0, 1) 'x top right point 2
                MPnt(7) = IntResult(1, 1) 'y
                MPnt(8) = IntResult(2, 1) + BoxHeight + HeightVar  'z'c

                MPnt(9) = IntResult(0, 1) 'x bottom right point 3
                MPnt(10) = IntResult(1, 1) 'y
                MPnt(11) = IntResult(2, 1) + HeightVar  'z'c

                MPnt(12) = IntResult(0, 0) 'x bottom left point 4
                MPnt(13) = IntResult(1, 0) 'y
                MPnt(14) = IntResult(2, 0) + HeightVar  'z'c

            Case "Rrs"
                MPnt(0) = IntResult(0, 1)
                MPnt(1) = IntResult(1, 1)
                MPnt(2) = SurroundHeight 'Surround height is established by the Ls speaker IntResult(2, 2) - BoxHeight / 2 - HeightVar
                ModInitX = IntResult(0, 1)
                ModInitY = IntResult(1, 1)
                ModInitZ = IntResult(2, 2)
                'move position inward until a fit
                OutArr = FindWallIndex("RearR", CSng(MainLP(0)), CSng(ModInitY), CSng(ModInitZ), WallArr, 120, BoxWidth)
                MPnt(0) = OutArr(0) '
                MPnt(1) = OutArr(1) '
                MPnt(2) = OutArr(2) '
                MPnt(15) = OutArr(3) '
                'Check if too far right
                OutArr = FindWallIndex("Right", CSng(MainLP(0)), CSng(MainLP(1)), CSng(MainLP(2)), WallArr, 1, BoxWidth)
                YMPnt(0) = OutArr(0) '
                YMPnt(1) = OutArr(1) '
                YMPnt(2) = OutArr(2) '
                YMPnt(15) = OutArr(3) '
                If YMPnt(1) - 1 < MPnt(1) Then
                    OutArr = FindWallIndex("RearR", CSng(MainLP(0)), CSng(YMPnt(1) - 1), CSng(ModInitZ), WallArr, 120, BoxWidth)
                    MPnt(0) = OutArr(0) '
                    MPnt(1) = OutArr(1) '
                    MPnt(2) = OutArr(2) '
                    MPnt(15) = OutArr(3) '
                End If

                'Check height spacing
                If FloorHeight = 0 Then 'not floorstanding speakers
                    'move position downward until a fit
                    If MPnt(15) = 99 Then MPnt(1) = ModInitY
                    OutArr = FindWallIndex("HeightRrD", CSng(MainLP(0)), CSng(MPnt(1)), CSng(SurroundHeight), WallArr, 100, BoxHeight)
                    MPnt(0) = OutArr(0) '
                    MPnt(1) = OutArr(1) '
                    MPnt(2) = OutArr(2) '
                    MPnt(15) = OutArr(3) '
                    'if location is too low
                    If MPnt(2) < MainLP(2) Then
                        If MPnt(15) = 99 Then MPnt(1) = ModInitY
                        OutArr = FindWallIndex("HeightRrU", CSng(MainLP(0)), CSng(MPnt(1)), CSng(SurroundHeight), WallArr, 100, BoxHeight)
                        MPnt(0) = OutArr(0) '
                        MPnt(1) = OutArr(1) '
                        MPnt(2) = OutArr(2) '
                        MPnt(15) = OutArr(3) '
                    End If
                Else
                    'Determine BaseHeight; height above floor for floorstanding speakers
                    OutArr = FindWallIndex("BaseHeightL", CSng(MPnt(0) - 0.4), CSng(MPnt(1)), CSng(FloorHeight), WallArr, 120, BoxHeight)
                    BaseHeight(SpeakerCnt) = OutArr(2) '
                    FloorHeight = BaseHeight(SpeakerCnt) + BoxHeight / 2
                    MPnt(2) = FloorHeight
                    IntResult(2, 0) = FloorHeight
                    IntResult(2, 1) = FloorHeight
                End If

                MPnt(3) = IntResult(0, 0) 'x top left point 1
                MPnt(4) = IntResult(1, 0) 'y
                MPnt(5) = IntResult(2, 0) + BoxHeight + HeightVar  'z'c

                MPnt(6) = IntResult(0, 1) 'x top right point 2
                MPnt(7) = IntResult(1, 1) 'y
                MPnt(8) = IntResult(2, 1) + BoxHeight + HeightVar  'z'c

                MPnt(9) = IntResult(0, 1) 'x bottom right point 3
                MPnt(10) = IntResult(1, 1) 'y
                MPnt(11) = IntResult(2, 1) + HeightVar  'z'c

                MPnt(12) = IntResult(0, 0) 'x bottom left point 4
                MPnt(13) = IntResult(1, 0) 'y
                MPnt(14) = IntResult(2, 0) + HeightVar  'z'c

            Case "Rs"
                MPnt(0) = IntResult(0, 2) 'preferred position
                MPnt(1) = IntResult(1, 2)
                MPnt(2) = SurroundHeight 'IntResult(2, 2) - BoxHeight / 2 - HeightVar 'z
                ModInitX = IntResult(0, 2)
                ModInitY = IntResult(1, 2)
                ModInitZ = IntResult(2, 2)
                'move position forward until a fit
                OutArr = FindWallIndex("RightF", CSng(ModInitX), CSng(MainLP(1)), CSng(ModInitZ), WallArr, 120, BoxWidth)
                MPnt(0) = OutArr(0) '
                MPnt(1) = OutArr(1) '
                MPnt(2) = OutArr(2) '
                MPnt(15) = OutArr(3) '
                If MPnt(15) = 99 Then
                    'if no result try moving backward
                    OutArr = FindWallIndex("Right", CSng(ModInitX), CSng(MainLP(1)), CSng(ModInitZ), WallArr, 120, BoxWidth)
                    MPnt(0) = OutArr(0) '
                    MPnt(1) = OutArr(1) '
                    MPnt(2) = OutArr(2) '
                    MPnt(15) = OutArr(3) '
                End If

                'Check height spacing
                If FloorHeight = 0 Then 'not for floorstanding speakers
                    'move downward to find a fit
                    If MPnt(15) = 99 Then MPnt(0) = ModInitX
                    OutArr = FindWallIndex("HeightRD", CSng(MPnt(0)), CSng(MainLP(1)), CSng(SurroundHeight), WallArr, 100, BoxHeight)
                    MPnt(0) = OutArr(0) '
                    MPnt(1) = OutArr(1) '
                    MPnt(2) = OutArr(2) '
                    MPnt(15) = OutArr(3) '
                    'if location is too low
                    If MPnt(2) < MainLP(2) + 0.5 Or MPnt(15) = 99 Then
                        If MPnt(15) = 99 Then MPnt(0) = ModInitX
                        OutArr = FindWallIndex("HeightRU", CSng(MPnt(0)), CSng(MainLP(1)), CSng(SurroundHeight), WallArr, 100, BoxHeight)
                        MPnt(0) = OutArr(0) '
                        MPnt(1) = OutArr(1) '
                        MPnt(2) = OutArr(2) '
                        MPnt(15) = OutArr(3) '
                    End If
                Else
                    'Determine BaseHeight; height above floor for floorstanding speakers
                    OutArr = FindWallIndex("BaseHeightL", CSng(MPnt(0)), CSng(MPnt(1) - 0.4), CSng(FloorHeight), WallArr, 120, BoxHeight)
                    BaseHeight(SpeakerCnt) = OutArr(2) '
                    FloorHeight = BaseHeight(SpeakerCnt) + BoxHeight / 2
                    MPnt(2) = FloorHeight
                    IntResult(2, 0) = FloorHeight
                    IntResult(2, 1) = FloorHeight
                End If

                MPnt(3) = IntResult(0, 0) 'x top left point 1
                MPnt(4) = IntResult(1, 0) 'y
                MPnt(5) = IntResult(2, 0) + BoxHeight + HeightVar 'z'c

                MPnt(6) = IntResult(0, 1) 'x top right point 2
                MPnt(7) = IntResult(1, 1) 'y
                MPnt(8) = IntResult(2, 1) + BoxHeight + HeightVar  'z'c

                MPnt(9) = IntResult(0, 1) 'x bottom right point 3
                MPnt(10) = IntResult(1, 1) 'y
                MPnt(11) = IntResult(2, 1) + HeightVar 'z'c

                MPnt(12) = IntResult(0, 0) 'x bottom left point 4
                MPnt(13) = IntResult(1, 0) 'y
                MPnt(14) = IntResult(2, 0) + HeightVar  'z'c

            Case "Ls"
                MPnt(0) = IntResult(0, 2) 'preferred position is farthest back
                MPnt(1) = IntResult(1, 2)
                If FloorHeight = 0 Then
                    MPnt(2) = IntResult(2, 2) + BoxHeight / 2 + HeightVar 'z 
                Else
                    MPnt(2) = IntResult(2, 2)
                End If
                SurroundHeight = MPnt(2)
                ModInitX = IntResult(0, 2)
                ModInitY = IntResult(1, 2)
                ModInitZ = IntResult(2, 2)
                OutArr = FindWallIndex("LeftF", CSng(ModInitX), CSng(MainLP(1)), CSng(ModInitZ), WallArr, 120, BoxWidth)
                MPnt(0) = OutArr(0) '
                MPnt(1) = OutArr(1) '
                MPnt(2) = OutArr(2) '
                MPnt(15) = OutArr(3) '
                If MPnt(15) = 99 Then
                    'if no fit is found
                    OutArr = FindWallIndex("Left", CSng(ModInitX), CSng(MainLP(1)), CSng(ModInitZ), WallArr, 120, BoxWidth)
                    MPnt(0) = OutArr(0) '
                    MPnt(1) = OutArr(1) '
                    MPnt(2) = OutArr(2) '
                    MPnt(15) = OutArr(3) '
                End If

                'Check height spacing
                If FloorHeight = 0 Then 'not for floorstanding speakers
                    'move downward to find a fit
                    If MPnt(15) = 99 Then MPnt(0) = ModInitX
                    OutArr = FindWallIndex("HeightLD", CSng(MPnt(0)), CSng(MainLP(1)), CSng(SurroundHeight), WallArr, 100, BoxHeight)
                    MPnt(0) = OutArr(0) '
                    MPnt(1) = OutArr(1) '
                    MPnt(2) = OutArr(2) '
                    MPnt(15) = OutArr(3) '
                    'if location is too low
                    If MPnt(2) < MainLP(2) + 0.5 Or MPnt(15) = 99 Then
                        If MPnt(15) = 99 Then MPnt(0) = ModInitX
                        OutArr = FindWallIndex("HeightLU", CSng(MPnt(0)), CSng(MainLP(1)), CSng(SurroundHeight), WallArr, 100, BoxHeight)
                        MPnt(0) = OutArr(0) '
                        MPnt(1) = OutArr(1) '
                        MPnt(2) = OutArr(2) '
                        MPnt(15) = OutArr(3) '
                    End If
                Else
                    'Determine BaseHeight; height above floor for floorstanding speakers
                    OutArr = FindWallIndex("BaseHeightR", CSng(MPnt(0)), CSng(MPnt(1) + 0.4), CSng(FloorHeight), WallArr, 120, BoxHeight)
                    BaseHeight(SpeakerCnt) = OutArr(2) '
                    FloorHeight = BaseHeight(SpeakerCnt) + BoxHeight / 2
                    MPnt(2) = FloorHeight
                    IntResult(2, 0) = FloorHeight
                    IntResult(2, 1) = FloorHeight
                End If

                SurroundHeight = MPnt(2)

                MPnt(3) = IntResult(0, 0) 'x top left point 1
                MPnt(4) = IntResult(1, 0) 'y
                MPnt(5) = IntResult(2, 0) + BoxHeight + HeightVar 'z'c

                MPnt(6) = IntResult(0, 1) 'x top right point 2
                MPnt(7) = IntResult(1, 1) 'y
                MPnt(8) = IntResult(2, 1) + BoxHeight + HeightVar  'z'c

                MPnt(9) = IntResult(0, 1) 'x bottom right point 3
                MPnt(10) = IntResult(1, 1) 'y
                MPnt(11) = IntResult(2, 1) + HeightVar 'z'c

                MPnt(12) = IntResult(0, 0) 'x bottom left point 4
                MPnt(13) = IntResult(1, 0) 'y
                MPnt(14) = IntResult(2, 0) + HeightVar  'z'c

            Case "Rw"
                MPnt(0) = IntResult(0, 2) 'preferred position
                MPnt(1) = IntResult(1, 2)
                MPnt(2) = IntResult(2, 2)  'LeftHeight 'z
                ModInitX = IntResult(0, 2)
                ModInitY = IntResult(1, 2)
                ModInitZ = IntResult(2, 2)
                'move position backward until a fit
                OutArr = FindWallIndex("Right", CSng(ModInitX), CSng(MainLP(1)), CSng(ModInitZ), WallArr, 100, BoxWidth)
                MPnt(0) = OutArr(0) '
                MPnt(1) = OutArr(1) '
                MPnt(2) = OutArr(2) '
                MPnt(15) = OutArr(3) '
                'Check if too far left
                OutArr = FindWallIndex("Right", CSng(2), CSng(MainLP(1)), CSng(MainLP(2)), WallArr, 1, BoxWidth) 'find width position of front-side wall
                YMPnt(0) = OutArr(0) '
                YMPnt(1) = OutArr(1) '
                YMPnt(2) = OutArr(2) '
                YMPnt(15) = OutArr(3) '
                If YMPnt(1) > MPnt(1) - 2 Or MPnt(15) = 99 Then
                    'move position forward until a fit
                    OutArr = FindWallIndex("RightF", CSng(ModInitX), CSng(MainLP(1)), CSng(ModInitZ), WallArr, 100, BoxWidth)
                    MPnt(0) = OutArr(0) '
                    MPnt(1) = OutArr(1) '
                    MPnt(2) = OutArr(2) '
                    MPnt(15) = OutArr(3)
                End If

                'Check height spacing
                'move downward to find a fit
                If FloorHeight = 0 Then 'do not use for floorstanding speakers
                    If MPnt(15) = 99 Then MPnt(0) = ModInitX
                    OutArr = FindWallIndex("HeightRD", CSng(MPnt(0)), CSng(MainLP(1)), CSng(ModInitZ), WallArr, 100, BoxHeight)
                    MPnt(0) = OutArr(0) '
                    MPnt(1) = OutArr(1) '
                    MPnt(2) = OutArr(2) '
                    MPnt(15) = OutArr(3) '
                    'if location is too low
                    If MPnt(2) < MainLP(2) + 0.5 Or MPnt(15) = 99 Then
                        If MPnt(15) = 99 Then MPnt(0) = ModInitX
                        OutArr = FindWallIndex("HeightRU", CSng(MPnt(0)), CSng(MainLP(1)), CSng(ModInitZ), WallArr, 100, BoxHeight)
                        MPnt(0) = OutArr(0) '
                        MPnt(1) = OutArr(1) '
                        MPnt(2) = OutArr(2) '
                        MPnt(15) = OutArr(3) '
                    End If
                Else
                    'Determine BaseHeight; height above floor for floorstanding speakers
                    OutArr = FindWallIndex("BaseHeightL", CSng(MPnt(0)), CSng(MPnt(1) - 0.4), CSng(FloorHeight), WallArr, 120, BoxHeight)
                    BaseHeight(SpeakerCnt) = OutArr(2) '
                    FloorHeight = BaseHeight(SpeakerCnt) + BoxHeight / 2
                    MPnt(2) = FloorHeight
                    IntResult(2, 0) = FloorHeight
                    IntResult(2, 1) = FloorHeight
                End If

                MPnt(3) = IntResult(0, 0) 'x top left point 1
                MPnt(4) = IntResult(1, 0) 'y
                MPnt(5) = IntResult(2, 0) - BoxHeight / 2  'z

                MPnt(6) = IntResult(0, 1) 'x top right point 2
                MPnt(7) = IntResult(1, 1) 'y
                MPnt(8) = IntResult(2, 1) - BoxHeight / 2   'z

                MPnt(9) = IntResult(0, 1) 'x bottom right point 3
                MPnt(10) = IntResult(1, 1) 'y
                MPnt(11) = IntResult(2, 1) + BoxHeight / 2   'z

                MPnt(12) = IntResult(0, 0) 'x bottom left point 4
                MPnt(13) = IntResult(1, 0) 'y
                MPnt(14) = IntResult(2, 0) + BoxHeight / 2   'z

            Case "Lw"
                MPnt(0) = IntResult(0, 2) 'preferred position
                MPnt(1) = IntResult(1, 2)
                MPnt(2) = IntResult(2, 2)  'LeftHeight 'z
                ModInitX = IntResult(0, 2)
                ModInitY = IntResult(1, 2)
                ModInitZ = IntResult(2, 2)
                'move position backward until a fit
                OutArr = FindWallIndex("Left", CSng(ModInitX), CSng(MainLP(1)), CSng(ModInitZ), WallArr, 100, BoxWidth)
                MPnt(0) = OutArr(0) '
                MPnt(1) = OutArr(1) '
                MPnt(2) = OutArr(2) '
                MPnt(15) = OutArr(3) '
                'Check if too far left
                OutArr = FindWallIndex("Left", CSng(2), CSng(MainLP(1)), CSng(MainLP(2)), WallArr, 1, BoxWidth) 'find width position of front-side wall
                YMPnt(0) = OutArr(0) '
                YMPnt(1) = OutArr(1) '
                YMPnt(2) = OutArr(2) '
                YMPnt(15) = OutArr(3) '
                If YMPnt(1) > MPnt(1) + 2 Or MPnt(15) = 99 Then
                    'move position forward until a fit
                    OutArr = FindWallIndex("LeftF", CSng(ModInitX), CSng(MainLP(1)), CSng(ModInitZ), WallArr, 100, BoxWidth)
                    MPnt(0) = OutArr(0) '
                    MPnt(1) = OutArr(1) '
                    MPnt(2) = OutArr(2) '
                    MPnt(15) = OutArr(3)
                End If

                'Check height spacing
                If FloorHeight = 0 Then 'do not use for floorstanding speakers
                    'move downward to find a fit
                    If MPnt(15) = 99 Then MPnt(0) = ModInitX
                    OutArr = FindWallIndex("HeightLD", CSng(MPnt(0)), CSng(MainLP(1)), CSng(ModInitZ), WallArr, 100, BoxHeight)
                    MPnt(0) = OutArr(0) '
                    MPnt(1) = OutArr(1) '
                    MPnt(2) = OutArr(2) '
                    MPnt(15) = OutArr(3) '
                    'if location is too low
                    If MPnt(2) < MainLP(2) + 0.5 Or MPnt(15) = 99 Then
                        If MPnt(15) = 99 Then MPnt(0) = ModInitX
                        OutArr = FindWallIndex("HeightLU", CSng(MPnt(0)), CSng(MainLP(1)), CSng(ModInitZ), WallArr, 100, BoxHeight)
                        MPnt(0) = OutArr(0) '
                        MPnt(1) = OutArr(1) '
                        MPnt(2) = OutArr(2) '
                        MPnt(15) = OutArr(3) '
                    End If
                Else
                    'Determine BaseHeight; height above floor for floorstanding speakers
                    OutArr = FindWallIndex("BaseHeightR", CSng(MPnt(0)), CSng(MPnt(1) + 0.4), CSng(FloorHeight), WallArr, 120, BoxHeight)
                    BaseHeight(SpeakerCnt) = OutArr(2) '
                    FloorHeight = BaseHeight(SpeakerCnt) + BoxHeight / 2
                    MPnt(2) = FloorHeight
                    IntResult(2, 0) = FloorHeight
                    IntResult(2, 1) = FloorHeight
                End If

                MPnt(3) = IntResult(0, 0) 'x top left point 1
                MPnt(4) = IntResult(1, 0) 'y
                MPnt(5) = IntResult(2, 0) - BoxHeight / 2  'z

                MPnt(6) = IntResult(0, 1) 'x top right point 2
                MPnt(7) = IntResult(1, 1) 'y
                MPnt(8) = IntResult(2, 1) - BoxHeight / 2   'z

                MPnt(9) = IntResult(0, 1) 'x bottom right point 3
                MPnt(10) = IntResult(1, 1) 'y
                MPnt(11) = IntResult(2, 1) + BoxHeight / 2  'z

                MPnt(12) = IntResult(0, 0) 'x bottom left point 4
                MPnt(13) = IntResult(1, 0) 'y
                MPnt(14) = IntResult(2, 0) + BoxHeight / 2  'z

            Case "Ltr"
                MPnt(0) = IntResult(0, 2)
                MPnt(1) = IntResult(1, 2)
                MPnt(2) = IntResult(2, 2)
                ModInitX = IntResult(0, 2)
                ModInitY = IntResult(1, 2)
                IncrementX = Math.Abs(MPnt(0) - MainLP(0)) / Math.Abs(MPnt(1) - MainLP(1)) * 0.1
                IncrementY = Math.Abs(MPnt(1) - MainLP(1)) / Math.Abs(MPnt(0) - MainLP(0)) * 0.1
                FactorY = -1
                FactorX = -1
                'move position backward until a fit
                OutArr = FindWallIndex("CeilingXB", CSng(ModInitX), CSng(ModInitY), CSng(MainLP(2)), WallArr, 100, BoxHeight)
                MPnt(0) = OutArr(0) '
                MPnt(1) = OutArr(1) '
                MPnt(2) = OutArr(2) '
                MPnt(15) = OutArr(3) '
                'Check if too far back if so move forward
                If MPnt(0) > IntResult(0, 0) Or MPnt(15) = 99 Then
                    OutArr = FindWallIndex("CeilingXF", CSng(ModInitX), CSng(ModInitY), CSng(MainLP(2)), WallArr, 100, BoxHeight)
                    MPnt(0) = OutArr(0) '
                    MPnt(1) = OutArr(1) '
                    MPnt(2) = OutArr(2) '
                    MPnt(15) = OutArr(3) '
                End If

                'check width position
                'Incrementally move diagonally toward MLP until spacing is correct
                If MPnt(15) = 99 Or MainLP(0) + 2 > MPnt(0) Then 'if previous results are incalcuable  or spkr is too far forward start over
                    MPnt(0) = ModInitX
                    MPnt(1) = ModInitY
                End If
                Cnt = 1
                PM15 = 99
                Do Until PM15 <> 99 Or Cnt = 300
                    OutArr = FindWallIndex("CeilingYL", CSng(MPnt(0)), CSng(MPnt(1)), CSng(MainLP(2)), WallArr, 1, BoxWidth)
                    YMPnt(0) = OutArr(0) '
                    YMPnt(1) = OutArr(1) '
                    YMPnt(2) = OutArr(2) '
                    YMPnt(15) = OutArr(3) '
                    If YMPnt(15) = 99 Then YMPnt(0) = MPnt(0)
                    If YMPnt(15) = 99 Then YMPnt(1) = ModInitY + (CSng(Cnt) * FactorY * IncrementY) 'left=1 right=-1
                    OutArr = FindWallIndex("CeilingXF", CSng(YMPnt(0)), CSng(YMPnt(1)), CSng(MainLP(2)), WallArr, 1, BoxHeight)
                    MPnt(0) = OutArr(0) '
                    MPnt(1) = OutArr(1) '
                    MPnt(2) = OutArr(2) '
                    MPnt(15) = OutArr(3) '
                    If MPnt(15) = 99 Then
                        MPnt(0) = ModInitX + (CSng(Cnt) * FactorX * IncrementX) 'front=1 rear=-1
                        MPnt(1) = YMPnt(1)
                    End If
                    Cnt = Cnt + 1
                    If MPnt(15) <> 99 And YMPnt(15) <> 99 Then PM15 = 0
                Loop

                'if too low
                ModInitY = MPnt(1)
                Cnt = 1
                Do Until MPnt(2) > MainLP(2) + 3 Or Cnt = 100 'Do Until MPnt(2) > MainLP(2) + 3 Or Cnt = 100
                    OutArr = FindWallIndex("CeilingYL", CSng(MPnt(0)), CSng(MPnt(1)), CSng(MainLP(2) + 2), WallArr, 1, BoxWidth)
                    MPnt(0) = OutArr(0) '
                    MPnt(1) = OutArr(1) '
                    MPnt(2) = OutArr(2) '
                    MPnt(15) = OutArr(3) '
                    MPnt(1) = ModInitY + (CSng(Cnt) * FactorY * IncrementY) 'left=1 right=-1
                    Cnt = Cnt + 1
                Loop

                ModInitX = MPnt(0)
                Cnt = 1
                Do Until MPnt(2) > MainLP(2) + 3 Or Cnt = 100
                    OutArr = FindWallIndex("CeilingXF", CSng(MPnt(0)), CSng(MPnt(1)), CSng(MainLP(2) + 2), WallArr, 1, BoxHeight)
                    MPnt(0) = OutArr(0) '
                    MPnt(1) = OutArr(1) '
                    MPnt(2) = OutArr(2) '
                    MPnt(15) = OutArr(3) '
                    MPnt(1) = ModInitX + (CSng(Cnt) * FactorX * IncrementX) 'left=1 right=-1
                    Cnt = Cnt + 1
                Loop

                If Cnt = 100 Then MPnt(15) = 99

                BackwardLength = MPnt(0)

                TopLeft = IntResult(1, 2) + 1 'TopWidth(0, 0)
                TopRight = IntResult(1, 2) - 1 'TopWidth(1, 0)

                MPnt(3) = IntResult(0, 0) 'x top left point 1
                MPnt(4) = TopLeft 'y
                MPnt(5) = OutArr(2) 'IntResult(1, 0) 'IntResult(2, 0) 'z

                MPnt(6) = IntResult(0, 0) 'x top right point 2
                MPnt(7) = TopRight 'y
                MPnt(8) = OutArr(2) 'IntResult(1, 0) 'IntResult(2, 0) 'z

                MPnt(9) = IntResult(0, 1)  'x bottom right point 3
                MPnt(10) = TopRight 'y
                MPnt(11) = OutArr(2) 'IntResult(1, 1) 'IntResult(2, 1) 'z

                MPnt(12) = IntResult(0, 1) 'x bottom left point 4
                MPnt(13) = TopLeft 'y
                MPnt(14) = OutArr(2) 'z'IntResult(1, 1) 'IntResu

            Case "Ltf"
                MPnt(0) = IntResult(0, 2)
                MPnt(1) = IntResult(1, 2)
                MPnt(2) = IntResult(2, 2)
                ModInitX = IntResult(0, 2)
                ModInitY = IntResult(1, 2)
                IncrementX = Math.Abs(MPnt(0) - MainLP(0)) / Math.Abs(MPnt(1) - MainLP(1)) * 0.1
                IncrementY = Math.Abs(MPnt(1) - MainLP(1)) / Math.Abs(MPnt(0) - MainLP(0)) * 0.1
                FactorY = -1 'movement is toward right
                FactorX = 1 'movement is tward back
                'move position forward until a fit
                OutArr = FindWallIndex("CeilingXF", CSng(ModInitX), CSng(ModInitY), CSng(MainLP(2)), WallArr, 100, BoxHeight)
                MPnt(0) = OutArr(0) '
                MPnt(1) = OutArr(1) '
                MPnt(2) = OutArr(2) '
                MPnt(15) = OutArr(3) '
                'Check if too far forward if so move back
                If MPnt(0) < IntResult(0, 1) Or MPnt(15) = 99 Then
                    OutArr = FindWallIndex("CeilingXB", CSng(ModInitX), CSng(ModInitY), CSng(MainLP(2)), WallArr, 100, BoxHeight)
                    MPnt(0) = OutArr(0) '
                    MPnt(1) = OutArr(1) '
                    MPnt(2) = OutArr(2) '
                    MPnt(15) = OutArr(3) '
                End If

                'check width position
                'Incrementally move diagonally toward MLP until spacing is correct
                If MPnt(15) = 99 Or MainLP(0) - 2 < MPnt(0) Then 'if previous results are incalcuable start over
                    MPnt(0) = ModInitX
                    MPnt(1) = ModInitY
                End If
                Cnt = 1
                PM15 = 99
                Do Until PM15 <> 99 Or Cnt = 300
                    OutArr = FindWallIndex("CeilingYL", CSng(MPnt(0)), CSng(MPnt(1)), CSng(MainLP(2)), WallArr, 1, BoxWidth)
                    YMPnt(0) = OutArr(0) '
                    YMPnt(1) = OutArr(1) '
                    YMPnt(2) = OutArr(2) '
                    YMPnt(15) = OutArr(3) '
                    If YMPnt(15) = 99 Then YMPnt(0) = MPnt(0)
                    If YMPnt(15) = 99 Then YMPnt(1) = ModInitY + (CSng(Cnt) * FactorY * IncrementY) 'left=1 right=-1
                    OutArr = FindWallIndex("CeilingXB", CSng(YMPnt(0)), CSng(YMPnt(1)), CSng(MainLP(2)), WallArr, 1, BoxHeight)
                    MPnt(0) = OutArr(0) '
                    MPnt(1) = OutArr(1) '
                    MPnt(2) = OutArr(2) '
                    MPnt(15) = OutArr(3) '
                    If MPnt(15) = 99 Then
                        MPnt(0) = ModInitX + (CSng(Cnt) * FactorX * IncrementX) 'front=1 rear=-1
                        MPnt(1) = YMPnt(1)
                    End If
                    Cnt = Cnt + 1
                    If MPnt(15) <> 99 And YMPnt(15) <> 99 Then PM15 = 0
                Loop

                'if too low
                ModInitY = MPnt(1)
                Cnt = 1
                Do Until MPnt(2) > MainLP(2) + 3 Or Cnt = 100
                    OutArr = FindWallIndex("CeilingYL", CSng(MPnt(0)), CSng(MPnt(1)), CSng(MainLP(2) + 2), WallArr, 1, BoxWidth)
                    MPnt(0) = OutArr(0) '
                    MPnt(1) = OutArr(1) '
                    MPnt(2) = OutArr(2) '
                    MPnt(15) = OutArr(3) '
                    MPnt(1) = ModInitY + (CSng(Cnt) * FactorY * IncrementY) 'left=1 right=-1
                    Cnt = Cnt + 1
                    If MPnt(15) = 99 Then Cnt = 100
                Loop

                If Cnt = 100 Then MPnt(15) = 99
                ForwardLength = MPnt(0)

                TopLeft = IntResult(1, 2) + 1 'TopWidth(0, 0)
                TopRight = IntResult(1, 2) - 1 'TopWidth(1, 0)

                MPnt(3) = IntResult(0, 0) 'x top left point 1
                MPnt(4) = TopLeft 'y
                MPnt(5) = OutArr(2) 'IntResult(1, 0) 'IntResult(2, 0) 'z

                MPnt(6) = IntResult(0, 0) 'x top right point 2
                MPnt(7) = TopRight 'y
                MPnt(8) = OutArr(2) 'IntResult(1, 0) 'IntResult(2, 0) 'z

                MPnt(9) = IntResult(0, 1)  'x bottom right point 3
                MPnt(10) = TopRight 'y
                MPnt(11) = OutArr(2) 'IntResult(1, 1) 'IntResult(2, 1) 'z

                MPnt(12) = IntResult(0, 1) 'x bottom left point 4
                MPnt(13) = TopLeft 'y
                MPnt(14) = OutArr(2) 'z'IntResult(1, 1) 'IntResu

            Case "Rtr"
                MPnt(0) = IntResult(0, 2)
                MPnt(1) = IntResult(1, 2)
                MPnt(2) = IntResult(2, 2)
                ModInitX = BackwardLength 'IntResult(0, 2)
                ModInitY = IntResult(1, 2)
                IncrementX = Math.Abs(MPnt(0) - MainLP(0)) / Math.Abs(MPnt(1) - MainLP(1)) * 0.1
                IncrementY = Math.Abs(MPnt(1) - MainLP(1)) / Math.Abs(MPnt(0) - MainLP(0)) * 0.1
                FactorY = 1
                FactorX = -1
                'move position backward until a fit
                OutArr = FindWallIndex("CeilingXB", CSng(ModInitX), CSng(ModInitY), CSng(MainLP(2)), WallArr, 100, BoxHeight)
                MPnt(0) = OutArr(0) '
                MPnt(1) = OutArr(1) '
                MPnt(2) = OutArr(2) '
                MPnt(15) = OutArr(3) '
                'Check if too far back if so move forward
                If MPnt(0) > IntResult(0, 0) Or MPnt(15) = 99 Then
                    OutArr = FindWallIndex("CeilingXF", CSng(ModInitX), CSng(ModInitY), CSng(MainLP(2)), WallArr, 100, BoxHeight)
                    MPnt(0) = OutArr(0) '
                    MPnt(1) = OutArr(1) '
                    MPnt(2) = OutArr(2) '
                    MPnt(15) = OutArr(3) '
                End If

                'check width position
                'Incrementally move diagonally toward MLP until spacing is correct
                If MPnt(15) = 99 Or MainLP(0) + 2 > MPnt(0) Then 'if previous results are incalcuable or spkr too far forward start over
                    ModInitX = IntResult(0, 2)
                    MPnt(0) = ModInitX
                    MPnt(1) = ModInitY
                End If
                Cnt = 1
                PM15 = 99
                Do Until PM15 <> 99 Or Cnt = 300
                    OutArr = FindWallIndex("CeilingYR", CSng(MPnt(0)), CSng(MPnt(1)), CSng(MainLP(2)), WallArr, 1, BoxWidth)
                    YMPnt(0) = OutArr(0) '
                    YMPnt(1) = OutArr(1) '
                    YMPnt(2) = OutArr(2) '
                    YMPnt(15) = OutArr(3) '
                    If YMPnt(15) = 99 Then YMPnt(0) = MPnt(0)
                    If YMPnt(15) = 99 Then YMPnt(1) = ModInitY + (CSng(Cnt) * FactorY * IncrementY) 'left=1 right=-1
                    OutArr = FindWallIndex("CeilingXF", CSng(YMPnt(0)), CSng(YMPnt(1)), CSng(MainLP(2)), WallArr, 1, BoxHeight)
                    MPnt(0) = OutArr(0) '
                    MPnt(1) = OutArr(1) '
                    MPnt(2) = OutArr(2) '
                    MPnt(15) = OutArr(3) '
                    If MPnt(15) = 99 Then
                        MPnt(0) = ModInitX + (CSng(Cnt) * FactorX * IncrementX) 'front=1 rear=-1
                        MPnt(1) = YMPnt(1)
                    End If
                    Cnt = Cnt + 1
                    If MPnt(15) <> 99 And YMPnt(15) <> 99 Then PM15 = 0
                Loop

                'if too low
                ModInitY = MPnt(1)
                Cnt = 1
                Do Until MPnt(2) > MainLP(2) + 3 Or Cnt = 100
                    OutArr = FindWallIndex("CeilingYR", CSng(MPnt(0)), CSng(MPnt(1)), CSng(MainLP(2) + 2), WallArr, 1, BoxWidth)
                    MPnt(0) = OutArr(0) '
                    MPnt(1) = OutArr(1) '
                    MPnt(2) = OutArr(2) '
                    MPnt(15) = OutArr(3) '
                    MPnt(1) = ModInitY + (CSng(Cnt) * FactorY * IncrementY) 'left=1 right=-1
                    Cnt = Cnt + 1
                Loop

                If Cnt = 100 Then MPnt(15) = 99

                TopLeft = IntResult(1, 2) + 1 'TopWidth(0, 0)
                TopRight = IntResult(1, 2) - 1 'TopWidth(1, 0)

                MPnt(3) = IntResult(0, 0) 'x top left point 1
                MPnt(4) = TopLeft 'y
                MPnt(5) = OutArr(2) 'IntResult(1, 0) 'IntResult(2, 0) 'z

                MPnt(6) = IntResult(0, 0) 'x top right point 2
                MPnt(7) = TopRight 'y
                MPnt(8) = OutArr(2) 'IntResult(1, 0) 'IntResult(2, 0) 'z

                MPnt(9) = IntResult(0, 1)  'x bottom right point 3
                MPnt(10) = TopRight 'y
                MPnt(11) = OutArr(2) 'IntResult(1, 1) 'IntResult(2, 1) 'z

                MPnt(12) = IntResult(0, 1) 'x bottom left point 4
                MPnt(13) = TopLeft 'y
                MPnt(14) = OutArr(2) 'z'IntResult(1, 1) 'IntResu

            Case "Rtf"
                MPnt(0) = ForwardLength 'IntResult(0, 2)
                MPnt(1) = IntResult(1, 2)
                MPnt(2) = IntResult(2, 2)
                ModInitX = ForwardLength 'IntResult(0, 2)
                ModInitY = IntResult(1, 2)
                IncrementX = Math.Abs(MPnt(0) - MainLP(0)) / Math.Abs(MPnt(1) - MainLP(1)) * 0.1
                IncrementY = Math.Abs(MPnt(1) - MainLP(1)) / Math.Abs(MPnt(0) - MainLP(0)) * 0.1
                FactorY = 1 'movement is toward left
                FactorX = 1 'movement is toward back
                'move position forward until a fit
                OutArr = FindWallIndex("CeilingXF", CSng(ModInitX), CSng(ModInitY), CSng(MainLP(2)), WallArr, 100, BoxHeight)
                MPnt(0) = OutArr(0) '
                MPnt(1) = OutArr(1) '
                MPnt(2) = OutArr(2) '
                MPnt(15) = OutArr(3) '
                'Check if too far forward if so move back
                If MPnt(0) > IntResult(0, 1) Or MPnt(15) = 99 Then
                    OutArr = FindWallIndex("CeilingXB", CSng(ModInitX), CSng(ModInitY), CSng(MainLP(2)), WallArr, 100, BoxHeight)
                    MPnt(0) = OutArr(0) '
                    MPnt(1) = OutArr(1) '
                    MPnt(2) = OutArr(2) '
                    MPnt(15) = OutArr(3) '
                End If

                'check width position
                'Incrementally move diagonally toward MLP until spacing is correct
                If MPnt(15) = 99 Or MainLP(0) - 2 < MPnt(0) Then 'if previous results are incalcuable start over or position is too far backward
                    MPnt(0) = ModInitX
                    MPnt(1) = ModInitY
                End If
                Cnt = 1
                PM15 = 99
                Do Until PM15 <> 99 Or Cnt = 300
                    OutArr = FindWallIndex("CeilingYR", CSng(MPnt(0)), CSng(MPnt(1)), CSng(MainLP(2)), WallArr, 1, BoxWidth)
                    YMPnt(0) = OutArr(0) '
                    YMPnt(1) = OutArr(1) '
                    YMPnt(2) = OutArr(2) '
                    YMPnt(15) = OutArr(3) '
                    If YMPnt(15) = 99 Then YMPnt(0) = MPnt(0)
                    If YMPnt(15) = 99 Then YMPnt(1) = ModInitY + (CSng(Cnt) * FactorY * IncrementY) 'left=1 right=-1
                    OutArr = FindWallIndex("CeilingXB", CSng(YMPnt(0)), CSng(YMPnt(1)), CSng(MainLP(2)), WallArr, 1, BoxHeight)
                    MPnt(0) = OutArr(0) '
                    MPnt(1) = OutArr(1) '
                    MPnt(2) = OutArr(2) '
                    MPnt(15) = OutArr(3) '
                    If MPnt(15) = 99 Then
                        MPnt(0) = ModInitX + (CSng(Cnt) * FactorX * IncrementX) 'front=1 rear=-1
                        MPnt(1) = YMPnt(1)
                    End If
                    Cnt = Cnt + 1
                    If MPnt(15) <> 99 And YMPnt(15) <> 99 Then PM15 = 0
                Loop

                'if too low
                ModInitY = MPnt(1)
                Cnt = 1
                Do Until MPnt(2) > MainLP(2) + 3 Or Cnt = 100
                    OutArr = FindWallIndex("CeilingYR", CSng(MPnt(0)), CSng(MPnt(1)), CSng(MainLP(2) + 2), WallArr, 1, BoxWidth)
                    MPnt(0) = OutArr(0) '
                    MPnt(1) = OutArr(1) '
                    MPnt(2) = OutArr(2) '
                    MPnt(15) = OutArr(3) '
                    MPnt(1) = ModInitY + (CSng(Cnt) * FactorY * IncrementY) 'left=1 right=-1
                    Cnt = Cnt + 1
                Loop

                If Cnt = 100 Then MPnt(15) = 99

                TopLeft = IntResult(1, 2) + 1 'TopWidth(0, 0)
                TopRight = IntResult(1, 2) - 1 'TopWidth(1, 0)

                MPnt(3) = IntResult(0, 0) 'x top left point 1
                MPnt(4) = TopLeft 'y
                MPnt(5) = OutArr(2) 'IntResult(1, 0) 'IntResult(2, 0) 'z

                MPnt(6) = IntResult(0, 0) 'x top right point 2
                MPnt(7) = TopRight 'y
                MPnt(8) = OutArr(2) 'IntResult(1, 0) 'IntResult(2, 0) 'z

                MPnt(9) = IntResult(0, 1)  'x bottom right point 3
                MPnt(10) = TopRight 'y
                MPnt(11) = OutArr(2) 'IntResult(1, 1) 'IntResult(2, 1) 'z

                MPnt(12) = IntResult(0, 1) 'x bottom left point 4
                MPnt(13) = TopLeft 'y
                MPnt(14) = OutArr(2) 'z'IntResult(1, 1) 'IntResu

            Case Else

                MsgBox("Speaker designation not recognized", vbOKOnly, "Error")
        End Select

        If MPnt(15) = 99 Then MsgBox(SpkrConfig & " does not have a calculable placement", vbOKOnly, "Speaker Placement not possible")

        CalcMirrorPnts = MPnt
    End Function

    Private Function FindWallIndex(ByRef WallName As String, ByRef PosX As Single, ByRef PosY As Single,
    ByRef PosZ As Single, ByRef WallArray As Object, ByRef LimitVar As Integer, ByRef ToleranceIn As Single) As Object 'Find the wall that contains the point
        Dim n As Integer, TempArr As Object, P1(2) As Single, P2(2) As Single, ResultArr As Object, Cnt As Integer
        Dim Corn1(2) As Single, Corn2(2) As Single, Corn3(2) As Single, Corn4(2) As Single, ClickCnt As Integer
        Dim JumpP(2) As Single, z As Integer, P1T(2) As Single, P2T(2) As Single, StopCnt As Integer, CompareVar(2, 1) As Int16
        Dim GoVar As Boolean, CCnt As Integer, OriVar As Integer, ResultArr0 As Object, ResultArr1 As Object, IsALine As Single, m As Integer, Tolerance As Single

        ReDim ResultArr(3)
        ReDim ResultArr0(3)
        ReDim ResultArr1(3)

        Tolerance = ToleranceIn / 2 + 0.2

        Select Case WallName
            Case "BaseHeightF"
                P1(0) = PosX
                P1(1) = PosY
                P1(2) = PosZ
                P2(0) = PosX
                P2(1) = PosY
                P2(2) = PosZ - 50 'c
                JumpP(0) = 0.2 'c
                JumpP(1) = 0
                JumpP(2) = 0 'c
                OriVar = 2
            Case "BaseHeightB"
                P1(0) = PosX
                P1(1) = PosY
                P1(2) = PosZ
                P2(0) = PosX
                P2(1) = PosY
                P2(2) = PosZ - 50 'c
                JumpP(0) = -0.2 'c
                JumpP(1) = 0
                JumpP(2) = 0 'c
                OriVar = 2
            Case "BaseHeightR"
                P1(0) = PosX
                P1(1) = PosY
                P1(2) = PosZ
                P2(0) = PosX
                P2(1) = PosY
                P2(2) = PosZ - 50 'c
                JumpP(0) = 0
                JumpP(1) = 0.2 'c
                JumpP(2) = 0 'c
                OriVar = 2
            Case "BaseHeightL"
                P1(0) = PosX
                P1(1) = PosY
                P1(2) = PosZ
                P2(0) = PosX
                P2(1) = PosY
                P2(2) = PosZ - 50 'c
                JumpP(0) = 0
                JumpP(1) = -0.2 'c
                JumpP(2) = 0 'c
                OriVar = 2
            Case "HeightRU" 'height right upward
                P1(0) = PosX
                P1(1) = PosY - 2 'c
                P1(2) = PosZ
                P2(0) = PosX
                P2(1) = PosY - 50 'c
                P2(2) = PosZ
                JumpP(0) = 0
                JumpP(1) = 0
                JumpP(2) = 1 * Tolerance 'c
                OriVar = 2
            Case "HeightRD"
                P1(0) = PosX
                P1(1) = PosY - 2 'c
                P1(2) = PosZ
                P2(0) = PosX
                P2(1) = PosY - 50 'c
                P2(2) = PosZ
                JumpP(0) = 0
                JumpP(1) = 0
                JumpP(2) = -1 * Tolerance 'c
                OriVar = 2
            Case "HeightLU"
                P1(0) = PosX
                P1(1) = PosY + 2 'c
                P1(2) = PosZ
                P2(0) = PosX
                P2(1) = PosY + 50 'c
                P2(2) = PosZ
                JumpP(0) = 0
                JumpP(1) = 0
                JumpP(2) = 1 * Tolerance 'c
                OriVar = 2
            Case "HeightLD"
                P1(0) = PosX
                P1(1) = PosY + 2 'c
                P1(2) = PosZ
                P2(0) = PosX
                P2(1) = PosY + 50 'c
                P2(2) = PosZ
                JumpP(0) = 0
                JumpP(1) = 0
                JumpP(2) = -1 * Tolerance 'c
                OriVar = 2
            Case "HeightFU"
                P1(0) = PosX - 2
                P1(1) = PosY
                P1(2) = PosZ
                P2(0) = PosX - 50
                P2(1) = PosY
                P2(2) = PosZ
                JumpP(0) = 0
                JumpP(1) = 0
                JumpP(2) = 1 * Tolerance 'c
                OriVar = 2
            Case "HeightFD"
                P1(0) = PosX - 2
                P1(1) = PosY
                P1(2) = PosZ
                P2(0) = PosX - 50
                P2(1) = PosY
                P2(2) = PosZ
                JumpP(0) = 0
                JumpP(1) = 0
                JumpP(2) = -1 * Tolerance 'c
                OriVar = 2
            Case "HeightRrU"
                P1(0) = PosX + 2
                P1(1) = PosY
                P1(2) = PosZ
                P2(0) = PosX + 50
                P2(1) = PosY
                P2(2) = PosZ
                JumpP(0) = 0
                JumpP(1) = 0
                JumpP(2) = 1 * Tolerance 'c
                OriVar = 2
            Case "HeightRrD"
                P1(0) = PosX + 2
                P1(1) = PosY
                P1(2) = PosZ
                P2(0) = PosX + 50
                P2(1) = PosY
                P2(2) = PosZ
                JumpP(0) = 0
                JumpP(1) = 0
                JumpP(2) = -1 * Tolerance 'c
                OriVar = 2
            Case "Left"
                P1(0) = PosX
                P1(1) = PosY + 2 'c
                P1(2) = PosZ
                P2(0) = PosX
                P2(1) = PosY + 50 'c
                P2(2) = PosZ
                JumpP(0) = -1 * Tolerance 'c
                JumpP(1) = 0
                JumpP(2) = 0
                OriVar = 1
            Case "LeftF"
                P1(0) = PosX
                P1(1) = PosY + 2 'c
                P1(2) = PosZ
                P2(0) = PosX
                P2(1) = PosY + 50 'c
                P2(2) = PosZ
                JumpP(0) = -1 * Tolerance
                JumpP(1) = 0
                JumpP(2) = 0
                OriVar = 1
            Case "Right"
                P1(0) = PosX
                P1(1) = PosY - 2 'c
                P1(2) = PosZ
                P2(0) = PosX
                P2(1) = PosY - 50 'c
                P2(2) = PosZ
                JumpP(0) = 1 * Tolerance
                JumpP(1) = 0
                JumpP(2) = 0
                OriVar = 1
            Case "RightF"
                P1(0) = PosX
                P1(1) = PosY - 2 'c
                P1(2) = PosZ
                P2(0) = PosX
                P2(1) = PosY - 50 'c
                P2(2) = PosZ
                JumpP(0) = -1 * Tolerance
                JumpP(1) = 0
                JumpP(2) = 0
                OriVar = 1
            Case "CeilingYR"
                P1(0) = PosX
                P1(1) = PosY
                P1(2) = PosZ + 50 'c
                P2(0) = PosX
                P2(1) = PosY
                P2(2) = PosZ - 1 'c
                JumpP(0) = 0
                JumpP(1) = 1 * Tolerance 'c
                JumpP(2) = 0
                OriVar = 2
            Case "CeilingYL"
                P1(0) = PosX
                P1(1) = PosY
                P1(2) = PosZ + 50 'c
                P2(0) = PosX
                P2(1) = PosY
                P2(2) = PosZ - 1 'c
                JumpP(0) = 0
                JumpP(1) = -1 * Tolerance 'c
                JumpP(2) = 0
                OriVar = 2
            Case "CeilingXF"
                P1(0) = PosX
                P1(1) = PosY
                P1(2) = PosZ + 50 'c
                P2(0) = PosX
                P2(1) = PosY
                P2(2) = PosZ - 1 'c
                JumpP(0) = -1 * Tolerance
                JumpP(1) = 0
                JumpP(2) = 0
                OriVar = 2
            Case "CeilingXB"
                P1(0) = PosX
                P1(1) = PosY
                P1(2) = PosZ + 50 'c
                P2(0) = PosX
                P2(1) = PosY
                P2(2) = PosZ - 1 'c
                JumpP(0) = 1 * Tolerance
                JumpP(1) = 0
                JumpP(2) = 0
                OriVar = 2
            Case "FrontR"
                P1(0) = PosX - 50
                P1(1) = PosY
                P1(2) = PosZ
                P2(0) = PosX - 2
                P2(1) = PosY
                P2(2) = PosZ
                JumpP(0) = 0
                JumpP(1) = 1 * Tolerance 'c
                JumpP(2) = 0
                OriVar = 0
            Case "FrontL"
                P1(0) = PosX - 50
                P1(1) = PosY
                P1(2) = PosZ
                P2(0) = PosX - 2
                P2(1) = PosY
                P2(2) = PosZ
                JumpP(0) = 0
                JumpP(1) = -1 * Tolerance 'c
                JumpP(2) = 0
                OriVar = 0
            Case "RearR"
                P1(0) = PosX + 2
                P1(1) = PosY
                P1(2) = PosZ
                P2(0) = PosX + 50
                P2(1) = PosY
                P2(2) = PosZ
                JumpP(0) = 0
                JumpP(1) = 1 * Tolerance 'c
                JumpP(2) = 0
                OriVar = 0
            Case "RearL"
                P1(0) = PosX + 2
                P1(1) = PosY
                P1(2) = PosZ
                P2(0) = PosX + 50
                P2(1) = PosY
                P2(2) = PosZ
                JumpP(0) = 0
                JumpP(1) = -1 * Tolerance 'c
                JumpP(2) = 0
                OriVar = 0
        End Select

        P1T(0) = P1(0)
        P1T(1) = P1(1)
        P1T(2) = P1(2)
        P2T(0) = P2(0)
        P2T(1) = P2(1)
        P2T(2) = P2(2)

        StopCnt = 0
        ClickCnt = 0

        Do Until ClickCnt = LimitVar
            If ClickCnt > 0 Then
                P1T(0) = P1T(0) + JumpP(0) * 0.1 'move the speaker farther to find next surface
                P1T(1) = P1T(1) + JumpP(1) * 0.1
                P1T(2) = P1T(2) + JumpP(2) * 0.1
                P2T(0) = P2T(0) + JumpP(0) * 0.1
                P2T(1) = P2T(1) + JumpP(1) * 0.1
                P2T(2) = P2T(2) + JumpP(2) * 0.1

                P1(0) = P1T(0)
                P1(1) = P1T(1)
                P1(2) = P1T(2)
                P2(0) = P2T(0)
                P2(1) = P2T(1)
                P2(2) = P2T(2)
            End If
            ClickCnt = ClickCnt + 1
            Cnt = 0

            For z = 0 To 2
                If Cnt = 1 Then
                    P1(0) = P1(0) - JumpP(0) 'test to see if there is room for the speaker
                    P1(1) = P1(1) - JumpP(1)
                    P1(2) = P1(2) - JumpP(2)
                    P2(0) = P2(0) - JumpP(0)
                    P2(1) = P2(1) - JumpP(1)
                    P2(2) = P2(2) - JumpP(2)
                ElseIf Cnt = 2 Then
                    P1(0) = P1(0) + JumpP(0) * 2 'test to see if there is room for the speaker
                    P1(1) = P1(1) + JumpP(1) * 2
                    P1(2) = P1(2) + JumpP(2) * 2
                    P2(0) = P2(0) + JumpP(0) * 2
                    P2(1) = P2(1) + JumpP(1) * 2
                    P2(2) = P2(2) + JumpP(2) * 2
                End If
                For n = 0 To UBound(WallArray, 2)
                    If WallArray(4, n) <> "00000" Then 'if the wall is real and not one of the imaginary ones
                        Corn1(0) = WallArray(5, n)
                        Corn1(1) = WallArray(6, n)
                        Corn1(2) = WallArray(7, n)
                        Corn2(0) = WallArray(8, n)
                        Corn2(1) = WallArray(9, n)
                        Corn2(2) = WallArray(10, n)
                        Corn3(0) = WallArray(11, n)
                        Corn3(1) = WallArray(12, n)
                        Corn3(2) = WallArray(13, n)
                        Corn4(0) = WallArray(14, n)
                        Corn4(1) = WallArray(15, n)
                        Corn4(2) = WallArray(16, n)
                        On Error Resume Next
                        TempArr = Functions.Intersection(P1, P2, Corn1, Corn2, Corn3, Corn4)

                        If TempArr(0) <> 0 And TempArr(1) <> 0 And TempArr(2) <> 0 Then
                            GoVar = True

                            If Cnt < 2 Then
                                GoVar = True
                            Else
                                GoVar = False
                            End If

                            If Cnt = 2 Then
                                IsALine = Functions.TriangleArea(ResultArr, ResultArr1, TempArr) ' if area of the triangle is 0 it's a line
                                If Math.Abs(IsALine) <= 0.05 Then
                                    GoVar = True
                                    'Compare the three points to see if any match.  If a match is found its a bad hit
                                    ReDim CompareVar(2, 2)
                                    For m = 0 To 2
                                        If Math.Abs(ResultArr(m) - ResultArr1(m)) <= 0.05 Then CompareVar(m, 0) = 1
                                        If Math.Abs(ResultArr1(m) - TempArr(m)) <= 0.05 Then CompareVar(m, 1) = 1
                                        If Math.Abs(TempArr(m) - ResultArr(m)) <= 0.05 Then CompareVar(m, 2) = 1
                                    Next m

                                    If CompareVar(0, 0) = 1 And CompareVar(1, 0) = 1 And CompareVar(2, 0) = 1 Then GoVar = False
                                    If CompareVar(0, 1) = 1 And CompareVar(1, 1) = 1 And CompareVar(2, 1) = 1 Then GoVar = False
                                    If CompareVar(0, 2) = 1 And CompareVar(1, 2) = 1 And CompareVar(2, 2) = 1 Then GoVar = False
                                    'if the points are at the same jump point
                                    If Math.Abs(JumpP(0)) = Tolerance Then
                                        If CompareVar(0, 0) = 1 Or CompareVar(0, 1) = 1 Or CompareVar(0, 2) = 1 Then GoVar = False
                                    End If
                                    If Math.Abs(JumpP(1)) = Tolerance Then
                                        If CompareVar(1, 0) = 1 Or CompareVar(1, 1) = 1 Or CompareVar(1, 2) = 1 Then GoVar = False
                                    End If
                                    If Math.Abs(JumpP(2)) = Tolerance Then
                                        If CompareVar(2, 0) = 1 Or CompareVar(2, 1) = 1 Or CompareVar(2, 2) = 1 Then GoVar = False
                                    End If
                                End If
                            End If

                            If GoVar Then
                                If Cnt = 0 Then
                                    ResultArr(0) = TempArr(0)
                                    ResultArr(1) = TempArr(1)
                                    ResultArr(2) = TempArr(2)
                                    ResultArr(3) = WallArray(0, n)
                                ElseIf Cnt = 1 Then
                                    ResultArr1(0) = TempArr(0)
                                    ResultArr1(1) = TempArr(1)
                                    ResultArr1(2) = TempArr(2)
                                    ResultArr1(3) = WallArray(0, n)
                                End If
                                Cnt = Cnt + 1
                                If Cnt = 3 Then
                                    n = UBound(WallArray, 2)
                                    ClickCnt = LimitVar
                                    'z = 2
                                End If
                            End If
                        End If
                    End If
                Next n
            Next z
        Loop

        If Cnt < 3 Then
            ResultArr(3) = 99
        End If

        FindWallIndex = ResultArr
    End Function
    Public Function CalcSpeakerToe(ByRef MainLP As Object, ByRef SpeakerArr As Object, ByRef SpkrPick As Integer, ByRef Xer As Single, ByRef Yer As Single, ByRef Zer As Single) As Object
        Dim AngleNum As Single, P1(2) As Single, P2(2) As Single, SpkrUse As String, AngleArr As Object
        Dim Oppo As Single, Adj As Single, XAng As Single, YAng As Single, ZAng As Single

        ReDim AngleArr(2)

        P1(0) = MainLP(0)
        P1(1) = MainLP(1)
        P1(2) = MainLP(2)

        P2(0) = Xer
        P2(1) = Yer
        P2(2) = Zer


        SpkrUse = SpeakerArr(3, SpkrPick)
        'Azimuth or toein b=Atan(O/A)
        Oppo = Math.Abs(P1(1) - P2(1))
        Adj = Math.Abs(P1(0) - P2(0))
        If Adj = 0 Then Adj = 0.0001
        AngleNum = Functions.Degrees(Math.Atan(Oppo / Adj))

        If AngleNum < 0 Then
            AngleNum = 360 - AngleNum
        End If

        If P2(1) > P1(1) And P2(0) <= P1(0) Then 'quad 1 Left is in Q1
            XAng = AngleNum
            YAng = 90 + XAng
        ElseIf P2(1) <= P1(1) And P2(0) < P1(0) Then 'quad 2 Right is in Q2
            XAng = AngleNum
            YAng = 90 - XAng
        ElseIf P2(1) > P1(1) And P2(0) >= P1(0) Then 'quad 3 Ls is in Q3
            XAng = 180 - AngleNum
            YAng = AngleNum + 90
        ElseIf P2(1) < P1(1) And P2(0) >= P1(0) Then 'quad 4 Rs is in Q3
            XAng = 180 - AngleNum
            YAng = 90 - AngleNum
        End If

        'Test
        If P2(1) > P1(1) And P2(0) <= P1(0) Then 'quad 1
            XAng = 180 - AngleNum
            YAng = 90 - AngleNum
        ElseIf P2(1) <= P1(1) And P2(0) < P1(0) Then 'quad 2
            XAng = 180 - AngleNum
            YAng = 90 + AngleNum
        ElseIf P2(1) > P1(1) And P2(0) >= P1(0) Then 'quad 3
            XAng = AngleNum
            YAng = 90 - XAng
        ElseIf P2(1) < P1(1) And P2(0) >= P1(0) Then 'quad 4
            XAng = AngleNum
            YAng = 90 + XAng
        End If

        'RA or toe-up
        Oppo = Math.Abs(P1(2) - P2(2))
        Adj = Functions.PythagDist(P1(0), P2(0), P1(1), P2(1))
        AngleNum = Functions.Degrees(Math.Atan(Oppo / Adj))

        If P2(2) >= P1(2) Then
            ZAng = 90 - AngleNum
        Else
            ZAng = 90 + AngleNum
        End If

        'If ZAng >= 135 Then
        ' '"Ceiling
        ' AngleArr(0) = 0 'toe-in
        ' AngleArr(1) = 90 - XAng + 90 'toe-up
        ' AngleArr(2) = 90 + YAng  'roll
        If YAng <= 135 And YAng >= 45 Then 'Front or Rear
            If XAng < 90 Then '
                '"Front"
                AngleArr(0) = 90 + YAng 'toe-in
                AngleArr(1) = ZAng - 90 'toe-up
                AngleArr(2) = 180 'roll
            ElseIf XAng >= 135 And XAng <= 225 Then
                '"Rear"
                AngleArr(0) = 90 - YAng 'toe-in
                AngleArr(1) = 90 - ZAng 'toe-up
                AngleArr(2) = 180 'roll
            End If
        ElseIf XAng <= 135 And XAng >= 45 Then 'Left or Right
            If YAng < 90 Then '
                '"Right"
                AngleArr(0) = 90 + (90 - XAng) 'toe-in
                AngleArr(1) = 0 'toe-up
                AngleArr(2) = (180 - ZAng) + 90 'roll
            ElseIf YAng >= 135 And YAng <= 225 Then '
                '"Left"
                AngleArr(0) = 360 - (180 - XAng) 'toe-in
                AngleArr(1) = 0 'toe-up
                AngleArr(2) = ZAng + 90 'roll
            End If
        End If

        CalcSpeakerToe = AngleArr

    End Function
End Class
