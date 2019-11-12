Option Explicit On

Public Class SpeakerConfig
    Dim DolbyName As Object
    Dim ConfigBrand As Object
    Dim dbServercls As New dbServer
    Dim Functions As New Functions
    Dim DataFunctions As New DataFunctions
    Dim Library As New Library
    Public Function OrientSpkr(ByRef ProjNum As Long, ByRef SpkrLstIn As Object) As Object
        Dim SpkrLst As Object, n As Integer, Coord As Object, WallType As Object, FloorHeight As Single, WallSep As Single = 0

        'SpkrLst(0, n) = CurrentProjectID 'dbRecordsetProject!ProjectIndex = SpeakerData(0, Cnt)
        'SpkrLst(1, n) = NewProjectNumber 'dbRecordsetProject!Index = SpeakerData(1, Cnt)
        'SpkrLst(2, n) = SpkrInfoEach(1) 'dbRecordsetProject!DrawRefIndex = SpeakerData(2, Cnt)
        'SpkrLst(3, n) = SpkrInfoEach(2) 'dbRecordsetProject!Brand = SpeakerData(3, Cnt)
        'SpkrLst(4, n) = SpkrInfoEach(3) 'dbRecordsetProject!Model = SpeakerData(4, Cnt)
        'SpkrLst(5, n) = SpkrInfoEach(4) 'dbRecordsetProject!Version = SpeakerData(5, Cnt)
        'SpkrLst(6, n) = SpkrInfoEach(5) 'dbRecordsetProject!RetailPrice = SpeakerData(6, Cnt)
        'SpkrLst(8, n) = SpkrInfoEach(6) 'dbRecordsetProject!Height = SpeakerData(7, Cnt)
        'SpkrLst(9, n) = SpkrInfoEach(7) 'dbRecordsetProject!Width = SpeakerData(8, Cnt)
        'SpkrLst(10, n) = SpkrInfoEach(8) 'dbRecordsetProject!Depth = SpeakerData(9, Cnt)
        'SpkrLst(10, n) = SpkrInfoEach(9) 'dbRecordsetProject!ACOffsetX = SpeakerData(10, Cnt)
        'SpkrLst(11, n) = SpkrInfoEach(10) 'dbRecordsetProject!ACOffsetY = SpeakerData(11, Cnt)
        'SpkrLst(12, n) = SpkrInfoEach(11) 'dbRecordsetProject!ACOffsetZ = SpeakerData(12, Cnt)
        'SpkrLst(13, n) = SpkrInfoEach(12) 'dbRecordsetProject!WooferOffsetX = SpeakerData(13, Cnt)
        'SpkrLst(14, n) = SpkrInfoEach(13) 'dbRecordsetProject!WooferOffsetY = SpeakerData(14, Cnt)
        'SpkrLst(15, n) = SpkrInfoEach(14) 'dbRecordsetProject!WooferOffsetZ = SpeakerData(15, Cnt)
        'SpkrLst(16, n) = SpkrInfoEach(15) 'dbRecordsetProject!TopImage = SpeakerData(16, Cnt)
        'SpkrLst(17, n) = SpkrInfoEach(16) 'dbRecordsetProject!SideImage = SpeakerData(17, Cnt)
        'SpkrLst(18, n) = SpkrInfoEach(17) 'dbRecordsetProject!FrontImage = SpeakerData(18, Cnt)
        'SpkrLst(19, n) = SpkrInfoEach(18) 'dbRecordsetProject!BackImage = SpeakerData(19, Cnt)
        'SpkrLst(20, n) = SpkrInfoEach(19) 'dbRecordsetProject!FreqResponse = SpeakerData(20, Cnt)
        'SpkrLst(21, n) = SpkrInfoEach(20) 'dbRecordsetProject!FreqResponsePlot = SpeakerData(21, Cnt)
        'SpkrLst(22, n) = SpkrInfoEach(21) 'dbRecordsetProject!OffFreqResponsePlot = SpeakerData(22, Cnt)
        'SpkrLst(23, n) = SpkrInfoEach(22) 'dbRecordsetProject!HorCoverage = SpeakerData(23, Cnt)
        'SpkrLst(24, n) = SpkrInfoEach(23) 'dbRecordsetProject!VerCoverage = SpeakerData(24, Cnt)
        'SpkrLst(25, n) = SpkrInfoEach(24) 'dbRecordsetProject!NominalImpedance = SpeakerData(25, Cnt)
        'SpkrLst(26, n) = SpkrInfoEach(25) 'dbRecordsetProject!Sensitivity = SpeakerData(26, Cnt)
        'SpkrLst(27, n) = SpkrInfoEach(26) 'dbRecordsetProject!MaxSoundLevel = SpeakerData(27, Cnt)
        'SpkrLst(28, n) = SpkrInfoEach(27) 'dbRecordsetProject!Crossover = SpeakerData(28, Cnt)
        'SpkrLst(29, n) = SpkrInfoEach(28) 'dbRecordsetProject!Mounting = SpeakerData(29, Cnt)
        'SpkrLst(30, n) = SpkrInfoEach(29) 'dbRecordsetProject!Type = SpeakerData(30, Cnt)
        'SpkrLst(31, n) = SpkrInfoEach(30) 'dbRecordsetProject!Slope = SpeakerData(31, Cnt)
        'SpkrLst(33, n) = (SpkrConfigResults(0, 0, n)) 'X
        'SpkrLst(34, n) = (SpkrConfigResults(1, 0, n)) 'Y
        'SpkrLst(35, n) = (SpkrConfigResults(2, 0, n)) 'Z
        'SpkrLst(36, n) = SpkrInfoEach(34) 'Status = SpeakerData(35, Cnt)
        'SpkrLst(37, n) = SpkrConfigResults(0, 5, n) 'Designation = SpeakerData(36, Cnt)
        'SpkrLst(38, n) = SpkrConfigResults(2, 5, n) 'Wallindex...dbRecordsetProject!Wall = SpeakerData(37, Cnt)
        'SpkrLst(39, n) = SpkrInfoEach(37) 'Family = SpeakerData(38, Cnt)
        'SpkrLst(40, n) = 0 'ToeIn = SpeakerData(39, Cnt)
        'SpkrLst(41, n) = 0 'ToeUp = SpeakerData(40, Cnt)
        'SpkrLst(42, n) = " " 'Roll = SpeakerData(41, Cnt)

        SpkrLst = SpkrLstIn
        Dim AdjustZ As Single
        Dim AdjustX As Single
        For n = 0 To UBound(SpkrLst, 2)
            If SpkrLst(38, n) <> 99 Then
                If SpkrLst(37, n) = "Rtr" Then
                    n = n
                End If
                dbServercls.ProjectInitial = ProjNum
                dbServercls.ProjectWallIndx = SpkrLst(38, n)
                Coord = dbServercls.GetProjectCoord 'send project number and wall index
                FloorHeight = Coord(11)
                WallType = DataFunctions.DetermineWallType(Coord, CStr(SpkrLst(37, n)))
                If Mid(SpkrLst(37, n), 2, 1) = "t" Then
                    WallType(0) = "Ceiling"
                End If
                SpkrLst(40, n) = WallType(1) 'toe-in
                SpkrLst(41, n) = WallType(2) 'toe-up
                SpkrLst(42, n) = WallType(3) 'roll
                Select Case SpkrLst(30, n) 'mounting
                    Case "Inwall"
                        Select Case WallType(0) 'wall
                            Case "Front"
                                SpkrLst(33, n) = SpkrLst(33, n) - (SpkrLst(10, n) / 12) / 2 'X - Depth/2
                                AdjustZ = Math.Sin(Functions.Radians((SpkrLst(41, n)))) * ((SpkrLst(10, n) / 12) / 2)
                                SpkrLst(35, n) = SpkrLst(35, n) + AdjustZ
                                AdjustX = ((SpkrLst(10, n) / 12) / 2) - (Math.Cos(Functions.Radians((SpkrLst(41, n)))) * ((SpkrLst(10, n) / 12) / 2))
                                SpkrLst(33, n) = SpkrLst(33, n) + AdjustX
                                AdjustZ = Math.Sin(Functions.Radians((SpkrLst(40, n)))) * ((SpkrLst(10, n) / 12) / 2)
                                SpkrLst(34, n) = SpkrLst(34, n) + AdjustZ 'Y
                                AdjustX = ((SpkrLst(10, n) / 12) / 2) - (Math.Cos(Functions.Radians((SpkrLst(40, n)))) * ((SpkrLst(10, n) / 12) / 2))
                                SpkrLst(33, n) = SpkrLst(33, n) + AdjustX
                            Case "Rear"
                                SpkrLst(33, n) = SpkrLst(33, n) + (SpkrLst(10, n) / 12) / 2 'X + Depth/2
                                AdjustZ = Math.Sin(Functions.Radians((SpkrLst(41, n)))) * ((SpkrLst(10, n) / 12) / 2)
                                SpkrLst(35, n) = SpkrLst(35, n) - AdjustZ
                                AdjustX = ((SpkrLst(10, n) / 12) / 2) - (Math.Cos(Functions.Radians((SpkrLst(41, n)))) * ((SpkrLst(10, n) / 12) / 2))
                                SpkrLst(33, n) = SpkrLst(33, n) - AdjustX
                                AdjustZ = Math.Sin(Functions.Radians((SpkrLst(40, n)))) * ((SpkrLst(10, n) / 12) / 2)
                                SpkrLst(34, n) = SpkrLst(34, n) + AdjustZ 'Y
                                AdjustX = ((SpkrLst(10, n) / 12) / 2) - (Math.Cos(Functions.Radians((SpkrLst(40, n) - 180))) * ((SpkrLst(10, n) / 12) / 2))
                                SpkrLst(33, n) = SpkrLst(33, n) - AdjustX
                            Case "Left" '"SurroundL"
                                SpkrLst(34, n) = SpkrLst(34, n) - (SpkrLst(10, n) / 12) / 2 'Y - Depth/2
                                AdjustZ = Math.Sin(Functions.Radians((SpkrLst(42, n)))) * ((SpkrLst(10, n) / 12) / 2)
                                SpkrLst(35, n) = SpkrLst(35, n) - AdjustZ
                                AdjustX = ((SpkrLst(10, n) / 12) / 2) - (Math.Cos(Functions.Radians((SpkrLst(42, n)))) * ((SpkrLst(10, n) / 12) / 2))
                                SpkrLst(34, n) = SpkrLst(34, n) + AdjustX
                                AdjustZ = Math.Sin(Functions.Radians((90 - SpkrLst(40, n)))) * ((SpkrLst(10, n) / 12) / 2)
                                SpkrLst(33, n) = SpkrLst(33, n) - AdjustZ 'X
                                AdjustX = ((SpkrLst(10, n) / 12) / 2) - (Math.Cos(Functions.Radians((90 - SpkrLst(40, n)))) * ((SpkrLst(10, n) / 12) / 2))
                                SpkrLst(34, n) = SpkrLst(34, n) + AdjustX 'Y
                            Case "Right" '"SurroundR"
                                SpkrLst(34, n) = SpkrLst(34, n) + (SpkrLst(10, n) / 12) / 2 'Y + Depth/2
                                AdjustZ = Math.Sin(Functions.Radians((SpkrLst(42, n)))) * ((SpkrLst(10, n) / 12) / 2)
                                SpkrLst(35, n) = SpkrLst(35, n) + AdjustZ
                                AdjustX = ((SpkrLst(10, n) / 12) / 2) - (Math.Cos(Functions.Radians((SpkrLst(42, n)))) * ((SpkrLst(10, n) / 12) / 2))
                                SpkrLst(34, n) = SpkrLst(34, n) - AdjustX
                                AdjustZ = Math.Sin(Functions.Radians((SpkrLst(40, n) - 270))) * ((SpkrLst(10, n) / 12) / 2)
                                SpkrLst(33, n) = SpkrLst(33, n) - AdjustZ 'X
                                AdjustX = ((SpkrLst(10, n) / 12) / 2) - (Math.Cos(Functions.Radians((SpkrLst(40, n) - 270))) * ((SpkrLst(10, n) / 12) / 2))
                                SpkrLst(34, n) = SpkrLst(34, n) - AdjustX 'Y
                            Case "Ceiling"
                                SpkrLst(35, n) = SpkrLst(35, n) - (SpkrLst(10, n) / 12) / 2 'Y - Depth/2
                                AdjustX = Math.Sin(Functions.Radians((SpkrLst(42, n)))) * ((SpkrLst(10, n) / 12) / 2)
                                SpkrLst(34, n) = SpkrLst(34, n) + AdjustX 'y
                                AdjustZ = ((SpkrLst(10, n) / 12) / 2) - (Math.Cos(Functions.Radians((SpkrLst(42, n)))) * ((SpkrLst(10, n) / 12) / 2))
                                SpkrLst(35, n) = SpkrLst(35, n) + AdjustZ 'z
                                AdjustX = Math.Sin(Functions.Radians((SpkrLst(41, n) - 90))) * ((SpkrLst(10, n) / 12) / 2)
                                SpkrLst(33, n) = SpkrLst(33, n) + AdjustX 'x
                                AdjustZ = ((SpkrLst(10, n) / 12) / 2) - (Math.Cos(Functions.Radians((SpkrLst(41, n) - 90))) * ((SpkrLst(10, n) / 12) / 2))
                                SpkrLst(35, n) = SpkrLst(35, n) + AdjustZ 'z
                        End Select
                    Case "Onwall"
                        Select Case WallType(0) 'wall
                            Case "Front"
                                AdjustZ = Math.Sin(Functions.Radians((SpkrLst(41, n)))) * ((SpkrLst(10, n) / 12) / 2)
                                SpkrLst(35, n) = SpkrLst(35, n) - AdjustZ
                                AdjustX = (Math.Cos(Functions.Radians((SpkrLst(41, n)))) * ((SpkrLst(10, n) / 12) / 2))
                                SpkrLst(33, n) = SpkrLst(33, n) + AdjustX
                                AdjustZ = Math.Sin(Functions.Radians((SpkrLst(40, n)))) * ((SpkrLst(10, n) / 12) / 2)
                                SpkrLst(34, n) = SpkrLst(34, n) + AdjustZ 'Y
                                AdjustX = ((SpkrLst(10, n) / 12) / 2) - (Math.Cos(Functions.Radians((SpkrLst(40, n)))) * ((SpkrLst(10, n) / 12) / 2))
                                SpkrLst(33, n) = SpkrLst(33, n) - AdjustX 'x
                            Case "Rear"
                                AdjustZ = Math.Sin(Functions.Radians((SpkrLst(41, n)))) * ((SpkrLst(10, n) / 12) / 2)
                                SpkrLst(35, n) = SpkrLst(35, n) + AdjustZ
                                AdjustX = (Math.Cos(Functions.Radians((SpkrLst(41, n)))) * ((SpkrLst(10, n) / 12) / 2))
                                SpkrLst(33, n) = SpkrLst(33, n) - AdjustX 'x
                                AdjustZ = Math.Sin(Functions.Radians((SpkrLst(40, n)))) * ((SpkrLst(10, n) / 12) / 2)
                                SpkrLst(34, n) = SpkrLst(34, n) + AdjustZ 'Y
                                AdjustX = ((SpkrLst(10, n) / 12) / 2) - (Math.Cos(Functions.Radians((SpkrLst(40, n) - 180))) * ((SpkrLst(10, n) / 12) / 2))
                                SpkrLst(33, n) = SpkrLst(33, n) + AdjustX 'x
                            Case "Left"
                                AdjustZ = Math.Sin(Functions.Radians((SpkrLst(42, n)))) * ((SpkrLst(10, n) / 12) / 2)
                                SpkrLst(35, n) = SpkrLst(35, n) + AdjustZ 'z
                                AdjustX = (Math.Cos(Functions.Radians((SpkrLst(42, n)))) * ((SpkrLst(10, n) / 12) / 2))
                                SpkrLst(34, n) = SpkrLst(34, n) + AdjustX 'y
                                AdjustZ = Math.Sin(Functions.Radians((90 - SpkrLst(40, n)))) * ((SpkrLst(10, n) / 12) / 2)
                                SpkrLst(33, n) = SpkrLst(33, n) + AdjustZ 'X
                                AdjustX = ((SpkrLst(10, n) / 12) / 2) - (Math.Cos(Functions.Radians((90 - SpkrLst(40, n)))) * ((SpkrLst(10, n) / 12) / 2))
                                SpkrLst(34, n) = SpkrLst(34, n) + AdjustX 'Y
                            Case "Right"
                                AdjustZ = Math.Sin(Functions.Radians((SpkrLst(42, n)))) * ((SpkrLst(10, n) / 12) / 2)
                                SpkrLst(35, n) = SpkrLst(35, n) - AdjustZ 'z
                                AdjustX = (Math.Cos(Functions.Radians((SpkrLst(42, n)))) * ((SpkrLst(10, n) / 12) / 2))
                                SpkrLst(34, n) = SpkrLst(34, n) - AdjustX 'y
                                AdjustZ = Math.Sin(Functions.Radians((SpkrLst(40, n) - 270))) * ((SpkrLst(10, n) / 12) / 2)
                                SpkrLst(33, n) = SpkrLst(33, n) + AdjustZ 'X
                                AdjustX = ((SpkrLst(10, n) / 12) / 2) - (Math.Cos(Functions.Radians((SpkrLst(40, n) - 270))) * ((SpkrLst(10, n) / 12) / 2))
                                SpkrLst(34, n) = SpkrLst(34, n) - AdjustX 'Y
                            Case "Ceiling"
                                AdjustX = Math.Sin(Functions.Radians((SpkrLst(42, n)))) * ((SpkrLst(10, n) / 12) / 2)
                                SpkrLst(34, n) = SpkrLst(34, n) - AdjustX 'y
                                AdjustZ = (Math.Cos(Functions.Radians((SpkrLst(42, n)))) * ((SpkrLst(10, n) / 12) / 2))
                                SpkrLst(35, n) = SpkrLst(35, n) + AdjustZ 'z
                                AdjustX = Math.Sin(Functions.Radians((SpkrLst(41, n) - 90))) * ((SpkrLst(10, n) / 12) / 2)
                                SpkrLst(33, n) = SpkrLst(33, n) - AdjustX 'x
                                AdjustZ = ((SpkrLst(10, n) / 12) / 2) - (Math.Cos(Functions.Radians((SpkrLst(41, n) - 90))) * ((SpkrLst(10, n) / 12) / 2))
                                SpkrLst(35, n) = SpkrLst(35, n) - AdjustZ 'z
                        End Select
                    Case "Bookshelf"
                        WallSep = (0.25 / 12)
                        Select Case WallType(0) 'wall
                            Case "Front"
                                WallSep = WallSep
                                If SpkrLst(37, n) = "C" Then
                                    WallSep = WallSep * 2
                                    SpkrLst(33, n) = SpkrLst(33, n) + (SpkrLst(10, n) / 12) / 2 + WallSep  'X - Depth/2
                                    SpkrLst(40, n) = DataFunctions.SpeakerToe(0, n) 'toe-in
                                Else
                                    WallSep = WallSep * 3
                                    SpkrLst(33, n) = SpkrLst(33, n) + (SpkrLst(10, n) / 12) / 2 + WallSep 'X - Depth/2
                                    SpkrLst(40, n) = DataFunctions.SpeakerToe(0, n) 'toe-in
                                End If
                                'Shift position to maintain proper angle toward MLP
                                SpkrLst(34, n) = SpkrLst(34, n) - ((SpkrLst(10, n) / 12) / 2 + WallSep) * Math.Tan(Functions.Radians((DataFunctions.SpeakerToe(0, n))))
                                SpkrLst(41, n) = DataFunctions.SpeakerToe(1, n) 'toe-up
                                SpkrLst(42, n) = DataFunctions.SpeakerToe(2, n) 'roll
                                'compensate position for rotation
                                AdjustZ = Math.Sin(Functions.Radians((SpkrLst(41, n)))) * ((SpkrLst(10, n) / 12) / 2)
                                SpkrLst(35, n) = SpkrLst(35, n) + AdjustZ
                                AdjustX = (Math.Cos(Functions.Radians((SpkrLst(41, n)))) * ((SpkrLst(10, n) / 12) / 2))
                                SpkrLst(33, n) = SpkrLst(33, n) + AdjustX
                                AdjustZ = Math.Sin(Functions.Radians((SpkrLst(40, n)))) * ((SpkrLst(10, n) / 12) / 2)
                                SpkrLst(34, n) = SpkrLst(34, n) - AdjustZ 'Y
                                AdjustX = ((SpkrLst(10, n) / 12) / 2) - (Math.Cos(Functions.Radians((SpkrLst(40, n)))) * ((SpkrLst(10, n) / 12) / 2))
                                SpkrLst(33, n) = SpkrLst(33, n) - AdjustX 'x
                            Case "Rear"
                                SpkrLst(33, n) = SpkrLst(33, n) - (SpkrLst(10, n) / 12) / 2 - WallSep 'X + Depth/2
                                SpkrLst(40, n) = DataFunctions.SpeakerToe(0, n) 'toe-in
                                SpkrLst(41, n) = DataFunctions.SpeakerToe(1, n) 'toe-up
                                SpkrLst(42, n) = DataFunctions.SpeakerToe(2, n) 'roll
                                'Shift position to maintain proper angle toward MLP
                                SpkrLst(34, n) = SpkrLst(34, n) + ((SpkrLst(10, n) / 12) / 2 + WallSep) * Math.Tan(Functions.Radians((DataFunctions.SpeakerToe(0, n))))
                                'compensate position for rotation
                                AdjustZ = Math.Sin(Functions.Radians((SpkrLst(41, n)))) * ((SpkrLst(10, n) / 12) / 2)
                                SpkrLst(35, n) = SpkrLst(35, n) - AdjustZ
                                AdjustX = (Math.Cos(Functions.Radians((SpkrLst(41, n)))) * ((SpkrLst(10, n) / 12) / 2))
                                SpkrLst(33, n) = SpkrLst(33, n) - AdjustX 'x
                                AdjustZ = Math.Sin(Functions.Radians((SpkrLst(40, n)))) * ((SpkrLst(10, n) / 12) / 2)
                                SpkrLst(34, n) = SpkrLst(34, n) - AdjustZ 'Y
                                AdjustX = ((SpkrLst(10, n) / 12) / 2) - (Math.Cos(Functions.Radians((SpkrLst(40, n) - 180))) * ((SpkrLst(10, n) / 12) / 2))
                                SpkrLst(33, n) = SpkrLst(33, n) + AdjustX 'x
                            Case "Left"
                                SpkrLst(34, n) = SpkrLst(34, n) - (SpkrLst(10, n) / 12) / 2 + WallSep 'Y - Depth/2
                                SpkrLst(40, n) = DataFunctions.SpeakerToe(0, n) 'toe-in
                                SpkrLst(41, n) = DataFunctions.SpeakerToe(1, n) 'toe-up
                                SpkrLst(42, n) = DataFunctions.SpeakerToe(2, n) 'roll
                                'Shift position to maintain proper angle toward MLP
                                If DataFunctions.SpeakerToe(0, n) < 90 Then
                                    SpkrLst(33, n) = SpkrLst(33, n) + ((SpkrLst(10, n) / 12) / 2 + WallSep) * Math.Tan(Functions.Radians((90 - DataFunctions.SpeakerToe(0, n))))
                                Else
                                    SpkrLst(33, n) = SpkrLst(33, n) - ((SpkrLst(10, n) / 12) / 2 + WallSep) * Math.Tan(Functions.Radians((DataFunctions.SpeakerToe(0, n) - 90)))
                                End If
                                'compensate position for rotation
                                AdjustZ = Math.Sin(Functions.Radians((SpkrLst(42, n)))) * ((SpkrLst(10, n) / 12) / 2)
                                SpkrLst(35, n) = SpkrLst(35, n) + AdjustZ 'z
                                AdjustX = (Math.Cos(Functions.Radians((SpkrLst(42, n)))) * ((SpkrLst(10, n) / 12) / 2))
                                SpkrLst(34, n) = SpkrLst(34, n) + AdjustX 'y
                                AdjustZ = Math.Sin(Functions.Radians((90 - SpkrLst(40, n)))) * ((SpkrLst(10, n) / 12) / 2)
                                SpkrLst(33, n) = SpkrLst(33, n) + AdjustZ 'X
                                AdjustX = ((SpkrLst(10, n) / 12) / 2) - (Math.Cos(Functions.Radians((90 - SpkrLst(40, n)))) * ((SpkrLst(10, n) / 12) / 2))
                                SpkrLst(34, n) = SpkrLst(34, n) - AdjustX 'Y
                            Case "Right"
                                SpkrLst(34, n) = SpkrLst(34, n) + (SpkrLst(10, n) / 12) / 2 - WallSep 'Y + Depth/2
                                SpkrLst(40, n) = DataFunctions.SpeakerToe(0, n) 'toe-in
                                SpkrLst(41, n) = DataFunctions.SpeakerToe(1, n) 'toe-up
                                SpkrLst(42, n) = DataFunctions.SpeakerToe(2, n) 'roll
                                'Shift position to maintain proper angle toward MLP
                                If DataFunctions.SpeakerToe(0, n) > 270 Then
                                    SpkrLst(33, n) = SpkrLst(33, n) - ((SpkrLst(10, n) / 12) / 2 + WallSep) * Math.Tan(Functions.Radians((270 - DataFunctions.SpeakerToe(0, n))))
                                Else
                                    SpkrLst(33, n) = SpkrLst(33, n) + ((SpkrLst(10, n) / 12) / 2 + WallSep) * Math.Tan(Functions.Radians((DataFunctions.SpeakerToe(0, n) - 270)))
                                End If
                                'compensate position for rotation
                                AdjustZ = Math.Sin(Functions.Radians((SpkrLst(42, n)))) * ((SpkrLst(10, n) / 12) / 2)
                                SpkrLst(35, n) = SpkrLst(35, n) - AdjustZ 'z roll
                                AdjustX = (Math.Cos(Functions.Radians((SpkrLst(42, n)))) * ((SpkrLst(10, n) / 12) / 2))
                                SpkrLst(34, n) = SpkrLst(34, n) - AdjustX 'y
                                AdjustZ = Math.Sin(Functions.Radians((SpkrLst(40, n) - 270))) * ((SpkrLst(10, n) / 12) / 2)
                                SpkrLst(33, n) = SpkrLst(33, n) + AdjustZ 'X
                                AdjustX = ((SpkrLst(10, n) / 12) / 2) - (Math.Cos(Functions.Radians((SpkrLst(40, n) - 270))) * ((SpkrLst(10, n) / 12) / 2))
                                SpkrLst(34, n) = SpkrLst(34, n) - AdjustX 'Y
                            Case "Ceiling"
                                AdjustX = Math.Sin(Functions.Radians((SpkrLst(42, n)))) * ((SpkrLst(10, n) / 12) / 2)
                                SpkrLst(34, n) = SpkrLst(34, n) - AdjustX 'y
                                AdjustZ = (Math.Cos(Functions.Radians((SpkrLst(42, n)))) * ((SpkrLst(10, n) / 12) / 2))
                                SpkrLst(35, n) = SpkrLst(35, n) + AdjustZ 'z
                                AdjustX = Math.Sin(Functions.Radians((SpkrLst(41, n) - 90))) * ((SpkrLst(10, n) / 12) / 2)
                                SpkrLst(33, n) = SpkrLst(33, n) - AdjustX 'x
                                AdjustZ = ((SpkrLst(10, n) / 12) / 2) - (Math.Cos(Functions.Radians((SpkrLst(41, n) - 90))) * ((SpkrLst(10, n) / 12) / 2))
                                SpkrLst(35, n) = SpkrLst(35, n) - AdjustZ 'z
                        End Select
                    Case "Floorstanding"
                        WallSep = (3 / 12)
                        Select Case WallType(0) 'wall
                            Case "Front"
                                If SpkrLst(37, n) = "C" Then
                                    WallSep = WallSep * 2
                                    SpkrLst(33, n) = SpkrLst(33, n) + (SpkrLst(10, n) / 12) / 2 + WallSep 'X - Depth/2
                                    SpkrLst(40, n) = DataFunctions.SpeakerToe(0, n) 'toe-in
                                Else
                                    WallSep = WallSep * 3
                                    SpkrLst(33, n) = SpkrLst(33, n) + (SpkrLst(10, n) / 12) / 2 + WallSep 'X - Depth/2
                                    SpkrLst(40, n) = DataFunctions.SpeakerToe(0, n) 'toe-in
                                End If
                                SpkrLst(35, n) = DataFunctions.BaseHeight(n) + (SpkrLst(8, n) / 12) / 2 'Z = Height/2
                                'Shift position to maintain proper angle toward MLP
                                SpkrLst(34, n) = SpkrLst(34, n) - ((SpkrLst(10, n) / 12) / 2 + WallSep) * Math.Tan(Functions.Radians((DataFunctions.SpeakerToe(0, n))))
                                SpkrLst(41, n) = 0 'toe-up
                                SpkrLst(42, n) = 180 'roll
                            Case "Rear"
                                SpkrLst(33, n) = SpkrLst(33, n) - (SpkrLst(10, n) / 12) / 2 - WallSep 'X + Depth/2
                                SpkrLst(40, n) = DataFunctions.SpeakerToe(0, n) 'toe-in
                                'Shift position to maintain proper angle toward MLP
                                SpkrLst(34, n) = SpkrLst(34, n) + ((SpkrLst(10, n) / 12) / 2 + WallSep) * Math.Tan(Functions.Radians((DataFunctions.SpeakerToe(0, n))))
                                SpkrLst(35, n) = DataFunctions.BaseHeight(n) + (SpkrLst(8, n) / 12) / 2 'Z = Height/2
                                SpkrLst(41, n) = 0 'toe-up
                                SpkrLst(42, n) = 180 'roll
                            Case "Left"
                                SpkrLst(34, n) = SpkrLst(34, n) - (SpkrLst(10, n) / 12) / 2 - WallSep 'Y - Depth/2
                                SpkrLst(40, n) = DataFunctions.SpeakerToe(0, n) 'toe-in
                                SpkrLst(35, n) = DataFunctions.BaseHeight(n) + (SpkrLst(8, n) / 12) / 2 'Z = Height/2
                                'Shift position to maintain proper angle toward MLP
                                If DataFunctions.SpeakerToe(0, n) < 90 Then
                                    SpkrLst(33, n) = SpkrLst(33, n) + ((SpkrLst(10, n) / 12) / 2 + WallSep) * Math.Tan(Functions.Radians((90 - DataFunctions.SpeakerToe(0, n))))
                                Else
                                    SpkrLst(33, n) = SpkrLst(33, n) - ((SpkrLst(10, n) / 12) / 2 + WallSep) * Math.Tan(Functions.Radians((DataFunctions.SpeakerToe(0, n) - 90)))
                                End If
                                SpkrLst(41, n) = 0 'toe-up
                                SpkrLst(42, n) = 180 'roll
                            Case "Right"
                                SpkrLst(34, n) = SpkrLst(34, n) + (SpkrLst(10, n) / 12) / 2 + WallSep 'Y + Depth/2
                                SpkrLst(40, n) = DataFunctions.SpeakerToe(0, n) 'toe-in
                                SpkrLst(35, n) = DataFunctions.BaseHeight(n) + (SpkrLst(8, n) / 12) / 2 'Z = Height/2
                                'Shift position to maintain proper angle toward MLP
                                If DataFunctions.SpeakerToe(0, n) > 270 Then
                                    SpkrLst(33, n) = SpkrLst(33, n) - ((SpkrLst(10, n) / 12) / 2 + WallSep) * Math.Tan(Functions.Radians((270 - DataFunctions.SpeakerToe(0, n))))
                                Else
                                    SpkrLst(33, n) = SpkrLst(33, n) + ((SpkrLst(10, n) / 12) / 2 + WallSep) * Math.Tan(Functions.Radians((DataFunctions.SpeakerToe(0, n) - 270)))
                                End If
                                SpkrLst(41, n) = 0 'toe-up
                                SpkrLst(42, n) = 180 'roll
                            Case "Ceiling"
                                AdjustX = Math.Sin(Functions.Radians((SpkrLst(42, n)))) * ((SpkrLst(10, n) / 12) / 2)
                                SpkrLst(34, n) = SpkrLst(34, n) - AdjustX 'y
                                AdjustZ = (Math.Cos(Functions.Radians((SpkrLst(42, n)))) * ((SpkrLst(10, n) / 12) / 2))
                                SpkrLst(35, n) = SpkrLst(35, n) + AdjustZ 'z
                                AdjustX = Math.Sin(Functions.Radians((SpkrLst(41, n) - 90))) * ((SpkrLst(10, n) / 12) / 2)
                                SpkrLst(33, n) = SpkrLst(33, n) - AdjustX 'x
                                AdjustZ = ((SpkrLst(10, n) / 12) / 2) - (Math.Cos(Functions.Radians((SpkrLst(41, n) - 90))) * ((SpkrLst(10, n) / 12) / 2))
                                SpkrLst(35, n) = SpkrLst(35, n) - AdjustZ 'z
                        End Select
                End Select
            End If
        Next n
        Return SpkrLst
    End Function

    Public Function UpdateSpkrConfig(ByRef ProjNumb As String, ByRef SelectedSpkrConfigur As String _
, ByRef ConfigTypeDolby As String, ByRef Seatss As Object, ByRef SpkrInfo As Object) As Object
        Dim MLP As Object = 0, Spkrs As Object, Walls As Object, ImagePnt As Object, SpkrCnt As Integer, CeilingPos(2, 3) As Single
        Dim MirrorPnts As Object, MirrorArr As Object, TopPoint(2, 1) As Object, SpkrCon As String, SpkrLabel As String, ToeArr As Object
        Dim SpkrInfoEach As Object

        DataFunctions.SurroundHeight = 100 ' initialize variable for overall highest surround height used in CalcMirrorPnts

        'dbServercls.ProjectSeatUpdate = Seatss 'update seats with latest list using clsdbServer
        ReDim MLP(2)
        For n = 0 To UBound(Seatss, 2)
            If Seatss(10, n) = 3 Then
                MLP(0) = Seatss(4, n)
                MLP(1) = Seatss(5, n)
                MLP(2) = Seatss(6, n)
            End If
        Next

        Spkrs = Library.SpeakerAngles(SelectedSpkrConfigur, ConfigTypeDolby) ' get selected speaker configuration from Library module

        ReDim DataFunctions.BaseHeight(UBound(Spkrs, 2))
        ReDim DataFunctions.SpeakerToe(2, UBound(Spkrs, 2))
        ReDim MirrorPnts(2, 5, UBound(Spkrs, 2))
        ReDim SpkrInfoEach(4)

        dbServercls.ProjectInitial = ProjNumb
        Walls = dbServercls.GetProjectWalls() 'get project walls from clsdbServer

        For SpkrCnt = 0 To UBound(Spkrs, 2)
            SpkrInfoEach(0) = SpkrInfo(0, SpkrCnt)
            SpkrInfoEach(1) = SpkrInfo(1, SpkrCnt)
            SpkrInfoEach(2) = SpkrInfo(2, SpkrCnt)
            SpkrInfoEach(3) = SpkrInfo(3, SpkrCnt)
            SpkrInfoEach(4) = SpkrInfo(4, SpkrCnt)
            ImagePnt = DataFunctions.CalcImagePoint(MLP, Spkrs, SpkrCnt, TopPoint)
            SpkrCon = Spkrs(4, SpkrCnt) 'use of spkrs; Front, SurroundR, SurroundL, Ceiling, Rear
            SpkrLabel = Spkrs(3, SpkrCnt)
            MirrorArr = DataFunctions.CalcMirrorPnts(MLP, Walls, ImagePnt, SpkrLabel, TopPoint, SpkrCnt, SpkrInfoEach)
            ToeArr = DataFunctions.CalcSpeakerToe(MLP, Spkrs, SpkrCnt, CSng(MirrorArr(0)), CSng(MirrorArr(1)), CSng(MirrorArr(2))) 'Calculate speaker toe to point speaker at listener for bookshelf and floor speakers
            DataFunctions.SpeakerToe(0, SpkrCnt) = ToeArr(0) 'toe-in for bookshelf
            DataFunctions.SpeakerToe(1, SpkrCnt) = ToeArr(1) 'toe-up for bookshelf
            DataFunctions.SpeakerToe(2, SpkrCnt) = ToeArr(2) 'roll for bookshelf
            MirrorPnts(0, 0, SpkrCnt) = MirrorArr(0) 'x center point
            MirrorPnts(1, 0, SpkrCnt) = MirrorArr(1) 'y center point
            MirrorPnts(2, 0, SpkrCnt) = MirrorArr(2) 'z center point
            MirrorPnts(0, 1, SpkrCnt) = MirrorArr(3) 'x box points
            MirrorPnts(1, 1, SpkrCnt) = MirrorArr(4) 'y
            MirrorPnts(2, 1, SpkrCnt) = MirrorArr(5) 'z
            MirrorPnts(0, 2, SpkrCnt) = MirrorArr(6) 'x
            MirrorPnts(1, 2, SpkrCnt) = MirrorArr(7) 'y
            MirrorPnts(2, 2, SpkrCnt) = MirrorArr(8) 'z
            MirrorPnts(0, 3, SpkrCnt) = MirrorArr(9) 'x
            MirrorPnts(1, 3, SpkrCnt) = MirrorArr(10) 'y
            MirrorPnts(2, 3, SpkrCnt) = MirrorArr(11) 'z
            MirrorPnts(0, 4, SpkrCnt) = MirrorArr(12) 'x
            MirrorPnts(1, 4, SpkrCnt) = MirrorArr(13) 'y
            MirrorPnts(2, 4, SpkrCnt) = MirrorArr(14) 'z
            MirrorPnts(0, 5, SpkrCnt) = SpkrLabel 'speaker label
            MirrorPnts(1, 5, SpkrCnt) = SpkrCon 'speaker use
            MirrorPnts(2, 5, SpkrCnt) = MirrorArr(15) 'wall index
            ' determine ceiling rails y position
            If Spkrs(3, SpkrCnt) = "R" Then
                TopPoint(0, 0) = MirrorArr(4) 'right point
                TopPoint(1, 0) = MirrorArr(7) 'left point
                TopPoint(2, 0) = MirrorArr(1) 'center point
            End If
            If Spkrs(3, SpkrCnt) = "L" Then
                TopPoint(0, 1) = MirrorArr(4) 'left point
                TopPoint(1, 1) = MirrorArr(7) 'right point
                TopPoint(2, 1) = MirrorArr(1) 'center point
            End If
        Next SpkrCnt

        'MirrorPnts = EqualizePositions(MirrorPnts)
        'CreateCSV (MirrorPnts)
        UpdateSpkrConfig = MirrorPnts

    End Function
    Public Function SpeakerAngles(ByVal DolbyName As Object, ByVal ConfigBrand As Object) As Object
        Dim Outdata As Object

        Outdata = Library.SpeakerAngles(DolbyName, ConfigBrand)

        Return Outdata

    End Function
    Public Function GetDolbyArray() As Object
        Dim DolbyArray As Object

        ReDim DolbyArray(8, 1)

        'Configuration
        DolbyArray(0, 0) = "2.0"
        DolbyArray(1, 0) = "5.1"
        DolbyArray(2, 0) = "5.1.2"
        DolbyArray(3, 0) = "5.1.4"
        DolbyArray(4, 0) = "7.1"
        DolbyArray(5, 0) = "7.1.2"
        DolbyArray(6, 0) = "7.1.4"
        DolbyArray(7, 0) = "9.1.2"
        DolbyArray(8, 0) = "9.1.4"

        'Number of speakers
        DolbyArray(0, 1) = "2"
        DolbyArray(1, 1) = "5"
        DolbyArray(2, 1) = "7"
        DolbyArray(3, 1) = "9"
        DolbyArray(4, 1) = "7"
        DolbyArray(5, 1) = "9"
        DolbyArray(6, 1) = "11"
        DolbyArray(7, 1) = "11"
        DolbyArray(8, 1) = "13"

        Return DolbyArray
    End Function
End Class