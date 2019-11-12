Public Class Library
    Public Function SpeakerAngles(ByRef Config As String, ByRef Brand As String) As Object
        Dim ConfigArray As Object

        ReDim ConfigArray(4, 0)
        ConfigArray(0, 0) = 22 'start with left most
        ConfigArray(1, 0) = 30
        ConfigArray(2, 0) = 27
        ConfigArray(3, 0) = "L"
        ConfigArray(4, 0) = "Front"

        Select Case Brand
            Case "Dolby"
                Select Case Config
                    Case "2.0"
                        ReDim ConfigArray(4, 1)
                        ConfigArray(0, 0) = 330 'start with left most
                        ConfigArray(1, 0) = 338
                        ConfigArray(2, 0) = 333
                        ConfigArray(3, 0) = "L"  'speaker designation based on Trinnov name guidelines
                        ConfigArray(4, 0) = "Front"  'speaker function Front, Rear, SurroundL, SurroundR, Ceiling

                        ConfigArray(0, 1) = 22 '330
                        ConfigArray(1, 1) = 30 '338
                        ConfigArray(2, 1) = 27 '333
                        ConfigArray(3, 1) = "R"
                        ConfigArray(4, 1) = "Front"

                    Case "5.1"
                        ReDim ConfigArray(4, 4)
                        ConfigArray(0, 0) = 330 'start with left most
                        ConfigArray(1, 0) = 338
                        ConfigArray(2, 0) = 333
                        ConfigArray(3, 0) = "L"  'speaker designation based on Trinnov name guidelines
                        ConfigArray(4, 0) = "Front"  'speaker function Front, Rear, SurroundL, SurroundR, Ceiling

                        ConfigArray(0, 1) = 22 '330
                        ConfigArray(1, 1) = 30 '338
                        ConfigArray(2, 1) = 27 '333
                        ConfigArray(3, 1) = "R"
                        ConfigArray(4, 1) = "Front"

                        ConfigArray(0, 2) = 1
                        ConfigArray(1, 2) = 359.0#
                        ConfigArray(2, 2) = 0
                        ConfigArray(3, 2) = "C"
                        ConfigArray(4, 2) = "Front"

                        ConfigArray(0, 3) = 240
                        ConfigArray(1, 3) = 250
                        ConfigArray(2, 3) = 245
                        ConfigArray(3, 3) = "Ls"
                        ConfigArray(4, 3) = "SurroundL"

                        ConfigArray(0, 4) = 110
                        ConfigArray(1, 4) = 120
                        ConfigArray(2, 4) = 115
                        ConfigArray(3, 4) = "Rs"
                        ConfigArray(4, 4) = "SurroundR"

                    Case "7.1"
                        ReDim ConfigArray(4, 6)
                        ConfigArray(0, 0) = 330 'start with left most
                        ConfigArray(1, 0) = 338
                        ConfigArray(2, 0) = 333
                        ConfigArray(3, 0) = "L"  'speaker designation based on Trinnov name guidelines
                        ConfigArray(4, 0) = "Front"  'speaker function Front, Rear, SurroundL, SurroundR, Ceiling

                        ConfigArray(0, 1) = 22 '330
                        ConfigArray(1, 1) = 30 '338
                        ConfigArray(2, 1) = 27 '333
                        ConfigArray(3, 1) = "R"
                        ConfigArray(4, 1) = "Front"

                        ConfigArray(0, 2) = 1
                        ConfigArray(1, 2) = 359.0#
                        ConfigArray(2, 2) = 0
                        ConfigArray(3, 2) = "C"
                        ConfigArray(4, 2) = "Front"

                        ConfigArray(0, 3) = 250
                        ConfigArray(1, 3) = 270
                        ConfigArray(2, 3) = 250
                        ConfigArray(3, 3) = "Ls"
                        ConfigArray(4, 3) = "SurroundL"

                        ConfigArray(0, 4) = 90
                        ConfigArray(1, 4) = 110
                        ConfigArray(2, 4) = 110
                        ConfigArray(3, 4) = "Rs"
                        ConfigArray(4, 4) = "SurroundR"

                        ConfigArray(0, 5) = 210
                        ConfigArray(1, 5) = 225
                        ConfigArray(2, 5) = 215
                        ConfigArray(3, 5) = "Lrs"
                        ConfigArray(4, 5) = "RearL"

                        ConfigArray(0, 6) = 135
                        ConfigArray(1, 6) = 150
                        ConfigArray(2, 6) = 145
                        ConfigArray(3, 6) = "Rrs"
                        ConfigArray(4, 6) = "RearR"
                    Case "7.1.2"
                        ReDim ConfigArray(4, 8)
                        ConfigArray(0, 0) = 330 'start with left most
                        ConfigArray(1, 0) = 338
                        ConfigArray(2, 0) = 333
                        ConfigArray(3, 0) = "L"  'speaker designation based on Trinnov name guidelines
                        ConfigArray(4, 0) = "Front"  'speaker function Front, Rear, SurroundL, SurroundR, Ceiling

                        ConfigArray(0, 1) = 22 '330
                        ConfigArray(1, 1) = 30 '338
                        ConfigArray(2, 1) = 27 '333
                        ConfigArray(3, 1) = "R"
                        ConfigArray(4, 1) = "Front"

                        ConfigArray(0, 2) = 1
                        ConfigArray(1, 2) = 359.0#
                        ConfigArray(2, 2) = 0
                        ConfigArray(3, 2) = "C"
                        ConfigArray(4, 2) = "Front"

                        ConfigArray(0, 3) = 250
                        ConfigArray(1, 3) = 270
                        ConfigArray(2, 3) = 250
                        ConfigArray(3, 3) = "Ls"
                        ConfigArray(4, 3) = "SurroundL"

                        ConfigArray(0, 4) = 90
                        ConfigArray(1, 4) = 110
                        ConfigArray(2, 4) = 110
                        ConfigArray(3, 4) = "Rs"
                        ConfigArray(4, 4) = "SurroundR"

                        ConfigArray(0, 5) = 210
                        ConfigArray(1, 5) = 225
                        ConfigArray(2, 5) = 215
                        ConfigArray(3, 5) = "Lrs"
                        ConfigArray(4, 5) = "RearL"

                        ConfigArray(0, 6) = 135
                        ConfigArray(1, 6) = 150
                        ConfigArray(2, 6) = 145
                        ConfigArray(3, 6) = "Rrs"
                        ConfigArray(4, 6) = "RearR"

                        ConfigArray(0, 7) = 100
                        ConfigArray(1, 7) = 65
                        ConfigArray(2, 7) = 82
                        ConfigArray(3, 7) = "Ltf"
                        ConfigArray(4, 7) = "CeilingL"

                        ConfigArray(0, 8) = 100
                        ConfigArray(1, 8) = 65
                        ConfigArray(2, 8) = 82
                        ConfigArray(3, 8) = "Rtf"
                        ConfigArray(4, 8) = "CeilingR"

                    Case "7.1.4"
                        ReDim ConfigArray(4, 10)
                        ConfigArray(0, 0) = 330 'start with left most
                        ConfigArray(1, 0) = 338
                        ConfigArray(2, 0) = 333
                        ConfigArray(3, 0) = "L"  'speaker designation based on Trinnov name guidelines
                        ConfigArray(4, 0) = "Front"  'speaker function Front, Rear, SurroundL, SurroundR, Ceiling

                        ConfigArray(0, 1) = 22 '330
                        ConfigArray(1, 1) = 30 '338
                        ConfigArray(2, 1) = 27 '333
                        ConfigArray(3, 1) = "R"
                        ConfigArray(4, 1) = "Front"

                        ConfigArray(0, 2) = 1
                        ConfigArray(1, 2) = 359.0#
                        ConfigArray(2, 2) = 0
                        ConfigArray(3, 2) = "C"
                        ConfigArray(4, 2) = "Front"

                        ConfigArray(0, 3) = 250
                        ConfigArray(1, 3) = 270
                        ConfigArray(2, 3) = 250
                        ConfigArray(3, 3) = "Ls"
                        ConfigArray(4, 3) = "SurroundL"

                        ConfigArray(0, 4) = 90
                        ConfigArray(1, 4) = 110
                        ConfigArray(2, 4) = 110
                        ConfigArray(3, 4) = "Rs"
                        ConfigArray(4, 4) = "SurroundR"

                        ConfigArray(0, 5) = 210
                        ConfigArray(1, 5) = 225
                        ConfigArray(2, 5) = 215
                        ConfigArray(3, 5) = "Lrs"
                        ConfigArray(4, 5) = "RearL"

                        ConfigArray(0, 6) = 135
                        ConfigArray(1, 6) = 150
                        ConfigArray(2, 6) = 145
                        ConfigArray(3, 6) = "Rrs"
                        ConfigArray(4, 6) = "RearR"

                        ConfigArray(0, 7) = 55
                        ConfigArray(1, 7) = 30
                        ConfigArray(2, 7) = 42
                        ConfigArray(3, 7) = "Ltf"
                        ConfigArray(4, 7) = "CeilingL"

                        ConfigArray(0, 8) = 150
                        ConfigArray(1, 8) = 125
                        ConfigArray(2, 8) = 137
                        ConfigArray(3, 8) = "Ltr"
                        ConfigArray(4, 8) = "CeilingL"

                        ConfigArray(0, 9) = 55
                        ConfigArray(1, 9) = 30
                        ConfigArray(2, 9) = 42
                        ConfigArray(3, 9) = "Rtf"
                        ConfigArray(4, 9) = "CeilingR"

                        ConfigArray(0, 10) = 150
                        ConfigArray(1, 10) = 125
                        ConfigArray(2, 10) = 137
                        ConfigArray(3, 10) = "Rtr"
                        ConfigArray(4, 10) = "CeilingR"
                    Case "5.1.2"
                        ReDim ConfigArray(4, 6)
                        ConfigArray(0, 0) = 330 'start with left most
                        ConfigArray(1, 0) = 338
                        ConfigArray(2, 0) = 333
                        ConfigArray(3, 0) = "L"  'speaker designation based on Trinnov name guidelines
                        ConfigArray(4, 0) = "Front"  'speaker function Front, Rear, SurroundL, SurroundR, Ceiling

                        ConfigArray(0, 1) = 22 '330
                        ConfigArray(1, 1) = 20 '338
                        ConfigArray(2, 1) = 27 '333
                        ConfigArray(3, 1) = "R"
                        ConfigArray(4, 1) = "Front"

                        ConfigArray(0, 2) = 1
                        ConfigArray(1, 2) = 359.0#
                        ConfigArray(2, 2) = 0
                        ConfigArray(3, 2) = "C"
                        ConfigArray(4, 2) = "Front"

                        ConfigArray(0, 3) = 240
                        ConfigArray(1, 3) = 250
                        ConfigArray(2, 3) = 245
                        ConfigArray(3, 3) = "Ls"
                        ConfigArray(4, 3) = "SurroundL"

                        ConfigArray(0, 4) = 110
                        ConfigArray(1, 4) = 120
                        ConfigArray(2, 4) = 115
                        ConfigArray(3, 4) = "Rs"
                        ConfigArray(4, 4) = "SurroundR"

                        ConfigArray(0, 5) = 100
                        ConfigArray(1, 5) = 65
                        ConfigArray(2, 5) = 82
                        ConfigArray(3, 5) = "Ltf"
                        ConfigArray(4, 5) = "CeilingL"

                        ConfigArray(0, 6) = 100
                        ConfigArray(1, 6) = 65
                        ConfigArray(2, 6) = 82
                        ConfigArray(3, 6) = "Rtf"
                        ConfigArray(4, 6) = "CeilingR"
                    Case "5.1.4"
                        ReDim ConfigArray(4, 8)
                        ConfigArray(0, 0) = 330 'start with left most
                        ConfigArray(1, 0) = 338
                        ConfigArray(2, 0) = 333
                        ConfigArray(3, 0) = "L"  'speaker designation based on Trinnov name guidelines
                        ConfigArray(4, 0) = "Front"  'speaker function Front, Rear, SurroundL, SurroundR, Ceiling

                        ConfigArray(0, 1) = 22 '330
                        ConfigArray(1, 1) = 30 '338
                        ConfigArray(2, 1) = 27 '333
                        ConfigArray(3, 1) = "R"
                        ConfigArray(4, 1) = "Front"

                        ConfigArray(0, 2) = 1
                        ConfigArray(1, 2) = 359.0#
                        ConfigArray(2, 2) = 0
                        ConfigArray(3, 2) = "C"
                        ConfigArray(4, 2) = "Front"

                        ConfigArray(0, 3) = 240
                        ConfigArray(1, 3) = 250
                        ConfigArray(2, 3) = 245
                        ConfigArray(3, 3) = "Ls"
                        ConfigArray(4, 3) = "SurroundL"

                        ConfigArray(0, 4) = 110
                        ConfigArray(1, 4) = 120
                        ConfigArray(2, 4) = 115
                        ConfigArray(3, 4) = "Rs"
                        ConfigArray(4, 4) = "SurroundR"

                        ConfigArray(0, 5) = 55
                        ConfigArray(1, 5) = 30
                        ConfigArray(2, 5) = 42
                        ConfigArray(3, 5) = "Rtf"
                        ConfigArray(4, 5) = "CeilingR"

                        ConfigArray(0, 6) = 150
                        ConfigArray(1, 6) = 125
                        ConfigArray(2, 6) = 137
                        ConfigArray(3, 6) = "Rtr"
                        ConfigArray(4, 6) = "CeilingR"

                        ConfigArray(0, 7) = 55
                        ConfigArray(1, 7) = 30
                        ConfigArray(2, 7) = 42
                        ConfigArray(3, 7) = "Ltf"
                        ConfigArray(4, 7) = "CeilingL"

                        ConfigArray(0, 8) = 150
                        ConfigArray(1, 8) = 125
                        ConfigArray(2, 8) = 137
                        ConfigArray(3, 8) = "Ltr"
                        ConfigArray(4, 8) = "CeilingL"
                    Case "9.1.4"
                        ReDim ConfigArray(4, 12)
                        ConfigArray(0, 0) = 330 'Result 0 'start with left most
                        ConfigArray(1, 0) = 338 'Result 1
                        ConfigArray(2, 0) = 333 'Result 2
                        ConfigArray(3, 0) = "L"  'speaker designation based on Trinnov name guidelines
                        ConfigArray(4, 0) = "Front"  'speaker function Front, Rear, SurroundL, SurroundR, Ceiling

                        ConfigArray(0, 1) = 22 '330 'Result 0
                        ConfigArray(1, 1) = 30 '338 'Result 1
                        ConfigArray(2, 1) = 27 '333 'Result 2
                        ConfigArray(3, 1) = "R"
                        ConfigArray(4, 1) = "Front"

                        ConfigArray(0, 2) = 1
                        ConfigArray(1, 2) = 359.0#
                        ConfigArray(2, 2) = 0
                        ConfigArray(3, 2) = "C"
                        ConfigArray(4, 2) = "Front"

                        ConfigArray(0, 3) = 250 'Result 0
                        ConfigArray(1, 3) = 270 'Result 1
                        ConfigArray(2, 3) = 250 'Result 2
                        ConfigArray(3, 3) = "Ls"
                        ConfigArray(4, 3) = "SurroundL"

                        ConfigArray(0, 4) = 90
                        ConfigArray(1, 4) = 110
                        ConfigArray(2, 4) = 110
                        ConfigArray(3, 4) = "Rs"
                        ConfigArray(4, 4) = "SurroundR"

                        ConfigArray(0, 5) = 210
                        ConfigArray(1, 5) = 225
                        ConfigArray(2, 5) = 215
                        ConfigArray(3, 5) = "Lrs"
                        ConfigArray(4, 5) = "RearL"

                        ConfigArray(0, 6) = 135
                        ConfigArray(1, 6) = 150
                        ConfigArray(2, 6) = 145
                        ConfigArray(3, 6) = "Rrs"
                        ConfigArray(4, 6) = "RearR"

                        ConfigArray(0, 7) = 55
                        ConfigArray(1, 7) = 30
                        ConfigArray(2, 7) = 42
                        ConfigArray(3, 7) = "Ltf"
                        ConfigArray(4, 7) = "CeilingL"

                        ConfigArray(0, 8) = 150 'Result 0
                        ConfigArray(1, 8) = 125 'Result 1
                        ConfigArray(2, 8) = 137 'Result 2
                        ConfigArray(3, 8) = "Ltr"
                        ConfigArray(4, 8) = "CeilingL"

                        ConfigArray(0, 9) = 55
                        ConfigArray(1, 9) = 30
                        ConfigArray(2, 9) = 42
                        ConfigArray(3, 9) = "Rtf"
                        ConfigArray(4, 9) = "CeilingR"

                        ConfigArray(0, 10) = 150
                        ConfigArray(1, 10) = 125
                        ConfigArray(2, 10) = 137
                        ConfigArray(3, 10) = "Rtr"
                        ConfigArray(4, 10) = "CeilingR"

                        ConfigArray(0, 11) = 290 'Result 0
                        ConfigArray(1, 11) = 310 'Result 1
                        ConfigArray(2, 11) = 300 'Result 2
                        ConfigArray(3, 11) = "Lw"
                        ConfigArray(4, 11) = "WideL"

                        ConfigArray(0, 12) = 50 '290 'Result 0
                        ConfigArray(1, 12) = 70 '310 'Result 1
                        ConfigArray(2, 12) = 60 '300 'Result 2
                        ConfigArray(3, 12) = "Rw"
                        ConfigArray(4, 12) = "WideR"
                    Case "9.1.2"
                        ReDim ConfigArray(4, 10)
                        ConfigArray(0, 0) = 330 'start with left most
                        ConfigArray(1, 0) = 338
                        ConfigArray(2, 0) = 333
                        ConfigArray(3, 0) = "L"  'speaker designation based on Trinnov name guidelines
                        ConfigArray(4, 0) = "Front"  'speaker function Front, Rear, SurroundL, SurroundR, Ceiling

                        ConfigArray(0, 1) = 22 '330
                        ConfigArray(1, 1) = 30 '338
                        ConfigArray(2, 1) = 27 '333
                        ConfigArray(3, 1) = "R"
                        ConfigArray(4, 1) = "Front"

                        ConfigArray(0, 2) = 1
                        ConfigArray(1, 2) = 359.0#
                        ConfigArray(2, 2) = 0
                        ConfigArray(3, 2) = "C"
                        ConfigArray(4, 2) = "Front"

                        ConfigArray(0, 3) = 250
                        ConfigArray(1, 3) = 270
                        ConfigArray(2, 3) = 250
                        ConfigArray(3, 3) = "Ls"
                        ConfigArray(4, 3) = "SurroundL"

                        ConfigArray(0, 4) = 90 '250
                        ConfigArray(1, 4) = 110 '270
                        ConfigArray(2, 4) = 110 '250
                        ConfigArray(3, 4) = "Rs"
                        ConfigArray(4, 4) = "SurroundR"

                        ConfigArray(0, 5) = 210
                        ConfigArray(1, 5) = 225
                        ConfigArray(2, 5) = 215
                        ConfigArray(3, 5) = "Lrs"
                        ConfigArray(4, 5) = "RearL"

                        ConfigArray(0, 6) = 135 '210
                        ConfigArray(1, 6) = 150 '225
                        ConfigArray(2, 6) = 145 '215
                        ConfigArray(3, 6) = "Rrs"
                        ConfigArray(4, 6) = "RearR"

                        ConfigArray(0, 7) = 100
                        ConfigArray(1, 7) = 65
                        ConfigArray(2, 7) = 82
                        ConfigArray(3, 7) = "Ltf"
                        ConfigArray(4, 7) = "CeilingL"

                        ConfigArray(0, 8) = 100
                        ConfigArray(1, 8) = 65
                        ConfigArray(2, 8) = 82
                        ConfigArray(3, 8) = "Rtf"
                        ConfigArray(4, 8) = "CeilingR"

                        ConfigArray(0, 9) = 290
                        ConfigArray(1, 9) = 310
                        ConfigArray(2, 9) = 300
                        ConfigArray(3, 9) = "Lw"
                        ConfigArray(4, 9) = "WideL"

                        ConfigArray(0, 10) = 50 '290
                        ConfigArray(1, 10) = 70 '310
                        ConfigArray(2, 10) = 60 '300
                        ConfigArray(3, 10) = "Rw"
                        ConfigArray(4, 10) = "WideR"
                End Select

            Case "Imax"

            Case "Trinnov"

        End Select

        Return ConfigArray
    End Function

End Class