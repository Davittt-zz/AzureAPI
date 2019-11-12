Imports System.Data.SqlClient
Public Class dbServer
    Dim connector As String = "Server=tcp:lemaytest.database.windows.net,1433;Initial Catalog=SQLAcousticalc;Persist Security Info=False;User ID=LemayTestdb;Password=34r%d$$#qwfhe$;MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;"
    Dim OutputData As Object()
    Dim OutputData12(1, 0) As Object
    Dim OutputData32(3, 0) As Object
    Dim OutputData62(6, 0) As Object
    Dim OutputData102(10, 0) As Object
    Dim OutputData172(17, 0) As Object
    Dim OutputData372(37, 0) As Object
    Dim User As String
    Dim Family(1) As Object
    Dim Project As String
    Dim WallIndx As Int16
    Dim Completed As Boolean
    Dim RowInputKey(1) As Object
    Public WriteOnly Property ProjectInitial() As String
        Set(value As String)
            Me.Project = value
        End Set
    End Property
    Public WriteOnly Property ProjectWallIndx() As String
        Set(value As String)
            Me.WallIndx = value
        End Set
    End Property
    Public ReadOnly Property GetProjectWalls() As Object
        Get
            ProjectWallsGet()
            Return Me.OutputData172
        End Get
    End Property
    Public ReadOnly Property GetProjectCoord() As Object
        Get
            WallCoordGet()
            Return Me.OutputData
        End Get
    End Property

    Public Sub WallCoordGet()

        Using connection As New SqlConnection(connector)
            connection.Open()
            ReDim OutputData(11)
            Dim TableName As String = "ProjectWalls"
            Dim SQLadapter As New SqlDataAdapter("SELECT * FROM " & TableName & " ORDER by ProjectNumber", connection)
            Dim SQLbuilder As New SqlCommandBuilder(SQLadapter)
            Dim SQLds As New DataSet()
            SQLadapter.Fill(SQLds, TableName)
            Dim ChosenTable As DataTable = SQLds.Tables(TableName)
            Dim row As DataRow = ChosenTable.NewRow()
            For Each row In ChosenTable.Rows

                If Me.Project = row("ProjectNumber") And Me.WallIndx = row("Index") Then
                    OutputData(0) = row("X1")
                    OutputData(1) = row("Y1")
                    OutputData(2) = row("Z1")
                    OutputData(3) = row("X2")
                    OutputData(4) = row("Y2")
                    OutputData(5) = row("Z2")
                    OutputData(6) = row("X3")
                    OutputData(7) = row("Y3")
                    OutputData(8) = row("Z3")
                    OutputData(9) = row("X4")
                    OutputData(10) = row("Y4")
                    OutputData(11) = row("Z4")
                End If
            Next

            connection.Close()
        End Using

    End Sub
    Public Sub ProjectWallsGet()

        Using connection As New SqlConnection(connector)
            connection.Open()
            Dim Cnt As Int16 = 0
            Dim TableName As String = "ProjectWalls"
            Dim SQLadapter As New SqlDataAdapter("SELECT * FROM " & TableName & " ORDER by ProjectNumber", connection)
            Dim SQLbuilder As New SqlCommandBuilder(SQLadapter)
            Dim SQLds As New DataSet()
            SQLadapter.Fill(SQLds, TableName)
            Dim ChosenTable As DataTable = SQLds.Tables(TableName)
            Dim row As DataRow = ChosenTable.NewRow()
            For Each row In ChosenTable.Rows

                If Project = row("ProjectNumber") Then
                    OutputData172(0, Cnt) = row("Index")
                    OutputData172(2, Cnt) = row("WallName")
                    OutputData172(1, Cnt) = row("ProjectNumber")
                    OutputData172(3, Cnt) = row("DrawRefIndex")
                    OutputData172(4, Cnt) = row("MaterialIndex")
                    OutputData172(5, Cnt) = row("X1")
                    OutputData172(6, Cnt) = row("Y1")
                    OutputData172(7, Cnt) = row("Z1")
                    OutputData172(8, Cnt) = row("X2")
                    OutputData172(9, Cnt) = row("Y2")
                    OutputData172(10, Cnt) = row("Z2")
                    OutputData172(11, Cnt) = row("X3")
                    OutputData172(12, Cnt) = row("Y3")
                    OutputData172(13, Cnt) = row("Z3")
                    OutputData172(14, Cnt) = row("X4")
                    OutputData172(15, Cnt) = row("Y4")
                    OutputData172(16, Cnt) = row("Z4")
                    OutputData172(17, Cnt) = row("Thickness")
                    Cnt = Cnt + 1
                    ReDim Preserve OutputData172(17, Cnt)
                End If
            Next

            If Cnt = 0 Then OutputData172(0, 0) = "Walls Not Found"
            If Cnt > 0 Then
                ReDim Preserve OutputData172(17, Cnt - 1)
            End If
            connection.Close()
        End Using
    End Sub
    Public WriteOnly Property ProjectSeatUpdate()
        Set(SeatData As Object)

            Using connection As New SqlConnection(connector)
                connection.Open()
                Dim Cnt As Int16 = 0
                Dim TableName As String = "ProjectSeating"
                Dim SQLadapter As New SqlDataAdapter("SELECT * FROM " & TableName, connection)
                Dim SQLbuilder As New SqlCommandBuilder(SQLadapter)
                Dim SQLds As New DataSet()
                SQLadapter.Fill(SQLds, TableName)
                Dim ChosenTable As DataTable = SQLds.Tables(TableName)
                Dim row As DataRow = ChosenTable.NewRow()
                For Each row In ChosenTable.Rows

                    If SeatData(1, 0) = row("ProjectNumber") And SeatData(0, Cnt) = row("Index") Then
                        row("Index") = SeatData(0, Cnt)
                        row("ProjectNumber") = SeatData(1, Cnt)
                        row("ObjectIndex") = SeatData(2, Cnt)
                        row("DrawRefIndex") = SeatData(3, Cnt)
                        row("X") = SeatData(4, Cnt)
                        row("Y") = SeatData(5, Cnt)
                        row("z") = SeatData(6, Cnt)
                        row("EarOffsetX") = SeatData(7, Cnt)
                        row("EarOffsetY") = SeatData(8, Cnt)
                        row("EarOffsetZ") = SeatData(9, Cnt)
                        row("Status") = SeatData(10, Cnt)
                        Cnt = Cnt + 1
                        SQLadapter.Update(SQLds, TableName)
                    End If
                Next
                If Cnt <= UBound(SeatData, 2) Or Cnt = 0 Then
                    For n = Cnt To UBound(SeatData, 2)
                        row("Index") = vbNull 'seat index; each seat will have a unique index i.e. 1,2,3,4
                        row("ProjectNumber") = SeatData(1, n) ' ViewIt project number as long
                        row("ObjectIndex") = SeatData(2, n) 'references the actual seat model used(useful for calculating ear offsets)
                        row("DrawRefIndex") = SeatData(3, n) ' unused for now do not leave null
                        row("X") = SeatData(4, n) 'listener ear/eye X position
                        row("Y") = SeatData(5, n) 'listener ear/eye Y position
                        row("z") = SeatData(6, n) 'listener ear/eye Z position
                        row("EarOffsetX") = SeatData(7, n) 'if the X,Y,&Z position is that of the seat, designates the correction to acheive listener actual position
                        row("EarOffsetY") = SeatData(8, n)
                        row("EarOffsetZ") = SeatData(9, n)
                        row("Status") = SeatData(10, n) ' indicates if listener is active "1" or inactive "0" or if the entry represents the central listening position (donut) "3"
                        ChosenTable.Rows.Add(row)
                        SQLadapter.Update(SQLds, TableName)
                    Next
                End If
                connection.Close()
            End Using
        End Set

    End Property
End Class
