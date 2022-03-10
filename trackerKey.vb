Module trackerKey

    Public Class keyCheck
        Public loc As String
        Public area As String
        Public name As String
        Public scan As Boolean
        Public checked As Boolean
        Public gs As Boolean
        Public cow As Boolean
        Public scrub As Boolean
        Public shop As Boolean
        Public dungeon As Integer
        Public forced As Boolean

        Public Sub New()
            loc = String.Empty
            area = String.Empty
            name = String.Empty
            checked = False
            scan = True
            gs = False
            cow = False
            scrub = False
            shop = False
            dungeon = 0
            forced = False
        End Sub
    End Class
End Module
