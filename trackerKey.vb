Module trackerKey

    Public Class keyCheck
        Public loc As String
        Public area As String
        Public zone As Byte
        Public name As String
        Public scan As Boolean
        Public checked As Boolean
        Public gs As Boolean
        Public cow As Boolean
        Public scrub As Boolean
        Public shop As Boolean
        Public forced As Boolean
        Public logic As String

        Public Sub New()
            loc = String.Empty
            area = String.Empty
            zone = 99
            name = String.Empty
            checked = False
            scan = False
            gs = False
            cow = False
            scrub = False
            shop = False
            forced = False
            logic = String.Empty
        End Sub
    End Class
End Module
