Public Class CommonResources
#Region "Public constants"
    Public Const KGauss As Double = 0.01720209895
    Public Const EpochB1950 As Double = -0.500002108145107 * 36525 + 2451545
    Public Const EpochJ2000 As Double = 0 * 36525 + 2451545
    Public Const c As Double = 299792458
    Public Const AU As Double = 149.6 * 10 ^ 9
    Public Const OblEclJ2000 As Double = 23.4392911111111
    Public Enum CoordType
        Geometric = 1
        Astrometric = 2
        Apparent = 3
    End Enum
#End Region
#Region "Public Structures"
    Public Structure PosSpeed
        Dim XPosition As Double
        Dim YPosition As Double
        Dim ZPosition As Double

        Dim XSpeed As Double
        Dim YSpeed As Double
        Dim ZSpeed As Double

        Sub New(ByVal X As Double, ByVal Y As Double, ByVal Z As Double, ByVal vX As Double, ByVal vY As Double, ByVal vZ As Double)
            XPosition = X
            YPosition = Y
            ZPosition = Z

            XSpeed = vX
            YSpeed = vY
            ZSpeed = vZ
        End Sub
    End Structure
    Public Structure PolarCoord
        Dim Radius As Double
        Dim Inclination As Double
        Dim Azimuth As Double

        Sub New(ByVal R As Double, ByVal I As Double, ByVal A As Double)
            Radius = R
            Inclination = I
            Azimuth = A
        End Sub
    End Structure

    Public Structure Ephemeris
        Dim Instant As TimeFunctions.DateWithEra
        Dim RA As Double
        Dim Decl As Double
        Dim DistSun As Double
        Dim DistEarth As Double
        Dim Comment As String

        Sub New(ByVal CalcInst As TimeFunctions.DateWithEra, ByVal RightAsc As Double, ByVal Declination As Double, ByVal DistanceSun As Double, ByVal DistanceEarth As Double, ByVal Note As String)
            Instant = CalcInst
            RA = RightAsc
            Decl = Declination
            DistSun = DistanceSun
            DistEarth = DistanceEarth
            Comment = Note
        End Sub
    End Structure
#End Region
#Region "Common Functions"
    Public Function RadiansToDegrees(ByVal Radians As Double) As Double
        Return Radians * 180.0 / Math.PI
    End Function
    Public Function DegreesToRadians(ByVal Degrees As Double) As Double
        Return (Math.PI / 180) * Degrees
    End Function
    Public Function DegreesToRadiansCorrect(ByVal Degrees As Double) As Double
        Return ReduceRad(DegreesToRadians(Degrees))
    End Function
    Public Function ReduceRad(ByVal FullAngle As Double) As Double
        Return DecimalPart(FullAngle / (2 * Math.PI)) * 2 * Math.PI
    End Function
    Public Function DecimalPart(ByVal number As Double) As Double
        If number >= 0 Then
            Return number - Math.Floor(number)
        Else
            Return number - Math.Ceiling(number)
        End If
    End Function
    ''' <summary>
    ''' Return a Time object starting from Hours and decimal
    ''' </summary>
    ''' <param name="HoursDecimal"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function TimeFromHours(ByVal HoursDecimal As Double) As DateTime
        'TODO Check for negative values
        Dim Hours As Integer
        Dim Minutes, Seconds As Double
        Dim FullTime As New DateTime

        Hours = Math.Floor(HoursDecimal)
        Minutes = DecimalPart(HoursDecimal) * 60
        Seconds = DecimalPart(Minutes) * 60

        FullTime = DateTime.Parse(Hours & ":" & Math.Floor(Minutes) & ":" & Math.Floor(Seconds))
        Return FullTime
    End Function
    ''' <summary>
    ''' Return Hours and decimal starting from DateTime object
    ''' </summary>
    ''' <param name="FullTime"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function HoursFromTime(ByVal FullTime As DateTime) As Double
        'TODO Check for negative values
        Dim HoursDecimal As Double

        With FullTime
            HoursDecimal = (((.Second / 60) + .Minute) / 60) + .Hour
        End With

        Return HoursDecimal
    End Function
    Public Function DegreesToHours(ByVal Angle As Double) As Double
        Return Angle * 24 / 360
    End Function
    Public Function HoursToDegrees(ByVal Angle As Double) As Double
        Return Angle * 360 / 24
    End Function
#End Region
End Class
