Public Class TimeFunctions
    Public Const EndGregorianCal As DateTime = #10/15/1582#

    Public Structure DateWithEra
        Dim FullDate As DateTime
        Dim IsBC As Boolean

        Sub New(ByVal AbsoluteDate As DateTime, ByVal BC As Boolean)
            FullDate = AbsoluteDate
            IsBC = BC
        End Sub
    End Structure

    Friend MiscFunctions As New CommonResources
    ''' <summary>
    ''' Return Julian date starting from Common Date and Time
    ''' </summary>
    ''' <param name="CommonDate">Common Date and Time to be converted</param>
    ''' <param name="IsBC">If the year is BC, turn this flag on</param>
    ''' <returns>Julian Date corresponding to the Common date given</returns>
    ''' <remarks></remarks>
    Public Function DateToJD(ByVal CommonDate As DateTime, ByVal IsBC As Boolean) As Double
        Dim y As Integer
        Dim m As Integer
        Dim d As Double

        Dim A As Integer
        Dim B As Integer

        Dim JD As Double

        If CommonDate.Month > 2 Then
            y = CommonDate.Year
            m = CommonDate.Month
        Else
            y = CommonDate.Year - 1
            m = CommonDate.Month + 12
        End If

        With CommonDate
            d = ((((.Second / 60) + .Minute) / 60) + .Hour) / 24 + .Day
        End With

        If IsBC Then
            y *= -1
        End If

        If CommonDate >= EndGregorianCal Then
            A = Math.Floor(y / 100)
            B = 2 - A + Math.Floor(A / 4)
        Else
            A = 0
            B = 0
        End If

        JD = Math.Floor(365.25 * (y + 4716)) + Math.Floor(30.6 * (m + 1)) + d + B - 1524.5

        Return JD
    End Function
    ''' <summary>
    ''' Return Common Date and Time starting from a given Julian Date
    ''' </summary>
    ''' <param name="JD">The Julian Date to be converted</param>
    ''' <returns>Common Date and Time corresponding to the given Julian Date</returns>
    ''' <remarks></remarks>
    Public Function JDToDate(ByVal JD As Double) As DateWithEra
        Dim Z As Integer
        Dim F As Double
        Dim A As Integer
        Dim Alpha As Double
        Dim B, C, D, E As Integer

        Dim Day As Double
        Dim m, y As Integer
        Dim hours, minutes, seconds As Double

        Dim CommonDate As String
        Dim FullDate As New DateWithEra

        Z = Math.Floor(JD)
        F = JD - Z + 0.5

        If Z < 2299161 Then
            A = Z
        Else
            Alpha = Math.Floor((Z - 1867216.25) / 36524.25)
            A = Z + 1 + Alpha - Math.Floor(Alpha / 4)
        End If

        B = A + 1524
        C = Math.Floor((B - 122.1) / 365.25)
        D = Math.Floor(365.25 * C)
        E = Math.Floor((B - D) / 30.6001)

        Day = B - D - Math.Floor(30.6001 * E) + F

        If E < 13.5 Then
            m = E - 1
        Else
            m = E - 13
        End If

        If m > 2.5 Then
            y = C - 4716
        Else
            y = C - 4715
        End If

        hours = MiscFunctions.DecimalPart(Day) * 24
        minutes = MiscFunctions.DecimalPart(hours) * 60
        seconds = MiscFunctions.DecimalPart(minutes) * 60

        If y < 0 Then
            FullDate.IsBC = True
            y *= -1
        Else
            FullDate.IsBC = False
        End If

        CommonDate = y & "-" & m & "-" & Math.Floor(Day) & " " & Math.Floor(hours) & ":" & Math.Floor(minutes) & ":" & Math.Floor(seconds)

        FullDate.FullDate = DateTime.Parse(CommonDate)

        Return FullDate
    End Function
    ''' <summary>
    ''' It calculates the difference in minutes between Terrestrial Time (aka Terrestrial Dynamical Time or Ephemeris Time) and Universal Time for a given year
    ''' </summary>
    ''' <param name="Year">The year where to calculate the difference</param>
    ''' <returns>The difference in minutes and decimal</returns>
    ''' <remarks></remarks>
    Public Function DeltaT(ByVal Year As Integer) As Double
        Const BasicYear As Integer = 1900
        Dim Difference As Double

        Difference = (Year - BasicYear) / 100

        Return 0.41 + 1.2053 * Difference + 0.4992 * Difference ^ 2

    End Function
    ''' <summary>
    ''' Return the Terrestrial Time (aka Terrestrial Dynamical Time or Ephemeris Time) corresponding to the given Universal Time value
    ''' </summary>
    ''' <param name="UniversalTime"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function EphemerisTime(ByVal UniversalTime As DateTime) As DateTime
        Dim Delta As Double

        Delta = DeltaT(UniversalTime.Year)
        Return UniversalTime.AddMinutes(Delta)
    End Function
    ''' <summary>
    ''' Return the Universal Time corresponding to the given Terrestrial Time (aka Terrestrial Dynamical Time or Ephemeris Time) value
    ''' </summary>
    ''' <param name="EphemerisTime"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function UniversalTime(ByVal EphemerisTime As DateTime) As DateTime
        Dim Delta As Double

        Delta = DeltaT(EphemerisTime.Year) * (-1)
        Return EphemerisTime.AddMinutes(Delta)
    End Function
    ''' <summary>
    ''' Return Easter Date for a given year in the calendar
    ''' </summary>
    ''' <param name="Year"></param>
    ''' <returns>A Date object containing the full date of the easter for the given year</returns>
    ''' <remarks>The function automatically distinguish between Julian and Gregorian calendar</remarks>
    Public Function EasterDate(ByVal Year As Integer) As Date
        Dim a, b, c, d, e, f, g, h, i, k, l, m, n, p As Double
        Dim Day As Integer
        Dim Month As Integer
        Dim FullData As New Date

        If Year < 1583 Then
            'Julian Calendar
            a = Year Mod 4

            b = Year Mod 7

            c = Year Mod 19

            d = (19 * c + 15) Mod 30

            e = (2 * a + 4 * b - d + 34) Mod 7

            f = Math.Floor((d + e + 114) / 31)
            g = (d + e + 114) Mod 31

            Day = g + 1
            Month = f
        Else
            'Gregorian Calendar
            a = Year Mod 19

            b = Math.Floor(Year / 100)
            c = Year Mod 100

            d = Math.Floor(b / 4)
            e = b Mod 4

            f = Math.Floor((b + 8) / 25)

            g = Math.Floor((b - f + 1) / 3)

            h = (19 * a + b - d - g + 15) Mod 30

            i = Math.Floor(c / 4)
            k = c Mod 4

            l = (32 + 2 * e + 2 * i - h - k) Mod 7

            m = Math.Floor((a + 11 * h + 22 * l) / 451)

            n = Math.Floor((h + l - 7 * m + 114) / 31)
            p = (h + l - 7 * m + 114) Mod 31

            Day = p + 1
            Month = n
        End If

        FullData = Date.Parse(Year & "-" & Month & "-" & Day)
        Return FullData
    End Function
    Public Function SiderealTimeGreenwich(ByVal Instant As DateTime, ByVal IsBC As Boolean) As DateTime
        Dim T As Double
        Dim JD As Double
        Dim InstantZero As New Date
        Dim InstantGiven As Double
        Dim Theta As Double

        InstantZero = Date.Parse(Instant.Year & "-" & Instant.Month & "-" & Instant.Day)
        JD = DateToJD(InstantZero, IsBC)

        T = (JD - 2415020.0) / 36525

        Theta = 6.6460656 + 2400.051262 * T + 0.00002581 * T ^ 2
        Theta = Theta Mod 24
        InstantGiven = MiscFunctions.HoursFromTime(Instant) * 1.002737908
        Theta += InstantGiven

        Return MiscFunctions.TimeFromHours(Theta)
    End Function
    Public Function RANutation(ByVal DeltaPsi As Double) As Double

    End Function
End Class
