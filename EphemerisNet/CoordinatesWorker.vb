Imports EphemerisNet.CommonResources

Public Class CoordinatesWorker
    Friend MiscFunction As New CommonResources
    Private _ObliquityEcliptic As Double

    Private _OrbitalPosSpeed As PosSpeed
    Private _HeliocentricPosSpeed As PosSpeed
    Private _SunPosSpeed As PosSpeed
    Private _EquatorialPosSpeed As PosSpeed
    Private _PolarPosition As PolarCoord

    Public ReadOnly Property OrbitalPosSpeed As PosSpeed
        Get
            Return _OrbitalPosSpeed
        End Get
    End Property
    Public Property HeliocentricPosSpeed As PosSpeed
        Get
            Return _HeliocentricPosSpeed
        End Get
        Set(ByVal value As PosSpeed)
            _HeliocentricPosSpeed = value
        End Set
    End Property
    Public Property SunPosSpeed As PosSpeed
        Get
            Return _SunPosSpeed
        End Get
        Set(ByVal value As PosSpeed)
            _SunPosSpeed = value
        End Set
    End Property
    Public Property EquatorialPosSpeed As PosSpeed
        Get
            Return _EquatorialPosSpeed
        End Get
        Set(ByVal value As PosSpeed)
            _EquatorialPosSpeed = value
        End Set
    End Property
    Public ReadOnly Property ObliquityEcliptic As Double
        Get
            Return _ObliquityEcliptic
        End Get
    End Property
    Public ReadOnly Property PolarPosition As PolarCoord
        Get
            Return _PolarPosition
        End Get
    End Property

    Private Function CoordTransformation(ByVal OriginalPosSpeed As PosSpeed, ByVal TransMatrix(,) As Double) As PosSpeed
        Dim Result As PosSpeed

        With Result
            .XPosition = TransMatrix(1, 1) * OriginalPosSpeed.XPosition + TransMatrix(1, 2) * OriginalPosSpeed.YPosition + TransMatrix(1, 3) * OriginalPosSpeed.ZPosition
            .YPosition = TransMatrix(2, 1) * OriginalPosSpeed.XPosition + TransMatrix(2, 2) * OriginalPosSpeed.YPosition + TransMatrix(2, 3) * OriginalPosSpeed.ZPosition
            .ZPosition = TransMatrix(3, 1) * OriginalPosSpeed.XPosition + TransMatrix(3, 2) * OriginalPosSpeed.YPosition + TransMatrix(3, 3) * OriginalPosSpeed.ZPosition

            .XSpeed = TransMatrix(1, 1) * OriginalPosSpeed.XSpeed + TransMatrix(1, 2) * OriginalPosSpeed.YSpeed + TransMatrix(1, 3) * OriginalPosSpeed.ZSpeed
            .YSpeed = TransMatrix(2, 1) * OriginalPosSpeed.XSpeed + TransMatrix(2, 2) * OriginalPosSpeed.YSpeed + TransMatrix(2, 3) * OriginalPosSpeed.ZSpeed
            .ZSpeed = TransMatrix(3, 1) * OriginalPosSpeed.XSpeed + TransMatrix(3, 2) * OriginalPosSpeed.YSpeed + TransMatrix(3, 3) * OriginalPosSpeed.ZSpeed
        End With
        Return Result
    End Function
    ''' <summary>
    ''' Calculates Position and Speed of an object along a elliptic orbit
    ''' </summary>
    ''' <param name="JDComputation">Julian Date of the Computeation instant</param>
    ''' <param name="JDPeriHelion">Julian Date of the object's Perihelion</param>
    ''' <param name="MajorAxis">Major Axis of the object's orbit in Astronomical Units</param>
    ''' <param name="Eccentricity">Object's orbit Eccentricity (1 > e ≥ 0)</param>
    ''' <returns>Position and speed of the object in AU and AU/day respectively in XY plane</returns>
    ''' <remarks>INTERNAL: Programma 11, page 16 - KO</remarks>
    Public Function EllipticOrbitPosSpeed(ByVal JDComputation As Double, ByVal JDPeriHelion As Double, ByVal MajorAxis As Double, ByVal Eccentricity As Double, ByVal MeanAnomaly As Single) As PosSpeed
        Dim N As Double
        Dim M As Double
        Dim E, DE, Tmp, Rho, K As Double
        'Dim Result As PosSpeed

        'Mean daily motion
        N = KGauss / (MajorAxis * Math.Sqrt(MajorAxis))
        'Mean anomaly
        'M = MiscFunction.ReduceRad(N * (JDComputation - JDPeriHelion) * MiscFunction.DegreesToRadians(MeanAnomaly))
        M = N * (JDComputation - JDPeriHelion) * MiscFunction.DegreesToRadians(MeanAnomaly)
        'Kepler Equation
        E = M * 0.85
        Do
            DE = (M + Eccentricity * Math.Sin(E) - E) / (1 - Eccentricity * Math.Cos(E))
            E += DE
        Loop Until DE > 1.0E-18

        Tmp = Math.Sqrt(1 - Eccentricity * Eccentricity)

        With _OrbitalPosSpeed
            .XPosition = MajorAxis * (Math.Cos(E) - Eccentricity)
            .YPosition = MajorAxis * Tmp * Math.Sin(E)

            Rho = 1 - Eccentricity * Math.Cos(E)
            K = KGauss / Math.Sqrt(MajorAxis)
            .XSpeed = -K * Math.Sin(E) / Rho
            .YSpeed = K * Tmp * Math.Cos(E) / Rho
        End With

        Return _OrbitalPosSpeed
    End Function
    ''' <summary>
    ''' Calculates Position and Speed of an object along a elliptic orbit
    ''' </summary>
    ''' <param name="JDComputation">Julian Date of the Computeation instant</param>
    ''' <param name="JDPerihelion">Julian Date of the object's Perihelion</param>
    ''' <param name="DistPerihelion">Perihelion distance</param>
    ''' <returns>Position and speed of the object in AU and AU/day respectively in XY plane</returns>
    ''' <remarks>INTERNAL: Programma 12/bis, page 18 - OK</remarks>
    Public Function ParabolicOrbitPosSpeed(ByVal JDComputation As Double, ByVal JDPerihelion As Double, ByVal DistPerihelion As Double) As PosSpeed
        Dim K, R, TAV2 As Double
        Dim WW, TMP As Double
        'Dim Result As PosSpeed

        'Barker's equation solved with iterative method
        WW = 3 * KGauss * (JDComputation - JDPerihelion) / (DistPerihelion * Math.Sqrt(2 * DistPerihelion))

        Do
            TMP = TAV2
            TAV2 = (2 * TMP ^ 3 + WW) / (3 * (TMP ^ 2 + 1))
        Loop While TAV2 <> TMP

        K = KGauss / Math.Sqrt(2 * DistPerihelion)
        R = DistPerihelion * (1 + TAV2 ^ 2)

        With _OrbitalPosSpeed
            .XPosition = DistPerihelion * (1 - TAV2 ^ 2)
            .YPosition = 2 * DistPerihelion * TAV2

            .XSpeed = -K * .YPosition / R
            .YSpeed = K * (.XPosition / R + 1)
        End With

        Return _OrbitalPosSpeed
    End Function
    ''' <summary>
    ''' Transform position and speed on the orbits to heliocentric eclictic coordiantes
    ''' </summary>
    ''' <param name="ArgPerihelion">Argument of Perihelion</param>
    ''' <param name="LongAscNode">Longitude of ascending node (degrees)</param>
    ''' <param name="Inclination">Orbit plane's inclination (degrees)</param>
    ''' <param name="OrbitalPosSpeed">A set of positions and speeds on the object's orbit</param>
    ''' <returns>Heliocentric position and speed of the object in AU and AU/day respectively in XYZ plane</returns>
    ''' <remarks>INTERNAL: Programma 13, page 20 - OK</remarks>
    Public Function OrbitalToHeliocentric(ByVal ArgPerihelion As Double, ByVal Inclination As Double, ByVal LongAscNode As Double, ByVal OrbitalPosSpeed As PosSpeed) As PosSpeed
        Dim MatGauss(3, 3) As Double
        Dim C1, C2, C3 As Double
        Dim S1, S2, S3 As Double
        'Dim Result As PosSpeed

        ArgPerihelion = MiscFunction.DegreesToRadians(ArgPerihelion)
        Inclination = MiscFunction.DegreesToRadians(Inclination)
        LongAscNode = MiscFunction.DegreesToRadians(LongAscNode)

        'Transformation Matrix
        C1 = Math.Cos(ArgPerihelion)
        C2 = Math.Cos(Inclination)
        C3 = Math.Cos(LongAscNode)

        S1 = Math.Sin(ArgPerihelion)
        S2 = Math.Sin(Inclination)
        S3 = Math.Sin(LongAscNode)

        MatGauss(1, 1) = C1 * C3 - S1 * C2 * S3
        MatGauss(1, 2) = -S1 * C3 - C1 * C2 * S3
        MatGauss(1, 3) = S2 * S3

        MatGauss(2, 1) = C1 * S3 + S1 * C2 * C3
        MatGauss(2, 2) = -S1 * S3 + C1 * C2 * C3
        MatGauss(2, 3) = -S2 * C3

        MatGauss(3, 1) = S1 * S2
        MatGauss(3, 2) = C1 * S2
        MatGauss(3, 3) = C2

        'Coordinates trasnformation
        _HeliocentricPosSpeed = CoordTransformation(OrbitalPosSpeed, MatGauss)

        Return _HeliocentricPosSpeed
    End Function
    ''' <summary>
    ''' Calculates Geocentric Sun position for a given instant
    ''' </summary>
    ''' <param name="CalcInst">Julian date of the instant</param>
    ''' <returns>Geocentric position and speed of the Sun</returns>
    ''' <remarks>INTERNAL: Programma 14, page 22 - OK</remarks>
    Public Function SunGeoPosition(ByVal CalcInst As Double) As PosSpeed
        Dim LA(50), RA(50) As Integer
        Dim A(50), V(50) As Double
        Dim T As Double
        Dim L, R, DL, DR As Double
        Dim U, TMP As Double
        'Dim Result As PosSpeed
        Const LongConst As Integer = 10000000

        LA = {403406, 195207, 119433, 112392, 3891, 2819, 1721, 0, 660, 350, 334, 314, 268, 242, 234, 158, 132, 129, 114, 99, 93, 86, 78, 72, 68, 64, 46, 38, 37, 32, 29, 28, 27, 27, 25, 24, 21, 21, 20, 18, 17, 14, 13, 13, 13, 12, 10, 10, 10, 10}
        RA = {0, -97597, -59715, -56188, -1556, -1126, -861, 941, -264, -163, 0, 309, -158, 0, -54, 0, -93, -20, 0, -47, 0, 0, -33, -32, 0, -10, -16, 0, 0, -24, -13, 0, -9, 0, -17, -11, 0, 31, -10, 0, -12, 0, -5, 0, 0, 0, 0, 0, 0, -9}

        A = {4.721964#, 5.937458#, 1.115589#, 5.781616#, 5.5474#, 1.512#, 4.1897#, 1.163#, 5.415#, 4.315#, 4.553#, 5.198#, 5.989#, 2.911#, 1.423#, 0.061#, 2.317#, 3.193#, 2.828#, 0.52#, 4.65#, 4.35#, 2.75#, 4.5#, 3.23#, 1.22#, 0.14#, 3.44#, 4.37#, 1.14#, 2.84#, 5.96#, 5.09#, 1.72#, 2.56#, 1.92#, 0.09#, 5.98#, 4.03#, 4.27#, 0.79#, 4.24#, 2.01#, 2.65#, 4.98#, 0.93#, 2.21#, 3.59#, 1.5#, 2.55#}
        V = {1.621043#, 62830.348067#, 62830.821524#, 62829.634302#, 125660.5691#, 125660.9845#, 62832.4766#, 0.813#, 125659.31#, 57533.85#, -33.931#, 777137.715#, 78604.191#, 5.412#, 39302.098#, -34.861#, 115067.698#, 15774.337#, 5296.67#, 58849.27#, 5296.11#, -3980.7#, 52237.6969#, 55076.47#, 261.08#, 15773.85#, 188491.03#, -7756.55#, 264.89#, 117906.27#, 55075.75#, -7961.39#, 188489.81#, 2132.19#, 109771.03#, 54868.56#, 25443.93#, -55731.43#, 60697.74#, 2132.79#, 109771.63#, -7752.82#, 188491.91#, 207.81#, 29424.63#, -7.99#, 46941.14#, -68.29#, 21463.25#, 157208.4#}

        T = (CalcInst - 2451545) / 3652500
        For i As Integer = 0 To 49
            U = MiscFunction.ReduceRad(A(i) + V(i) * T)
            L = L + LA(i) * Math.Sin(U)
            R = R + RA(i) * Math.Cos(U)
            DL = DL + LA(i) * V(i) * Math.Cos(U)
            DR = DR - RA(i) * V(i) * Math.Sin(U)
        Next

        TMP = MiscFunction.ReduceRad(62833.196168 * T)
        L = MiscFunction.ReduceRad(4.9353929 + TMP + L / LongConst)
        R = 1.0001026 + R / LongConst
        DL = (62833.196168 + DL / LongConst) / 3652500
        DR = (DR / LongConst) / 3652500

        With _SunPosSpeed
            .XPosition = R * Math.Cos(L)
            .YPosition = R * Math.Sin(L)

            .XSpeed = DR * Math.Cos(L) - DL * .YPosition
            .YSpeed = DR * Math.Sin(L) + DL * .XPosition

        End With
        Return _SunPosSpeed
    End Function
    ''' <summary>
    ''' Convert eclictic coordinates from one epoch to another calcultaing precession
    ''' </summary>
    ''' <param name="InitInstant">Initial Epoch or Julian date (see also EpochB1950 and EpochJ2000 constants)</param>
    ''' <param name="FinalInstant">Final Epoch or Julian date (see also EpochB1950 and EpochJ2000 constants)</param>
    ''' <param name="InitPosition">Initial Position and Speed</param>
    ''' <returns>An object representing postition and speed corrected for the precession</returns>
    ''' <remarks>INTERNAL: Programma 15, page 24 - OK</remarks>
    Public Function EclPrecession(ByVal InitInstant As Double, ByVal FinalInstant As Double, ByVal InitPosition As PosSpeed) As PosSpeed
        Dim T0, T, DT As Double
        Dim LgNodEcl, AngoEcli, Preces As Double
        Dim MatPrec(3, 3) As Double
        Dim C1, C2, C3 As Double
        Dim S1, S2, S3 As Double
        Dim Result As PosSpeed

        T0 = (InitInstant - 2451545) / 36525
        T = (FinalInstant - 2451545) / 36525
        DT = T - T0

        LgNodEcl = 174.876383888889 + (((3289.4789 + 0.60622 * T0) * T0) + ((-869.8089 - 0.50491 * T0) + 0.03536 * DT) * DT) / 3600
        AngoEcli = ((47.0029 - (0.06603 - 0.000598 * T0) * T0) + ((-0.03302 + 0.000598 * T0) + 0.00006 * DT) * DT) * DT / 3600
        Preces = ((5029.0966 + (2.22226 - 0.000042 * T0) * T0) + ((1.11113 - 0.000042 * T0) - 0.000006 * DT) * DT) * DT / 3600

        LgNodEcl = MiscFunction.DegreesToRadians(LgNodEcl)
        AngoEcli = MiscFunction.DegreesToRadians(AngoEcli)
        Preces = MiscFunction.DegreesToRadians(Preces)

        C1 = Math.Cos(LgNodEcl + Preces)
        C2 = Math.Cos(AngoEcli)
        C3 = Math.Cos(LgNodEcl)

        S1 = Math.Sin(LgNodEcl + Preces)
        S2 = Math.Sin(AngoEcli)
        S3 = Math.Sin(LgNodEcl)

        MatPrec(1, 1) = C1 * C3 + S1 * C2 * S3
        MatPrec(1, 2) = C1 * S3 - S1 * C2 * C3
        MatPrec(1, 3) = -S1 * S2

        MatPrec(2, 1) = S1 * C3 - C1 * C2 * S3
        MatPrec(2, 2) = S1 * S3 + C1 * C2 * C3
        MatPrec(2, 3) = C1 * S2

        MatPrec(3, 1) = S2 * S3
        MatPrec(3, 2) = -S2 * C3
        MatPrec(3, 3) = C2

        'Coordinates trasnformation
        Result = CoordTransformation(InitPosition, MatPrec)

        Return Result
    End Function
    ''' <summary>
    ''' Correct postition of an object for Aberration and transform Heliocentric Eclictc Coordinates to Gecoentric Ecliptic Coordinates
    ''' </summary>
    ''' <param name="ObjPosSpeed"></param>
    ''' <param name="SunPosSpeed"></param>
    ''' <param name="RequestedType"></param>
    ''' <returns></returns>
    ''' <remarks>INTERNAL: Programma 17, page 28 - Controlla la precisione del risultato</remarks>
    Public Function AberrationCorrection(ByVal ObjPosSpeed As PosSpeed, ByVal SunPosSpeed As PosSpeed, ByVal RequestedType As CoordType) As PosSpeed
        Dim X, Y, Z As Double
        Dim Delta, LightTime As Double
        Dim Result As PosSpeed

        X = ObjPosSpeed.XPosition + SunPosSpeed.XPosition
        Y = ObjPosSpeed.YPosition + SunPosSpeed.YPosition
        Z = ObjPosSpeed.ZPosition + SunPosSpeed.ZPosition

        Delta = Math.Sqrt(X ^ 2 + Y ^ 2 + Z ^ 2)
        LightTime = 0.0057755183041213 * Delta

        With Result
            Select Case RequestedType
                Case CoordType.Geometric
                    .XPosition = X
                    .YPosition = Y
                    .ZPosition = Z
                Case CoordType.Astrometric
                    .XPosition = X - LightTime * ObjPosSpeed.XSpeed
                    .YPosition = Y - LightTime * ObjPosSpeed.YSpeed
                    .ZPosition = Z - LightTime * ObjPosSpeed.ZSpeed
                Case CoordType.Apparent
                    .XPosition = X - LightTime * SunPosSpeed.XSpeed + ObjPosSpeed.XSpeed
                    .YPosition = Y - LightTime * SunPosSpeed.YSpeed + ObjPosSpeed.YSpeed
                    .ZPosition = Z - LightTime * SunPosSpeed.ZSpeed + ObjPosSpeed.ZSpeed
            End Select
        End With
        Return Result
    End Function
    ''' <summary>
    ''' Calculates the mean Obliquity of the ecliptic value and transform Gecoentric Ecliptic Coordinates to Geocentric Equatorial Coordinates
    ''' </summary>
    ''' <param name="EclPosition">Gecoentric Ecliptic Coordinates</param>
    ''' <param name="RefEpoch">Final Epoch or Julian date (see also EpochB1950 and EpochJ2000 constants)</param>
    ''' <returns>Geocentric Equatorial Coordinates</returns>
    ''' <remarks>INTERNAL: Programma 18, page 30 - OK</remarks>
    Public Function EclicticToEquatorial(ByVal EclPosition As PosSpeed, ByVal RefEpoch As Double) As PosSpeed
        Dim T, Tmp, Epsilon As Double
        Dim CE, SE As Double
        'Dim Result As PosSpeed

        T = (RefEpoch - 2451545) / 36525
        T /= 100

        Tmp = T * (27.87 + T * (5.79 + T * 2.45))
        Tmp = T * (-249.67 + T * (-39.05 + T * (7.12 + Tmp)))
        Tmp = T * (-1.55 + T * (1999.25 + T * (-51.38 + Tmp)))
        Tmp = (T * (-4680.93 + Tmp)) / 3600
        'Obliquity of the ecliptic in degrees
        Epsilon = OblEclJ2000 + Tmp
        'Obliquity of the ecliptic in radians
        Epsilon = MiscFunction.DegreesToRadians(Epsilon)
        _ObliquityEcliptic = Epsilon

        CE = Math.Cos(Epsilon)
        SE = Math.Sin(Epsilon)

        With _EquatorialPosSpeed
            .XPosition = EclPosition.XPosition
            .YPosition = EclPosition.YPosition * CE - EclPosition.ZPosition * SE
            .ZPosition = EclPosition.YPosition * SE + EclPosition.ZPosition * CE
        End With

        Return _EquatorialPosSpeed
    End Function
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="CalcInstant"></param>
    ''' <param name="EqPosition"></param>
    ''' <param name="Epsilon"></param>
    ''' <returns></returns>
    ''' <remarks>INTERNAL: Programma 19, page 32 - ??</remarks>
    Public Function EqNutation(ByVal CalcInstant As Double, ByVal EqPosition As PosSpeed, ByVal Epsilon As Double) As PosSpeed
        Dim T As Double
        Dim LML, AML, LNA, LMS, AMS As Double
        Dim DPsi, DEpsi As Double
        Dim CTMP, STMP As Double
        Dim Result As PosSpeed

        T = (CalcInstant - 2451545) / 36525
        'Mean Moon longitude
        LML = MiscFunction.DegreesToRadiansCorrect(218.316 + 481267.881 * T)
        'Mean Moon Anomaly
        AML = MiscFunction.DegreesToRadiansCorrect(134.963 + 477198.867 * T)
        'Ascending Node Moon Longitude
        LNA = MiscFunction.DegreesToRadiansCorrect(125.045 - 1934.136 * T)
        'Mean Sun Longitude
        LMS = MiscFunction.DegreesToRadiansCorrect(280.466 + 36000.77 * T)
        'Mean Sun Anomaly
        AMS = MiscFunction.DegreesToRadiansCorrect(357.528 + 35999.05 * T)

        'Longitude Nutation
        DPsi = -17.2 * Math.Sin(LNA) + 0.206 * Math.Sin(2 * LNA) - 1.309 * Math.Sin(2 * LMS) + 0.143 * Math.Sin(AMS) - 0.227 * Math.Sin(2 * LML) + 0.071 * Math.Sin(AML)
        DPsi = MiscFunction.DegreesToRadians(DPsi / 3600)
        'Obliquity Nutation
        DEpsi = 9.203 * Math.Cos(LNA) - 0.09 * Math.Cos(2 * LNA) + 0.574 * Math.Cos(2 * LMS) + 0.022 * Math.Cos(2 * LMS + AMS) + 0.098 * Math.Cos(2 * LML) + 0.02 * Math.Cos(2 * LML - LNA)
        DEpsi = MiscFunction.DegreesToRadians(DEpsi / 3600)

        CTMP = DPsi * Math.Cos(Epsilon)
        STMP = DPsi * Math.Sin(Epsilon)

        With EqPosition
            Result.XPosition = .XPosition + (-CTMP * .YPosition - STMP * .ZPosition)
            Result.YPosition = .YPosition + (CTMP * .XPosition - DEpsi * .ZPosition)
            Result.ZPosition = .ZPosition + (STMP * .XPosition + DEpsi * .YPosition)
        End With

        Return Result
    End Function
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="GeoCentricEqua"></param>
    ''' <returns></returns>
    ''' <remarks>INTERNAL: Programma 20, page 34 - ??</remarks>
    Public Function CoordEquToPolar(ByVal GeoCentricEqua As PosSpeed) As PolarCoord
        Dim R As Double
        Dim RA, Decl As Double
        'Dim Result As PolarCoord

        With GeoCentricEqua
            R = Math.Sqrt(.XPosition ^ 2 + .YPosition ^ 2 + .ZPosition ^ 2)
            RA = MiscFunction.ReduceRad(Math.Atan2(.YPosition, .XPosition))
            'Decl = Math.Atan(.ZPosition / Math.Sqrt(.XPosition ^ 2 + .YPosition ^ 2))
            Decl = Math.Acos(.ZPosition / R)
        End With

        With _PolarPosition
            .Radius = R
            .Inclination = RA
            .Azimuth = Decl
        End With

        Return _PolarPosition
    End Function
End Class
