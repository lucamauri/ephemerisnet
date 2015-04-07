Namespace OpenLM.AstroNet
	Public Class clsInput
		Dim _CalcInst As AstroTime.DateWithEra
		Dim _UniversalTime As Boolean
		Dim _PerihelionInst As AstroTime.DateWithEra
		Dim _Eccentricity As Double

		Dim _MajorAxis As Double
		Dim _MeanAnomaly As Double
		Dim _PerihelionDist As Double

		Dim _ArgPerihelion As Double
		Dim _LongAscNode As Double
		Dim _Inclination As Double
		Dim _OrbitalEpoch As Double

		Dim _CoordType As Common.CoordType
		Dim _ResultEpoch As Double

		Public ReadOnly Property CalcInst As AstroTime.DateWithEra
			Get
				Return _CalcInst
			End Get
		End Property
		Public ReadOnly Property UniversalTime As Boolean
			Get
				Return _UniversalTime
			End Get
		End Property
		Public ReadOnly Property PerihelionInst As AstroTime.DateWithEra
			Get
				Return _PerihelionInst
			End Get
		End Property
		Public ReadOnly Property Eccentricity As Double
			Get
				Return _Eccentricity
			End Get
		End Property
		Public ReadOnly Property OrbitType As String
			Get
				If _Eccentricity = 1 Then
					Return "Parabolic Orbit"
				Else
					Return "Elliptic Orbit"
				End If
			End Get
		End Property

		Public ReadOnly Property MajorAxis As Double
			Get
				Return _MajorAxis
			End Get
		End Property
		Public ReadOnly Property MeanAnomaly As Double
			Get
				Return _MeanAnomaly
			End Get
		End Property
		Public ReadOnly Property PerihelionDist As Double
			Get
				Return _PerihelionDist
			End Get
		End Property

		Public ReadOnly Property ArgPerihelion As Double
			Get
				Return _ArgPerihelion
			End Get
		End Property
		Public ReadOnly Property LongAscNode As Double
			Get
				Return _LongAscNode
			End Get
		End Property
		Public ReadOnly Property Inclination As Double
			Get
				Return _Inclination
			End Get
		End Property
		Public ReadOnly Property OrbitalEpoch As Double
			Get
				Return _OrbitalEpoch
			End Get
		End Property

		Public ReadOnly Property CoordType As Common.CoordType
			Get
				Return _CoordType
			End Get
		End Property
		Public ReadOnly Property ResultEpoch As Double
			Get
				Return _ResultEpoch
			End Get
		End Property

		Sub New(ByVal CalculationInstant As AstroTime.DateWithEra, ByVal IsUT As Boolean, ByVal PerihelionInstant As AstroTime.DateWithEra, ByVal OrbitalEccentricity As Double, ByVal MajorSemiAxis As Double, ByVal EpochMeanAnomaly As Double, ByVal PerhelionDistance As Double, ByVal ArgumentOfPerihelion As Double, ByVal LongitudeAscendingNode As Double, ByVal OrbitalInclination As Double, ByVal OrbitalDataEpoch As Double, ByVal CoordinatesType As Common.CoordType, ByVal RequestedResultEpoch As Double)

			_CalcInst = CalculationInstant
			_UniversalTime = IsUT
			_PerihelionInst = PerihelionInstant
			_Eccentricity = OrbitalEccentricity

			_MajorAxis = MajorSemiAxis
			_MeanAnomaly = EpochMeanAnomaly
			_PerihelionDist = PerhelionDistance

			_ArgPerihelion = ArgumentOfPerihelion
			_LongAscNode = LongitudeAscendingNode
			_Inclination = OrbitalInclination
			_OrbitalEpoch = OrbitalDataEpoch

			_CoordType = CoordinatesType
			_ResultEpoch = RequestedResultEpoch

		End Sub
		Function Process() As Common.Ephemeris
			Dim OrbitalFunctions As New Coord
			Dim MiscFunctions As New Common
			Dim TimeFunctions As New AstroTime

			Try
				If _Eccentricity = 1 Then
					'Parabolic Orbit
					'============
					'Programma 12
					'============
					Call OrbitalFunctions.ParabolicOrbitPosSpeed(TimeFunctions.DateToJD(_CalcInst.FullDate, _CalcInst.IsBC), TimeFunctions.DateToJD(_PerihelionInst.FullDate, _PerihelionInst.IsBC), _PerihelionDist)
				Else
					'Elliptic Orbit
					'============
					'Programma 11
					'============
					Call OrbitalFunctions.EllipticOrbitPosSpeed(TimeFunctions.DateToJD(_CalcInst.FullDate, _CalcInst.IsBC), TimeFunctions.DateToJD(_PerihelionInst.FullDate, _PerihelionInst.IsBC), _MajorAxis, _Eccentricity, _MeanAnomaly)
				End If

				'============
				'Programma 13
				'============
				Call OrbitalFunctions.OrbitalToHeliocentric(_ArgPerihelion, _Inclination, _LongAscNode, OrbitalFunctions.OrbitalPosSpeed)

				If _ResultEpoch <> _OrbitalEpoch Then
					'============
					'Programma 15 (object)
					'============
					OrbitalFunctions.HeliocentricPosSpeed = OrbitalFunctions.EclPrecession(_OrbitalEpoch, _ResultEpoch, OrbitalFunctions.HeliocentricPosSpeed)
				End If

				'============
				'Programma 14
				'============
				Call OrbitalFunctions.SunGeoPosition(TimeFunctions.DateToJD(_CalcInst.FullDate, _CalcInst.IsBC))

				If _ResultEpoch <> _MeanAnomaly Then
					'============
					'Programma 15 (Sun)
					'============
					OrbitalFunctions.SunPosSpeed = OrbitalFunctions.EclPrecession(_OrbitalEpoch, _ResultEpoch, OrbitalFunctions.SunPosSpeed)
				End If

				'============
				'Programma 17
				'============
				OrbitalFunctions.HeliocentricPosSpeed = OrbitalFunctions.AberrationCorrection(OrbitalFunctions.HeliocentricPosSpeed, OrbitalFunctions.SunPosSpeed, _CoordType)

				'============
				'Programma 18
				'============
				Call OrbitalFunctions.EclicticToEquatorial(OrbitalFunctions.HeliocentricPosSpeed, _OrbitalEpoch)

				If _CoordType = Common.CoordType.Apparent Then
					'============
					'Programma 19
					'============
					OrbitalFunctions.EquatorialPosSpeed = OrbitalFunctions.EqNutation(TimeFunctions.DateToJD(_CalcInst.FullDate, _CalcInst.IsBC), OrbitalFunctions.EquatorialPosSpeed, OrbitalFunctions.ObliquityEcliptic)
				End If

				'============
				'Programma 20
				'============
				Call OrbitalFunctions.CoordEquToPolar(OrbitalFunctions.EquatorialPosSpeed)

				'============
				'Programma 02
				'============

				Return New Common.Ephemeris(_CalcInst, OrbitalFunctions.PolarPosition.Azimuth, OrbitalFunctions.PolarPosition.Inclination, 0, OrbitalFunctions.PolarPosition.Radius, "OK")
			Catch ex As Exception
				Return New Common.Ephemeris(_CalcInst, OrbitalFunctions.PolarPosition.Azimuth, OrbitalFunctions.PolarPosition.Inclination, 0, OrbitalFunctions.PolarPosition.Radius, "OK")
			End Try

		End Function
	End Class
End Namespace

