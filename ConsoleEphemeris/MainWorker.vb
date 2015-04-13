Module MainWorker
    Public Const PromptSymbol As String = "EN > "

    Sub Main()
        Dim Statement As String
        Dim BrokenDownStatement As String()
        Dim Command As String
        Dim Args As String()
        Dim Result As String

        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Title = "Ephemeris.net command line console"

        Console.WriteLine("Welcome to Ephemeris.net console frontend")
        Console.WriteLine("This package is version " & GetType(EphemerisNet.BaseWorker).Assembly.GetName().Version.ToString)
        Console.WriteLine()
        Console.Write(PromptSymbol)

        Do While True
            Statement = Console.ReadLine()
            BrokenDownStatement = Statement.Split(" ")
            ReDim Args(BrokenDownStatement.Length - 1)
            Command = BrokenDownStatement(0)

            For i = 1 To BrokenDownStatement.Length - 1
                Args(i - 1) = BrokenDownStatement(i)
            Next

            Select Case Command
                Case "easter"
                    Result = EasterDateCalculator(Args(0))
                Case "exit", "quit"
                    Exit Do
                Case "ver"
                    Result = "This package is version " & GetType(EphemerisNet.BaseWorker).Assembly.GetName().Version.ToString
                Case Else
                    Result = "Command not acknowldged: -" & Command & "-"
            End Select
            Console.WriteLine(" " & Result)
            Console.Write(PromptSymbol)
        Loop

        Console.WriteLine("I am exiting, goodbye")
        Environment.Exit(0)

    End Sub
    Function EasterDateCalculator(InputYear As String) As String
        Dim TimeWorker As New EphemerisNet.TimeFunctions
        Dim NumericYear As Integer

        If Integer.TryParse(InputYear, NumericYear) Then
            Return "In " & InputYear & " Easter Sunday is: " & TimeWorker.EasterDate(NumericYear).ToString("yyyy-MM-dd")
        Else
            Return "Invalid year"
        End If
    End Function
End Module

