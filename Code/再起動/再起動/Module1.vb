Module Module1

    Sub Main()
        Dim PS As Process() = Process.GetProcessesByName("Excel")
        Dim P

        For Each P In PS
            P.WaitForExit()
        Next

        Process.Start(Command())
    End Sub

End Module
