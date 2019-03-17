Imports NDde.Client
Imports System.IO
Imports System.Text

Module Module1

    Sub Main(ByVal CmdArgs() As String)
        Dim Reader As New StreamReader(AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "取得銘柄.txt", Encoding.GetEncoding("Shift-JIS"))
        Dim Writer As New StreamWriter(AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "取得内容.csv", True, Encoding.GetEncoding("Shift-JIS"))
        Dim ErrorW As New StreamWriter(AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "エラー銘柄.txt", False, Encoding.GetEncoding("Shift-JIS"))
        Dim LineCount As Integer = UBound(My.Computer.FileSystem.OpenTextFileReader(AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "取得銘柄.txt").ReadToEnd.Split(Chr(13)))
        Dim LoopCount As Integer = 0
        Dim Info As String = ""
        Dim ErrorCode As String = ""
        Dim FirstFlag As Boolean = True
        Dim Temp As Byte()

        Do Until Reader.EndOfStream
            Dim Code As String = Reader.ReadLine()
            If Code = "" Then Exit Do
            Dim Client As New DdeClient("RSS", Code & ".T")
            Client.Connect() '接続

            Console.Write(vbCr & "進行状況：" & Math.Round(LoopCount / LineCount * 100) & "% (" & Code & ")") '進行状況

            Try
                Temp = Client.Request("銘柄コード", 1, 10000)

                Info &= Code
                For Each Item In CmdArgs
                    Temp = Client.Request(Item, 1, 10000)
                    Info &= "," & Encoding.Default.GetString(Temp, 0, Temp.Length - 1) '取得
                Next
                Info &= vbCrLf '改行
            Catch ex As NDde.DdeException
                If FirstFlag Then
                    FirstFlag = False
                    ErrorCode = Code
                Else
                    ErrorCode &= vbCrLf & Code
                End If
            End Try

            Client.Dispose() '切断
            LoopCount += 1
        Loop

        Writer.WriteLine(Info)
        ErrorW.WriteLine(ErrorCode)
        Reader.Close()
        Writer.Close()
        ErrorW.Close()
    End Sub

End Module
