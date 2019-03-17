Imports NDde.Client
Imports System.IO
Imports System.Text

Module Module1
    Sub Main(ByVal CmdArgs() As String)
        Dim Reader As New StreamReader(AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "取得銘柄.txt", Encoding.GetEncoding("Shift-JIS"))
        Dim Writer As New StreamWriter(AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "取得内容.csv", False, Encoding.GetEncoding("Shift-JIS"))
        Dim LineCount As Integer = UBound(My.Computer.FileSystem.OpenTextFileReader(AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "取得銘柄.txt").ReadToEnd.Split(Chr(13)))
        Dim Code(LineCount) As String
        Dim Client(LineCount)
        Dim Info As String = ""

        On Error Resume Next

        For i As Integer = 0 To LineCount - 1
            Code(i) = Reader.ReadLine()
            If Code(i) = "" Then Exit Sub

            Client(i) = New DdeClient("RSS", Code(i) & ".T")
            Client(i).Connect() '接続

            For Each Item In CmdArgs
                Dim Temp As Byte() = Client(i).Request(Item, 1, 10000)
            Next
        Next

        Reader.Close()

        Do
            If File.Exists(AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "取得開始") Then Exit Do
            System.Threading.Thread.Sleep(10)
        Loop

        File.Delete(AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "取得開始")

        For i As Integer = 0 To LineCount - 1
            If i <> 0 Then Info &= vbCrLf '改行
            Info &= Code(i)
            For Each Item In CmdArgs
                Dim Temp As Byte() = Client(i).Request(Item, 1, 10000)
                Info &= "," & Encoding.Default.GetString(Temp, 0, Temp.Length - 1) '取得
            Next

            Client(i).Dispose() '切断
        Next

        Writer.WriteLine(Info)
        Writer.Close()
    End Sub

End Module
