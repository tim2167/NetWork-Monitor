'----------------------------------------------------------------------------------------------------------------------------
'Script Name : BandwidthMonitor.sln
'Author      : Andrew Samuel
'Created     : 16/09/2010
'Description : Tracks the inbound and outbound traffic on a PC
'----------------------------------------------------------------------------------------------------------------------------
'We need to import System.Math to allow a bit of rounding
'----------------------------------------------------------------------------------------------------------------------------
Imports System.Math

Public Class Form1
    '----------------------------------------------------------------------------------------------------------------------------
    'Here we declare all our variables
    '----------------------------------------------------------------------------------------------------------------------------
    Dim objProcess As New Process               'Used to process the netstat command
    Dim strOutput As String                     'A string to store the output of the netstat process
    Dim strError As String                      'A string to store any errors of the netstat process
    Dim DLoad As Long                           'This will be the raw download data from netstat
    Dim ULoad As Long                           'This will be the raw upload data from netstat
    Dim LastULoad As Long = 0                   'This will be used to calculate the surrent usage
    Dim LastDLoad As Long = 0                   'This will be used to calculate the current usage
    Dim SString As String()                     'We need this string array to give us the ability to get the data we need from the netstat command
    Dim doAfterSecondPass As Boolean = False    'This we will use to allow the program to process something on and after the second pass through

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Timer1.Start()                          'This will start polling the netstat each second
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        'This is where the fun starts!
        '
        'We need to set all the variables inside the objProcess
        objProcess.StartInfo.RedirectStandardOutput = True
        objProcess.StartInfo.CreateNoWindow = True
        objProcess.StartInfo.RedirectStandardError = True
        objProcess.StartInfo.FileName() = "netstat"
        objProcess.StartInfo.Arguments() = "-e"
        objProcess.StartInfo.UseShellExecute = False
        'Now we will start the process
        objProcess.Start()
        'Here we will store the outputs to our 2 strings
        strOutput = objProcess.StandardOutput.ReadToEnd
        strError = objProcess.StandardError.ReadToEnd
        'And then we wait for the process to end before carrying on
        objProcess.WaitForExit()
        'Now we output the information we have
        TextBox1.Text = strOutput 'This is the raw data from netstat -e
        SString = Split(strOutput, "Bytes") 'We split that string at the word Bytes
        DLoad = Int(Mid(SString(1), 14, 17)) 'Find the download figure and store it to DLoad variable
        ULoad = Int(Mid(SString(1), 31, 16)) 'Find the upload figure and store it to ULoad variable

        'Here is where we work out how much data has passed in the last second
        If doAfterSecondPass = True Then
            Label1.Text = "Upload: " & roundBytes(ULoad - LastULoad) & "/s" & vbCrLf & "Download: " & roundBytes(DLoad - LastDLoad) & "/s"
        Else
            doAfterSecondPass = True
        End If
        'And this is where we set the las values ready to use next time
        LastDLoad = DLoad
        LastULoad = ULoad
    End Sub

    '----------------------------------------------------------------------------------------------------------------------------
    'This is the function to return a slightly friendlier value for the bytes we got from netstat
    '----------------------------------------------------------------------------------------------------------------------------
    Function roundBytes(ByVal bytes As String)
        If bytes > 1073741824 Then
            Return Round(bytes / 1073741824, 2) & " GB"
        ElseIf bytes > 1048576 Then
            Return Round(bytes / 1048576, 2) & " MB"
        ElseIf bytes > 1024 Then
            Return Round(bytes / 1024, 2) & " KB"
        ElseIf bytes = 0 Then
            Return bytes & " B"
        Else
            Return bytes & " B"
        End If
    End Function

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
    End Sub
End Class