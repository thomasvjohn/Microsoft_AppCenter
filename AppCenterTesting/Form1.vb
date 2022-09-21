Imports Microsoft.AppCenter
Imports Microsoft.AppCenter.Analytics
Imports Microsoft.AppCenter.Crashes
Imports System.Globalization
Imports System.IO
Imports System.Text

Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        AppCenter.SetCountryCode(RegionInfo.CurrentRegion.TwoLetterISORegionName)

        AppCenter.Start("YOUR SECRET PASSCODE", GetType(Analytics), GetType(Crashes))

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Analytics.TrackEvent("Video clicked")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            Crashes.GenerateTestCrash()
        Catch ex As Exception
            Crashes.TrackError(ex)

        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            Crashes.GenerateTestCrash()
        Catch exception As Exception
            Dim attachments = New ErrorAttachmentLog() {ErrorAttachmentLog.AttachmentWithText("MASTER_Log", "MASTER>Log"), ErrorAttachmentLog.AttachmentWithBinary(Encoding.UTF8.GetBytes(File.ReadAllText("C:\SRMS\SRMSMaster\Application\Log\2021\6\20210624_MASTER.log")), "20210624_MASTER.txt", "Log/Master")}
            Crashes.TrackError(exception, Nothing, attachments)

        End Try
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Try
            AppCenter.SetUserId("Chilton")
            Crashes.GenerateTestCrash()
        Catch exception As Exception
            Dim attachments = New ErrorAttachmentLog() {ErrorAttachmentLog.AttachmentWithBinary(Encoding.UTF8.GetBytes(File.ReadAllText("C:\SRMS\SRMSMaster\Application\Log\2021\6\20210624_MASTER.log")), "20210624_MASTER.txt", "Log")}
            Dim dict = New Dictionary(Of String, String) From {
                        {"Filename", "saved_game001.txt"},
                        {"Where", "Reload game"},
                        {"Issue", "Index of available games is corrupted"}}
            Crashes.TrackError(exception, dict, attachments)
        End Try
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Try
            AppCenter.SetUserId("Chilton")
            Crashes.GenerateTestCrash()
        Catch exception As Exception
            CopyWindowsEventsToLog("C:\SRMS\SRMSMaster\Application\Log\2021\6\20210624_MASTER.log")
            Dim attachments = New ErrorAttachmentLog() {ErrorAttachmentLog.AttachmentWithBinary(Encoding.UTF8.GetBytes(File.ReadAllText("C:\SRMS\SRMSMaster\Application\Log\2021\6\20210624_MASTER.log")), "20210624_MASTER.txt", "Log")}
            Crashes.TrackError(exception, Nothing, attachments)
        End Try

    End Sub

    Sub CopyWindowsEventsToLog(LogFile As String)

        'Opening the Log File
        Dim SW As New StreamWriter(LogFile, True)
        SW.WriteLine(vbNewLine & vbNewLine & vbNewLine & "---------------- WINDOWS EVENT LOG ----------------" & vbNewLine & vbNewLine & vbNewLine)

        Try

            'Application Event Log
            Dim elEvent As New System.Diagnostics.EventLog("Application")

            For i = elEvent.Entries.Count - 1 To 0 Step -1

                'Limiting to Last One hour
                If DateDiff(DateInterval.Minute, elEvent.Entries(i).TimeGenerated, Now) > 60 Then Exit For

                SW.WriteLine("Entry Type: " & elEvent.Entries(i).EntryType.ToString)
                SW.WriteLine("Date/Time Generated: " & elEvent.Entries(i).TimeGenerated.ToString)
                SW.WriteLine("Source: " & elEvent.Entries(i).Source.ToString)
                SW.WriteLine("Category: " & elEvent.Entries(i).Category.ToString)
                SW.WriteLine("Event: " & elEvent.Entries(i).EventID.ToString)
                If elEvent.Entries(i).UserName Is Nothing Then
                    SW.WriteLine("User: N/A")
                Else
                    SW.WriteLine("User: " & elEvent.Entries(i).UserName.ToString)
                End If
                SW.WriteLine("Computer: " & elEvent.Entries(i).MachineName.ToString)
                SW.WriteLine("Description: " & elEvent.Entries(i).Message.ToString)
                SW.WriteLine("--------------------------------------------")

            Next

            'System Event Log
            elEvent = New System.Diagnostics.EventLog("System")

            For i = elEvent.Entries.Count - 1 To 0 Step -1

                'Limiting to Last One hour
                If DateDiff(DateInterval.Minute, elEvent.Entries(i).TimeGenerated, Now) > 60 Then Exit For

                SW.WriteLine("Entry Type: " & elEvent.Entries(i).EntryType.ToString)
                SW.WriteLine("Date/Time Generated: " & elEvent.Entries(i).TimeGenerated.ToString)
                SW.WriteLine("Source: " & elEvent.Entries(i).Source.ToString)
                SW.WriteLine("Category: " & elEvent.Entries(i).Category.ToString)
                SW.WriteLine("Event: " & elEvent.Entries(i).EventID.ToString)
                If elEvent.Entries(i).UserName Is Nothing Then
                    SW.WriteLine("User: N/A")
                Else
                    SW.WriteLine("User: " & elEvent.Entries(i).UserName.ToString)
                End If
                SW.WriteLine("Computer: " & elEvent.Entries(i).MachineName.ToString)
                SW.WriteLine("Description: " & elEvent.Entries(i).Message.ToString)
                SW.WriteLine("--------------------------------------------")

            Next
        Catch ex As Exception
            Crashes.TrackError(ex)
        End Try

        SW.Close()

    End Sub

End Class