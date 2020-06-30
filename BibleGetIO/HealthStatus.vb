Imports System.Globalization
Imports System.Speech.Synthesis

Public Class HealthStatus

    Private msg As String
    Private bibleGetDB As New BibleGetDatabase
    Private state_off_text As String = "ENABLE DEBUG MODE"
    Private state_on_text As String = "DISABLE DEBUG MODE"
    Private synth As New SpeechSynthesizer
    Private DEBUG_MODE As Boolean

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Close()
    End Sub

    Private Sub HealthStatus_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        synth.Dispose()
    End Sub

    Private Sub HealthStatus_Load(sender As Object, e As EventArgs) Handles Me.Load
        DEBUG_MODE = My.Settings.DEBUG_MODE
        If DEBUG_MODE Then
            Button2.Text = state_on_text
            Button2.BackColor = Drawing.Color.Maroon
            Button2.ForeColor = Drawing.Color.Red
        Else
            Button2.Text = state_off_text
            Button2.BackColor = Drawing.Color.Lime
            Button2.ForeColor = Drawing.Color.DarkGreen
        End If

        Dim curTime As Date = DateTime.Now
        Dim dt As String = curTime.ToString("F")

        If bibleGetDB.IsInitialized Then
            msg = BibleGetAddIn.__("The BibleGet Plug-in has been correctly initialized!")
        Else
            msg = BibleGetAddIn.__("The BibleGet Plug-in has not been correctly initialized...")
        End If

        Label1.Text = msg & Environment.NewLine & Environment.NewLine & dt

        Dim speakStr As String = "<speak version=""1.0"""
        speakStr += " xmlns=""http://www.w3.org/2001/10/synthesis"""
        speakStr += " xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"""
        speakStr += " xsi:schemaLocation=""http://www.w3.org/2001/10/synthesis"
        speakStr += "           http://www.w3.org/TR/speech-synthesis/synthesis.xsd"""
        speakStr += " xml:lang=""" & CultureInfo.CurrentUICulture.Name & """>"
        speakStr += msg
        'If bibleGetDB.IsInitialized Then
        'speakStr += "The BibleGet Plug-in <prosody volume=""x-loud""> has been </prosody> <break strength=""weak"" /> correctly <break strength=""weak"" /> initialized!"
        'Else
        'speakStr += "The BibleGet Plug-in <prosody volume=""x-loud""> has </prosody> <break strength=""weak"" /> not been <break strength=""weak"" /> correctly initialized..."
        'End If
        speakStr += "</speak>"

        synth.SetOutputToDefaultAudioDevice()
        'Dim engLocale As New CultureInfo("en")
        'synth.SelectVoiceByHints(VoiceGender.Female, VoiceAge.Adult, 1, CultureInfo.CurrentUICulture)
        synth.SpeakSsmlAsync(speakStr)

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        DEBUG_MODE = Not DEBUG_MODE
        My.Settings.DEBUG_MODE = DEBUG_MODE
        My.Settings.Save()
        If DEBUG_MODE Then
            Button2.Text = state_on_text
            Button2.BackColor = Drawing.Color.Maroon
            Button2.ForeColor = Drawing.Color.Red
        Else
            Button2.Text = state_off_text
            Button2.BackColor = Drawing.Color.Lime
            Button2.ForeColor = Drawing.Color.DarkGreen
        End If
    End Sub
End Class