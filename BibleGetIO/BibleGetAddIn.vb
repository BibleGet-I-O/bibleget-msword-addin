Imports System.Globalization
Imports System.IO
Imports System.Management


Public Class BibleGetAddIn

    Private Shared _RM As Resources.ResourceManager = New Resources.ResourceManager("BibleGetIO.BibleGetResource", System.Reflection.Assembly.GetExecutingAssembly())
    Private Shared _locale As CultureInfo = CultureInfo.CurrentUICulture
    'Public Shared helpFile As String    
    Public Shared ThisAppDataHome As String = "BibleGetMSOfficePlugin"
    Public Shared ThisAppDataDirectory As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), ThisAppDataHome)
    Public Shared logFile As String = Path.Combine(ThisAppDataDirectory, "BibleGet.log")
    Public Shared BGET_ENDPOINT = "https://query.bibleget.io/index.php"
    Public Shared BGET_METADATA_ENDPOINT = "https://query.bibleget.io/metadata.php"
    Public Shared BGET_SEARCH_ENDPOINT = "https://query.bibleget.io/search.php"
    'Private DEBUG_MODE = My.Settings.DEBUG_MODE

    Shared ReadOnly Property RM As Resources.ResourceManager
        Get
            Return _RM
        End Get
    End Property

    Shared ReadOnly Property locale As CultureInfo
        Get
            Return _locale
        End Get
    End Property


    Public Shared Function __(ByVal myStr As String) As String
        Dim myTranslation As String = _RM.GetString(myStr, _locale)
        If Not String.IsNullOrEmpty(myTranslation) Then
            Return myTranslation
        Else
            Return myStr
        End If
    End Function

    Private Shared Sub BibleGetAddIn_Startup() Handles Me.Startup
        Dim DEBUG_MODE As Boolean = My.Settings.DEBUG_MODE
        Dim Application As Word.Application = Globals.BibleGetAddIn.Application
        Dim lang As Office.MsoLanguageID = Application.LanguageSettings.LanguageID(Office.MsoAppLanguageID.msoLanguageIDUI)
        Threading.Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo(lang)
        Threading.Thread.CurrentThread.CurrentUICulture = CultureInfo.GetCultureInfo(lang)
        _locale = CultureInfo.GetCultureInfo(lang)
        'Select Case lang
        '    Case Office.MsoLanguageID.msoLanguageIDAfrikaans

        '    Case Office.MsoLanguageID.msoLanguageIDAlbanian
        '    Case Office.MsoLanguageID.msoLanguageIDAmharic
        '    Case Office.MsoLanguageID.msoLanguageIDArabic
        '    Case Office.MsoLanguageID.msoLanguageIDArabicAlgeria
        '    Case Office.MsoLanguageID.msoLanguageIDArabicBahrain
        '    Case Office.MsoLanguageID.msoLanguageIDArabicEgypt
        '    Case Office.MsoLanguageID.msoLanguageIDArabicIraq
        '    Case Office.MsoLanguageID.msoLanguageIDArabicJordan
        '    Case Office.MsoLanguageID.msoLanguageIDArabicKuwait
        '    Case Office.MsoLanguageID.msoLanguageIDArabicLebanon
        '    Case Office.MsoLanguageID.msoLanguageIDArabicLibya
        '    Case Office.MsoLanguageID.msoLanguageIDArabicMorocco
        '    Case Office.MsoLanguageID.msoLanguageIDArabicOman
        '    Case Office.MsoLanguageID.msoLanguageIDArabicQatar
        '    Case Office.MsoLanguageID.msoLanguageIDArabicSyria
        '    Case Office.MsoLanguageID.msoLanguageIDArabicTunisia
        '    Case Office.MsoLanguageID.msoLanguageIDArabicUAE
        '    Case Office.MsoLanguageID.msoLanguageIDArabicYemen
        '    Case Office.MsoLanguageID.msoLanguageIDArmenian
        '    Case Office.MsoLanguageID.msoLanguageIDAssamese
        '    Case Office.MsoLanguageID.msoLanguageIDAzeriCyrillic
        '    Case Office.MsoLanguageID.msoLanguageIDAzeriLatin
        '    Case Office.MsoLanguageID.msoLanguageIDBasque
        '    Case Office.MsoLanguageID.msoLanguageIDBelgianDutch
        '    Case Office.MsoLanguageID.msoLanguageIDBelgianFrench
        '    Case Office.MsoLanguageID.msoLanguageIDBengali
        '    Case Office.MsoLanguageID.msoLanguageIDBosnian
        '    Case Office.MsoLanguageID.msoLanguageIDBosnianBosniaHerzegovinaCyrillic
        '    Case Office.MsoLanguageID.msoLanguageIDBosnianBosniaHerzegovinaLatin
        '    Case Office.MsoLanguageID.msoLanguageIDBrazilianPortuguese
        '    Case Office.MsoLanguageID.msoLanguageIDBulgarian
        '    Case Office.MsoLanguageID.msoLanguageIDBurmese
        '    Case Office.MsoLanguageID.msoLanguageIDByelorussian
        '    Case Office.MsoLanguageID.msoLanguageIDCatalan
        '    Case Office.MsoLanguageID.msoLanguageIDCherokee
        '    Case Office.MsoLanguageID.msoLanguageIDChineseHongKongSAR
        '    Case Office.MsoLanguageID.msoLanguageIDChineseMacaoSAR
        '    Case Office.MsoLanguageID.msoLanguageIDChineseSingapore
        '    Case Office.MsoLanguageID.msoLanguageIDCroatian
        '    Case Office.MsoLanguageID.msoLanguageIDCzech
        '    Case Office.MsoLanguageID.msoLanguageIDDanish
        '    Case Office.MsoLanguageID.msoLanguageIDDivehi
        '    Case Office.MsoLanguageID.msoLanguageIDDutch
        '    Case Office.MsoLanguageID.msoLanguageIDEdo
        '    Case Office.MsoLanguageID.msoLanguageIDEnglishAUS
        '    Case Office.MsoLanguageID.msoLanguageIDEnglishBelize
        '    Case Office.MsoLanguageID.msoLanguageIDEnglishCanadian
        '    Case Office.MsoLanguageID.msoLanguageIDEnglishCaribbean
        '    Case Office.MsoLanguageID.msoLanguageIDEnglishIndonesia
        '    Case Office.MsoLanguageID.msoLanguageIDEnglishIreland
        '    Case Office.MsoLanguageID.msoLanguageIDEnglishJamaica
        '    Case Office.MsoLanguageID.msoLanguageIDEnglishNewZealand
        '    Case Office.MsoLanguageID.msoLanguageIDEnglishPhilippines
        '    Case Office.MsoLanguageID.msoLanguageIDEnglishSouthAfrica
        '    Case Office.MsoLanguageID.msoLanguageIDEnglishTrinidadTobago
        '    Case Office.MsoLanguageID.msoLanguageIDEnglishUK
        '    Case Office.MsoLanguageID.msoLanguageIDEnglishUS
        '    Case Office.MsoLanguageID.msoLanguageIDEnglishZimbabwe
        '    Case Office.MsoLanguageID.msoLanguageIDEstonian
        '    Case Office.MsoLanguageID.msoLanguageIDFaeroese
        '    Case Office.MsoLanguageID.msoLanguageIDFarsi
        '    Case Office.MsoLanguageID.msoLanguageIDFilipino
        '    Case Office.MsoLanguageID.msoLanguageIDFinnish
        '    Case Office.MsoLanguageID.msoLanguageIDFrench
        '    Case Office.MsoLanguageID.msoLanguageIDFrenchCameroon
        '    Case Office.MsoLanguageID.msoLanguageIDFrenchCanadian
        '    Case Office.MsoLanguageID.msoLanguageIDFrenchCotedIvoire
        '    Case Office.MsoLanguageID.msoLanguageIDFrenchHaiti
        '    Case Office.MsoLanguageID.msoLanguageIDFrenchLuxembourg
        '    Case Office.MsoLanguageID.msoLanguageIDFrenchMali
        '    Case Office.MsoLanguageID.msoLanguageIDFrenchMonaco
        '    Case Office.MsoLanguageID.msoLanguageIDFrenchMorocco
        '    Case Office.MsoLanguageID.msoLanguageIDFrenchReunion
        '    Case Office.MsoLanguageID.msoLanguageIDFrenchSenegal
        '    Case Office.MsoLanguageID.msoLanguageIDFrenchWestIndies
        '    Case Office.MsoLanguageID.msoLanguageIDFrenchCongoDRC
        '    Case Office.MsoLanguageID.msoLanguageIDFrisianNetherlands
        '    Case Office.MsoLanguageID.msoLanguageIDFulfulde
        '    Case Office.MsoLanguageID.msoLanguageIDGaelicIreland
        '    Case Office.MsoLanguageID.msoLanguageIDGaelicScotland
        '    Case Office.MsoLanguageID.msoLanguageIDGalician
        '    Case Office.MsoLanguageID.msoLanguageIDGeorgian
        '    Case Office.MsoLanguageID.msoLanguageIDGerman
        '    Case Office.MsoLanguageID.msoLanguageIDGermanAustria
        '    Case Office.MsoLanguageID.msoLanguageIDGermanLiechtenstein
        '    Case Office.MsoLanguageID.msoLanguageIDGermanLuxembourg
        '    Case Office.MsoLanguageID.msoLanguageIDGreek
        '    Case Office.MsoLanguageID.msoLanguageIDGuarani
        '    Case Office.MsoLanguageID.msoLanguageIDGujarati
        '    Case Office.MsoLanguageID.msoLanguageIDHausa
        '    Case Office.MsoLanguageID.msoLanguageIDHawaiian
        '    Case Office.MsoLanguageID.msoLanguageIDHebrew
        '    Case Office.MsoLanguageID.msoLanguageIDHindi
        '    Case Office.MsoLanguageID.msoLanguageIDHungarian
        '    Case Office.MsoLanguageID.msoLanguageIDIbibio
        '    Case Office.MsoLanguageID.msoLanguageIDIcelandic
        '    Case Office.MsoLanguageID.msoLanguageIDIgbo
        '    Case Office.MsoLanguageID.msoLanguageIDIndonesian
        '    Case Office.MsoLanguageID.msoLanguageIDInuktitut
        '    Case Office.MsoLanguageID.msoLanguageIDItalian
        '    Case Office.MsoLanguageID.msoLanguageIDJapanese
        '    Case Office.MsoLanguageID.msoLanguageIDKannada
        '    Case Office.MsoLanguageID.msoLanguageIDKanuri
        '    Case Office.MsoLanguageID.msoLanguageIDKashmiri
        '    Case Office.MsoLanguageID.msoLanguageIDKashmiriDevanagari
        '    Case Office.MsoLanguageID.msoLanguageIDKazakh
        '    Case Office.MsoLanguageID.msoLanguageIDKhmer
        '    Case Office.MsoLanguageID.msoLanguageIDKirghiz
        '    Case Office.MsoLanguageID.msoLanguageIDKonkani
        '    Case Office.MsoLanguageID.msoLanguageIDKorean
        '    Case Office.MsoLanguageID.msoLanguageIDKyrgyz
        '    Case Office.MsoLanguageID.msoLanguageIDLao
        '    Case Office.MsoLanguageID.msoLanguageIDLatin
        '    Case Office.MsoLanguageID.msoLanguageIDLatvian
        '    Case Office.MsoLanguageID.msoLanguageIDLithuanian
        '    Case Office.MsoLanguageID.msoLanguageIDMacedonianFYROM
        '    Case Office.MsoLanguageID.msoLanguageIDMalayalam
        '    Case Office.MsoLanguageID.msoLanguageIDMalayBruneiDarussalam
        '    Case Office.MsoLanguageID.msoLanguageIDMalaysian
        '    Case Office.MsoLanguageID.msoLanguageIDMaltese
        '    Case Office.MsoLanguageID.msoLanguageIDManipuri
        '    Case Office.MsoLanguageID.msoLanguageIDMaori
        '    Case Office.MsoLanguageID.msoLanguageIDMarathi
        '    Case Office.MsoLanguageID.msoLanguageIDMexicanSpanish
        '    Case Office.MsoLanguageID.msoLanguageIDMixed
        '    Case Office.MsoLanguageID.msoLanguageIDMongolian
        '    Case Office.MsoLanguageID.msoLanguageIDNepali
        '    Case Office.MsoLanguageID.msoLanguageIDNone
        '    Case Office.MsoLanguageID.msoLanguageIDNoProofing
        '    Case Office.MsoLanguageID.msoLanguageIDNorwegianBokmol
        '    Case Office.MsoLanguageID.msoLanguageIDNorwegianNynorsk
        '    Case Office.MsoLanguageID.msoLanguageIDOriya
        '    Case Office.MsoLanguageID.msoLanguageIDOromo
        '    Case Office.MsoLanguageID.msoLanguageIDPashto
        '    Case Office.MsoLanguageID.msoLanguageIDPolish
        '    Case Office.MsoLanguageID.msoLanguageIDPortuguese
        '    Case Office.MsoLanguageID.msoLanguageIDPunjabi
        '    Case Office.MsoLanguageID.msoLanguageIDQuechuaBolivia
        '    Case Office.MsoLanguageID.msoLanguageIDQuechuaEcuador
        '    Case Office.MsoLanguageID.msoLanguageIDQuechuaPeru
        '    Case Office.MsoLanguageID.msoLanguageIDRhaetoRomanic
        '    Case Office.MsoLanguageID.msoLanguageIDRomanian
        '    Case Office.MsoLanguageID.msoLanguageIDRomanianMoldova
        '    Case Office.MsoLanguageID.msoLanguageIDRussian
        '    Case Office.MsoLanguageID.msoLanguageIDRussianMoldova
        '    Case Office.MsoLanguageID.msoLanguageIDSamiLappish
        '    Case Office.MsoLanguageID.msoLanguageIDSanskrit
        '    Case Office.MsoLanguageID.msoLanguageIDSepedi
        '    Case Office.MsoLanguageID.msoLanguageIDSerbianBosniaHerzegovinaCyrillic
        '    Case Office.MsoLanguageID.msoLanguageIDSerbianBosniaHerzegovinaLatin
        '    Case Office.MsoLanguageID.msoLanguageIDSerbianCyrillic
        '    Case Office.MsoLanguageID.msoLanguageIDSerbianLatin
        '    Case Office.MsoLanguageID.msoLanguageIDSesotho
        '    Case Office.MsoLanguageID.msoLanguageIDSimplifiedChinese
        '    Case Office.MsoLanguageID.msoLanguageIDSindhi
        '    Case Office.MsoLanguageID.msoLanguageIDSindhiPakistan
        '    Case Office.MsoLanguageID.msoLanguageIDSinhalese
        '    Case Office.MsoLanguageID.msoLanguageIDSlovak
        '    Case Office.MsoLanguageID.msoLanguageIDSlovenian
        '    Case Office.MsoLanguageID.msoLanguageIDSomali
        '    Case Office.MsoLanguageID.msoLanguageIDSorbian
        '    Case Office.MsoLanguageID.msoLanguageIDSpanish
        '    Case Office.MsoLanguageID.msoLanguageIDSpanishArgentina
        '    Case Office.MsoLanguageID.msoLanguageIDSpanishBolivia
        '    Case Office.MsoLanguageID.msoLanguageIDSpanishChile
        '    Case Office.MsoLanguageID.msoLanguageIDSpanishColombia
        '    Case Office.MsoLanguageID.msoLanguageIDSpanishCostaRica
        '    Case Office.MsoLanguageID.msoLanguageIDSpanishDominicanRepublic
        '    Case Office.MsoLanguageID.msoLanguageIDSpanishEcuador
        '    Case Office.MsoLanguageID.msoLanguageIDSpanishElSalvador
        '    Case Office.MsoLanguageID.msoLanguageIDSpanishGuatemala
        '    Case Office.MsoLanguageID.msoLanguageIDSpanishHonduras
        '    Case Office.MsoLanguageID.msoLanguageIDSpanishModernSort
        '    Case Office.MsoLanguageID.msoLanguageIDSpanishNicaragua
        '    Case Office.MsoLanguageID.msoLanguageIDSpanishPanama
        '    Case Office.MsoLanguageID.msoLanguageIDSpanishParaguay
        '    Case Office.MsoLanguageID.msoLanguageIDSpanishPeru
        '    Case Office.MsoLanguageID.msoLanguageIDSpanishPuertoRico
        '    Case Office.MsoLanguageID.msoLanguageIDSpanishUruguay
        '    Case Office.MsoLanguageID.msoLanguageIDSpanishVenezuela
        '    Case Office.MsoLanguageID.msoLanguageIDSutu
        '    Case Office.MsoLanguageID.msoLanguageIDSwahili
        '    Case Office.MsoLanguageID.msoLanguageIDSwedish
        '    Case Office.MsoLanguageID.msoLanguageIDSwedishFinland
        '    Case Office.MsoLanguageID.msoLanguageIDSwissFrench
        '    Case Office.MsoLanguageID.msoLanguageIDSwissGerman
        '    Case Office.MsoLanguageID.msoLanguageIDSwissItalian
        '    Case Office.MsoLanguageID.msoLanguageIDSyriac
        '    Case Office.MsoLanguageID.msoLanguageIDTajik
        '    Case Office.MsoLanguageID.msoLanguageIDTamazight
        '    Case Office.MsoLanguageID.msoLanguageIDTamazightLatin
        '    Case Office.MsoLanguageID.msoLanguageIDTamil
        '    Case Office.MsoLanguageID.msoLanguageIDTatar
        '    Case Office.MsoLanguageID.msoLanguageIDTelugu
        '    Case Office.MsoLanguageID.msoLanguageIDThai
        '    Case Office.MsoLanguageID.msoLanguageIDTibetan
        '    Case Office.MsoLanguageID.msoLanguageIDTigrignaEritrea
        '    Case Office.MsoLanguageID.msoLanguageIDTigrignaEthiopic
        '    Case Office.MsoLanguageID.msoLanguageIDTraditionalChinese
        '    Case Office.MsoLanguageID.msoLanguageIDTsonga
        '    Case Office.MsoLanguageID.msoLanguageIDTswana
        '    Case Office.MsoLanguageID.msoLanguageIDTurkish
        '    Case Office.MsoLanguageID.msoLanguageIDTurkmen
        '    Case Office.MsoLanguageID.msoLanguageIDUkrainian
        '    Case Office.MsoLanguageID.msoLanguageIDUrdu
        '    Case Office.MsoLanguageID.msoLanguageIDUzbekCyrillic
        '    Case Office.MsoLanguageID.msoLanguageIDUzbekLatin
        '    Case Office.MsoLanguageID.msoLanguageIDVenda
        '    Case Office.MsoLanguageID.msoLanguageIDVietnamese
        '    Case Office.MsoLanguageID.msoLanguageIDWelsh
        '    Case Office.MsoLanguageID.msoLanguageIDXhosa
        '    Case Office.MsoLanguageID.msoLanguageIDYi
        '    Case Office.MsoLanguageID.msoLanguageIDYiddish
        '    Case Office.MsoLanguageID.msoLanguageIDYoruba
        '    Case Office.MsoLanguageID.msoLanguageIDZulu
        'End Select

        If Not Directory.Exists(ThisAppDataDirectory) Then
            Try
                Dim dirInfo As DirectoryInfo
                dirInfo = Directory.CreateDirectory(ThisAppDataDirectory)
                Diagnostics.Debug.WriteLine(TimeOfDay.ToLongTimeString & " >> Directory created successfully: " & dirInfo.ToString)
            Catch ex As UnauthorizedAccessException
                Diagnostics.Debug.WriteLine("UnauthorizedAccessException caught while trying to create directory: " & ThisAppDataDirectory)
                Diagnostics.Debug.WriteLine(ex.Message)
            Catch ex As ArgumentNullException
                Diagnostics.Debug.WriteLine("UnauthorizedAccessException caught while trying to create directory: " & ThisAppDataDirectory)
                Diagnostics.Debug.WriteLine(ex.Message)
            Catch ex As ArgumentException
                Diagnostics.Debug.WriteLine("UnauthorizedAccessException caught while trying to create directory: " & ThisAppDataDirectory)
                Diagnostics.Debug.WriteLine(ex.Message)
            Catch ex As PathTooLongException
                Diagnostics.Debug.WriteLine("UnauthorizedAccessException caught while trying to create directory: " & ThisAppDataDirectory)
                Diagnostics.Debug.WriteLine(ex.Message)
            Catch ex As DirectoryNotFoundException
                Diagnostics.Debug.WriteLine("UnauthorizedAccessException caught while trying to create directory: " & ThisAppDataDirectory)
                Diagnostics.Debug.WriteLine(ex.Message)
            Catch ex As IOException
                Diagnostics.Debug.WriteLine("UnauthorizedAccessException caught while trying to create directory: " & ThisAppDataDirectory)
                Diagnostics.Debug.WriteLine(ex.Message)
            Catch ex As NotSupportedException
                Diagnostics.Debug.WriteLine("UnauthorizedAccessException caught while trying to create directory: " & ThisAppDataDirectory)
                Diagnostics.Debug.WriteLine(ex.Message)
            End Try

        End If
        If Not File.Exists(logFile) Then
            Dim objOS As ManagementObjectSearcher
            Dim objCs As ManagementObjectSearcher
            Dim objMgmt As ManagementObject
            Dim m_strComputerName As String = String.Empty
            Dim m_strManufacturer As String = String.Empty
            Dim m_StrModel As String = String.Empty
            Dim m_strOSName As String = String.Empty
            Dim m_strOSVersion As String = String.Empty
            Dim m_strSystemType As String = String.Empty
            Dim m_strTPM As String = String.Empty
            Dim m_strWindowsDir As String = String.Empty
            Dim bit As String = String.Empty
            If My.Computer.Registry.LocalMachine.OpenSubKey("Hardware\Description\System\CentralProcessor\0").GetValue("Identifier").ToString.Contains("x86") Then
                bit = "32-bit"
            Else
                bit = "64-bit"
            End If

            objOS = New ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem")
            objCs = New ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystem")
            For Each objMgmt In objOS.Get
                m_strOSName = objMgmt("name").ToString()
                m_strOSVersion = objMgmt("version").ToString()
                m_strComputerName = objMgmt("csname").ToString()
                m_strWindowsDir = objMgmt("windowsdirectory").ToString()
            Next
            For Each objMgmt In objCs.Get
                m_strManufacturer = objMgmt("manufacturer").ToString()
                m_StrModel = objMgmt("model").ToString()
                m_strSystemType = objMgmt("systemtype").ToString
                m_strTPM = objMgmt("totalphysicalmemory").ToString()
            Next

            Using fs As FileStream = File.Create(logFile)
                LogInfo(fs, "BibleGetIO for MSWord Debug Log File, created " & DateTime.Now.ToString("F", New CultureInfo("en-US")) & Environment.NewLine)
                LogInfo(fs, "###############################################" & Environment.NewLine)
                LogInfo(fs, "Operating System:" & vbTab & vbTab & My.Computer.Info.OSFullName.ToString() & Environment.NewLine)
                LogInfo(fs, "OSPlatform:" & vbTab & vbTab & vbTab & My.Computer.Info.OSPlatform.ToString() & Environment.NewLine)
                LogInfo(fs, "OSVersion:" & vbTab & vbTab & vbTab & My.Computer.Info.OSVersion.ToString() & Environment.NewLine)
                LogInfo(fs, "Windows bit version: " & vbTab & vbTab & bit & Environment.NewLine)
                'LogInfo(fs, "Computer Name: " & vbTab & vbTab & My.Computer.Name.ToString() & Environment.NewLine)
                LogInfo(fs, "Computer Language: " & vbTab & vbTab & System.Globalization.CultureInfo.CurrentCulture.DisplayName & Environment.NewLine)
                LogInfo(fs, "Current Date/Time: " & vbTab & vbTab & Date.Now.ToLongDateString + ", " + Date.Now.ToLongTimeString & Environment.NewLine)
                'LogInfo(fs, "" & Environment.NewLine)
                LogInfo(fs, "Computer Manufacturer:" & vbTab & vbTab & m_strManufacturer & Environment.NewLine)
                LogInfo(fs, "Computer Model:" & vbTab & vbTab & vbTab & m_StrModel & Environment.NewLine)
                LogInfo(fs, "OS Version:" & vbTab & vbTab & vbTab & m_strOSVersion & Environment.NewLine)
                LogInfo(fs, "System Type:" & vbTab & vbTab & vbTab & m_strSystemType & Environment.NewLine)
                LogInfo(fs, "Windows Directory:" & vbTab & vbTab & m_strWindowsDir & Environment.NewLine)
                'LogInfo(fs, "" & Environment.NewLine)
                LogInfo(fs, "Number of Processors: " & vbTab & vbTab & Environment.ProcessorCount.ToString & Environment.NewLine)
                Dim moSearch As New ManagementObjectSearcher("Select * from Win32_Processor")
                Dim moReturn As ManagementObjectCollection = moSearch.Get
                For Each mo As ManagementObject In moReturn
                    LogInfo(fs, "Processor: " & vbTab & vbTab & vbTab & (mo("name")) & Environment.NewLine)
                Next
                Dim ramsize As Integer
                ramsize = My.Computer.Info.TotalPhysicalMemory / 1024 / 1024
                LogInfo(fs, "Memory: " & vbTab & vbTab & vbTab & ramsize.ToString & "MB RAM" & Environment.NewLine)
                LogInfo(fs, "" & Environment.NewLine)
                Dim WmiSelect As New ManagementObjectSearcher("root\CIMV2", "SELECT * FROM Win32_VideoController")
                Dim VGA As String = String.Empty
                For Each WmiResults As ManagementObject In WmiSelect.Get()
                    VGA = WmiResults.GetPropertyValue("Name").ToString
                Next
                LogInfo(fs, "Computer Display Info: " & vbTab & vbTab & VGA & Environment.NewLine)
                Dim intX As Integer = Windows.Forms.Screen.PrimaryScreen.Bounds.Width
                Dim intY As Integer = Windows.Forms.Screen.PrimaryScreen.Bounds.Height
                LogInfo(fs, "Screen Resolution: " & vbTab & vbTab & intX & " X " & intY & Environment.NewLine)
                'LogInfo(fs, "" & Environment.NewLine)
                Dim memory As Integer
                memory = My.Computer.Info.TotalPhysicalMemory / 1024 / 1024
                LogInfo(fs, "Total Physical Memory: " & vbTab & vbTab & memory.ToString() & "MB" & Environment.NewLine)
                memory = My.Computer.Info.TotalVirtualMemory / 1024 / 1024 / 1024
                LogInfo(fs, "Total Virtual Memory: " & vbTab & vbTab & memory.ToString() & "GB" & Environment.NewLine)
                memory = My.Computer.Info.AvailableVirtualMemory / 1024 / 1024 / 1024
                LogInfo(fs, "Available Virtual Memory: " & vbTab & memory.ToString() & "GB" & Environment.NewLine)
                memory = My.Computer.Info.AvailablePhysicalMemory / 1024 / 1024
                LogInfo(fs, "Available Physical Memory: " & vbTab & memory.ToString() & "MB" & Environment.NewLine)
                LogInfo(fs, "Network Available: " & vbTab & vbTab & My.Computer.Network.IsAvailable.ToString() & Environment.NewLine)
                LogInfo(fs, "###############################################" & Environment.NewLine)
                LogInfo(fs, "" & Environment.NewLine)
            End Using
        End If

        Dim lastUpdateCheck As DateTime = My.Settings.UpdateCheck.AddDays(7)
        Dim nowDateTime As DateTime = DateTime.Now
        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug("ThisAddIn.vb" & vbTab & "lastUpdateCheck = " & My.Settings.UpdateCheck.ToString)
        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug("ThisAddIn.vb" & vbTab & "lastUpdateCheck + 7 days = " & lastUpdateCheck.ToString)
        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug("ThisAddIn.vb" & vbTab & "now = " & nowDateTime.ToString)

        Dim lastUpdateFromNow As Int16 = DateTime.Compare(nowDateTime, lastUpdateCheck)
        If lastUpdateFromNow > 0 Then
            If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug("ThisAddIn.vb" & vbTab & "It has been more than 7 days since last update check")
            BibleGetAddIn.checkForUpdate()
        ElseIf lastUpdateFromNow = 0 Then
            If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug("ThisAddIn.vb" & vbTab & "It has been exactly 7 days since last update check")
        Else
            If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug("ThisAddIn.vb" & vbTab & "It has been less than 7 days since last update check")
        End If

    End Sub

    Private Shared Sub BibleGetAddIn_Shutdown() Handles Me.Shutdown

    End Sub

    'Private Sub Application_DocumentBeforeSave(ByVal Doc As Word.Document, ByRef SaveAsUI As Boolean, _
    '    ByRef Cancel As Boolean) Handles Application.DocumentBeforeSave
    '    'Doc.Paragraphs(1).Range.InsertParagraphBefore()
    '    'Doc.Paragraphs(1).Range.Text = "This text was added by using code."
    'End Sub

    'Private Sub Application_Startup() Handles Application.Startup
    '    ''Dim oCult As String = Application.Language.ToString
    '    ''Dim culture2 As CultureInfo = Thread.CurrentThread.CurrentCulture
    '    'Dim culture1 As CultureInfo = CultureInfo.CurrentUICulture
    '    'Dim lang As String = culture1.TwoLetterISOLanguageName
    '    If Me.DEBUG_MODE Then ThisAddIn.LogInfoToDebug(Me.GetType().FullName & vbTab & String.Fo_RMat("The current culture is {0}", lang))
    '    '_RM = New Resources.ResourceManager("BibleGetIO.BibleGetResource", System.Reflection.Assembly.GetExecutingAssembly())
    '    If Me.DEBUG_MODE Then ThisAddIn.LogInfoToDebug(Me.GetType().FullName & vbTab & _RM.BaseName)
    '    'Dim greeting As String = _RM.GetString("About this plugin", CultureInfo.CurrentUICulture)
    '    If Me.DEBUG_MODE Then ThisAddIn.LogInfoToDebug(Me.GetType().FullName & vbTab & greeting)
    'End Sub

    Public Shared Sub checkForUpdate()
        Dim onlineVersion As Version = HTTPCaller.GetCurrentVersion
        My.Settings.NewVersion = onlineVersion.ToString
        If Version.op_GreaterThan(onlineVersion, My.Application.Info.Version) Then
            My.Settings.NewVersionExists = True
        Else
            My.Settings.NewVersionExists = False
        End If
        My.Settings.UpdateCheck = DateTime.Now
        My.Settings.Save()
    End Sub

    Private Shared Sub LogInfo(ByVal fs As FileStream, ByVal value As String)
        Dim info As Byte() = New UTF8Encoding(True).GetBytes(value)
        fs.Write(info, 0, info.Length)
    End Sub

    Public Shared Sub LogInfoToDebug(ByVal value As String)
        Using fs As StreamWriter = File.AppendText(logFile)
            'Dim info As Byte() = New UTF8Encoding(True).GetBytes(value)
            fs.WriteLine(DateTime.Now.ToString("ddd MMM dd, yyyy HH:mm:ss.ffzz", New CultureInfo("en-US")) & vbTab & value)
            fs.WriteLine("")
        End Using
    End Sub


End Class

Public Class CSSRULE
    Public Shared ALIGN() As String = New String(3) {"left", "center", "right", "justify"}
    Public Shared TEXTSTYLES() As String = New String(3) {"bold", "italic", "underline", "line-through"}
    Public Shared BORDERSTYLE() As String = New String(8) {"none", "dotted", "dashed", "solid", "double", "groove", "ridge", "inset", "outset"}
End Class

Public Enum ALIGN
    LEFT
    CENTER
    RIGHT
    JUSTIFY
End Enum
Public Enum VALIGN
    SUPERSCRIPT
    SUBSCRIPT
    NORMAL
End Enum
Public Enum WRAP
    NONE
    PARENTHESES
    BRACKETS
End Enum
Public Enum POS
    TOP
    BOTTOM
    BOTTOMINLINE
End Enum
Public Enum FORMAT
    USERLANG
    BIBLELANG
    USERLANGABBREV
    BIBLELANGABBREV
End Enum
Public Enum VISIBILITY
    SHOW
    HIDE
End Enum

Public Enum PARAGRAPHTYPE
    BIBLEVERSION
    BOOKCHAPTER
    VERSES
    VERSENUMBER
    VERSETEXT
End Enum

Public Enum PREFERORIGIN
    HEBREW
    GREEK
End Enum