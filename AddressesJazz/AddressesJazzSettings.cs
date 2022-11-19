using System;
using System.IO;
using System.Xml.Serialization;
using AdressesUtility;

namespace AddressesJazz
{
    /// <summary>Holds the settings of the application.</summary>
    public sealed class AddressesJazzSettings
    {
        static AddressesJazzSettings defaultSettings = new AddressesJazzSettings();

        /// <summary>FTP host</summary>
        public string FtpHost = "www.jazzliveaarau.ch";

        /// <summary>FTP user</summary>
        public string FtpUser = "jazzliv1";

        /// <summary>Configuration XML root element</summary>
        public string ConfigRootElement = "AddressesJazzSettings";

        /// <summary>Name of the file with the addresses</summary>
        public string AddressesFileName = "JazzLiveAdressen.csv";

        /// <summary>Server directory for the address file AddressesFileName</summary>
        public string AddressesServerDir = "/adressen/";

        /// <summary>Name of the file with the addresses for Beta (test) version</summary>
        public string AddressesFileNameBeta = "JazzLiveAdressenBeta.csv";

        /// <summary>Flag telling if it a Beta (test) version</summary>
        public bool BetaVersion = false;

        /// <summary>Flag telling if current addresses shall be copied to Beta (start) addresses</summary>
        public bool CopyCurrentAddressesForBetaVersion = false;

        /// <summary>Name of the help file</summary>
        public string FileHelp = @"JAZZ_live_AARAU_Adressen.rtf";

        /// <summary>Name of the test protocol file</summary>
        public string FileTest = @"TestProtokoll.txt";

        /// <summary>Name of the checkin-checkout logfile for the addresses</summary>
        public string CheckInOutLogFileName = "CheckInCheckOutAddresses.log";

        /// <summary>Start string for checkin row in logfile</summary>
        public string CheckInLogFile = @"Checked in by:";

        /// <summary>Start string for checkout row in logfile</summary>
        public string CheckOutLogFile = @"Checked out by:";

        /// <summary>Name of the directory with file AddressFileName</summary>
        public string AddressesDir = @"Excel";

        /// <summary>Name of the directory with file AddressFileName backups</summary>
        public string AddressesBackupsDir = @"Backups";

        /// <summary>Name of the directory with output files</summary>
        public string OutputDir = @"Output";

        /// <summary>Name of the directory for the help file</summary>
        public string HelpDir = @"Help";

        /// <summary>Setup server directory for application Adressen</summary>
        public string AdressenServerDir = @"setup/Adressen";

        /// <summary>Name of the server directory that has a file with the latest version info</summary>
        public string LatestVersionInfoDir = @"LatestVersionInfo";

        /// <summary>Name of the local and server directory for the installer for a new version</summary>
        public string NewVersionDir = @"NeueVersion";

        /// <summary>URL for the program documentation</summary>
        public string ProgramDocumentation = @"http://www.jazzliveaarau.ch/Administration/Wiki/wiki_documentation.htm";

        #region Output lists

        /// <summary>Output header </summary>
        public string OutputList_00 = "Output wählen";

        /// <summary>Output mail addresses as Excel XSLX</summary>
        public string OutputList_01 = "Versand.xlsx";

        /// <summary>Output mail addresses as Excel CSV</summary>
        public string OutputList_02 = "Versand.csv";

        /// <summary>Output supporter as text</summary>
        public string OutputList_03 = "Supporter.txt";

        /// <summary>Output sponsors as a text file</summary>
        public string OutputList_04 = "Sponsor.txt";

        /// <summary>Output E-Mail addresses as a text file</summary>
        public string OutputList_05 = "Email.txt";

        /// <summary>Output supporter as Excel CSV</summary>
        public string OutputList_06 = "Supporter.csv";

        /// <summary>Output supporter as XML</summary>
        public string OutputList_07 = "Supporter.xml";

        /// <summary>Output check results as a text file</summary>
        public string OutputList_08 = "Check.txt";

        #endregion // Output lists

        #region Captions
        /// <summary>Caption for the combobox season</summary>
        public string Caption_Season = "Saison";

        /// <summary>Caption for the support</summary>
        public string Caption_Support = "Beitrag";

        /// <summary>Caption for the button next record</summary>
        public string Caption_Next = ">";

        /// <summary>Caption for the button previous record</summary>
        public string Caption_Previous = "<";

        /// <summary>Caption for the label search record</summary>
        public string Caption_Search = "Suchen";

        /// <summary>Caption for the button delete record</summary>
        public string Caption_Delete = "Löschen";

        /// <summary>Caption for the button add record</summary>
        public string Caption_Add = "Neu";

        /// <summary>Caption for the button Checkin/Checkout Undefined</summary>
        public string Caption_CheckInOutUndefined = "---";

        /// <summary>Caption for the button Checkin/CheckOut</summary>
        public string Caption_CheckOut = "Check out";

        /// <summary>Caption for the button Checkin/CheckOut</summary>
        public string Caption_CheckIn = "Speichern";

        /// <summary>Caption for the exit button</summary>
        public string Caption_Exit = "Ende";

        /// <summary>Caption selection of backup file for reset</summary>
        public string Caption_BackupFileSelect = "Backup Datei wählen";

        #endregion // Captions

        #region Columns

        /// <summary>Record 1 name</summary>
        public string Name_Record_01 = "Vorname";
        /// <summary>Caption for record 1</summary>
        public string Caption_Record_01 = "Vorname";
        /// <summary>Type record 1</summary>
        public string Type_Record_01 = "string";
        /// <summary>Help text record 1</summary>
        public string Help_Record_01 = "Vorname eingeben. Darf leer sein.";

        /// <summary>Record 2 name</summary>
        public string Name_Record_02 = "FamilienName";
        /// <summary>Caption for record 2</summary>
        public string Caption_Record_02 = "FamilienName";
        /// <summary>Type record 2</summary>
        public string Type_Record_02 = "string";
        /// <summary>Help text record 2</summary>
        public string Help_Record_02 = "Familienname eingeben. Sollte nicht leer sein.";

        /// <summary>Record 3 name</summary>
        public string Name_Record_03 = "Strasse";
        /// <summary>Caption for record 3</summary>
        public string Caption_Record_03 = "Strasse";
        /// <summary>Type record 3</summary>
        public string Type_Record_03 = "string";
        /// <summary>Help text record 3</summary>
        public string Help_Record_03 = "Strasse eingeben. Darf leer sein.";

        /// <summary>Record 4 name</summary>
        public string Name_Record_04 = "Hausnummer";
        /// <summary>Caption for record 4</summary>
        public string Caption_Record_04 = "Hausnummer";
        /// <summary>Type record 4</summary>
        public string Type_Record_04 = "string";
        /// <summary>Help text record 4</summary>
        public string Help_Record_04 = "Hausnummer eingeben. Darf leer sein.";

        /// <summary>Record 5 name</summary>
        public string Name_Record_05 = "PLZ";
        /// <summary>Caption for record 5</summary>
        public string Caption_Record_05 = "PLZ";
        /// <summary>Type record 5</summary>
        public string Type_Record_05 = "string";
        /// <summary>Help text record 5</summary>
        public string Help_Record_05 = "PLZ eingeben. Darf leer sein.";

        /// <summary>Record 6 name</summary>
        public string Name_Record_06 = "Wohnort";
        /// <summary>Caption for record 6</summary>
        public string Caption_Record_06 = "Wohnort";
        /// <summary>Type record 6</summary>
        public string Type_Record_06 = "string";
        /// <summary>Help text record 6</summary>
        public string Help_Record_06 = "Wohnort eingeben. Darf leer sein.";

        /// <summary>Record 7 name</summary>
        public string Name_Record_07 = "e-Mail";
        /// <summary>Caption for record 7</summary>
        public string Caption_Record_07 = "e-Mail";
        /// <summary>Type record 7</summary>
        public string Type_Record_07 = "string";
        /// <summary>Help text record 7</summary>
        public string Help_Record_07 = "e-Mail eingeben. Darf leer sein.";

        /// <summary>Record 8 name</summary>
        public string Name_Record_08 = "Post";
        /// <summary>Caption for record 8</summary>
        public string Caption_Record_08 = "Post";
        /// <summary>Type record 8</summary>
        public string Type_Record_08 = "boolean";
        /// <summary>Help text record 8</summary>
        public string Help_Record_08 = "Saisonsprogramm schicken";

        /// <summary>Record 9 name</summary>
        public string Name_Record_09 = "NewsletterJazz";
        /// <summary>Caption for record 9</summary>
        public string Caption_Record_09 = "Newsletter";
        /// <summary>Type record 9</summary>
        public string Type_Record_09 = "boolean";
        /// <summary>Help text record 9</summary>
        public string Help_Record_09 = "Newsletter schicken";

        /// <summary>Record 10 name</summary>
        public string Name_Record_10 = "Sponsor";
        /// <summary>Caption for record 10</summary>
        public string Caption_Record_10 = "Sponsor";
        /// <summary>Type record 10</summary>
        public string Type_Record_10 = "boolean";
        /// <summary>Help text record 10</summary>
        public string Help_Record_10 = "Sponsor flag";

        /// <summary>Record 11 name</summary>
        public string Name_Record_11 = "Kommentar";
        /// <summary>Caption for record 11</summary>
        public string Caption_Record_11 = "Kommentar";
        /// <summary>Type record 11</summary>
        public string Type_Record_11 = "string";
        /// <summary>Help text record 11</summary>
        public string Help_Record_11 = "Kommentar eingeben";

        /// <summary>Start part for record supporter. Season years (starting with 2009-2010) shall be added</summary>
        public string Name_Start_Record_Supporter = "Beitrag-";
        /// <summary>Type record supporter</summary>
        public string Type_Record_Supporter = "float";
        /// <summary>Help text start for supporter. Season years (starting with 2009-2010) shall be added</summary>
        public string Help_Record_Start_Supporter = "Beitrag eingeben für Saison ";

        #endregion // Columns

        #region Status and error messages

        /// <summary>Error message: There is no input Excel file </summary>
        public string ErrMsgNoExcelFile = @"Keine Excel Datei: ";

        /// <summary>Error message: No connection to Internet is available</summary>
        public string ErrMsgNoInternetConnection = @"Keine Verbindung zu Internet ist vorhanden";

        /// <summary>Error message: Failure downloading Excel file </summary>
        public string ErrMsgNoExcelFileDownload = @"Adressen sind nicht heruntergeladen";

        /// <summary>Error message: Failure downloading installer for a new version</summary>
        public string ErrMsgNewVersionDownload = @"Installationsprogramm für eine neue Vesion wurde nicht heruntergeladen";

        /// <summary>Message: Installer for a new version is downloaded</summary>
        public string MsgNewVersionDownload = @"Installationsprogramm neuester Version ist heruntergeladen." +
            "\r\nDie Datei ist im Ordner " + @"C:\Apps\JazzLiveAarau\Adressen\NeueVersion gespeichert." +
            "\r\nDiese Applikation beenden bevor Installation." +
             "\r\nDoppelklick auf die .exe (Anwendung) Datei im Ordner NeueVersion für Installation der neuen Version.";

        /// <summary>Error message: Failure uploading Excel file </summary>
        public string ErrMsgUploadAddressesFailed = @"Hochladen von Adressdatei zum Server scheiterte";

        /// <summary>Error message: Failure uploading backup Excel file </summary>
        public string ErrMsgUploadBackupAddressesFailed = @"Hochladen von Backup-Datei zum Server scheiterte";

        /// <summary>Error message: Failure downloading Excel file </summary>
        public string ErrMsgNoCheckInOutLogFileDownload = @"Checkin Log-Datei nicht heruntergeladen";

        /// <summary>Error message: Adresses are already checked out by somebody else </summary>
        public string ErrMsgAddressesCheckOutBy = @"Adressen sind schon checked out bei ";

        /// <summary>Error message: E-Mail address contains no At Sign @ </summary>
        public string ErrMsgEmailAddressNoAtSign = @"Kein @ in der E-Mail Adresse ";

        /// <summary>Error message: There must a first name or a family name </summary>
        public string ErrMsgNoFirstNameNoFamilyName = @"Kein Vorname oder Familienname";

        /// <summary>Error message: Street must be defined for Mail</summary>
        public string ErrMsgNoStreetForMail = @"Für Post musst Strasse definiert sein";

        /// <summary>Error message: Postal code must be defined for Mail</summary>
        public string ErrMsgNoPostalCodeForMail = @"Für Post musst PLZ definiert sein";

        /// <summary>Error message: City must be defined for Mail</summary>
        public string ErrMsgNoCityForMail = @"Für Post musst Wohnort definiert sein";

        /// <summary>Error message: Email address must be defined for Newsletter</summary>
        public string ErrMsgNoEmailAddressForNewsletter = @"Für Newsletter musst E-Mail Adresse definiert sein";

        /// <summary>Error message: Checked out addresses are not saved</summary>
        public string ErrMsgRecordNotSaved = @"Geänderte Daten wurden nicht gespeichert";

        /// <summary>Error message: Checked out addresses are not saved</summary>
        public string ErrMsgCancelWithoutSave = @"Adressen wurden nicht gespeichert";

        /// <summary>Message: Force checkout anyhow? </summary>
        public string MsgAddressesForceCheckOut = "Trotzdem check out?";

        /// <summary>Error message: Adresses are already checked out by somebody else </summary>
        public string ErrMsgUploadLogfile = @"Checkin-Checkout Logfile konnte nicht hochgeladen werden";

        /// <summary>Error message: Output of the requested list is not yet implemented </summary>
        public string ErrMsgOutputNotYetImplemented = @" ist noch nicht implementiert";

        /// <summary>Message: Message when the user makes Exit or Cancel and the addresses are checked out</summary>
        public string MsgShallAddressesBeUploaded = @"Möchten Sie dass die Adressen hochgeladen (gespeichert) werden?";

        /// <summary>Message: Message cation when the user makes Exit or Cancel and the addresses are checked out</summary>
        public string MsgCaptionShallAddressesBeUploaded = @"Adressen sind ausgecheckt";

        /// <summary>Message: Excel file is downloaded from the server</summary>
        public string MsgExcelFileDownload = @"Adressen sind vom Server heruntergeladen";

        /// <summary>Message: Addresses are checked out</summary>
        public string MsgAddressesAreCheckedOut = @"Adressen sind ausgecheckt";

        /// <summary>Message: Addresses are checked in</summary>
        public string MsgAddressesAreCheckedIn = @"Adressen sind eingecheckt";

        /// <summary>Error message: There is no local address file </summary>
        public string ErrMsgNoLocalExcelFile = @"Es gibt kein Datei ";

        /// <summary>Error message: The addresses have not been checked out </summary>
        public string ErrMsgAddressesNotCheckedOut = @"Adressen sind nicht checked out";

        /// <summary>Message: Exit application</summary>
        public string MsgExitApplication = @"Bitte die Applikation beenden!";

        /// <summary>Not valid character(s) in string. String has been modified</summary>
        public string ErrMsgNotValidCharsHaveBeenRemoved = @"Nicht erlaubte Zeichen wie zum Beispiel Kommas sind weggenommen";

        /// <summary>Only numbers are allowed</summary>
        public string ErrMsgAllCharsExceptNumbersHaveBeenRemoved = @"Nur Zahlen sind erlaubt: Kein Komma, Punkt, Leerzeichen, Etc.";

        /// <summary>Error message: Failure getting backup address files </summary>
        public string ErrMsgGettingBackupFilesFailed = @"Backup-Adressen fehlen";

        /// <summary>Error message: Reset only if addresses are checked out </summary>
        public string ErrMsgAddressesMustBeCheckedOutForReset = @"Bevor Reset muss man Checkout machen";

        /// <summary>Error message: File is locked, i.e. it is opened by another program</summary>
        public string ErrMsgFileIsLocked = @"Datei kann nicht kreiert werden" + "\r\n" + "Bitte Applikation beenden, die diese Datei geöffnet hat:" + "\r\n";

        /// <summary>Status message: A new version is available</summary>
        public string MsgNewVersionIsAvailable = @"Es gibt eine neue Version ";

        #endregion // Status and error messages


        #region GUI strings

        /// <summary>GUI program title</summary>
        public string GuiTextProgramTitle = @"JAZZ live AARAU Adressen";

        /// <summary>GUI default band text</summary>
        public string GuiTextFamilyName = @"Familienname";

        /// <summary>GUI help dialog title</summary>
        public string GuiHelpDialogTitle = @"JAZZ live AARAU Adressen Hilfe";

        /// <summary>GUI help dialog exit button</summary>
        public string GuiHelpDialogExit = @"Schliessen";

        #endregion


        #region Tool tips

        /// <summary>GUI Tool tip application</summary>
        public string ToolTipApplication = @"Adressdatenbank JAZZ live AARAU";

        /// <summary>GUI Tool tip first name</summary>
        public string ToolTipTextBoxFirstName = @"Vorname eingeben";

        /// <summary>GUI Tool tip family name</summary>
        public string ToolTipTextBoxFamyliName = @"Familienname eingeben";

        /// <summary>GUI Tool tip street name</summary>
        public string ToolTipTextBoxStreetName = @"Strassenname eingeben";

        /// <summary>GUI Tool tip house number</summary>
        public string ToolTipTextBoxHouseNumber = @"Hausnummer eingeben";

        /// <summary>GUI Tool tip postal code</summary>
        public string ToolTipTextBoxPostalCode = @"ZIP eingeben";

        /// <summary>GUI Tool tip city name</summary>
        public string ToolTipTextBoxCityName = @"Stadt eingeben";

        /// <summary>GUI Tool tip E-Mail address</summary>
        public string ToolTipTextBoxEmailAddress = @"E-Mail Adresse eingeben";

        /// <summary>GUI Tool tip comment one</summary>
        public string ToolTipTextBoxCommentOne = @"Kommentar eingeben";

        /// <summary>GUI Tool tip support</summary>
        public string ToolTipTextBoxSupport = @"Supporter Beitrag eingeben";

        /// <summary>GUI Tool tip delete record</summary>
        public string ToolTipButtonDelete = @"Person löschen";

        /// <summary>GUI Tool tip add record</summary>
        public string ToolTipButtonAdd = @"Person (Datensatz) zufügen";

        /// <summary>GUI Tool tip sort with first name</summary>
        public string ToolTipButtonSortFirstName = @"Sortieren mit Vorname";

        /// <summary>GUI Tool tip sort with family name</summary>
        public string ToolTipButtonSortFamilyName = @"Sortieren mit Familienname";

        /// <summary>GUI Tool tip sort with postal code</summary>
        public string ToolTipButtonSortPostalCode = @"Sortieren mit PLZ";

        /// <summary>GUI Tool tip checkin-checkout button</summary>
        public string ToolTipButtonCheckInOut = @"Einchecken oder Auschecken von Adressen";

        /// <summary>GUI Tool tip previous record</summary>
        public string ToolTipButtonPreviousRecord = @"Voriger Datensatz";

        /// <summary>GUI Tool tip previous record</summary>
        public string ToolTipButtonNextRecord = @"Nächster Datensatz";

        /// <summary>GUI Tool tip output data</summary>
        public string ToolTipOutputData = @"Output wählen."+"\r\nDie Dateien werden in diesem Ordner gespeichert:" +
            "\r\n" + @"C:\Apps\JazzLiveAarau\Adressen\Output";

        /// <summary>GUI Tool tip season</summary>
        public string ToolTipSeason = @"Saison wählen";

        /// <summary>GUI Tool tip search string</summary>
        public string ToolTipSearchString = @"Such-String eingeben";

        /// <summary>GUI Tool tip search hits</summary>
        public string ToolTipSearchStringHits = @"Anzahl Datensatzen mit Such-String";

        /// <summary>GUI Tool tip status message</summary>
        public string ToolTipStatusMessage = @"Status";

        /// <summary>GUI Tool tip exit application</summary>
        public string ToolTipExitApplication = @"Ende Applikation";

        /// <summary>GUI Tool tip check box mail</summary>
        public string ToolTipCheckPost = @"Post schicken";

        /// <summary>GUI Tool tip check box newsletter</summary>
        public string ToolTipCheckNewsletter = @"Newsletter schicken";

        /// <summary>GUI Tool tip check box sponsor</summary>
        public string ToolTipCheckSponsor = @"Sponsor für JAZZ live AARAU";

        /// <summary>GUI Tool tip search results</summary>
        public string ToolTipSearchResults = @"Person (Datensatz) wählen";

        /// <summary>GUI Tool tip help</summary>
        public string ToolTipButtonHelp = @"Instruktionen für JAZZ live AARAU Adressen";

        /// <summary>GUI Tool tip help dialog text</summary>
        public string ToolTipHelpTextBox = @"Text vom Hilfsdokument im Hilfeordner wird gezeigt";

        /// <summary>GUI Tool tip help dialog exit</summary>
        public string ToolTipHelpExit = @"Fenster wird geschlossen";

        /// <summary>GUI Tool tip update</summary>
        public string ToolTipButtonUpdate = @"Installationsprogramm neuester Version herunterladen." + 
            "\r\nDie Datei wird im Ordner " + @"C:\Apps\JazzLiveAarau\Adressen\NeueVersion gespeichert." +
            "\r\nDiese Applikation beenden bevor Installation." +
            "\r\nDoppelklick auf die .exe Datei im Ordner NeueVersion für Installation der neuen Version.";

        /// <summary>GUI Tool tip reset</summary>
        public string ToolTipReset = @"Beim einen Problem mit dem Adressdaten (CSV Datei) kann man eine alte Backup-Datei laden";

        /// <summary>GUI Tool tip new version is available</summary>
        public string ToolTipNewVersionAvailable = @"Bitte die neue Version installieren! Siehe Help.";

        #endregion

        /// <summary>Constructor</summary>
        public AddressesJazzSettings() { }

        /// <summary>Gets the default settings instance.</summary>
        /// <remarks>
        /// <para>On first access, an attempt is made to load the settings from an application-specific location. If the
        /// file is not found or corrupt, then all fields of the returned instance are set to their default values.
        /// </para>
        /// </remarks>
        internal static AddressesJazzSettings Default
        {
            get { return defaultSettings; }
        }

        /// <summary>Saves all settings.</summary>
        internal void Save()
        {
            // Always existing Directory.CreateDirectory(FileUtil.GetPathToExeDirectory());

            using (FileStream fileStream = new FileStream(AdressesUtility.FileUtil.ConfigFileName(ConfigRootElement, JazzMain.m_exe_directory), FileMode.Create))
            using (StreamWriter streamWriter = new StreamWriter(fileStream))
            {
                new XmlSerializer(typeof(AddressesJazzSettings)).Serialize(streamWriter, defaultSettings);
            }
        }

        /// <summary>Reads the configuration file and sets values in defaultSettings.</summary>
        internal void ReadFromConfigFile()
        {
            using (FileStream fileStream = new FileStream(AdressesUtility.FileUtil.ConfigFileName(ConfigRootElement, JazzMain.m_exe_directory), FileMode.Open, FileAccess.Read, FileShare.Read))
            using (StreamReader streamReader = new StreamReader(fileStream))
            {
                defaultSettings = (AddressesJazzSettings)new XmlSerializer(typeof(AddressesJazzSettings)).Deserialize(streamReader);
            }
        }

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        static AddressesJazzSettings()
        {
            try
            {
                using (FileStream fileStream = new FileStream(AdressesUtility.FileUtil.ConfigFileName(AddressesJazzSettings.defaultSettings.ConfigRootElement, JazzMain.m_exe_directory), FileMode.Open, FileAccess.Read, FileShare.Read))
                using (StreamReader streamReader = new StreamReader(fileStream))
                {
                    defaultSettings = (AddressesJazzSettings)new XmlSerializer(typeof(AddressesJazzSettings)).Deserialize(streamReader);
                }
            }
            catch (FileNotFoundException) { }
            catch (DirectoryNotFoundException) { }
            catch (InvalidOperationException) { } // Thrown when there is an error in the XML document
            catch (InvalidCastException) { } // Thrown occasionally in Visual Studio when opening designer
            catch (Exception e)
            {
                using (StreamWriter w = File.AppendText(Path.Combine(JazzMain.m_exe_directory, "Settings-debug-log.txt")))
                {
                    w.WriteLine();
                    w.WriteLine(">>> Unhandled Exception " + e.GetType() + " occurred at " + DateTime.Now + "!");
                    w.WriteLine();
                    w.WriteLine(e);
                    w.WriteLine();

                    // Close the writer and underlying file.
                    w.Close();
                }
            }
        }
    }
}

