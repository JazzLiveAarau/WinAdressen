using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using ExcelUtil;
using System.IO;
using Ftp;

namespace AddressesJazz
{
    /// <summary>Class with functions that handle the seasons</summary>
    static public class Season
    {
        /// <summary>Start year for the storing of season data, i.e. for the columns with supporter amount columns</summary>
        static private int m_start_season = 2009;

        /// <summary>Get supporter season start year</summary>
        static public int GetSupporterSeasonStartYear()
        {
            return m_start_season;

        } // GetSupporterStartYear

        /// <summary>Start month for a new season is four (April). </summary>
        static private int m_start_date_new_season = 4; // April 

        /// <summary>Returns the current season as string. </summary>
        static public string GetCurrentSeason()
        {
            string ret_current_season = "";

            int current_season_start_year = GetCurrentSeasonStartYear();

            ret_current_season = current_season_start_year.ToString() + "-" + (current_season_start_year + 1).ToString();

            return ret_current_season;

        } // GetCurrentSeason

        /// <summary>Returns the current season start year as integer. </summary>
        static public int GetCurrentSeasonStartYear()
        {
            int ret_current_season_int = -12345;

            DateTime current_time = DateTime.Now;
            int now_year = current_time.Year;
            int now_month = current_time.Month;

            if (now_month < m_start_date_new_season)
            {
                ret_current_season_int = now_year - 1;
            }
            else
            {
                ret_current_season_int = now_year;
            }

            return ret_current_season_int;

        } // GetCurrentSeasonStartYear

        /// <summary>Returns record name for supporter data (column)</summary>
        static public string RecordNameSupporter(int i_start_year_season)
        {
            return AddressesJazzSettings.Default.Name_Start_Record_Supporter + SeasonString(i_start_year_season);

        } // RecordNameSupporter

        /// <summary>Returns record name for supporter data (column)</summary>
        static public string SeasonString(int i_start_year_season)
        {
            return i_start_year_season.ToString() + "-" + (i_start_year_season + 1).ToString();

        } // SeasonString

        /// <summary>Returns record type for supporter data</summary>
        static public string RecordTypeSupporter()
        {
            return AddressesJazzSettings.Default.Type_Record_Supporter;

        } // RecordTypeSupporter

        /// <summary>Returns record help for supporter data</summary>
        static public string RecordHelpSupporter(int i_start_year_season)
        {
            return AddressesJazzSettings.Default.Help_Record_Start_Supporter + SeasonString(i_start_year_season);

        } // RecordHelpSupporter


        /// <summary>Get all seasons as an array of strings.</summary>
        static public string[] GetAllSeasons()
        {
            string[] list_saisons;
            ArrayList array_list_saisons = new ArrayList();
            DateTime current_time = DateTime.Now;
            int now_year = current_time.Year;
            int now_month = current_time.Month;
            // Peter H also wanted next season int end_year = now_year + 1;
            int end_year = now_year;
            if (now_month < m_start_date_new_season)
            {
                end_year = end_year - 1;
            }

            for (int i_year = m_start_season; i_year <= end_year; i_year++)
            {
                string year_season = i_year.ToString() + "-" + (i_year+1).ToString();

                array_list_saisons.Add(year_season);
            }

            list_saisons = (string[])array_list_saisons.ToArray(typeof(string));

            return list_saisons;

        } // GetAllSeasons

        /// <summary>Get previous seasons as an array of strings.</summary>
        static public string[] GetPreviousSeasons()
        {
            string[] all_seasons = GetAllSeasons();

            int end_previous_seasons = all_seasons.Length - 2;

            string[] prev_seasons;
            prev_seasons = new string[end_previous_seasons];

            for (int i_prev = 0; i_prev < end_previous_seasons; i_prev++)
            {
                prev_seasons[i_prev] = all_seasons[i_prev];
            }

            return prev_seasons;

        } // GetPreviousSeasons

        /// <summary>Add the support from previous seasons to the input table</summary>
        static public bool DownloadPreviousSeasons(string i_addresses_file_name, string i_ftp_host, string i_ftp_user, string i_ftp_password, out string[] o_previous_file_names, out string o_error)
        {
            o_error = "";

            o_previous_file_names = _PreviuosFileNames(i_addresses_file_name);

            if (!InternetUtil.IsInternetConnectionAvailable())
            {
                o_error = AddressesJazzSettings.Default.ErrMsgNoInternetConnection;

                return false;
            }

            Ftp.DownLoad ftp_download = new DownLoad(i_ftp_host, i_ftp_user, i_ftp_password);

            for (int i_prev_file = 0; i_prev_file < o_previous_file_names.Length; i_prev_file++)
            {
                string previous_addresses_file_name = o_previous_file_names[i_prev_file];

                if (!ftp_download.DownloadBinary(AdressesUtility.FileUtil.GetServerFileName(previous_addresses_file_name), previous_addresses_file_name, out o_error))
                {
                    return false;
                }
            }


            return true;

        } // DownloadPreviousSeasons

        /// <summary>Returns the previous file names</summary>
        static private string[] _PreviuosFileNames(string i_addresses_file_name)
        {
            string[] ret_file_names;

            ArrayList array_list_file_names = new ArrayList();

            string[] all_seasons = GetAllSeasons();

            int end_previous_seasons = all_seasons.Length - 2;

            for (int i_season = 0; i_season < end_previous_seasons; i_season++)
            {
                string previuos_season = all_seasons[i_season];

                string previous_addresses_file_name = Path.GetDirectoryName(i_addresses_file_name) + @"\" + previuos_season + "-" + Path.GetFileName(i_addresses_file_name);

                array_list_file_names.Add(previous_addresses_file_name);
            }

            ret_file_names = (string[])array_list_file_names.ToArray(typeof(string));

            return ret_file_names;

        } // _PreviuosFileNames

    } // Season

} // namespace
