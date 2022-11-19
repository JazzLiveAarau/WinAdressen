using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelUtil;
using System.Collections;
using System.IO;

namespace AddressesJazz
{
    /// <summary>Class with functions to convert csv files by removing, adding and modifying columns.
    /// <para>These functions were only used once for converting data</para>
    /// <para>The class is kept since the functions are good sample code for Table functions</para>
    /// </summary>
    public static class ConvertCsv
    {
        /// <summary>Convert format from C++ application to the C# application. This is a function that only is used one time. Everything is hardcoded.</summary>
        /// <param name="i_file_csv_in">Input CSV file name</param>
        /// <param name="i_previous_file_names">Previous file names</param>
        /// <param name="i_file_csv_out">Output CSV file name</param>
        /// <param name="o_error">Error description</param>
        public static bool CPlusPlusToCSharp(string i_file_csv_in, string[] i_previous_file_names, string i_file_csv_out, out string o_error)
        {
            o_error = "";

            if (i_file_csv_in == i_file_csv_out)
            {
                o_error = "ConvertCsv.CPlusPlusToCSharp It is not allowed to overwrite the input file";
                return false;
            }

            // No longer used columns that shall be deleted
            string[] not_used_columns = { "0", "Telefon", "Zeitung", "Umfrage", "Löschen"};

            // Columns that shall change names
            string[] current_names = {"Beitrag", "BeitragNächste"};
            string[] new_names = {"Beitrag-2014-2015", "Beitrag-2015-2016"};

            // Columns that shall be moved
            string[] move_names = { "Kommentar" };
            string[] to_names = { "Beitrag-2014-2015"};   
      
            // Columns that shall be appended
            string[] append_names = {"Beitrag-2016-2017", "Beitrag-2017-2018"};

            Table table_csv = new Table("Table created from a CSV file");
            string csv_delimiter = "";
            Encoding file_encoding = null;
            if (!ExcelUtil.ToTable.CsvToTable(i_file_csv_in, ref table_csv, out csv_delimiter, ref file_encoding, out o_error)) return false;

            for (int i_delete = 0; i_delete < not_used_columns.Length; i_delete++)
            {
                string not_used_column = not_used_columns[i_delete];

                if (!ExcelUtil.TableTools.RemoveColumn(ref table_csv, not_used_column, out o_error)) return false;
            }

            // No support = '0'
            // Note start from row 1
            for (int i_row = 1; i_row < table_csv.NumberRows; i_row++)
            {
                string support_this_year = table_csv.GetFieldString(i_row, current_names[0], out o_error);
                if (o_error != "") return false;

                string support_next_year = table_csv.GetFieldString(i_row, current_names[1], out o_error);
                if (o_error != "") return false;

                if (support_this_year.Trim() == "")
                {
                    if (!table_csv.SetFieldString(i_row, current_names[0], "0", out o_error)) return false;
                }

                if (support_next_year.Trim() == "")
                {
                    if (!table_csv.SetFieldString(i_row, current_names[1], "0", out o_error)) return false;
                }
            }

            for (int i_rename = 0; i_rename < current_names.Length; i_rename++)
            {
                string current_name = current_names[i_rename];
                string new_name = new_names[i_rename];

                if (!ExcelUtil.TableTools.ChangeColumnName(ref table_csv, current_name, new_name, out o_error)) return false;
            }

            for (int i_move = 0; i_move < move_names.Length; i_move++)
            {
                string move_name = move_names[i_move];
                string to_name = to_names[i_move];

                if (!ExcelUtil.TableTools.MoveColumn(ref table_csv, move_name, to_name, out o_error)) return false;
            }

            int n_rows = table_csv.NumberRows;
            string[] fields_as_strings = _SupportColumnInitialValuesAsStrings(n_rows);

            for (int i_append = 0; i_append < append_names.Length; i_append++)
            {
                string append_name = append_names[i_append];
                fields_as_strings[0] = append_name;

                Column append_column;
                if (!ExcelUtil.TableTools.CreateColumn(fields_as_strings, out append_column, out o_error)) return false;

                int index_column = table_csv.NumberColumns;
                if (!ExcelUtil.TableTools.InsertColumn(ref table_csv, index_column, append_column, out o_error)) return false;

            }

            if (!_AddColumnsForPreviousSeasons(ref table_csv, i_previous_file_names, out o_error)) return false;

            if (!_AddSupportPreviousSeasons(ref table_csv, i_previous_file_names, out o_error)) return false;

            if (!ExcelUtil.FromTable.TableToCsv(table_csv, i_file_csv_out, csv_delimiter, file_encoding, out o_error)) return false;

            return true;
        }

        /// <summary>Create array of default values for a support column</summary>
        static private string[] _SupportColumnInitialValuesAsStrings(int i_number_rows)
        {
            string[] fields_as_strings;
            ArrayList array_list_fields = new ArrayList();

            for (int i_row = 0; i_row < i_number_rows; i_row++)
            {
                string field_value = "";
                if (0 == i_row)
                {
                    field_value = "AppendColumn";
                }
                else
                {
                    field_value = "0"; 
                }

                array_list_fields.Add(field_value);
            }

            fields_as_strings = (string[])array_list_fields.ToArray(typeof(string));

            return fields_as_strings;
        }

        /// <summary>Add columns in the table for previuos seasons. TODO What about the column headers ??? </summary>
        static private bool _AddColumnsForPreviousSeasons(ref Table io_table_csv, string[] i_previous_file_names, out string o_error)
        {
            o_error = "";

            if (i_previous_file_names == null)
            {
                o_error = "ConvertCsv.AddColumnsForPreviousSeasons Previous files is null";
                return true;
            }

            int n_rows = io_table_csv.NumberRows;
            string[] fields_as_strings = _SupportColumnInitialValuesAsStrings(n_rows);

            string[] insert_previous_names;

            insert_previous_names = new string[i_previous_file_names.Length];

            for (int i_support = 0; i_support < i_previous_file_names.Length; i_support++)
            {
                insert_previous_names[i_support] = _SupportColumnNameFromFile(i_previous_file_names[i_support]);
            }

            string current_season = Season.GetCurrentSeason();
            string insert_before_name = "Beitrag-" + current_season;

            for (int i_insert = 0; i_insert < insert_previous_names.Length; i_insert++)
            {

                string insert_previous_name = insert_previous_names[i_insert];
                fields_as_strings[0] = insert_previous_name;

                Column insert_column;
                if (!ExcelUtil.TableTools.CreateColumn(fields_as_strings, out insert_column, out o_error)) return false;

                if (!ExcelUtil.TableTools.InsertColumn(ref io_table_csv, insert_before_name, insert_column, out o_error)) return false;
            }

            return true;
        }

        /// <summary>Add the support from previous seasons to the input table</summary>
        static private bool _AddSupportPreviousSeasons(ref Table io_table_addresses, string[] i_previous_file_names, out string o_error)
        {
            o_error = "";

            if (i_previous_file_names == null)
            {
                o_error = "_AddSupportPreviousSeasons Previous files is null";
                return true;
            }

            for (int i_prev_file = 0; i_prev_file < i_previous_file_names.Length; i_prev_file++)
            {
                string previous_addresses_file_name = i_previous_file_names[i_prev_file];

                if (!File.Exists(previous_addresses_file_name))
                {
                    o_error = "_AddSupportPreviousSeasons Missing file " + previous_addresses_file_name;
                    return false;
                }

                Table prev_table_addresses = new Table("Table previous");
                if (!ToTable.CsvToTable(previous_addresses_file_name, ref prev_table_addresses, out o_error))
                    return false;

                string support_column_name = _SupportColumnNameFromFile(previous_addresses_file_name);

                // Note start from row 1
                for (int i_row = 1; i_row < prev_table_addresses.NumberRows; i_row++)
                {
                    string support_value_str = prev_table_addresses.GetFieldString(i_row, "Beitrag", out o_error);
                    if (o_error != "") return false;
                    if (support_value_str.Trim() != "")
                    {
                        string family_name = prev_table_addresses.GetFieldString(i_row, "FamilienName", out o_error);
                        if (o_error != "") return false;

                        string first_name = prev_table_addresses.GetFieldString(i_row, "Vorname", out o_error);
                        if (o_error != "") return false;

                        if (!_UpdateSupportColumn(ref io_table_addresses, family_name, first_name, support_value_str.Trim(), support_column_name, out o_error)) return false;
                    }
                }
            }


            return true;
        }

        /// <summary>Update the support column value if the same first and family name exists</summary>
        static private bool _UpdateSupportColumn(ref Table io_table_addresses, string i_family_name, string i_first_name, string i_support_str, string i_support_column_name, out string o_error)
        {
            o_error = "";

            // Note start from row 1
            for (int i_row = 1; i_row < io_table_addresses.NumberRows; i_row++)
            {
                string family_name = io_table_addresses.GetFieldString(i_row, "FamilienName", out o_error);
                if (o_error != "") return false;

                string first_name = io_table_addresses.GetFieldString(i_row, "Vorname", out o_error);
                if (o_error != "") return false;

                if (family_name == i_family_name && first_name == i_first_name)
                {
                    if (!io_table_addresses.SetFieldString(i_row, i_support_column_name, i_support_str, out o_error)) return false;
                    break;
                }
            }

            return true;
        }

        /// <summary>Returns support colum name constructed from the input full file name</summary>
        static private string _SupportColumnNameFromFile(string i_full_file_name)
        {
            string file_name = Path.GetFileName(i_full_file_name);
            string season_str = file_name.Substring(0, 9);
            string support_column_name = "Beitrag-" + season_str; // TODO Define Beitrag as config parameter 
            return support_column_name;
        }
    }
}
