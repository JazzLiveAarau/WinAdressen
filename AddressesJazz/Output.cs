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
    /// <summary>Class that handles output of lists
    /// <para>Output is for instance Excel files and text files with names and addresses for the yearly letter and for the supporters.</para>
    /// <para>The requests are commands (calls) normally coming from the JazzForm class.</para>
    /// </summary>
    static public class Output
    {
        /// <summary>Get all available output lists as an array of strings.
        /// <para>The function should be called from a Form class for e.g. a Combo Box.</para>
        /// <para>The returned (hardcoded) array defines the available output lists.</para>
        /// <para>The first element in the array is a header (caption) for the elements</para>
        /// <para>All array names are defined in the configuration file (class AddressesJazzSettings).</para>
        /// </summary>
        static public string[] GetAllOutput()
        {
            string[] list_output;
            ArrayList array_list_output = new ArrayList();

            array_list_output.Add(AddressesJazzSettings.Default.OutputList_00);
            array_list_output.Add(AddressesJazzSettings.Default.OutputList_01);
            array_list_output.Add(AddressesJazzSettings.Default.OutputList_02);
            array_list_output.Add(AddressesJazzSettings.Default.OutputList_03);
            array_list_output.Add(AddressesJazzSettings.Default.OutputList_04);
            array_list_output.Add(AddressesJazzSettings.Default.OutputList_05);
            array_list_output.Add(AddressesJazzSettings.Default.OutputList_06);
            array_list_output.Add(AddressesJazzSettings.Default.OutputList_07);
            array_list_output.Add(AddressesJazzSettings.Default.OutputList_08);

            list_output = (string[])array_list_output.ToArray(typeof(string));

            return list_output;
        }

        /// <summary>Get the header (prompt/caption),i.e. the first item in the array of output lists.</summary>
        static public string GetHeaderItem()
        {
            return AddressesJazzSettings.Default.OutputList_00;
        }

        /// <summary>Output the requested list
        /// <para>One of the following execution functions are called</para>
        /// <para>- MailAdressesAsXslx</para>
        /// <para>- MailAdressesAsCsv</para>
        /// <para>- SupportListTxt</para>
        /// <para>- SponsorListTxt</para>
        /// <para>- SupportAddressesAsCsv</para>
        /// </summary>
        /// <param name="i_table_addresses">Table with addresse</param>
        /// <param name="i_season_column_name">Season column name</param>
        /// <param name="i_selected_index">Index defining the output list</param>
        /// <param name="i_selected_inner_text">Text corresponding to the selected index</param>
        /// <param name="i_output_file_name">Name of the output file</param>
        /// <param name="o_error">Error message</param>
        static public bool ExecuteRequest(Table i_table_addresses, string i_season_column_name, int i_selected_index, 
            string i_selected_inner_text, string i_output_file_name, out string o_error)
        {
            o_error = "";

            if (i_selected_index > 0 && FileIsLocked(i_output_file_name))
            {
                o_error =  AddressesJazzSettings.Default.ErrMsgFileIsLocked + i_output_file_name;
                return false;
            }

            if (0 == i_selected_index)
            {
                o_error = i_selected_inner_text + @" is a header. No execution";
                return true;
            }
            else if (1 == i_selected_index)
            {
                if (!MailAdressesAsXslx(i_table_addresses, i_output_file_name, out o_error))
                {
                    return false;
                }

                return true;
            }
            else if (2 == i_selected_index)
            {
                if (!MailAdressesAsCsv(i_table_addresses, i_output_file_name, out o_error))
                {
                    return false;
                }

                return true;
            }
            else if (3 == i_selected_index)
            {
                if (!SupportListTxt(i_table_addresses, i_season_column_name, i_output_file_name, out o_error))
                {
                    return false;
                }

                return true;
            }
            else if (4 == i_selected_index)
            {
                if (!SponsorListTxt(i_table_addresses, i_output_file_name, out o_error))
                {
                    return false;
                }

                return true;
            }
            else if (5 == i_selected_index)
            {
                if (!NewsletterListTxt(i_table_addresses, i_output_file_name, out o_error))
                {
                    return false;
                }

                return true;
            }
            else if (6 == i_selected_index)
            {
                if (!SupportAddressesAsCsv(i_table_addresses, i_season_column_name, i_output_file_name, out o_error))
                {
                    return false;
                }

                return true;
            }
            else if (7 == i_selected_index)
            {
                if (!SupportersAsXml(i_table_addresses, i_season_column_name, i_output_file_name, out o_error))
                {
                    return false;
                }

                System.Diagnostics.Process.Start("notepad.exe", i_output_file_name);

                return true;
            }
            else if (8 == i_selected_index)
            {
                if (!CheckListTxt(i_table_addresses, i_output_file_name, i_season_column_name, out o_error))
                {
                    return false;
                }

                return true;
            }
            else 
            {
                o_error = @"Programming error Output.ExecuteRequest Index= " + i_selected_index.ToString() 
                               + @" Inner text= " + i_selected_inner_text;
                return false;
            }

        } // ExecuteRequest

        /// <summary>Output mail addresses as XLSX file
        /// <para>A Table with mail addresses for the XLSX output is created with condition Post=WAHR and Sponsor=FALSCH.</para>
        /// <para>Function TableToXlsx in class FromTable is called.</para>
        /// </summary>
        /// <param name="i_table_addresses">Table with addresses</param>
        /// <param name="i_output_file_name">Output file name</param>
        /// <param name="o_error">Error message</param>
        static public bool MailAdressesAsXslx(Table i_table_addresses, string i_output_file_name, out string o_error)
        {
            o_error = "";

            Table table_xlsx = new Table("Table with mail addresses for XLSX output");

            Row first_row = i_table_addresses.GetRow(0, out o_error);
            table_xlsx.AddRow(first_row, out o_error);
            table_xlsx.NumberColumns = first_row.NumberColumns; // TODO Should be done automatic by AddRow for first row

            for (int row_index = 1; row_index < i_table_addresses.NumberRows; row_index++)
            {
                string post_str = i_table_addresses.GetFieldString(row_index, "Post", out o_error);
                string sponsor_str = i_table_addresses.GetFieldString(row_index, "Sponsor", out o_error);

                if (post_str.CompareTo("WAHR") == 0 && sponsor_str.CompareTo("FALSCH") == 0)
                {
                    Row current_row = i_table_addresses.GetRow(row_index, out o_error);
                    table_xlsx.AddRow(current_row, out o_error);
                }
  
            }

            if (!MailAddressesModify(table_xlsx, out o_error))
            {
                return false;
            }

            if (!FromTable.TableToXlsx(table_xlsx, i_output_file_name, out o_error))
            {
                return false;
            }

            string error_message = @"";
            StartExcel(i_output_file_name, out error_message);

            return true;

        } // MailAdressesAsXslx

        /// <summary>Output mail addresses as CSV file
        /// <para>A Table with mail addresses for the CSV output is created with condition Post=WAHR and Sponsor=FALSCH.</para>
        /// <para>Function TableToCsv in class FromTable is called.</para>
        /// </summary>
        /// <param name="i_table_addresses">Table with addresses</param>
        /// <param name="i_output_file_name">Output file name</param>
        /// <param name="o_error">Error message</param>
        static public bool MailAdressesAsCsv(Table i_table_addresses, string i_output_file_name, out string o_error)
        {
            o_error = "";

            Table table_csv = new Table("Table with mail addresses for CSV output");

            Row first_row = i_table_addresses.GetRow(0, out o_error);
            table_csv.AddRow(first_row, out o_error);
            // ???? Changes ???? table_csv.NumberColumns = first_row.NumberColumns; // TODO Should be done automatic by AddRow for first row

            table_csv.NumberColumns = i_table_addresses.NumberColumns;

            for (int row_index = 1; row_index < i_table_addresses.NumberRows; row_index++)
            {
                string post_str = i_table_addresses.GetFieldString(row_index, "Post", out o_error);
                string sponsor_str = i_table_addresses.GetFieldString(row_index, "Sponsor", out o_error);

                if (post_str.CompareTo("WAHR") == 0 && sponsor_str.CompareTo("FALSCH") == 0)
                {
                    Row current_row = i_table_addresses.GetRow(row_index, out o_error);
                    table_csv.AddRow(current_row, out o_error);
                }

            }

            if (!MailAddressesModify(table_csv, out o_error))
            {
                return false;
            }

            if (!FromTable.TableToCsv(table_csv, i_output_file_name, ";", null, out o_error))
            {
                return false;
            }

            string error_message = @"";
            if (!StartExcel(i_output_file_name, out error_message))
            {
                o_error = "Output.MailAdressesAsCsv " + error_message;
                return false;
            }

            return true;

        } // MailAdressesAsCsv

        /// <summary>Modify output mail addresses
        /// <para>Only the address data (columns) that will be used for a mail will be kept</para>
        /// <para></para>
        /// </summary>
        /// <param name="i_table_mail_addresses">Table with addresses that will be modified (input and output)</param>
        /// <param name="o_error">Error message</param>
        static private bool MailAddressesModify(Table i_table_mail_addresses, out string o_error)
        {
            o_error = "";

            Row header_row = i_table_mail_addresses.GetRow(0, out o_error);
            if (null == header_row || o_error.Length > 0)
                return false;

            int index_email = Table.GetColumnIndex(header_row, "e-Mail");
            if (index_email <0 )
            {
                o_error = "Output.MailAddressesModify index_email < 0";
                return false;
            }

            int number_columns = i_table_mail_addresses.NumberColumns;

            for (int index_remove= number_columns-1; index_remove>= index_email;  index_remove--)
            {
                if (!TableTools.RemoveColumn(ref i_table_mail_addresses, index_remove, out o_error))
                    return false;
            }

            if (!TableTools.MoveColumn(ref i_table_mail_addresses, "FamilienName", "Vorname", out o_error))
            {
                return false;
            }
            
            return true;

        } // MailAddressesModify

        /// <summary>Output of list as text file with supporters for a given season
        /// <para>The fields (columns) that shall be outputted is defined in a hardcoded string array.</para>
        /// <para>Sizes of the output text fields and if the text shall be right or left aligned are also defined.</para>
        /// <para>Output records (supporters) are defined by the support sum (> 0 CHF) for the input season.</para>
        /// <para>The text file should be created in subdirectory (Output) in the exe directory.</para>
        /// <para>Supporter) data is written to the file.</para>
        /// <para>The output text file is opened with Notepad.</para>
        /// </summary>
        /// <param name="i_table_addresses">Table with addresses</param>
        /// <param name="i_season_column_name">The season that user has set</param>
        /// <param name="i_output_file_name">Output file name with full path</param>
        /// <param name="o_error">Error message</param>
        static public bool SupportListTxt(Table i_table_addresses, string i_season_column_name, string i_output_file_name, out string o_error)
        {
            o_error = "";

            string error_message = "";

            string[] used_columns = { "Vorname", "FamilienName", "Strasse", "Hausnummer", "PLZ", "Wohnort", i_season_column_name };
            string[] header_colums = { "Vorname", "FamilienName", "Strasse", "Hausn", "PLZ", "Wohnort", i_season_column_name.Substring(8) };
            string[] sum_colums = { "", "", "", "", "", "Summe:", "add..." };
            string[] number_colums = { "", "", "", "", "", "Anzahl:", "add..." };

            int[] size_columns = { 18, 18, 18, 8, 8, 15, 10 };
            bool[] right_aligned = { false, false, false, true, true, false, true };
            bool[] right_aligned_total = { false, false, false, true, true, true, true };
            string[] str_columns = { "", "", "", "", "", "", "" };

            double total_sum = 0.0;

            int total_number = 0;

            try
            {
                using (StreamWriter outfile = new StreamWriter(i_output_file_name))
                {
                    string output_header = _OutputSupportListRow(header_colums, size_columns, right_aligned);
                    outfile.Write(output_header);
                    outfile.Write(System.Environment.NewLine);

                    // Note start from row 1
                    for (int row_index = 1; row_index < i_table_addresses.NumberRows; row_index++)
                    {
                        string support_value_str = i_table_addresses.GetFieldString(row_index, i_season_column_name, out error_message);

                        if (support_value_str != "0")
                        {
                            str_columns[0] = i_table_addresses.GetFieldString(row_index, used_columns[0], out error_message);
                            str_columns[1] = i_table_addresses.GetFieldString(row_index, used_columns[1], out error_message);
                            str_columns[2] = i_table_addresses.GetFieldString(row_index, used_columns[2], out error_message);
                            str_columns[3] = i_table_addresses.GetFieldString(row_index, used_columns[3], out error_message);
                            str_columns[4] = i_table_addresses.GetFieldString(row_index, used_columns[4], out error_message);
                            str_columns[5] = i_table_addresses.GetFieldString(row_index, used_columns[5], out error_message);
                            str_columns[6] = support_value_str;

                            string output_row = _OutputSupportListRow(str_columns, size_columns, right_aligned);

                            outfile.Write(output_row);

                            total_sum = total_sum + System.Convert.ToDouble(support_value_str);

                            total_number = total_number + 1;
                        }
                    }

                    outfile.Write(System.Environment.NewLine);
                    outfile.Write(System.Environment.NewLine);

                    sum_colums[6] = total_sum.ToString();
                    number_colums[6] = total_number.ToString();

                    string output_total = _OutputSupportListRow(sum_colums, size_columns, right_aligned_total);
                    outfile.Write(output_total);
                    output_total = _OutputSupportListRow(number_colums, size_columns, right_aligned_total);
                    outfile.Write(output_total);
              }
            } // try

            catch (Exception e)
            {
                o_error = " Unhandled Exception " + e.GetType() + " occurred at " + DateTime.Now + "!";
                return false;
            }

            System.Diagnostics.Process.Start("notepad.exe", i_output_file_name);
            //System.Diagnostics.Process.Start("winword.exe", i_output_file_name);

            return true;

        } // SupportListTxt

        /// <summary>Output a list as XML file with supporters </summary>
        /// <param name="i_table_addresses">Table with addresses</param>
        /// <param name="i_season_column_name">The season that user has set</param>
        /// <param name="i_output_file_name">Output file name with full path</param>
        /// <param name="o_error">Error message</param>
        static public bool SupportersAsXml(Table i_table_addresses, string i_season_column_name, string i_output_file_name, out string o_error)
        {
            o_error = "";

            string error_message = "";

            ArrayList array_list_warnings = new ArrayList();

            Table table_xml = new Table("Table with supporters XML output");

            Row first_row = i_table_addresses.GetRow(0, out o_error);
            table_xml.AddRow(first_row, out o_error);
            table_xml.NumberColumns = first_row.NumberColumns; // TODO Should be done automatic by AddRow for first row

            int n_supporters = 0;

            for (int row_index = 1; row_index < i_table_addresses.NumberRows; row_index++)
            {
                string support_value_str = i_table_addresses.GetFieldString(row_index, i_season_column_name, out error_message);

                if (support_value_str != "0")
                {
                    Row current_row = i_table_addresses.GetRow(row_index, out o_error);
                    table_xml.AddRow(current_row, out o_error);

                    n_supporters = n_supporters + 1;
                }
            }

            if (n_supporters == 0)
            {
                string season_str = i_season_column_name.Substring(8);

                o_error = "Keine Supporter Saison " + season_str;

                return false;
            }

            string person_xml_tag = "Supporter";

            if (!FromTable.TableToXml(table_xml, i_output_file_name, person_xml_tag, System.Text.Encoding.UTF8, out o_error))
            {
                return false;
            }

            return true;

        } // SupportersAsXml

        /// <summary>Output a list as CSV Excel file with supporter addresses </summary>
        /// <param name="i_table_addresses">Table with addresses</param>
        /// <param name="i_season_column_name">The season that user has set</param>
        /// <param name="i_output_file_name">Output file name with full path</param>
        /// <param name="o_error">Error message</param>
        static public bool SupportAddressesAsCsv(Table i_table_addresses, string i_season_column_name, string i_output_file_name, out string o_error)
        {
            o_error = "";

            string error_message = "";

            ArrayList array_list_warnings = new ArrayList();

            Table table_csv = new Table("Table with supporter mail addresses for CSV output");

            Row first_row = i_table_addresses.GetRow(0, out o_error);
            table_csv.AddRow(first_row, out o_error);
            table_csv.NumberColumns = first_row.NumberColumns; // TODO Should be done automatic by AddRow for first row

            for (int row_index = 1; row_index < i_table_addresses.NumberRows; row_index++)
            {
                string support_value_str = i_table_addresses.GetFieldString(row_index, i_season_column_name, out error_message);
                string support_street_str = i_table_addresses.GetFieldString(row_index, "Strasse", out error_message);
                string support_city_str = i_table_addresses.GetFieldString(row_index, "Wohnort", out error_message);

                bool b_address_defined = true;
                if (support_street_str == "" || support_city_str == "")
                {
                    b_address_defined = false;
                }

                if (support_value_str != "0" && b_address_defined)
                {
                    Row current_row = i_table_addresses.GetRow(row_index, out o_error);
                    table_csv.AddRow(current_row, out o_error);
                }
                else if (support_value_str != "0" && !b_address_defined)
                {
                    string support_first_name_str = i_table_addresses.GetFieldString(row_index, "Vorname", out error_message);
                    string support_family_name_str = i_table_addresses.GetFieldString(row_index, "FamilienName", out error_message);
                    string support_email_str = i_table_addresses.GetFieldString(row_index, "e-Mail", out error_message);
                    string output_warning_line = "Keine Adresse für " + support_first_name_str + " " + support_family_name_str
                                                     + " " + support_email_str;
                    array_list_warnings.Add(output_warning_line);
                }
            }

            if (!MailAddressesModify(table_csv, out o_error))
            {
                return false;
            }

            if (!FromTable.TableToCsv(table_csv, i_output_file_name, ";", null, out o_error))
            {
                return false;
            }

            string[] warning_strings = (string[])array_list_warnings.ToArray(typeof(string));

            if (warning_strings.Length > 0)
            {
                _OutputWarningFile(warning_strings, i_output_file_name, out o_error);
            }

            StartExcel(i_output_file_name, out error_message);

            return true;

        } // SupportAddressesAsCsv

        /// <summary>Creates a file with warnings and opens it with notepad </summary>
        static private bool _OutputWarningFile(string[] i_warning_strings, string i_output_file_name, out string o_error)
        {
            o_error = "";

            if (i_warning_strings.Length == 0)
            {
                o_error = "_OutputWarningFile There are no warnings";
                return false;
            }
                string log_file = i_output_file_name + "_Log.txt";

            try
            {
                using (StreamWriter outfile = new StreamWriter(log_file))
                {

                    outfile.Write("Warnungen Warnungen  Warnungen Warnungen Warnungen Warnungen");
                    outfile.Write(System.Environment.NewLine);
                    outfile.Write(System.Environment.NewLine);

                    outfile.Write("File " + i_output_file_name);
                    outfile.Write(System.Environment.NewLine);
                    outfile.Write(System.Environment.NewLine);

                    for (int row_index = 0; row_index < i_warning_strings.Length; row_index++)
                    {
                        outfile.Write(i_warning_strings[row_index]);
                        outfile.Write(System.Environment.NewLine);
                    }

                    outfile.Write(System.Environment.NewLine);
                }
            }
            catch (Exception e)
            {
                o_error = " Unhandled Exception " + e.GetType() + " occurred at " + DateTime.Now + "!";
                return false;
            }

            System.Diagnostics.Process.Start("notepad.exe", log_file);

            return true;

        } // _OutputWarningFile

        /// <summary>Output list of sponsors as text file
        /// <para>The fields (columns) that shall be outputted is defined in a hardcoded string array.</para>
        /// <para>Sizes of the output text fields and if the text shall be right or left aligned are also defined.</para>
        /// <para>Output records (sponsors) are defined by the flag Sponsor=WAHR.</para>
        /// <para>The text file should be created in subdirectory (Output) in the exe directory.</para>
        /// <para>Sponsor data is written to the file. Function _OutputSupportListRow constructs the output line.</para>
        /// <para>The output text file is opened with Notepad.</para>
        /// </summary>
        /// <param name="i_table_addresses">Table with addresses</param>
        /// <param name="i_output_file_name">Output file name with path</param>
        /// <param name="o_error">Error message</param>
        static public bool SponsorListTxt(Table i_table_addresses, string i_output_file_name, out string o_error)
        {
            o_error = "";

            string error_message = "";

            string[] used_columns = { "Vorname", "FamilienName", "Strasse", "Hausnummer", "PLZ", "Wohnort" };
            string[] header_colums = { "Vorname", "FamilienName", "Strasse", "Hausn", "PLZ", "Wohnort" };

            int[] size_columns = { 18, 18, 18, 8, 8, 15 };
            bool[] right_aligned = { false, false, false, true, true, false };
            string[] str_columns = { "", "", "", "", "", "" };

            int total_number = 0;

            try
            {
                using (StreamWriter outfile = new StreamWriter(i_output_file_name))
                {
                    string output_header = _OutputSupportListRow(header_colums, size_columns, right_aligned);
                    outfile.Write(output_header);
                    outfile.Write(System.Environment.NewLine);

                    // Note start from row 1
                    for (int row_index = 1; row_index < i_table_addresses.NumberRows; row_index++)
                    {
                        string sponsor_value_str = i_table_addresses.GetFieldString(row_index, "Sponsor", out error_message);

                        if (sponsor_value_str == "WAHR")
                        {
                            str_columns[0] = i_table_addresses.GetFieldString(row_index, used_columns[0], out error_message);
                            str_columns[1] = i_table_addresses.GetFieldString(row_index, used_columns[1], out error_message);
                            str_columns[2] = i_table_addresses.GetFieldString(row_index, used_columns[2], out error_message);
                            str_columns[3] = i_table_addresses.GetFieldString(row_index, used_columns[3], out error_message);
                            str_columns[4] = i_table_addresses.GetFieldString(row_index, used_columns[4], out error_message);
                            str_columns[5] = i_table_addresses.GetFieldString(row_index, used_columns[5], out error_message);

                            string comment_str = i_table_addresses.GetFieldString(row_index, "Kommentar", out error_message);
                           
                            string output_row = _OutputSupportListRow(str_columns, size_columns, right_aligned);

                            outfile.Write(output_row);

                            outfile.Write("Kommentar: " + comment_str);
                            outfile.Write(System.Environment.NewLine);
                            outfile.Write(System.Environment.NewLine);

                            total_number = total_number + 1;
                        }
                    }

                    outfile.Write(System.Environment.NewLine);
                    outfile.Write(System.Environment.NewLine);
                }
            } // try

            catch (Exception e)
            {
                o_error = " Unhandled Exception " + e.GetType() + " occurred at " + DateTime.Now + "!";
                return false;
            }

            System.Diagnostics.Process.Start("notepad.exe", i_output_file_name);

            return true;

        } // SponsorListTxt

        /// <summary>Outputs a list of newsletter addressees as a text file
        /// <para>The fields (columns) that shall be outputted is defined in a hardcoded string array.</para>
        /// <para>Sizes of the output text fields and if the text shall be right or left aligned are also defined.</para>
        /// <para>Output records are defined by the flag NewsletterJazz=WAHR.</para>
        /// <para>The text file should be created in subdirectory (Output) in the exe directory.</para>
        /// <para>Function _OutputSupportListRow constructs the output line.</para>
        /// <para>The output text file is opened with Notepad.</para>
        /// </summary>
        /// <param name="i_table_addresses">Table with addresses</param>
        /// <param name="i_output_file_name">Output file name with path</param>
        /// <param name="o_error">Error message</param>
        static public bool NewsletterListTxt(Table i_table_addresses, string i_output_file_name, out string o_error)
        {
            o_error = "";

            string error_message = "";

            string[] used_columns = { "Vorname", "FamilienName", "e-Mail" };
            string[] header_colums = { "Vorname", "FamilienName", "E-Mail Adresse" };

            int[] size_columns = { 28, 28, 38 };
            bool[] right_aligned = { false, false, false };
            string[] str_columns = { "", "", "" };

            int total_number = 0;

            try
            {
                using (StreamWriter outfile = new StreamWriter(i_output_file_name))
                {
                    string output_header = _OutputSupportListRow(header_colums, size_columns, right_aligned);
                    outfile.Write(output_header);
                    outfile.Write(System.Environment.NewLine);

                    // Note start from row 1
                    for (int row_index = 1; row_index < i_table_addresses.NumberRows; row_index++)
                    {
                        string newsletter_value_str = i_table_addresses.GetFieldString(row_index, "NewsletterJazz", out error_message);

                        if (newsletter_value_str == "WAHR")
                        {
                            str_columns[0] = i_table_addresses.GetFieldString(row_index, used_columns[0], out error_message);
                            str_columns[1] = i_table_addresses.GetFieldString(row_index, used_columns[1], out error_message);
                            str_columns[2] = i_table_addresses.GetFieldString(row_index, used_columns[2], out error_message);

                            string output_row = _OutputSupportListRow(str_columns, size_columns, right_aligned);

                            outfile.Write(output_row);

                            total_number = total_number + 1;
                        }
                    }

                    outfile.Write(System.Environment.NewLine);
                    outfile.Write(System.Environment.NewLine);

                    string total_number_string = total_number.ToString();

                    outfile.Write("         Anzahl Newsletter Empfänger: " + total_number_string + System.Environment.NewLine);

                }
            } // try

            catch (Exception e)
            {
                o_error = " Unhandled Exception " + e.GetType() + " occurred at " + DateTime.Now + "!";
                return false;
            }

            System.Diagnostics.Process.Start("notepad.exe", i_output_file_name);

            return true;

        } // NewsletterListTxt

        /// <summary>Output list of check results as a text file
        /// <para>The fields (columns) that shall be outputted is defined in a hardcoded string array.</para>
        /// <para>Sizes of the output text fields and if the text shall be right or left aligned are also defined.</para>
        /// <para>Output records are defined by the flags Post=FALSCH, NewsletterJazz=FALSCH and Sponsor=FALSCH.</para>
        /// <para>The text file should be created in subdirectory (Output) in the exe directory.</para>
        /// <para>Sponsor data is written to the file. Function _OutputSupportListRow constructs the output line.</para>
        /// <para>The output text file is opened with Notepad.</para>
        /// </summary>
        /// <param name="i_table_addresses">Table with addresses</param>
        /// <param name="i_output_file_name">Output file name with path</param>
        /// <param name="i_season_column_name">Season column name</param>
        /// <param name="o_error">Error message</param>
        static public bool CheckListTxt(Table i_table_addresses, string i_output_file_name, string i_season_column_name, out string o_error)
        {
            o_error = "";

            string error_message = "";

            string[] used_columns = { "Vorname", "FamilienName", "Strasse", "Hausnummer", "PLZ", "Wohnort" };
            string[] header_colums = { "Vorname", "FamilienName", "Strasse", "Hausn", "PLZ", "Wohnort" };

            int[] size_columns = { 18, 18, 18, 8, 8, 15 };
            bool[] right_aligned = { false, false, false, true, true, false };
            string[] str_columns = { "", "", "", "", "", "" };

            int total_number = 0;

            try
            {
                using (StreamWriter outfile = new StreamWriter(i_output_file_name))
                {
                    string output_header = _OutputSupportListRow(header_colums, size_columns, right_aligned);
                    outfile.Write(output_header);
                    outfile.Write(System.Environment.NewLine);

                    // Note start from row 1
                    for (int row_index = 1; row_index < i_table_addresses.NumberRows; row_index++)
                    {
                        string newsletter_value_str = i_table_addresses.GetFieldString(row_index, "NewsletterJazz", out error_message);
                        string mail_value_str = i_table_addresses.GetFieldString(row_index, "Post", out error_message);
                        string sponsor_value_str = i_table_addresses.GetFieldString(row_index, "Sponsor", out error_message);

                        string support_value_str = i_table_addresses.GetFieldString(row_index, i_season_column_name, out error_message);

                        if (newsletter_value_str == "FALSCH" && mail_value_str == "FALSCH" && sponsor_value_str == "FALSCH" && support_value_str == "0")
                        {
                            str_columns[0] = i_table_addresses.GetFieldString(row_index, used_columns[0], out error_message);
                            str_columns[1] = i_table_addresses.GetFieldString(row_index, used_columns[1], out error_message);
                            str_columns[2] = i_table_addresses.GetFieldString(row_index, used_columns[2], out error_message);
                            str_columns[3] = i_table_addresses.GetFieldString(row_index, used_columns[3], out error_message);
                            str_columns[4] = i_table_addresses.GetFieldString(row_index, used_columns[4], out error_message);
                            str_columns[5] = i_table_addresses.GetFieldString(row_index, used_columns[5], out error_message);

                            string output_row = _OutputSupportListRow(str_columns, size_columns, right_aligned);

                            outfile.Write(output_row);

                            outfile.Write("Warnung: Kriegt keinen Versand oder Newsletter und ist nicht Supporter oder Sponsor");
                            outfile.Write(System.Environment.NewLine);
                            outfile.Write(System.Environment.NewLine);

                            total_number = total_number + 1;
                        }
                    }

                    outfile.Write(System.Environment.NewLine);
                    outfile.Write(System.Environment.NewLine);
                }
            } // try

            catch (Exception e)
            {
                o_error = " Unhandled Exception " + e.GetType() + " occurred at " + DateTime.Now + "!";
                return false;
            }

            System.Diagnostics.Process.Start("notepad.exe", i_output_file_name);

            return true;

        } // CheckListTxt
        /// <summary>Returns one row for an output list
        /// <para>Column values (text fields) that are too long will be shortened</para>
        /// </summary>
        /// <param name="i_str_columns">Array with column values</param>
        /// <param name="i_size_columns">Array with column sizes</param>
        /// <param name="i_right_aligned">Array with flags defining how the text shall be aligned</param>
        static private string _OutputSupportListRow(string[] i_str_columns, int[] i_size_columns, bool[] i_right_aligned)
        {
            string str_row = "";

            for (int col_index = 0; col_index < i_str_columns.Length; col_index++)
            {
                string str_column = i_str_columns[col_index];

                str_column = str_column.Trim();

                if (str_column.Length > i_size_columns[col_index] - 1)
                {
                    str_column = str_column.Substring(0, i_size_columns[col_index]-4) + "...";
                }

                if (i_right_aligned[col_index])
                {
                    bool first_char = true;
                    for (int add_index_right = str_column.Length; add_index_right < i_size_columns[col_index]; add_index_right++)
                    {
                        if (first_char)
                        {
                            str_column = str_column + " ";
                            first_char = false;
                        }
                        else
                        {
                            str_column = " " + str_column;
                        }
                    }
                }
                else
                {
                    for (int add_index_left = str_column.Length; add_index_left < i_size_columns[col_index]; add_index_left++)
                    {
                        str_column = str_column + " ";
                    }
                }


                str_row = str_row + str_column;

            }

            str_row = str_row + System.Environment.NewLine;

            return str_row;

        } // _OutputSupportListRow


        /// <summary>Starts Excel</summary>
        /// <param name="i_output_file_name">Output file name with full path</param>
        /// <param name="o_error">Error message</param>
        static private bool StartExcel(string i_output_file_name, out string o_error)
        {
            o_error = @"";

            try
            {
                System.Diagnostics.Process.Start("EXCEL.exe", i_output_file_name);
            }

            catch (Exception e)
            {
                o_error = "Output.StartExcel Unhandled Exception " + e.GetType();
                return false;
            }

            return true;
        } // StartExcel

        /// <summary>Checks if file is locked</summary>
        /// <param name="i_file_name">File name with full path</param>
        static public bool FileIsLocked(string i_file_name)
        {
            string exception_message = @"";
            bool b_file_locked = false;
            System.IO.FileStream file_stream;
            try
            {
                file_stream = System.IO.File.Open(i_file_name, FileMode.OpenOrCreate, FileAccess.Read, FileShare.None);
                file_stream.Close();
            }
            catch (System.IO.IOException ex)
            {
                exception_message = ex.Message;
                b_file_locked = true;
            }

            return b_file_locked;

        } // FileIsLocked

    } // class Output

} // namespace
