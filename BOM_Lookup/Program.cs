using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using ExcelDataReader;
using System.Data;
using System.Windows.Forms;
using System.Drawing;
using Microsoft.WindowsAPICodePack.Dialogs;
using System.Reflection;
using System.Diagnostics;
using System.Threading;

namespace BOM_Lookup
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            bool test = false;
            string excelPath;
            string assemblyPath;
            string partPath;
            string customPath;
            string destPath;
            string textPath;

            if (test == false)
            {
                excelPath = @"C:\Users\Admin\Desktop\test.xlsx";
                assemblyPath = @"\\10.30.30.5\mechdesign\ASSEMBLIES\PDF DXF STEP";
                partPath = @"\\10.30.30.5\mechdesign\PARTS\PDF DXF STEP";
                customPath = @"C:\Users\Admin\Desktop";
                destPath = @"C:\Users\Admin\Desktop";
                textPath = @"C:\Users\Admin\Desktop";
            }
            else
            {
                excelPath = @"C:\Users\Admin\Desktop\test.xlsx";
                assemblyPath = @"Z:\mechdesign\ASSEMBLIES\PDF DXF STEP";
                partPath = @"Z:\mechdesign\PARTS\PDF DXF STEP";
                customPath = @"C:\Users\Admin\Desktop";
                destPath = @"C:\Users\Admin\Desktop";
                textPath = @"C:\Users\Admin\Desktop";
            }

            bool custom_folder_used = false;

            Assembly assembly = Assembly.GetExecutingAssembly();

            FileVersionInfo fileVersionInfo = FileVersionInfo.GetVersionInfo(assembly.Location);
            string version = fileVersionInfo.ProductVersion;

            Console.Title = $"BOM Part Assembly Lookup Program - Version: {version}";

            string[] array = new string[1000];
            string[] array_flat = new string[1000];
            int row_length = 1;
            int counter = 0;
            int counter_n = 0;

            excelPath = FilePicker("Open Excel File");
            destPath = FolderPicker("Choose a folder to copy files to");
            textPath = destPath + "\\log_" + DateTime.Now.ToString("yyyy_M_dd_HH_mm_ss") + ".txt";

            //See whether custom folder is used
            Console.WriteLine("Looking for parts/assemblies in custom folder? Y/N");
            string custom_folder = Console.ReadLine();
            if (custom_folder == "Y" | custom_folder == "y")
            { 
                customPath = FolderPicker("Choose custom part/assemblies folder");
                custom_folder_used=true;
            }

            using (var stream = File.Open(excelPath, FileMode.Open, FileAccess.Read))
            {
                try
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {

                        int i = 0;
                        int max_sheets = 0;
                        int sheet_selection = 0, column;
                        string column_s = "a";
                        bool valid = false;

                        WritetoFile(textPath, $"BOM Lookup version: {version}");
                        WritetoFile(textPath, DateTime.Now.ToString("yyyy_M_dd_HH_mm_ss"));
                        WritetoFile(textPath, excelPath);

                        Console.WriteLine("Sheets found:");
                        //Print all sheets:
                        while (reader.Name != null)
                        {
                            Console.WriteLine($"{i + 1}." + reader.Name);
                            reader.NextResult();
                            max_sheets = i + 1;
                            i++;
                        }
                        reader.Reset();

                        while (valid == false)
                        {
                            //Prompt user to select sheet and go to it:
                            Console.Write("\nPlease select a sheet: ");
                            sheet_selection = Int32.Parse(Console.ReadLine());
                            if (sheet_selection > 0 & sheet_selection <= max_sheets)
                            {
                                valid = true;
                            }
                            else
                            {
                                Console.Write("Sheet selection out of bounds!");
                            }
                        }

                        reader.Reset();

                        int k = 0;
                        while (k < sheet_selection - 1)
                        {
                            reader.NextResult();
                            k++;
                        }

                        Console.WriteLine($"Sheet selected: {reader.Name}");
                        WritetoFile(textPath, $"\nSheet selected: {reader.Name}");

                        string name = reader.Name;

                        Console.Write("\nChoose a column (eg. A): ");

                        column_s = Console.ReadLine();
                        column = AlphatoNum(column_s);

                        WritetoFile(textPath, "Column selected: " + column_s);

                        row_length = reader.RowCount;

                        do
                        {
                            int j = 0;

                            while (reader.Read())
                            {
                                if (reader.Name == name)
                                {
                                    array[j] = reader.GetString(column - 1);

                                    if (array[j] == null) { break; }

                                    int part_length = array[j].Length;
                                    array_flat[j] = array[j].Insert(part_length - 2, "_flat");

                                    j++;
                                }
                            }
                        } while (reader.NextResult());
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine($"The file could not be opened:\n '{e}'");
                    WritetoFile(textPath, $"The file could not be opened:\n '{e}'");

                    Console.WriteLine("\nPress enter to exit.");
                    Console.ReadLine();
                    Environment.Exit(0);
                }
            }

            //int flat_counter = 0;
            //int part_length;
            //foreach (string part in array)
            //{
            //    part_length = part.Length;
            //    array_flat[flat_counter] = part.Insert(part_length - 2, "_flat");
            //    flat_counter++;
            //}

            Console.WriteLine("\nLooking for Parts/Assemblies/pdfs: \n");
            WritetoFile(textPath, "\nThe following files were not found: \n");

            //initalising strings
            string dxf, pdf, stp;
            string dxfPath, pdfPath, stpPath;
            string dxfPathPart, pdfPathPart, stpPathPart;
            string custom_dxfPath, custom_pdfPath, custom_stpPath;

            string dxf_flat, pdf_flat;
            string dxfPathflat, pdfPathflat, dxfPathPartflat, pdfPathPartflat, custom_dxfPathflat, custom_pdfPathflat;

            for (int i = 0; i < row_length; i++)
            {
                bool dxf_found = false, pdf_found = false, stp_found = false;

                dxf = array[i] + ".dxf";
                pdf = array[i] + ".pdf";
                stp = array[i] + ".stp"; //added step file

                dxfPath = assemblyPath + "\\" + dxf;
                pdfPath = assemblyPath + "\\" + pdf;
                stpPath = assemblyPath + "\\" + stp; //added step file
                dxfPathPart = partPath + "\\" + dxf;
                pdfPathPart = partPath + "\\" + pdf;
                stpPathPart = partPath + "\\" + stp; //added step file

                custom_dxfPath = customPath + "\\" + dxf;
                custom_pdfPath = customPath + "\\" + pdf;
                custom_stpPath = customPath + "\\" + stp; //added step file

                //SE21 now occasionally adds flat to the filename
                dxf_flat = array_flat[i] + ".dxf";
                pdf_flat = array_flat[i] + ".pdf";

                dxfPathflat = assemblyPath + "\\" + dxf_flat;
                pdfPathflat = assemblyPath + "\\" + pdf_flat;
                dxfPathPartflat = partPath + "\\" + dxf_flat;
                pdfPathPartflat = partPath + "\\" + pdf_flat;

                custom_dxfPathflat = customPath + "\\" + dxf_flat;
                custom_pdfPathflat = customPath + "\\" + pdf_flat;

                //Check which dxf is the newest:
                DateTime[] dates_dxf = new DateTime[6];

                dates_dxf[0] = File.GetLastWriteTime(dxfPath);
                dates_dxf[1] = File.GetLastWriteTime(dxfPathPart);
                dates_dxf[2] = File.GetLastWriteTime(custom_dxfPath);
                dates_dxf[3] = File.GetLastWriteTime(dxfPathflat);
                dates_dxf[4] = File.GetLastWriteTime(dxfPathPartflat);
                dates_dxf[5] = File.GetLastWriteTime(custom_dxfPathflat);

                int newest_dxf = NewestFile(dates_dxf);

                //Check which pdf is the newest:
                DateTime[] dates_pdf = new DateTime[6];

                dates_pdf[0] = File.GetLastWriteTime(pdfPath);
                dates_pdf[1] = File.GetLastWriteTime(pdfPathPart);
                dates_pdf[2] = File.GetLastWriteTime(custom_dxfPath);
                dates_pdf[3] = File.GetLastWriteTime(pdfPathflat);
                dates_pdf[4] = File.GetLastWriteTime(pdfPathPartflat);
                dates_pdf[5] = File.GetLastWriteTime(custom_pdfPathflat);

                int newest_pdf = NewestFile(dates_pdf);

                //Check which stp is the newest:
                DateTime[] dates_stp = new DateTime[6];

                dates_stp[0] = File.GetLastWriteTime(stpPath);
                dates_stp[1] = File.GetLastWriteTime(stpPathPart);
                dates_stp[2] = File.GetLastWriteTime(custom_stpPath);

                int newest_stp = NewestFile(dates_stp);

                if (custom_folder_used==true) 
                {
                    if ((System.IO.File.Exists(custom_dxfPath)) & (newest_dxf==2))
                    {
                        Console.WriteLine("Copying: " + dxf + "...");
                        //WritetoFile(textPath, "Copying: " + dxf + "...");
                        System.IO.File.Copy(custom_dxfPath, (destPath + "\\" + dxf), true);
                        dxf_found = true;
                        counter++;
                    }

                    if (System.IO.File.Exists(custom_pdfPath)&(newest_pdf ==2))
                    {
                        Console.WriteLine("Copying: " + pdf + "...");
                        //WritetoFile(textPath, "Copying: " + pdf + "...");
                        System.IO.File.Copy(custom_pdfPath, (destPath + "\\" + pdf), true);
                        pdf_found = true;
                        counter++;
                    }

                    if (System.IO.File.Exists(custom_stpPath) & (newest_stp == 2))
                    {
                        Console.WriteLine("Copying: " + stp + "...");
                        //WritetoFile(textPath, "Copying: " + pdf + "...");
                        System.IO.File.Copy(custom_stpPath, (destPath + "\\" + stp), true);
                        stp_found = true;
                        counter++;
                    }

                    //flat version in custom folder path
                    if (System.IO.File.Exists(custom_dxfPathflat)&(newest_dxf==5))
                    {
                        Console.WriteLine("Copying: " + dxf_flat + "...");
                        //WritetoFile(textPath, "Copying: " + dxf + "...");
                        System.IO.File.Copy(custom_dxfPathflat, (destPath + "\\" + dxf_flat), true);
                        dxf_found = true;
                        counter++;
                    }

                    if (System.IO.File.Exists(custom_pdfPathflat)&(newest_pdf==5))
                    {
                        Console.WriteLine("Copying: " + pdf_flat + "...");
                        //WritetoFile(textPath, "Copying: " + pdf + "...");
                        System.IO.File.Copy(custom_pdfPathflat, (destPath + "\\" + pdf_flat), true);
                        pdf_found = true;
                        counter++;
                    }
                }

                if (System.IO.File.Exists(dxfPath)&(newest_dxf==0))
                {
                    Console.WriteLine("Copying: " + dxf + "...");
                    //WritetoFile(textPath, "Copying: " + dxf + "...");
                    System.IO.File.Copy(dxfPath, (destPath + "\\" + dxf),  true);
                    dxf_found = true;
                    counter++;
                }

                if (System.IO.File.Exists(pdfPath)&(newest_pdf==0))
                {
                    Console.WriteLine("Copying: " + pdf + "...");
                    //WritetoFile(textPath, "Copying: " + pdf + "...");
                    System.IO.File.Copy(pdfPath, (destPath + "\\" + pdf),true);
                    pdf_found = true;
                    counter++;
                }

                if (System.IO.File.Exists(stpPath) & (newest_stp == 0))
                {
                    Console.WriteLine("Copying: " + stp + "...");
                    //WritetoFile(textPath, "Copying: " + pdf + "...");
                    System.IO.File.Copy(stpPath, (destPath + "\\" + stp), true);
                    stp_found = true;
                    counter++;
                }

                if (System.IO.File.Exists(dxfPathPart)&(newest_dxf==1))
                {
                    Console.WriteLine("Copying: " + dxf + "...");
                    //WritetoFile(textPath, "Copying: " + dxf + "...");
                    System.IO.File.Copy(dxfPathPart, (destPath + "\\" + dxf), true);
                    dxf_found = true;
                    counter++;
                }

                if (System.IO.File.Exists(pdfPathPart)&(newest_pdf==1))
                {
                    Console.WriteLine("Copying: " + pdf + "...");
                    //WritetoFile(textPath, "Copying: " + pdf + "...");
                    System.IO.File.Copy(pdfPathPart, (destPath + "\\" + pdf),   true);
                    pdf_found = true;
                    counter++;
                }

                if (System.IO.File.Exists(stpPathPart) & (newest_stp == 1))
                {
                    Console.WriteLine("Copying: " + stp + "...");
                    //WritetoFile(textPath, "Copying: " + pdf + "...");
                    System.IO.File.Copy(stpPathPart, (destPath + "\\" + stp), true);
                    stp_found = true;
                    counter++;
                }

                //FLAT VARIANT
                if (System.IO.File.Exists(dxfPathflat)&(newest_dxf==3))
                {
                    Console.WriteLine("Copying: " + dxf_flat + "...");
                    //WritetoFile(textPath, "Copying: " + dxf + "...");
                    System.IO.File.Copy(dxfPath, (destPath + "\\" + dxf), true);
                    dxf_found = true;
                    counter++;
                }

                if (System.IO.File.Exists(pdfPathflat)&(newest_pdf==3))
                {
                    Console.WriteLine("Copying: " + pdf_flat + "...");
                    //WritetoFile(textPath, "Copying: " + pdf + "...");
                    System.IO.File.Copy(pdfPathflat, (destPath + "\\" + pdf_flat), true);
                    pdf_found = true;
                    counter++;
                }

                if (System.IO.File.Exists(dxfPathPartflat)&(newest_dxf==4))
                {
                    Console.WriteLine("Copying: " + dxf_flat + "...");
                    //WritetoFile(textPath, "Copying: " + dxf + "...");
                    System.IO.File.Copy(dxfPathPartflat, (destPath + "\\" + dxf_flat), true);
                    dxf_found = true;
                    counter++;
                }

                if (System.IO.File.Exists(pdfPathPartflat)&(newest_pdf==4))
                {
                    Console.WriteLine("Copying: " + pdf_flat + "...");
                    //WritetoFile(textPath, "Copying: " + pdf + "...");
                    System.IO.File.Copy(pdfPathPartflat, (destPath + "\\" + pdf_flat), true);
                    pdf_found = true;
                    counter++;
                }

                if (dxf_found == false)
                {
                    Console.WriteLine($"{dxf} was not found.");
                    WritetoFile(textPath, $"{dxf}");
                    counter_n++;
                }

                if (pdf_found == false)
                {
                    Console.WriteLine($"{pdf} was not found.");
                    WritetoFile(textPath, $"{pdf}");
                    counter_n++;
                }

                if (stp_found == false)
                {
                    Console.WriteLine($"{stp} was not found.");
                    WritetoFile(textPath, $"{stp}");
                    counter_n++;
                }
            }
            
            if (counter == 0) 
            {
                Console.WriteLine($"\nNo files were copied.");
                WritetoFile(textPath, $"\nNo files were copied.");
            }
            else if (counter == 1) 
            {
                Console.WriteLine($"\n1 file was copied.");
                WritetoFile(textPath, $"\n1 file was copied.");
            }
            else 
            {
                Console.WriteLine($"\n{counter} files were copied.");
                WritetoFile(textPath, $"\n{counter} files were copied.");
            }

            if (counter_n == 0) 
            {
                Console.WriteLine($"All files were found.\n");
                WritetoFile(textPath, $"All files were found.\n");
            }
            else if (counter_n == 1) 
            {
                Console.WriteLine($"1 file was not found.\n");
                WritetoFile(textPath, $"1 file was not found.\n");
            }
            else 
            {
                Console.WriteLine($"{counter_n} files were not found.\n");
                WritetoFile(textPath, $"{counter_n} files were not found.\n");
            }

            System.Diagnostics.Process.Start(destPath);
            System.Diagnostics.Process.Start(textPath);
            
            System.Threading.Timer t = new System.Threading.Timer(timerC, null, 3000, 3000);

            Console.WriteLine("Closing window in 3 seconds.");
            Console.ReadLine();
        }
        public static int NewestFile(DateTime[] dates)
        {
            DateTime x = dates.Max();
            return Array.IndexOf(dates, x);
        }
        private static void timerC(object state)
        {
            Environment.Exit(0);
        }

        public static string FilePicker(string title)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();

            dialog.Title = title;

            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                //yolo
            }
            return @"" + dialog.FileName;
        }

        public static string FolderPicker(string title)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = true;

            dialog.Title = title;

            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                //yolo
            }
            return @"" + dialog.FileName;
        }

        public static int AlphatoNum(string s)
        {
            switch (s)
            {
                case "a":
                    return 1;
                    break;
                case "A":
                    return 1;
                    break;
                case "b":
                    return 2;
                    break;
                case "B":
                    return 2;
                    break;
                case "c":
                    return 3;
                    break;
                case "C":
                    return 3;
                    break;
                case "d":
                    return 4;
                    break;
                case "D":
                    return 4;
                    break;
                case "e":
                    return 5;
                    break;
                case "E":
                    return 5;
                    break;
                case "f":
                    return 6;
                    break;
                case "F":
                    return 6;
                    break;
                case "g":
                    return 7;
                    break;
                case "G":
                    return 7;
                    break;
                case "h":
                    return 8;
                    break;
                case "H":
                    return 8;
                    break;
                case "i":
                    return 9;
                    break;
                case "I":
                    return 9;
                    break;
                case "j":
                    return 10;
                    break;
                case "J":
                    return 10;
                    break;
                case "k":
                    return 11;
                    break;
                case "K":
                    return 11;
                    break;
                case "l":
                    return 12;
                    break;
                case "L":
                    return 12;
                    break;
                case "1":
                    return 1;
                    break;
                case "2":
                    return 2;
                case "3":
                    return 3;
                case "4":
                    return 4;
                case "5":
                    return 5;
                case "6":
                    return 6;
                case "7":
                    return 7;
                case "8":
                    return 8;
                case "9":
                    return 9;
                case "10":
                    return 10;
            }
            return 0;
        }

        public static void WritetoFile(string path, string text)
        {
            using (System.IO.StreamWriter file =
            new System.IO.StreamWriter(@path, true))
            {
                file.WriteLine(text);
            }
        }
    }
}
