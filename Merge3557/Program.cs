/*******************************************************************************
 * Program: MergePDF3557.exe
 * Description: Parses through the JH Extract CSV file and creates merged PDF 
 *              files based on Batch Names. Then it puts the merged PDF files
 *              in one ZIPPED file located in RemitWebFiles folder.
 *              
 *              Modeled from MergePDF4524.exe
 * Author: Thomas Sugimoto
 * Creation Date: 11/2/2020
 * 
 * Updates Modifications
 * 11/02/2020 - krp Initial Build.
 * ****************************************************************************/

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace Merge3557
{
    class Program
    {
        static List<Transactions> TList = new List<Transactions>();

        static List<string> CONP_and_CONP_CORR = new List<string>();
        static List<string> CORR = new List<string>();
        static List<string> CSIG_and_CMLT = new List<string>();
        
        static string ImageFolderPDF = string.Empty;
        static string ImageFolderPayments = string.Empty;
        static string BatchType = string.Empty;
        static string NEWLINE = Environment.NewLine;
        static string PaymentsExistFile = string.Empty;
        static string NoPaymentsFile = string.Empty;
        static string ZIP_PROGRAM = string.Empty;
        static string YYMMDD = string.Empty;
        static string YYYYMMDD = string.Empty;
        static string BOX = string.Empty;

        /// <summary>
        /// Main process
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            if (args.Length != 2)
            {
                Console.WriteLine(@"ABORT: Wrong number of arguments!");
                Console.WriteLine(@"USAGE: MergePDF3557 <YYMMDD> <BOX>");
                Environment.Exit(-1);
            }

            // Get and set the location of 7ZIP
            SetZipProgramPath();

            YYMMDD = args[0];
            YYYYMMDD = $"20{YYMMDD}";

            
            BOX = args[1];
            string JHFolder = string.Format(@"e:\jh\{0}\{1}", YYMMDD, BOX);
            string MetaFile = Path.Combine(JHFolder, String.Format(@"MD{0}.csv", YYMMDD));
            PaymentsExistFile = string.Format(@"e:\{0}\{1}\PAYMENTSEXIST.SENT", YYMMDD, BOX);
            NoPaymentsFile = string.Format(@"e:\{0}\{1}\NOPAYMENTS.SENT", YYMMDD, BOX);

            ImageFolderPDF = string.Format(@"e:\{0}\{1}\PDFS", YYMMDD, BOX);
            string ZippedFileName = $@"{ImageFolderPDF}\{YYYYMMDD}_PDFs.zip";
            try
            {
                #region Verify files and folders
                if (!Directory.Exists(JHFolder))
                {
                    Console.WriteLine(@"ABORT: {0} folder does NOT exist!", JHFolder);
                    Environment.Exit(-1);
                }

                if (!File.Exists(MetaFile))
                {
                    Console.WriteLine(@"ABORT: {0} file does NOT exist!", MetaFile);
                    Environment.Exit(-1);
                }

                if (!Directory.Exists(string.Format(@"e:\{0}\{1}", YYMMDD, BOX)))
                    Directory.CreateDirectory(string.Format(@"e:\{0}\{1}", YYMMDD, BOX));

                // Remove and create image folder.
                if (Directory.Exists(ImageFolderPDF))
                {
                    try
                    {
                        Directory.Delete(ImageFolderPDF, true);
                        Thread.Sleep(3000); // Give the folder time to mourn.
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(@"ABORT: Could not delete the folder: {0}!", ImageFolderPDF);
                        Console.WriteLine(@"Message: {0}", e.Message);
                        Environment.Exit(-1);
                    }
                }

                // Create the directory to hold the Merged PDF files
                try
                {
                    Directory.CreateDirectory(ImageFolderPDF);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(@"ABORT: Could not create the folder: {0}!", ImageFolderPDF);
                    Console.WriteLine(@"Message: {0}", ex.Message);
                    Environment.Exit(-1);
                }
                #endregion

                //Parse through the meta data file and add all records to a list.
                Console.WriteLine($"Parsing through meta file: {MetaFile}...");
                ParseCSV(MetaFile);
                Console.WriteLine($"    Finished parsing file.{NEWLINE}");
                string OutPutFile = string.Empty;

                // Add records to the correct lists
                Console.WriteLine($"Adding records to individual lists...");

                foreach (var t in TList.Where(x => x.MODE.Equals("CSIG") || x.MODE.Equals("CMLT")).GroupBy(a => a.FileName))
                {
                    CSIG_and_CMLT.Add(t.Key.ToString());
                }
                
                foreach (var t in TList.Where(x => x.MODE.Equals("CONP") || x.MODE.Equals("CONP CORR")).GroupBy(a => a.FileName))
                {
                    CONP_and_CONP_CORR.Add(t.Key.ToString());
                }
                foreach (var t in TList.Where(x => x.MODE.Equals("CORR")).GroupBy(a => a.FileName))
                {
                    CORR.Add(t.Key.ToString());
                }
                
                Console.WriteLine($"    Finished adding records to individual lists.{NEWLINE}");

                // Merge the files from each List into their own output file
                Console.WriteLine($"Creating merged PDF files...");
                if (CSIG_and_CMLT.Count > 0)
                {
                    OutPutFile = $@"{ImageFolderPDF}\{BOX}_{YYYYMMDD}_CPN.pdf";
                    if (File.Exists(OutPutFile)) { File.Delete(OutPutFile); }
                    MergePdf(OutPutFile, CSIG_and_CMLT);
                }
                if (CONP_and_CONP_CORR.Count > 0)
                {
                    OutPutFile = $@"{ImageFolderPDF}\{BOX}_{YYYYMMDD}_CORR.pdf";
                    if (File.Exists(OutPutFile)) { File.Delete(OutPutFile); }
                    MergePdf(OutPutFile, CONP_and_CONP_CORR);
                }
                if (CORR.Count > 0)
                {
                    OutPutFile = $@"{ImageFolderPDF}\{BOX}_{YYYYMMDD}_CORR-2.pdf";
                    if (File.Exists(OutPutFile)) { File.Delete(OutPutFile); }
                    MergePdf(OutPutFile, CORR);
                }
                Console.WriteLine($"    Finished creating merged PDF files.{NEWLINE}");

                // ZIP the merged files
                //ZIPALL(ZippedFileName);
            }
            catch (Exception exMain)
            {
                Console.WriteLine($"ABORT: {exMain.Message}");
            }
            Console.WriteLine("Finished running MergePDF3557");
        }

        /// <summary>
        ///     Calls the class PDFMerge to merge all the files into one output file per envelope
        /// </summary>
        /// <param name="output">File name to save the merged pdf files into.</param>
        /// <param name="files">List of all files belonging with an envelope.</param>
        static protected void MergePdf(string output, List<string> files)
        {
            PDFMerge.Merge(files.ToArray(), output);
        }

        /// <summary>
        ///  Adds all the records from CSV file to a List so we can sort them correctly.
        /// </summary>
        /// <param name="MetaFile">Meta Data file of RemitPlus data for the current date</param>
        static protected void ParseCSV(string MetaFile)
        {
            string line = string.Empty;
            string[] SplitMe;
            string MODE = string.Empty;
            string ProcessDate = string.Empty;
            string BankAccount = string.Empty;
            string CheckNo = string.Empty;
            double AmountInCents;
            string DocumentID = string.Empty;
            string SeqNo = string.Empty;
            string PageType = string.Empty;
            string FileName = string.Empty;
            string NewFileName = string.Empty;
            Int32 BatchNo;
            bool HasChecks = false;
            try
            {
                // Read the meta file
                using (StreamReader sr = new StreamReader(MetaFile))
                {
                    sr.ReadLine(); // Skip the Header line
                    while (!sr.EndOfStream)
                    {
                        line = sr.ReadLine();
                        if (line.ToUpper().Contains("CHECK"))
                        {
                            HasChecks = true;
                        }
                        SplitMe = line.Split('~');
                        ProcessDate = SplitMe[0];
                        BatchNo = Convert.ToInt32(SplitMe[1]);
                        SeqNo = SplitMe[2];
                        MODE = SplitMe[3];
                        BankAccount = SplitMe[6];
                        CheckNo = SplitMe[7];
                        AmountInCents = (Convert.ToDouble(SplitMe[9]) * 100);
                        PageType = SplitMe[19];
                        FileName = SplitMe[41];

                        // Add records to list
                        TList.Add(new Transactions()
                        {
                            ProcessDate = ProcessDate,
                            BatchNo = BatchNo,
                            SeqNo = SeqNo,
                            MODE = MODE,
                            BankAccount = BankAccount,
                            CheckNo = CheckNo,
                            AmountInCents = AmountInCents,
                            PageType = PageType,
                            FileName = FileName
                        });
                    }

                    // Sort the list by MODE, BatchNo, SeqNo
                    TList.OrderBy(x => x.MODE).ThenBy(x => x.BatchNo).ThenBy(x => x.SeqNo);
                    //sr.Close();
                }

                // Create xxx.SENT file based on found checks
                //if (HasChecks == true)
                //{
                //    using (StreamWriter sw = new StreamWriter(PaymentsExistFile))
                //    {
                //        sw.WriteLine($"{DateTime.Now}");
                //        sw.WriteLine($"Checks were found in export file!");
                //    }
                //}
                //else
                //{
                //    using (StreamWriter sw = new StreamWriter(NoPaymentsFile))
                //    {
                //        sw.WriteLine($"{DateTime.Now}");
                //        sw.WriteLine($"Zero checks were found in export file!");
                //    }
                //}

            }
            catch (Exception ex)
            {
                Console.WriteLine($"ERROR: {ex.Message}");
            }
        }

        /// <summary>
        /// Searches for 7Zip on users machine and sets a variable for its location.
        /// </summary>
        static private void SetZipProgramPath()
        {
            try
            {
                if (File.Exists(string.Format(@"C:\Progra~1\7-Zip\7z.exe")))
                {
                    ZIP_PROGRAM = string.Format(@"C:\Progra~1\7-Zip\7z.exe");
                }
                else if (File.Exists(string.Format(@"C:\Progra~2\7-Zip\7z.exe")))
                {
                    ZIP_PROGRAM = string.Format(@"C:\Progra~2\7-Zip\7z.exe");
                }
                else
                {
                    Console.WriteLine("I can not find 7Zip on this computer. Files NOT sent!");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// Adds the files to one zip file.
        /// </summary>
        /// <param name="ZIPFILENAME"></param>
        /// <param name="FTPFILENAME"></param>
        static private void ZIPALL(string ZIPFILENAME)
        {
            if (File.Exists(ZIPFILENAME))
            {
                File.Delete(ZIPFILENAME);
            }
            String WorkDir = $@"E:\{YYMMDD}\{BOX}";
            string ImagesFolder = $@"{ImageFolderPDF}";

            DirectoryInfo D = new DirectoryInfo(ImagesFolder);
            foreach (FileInfo f in D.GetFiles("*.pdf", SearchOption.TopDirectoryOnly))
            {
                try
                {
                    Process zip = new Process();
                    zip.StartInfo.FileName = ZIP_PROGRAM;
                    zip.StartInfo.Arguments = "a -tzip \"" + ZIPFILENAME + "\" \"" + f.FullName.ToString() + "\" -mx=9";
                    zip.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    zip.Start();
                    zip.WaitForExit();
                    zip.Dispose();
                    Thread.Sleep(1000);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"ERROR creating ZIP File: {ZIPFILENAME}! " + ex.Message);
                }
            }


            if (File.Exists(ZIPFILENAME))
            {
                Console.WriteLine($"Finished creating the ZIP file!{NEWLINE}");
                using (StreamWriter sw = new StreamWriter(Path.Combine(WorkDir, "ZIPFILECREATED.SENT")))
                {
                    sw.WriteLine($"Zip File was created at {DateTime.Now.ToString()}");
                    sw.Flush();

                }
            }
        } // End of ZIPALL
    }
}
