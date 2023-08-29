using Aspose.Cells.Drawing;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Reflection;

namespace ABSARecon
{
    public class Program
    {
        public static void Main(string[] args)
        {
            List<VISADATA> accountDetails = new List<VISADATA>();
            List<CardCentre> cardDetails = new List<CardCentre>();
            int counter = 0;
            try
            {
                Console.WriteLine("----------------Start--------------------");

                Console.WriteLine("-----------------------------------------");

                Console.WriteLine("Start Time ------------->   " + DateTime.Now);
                string source = ConfigurationManager.AppSettings["inputPath"];
                string destination = ConfigurationManager.AppSettings["destination"];
                string output = ConfigurationManager.AppSettings["outputPath"];
                string backup = ConfigurationManager.AppSettings["backup"];
                string visaRecon = ConfigurationManager.AppSettings["visaRecon"];
                string report = ConfigurationManager.AppSettings["report"];
                string sms = ConfigurationManager.AppSettings["sms"];
                string tempSave = ConfigurationManager.AppSettings["tempSave"];

                if (!UserFunctions.KillAllExcelInstaces())
                {
                    Console.WriteLine(" ");
                    Console.WriteLine("Unable to kill all excel instance");
                    Task.Factory.StartNew(() => UserFunctions.WriteLog(" ", " ", "Unable to kill all excel instance", ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));
                }

                Console.WriteLine(" ");
                Console.WriteLine("Excel instances killed successfully");
                Task.Factory.StartNew(() => UserFunctions.WriteLog(" ", " ", "Excel instances killed successfully", ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));

                int fileCount = UserFunctions.GetFileCountInFolder(source);

                UserFunctions.ReadAllFiles(source, out List<FileDetails> fileDetails);


                if (!fileDetails.Any())
                {
                    Console.WriteLine("No data found in location");
                    Task.Factory.StartNew(() => UserFunctions.WriteLog(" ", " ", "No data found in location", ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));
                    Thread.Sleep(10000);
                    return;
                }
                Console.WriteLine(" ");
                Task.Factory.StartNew(() => UserFunctions.WriteLog(" ", " ", "Data read from file successfully", ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));
                Console.WriteLine("Data read from file successfully");
                
                foreach (var item in fileDetails)
                {
                    string filePath = item.FilePath;
                    string fileName = Path.GetFileNameWithoutExtension(filePath);
                    string jsonInput = UserFunctions.ReadExcelToJson(filePath, destination, fileName);
                    string message = "";

                    if (string.IsNullOrEmpty(jsonInput))
                    {
                        Console.WriteLine("Unable to read data from " + filePath);
                        Task.Factory.StartNew(() => UserFunctions.WriteLog(item.FileNameWithoutExtension, " ", "Unable to read data from " + filePath, ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));
                        Thread.Sleep(10000);
                        return;
                    }

                    Console.WriteLine(" ");
                    Console.WriteLine("Data read from json successfully successfully");
                    Task.Factory.StartNew(() => UserFunctions.WriteLog(item.FileNameWithoutExtension, " ", "Data read from json successfully successfully", ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));

                    if (fileName.Contains(StaticVariables.CardCentreAccounts))
                    {
                       UserFunctions.ReadJsonTwo(jsonInput, out cardDetails);
                       UserFunctions.WriteToSheetTwo(visaRecon, tempSave, cardDetails, out message);

                       UserFunctions.MoveFile(item.FilePath, backup + Path.GetFileName(item.FilePath));
                    }

                    else
                    {
                        
                        UserFunctions.ReadJson(jsonInput, out accountDetails);

                        UserFunctions.CleanUpData(accountDetails, out List<CleanedData> cleanData, out message);

                        UserFunctions.RemoveDuplicates(cleanData, out List<CleanedDataTwo> removedDuplicate, out List<CleanedDataThree> duplicateRemoved, out string messages);
                        UserFunctions.GetDuplicates(cleanData, out List<CleanedData> getDuplicates, out message);

                        
                        string[] smss;

                        try
                        {
                            string[] lines = File.ReadAllLines(sms); // Read all lines into an array

                            smss = new string[lines.Length]; // Initialize the array size

                            for (int index = 0; index < lines.Length; index++)
                            {
                                string line = lines[index];
                               
                                smss[index] = line; // Store the line in the array
                            }
                            
                            UserFunctions.WriteToSheet(tempSave, tempSave, removedDuplicate, duplicateRemoved, smss[counter], fileCount, counter, out message);

                            
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error: " + ex.Message);
                        }

                        

                        if (string.IsNullOrEmpty(jsonInput))
                        {
                            Console.WriteLine(message + " " + filePath);
                            Task.Factory.StartNew(() => UserFunctions.WriteLog(item.FileNameWithoutExtension, " ", message + " " + filePath, ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));
                            Thread.Sleep(10000);
                            return;
                        }

                        Console.WriteLine(" ");
                        Console.WriteLine(message);
                        Task.Factory.StartNew(() => UserFunctions.WriteLog(item.FileNameWithoutExtension, " ", message, ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));


                        string generatedFile = "Clean Data " + DateTime.Now.ToString("yyyy-MM-dd hh_mm_ss");

                        UserFunctions.CreateExcel(generatedFile, new List<string> { JsonConvert.SerializeObject(removedDuplicate) }, out string outputFile, output);
                        





                        UserFunctions.MoveFile(item.FilePath, backup + Path.GetFileName(item.FilePath));

                        counter += 1;
                        
                    }

                }

                
                UserFunctions.KillAllExcelInstaces();
                string oldFileName = "Report" + "." + "xlsx";
                string newFileName = DateTime.Now.ToString("dd-MM-yyyy hh mm ") + "." + "xlsx";
                UserFunctions.RenameAndMoveExcelFile(tempSave, report,oldFileName, newFileName);

            }
            catch (Exception ex)
            {
                Task.Factory.StartNew(() => UserFunctions.WriteLog(" ", " ", ex.Message + "  || " + ex.StackTrace, ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));
                Console.WriteLine(" ");
                Console.WriteLine("Exception -------------------->    " + ex.Message + "  || " + ex.StackTrace);
            }
            Console.WriteLine("");

            Console.WriteLine(accountDetails.Count + " files  process and completed @ " + DateTime.Now);
            Console.WriteLine("");
            Console.WriteLine("Process completed");
            Thread.Sleep(15000);
        }
    }
}

