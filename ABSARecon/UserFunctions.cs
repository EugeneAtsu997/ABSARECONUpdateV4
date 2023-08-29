using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Aspose.Cells.Utility;
using Aspose.Cells;
using System.Net.Mail;
using System.Net;
using System.Runtime.InteropServices.ComTypes;
using Spire.Xls.Collections;
using System.Text.RegularExpressions;
using Aspose.Cells.Charts;
using System.Diagnostics.Metrics;
using System.Transactions;
using Spire.Xls;
using Aspose.Cells.Drawing;

namespace ABSARecon
{
    public static class UserFunctions
    {
        public static bool KillAllExcelInstaces()
        {
            bool worked = false;
            try
            {
                Process[] process = Process.GetProcessesByName("Excel");

                foreach (Process p in process)
                {
                    if (!string.IsNullOrEmpty(p.ProcessName))
                    {
                        try
                        {
                            p.Kill();
                            worked = true;
                        }
                        catch
                        {

                        }
                    }
                }
                worked = true;
            }
            catch (Exception ex)
            {
                Task.Factory.StartNew(() => WriteLog(" ", "", ex.Message + "  || " + ex.StackTrace, ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));

            }
            return worked;
        }
        public static void MoveFile(string sourcePath, string destinationPath)
        {
            try
            {
                if (!File.Exists(sourcePath))
                {
                    // This statement ensures that the file is created,  
                    // but the handle is not kept.  
                    using (FileStream fs = File.Create(sourcePath)) { }
                }
                // Ensure that the target does not exist.  
                if (File.Exists(destinationPath))
                    File.Delete(destinationPath);
                // Move the file.  
                File.Move(sourcePath, destinationPath);

                Task.Factory.StartNew(() => WriteLog(sourcePath, destinationPath, string.Format("{0} was moved to {1}.", sourcePath, destinationPath), ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));


                // See if the original exists now.  
                if (File.Exists(sourcePath))
                {
                    Task.Factory.StartNew(() => WriteLog(sourcePath, destinationPath, "The original file still exists, which is unexpected.", ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));

                }
                else
                {
                    Task.Factory.StartNew(() => WriteLog(sourcePath, destinationPath, "The original file no longer exists, which is expected.", ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));

                }
            }
            catch (Exception ex)
            {
                Task.Factory.StartNew(() => WriteLog(sourcePath, destinationPath, ex.Message + "  || " + ex.StackTrace, ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));
            }
        }

        public static void WriteLog(string sescureId, string request, string response, string serviceame, string mfunctionName, [CallerMemberName] string callerName = "")
        {
            mfunctionName = callerName;
            string logFilePath = "C:\\Logs\\" + serviceame + "\\";
            logFilePath = logFilePath + "Log-" + DateTime.Today.ToString("MM-dd-yyyy") + "." + "txt";
            try
            {
                using (FileStream fileStream = new FileStream(logFilePath, FileMode.Append))
                {
                    FileInfo logFileInfo;

                    logFileInfo = new FileInfo(logFilePath);
                    DirectoryInfo logDirInfo = new DirectoryInfo(logFileInfo.DirectoryName);
                    if (!logDirInfo.Exists) logDirInfo.Create();

                    StreamWriter log = new StreamWriter(fileStream);

                    if (!logFileInfo.Exists)
                    {
                        _ = logFileInfo.Create();
                    }
                    else
                    {
                        log.WriteLine(sescureId);
                        log.WriteLine(DateTime.Now.ToString());
                        log.WriteLine(request);
                        log.WriteLine(response);
                        log.WriteLine(mfunctionName);
                        log.WriteLine("_________________________________________________________________________________________________________");
                        log.Close();
                    }
                    fileStream.Close();
                    fileStream.Dispose();
                }
            }
            catch (Exception)
            {


            }

        }
        public static void ReadAllFiles(string sourcePath, out List<FileDetails> fileDetails)
        {
            fileDetails = new List<FileDetails>();

            try
            {
                foreach (string file in Directory.EnumerateFiles(sourcePath, "*.xlsx"))
                {
                    fileDetails.Add(new FileDetails
                    {

                        FileNameWithoutExtension = Path.GetFileNameWithoutExtension(file),
                        FilePath = file
                    });

                }
            }
            catch (Exception ex)
            {
                Task.Factory.StartNew(() => WriteLog(" ", "", ex.Message + "  || " + ex.StackTrace, ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));
            }

        }
        public static void ReadJson(string jsonInput, out List<VISADATA> accountDetails)
        {
            accountDetails = new List<VISADATA>();
            try
            {
                accountDetails = JsonConvert.DeserializeObject<List<VISADATA>>(jsonInput);
            }
            catch (Exception ex)
            {
                Task.Factory.StartNew(() => WriteLog(" ", " ", ex.Message + "  || " + ex.StackTrace, ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));

            }
        }

        public static void ReadJsonTwo(string jsonInput, out List<CardCentre> cardDetails)
        {
            cardDetails = new List<CardCentre>();
            try
            {
                cardDetails = JsonConvert.DeserializeObject<List<CardCentre>>(jsonInput);
            }
            catch (Exception ex)
            {
                Task.Factory.StartNew(() => WriteLog(" ", " ", ex.Message + "  || " + ex.StackTrace, ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));

            }
        }


        public static string ReadExcelToJson(string inputPath, string destination, string fileName)
        {
            string jsonInput = string.Empty;

            try
            {
                var workbook = new Aspose.Cells.Workbook(inputPath);

                string jsonPath = destination + fileName + ".json";
                workbook.Save(jsonPath);

                workbook.Dispose();

                jsonInput = File.ReadAllText(jsonPath);

            }
            catch (Exception ex)
            {
                Task.Factory.StartNew(() => WriteLog(inputPath, "", ex.Message + "  || " + ex.StackTrace, ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));
            }
            return jsonInput;
        }



        public static void ExcelUpdateAction(string workbookPath)
        {
            try
            {
                Application excelApp = new Application
                {
                    DisplayAlerts = false
                };

                Microsoft.Office.Interop.Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(workbookPath, 0, false, 5, "", "", false, XlPlatform.xlWindows, "", true, false, 0, true, false, false);

                Microsoft.Office.Interop.Excel.Sheets worksheets = excelWorkbook.Worksheets;

                excelWorkbook.Sheets["Evaluation Warning"].Delete();

                excelWorkbook.Save();

                excelWorkbook.Close();

                Marshal.ReleaseComObject(worksheets);

                excelApp.Quit();
            }
            catch (Exception)
            {

            }
        }

        public static bool CreateExcel(string fileName, List<string> jsonInput, out string filePath, string generatedExcelPath = null)
        {
            filePath = string.Empty;
            bool worked = false;
            if (string.IsNullOrEmpty(generatedExcelPath))
            {
                generatedExcelPath = ConfigurationManager.AppSettings["backup"];
            }

            string workbookPath = string.Empty;
            try
            {
                // Create a Workbook object
                Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();
                WorksheetCollection worksheets = workbook.Worksheets;

                int counter = 1;

                foreach (var item in jsonInput)
                {


                    if (counter == 1)
                    {
                        Aspose.Cells.Worksheet worksheet = workbook.Worksheets[0];

                        // Set JsonLayoutOptions
                        JsonLayoutOptions options = new JsonLayoutOptions
                        {
                            ArrayAsTable = true
                        };

                        // Import JSON Data
                        JsonUtility.ImportData(item, worksheet.Cells, 0, 0, options);
                    }
                    else
                    {
                        Aspose.Cells.Worksheet worksheet = worksheets.Add("Sheet" + counter);
                        // Set JsonLayoutOptions
                        JsonLayoutOptions options = new JsonLayoutOptions
                        {
                            ArrayAsTable = true
                        };

                        // Import JSON Data
                        JsonUtility.ImportData(item, worksheet.Cells, 0, 0, options);
                    }

                    counter++;
                }




                // Save Excel file

                filePath = generatedExcelPath + fileName + ".xlsx";
                workbookPath = filePath;
                workbook.Save(filePath);

                workbook.Dispose();
                worked = true;



                Task.Factory.StartNew(() => WriteLog(fileName, JsonConvert.SerializeObject(jsonInput), "File Created Successfully", ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));

            }
            catch (Exception ex)
            {
                Task.Factory.StartNew(() => WriteLog(" ", fileName, ex.Message + "  || " + ex.StackTrace, "Error", string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));

            }
            Task.Factory.StartNew(() => ExcelUpdateAction(workbookPath));
            
            return worked;
        }



        public static bool CleanUpData(List<VISADATA> accountDetails, out List<CleanedData> cleanData, out string message)
        {
            cleanData = new List<CleanedData>();
            message = "Unable to process this request";

            bool worked = false;
            try
            {

                accountDetails = accountDetails.Where(x => x != null && !string.IsNullOrEmpty(x.NUM) && x.CUR == StaticVariables.GHS && !string.IsNullOrEmpty(x.D)).OrderByDescending(x => x.NUM)/*.GroupBy(x => x.NUMBER).Where(group => group.Count() == 1).SelectMany(group => group)*/.ToList();

                var data = accountDetails.Where(x => x.NUM.Length < 3).ToList();

                data.ForEach(data => data.ConvertNumberToInteger = int.Parse(data.NUM));

                accountDetails = data.OrderBy(x => x.ConvertNumberToInteger).ToList();




                cleanData = accountDetails.Select(x => new CleanedData
                {
                    AMOUNT = x.AMOUNT,
                    AMOUNTUS = x.AMOUNTUS,
                    CARDNUMBER = x.CARDNUMBER,
                    CODE = x.CODE,
                    CUR = x.CUR,
                    D = x.D,
                    DATE = x.DATE,
                    NUM = x.NUM,
                    NUMBER = x.NUMBER,
                    TIME = x.TIME

                }).ToList();



                worked = true;
                message = "Request processed successfully";
            }
            catch (Exception)
            {


            }
            return worked;
        }

        public static bool GetDuplicates(List<CleanedData> accountDetails, out List<CleanedData> getDuplicates, out string message)
        {
            getDuplicates = new List<CleanedData>();
            bool success = false;
            message = "Duplicates not gotten";

            try
            {
                getDuplicates = accountDetails.GroupBy(x => x.NUMBER).Where(xx => xx.Count() > 1).SelectMany(x => x.ToList()).ToList();


                success = true;
                message = "Duplicates gotten successfully";
            }

            catch (Exception)
            {


            }
            return success;

        }


        public static bool RemoveDuplicates(List<CleanedData> accountDetails, out List<CleanedDataTwo> removedDuplicate, out List<CleanedDataThree> duplicateRemoved, out string message)
        {
            removedDuplicate = new List<CleanedDataTwo>();
            duplicateRemoved = new List<CleanedDataThree>();
            message = "Unable to remove data";
            bool success = false;


            try
            {
                var duplicates = accountDetails.GroupBy(x => x.NUMBER).Where(xx => xx.Count() <= 1).SelectMany(x => x.ToList()).ToList();

                removedDuplicate = duplicates.Select(x => new CleanedDataTwo
                {
                    AMOUNT = x.AMOUNT,
                    AMOUNTUS = x.AMOUNTUS,
                    CARDNUMBER = x.CARDNUMBER,
                    CUR = x.CUR,
                    D = x.D,
                    DATE = x.DATE,
                    NUM = x.NUM,
                    NUMBER = x.NUMBER,
                    TIME = x.TIME

                }).ToList();

                duplicateRemoved = duplicates.Select(x => new CleanedDataThree
                {
                    AMOUNT = x.AMOUNT,
                    AMOUNTUS = x.AMOUNTUS,
                    CARDNUMBER = x.CARDNUMBER,
                    CUR = x.CUR,
                    D = x.D,
                    DATE = x.DATE,
                    NUM = x.NUM,
                    NUMBER = x.NUMBER,
                    TIME = x.TIME

                }).ToList();

                success = true;
                message = "Request processes successfully";
            }
            catch (Exception ex)
            {
            }

            return success;
        }

        //Move data to new sheet on report template
        public static bool WriteToReconSheet(List<string> jsonInput, string excelPath, string savePath, out string message) //(string fileName, List<string> jsonInput, out string filePath, string generatedExcelPath = null)
        {
            bool success = false;
            message = "Unable to write to excel";
            string save = "";

            try
            {
                //string[] file = Directory.GetFiles(excelPath);

                //foreach (string file2 in file)
                //{
                    var workbook = new Aspose.Cells.Workbook(excelPath);
                    var worksheet = workbook.Worksheets.Add("Sheet2");

                    foreach (var item in jsonInput)
                    {
                        // Set JsonLayoutOptions
                        JsonLayoutOptions options = new JsonLayoutOptions
                        {
                            ArrayAsTable = true
                        };

                        // Import JSON Data
                        JsonUtility.ImportData(item, worksheet.Cells, 0, 0, options);
                    }

                    string fileName = "Clean Data " + DateTime.Now.ToString("yyyy-MM-dd hh_mm_ss") + "." + "xlsx"; //$"{DateTime.Now:dd.MM.yyyy}_bot.xlsx";  DateTime.Today.ToString("dd-MM-yyyy") + "." + "txt"
                    save = savePath + "\\" + fileName;
                    workbook.Save(save);

                //}
                success = true;
                message = "Write to excel was successful";


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                //Task.Factory.StartNew(() => WriteLog(" ", fileName, ex.Message + "  || " + ex.StackTrace, "Error", string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));

            }
            Task.Factory.StartNew(() => ExcelUpdateAction(save));
            return success;

        }


        // Sum the amount for day on sheet
        public static bool AmtEachDay(List<CleanedDataTwo> accountDetails, out List<CleanedDataTwo> amtEachDay, out string message)
        {
            amtEachDay = new List<CleanedDataTwo>();
            message = "Unable to sum numbers";
            bool success = false;

            try
            {

                //Get distinct date
                var getDistintDate = accountDetails.DistinctBy(x => x.DATE).ToList();
                

                //loop through distinct date using key

                foreach (var d in getDistintDate)
                {
                    //Sum amount by key
                    var Transactions = accountDetails.Where(x => x.DATE == d.DATE).ToList();

                    foreach (var item in Transactions)
                    {
                        amtEachDay.Add(item);
                    }

                    var amount = Transactions.Sum(x => decimal.Parse(x.AMOUNT));

                    //Create a clean object and use sum amount as amount
                    amtEachDay.Add(new CleanedDataTwo
                    {
                        DATE = "",
                        AMOUNT = amount.ToString(),
                        AMOUNTUS = "",
                        CARDNUMBER = "",
                        CUR = "",
                        D = "",
                        NUM = "",
                        NUMBER = "",
                        TIME = ""
                    });

                    //Add another row with empty 
                    amtEachDay.Add(new CleanedDataTwo
                    {
                        DATE = "",
                        AMOUNT = "",
                        AMOUNTUS = "",
                        CARDNUMBER = "",
                        CUR = "",
                        D = "",
                        NUM = "",
                        NUMBER = "",
                        TIME = ""
                    });
                }
                success = true;
                message = "Summation Successfully";
            }
            catch (Exception)
            {
            }

            return success;
        }

        public static bool SortedData(List<CardCentre> data, out List<cleanCardCentre> sortedData, out string message)
        {
            sortedData = new List<cleanCardCentre>();
            bool success = false;
            message = "Unable to sort data";

            try
            {
                data = data.Where(x => x.short_name == StaticVariables.ATMVISACASH && x.transaction_code == 11).ToList();

                foreach (var item in data)
                {

                    string pattern = @"(\d{6})(\d{4})";
                    var Narrative = item.narrative.ToString();
                    Match match = Regex.Match(Narrative, pattern);
                    if (match.Success)
                    {
                        item.short_name = $"{match.Groups[1].Value}-{match.Groups[2].Value}";
                    }

                    var curAmt = item.currency_amount.ToString();
                    item.currency_amount = curAmt.TrimStart('-');
                }

                sortedData = data.Select(x => new cleanCardCentre
                {


                    narrative = x.narrative,
                    Narrative = x.short_name,
                    currency_amount = x.currency_amount,
                    stmnt_date_and_time = x.stmnt_date_and_time
                }).ToList();

                success = true;
                message = "Data sorted successfully";


            }

            catch (Exception)
            {
            }
            return success;

        }


        //records with .99 at the end of the amount
        public static bool AmtWith99(List<CleanedDataTwo> data, out List<CleanedDataTwo> amtWith99, out string message)
        {
            amtWith99 = new List<CleanedDataTwo>();
            bool success = false;
            message = "Amount with .99 not found";

            try
            {
                //duplicates = accountDetails.GroupBy(x => x.NUMBER).Where(xx => xx.Count() <= 1).SelectMany(x => x.ToList()).ToList();
                amtWith99 = data.Where(x => x.AMOUNT.EndsWith(".99")).ToList();

                success = true;
                message = "Amount with .99 found";

            }
            catch (Exception)
            {
            }
            return success;
        }

        //subtract 25.99 from amount that ends with .99

        public static bool Subtract25(List<CleanedDataThree> data, out List<CleanedDataThree> subtract25, out string message)
        {
            subtract25 = new List<CleanedDataThree>();
            bool success = false;
            message = "Unable to subtract";

            try
            {

                foreach (var dd in data)
                {
                    if (dd.AMOUNT.EndsWith(".99"))
                    {
                        dd.AMOUNT = (decimal.Parse(dd.AMOUNT) - 25.99m).ToString();
                    }
                }
                subtract25 = data;

                success = true;
                message = "Subtraction done successfully";


            }
            catch (Exception)
            {
            }
            return success;
        }

        private static List<string> ListSheets(Aspose.Cells.Workbook workbook)

        {
            List<string> result = new List<string>();
            int index = 0;

            Aspose.Cells.Worksheet thisWorksheet = workbook.Worksheets[0];

            foreach (Aspose.Cells.Worksheet worksheet in workbook.Worksheets)

            {

                thisWorksheet.Cells[index, 0].Value = worksheet.Name;

                index++;

                result.Add(worksheet.Name);

            }
            return result;

        }
        public static bool WriteToSheet(string excelPath, string savePath, List<CleanedDataTwo> removedDuplicate, List<CleanedDataThree> duplicateRemoved, string sheetName, int fileCount, int counter, out string message)
        {
            KillAllExcelInstaces();
            bool success = false;
            message = "Unable to write to excel";
            string save = "";
            string filepath = "";
            try
            {
                string[] files = Directory.GetFiles(excelPath);
                /*foreach (string file in files)
                {
                    filepath = file;
                }*/
                var workbook = new Aspose.Cells.Workbook(/*filepath*/files[0]);

                // Get the worksheet collection
               WorksheetCollection worksheets = workbook.Worksheets;
                foreach (Aspose.Cells.Worksheet worksheet in worksheets)
                {
                    if (worksheet.Name == StaticVariables.VEP_745)
                    {
                        Subtract25(duplicateRemoved, out List<CleanedDataThree> subtract25, out message);
                        GetHalfDay(subtract25, out List<CleanedDataThree> getHalfDay, fileCount, counter, out message);

                        UpdateExcel(workbook, new List<string> { JsonConvert.SerializeObject(getHalfDay) }, StaticVariables.VEP_745);

                    }

                    if(worksheet.Name == StaticVariables.ACCESSFEE)
                    {
                        AmtWith99(removedDuplicate, out List<CleanedDataTwo> amtWith99, out message);
                        UpdateExcel(workbook, new List<string> { JsonConvert.SerializeObject(amtWith99) }, StaticVariables.ACCESSFEE);
                    }

                }
               
                var ws = workbook.Worksheets.Add(sheetName);
                AmtEachDay(removedDuplicate, out List<CleanedDataTwo> amtEachDay, out message);
                UpdateExcel(workbook, new List<string> { JsonConvert.SerializeObject(amtEachDay) }, sheetName);


                string fileName = "Report"+ "." + "xlsx"; 
                save = savePath + "\\" + fileName;
                workbook.Save(save);
                workbook.Dispose();
                
                


                success = true;
                message = "Write to excel was successful";


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                //Task.Factory.StartNew(() => WriteLog(" ", fileName, ex.Message + "  || " + ex.StackTrace, "Error", string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));

            }
            Task.Factory.StartNew(() => ExcelUpdateAction(save));
            KillAllExcelInstaces();
            return success;
        }

        //
        public static bool WriteToSheetTwo(string excelPath, string savePath, List<CardCentre> cardDetails, out string message)
        {
            bool success = false;
            message = "Unable to write to excel";
            string saveO = "";
            try
            {
                var workbook = new Aspose.Cells.Workbook(excelPath);
                
                SortedData(cardDetails, out List<cleanCardCentre> sortedData, out message);
                UpdateExcel(workbook, new List<string> { JsonConvert.SerializeObject(sortedData) }, StaticVariables.A009);
                
                string fileName = "Report" + "." + "xlsx";
                saveO = savePath + "\\" + fileName;
               
                
                workbook.Save(saveO);
                workbook.Dispose();
                KillAllExcelInstaces();




                success = true;
                message = "Write to excel was successful";


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                //Task.Factory.StartNew(() => WriteLog(" ", fileName, ex.Message + "  || " + ex.StackTrace, "Error", string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));

            }
            Task.Factory.StartNew(() => ExcelUpdateAction(saveO));
            KillAllExcelInstaces();
            return success;
        }

        public static void UpdateExcel(Aspose.Cells.Workbook workbook, List<string> jsonInput, string sheetName)
        {
            try
            {
                foreach (var item in jsonInput)
                {
                    Aspose.Cells.Worksheet workSheet = workbook.Worksheets[sheetName];

                    // Find the last row index with data in the worksheet
                    int lastRowIndex = workSheet.Cells.MaxDataRow + 1;

                    // Set JsonLayoutOptions
                    JsonLayoutOptions options = new JsonLayoutOptions
                    {
                        ArrayAsTable = true
                    };

                    // Import JSON Data
                    JsonUtility.ImportData(item, workSheet.Cells, lastRowIndex, 0, options);
                }
            }
            catch (Exception)
            {
            }
            

        }


         

        public static string FileExtension(string filepath)
        {
            string fileExtension = "";
            string result = "";
            try
            {
                
                string[] files = Directory.GetFiles(filepath);
                foreach (string file in files)
                {
                    fileExtension = Path.GetExtension(file);

                    if (fileExtension.Length >= 3)
                    {
                        string extractedSubstring = fileExtension.Substring(1,3);
                         
                        result = "SMS" + extractedSubstring;
                    }
                }
                    
            }
            catch (Exception)
            {
            }
            return result;
        }

        public static bool DoesWorksheetExist(Aspose.Cells.Workbook workbook, string sheetName)
        {
            return workbook.Worksheets.Any(sheet => sheet.Name == sheetName);
        }

        //rename and move excel file
        public static void RenameAndMoveExcelFile(string sourceFolderPath, string destinationFolderPath, string oldFileName, string newFileName)
        {
            try
            {
                string sourceFilePath = Path.Combine(sourceFolderPath, oldFileName);
                string destinationFilePath = Path.Combine(destinationFolderPath, newFileName);

                if (File.Exists(sourceFilePath))
                {
                    // Rename and move the file
                    File.Move(sourceFilePath, destinationFilePath);
                    Console.WriteLine("File renamed and moved successfully.");
                }
                else
                {
                    Console.WriteLine("Source file not found.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }


        public static void DeleteWorksheet(Aspose.Cells.Workbook workbook, string sheetName)
        {
            Aspose.Cells.Worksheet worksheet = workbook.Worksheets[sheetName];

            if (worksheet != null)
            {
                workbook.Worksheets.RemoveAt(workbook.Worksheets.IndexOf(worksheet));
                //workbook.Save("path_to_save_updated_workbook.xlsx");
            }
        }

        public static void DeleteWorksheetOne(Aspose.Cells.Workbook workbook, string savepath)
        {
            //Aspose.Cells.Worksheet worksheet = workbook.Worksheets[sheetName];
            WorksheetCollection worksheets = workbook.Worksheets;

            foreach(Aspose.Cells.Worksheet worksheet in worksheets)
            {
                if(worksheet.Name.Contains("Evaluation Warning"))
                {
                    workbook.Worksheets.RemoveAt(workbook.Worksheets.IndexOf(worksheet));
                }
            }

            string fileName = "Report" + "." + "xlsx";
            string save = savepath + "\\" + fileName;
        }

        //get number of file in folder
        public static int GetFileCountInFolder(string folderPath)
        {
            try
            {
                // Get a list of files in the folder
                string[] files = Directory.GetFiles(folderPath);

                // Return the count of files
                return files.Length - 1;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
                return -1; // Return -1 to indicate an error
            }
        }

        //Get half day from visa sheet
        public static bool GetHalfDay(List<CleanedDataThree> data, out List<CleanedDataThree> getHalfDay, int fileCount, int counter, out string message)
        {
            getHalfDay = new List<CleanedDataThree>();
            bool success = false;
            message = "Unable to get half day";

            try
            {
                if(fileCount == 2)
                {
                    if(counter == 0)
                    {
                        var getDistinctDate = data.DistinctBy(x => x.DATE).ToList();
                        var dd = getDistinctDate.ElementAt(1);
                        var Transactions = data.Where(x => x.DATE == dd.DATE).ToList();
                        getHalfDay = Transactions;

                    }

                    if(counter == 1)
                    {
                        var getDistinctDate = data.DistinctBy(x => x.DATE).ToList();
                        var dd = getDistinctDate.ElementAt(0);
                        var Transactions = data.Where(x => x.DATE == dd.DATE).ToList();
                        getHalfDay = Transactions;
                    }
                }

               

                else
                {
                    if (counter == 0)
                    {
                        var getDistinctDate = data.DistinctBy(x => x.DATE).ToList();
                        var dd = getDistinctDate.ElementAt(1);
                        var Transactions = data.Where(x => x.DATE == dd.DATE).ToList();
                        getHalfDay = Transactions;
                    }

                    if (counter > 0 && counter < fileCount-1)
                    {
                        getHalfDay = data;
                    }

                    if (counter == fileCount-1)
                    {
                        var getDistinctDate = data.DistinctBy(x => x.DATE).ToList();
                        var dd = getDistinctDate.ElementAt(0);
                        var Transactions = data.Where(x => x.DATE == dd.DATE).ToList();
                        getHalfDay = Transactions;
                    }
                }

            }
            catch (Exception)
            {
            }
            return success;
        }








    }
}
