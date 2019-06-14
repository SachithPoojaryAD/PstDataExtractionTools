using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Collections.Generic;
using MySql.Data.MySqlClient;
using OfficeOpenXml;

namespace PstDataExtractionTools
{
    [System.Runtime.InteropServices.Guid("C9AF260D-F666-41E5-BAC9-8699CC7020BE")]
    class Program
    {
        /*folder path of Aktiv1, Aktiv2, final copy destination and excel file*/
        //string Aktiv1FolderPath = @"D:\Sachith\PstTest\pstofbothAktiv1andAktiv2";
        //string Aktiv2FolderPath = @"D:\Sachith\PstTest\pstofbothAktiv1andAktiv2";
        //string FinalCopyFolderPath = @"D:\Sachith\PstTest\test";
        //string PSTFolderPath = @"D:\Sachith\PstTest\pstofbothAktiv1andAktiv2";
        //string PSTIgnoreFolderPath = @"D:\Sachith\PstTest\pstofbothAktiv1andAktiv2\_ignore";
        //string ExcelFilePath = @"D:\Sachith\PstTest\TestUsers.xlsx";

        public string Aktiv1FolderPath { get; set; }
        public string Aktiv2FolderPath { get; set; }
        public string FinalCopyFolderPath { get; set; }
        public string ExcelFilePath { get; set; }
        public string LogFilePath { get; set; } = @"D:\Sachith\PstTest\pstofbothAktiv1andAktiv2";
        public StringBuilder InitialLog { get; set; }

        string LogFileName = string.Format("Log{0}.txt", DateTime.Now.ToFileTime());
        int JobCount = 0;

        static void Main(string[] args)
        {
            Program prog = new Program();
            //prog.WriteToExcel(@"D:\Sachith\TestUsers.xlsx");
            //prog.CombineUserList();
            //prog.GetSimilarNames();
            //prog.MarkUserMapping();
            prog.InitialLog = new StringBuilder();

            DateTime startTime = DateTime.Now;
            Console.WriteLine("-----------------------------------Start----------------------------------------");
            prog.InitialLog.AppendLine("\n-----------------------------------Start of Log----------------------------------------");

            Console.WriteLine("Please select option \n" +
                "1) Read excel file and nove & rename folders \n" +
                "2) Remove .pst from folder name \n" +
                "3) Remove unwanted folders from destination path \n" +
                "4) Rename PST Files\n" +
                "5) Get mismatch count from log file\n" +
                "6) Search directory for backup folder\n" +
                "7) Get user list from aktiv folder\n" +
                "8) Update main excel sheet\n" +
                "9) Delete file type from folder\n" +
                "10) Update excel aktiv status\n" +
                "11) Get DB User Count\n" +
                "12) Generate CSV Customer Mapping\n" +
                "13) Generate User Status Excel\n" +
                "14) Generate Excel Sheet for AKtiv users\n" +
                "15) Get CSV from Postfach\n" +
                "16) Get Folder Name For User\n" +
                "17) Exit\n");
            int selection = 0;
            int.TryParse(Console.ReadLine(), out selection);

            Console.WriteLine("\nStarted the process at: " + startTime);
            prog.InitialLog.AppendLine("Started the process at: " + startTime);

            switch (selection)
            {
                case 1:
                    Console.WriteLine("Read Excel File and Move & Rename folders");
                    prog.InitialLog.AppendLine("\nRead Excel File and Move & Rename folders");
                    prog.ReadExcelFile();
                    break;
                case 2:
                    Console.WriteLine("Remove .pst from folder name");
                    prog.InitialLog.AppendLine("\nRemove .pst from folder name");
                    prog.RemovePSTFromFolderName();
                    break;
                case 3:
                    Console.WriteLine("Remove unwanted folders from destination path");
                    prog.InitialLog.AppendLine("\nRemove unwanted folders from destination path");
                    prog.RemoveUnwantedFolders();
                    break;
                case 4:
                    Console.WriteLine("Rename PST Files");
                    prog.InitialLog.AppendLine("\nRename PST Files");
                    prog.RenameInternalPSTFiles();
                    break;
                case 5:
                    Console.WriteLine("Get Mismatch count from log file");
                    //prog.InitialLog.AppendLine("\nGet Mismatch count from log file");
                    prog.GetMismatchCount();
                    break;
                case 6:
                    Console.WriteLine("Search Directory for backup folder");
                    prog.InitialLog.AppendLine("\nSearch Directory for backup folder");
                    prog.DirectorySearch();
                    break;
                case 7:
                    Console.WriteLine("Get user list from aktiv folder");
                    prog.InitialLog.AppendLine("\nGet user list from aktiv folder");
                    prog.GetUserListFromAKtivFolder();
                    break;
                case 8:
                    Console.WriteLine("Update main excel sheet");
                    prog.InitialLog.AppendLine("\nUpdate main excel sheet");
                    prog.UpdateExcelForAllAktivUsers();
                    break;
                case 9:
                    Console.WriteLine("Delete file type from folder");
                    prog.InitialLog.AppendLine("\nDelete file type from folder");
                    prog.DeleteFileType();
                    break;
                case 10:
                    Console.WriteLine("Update excel aktiv status");
                    prog.InitialLog.AppendLine("\nUpdate excel aktiv status");
                    prog.UpdateExcelAktivStatus();
                    break;
                case 11:
                    Console.WriteLine("Get DB User Count");
                    prog.InitialLog.AppendLine("\nGet DB User Count");
                    prog.UpdateExcelMatchedColoumn();
                    break;
                case 12:
                    Console.WriteLine("Generate CSV Customer Mapping");
                    prog.GenerateCSVCustomerMapping();
                    break;
                case 13:
                    Console.WriteLine("Generate User Status Excel");
                    prog.GenerateUserListStatusExcel();
                    break;
                case 14:
                    Console.WriteLine("Generate Excel Sheet for AKtiv users");
                    prog.CreateAktivExcelSheet();
                    break;
                case 15:
                    Console.WriteLine("Get CSV from Postfach");
                    prog.GetCSVFromPostfach();
                    break;
                case 16:
                    Console.WriteLine("Get Folder Name For User");
                    prog.GetFolderNameForUser();
                    break;
                case 17:
                    Environment.Exit(0);
                    break;
                default:
                    Console.WriteLine("Please select valid option");
                    prog.InitialLog.AppendLine("\nPlease select valid option");
                    break;
            }

            DateTime endTime = DateTime.Now;
            Console.WriteLine(string.Format("\nCompleted the process at: {0} Total time consumed: {1}", endTime, endTime - startTime));
            prog.AddLogs(prog.LogFilePath + "\\", string.Format("\nCompleted the process at: {0} Total time consumed: {1}", endTime, endTime - startTime));

            Console.WriteLine("Done. Completed process for " + prog.JobCount + " users.");
            prog.AddLogs(prog.LogFilePath + "\\", string.Format("Done. Completed process for {0} users.", prog.JobCount));
            Console.WriteLine("Check logs at " + prog.LogFilePath + "\\" + prog.LogFileName);

            Console.WriteLine("---------------------------------------End--------------------------------------");
            prog.AddLogs(prog.LogFilePath + "\\", "---------------------------------------End of Log--------------------------------------");
            Console.ReadKey();
        }

        /// <summary>
        /// Takes the input of excel file, path of aktiv1, aktiv2 and destination folder.
        /// Iterates through the excel file and calls GetAndMoveFoldersAndFiles method for each user
        /// </summary>
        private void ReadExcelFile()
        {
            try
            {
                Console.WriteLine("\nEnter path of excel file");
                ExcelFilePath = Console.ReadLine();
                InitialLog.AppendLine("\nExcel file path: " + ExcelFilePath);

                Console.WriteLine("\nEnter path of Aktiv1 folder");
                Aktiv1FolderPath = Console.ReadLine();
                InitialLog.AppendLine("Aktiv1 folder path: " + Aktiv1FolderPath);
                if (string.IsNullOrEmpty(Aktiv1FolderPath))
                {
                    throw new InvalidFilePathException("Please enter valid Aktiv 1 folder path\n");
                }

                Console.WriteLine("\nEnter path of Aktiv2 folder");
                Aktiv2FolderPath = Console.ReadLine();
                InitialLog.AppendLine("Aktiv2 folder path: " + Aktiv2FolderPath);
                if (string.IsNullOrEmpty(Aktiv2FolderPath))
                {
                    throw new InvalidFilePathException("Please enter valid Aktiv 2 folder path\n");
                }

                Console.WriteLine("\nEnter destination path");
                FinalCopyFolderPath = Console.ReadLine();
                InitialLog.AppendLine("Destination folder path: " + FinalCopyFolderPath);
                if (string.IsNullOrEmpty(FinalCopyFolderPath))
                {
                    throw new InvalidFilePathException("Please enter valid destination folder path\n");
                }
                LogFilePath = FinalCopyFolderPath;
                AddLogs(LogFilePath + "\\", InitialLog.ToString());

                #region EPPlus method

                var excelPackage = new ExcelPackage(new FileInfo(ExcelFilePath));

                Console.WriteLine("\nEnter excel sheet name");
                var sheetName = Console.ReadLine();

                var xlWorkSheet = excelPackage.Workbook.Worksheets[sheetName];

                //iterate the rows
                for (int index = 1; index <= xlWorkSheet.Dimension.End.Row; index++)
                {
                    var username = xlWorkSheet.Cells[index, 2].Value.ToString().Trim();
                    if (!string.IsNullOrWhiteSpace(username) && !username.Equals("User name"))
                    {
                        //Rename and move files and folder
                        GetAndMoveFoldersAndFiles(Aktiv1FolderPath.Replace(" ", string.Empty), username.Replace(" ", string.Empty));
                        GetAndMoveFoldersAndFiles(Aktiv2FolderPath.Replace(" ", string.Empty), username.Replace(" ", string.Empty));

                        //Search and Move already renamed folders
                        //MoveFolder("Aktiv1", Aktiv1FolderPath.Replace(" ", string.Empty), username.Replace(",", string.Empty));
                        //MoveFolder("Aktiv2", Aktiv2FolderPath.Replace(" ", string.Empty), username.Replace(",", string.Empty));

                        //Check if renamed folders exists for a user
                        //CheckFolders("Aktiv1", Aktiv1FolderPath.Replace(" ", string.Empty), username.Replace(",", string.Empty));
                        //CheckFolders("Aktiv2", Aktiv2FolderPath.Replace(" ", string.Empty), username.Replace(",", string.Empty));
                    }
                }

                excelPackage.Dispose();
                #endregion

                #region ExcelDataReader method
                //using (var excelStream = File.Open(ExcelFilePath, FileMode.Open, FileAccess.Read))
                //{
                //    using (var excelReader = ExcelReaderFactory.CreateReader(excelStream))
                //    {
                //        do
                //        {
                //            //read each row of the excel file
                //            while (excelReader.Read())
                //            {
                //                //check if the user name is not null and extracted status is not 'done' or 'no data found'
                //                //if (excelReader.GetString(2) == null || excelReader.GetString(3) == null) return;
                //                //if ( excelReader.GetString(3).Equals("Done") || excelReader.GetString(3).Equals("No Data Found"))
                //                //    return;
                //                //check if the user name is not null and extracted status is not 'done' or 'no data found'
                //                //if (!string.IsNullOrEmpty(excelReader.GetString(3)) && excelReader.GetString(3).ToLower().Equals("cross-checking"))
                //                //if (excelReader.GetString(3) == null) return;
                //                //if (!string.IsNullOrEmpty(excelReader.GetString(3)) && excelReader.GetString(3).Equals("Cross-checking"))
                //                //{

                //                //}

                //                //check if username is not null
                //                //if (!string.IsNullOrEmpty(excelReader.GetString(3)) && excelReader.GetString(3).ToLower().Equals("cross-checking"))
                //                //if (!string.IsNullOrEmpty(excelReader.GetString(2)) && !(excelReader.GetString(3).ToLower().Equals("Done") || excelReader.GetString(3).ToLower().Equals("No Data Found")))
                //                if (!string.IsNullOrWhiteSpace(excelReader.GetString(2)))
                //                {
                //                    //Rename and move files and folder
                //                    GetAndMoveFoldersAndFiles(Aktiv1FolderPath.Replace(" ", string.Empty), excelReader.GetString(2).Replace(" ", string.Empty));
                //                    GetAndMoveFoldersAndFiles(Aktiv2FolderPath.Replace(" ", string.Empty), excelReader.GetString(2).Replace(" ", string.Empty));

                //                    //Search and Move already renamed folders
                //                    //MoveFolder("Aktiv1", Aktiv1FolderPath.Replace(" ", string.Empty), excelReader.GetString(2).Replace(",", string.Empty));
                //                    //MoveFolder("Aktiv2", Aktiv2FolderPath.Replace(" ", string.Empty), excelReader.GetString(2).Replace(",", string.Empty));

                //                    //Check if renamed folders exists for a user
                //                    //CheckFolders("Aktiv1", Aktiv1FolderPath.Replace(" ", string.Empty), excelReader.GetString(2).Replace(",", string.Empty));
                //                    //CheckFolders("Aktiv2", Aktiv2FolderPath.Replace(" ", string.Empty), excelReader.GetString(2).Replace(",", string.Empty));
                //                }
                //            }
                //        } while (excelReader.NextResult());
                //    }
                //}
                #endregion

                #region Interop.Excel method

                //Microsoft.Office.Interop.Excel.Application xlApp;
                //Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                //Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                //Microsoft.Office.Interop.Excel.Sheets xlBigSheet;
                //Microsoft.Office.Interop.Excel.Range xlSheetRange;

                //xlApp = new Microsoft.Office.Interop.Excel.Application();
                ////sets whether the excel file will be open during this process
                //xlApp.Visible = false;
                ////open the excel file
                //xlWorkBook = xlApp.Workbooks.Open(ExcelFilePath, 0,
                //            false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows,
                //             "", true, false, 0, true, false, false);

                ////get all the worksheets in the excel  file
                //xlBigSheet = xlWorkBook.Worksheets;

                //Console.WriteLine("\nEnter excel sheet name");
                //var xlSheetName = Console.ReadLine();

                ////string x = "Extracted";
                ////get the specified worksheet
                //xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlBigSheet.get_Item(xlSheetName);

                //xlSheetRange = xlWorkSheet.UsedRange;

                //int colCount = xlSheetRange.Columns.Count;
                //int rowCount = xlSheetRange.Rows.Count;
                ////iterate the rows
                //for (int index = 1; index <= rowCount; index++)
                //{
                //    Microsoft.Office.Interop.Excel.Range cell = xlSheetRange.Cells[index, 2];
                //    if (cell.Value2 != null && !string.IsNullOrWhiteSpace(cell.Value2.ToString()) && !cell.Value2.ToString().Trim().Equals("User name"))
                //    {
                //        //Rename and move files and folder
                //        GetAndMoveFoldersAndFiles(Aktiv1FolderPath.Replace(" ", string.Empty), cell.Value2.ToString().Replace(" ", string.Empty));
                //        GetAndMoveFoldersAndFiles(Aktiv2FolderPath.Replace(" ", string.Empty), cell.Value2.ToString().Replace(" ", string.Empty));

                //        //Search and Move already renamed folders
                //        //MoveFolder("Aktiv1", Aktiv1FolderPath.Replace(" ", string.Empty), cell.Value2.ToString().Replace(",", string.Empty));
                //        //MoveFolder("Aktiv2", Aktiv2FolderPath.Replace(" ", string.Empty), cell.Value2.ToString().Replace(",", string.Empty));

                //        //Check if renamed folders exists for a user
                //        //CheckFolders("Aktiv1", Aktiv1FolderPath.Replace(" ", string.Empty), cell.Value2.ToString().Replace(",", string.Empty));
                //        //CheckFolders("Aktiv2", Aktiv2FolderPath.Replace(" ", string.Empty), cell.Value2.ToString().Replace(",", string.Empty));
                //    }
                //}

                ////xlWorkBook.Save();

                ////this line causes the excel file to get corrupted
                ////xlWorkBook.SaveAs(excelFilePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal,
                ////        Missing.Value, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive,
                ////        Missing.Value, Missing.Value, Missing.Value,
                ////        Missing.Value, Missing.Value);

                ////cleanup
                //xlWorkBook.Close(Missing.Value, Missing.Value, Missing.Value);
                //xlWorkBook = null;
                //xlApp.Quit();
                //GC.WaitForPendingFinalizers();
                //GC.Collect();
                //GC.WaitForPendingFinalizers();
                //GC.Collect();
                #endregion
            }
            catch (IOException)
            {
                Console.WriteLine("\nThe Excel file cannot be accessed if it is open. Please close the excel file and try again");
                AddLogs(LogFilePath + "\\", "The Excel file cannot be accessed if it is open. Please close the excel file and try again");
            }
            catch (ArgumentException ex)
            {
                Console.WriteLine("\nThe path of the excel file is not valid");
                AddLogs(LogFilePath + "\\", "The path of the excel file is not valid");
            }
            catch (InvalidFilePathException ex)
            {
                Console.WriteLine(ex.Message);
                AddLogs(LogFilePath + "\\", ex.Message);
            }
            catch (Exception ex)
            {
                Console.WriteLine("\nError please check logs at " + LogFilePath);
                AddLogs(LogFilePath + "\\", ex.Message + " stacktrace:- " + ex.StackTrace);
            }
        }

        /// <summary>
        /// Searches the given folder to find the user name, renames the folder, pst file and moves it to the destination folder
        /// </summary>
        /// <param name="folderPath">Path of the folder to search</param>
        /// <param name="searchString">User name</param>
        private void GetAndMoveFoldersAndFiles(string folderPath, string searchString)
        {
            //flag to check if backup of user has been found
            bool isBackupFoundFlag = false;
            foreach (var folder in Directory.GetDirectories(folderPath))
            {
                var directoryName = folder;
                var userName = searchString;
                var folderPathToSearch = folderPath;
                int folderCounter = 0;

                //remove external from the username
                if ("extern".Contains(userName.Substring(userName.LastIndexOf("-") + 1).ToLower().Trim()))
                {
                    userName = userName.Substring(0, userName.LastIndexOf("-"));
                }

                //get the frolder name from the folder directory path
                var strFolderName = folder.Substring(folder.LastIndexOf("\\") + 1).Replace(" ", string.Empty);

                //if dash(-) exists in folder name, then remove it and its proceding characters
                if (strFolderName.Contains("-") && int.TryParse(strFolderName.Substring(strFolderName.LastIndexOf("-") + 1), out int n))
                {
                    strFolderName = strFolderName.Substring(0, strFolderName.LastIndexOf("-"));
                }
                //if folder name contains 'extern', then remove it and its preceding string
                if (strFolderName.Contains("-") && "extern".Contains(strFolderName.Substring(strFolderName.LastIndexOf("-") + 1).ToLower().Trim()))
                {
                    strFolderName = strFolderName.Substring(0, strFolderName.LastIndexOf("-"));
                }

                //check if a backup folder with the users name exists
                if (strFolderName.Contains(userName) || userName.Contains(strFolderName))
                {
                    try
                    {
                        string folderName;
                        //if username is longer than folder name, then use username
                        if (strFolderName.Length < userName.Length)
                        {
                            folderName = userName;
                        }
                        else
                        {
                            folderName = strFolderName;
                        }

                        folderName = folderName.Replace(",", " ");

                        //append _Aktiv1 or _Aktiv2 to the folder name
                        if (folderPathToSearch.Substring(folderPathToSearch.LastIndexOf("\\") + 1).ToLower().Contains("Aktiv1".ToLower()))
                        {
                            folderName += "_Aktiv1" + (folderCounter == 0 ? string.Empty : folderCounter.ToString());
                            folderCounter++;
                        }
                        else if (folderPathToSearch.Substring(folderPathToSearch.LastIndexOf("\\") + 1).ToLower().Contains("Aktiv2".ToLower()))
                        {
                            folderName += "_Aktiv2" + (folderCounter == 0 ? string.Empty : folderCounter.ToString());
                            folderCounter++;
                        }

                        //move the folder to its destination path
                        var folderPathToMove = FinalCopyFolderPath + "\\" + folderName;
                        try
                        {
                            Directory.Move(folder, folderPathToMove);

                            //browse the destination folder to rename the .pst files
                            DirectoryInfo folderDirectory = new DirectoryInfo(folderPathToMove);

                            //used counter for file name in case of multiple pst files
                            int fileCounter = 0;
                            foreach (var file in folderDirectory.GetFiles().OrderBy(f => f.Name))
                            {
                                if (file.Name.Contains(".pst"))
                                {
                                    Directory.Move(file.FullName, folderPathToMove + "\\" + folderName + (fileCounter == 0 ? string.Empty : fileCounter.ToString()) + ".pst");
                                    fileCounter++;
                                }
                            }
                            isBackupFoundFlag = true;
                            Console.WriteLine("Copied and renamed file: " + folderPathToMove);
                            AddLogs(LogFilePath + "\\", "Copied and renamed file: " + folderPathToMove + " Username: " + searchString);
                            JobCount++;
                            return;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error for user: " + searchString + " Please check logs at " + LogFilePath);
                            AddLogs(LogFilePath + "\\", "Username:- " + searchString + " " + ex.Message + " stacktrace:- " + ex.StackTrace);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error for user: " + searchString + " Please check logs at " + LogFilePath);
                        AddLogs(LogFilePath + "\\", "Username:- " + searchString + " " + ex.Message + " stacktrace:- " + ex.StackTrace);
                    }
                }
                else //check if the name contains special characters(stored as ?)
                {
                    try
                    {
                        //dictionary to store the special character. key: index of the special character, value: special character
                        //Dictionary<int, char> listQuestionMarkOccurance = new Dictionary<int, char>();
                        var charArray = userName.ToCharArray();
                        directoryName = strFolderName;

                        for (int i = 0; i < charArray.Length; i++)
                        {
                            //if username contains '?', then replace it with the character from folder name at the same index
                            if (charArray[i].Equals('?') && directoryName.ElementAtOrDefault(i) != 0)
                            {
                                charArray[i] = directoryName[i];
                            }
                        }
                        userName = new string(charArray);

                        //check if username and directory name without the special characters match
                        if (directoryName.Contains(userName) || userName.Contains(directoryName))
                        {
                            string folderchar;
                            //if username is longer than folder name, then use username
                            if (directoryName.Length < userName.Length)
                            {
                                folderchar = userName;
                            }
                            else
                            {
                                folderchar = directoryName;
                            }

                            folderchar = folderchar.Replace(",", " ");

                            //append _Aktiv1 or _Aktiv2 to the folder name
                            if (folderPathToSearch.Contains("Aktiv1"))
                            {
                                folderchar += "_Aktiv1" + (folderCounter == 0 ? string.Empty : folderCounter.ToString());
                                folderCounter++;
                            }
                            else if (folderPathToSearch.Contains("Aktiv2"))
                            {
                                folderchar += "_Aktiv2" + (folderCounter == 0 ? string.Empty : folderCounter.ToString());
                                folderCounter++;
                            }

                            //move the folder to its destination path
                            var folderPathToMove = FinalCopyFolderPath + "\\" + folderchar;
                            try
                            {
                                Directory.Move(folder, folderPathToMove);

                                //browse the destination folder to rename the .pst files
                                DirectoryInfo folderDirectory = new DirectoryInfo(folderPathToMove);

                                //used counter for file name in case of multiple pst files
                                int fileCounter = 0;
                                foreach (var file in folderDirectory.GetFiles().OrderBy(f => f.Name))
                                {
                                    if (file.Name.Contains(".pst"))
                                    {
                                        Directory.Move(file.FullName, folderPathToMove + "\\" + folderchar + (fileCounter == 0 ? string.Empty : fileCounter.ToString()) + ".pst");
                                        fileCounter++;
                                    }
                                }
                                isBackupFoundFlag = true;
                                Console.WriteLine("Copied and renamed file: " + folderPathToMove);
                                AddLogs(LogFilePath + "\\", "Copied and renamed file: " + folderPathToMove + " Username: " + searchString);
                                JobCount++;
                                return;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("Error for user: " + searchString + " Please check logs at " + LogFilePath);
                                AddLogs(LogFilePath + "\\", "Username:- " + searchString + " " + ex.Message + " stacktrace:- " + ex.StackTrace);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error for user: " + searchString + " Please check logs at " + LogFilePath);
                        AddLogs(LogFilePath + "\\", "Username:- " + searchString + " " + ex.Message + " stacktrace:- " + ex.StackTrace);
                    }
                }
            }
            //if no backup is found, print it in log
            if (!isBackupFoundFlag)
            {
                Console.WriteLine(string.Format("Error for user: {0}. Could not find data. Please check logs at {1}", searchString, LogFilePath));
                AddLogs(LogFilePath + "\\", "Username:- " + searchString + ". Could not find data.");
            }
        }

        /// <summary>
        /// Search and move already renamed folders
        /// </summary>
        /// <param name="Aktiv">Aktiv number</param>
        /// <param name="folderPath">Path of the folder where you want to search</param>
        /// <param name="searchString">Name of the user that you want to find</param>
        private void MoveFolder(string Aktiv, string folderPath, string searchString)
        {
            //flag to check if backup of user has been found
            bool isBackupFoundFlag = false;
            foreach (var folder in Directory.GetDirectories(folderPath))
            {
                var directoryName = folder;
                var userName = searchString;
                var folderPathToSearch = folderPath;

                //remove external from the username
                if ("extern".Contains(userName.Substring(userName.LastIndexOf("-") + 1).ToLower().Trim()))
                {
                    userName = userName.Substring(0, userName.LastIndexOf("-"));
                }

                //get the frolder name from the folder directory path
                var strFolderName = folder.Substring(folder.LastIndexOf("\\") + 1);

                //if dash(-) exists in folder name, then remove it and its proceding characters
                if (int.TryParse(strFolderName.Substring(strFolderName.LastIndexOf("-") + 1), out int n))
                {
                    strFolderName = strFolderName.Substring(0, strFolderName.LastIndexOf("-"));
                }
                //if folder name contains 'extern', then remove it and its preceding string
                if ("extern".Contains(strFolderName.Substring(strFolderName.LastIndexOf("-") + 1).ToLower()))
                {
                    strFolderName = strFolderName.Substring(0, strFolderName.LastIndexOf("-"));
                }
                strFolderName = strFolderName.Trim();

                //check if a backup folder with the users name exists
                if ((strFolderName.Contains(userName) || userName.Contains(strFolderName)) && strFolderName.ToLower().Contains(Aktiv.ToLower()))
                {
                    try
                    {
                        var destinationFolderName = FinalCopyFolderPath + "\\" + folder.Substring(folder.LastIndexOf("\\") + 1);
                        Directory.Move(folder, destinationFolderName);

                        isBackupFoundFlag = true;
                        Console.WriteLine("Copied and renamed file: " + destinationFolderName);
                        AddLogs(LogFilePath + "\\", "Copied and renamed file: " + destinationFolderName + " Username: " + searchString);
                        JobCount++;
                        return;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error for user: " + searchString + " Please check logs at " + LogFilePath);
                        AddLogs(LogFilePath + "\\", "Username:- " + searchString + " " + ex.Message + " stacktrace:- " + ex.StackTrace);
                    }
                }
                else
                {
                    var charArray = userName.ToCharArray();
                    directoryName = strFolderName;

                    for (int i = 0; i < charArray.Length; i++)
                    {
                        //if username contains '?', then replace it with the character from folder name at the same index
                        if (charArray[i].Equals('?') && directoryName.ElementAtOrDefault(i) != 0)
                        {
                            charArray[i] = directoryName[i];
                        }
                    }
                    userName = new string(charArray);
                    userName = userName.Trim();

                    if ((strFolderName.Contains(userName) || userName.Contains(strFolderName)) && strFolderName.ToLower().Contains(Aktiv.ToLower()))
                    {
                        try
                        {
                            var destinationFolderName = FinalCopyFolderPath + "\\" + folder.Substring(folder.LastIndexOf("\\") + 1);
                            Directory.Move(folder, destinationFolderName);

                            isBackupFoundFlag = true;
                            Console.WriteLine("Copied and renamed file: " + destinationFolderName);
                            AddLogs(LogFilePath + "\\", "Copied and renamed file: " + destinationFolderName + " Username: " + searchString);
                            JobCount++;
                            return;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error for user: " + searchString + " Please check logs at " + LogFilePath);
                            AddLogs(LogFilePath + "\\", "Username:- " + searchString + " " + ex.Message + " stacktrace:- " + ex.StackTrace);
                        }
                    }
                }
            }
            //if no backup is found, print it in log
            if (!isBackupFoundFlag)
            {
                Console.WriteLine(string.Format("Error for user: {0}. Could not find data. Please check logs at {1}", searchString, LogFilePath));
                AddLogs(LogFilePath + "\\", "Username:- " + searchString + ". Could not find data.");
            }
        }

        void AddLogs(string path, string errorText)
        {
            StreamWriter sw = new StreamWriter(path + LogFileName, true, Encoding.UTF8);
            sw.WriteLine(errorText);
            sw.Close();
        }

        /// <summary>
        /// Removes '.pst' from the name of folders which has been extracted
        /// </summary>
        private void RemovePSTFromFolderName()
        {
            string FolderPath;
            try
            {
                Console.WriteLine("\nEnter path of folder");
                FolderPath = Console.ReadLine();
                InitialLog.AppendLine("\nFolder path: " + FolderPath);
                if (string.IsNullOrEmpty(FolderPath))
                {
                    throw new InvalidFilePathException("Please enter valid folder path\n");
                }
            }
            catch (InvalidFilePathException ex)
            {
                Console.WriteLine(ex.Message);
                AddLogs(LogFilePath + "\\", ex.Message);
                return;
            }

            LogFilePath = FolderPath;
            AddLogs(LogFilePath + "\\", InitialLog.ToString());

            foreach (var folder in Directory.GetDirectories(FolderPath))
            {
                //make sure that the current files being extracted doesn't get renamed
                //if (!(folder.Contains("Eilers Andre_Aktiv1.pst") || folder.Contains("Eilers Andre_Aktiv2.pst") || folder.Contains("Eisenschmidt Marco_Aktiv1.pst") || folder.Contains("Eisenschmidt Marco_Aktiv2.pst")))
                {
                    //check if folder name ends with '.pst'
                    if (folder.EndsWith(".pst"))
                    {
                        //replace '.pst' with empty string
                        var newFolderName = folder.Replace(".pst", string.Empty);
                        if (newFolderName.Contains("-Aktiv"))
                        {
                            newFolderName = newFolderName.Replace("-Aktiv", "_Aktiv");
                        }
                        try
                        {
                            //move the newly renamed folder
                            Directory.Move(folder, newFolderName);

                            Console.WriteLine("Renamed file: " + newFolderName);
                            AddLogs(LogFilePath + "\\", "Renamed file: " + newFolderName);
                        }
                        catch (IOException ex)
                        {
                            Console.WriteLine(string.Format("Error: Access to {0} folder was denied", folder));
                            AddLogs(LogFilePath + "\\", "Error: Access to " + folder + " folder was denied. " + ex.Message + " Stacktrace: " + ex.StackTrace);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error for folder: " + folder + " Please check logs at " + LogFilePath);
                            AddLogs(LogFilePath + "\\", "Folder:- " + folder + " " + ex.Message + " stacktrace:- " + ex.StackTrace);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Moves the folders which do not contain a .pst file to _ignore folder
        /// </summary>
        private void RemoveUnwantedFolders()
        {
            string FolderPath;
            string IgnoreFolderPath;
            try
            {
                Console.WriteLine("\nEnter path of folder");
                FolderPath = Console.ReadLine();
                InitialLog.AppendLine("\nFolder path: " + FolderPath);
                if (string.IsNullOrEmpty(FolderPath))
                {
                    throw new InvalidFilePathException("Please enter valid folder path\n");
                }

                Console.WriteLine("\nEnter path of ignore folder");
                IgnoreFolderPath = Console.ReadLine();
                if (string.IsNullOrEmpty(IgnoreFolderPath))
                {
                    throw new InvalidFilePathException("Please enter valid folder path\n");
                }
            }
            catch (InvalidFilePathException ex)
            {
                Console.WriteLine(ex.Message);
                AddLogs(LogFilePath + "\\", ex.Message);
                return;
            }

            LogFilePath = FolderPath;
            AddLogs(LogFilePath + "\\", InitialLog.ToString());

            foreach (var folder in Directory.GetDirectories(FolderPath))
            {
                //proceed only if folder name is not '_ignore'
                if (!folder.Contains("_ignore"))
                {
                    DirectoryInfo folderDirectory = new DirectoryInfo(folder);
                    int pstCounter = 0;
                    //check if folder contains pst files
                    foreach (var file in folderDirectory.GetFiles().OrderBy(f => f.Name))
                    {
                        if (file.Name.EndsWith(".pst"))
                        {
                            //if pst file is present, increase the pst counter
                            pstCounter++;
                        }
                    }
                    //move foler to _ignore folder if it does not contain any .pst files
                    if (pstCounter == 0)
                    {
                        var destinationFolderName = IgnoreFolderPath + "\\" + folder.Substring(folder.LastIndexOf("\\") + 1);
                        try
                        {
                            Directory.Move(folder, destinationFolderName);
                            Console.WriteLine("Copied folder: " + folder);
                            AddLogs(LogFilePath + "\\", "Copied folder: " + folder);
                        }
                        catch (IOException ex)
                        {
                            Console.WriteLine(string.Format("Error: {0} folder is in use by other process", folder));
                            AddLogs(LogFilePath + "\\", "Error: " + folder + " folder is in use by other process. " + ex.Message + " Stacktrace: " + ex.StackTrace);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error for folder: " + folder + " Please check logs at " + LogFilePath);
                            AddLogs(LogFilePath + "\\", "Folder:- " + folder + " " + ex.Message + " stacktrace:- " + ex.StackTrace);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Renames the .pst files inside of a given folder to the correct format
        /// </summary>
        private void RenameInternalPSTFiles()
        {
            string currentActivFolder;
            string FolderPath;
            try
            {
                Console.WriteLine("\nEnter path of folder");
                FolderPath = Console.ReadLine();
                InitialLog.AppendLine("\nFolder path: " + FolderPath);
                if (string.IsNullOrEmpty(FolderPath))
                {
                    throw new InvalidFilePathException("Please enter valid folder path\n");
                }

                Console.WriteLine("\nEnter current Aktiv Path(eg. Aktiv1, Aktiv2...)");
                currentActivFolder = Console.ReadLine();
                InitialLog.AppendLine("\nAktiv path: " + currentActivFolder);
                if (string.IsNullOrEmpty(FolderPath))
                {
                    throw new InvalidFilePathException("Please enter valid folder path\n");
                }
            }
            catch (InvalidFilePathException ex)
            {
                Console.WriteLine(ex.Message);
                AddLogs(LogFilePath + "\\", ex.Message);
                return;
            }

            LogFilePath = FolderPath;
            AddLogs(LogFilePath + "\\", InitialLog.ToString());

            foreach (var folder in Directory.GetDirectories(FolderPath))
            {
                //proceed only if folder name is not '_ignore'
                if (!folder.Contains("_ignore"))
                {

                    //get the frolder name from the folder directory path
                    var strFolderName = folder.Substring(folder.LastIndexOf("\\") + 1).Replace(",", " ");

                    /*uncomment if you want to remove count from file name*/
                    //if dash(-) exists in folder name, then remove it and its proceding characters
                    if (strFolderName.Contains("-") && int.TryParse(strFolderName.Substring(strFolderName.LastIndexOf("-") + 1), out int n))
                    {
                        strFolderName = strFolderName.Substring(0, strFolderName.LastIndexOf("-"));
                    }

                    /*uncomment if you want to remove 'extern' from file name*/
                    //if folder name contains 'extern', then remove it and its preceding string
                    if (strFolderName.Contains("-") && "extern".Contains(strFolderName.Substring(strFolderName.LastIndexOf("-") + 1).ToLower().Trim()))
                    {
                        strFolderName = strFolderName.Substring(0, strFolderName.LastIndexOf("-"));
                    }

                    strFolderName += "_" + currentActivFolder;

                    //used counter for file name in case of multiple pst files
                    int fileCounter = 0;
                    DirectoryInfo folderDirectory = new DirectoryInfo(folder);
                    //check if folder contains pst files
                    foreach (var file in folderDirectory.GetFiles().OrderBy(f => f.Name))
                    {
                        if (file.Name.EndsWith(".pst"))
                        {
                            /*uncomment only if you want to rename numeric pst files(eg. 001.pst)*/
                            //check if pst file name is numeric
                            //bool isFileNameInt = int.TryParse(file.Name.Substring(0, file.Name.LastIndexOf(".")), out int s);
                            //if (isFileNameInt)
                            //{
                            try
                            {
                                //var newFileName = file.FullName.Substring(file.FullName.LastIndexOf("\\") + 1).Replace(".pst", string.Empty);
                                //newFileName += "_" + strFolderName;
                                //var newFileName = file.DirectoryName + "\\" + file.DirectoryName.Substring(file.DirectoryName.LastIndexOf("\\") + 1) + ".pst";
                                //var newFilePath = file.DirectoryName + "\\" + newFileName + ".pst";
                                var newFilePath = file.DirectoryName + "\\" + strFolderName + (fileCounter == 0 ? string.Empty : fileCounter.ToString()) + ".pst";
                                JobCount++;
                                fileCounter++;
                                Directory.Move(file.FullName, newFilePath);
                                Console.WriteLine("Renamed File: " + file.FullName + " to " + newFilePath);
                                AddLogs(LogFilePath + "\\", "Renamed File: " + file.FullName + " to " + newFilePath);
                            }
                            catch (IOException ex)
                            {
                                Console.WriteLine(string.Format("Error: {0} folder is in use by other process", folder));
                                AddLogs(LogFilePath + "\\", "Error: " + file.FullName + " folder is in use by other process. " + ex.Message + " Stacktrace: " + ex.StackTrace);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("Error for file: " + file.FullName + " Please check logs at " + LogFilePath);
                                AddLogs(LogFilePath + "\\", "file:- " + file.FullName + " " + ex.Message + " stacktrace:- " + ex.StackTrace);
                            }
                            //}
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Check if renamed folders exists for a user
        /// </summary>
        /// <param name="Aktiv">Aktiv number</param>
        /// <param name="folderPath">Path of the folder where you want to search</param>
        /// <param name="searchString">Name of the user that you want to find</param>
        private void CheckFolders(string Aktiv, string folderPath, string searchString)
        {

            //flag to check if backup of user has been found
            bool isBackupFoundFlag = false;
            foreach (var folder in Directory.GetDirectories(folderPath))
            {
                var directoryName = folder;
                var userName = searchString;
                var folderPathToSearch = folderPath;

                //remove external from the username
                if ("extern".Contains(userName.Substring(userName.LastIndexOf("-") + 1).ToLower().Trim()))
                {
                    userName = userName.Substring(0, userName.LastIndexOf("-"));
                }

                //get the frolder name from the folder directory path
                var strFolderName = folder.Substring(folder.LastIndexOf("\\") + 1);

                //if dash(-) exists in folder name, then remove it and its proceding characters
                if (int.TryParse(strFolderName.Substring(strFolderName.LastIndexOf("-") + 1), out int n))
                {
                    strFolderName = strFolderName.Substring(0, strFolderName.LastIndexOf("-"));
                }
                //if folder name contains 'extern', then remove it and its preceding string
                if ("extern".Contains(strFolderName.Substring(strFolderName.LastIndexOf("-") + 1).ToLower().Trim()))
                {
                    strFolderName = strFolderName.Substring(0, strFolderName.LastIndexOf("-"));
                }
                strFolderName = strFolderName.Trim();

                //check if a backup folder with the users name exists
                if ((strFolderName.Contains(userName) || userName.Contains(strFolderName)) && strFolderName.ToLower().Contains(Aktiv.ToLower()))
                {
                    isBackupFoundFlag = true;
                    return;
                }
                else
                {
                    var charArray = userName.ToCharArray();
                    directoryName = strFolderName;

                    for (int i = 0; i < charArray.Length; i++)
                    {
                        //if username contains '?', then replace it with the character from folder name at the same index
                        if (charArray[i].Equals('?') && directoryName.ElementAtOrDefault(i) != 0)
                        {
                            charArray[i] = directoryName[i];
                        }
                    }
                    userName = new string(charArray);
                    userName = userName.Trim();

                    if ((strFolderName.Contains(userName) || userName.Contains(strFolderName)) && strFolderName.ToLower().Contains(Aktiv.ToLower()))
                    {
                        isBackupFoundFlag = true;
                        return;
                    }
                }
            }
            //if no backup is found, print it in log
            if (!isBackupFoundFlag)
            {
                Console.WriteLine(string.Format("Error for user: {0}. Could not find data. Please check logs at {1}", searchString + " " + Aktiv, LogFilePath));
                //AddLogs(LogFilePath + "\\", "Username:- " + searchString + ". Could not find data.");
                AddLogs(LogFilePath + "\\", searchString + " " + Aktiv);
            }
        }

        /// <summary>
        /// Gets the mismatched counts with folder name from the pst extraction logs
        /// </summary>
        private void GetMismatchCount()
        {
            ////path of log file
            //string filePath = @"C:\Users\s.poojary\Desktop\SavingLogPC20.txt";
            ////path of destination text file
            //string destinationPath = @"C:\Users\s.poojary\Desktop\FolderToSearchPC20.txt";


            //path of log file
            string filePath;
            //path of destination text file
            string destinationPath;

            try
            {
                Console.WriteLine("\nEnter path of extraction log file");
                filePath = Console.ReadLine();
                //InitialLog.AppendLine("\nLog file path: " + filePath);
                if (string.IsNullOrEmpty(filePath))
                {
                    throw new InvalidFilePathException("Please enter valid file path\n");
                }

                Console.WriteLine("\nEnter path of destination log file");
                destinationPath = Console.ReadLine();
                //InitialLog.AppendLine("\nDestination log file path: " + destinationPath);
                if (string.IsNullOrEmpty(destinationPath))
                {
                    throw new InvalidFilePathException("Please enter valid file path\n");
                }
            }
            catch (InvalidFilePathException ex)
            {
                Console.WriteLine(ex.Message);
                //AddLogs(LogFilePath + "\\", ex.Message);
                return;
            }

            if (File.Exists(filePath))
            {
                StreamWriter sw = new StreamWriter(destinationPath, true, Encoding.UTF8);

                var lines = File.ReadAllLines(filePath);
                for (var i = 0; i < lines.Length; i += 1)
                {
                    var line = lines[i];
                    //check if the current line contains "Items converted :"
                    if (line.Contains("Items converted :"))
                    {
                        //take the string after the colon(:) and split it via '/' to get the count
                        var strCount = line.Split(':');
                        var count = strCount[1].Trim().Split('/');
                        //check if count is mismatched and that it is greater than 0
                        if ((Convert.ToInt32(count[0]) != Convert.ToInt32(count[1])) && Convert.ToInt32(count[0]) > 0)
                        {
                            //format the string to be printed in the txt file
                            var strFolderName = lines[i - 1].Substring(lines[i - 1].LastIndexOf("\\") + 1);
                            var strFolderNameCount = string.Format("{0}:{1}", strFolderName, count[0]);
                            sw.WriteLine(strFolderNameCount);

                            Console.WriteLine(strFolderNameCount);
                        }
                    }
                }
                sw.Close();
            }
        }

        /// <summary>
        /// Takes the folder path and text file path input from user to search the given folder and calls SearchFolder method
        /// </summary>
        private void DirectorySearch()
        {
            //path of folder to search
            string searchDirectory;
            //path of log file
            string filePath;
            //path of destination text file
            //string destinationPath;

            try
            {
                Console.WriteLine("\nEnter path of folder to search");
                searchDirectory = Console.ReadLine();
                LogFilePath = searchDirectory;
                AddLogs(LogFilePath + "\\", "\nFolder to search: " + searchDirectory);
                if (string.IsNullOrEmpty(searchDirectory))
                {
                    throw new InvalidFilePathException("Please enter valid folder path\n");
                }

                Console.WriteLine("\nEnter path of text file");
                filePath = Console.ReadLine();
                AddLogs(LogFilePath + "\\", "\nText file path: " + filePath);
                if (string.IsNullOrEmpty(filePath))
                {
                    throw new InvalidFilePathException("Please enter valid file path\n");
                }

                //Console.WriteLine("\nEnter path of destination log file");
                //destinationPath = Console.ReadLine();
                //AddLogs(LogFilePath + "\\", "\nDestination log file path: " + destinationPath);
                //if (string.IsNullOrEmpty(destinationPath))
                //{
                //    throw new InvalidFilePathException("Please enter valid file path\n");
                //}
            }
            catch (InvalidFilePathException ex)
            {
                Console.WriteLine(ex.Message);
                AddLogs(LogFilePath + "\\", ex.Message);
                return;
            }

            SearchFolders(searchDirectory, filePath);
        }

        /// <summary>
        /// Searches a folder recursively to find a user name. Called by DirectorySearch method
        /// </summary>
        /// <param name="searchDir">Path of the folder to search</param>
        /// <param name="filePath">Path of the text file which contains list to user names</param>
        private void SearchFolders(string searchDir, string filePath)
        {
            try
            {
                foreach (var directory in Directory.GetDirectories(searchDir))
                {
                    //if the folder does not contain extracted data, recursively search that folder
                    if (!directory.Contains("Oberste Ebene der Outlook-Datendatei"))
                    {
                        //StreamWriter sw = new StreamWriter(destinationPath, true, Encoding.UTF8);
                        var lines = File.ReadAllLines(filePath);
                        for (var i = 0; i < lines.Length; i += 1)
                        {
                            var line = lines[i];
                            line = line.Replace(",", string.Empty);
                            //check if the current line contains the user name
                            if (directory.Trim().ToLower().Contains(line.Trim().ToLower()))
                            {
                                //sw.WriteLine(directory);
                                AddLogs(LogFilePath + "\\", "\nUser Found: " + line + " - " + directory);
                                Console.WriteLine(directory);
                            }
                        }

                        //search the current directory and check if username matches
                        SearchFolders(directory, filePath);
                    }
                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// Get the list of users in a aktiv folder
        /// </summary>
        private void GetUserListFromAKtivFolder()
        {
            string FolderPath;
            //string DestinationLogPath;
            try
            {
                Console.WriteLine("\nEnter path of folder");
                FolderPath = Console.ReadLine();
                InitialLog.AppendLine("\nFolder path: " + FolderPath);
                if (string.IsNullOrEmpty(FolderPath))
                {
                    throw new InvalidFilePathException("Please enter valid folder path\n");
                }

                //Console.WriteLine("\nEnter path of destination text file");
                //DestinationLogPath = Console.ReadLine();
                //InitialLog.AppendLine("\nDestination log path: " + DestinationLogPath);
                //if (string.IsNullOrEmpty(DestinationLogPath))
                //{
                //    throw new InvalidFilePathException("Please enter valid folder path\n");
                //}
            }
            catch (InvalidFilePathException ex)
            {
                Console.WriteLine(ex.Message);
                AddLogs(LogFilePath + "\\", ex.Message);
                return;
            }

            LogFilePath = FolderPath;
            AddLogs(LogFilePath + "\\", InitialLog.ToString());

            foreach (var folder in Directory.GetDirectories(FolderPath))
            {
                ////get the frolder name from the folder directory path
                ////var strFolderName = folder.Substring(folder.LastIndexOf("\\") + 1).Replace(",", " ");
                //var strFolderName = folder.Substring(folder.LastIndexOf("\\") + 1);

                ////if dash(-) exists in folder name, then remove it and its proceding characters
                //if (strFolderName.Contains("-") && int.TryParse(strFolderName.Substring(strFolderName.LastIndexOf("-") + 1), out int n))
                //{
                //    strFolderName = strFolderName.Substring(0, strFolderName.LastIndexOf("-"));
                //}
                ////if folder name contains 'extern', then remove it and its preceding string
                //if (strFolderName.Contains("-") && "extern".Contains(strFolderName.Substring(strFolderName.LastIndexOf("-") + 1).ToLower().Trim()))
                //{
                //    strFolderName = strFolderName.Substring(0, strFolderName.LastIndexOf("-"));
                //}

                //Console.WriteLine(strFolderName);

                DirectoryInfo folderDirectory = new DirectoryInfo(folder);
                //StreamWriter sw = new StreamWriter(DestinationLogPath, true, Encoding.UTF8);
                //check if folder contains log files
                foreach (var file in folderDirectory.GetFiles().OrderBy(f => f.Name))
                {
                    if (file.Name.EndsWith(".txt"))
                    {
                        //if txt file is present, read the name
                        var lines = File.ReadAllLines(file.FullName);
                        for (var i = 0; i < lines.Length; i += 1)
                        {
                            var line = lines[i];
                            //check if the current line contains "Archive Name:"
                            if (line.Contains("Archive Name:"))
                            {
                                //take the string after the colon(:) to get the user's name
                                var strLine = line.Split(':');
                                var userName = strLine[1].Trim();

                                //sw.WriteLine(userName);

                                Console.WriteLine(userName);

                                AddLogs(LogFilePath + "\\", userName);
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Combine list of users of aktiv1-5. Combines all the users and removes duplicate users
        /// </summary>
        private void CombineUserList()
        {
            string aktiv12FilePath = @"C:\Users\s.poojary\Desktop\Aktiv12Users.txt";
            string aktiv1UsersFilePath = @"C:\Users\s.poojary\Desktop\Aktiv1Users.txt";
            string aktiv2UsersFilePath = @"C:\Users\s.poojary\Desktop\Aktiv2Users.txt";
            string aktiv3UsersFilePath = @"C:\Users\s.poojary\Desktop\Aktiv3Users.txt";
            string aktiv4UsersFilePath = @"C:\Users\s.poojary\Desktop\Aktiv4Users.txt";
            string aktiv5UsersFilePath = @"C:\Users\s.poojary\Desktop\Aktiv5Users.txt";
            string sharedUsersFilePath = @"C:\Users\s.poojary\Desktop\Shared52Users.txt";
            string completeUsersFilePath = @"C:\Users\s.poojary\Desktop\CompleteUsers.txt";

            var lstAktiv12Users = File.ReadLines(aktiv12FilePath, Encoding.UTF8);
            //var lstAktiv1Users = File.ReadLines(aktiv1UsersFilePath, Encoding.UTF8);
            //var lstAktiv2Users = File.ReadLines(aktiv2UsersFilePath, Encoding.UTF8);
            var lstAktiv3Users = File.ReadLines(aktiv3UsersFilePath, Encoding.UTF8);
            var lstAktiv4Users = File.ReadLines(aktiv4UsersFilePath, Encoding.UTF8);
            var lstAktiv5Users = File.ReadLines(aktiv5UsersFilePath, Encoding.UTF8);
            var lstSharedUsers = File.ReadLines(sharedUsersFilePath, Encoding.UTF8);

            //var lstAktiv12Users = fullUsersArray.ToList();
            //var lstAktiv1Users = aktiv1UsersArray.ToList();
            //var lstAktiv2Users = aktiv2UsersArray.ToList();
            //var lstAktiv3Users = aktiv3UsersArray.ToList();
            //var lstAktiv4Users = aktiv4UsersArray.ToList();
            //var lstAktiv5Users = aktiv5UsersArray.ToList();
            //var lstSharedUsers = sharedUsersArray.ToList();

            //var lstCompleteUsers = lstAktiv1Users.Union(lstAktiv2Users).Union(lstSharedUsers).Union(lstAktiv3Users).Union(lstAktiv4Users).Union(lstAktiv5Users).ToList();
            var lstCompleteUsers = lstAktiv12Users.Union(lstSharedUsers).Union(lstAktiv3Users).Union(lstAktiv4Users).Union(lstAktiv5Users).ToList();
            lstCompleteUsers.Sort();

            File.WriteAllLines(completeUsersFilePath, lstCompleteUsers);
        }

        /// <summary>
        /// Update excel sheet to display which users are present in the aktiv folders
        /// </summary>
        private void UpdateExcelForAllAktivUsers()
        {
            try
            {
                var SourceExcelFilePath = @"C:\Users\s.poojary\Desktop\Extracted_Report_Local 30-4-19 12.10.xlsx";
                var DestinationExcelFilePath = @"C:\Users\s.poojary\Desktop\Aktiv1-5Users.xlsx";
                //InitialLog = new StringBuilder();
                AddLogs(LogFilePath + "\\", "\nSource Excel file path: " + SourceExcelFilePath);
                AddLogs(LogFilePath + "\\", "\nDestination Excel file path: " + DestinationExcelFilePath);

                string aktiv1UsersFilePath = @"C:\Users\s.poojary\Desktop\Aktiv1Users.txt";
                string aktiv2UsersFilePath = @"C:\Users\s.poojary\Desktop\Aktiv2Users.txt";
                string aktiv3UsersFilePath = @"C:\Users\s.poojary\Desktop\Aktiv3Users.txt";
                string aktiv4UsersFilePath = @"C:\Users\s.poojary\Desktop\Aktiv4Users.txt";
                string aktiv5UsersFilePath = @"C:\Users\s.poojary\Desktop\Aktiv5Users.txt";
                string sharedUsersFilePath = @"C:\Users\s.poojary\Desktop\Shared52Users.txt";

                var lstAktiv1Users = File.ReadLines(aktiv1UsersFilePath, Encoding.UTF8);
                var lstAktiv2Users = File.ReadLines(aktiv2UsersFilePath, Encoding.UTF8);
                var lstAktiv3Users = File.ReadLines(aktiv3UsersFilePath, Encoding.UTF8);
                var lstAktiv4Users = File.ReadLines(aktiv4UsersFilePath, Encoding.UTF8);
                var lstAktiv5Users = File.ReadLines(aktiv5UsersFilePath, Encoding.UTF8);
                var lstSharedUsers = File.ReadLines(sharedUsersFilePath, Encoding.UTF8);

                var sourceExcelPackage = new ExcelPackage(new FileInfo(SourceExcelFilePath));
                var destinationExcelPackage = new ExcelPackage(new FileInfo(DestinationExcelFilePath));

                var sourceWorkSheet = sourceExcelPackage.Workbook.Worksheets["extracted"];
                var destinationWorkSheet = destinationExcelPackage.Workbook.Worksheets["Sheet1"];

                Dictionary<int, string> dictSourceColoumns = new Dictionary<int, string>();
                Dictionary<int, string> dictDestinationColoumns = new Dictionary<int, string>();

                for (int i = 1; i <= sourceWorkSheet.Dimension.End.Column; i++)
                {
                    var colCell = sourceWorkSheet.Cells[1, i].Value.ToString();
                    dictSourceColoumns.Add(i, colCell);
                }
                for (int i = 1; i <= destinationWorkSheet.Dimension.End.Column; i++)
                {
                    var colCell = destinationWorkSheet.Cells[1, i].Value.ToString();
                    dictDestinationColoumns.Add(i, colCell);
                }

                int nameIndex = 1;
                //iterate the rows
                for (int sourceRowIndex = 1; sourceRowIndex <= sourceWorkSheet.Dimension.End.Row; sourceRowIndex++)
                {
                    var sourceUserName = sourceWorkSheet.Cells[sourceRowIndex, 2].Value.ToString().Trim();
                    if (!string.IsNullOrWhiteSpace(sourceUserName) && !sourceUserName.Equals("Users"))
                    {

                        for (int destinationRowIndex = 1; destinationRowIndex <= destinationWorkSheet.Dimension.End.Row; destinationRowIndex++)
                        {
                            bool isUserFound = false;
                            var destinationUserName = destinationWorkSheet.Cells[destinationRowIndex, 2].Value.ToString().Trim();
                            if (!string.IsNullOrWhiteSpace(destinationUserName) && !destinationUserName.Equals("User Name"))
                            {
                                #region code for updating aktiv1-5 status
                                //if (lstAktiv1Users.Any(str => str.Contains(destinationUserName)))
                                //{
                                //    xlDestinationSheetRange.Cells[destinationRowIndex, 3] = "Yes";
                                //}
                                //else
                                //{
                                //    xlDestinationSheetRange.Cells[destinationRowIndex, 3] = "No";
                                //}

                                //if (lstAktiv2Users.Any(str => str.Contains(destinationUserName)))
                                //{
                                //    xlDestinationSheetRange.Cells[destinationRowIndex, 4] = "Yes";
                                //}
                                //else
                                //{
                                //    xlDestinationSheetRange.Cells[destinationRowIndex, 4] = "No";
                                //}

                                //if (lstAktiv3Users.Any(str => str.Contains(destinationUserName)))
                                //{
                                //    xlDestinationSheetRange.Cells[destinationRowIndex, 5] = "Yes";
                                //}
                                //else
                                //{
                                //    xlDestinationSheetRange.Cells[destinationRowIndex, 5] = "No";
                                //}

                                //if (lstAktiv4Users.Any(str => str.Contains(destinationUserName)))
                                //{
                                //    xlDestinationSheetRange.Cells[destinationRowIndex, 6] = "Yes";
                                //}
                                //else
                                //{
                                //    xlDestinationSheetRange.Cells[destinationRowIndex, 6] = "No";
                                //}

                                //if (lstAktiv5Users.Any(str => str.Contains(destinationUserName)))
                                //{
                                //    xlDestinationSheetRange.Cells[destinationRowIndex, 7] = "Yes";
                                //}
                                //else
                                //{
                                //    xlDestinationSheetRange.Cells[destinationRowIndex, 7] = "No";
                                //}

                                //if (lstSharedUsers.Any(str => str.Contains(destinationUserName)))
                                //{
                                //    xlDestinationSheetRange.Cells[destinationRowIndex, 8] = "Yes";
                                //}
                                //else
                                //{
                                //    xlDestinationSheetRange.Cells[destinationRowIndex, 8] = "No";
                                //}
                                #endregion

                                if (sourceUserName.ToLower().Trim().Equals(destinationUserName.ToLower().Trim()))
                                {
                                    nameIndex = destinationRowIndex + 1;
                                    JobCount++;
                                    isUserFound = true;

                                    var matches = dictDestinationColoumns.Values.Intersect(dictSourceColoumns.Values);
                                    foreach (var m in matches)
                                    {
                                        var sourceColoumn = dictSourceColoumns.Where(x => m.Equals(x.Value)).Select(x => x.Key).First();
                                        var destinationColoumn = dictDestinationColoumns.Where(x => m.Equals(x.Value)).Select(x => x.Key).First();


                                        destinationWorkSheet.Cells[destinationRowIndex, destinationColoumn].Value = sourceWorkSheet.Cells[sourceRowIndex, sourceColoumn].Value;


                                    }
                                }

                            }
                            if (isUserFound)
                            {
                                break;
                            }
                        }
                        Console.WriteLine(sourceRowIndex);
                        JobCount++;
                        //AddLogs(LogFilePath + "\\", "\n" + sourceRowIndex.ToString());
                    }
                }

                destinationExcelPackage.Save();

                sourceExcelPackage.Dispose();
                destinationExcelPackage.Dispose();

                Console.WriteLine("Done");
            }
            catch (IOException)
            {
                Console.WriteLine("\nThe Excel file cannot be accessed if it is open. Please close the excel file and try again");
                AddLogs(LogFilePath + "\\", "The Excel file cannot be accessed if it is open. Please close the excel file and try again");
            }
            catch (ArgumentException ex)
            {
                Console.WriteLine("\nThe path of the excel file is not valid");
                AddLogs(LogFilePath + "\\", "The path of the excel file is not valid");
            }
            catch (InvalidFilePathException ex)
            {
                Console.WriteLine(ex.Message);
                AddLogs(LogFilePath + "\\", ex.Message);
            }
            catch (Exception ex)
            {
                Console.WriteLine("\nError please check logs at " + LogFilePath);
                AddLogs(LogFilePath + "\\", ex.Message + " stacktrace:- " + ex.StackTrace);
            }
        }

        /// <summary>
        /// Takes the folder path and file extension input from user to search the given folder and calls DeleteFileTypeFromFolder method
        /// </summary>
        private void DeleteFileType()
        {
            //path of folder to search
            string FolderPath;
            //extension of file type that is to be deleted
            string Extension;
            try
            {
                Console.WriteLine("\nEnter path of folder");
                FolderPath = Console.ReadLine();
                InitialLog.AppendLine("\nFolder path: " + FolderPath);
                if (string.IsNullOrEmpty(FolderPath))
                {
                    throw new InvalidFilePathException("Please enter valid folder path\n");
                }

                Console.WriteLine("\nEnter extension of file type to delete");
                Extension = Console.ReadLine();
                InitialLog.AppendLine("\nExtension Type: " + Extension);
                if (string.IsNullOrEmpty(Extension))
                {
                    throw new InvalidFilePathException("Please enter valid file extension\n");
                }
            }
            catch (InvalidFilePathException ex)
            {
                Console.WriteLine(ex.Message);
                AddLogs(LogFilePath + "\\", ex.Message);
                return;
            }

            LogFilePath = FolderPath;
            AddLogs(LogFilePath + "\\", InitialLog.ToString());

            DeleteFileTypeFromFolder(FolderPath, Extension);
        }

        /// <summary>
        /// Searches a folder recursively to delete all files of the specified extension. Called by DeleteFileType method
        /// </summary>
        /// <param name="path">Path of the folder to search</param>
        /// <param name="extension">Extension of the file type to delete</param>
        private void DeleteFileTypeFromFolder(string path, string extension)
        {
            foreach (var directory in Directory.GetDirectories(path))
            {
                DirectoryInfo di = new DirectoryInfo(directory);

                //FileInfo[] files = di.GetFiles("*.eml")
                //                     .Where(p => p.Extension == ".eml").ToArray();

                //get all the files of the specified extension
                FileInfo[] files = di.GetFiles("*" + extension)
                                     .Where(p => p.Extension == extension).ToArray();
                foreach (FileInfo file in files)
                {
                    try
                    {
                        //set the attribute to normal if it is something different ie. read only etc
                        file.Attributes = FileAttributes.Normal;
                        //delete the file
                        File.Delete(file.FullName);

                        Console.WriteLine("Deleted File: " + file.FullName);
                        AddLogs(LogFilePath + "\\", "\nDeleted File: " + file.FullName);
                        JobCount++;
                    }
                    catch (System.Exception ex)
                    {
                        Console.WriteLine("Error: " + ex.Message + ". File: " + file.FullName);
                        AddLogs(LogFilePath + "\\", "\nError: " + ex.Message + ". File: " + file.FullName);
                    }
                }

                //search the current directory and delete file if it matches the extension
                DeleteFileTypeFromFolder(directory, extension);
            }
        }

        /// <summary>
        /// Gets Duplicate list of users from a list of users in a text file
        /// </summary>
        private void GetSimilarNames()
        {
            string completeUsersFilePath = @"C:\Users\s.poojary\Desktop\Users.txt";
            string duplicateUsersFilePath = @"C:\Users\s.poojary\Desktop\DuplicateUsers3.txt";

            ////path of folder to search
            //string completeUsersFilePath;
            ////extension of file type that is to be deleted
            //string duplicateUsersFilePath;
            //try
            //{

            //    Console.WriteLine("\nEnter path of UserList text file");
            //    completeUsersFilePath = Console.ReadLine();
            //    //InitialLog.AppendLine("\nFolder path: " + completeUsersFilePath);
            //    if (string.IsNullOrEmpty(completeUsersFilePath))
            //    {
            //        throw new InvalidFilePathException("Please enter valid folder path\n");
            //    }

            //    Console.WriteLine("\nEnter path of duplicate users text file");
            //    duplicateUsersFilePath = Console.ReadLine();
            //    //InitialLog.AppendLine("\nExtension Type: " + duplicateUsersFilePath);
            //    if (string.IsNullOrEmpty(duplicateUsersFilePath))
            //    {
            //        throw new InvalidFilePathException("Please enter valid file extension\n");
            //    }
            //}
            //catch (InvalidFilePathException ex)
            //{
            //    Console.WriteLine(ex.Message);
            //    //AddLogs(LogFilePath + "\\", ex.Message);
            //    return;
            //}

            var lstCompleteUsers = File.ReadLines(completeUsersFilePath, Encoding.UTF8);

            //remove space, comma and hyphen(-) from the names
            var lstSanitizedUsers = lstCompleteUsers.Select(x => x.Replace("-", string.Empty).Replace(" ", string.Empty).Replace(",", string.Empty).Replace(".", string.Empty).ToLowerInvariant()).ToList();

            //get duplicate names from the list
            var lstDuplicates = lstSanitizedUsers.GroupBy(x => x).Where(group => group.Count() > 1).Select(group => group.Key).ToList();

            File.WriteAllLines(duplicateUsersFilePath, lstDuplicates);
        }

        /// <summary>
        /// Updated the aktiv folder status in the excel sheet
        /// </summary>
        private void UpdateExcelAktivStatus()
        {
            try
            {
                var DestinationExcelFilePath = @"C:\Users\s.poojary\Desktop\Aktiv1-5Users.xlsx";

                string aktiv1UsersFilePath = @"C:\Users\s.poojary\Desktop\Aktiv1Users.txt";
                string aktiv2UsersFilePath = @"C:\Users\s.poojary\Desktop\Aktiv2Users.txt";
                string aktiv3UsersFilePath = @"C:\Users\s.poojary\Desktop\Aktiv3Users.txt";
                string aktiv4UsersFilePath = @"C:\Users\s.poojary\Desktop\Aktiv4Users.txt";
                string aktiv5UsersFilePath = @"C:\Users\s.poojary\Desktop\Aktiv5Users.txt";
                string shared52UsersFilePath = @"C:\Users\s.poojary\Desktop\Shared52Users.txt";

                var lstAktiv1Users = File.ReadLines(aktiv1UsersFilePath, Encoding.UTF8);
                var lstAktiv2Users = File.ReadLines(aktiv2UsersFilePath, Encoding.UTF8);
                var lstAktiv3Users = File.ReadLines(aktiv3UsersFilePath, Encoding.UTF8);
                var lstAktiv4Users = File.ReadLines(aktiv4UsersFilePath, Encoding.UTF8);
                var lstAktiv5Users = File.ReadLines(aktiv5UsersFilePath, Encoding.UTF8);
                var lstShared52Users = File.ReadLines(shared52UsersFilePath, Encoding.UTF8);

                var excelPackage = new ExcelPackage(new FileInfo(DestinationExcelFilePath));
                var destinationWorkSheet = excelPackage.Workbook.Worksheets["Sheet1"];

                //iterate the rows
                for (int rowIndex = 1; rowIndex <= destinationWorkSheet.Dimension.End.Row; rowIndex++)
                {
                    var name = destinationWorkSheet.Cells[rowIndex, 2].Value.ToString().Trim(); ;
                    if (!string.IsNullOrWhiteSpace(name) && !name.ToLower().Equals("user name"))
                    {
                        if (lstAktiv1Users.Any(str => str.Contains(name)))
                        {
                            destinationWorkSheet.Cells[rowIndex, 3].Value = "Yes";
                        }
                        else
                        {
                            destinationWorkSheet.Cells[rowIndex, 3].Value = "No";
                        }

                        if (lstAktiv2Users.Any(str => str.Contains(name)))
                        {
                            destinationWorkSheet.Cells[rowIndex, 4].Value = "Yes";
                        }
                        else
                        {
                            destinationWorkSheet.Cells[rowIndex, 4].Value = "No";
                        }

                        if (lstAktiv3Users.Any(str => str.Contains(name)))
                        {
                            destinationWorkSheet.Cells[rowIndex, 5].Value = "Yes";
                        }
                        else
                        {
                            destinationWorkSheet.Cells[rowIndex, 5].Value = "No";
                        }

                        if (lstAktiv4Users.Any(str => str.Contains(name)))
                        {
                            destinationWorkSheet.Cells[rowIndex, 6].Value = "Yes";
                        }
                        else
                        {
                            destinationWorkSheet.Cells[rowIndex, 6].Value = "No";
                        }

                        if (lstAktiv5Users.Any(str => str.Contains(name)))
                        {
                            destinationWorkSheet.Cells[rowIndex, 7].Value = "Yes";
                        }
                        else
                        {
                            destinationWorkSheet.Cells[rowIndex, 7].Value = "No";
                        }

                        if (lstShared52Users.Any(str => str.Contains(name)))
                        {
                            destinationWorkSheet.Cells[rowIndex, 8].Value = "Yes";
                        }
                        else
                        {
                            destinationWorkSheet.Cells[rowIndex, 8].Value = "No";
                        }

                        Console.WriteLine(rowIndex);
                        JobCount++;
                    }
                }

                excelPackage.Save();
                excelPackage.Dispose();

                Console.WriteLine("Done");
            }
            catch (IOException ex)
            {
                Console.WriteLine("\nThe Excel file cannot be accessed if it is open. Please close the excel file and try again");
                AddLogs(LogFilePath + "\\", "The Excel file cannot be accessed if it is open. Please close the excel file and try again");
            }
            catch (ArgumentException ex)
            {
                Console.WriteLine("\nThe path of the excel file is not valid");
                AddLogs(LogFilePath + "\\", "The path of the excel file is not valid");
            }
            catch (InvalidFilePathException ex)
            {
                Console.WriteLine(ex.Message);
                AddLogs(LogFilePath + "\\", ex.Message);
            }
            catch (Exception ex)
            {
                Console.WriteLine("\nError please check logs at " + LogFilePath);
                AddLogs(LogFilePath + "\\", ex.Message + " stacktrace:- " + ex.StackTrace);
            }
        }

        /// <summary>
        /// Gets the count of messages from database, compare it with the value in local file and update excel sheet if it matches
        /// </summary>
        private void UpdateExcelMatchedColoumn()
        {
            try
            {
                //connection string of database
                string constring = @"server=192.168.90.61;port=3321;User Id=alr;password=auW+2Dc.sAso;database=emailarchiv_test_final;CharacterSet=utf8;Convert Zero Datetime=True;Allow Zero Datetime=True;AllowUserVariables=True";

                //path of folder where pst files are located
                string FolderPath;

                try
                {
                    Console.WriteLine("\nEnter path of pst folder");
                    FolderPath = Console.ReadLine();
                    LogFilePath = FolderPath;
                    AddLogs(LogFilePath + "\\", InitialLog.ToString());
                    AddLogs(LogFilePath + "\\", "\nPath of pst folder: " + FolderPath);
                    if (string.IsNullOrEmpty(FolderPath))
                    {
                        throw new InvalidFilePathException("Please enter valid folder path\n");
                    }
                }
                catch (InvalidFilePathException ex)
                {
                    Console.WriteLine(ex.Message);
                    AddLogs(LogFilePath + "\\", ex.Message);
                    return;
                }

                Console.WriteLine("\nEnter path of excel file");
                ExcelFilePath = Console.ReadLine();
                InitialLog.AppendLine("\nExcel file path: " + ExcelFilePath);

                var excelPackage = new ExcelPackage(new FileInfo(ExcelFilePath));
                var UserListWorkSheet = excelPackage.Workbook.Worksheets["User_list_fnl"];

                MySqlConnection con = new MySqlConnection(constring);
                con.Open();

                StreamWriter sw = new StreamWriter(LogFilePath + "\\" + string.Format("Mismatch{0}.txt", DateTime.Now.ToFileTime()), true, Encoding.UTF8);

                foreach (var folder in Directory.GetDirectories(FolderPath))
                {
                    var directoryName = folder;
                    var folderPathToSearch = FolderPath;
                    int folderCounter = 0;


                    //get the frolder name from the folder directory path
                    var strFolderName = folder.Substring(folder.LastIndexOf("\\") + 1).Replace(" ", string.Empty);
                    directoryName = strFolderName;

                    //if dash(-) exists in folder name, then remove it and its proceding characters
                    if (strFolderName.Contains("-") && int.TryParse(strFolderName.Substring(strFolderName.LastIndexOf("-") + 1, strFolderName.Length - strFolderName.LastIndexOf("_") - 2), out int n))
                    {
                        strFolderName = strFolderName.Remove(strFolderName.LastIndexOf("-"), strFolderName.Length - strFolderName.LastIndexOf("_") - 1);
                    }
                    //if folder name contains 'extern', then remove it and its preceding string
                    //if (strFolderName.Contains("-") && "extern".Contains(strFolderName.Substring(strFolderName.LastIndexOf("-") + 1).ToLower().Trim()))
                    //{
                    //    strFolderName = strFolderName.Substring(0, strFolderName.LastIndexOf("-"));
                    //}

                    strFolderName = strFolderName.Replace(",", string.Empty).Replace(" ", string.Empty).Trim();

                    //iterate the rows in excel sheet
                    for (int rowIndex = 1; rowIndex <= UserListWorkSheet.Dimension.End.Row; rowIndex++)
                    {
                        var username = UserListWorkSheet.Cells[rowIndex, 2].Value.ToString().Trim();
                        if (!string.IsNullOrWhiteSpace(username) && !username.Equals("User Name"))
                        {

                            //format username similar to folder name
                            var name = username.Replace(",", string.Empty).Replace(" ", string.Empty).Trim();


                            //remove external from the name
                            //if ("extern".Contains(name.Substring(name.LastIndexOf("-") + 1).ToLower().Trim()))
                            //{
                            //    name = name.Substring(0, name.LastIndexOf("-"));
                            //}

                            //check if a backup folder with the users name exists
                            if (strFolderName.Contains(name) || name.Contains(strFolderName))
                            {
                                try
                                {
                                    DirectoryInfo folderDirectory = new DirectoryInfo(folder);
                                    //check if folder contains log files
                                    foreach (var file in folderDirectory.GetFiles().OrderBy(f => f.Name))
                                    {
                                        if (file.Name.EndsWith(".txt"))
                                        {
                                            //if txt file is present, iterate through it
                                            var lines = File.ReadAllLines(file.FullName);
                                            for (var i = 0; i < lines.Length; i += 1)
                                            {
                                                var line = lines[i];
                                                //check if the current line contains "Items exported:"
                                                if (line.Contains("Items exported:"))
                                                {
                                                    //take the string after the colon(:) to get message count
                                                    var strLine = line.Split(':');
                                                    var count = strLine[1].Trim().ToInt();

                                                    //AddLogs(LogFilePath + "\\", count);

                                                    if (count > 0)
                                                    {
                                                        var postfach = username.Replace(",", string.Empty).Trim();
                                                        //postfach.SplitOnCapitalLetters();
                                                        postfach = string.Format("%{0}%", postfach);

                                                        string aktiv = "";

                                                        //if (directoryName.Contains("-") && int.TryParse(directoryName.Substring(directoryName.LastIndexOf("-") + 1, directoryName.Length - directoryName.LastIndexOf("_") - 2), out int m))
                                                        //{
                                                        //    directoryName = directoryName.Replace(directoryName.Substring(directoryName.LastIndexOf("-"), directoryName.Length - directoryName.LastIndexOf("_") - 1), string.Empty);
                                                        //}

                                                        //get the aktiv folder to search
                                                        if (strFolderName.Contains("-") && strFolderName.Substring(strFolderName.LastIndexOf("-")).ToLower().Trim().Contains("extern"))
                                                        {
                                                            aktiv = strFolderName.Substring(strFolderName.LastIndexOf("-") + 1);
                                                        }
                                                        else
                                                        {
                                                            aktiv = strFolderName.Substring(strFolderName.LastIndexOf("_") + 1);
                                                        }

                                                        aktiv = string.Format("%{0}%", aktiv);

                                                        Console.Write("\n" + strFolderName + " Local Count: " + count);

                                                        //get count from database
                                                        var dbCount = GetDBUserCount(con, postfach, aktiv);

                                                        //check if count of database and local value matches
                                                        if (count == dbCount.ToInt())
                                                        {

                                                            var matchedCell = UserListWorkSheet.Cells[rowIndex, 31].Value.ToString().Trim();
                                                            string cellValue = "";

                                                            //check if the cell is already populated, if it is populated append the value of the aktiv folder
                                                            if (!string.IsNullOrWhiteSpace(matchedCell))
                                                            {
                                                                cellValue = matchedCell;

                                                                if (strFolderName.Contains("Aktiv1"))
                                                                {
                                                                    cellValue += ", Aktiv1";
                                                                }
                                                                else if (strFolderName.Contains("Aktiv2"))
                                                                {
                                                                    cellValue += ", Aktiv2";
                                                                }
                                                                else if (strFolderName.Contains("Aktiv3"))
                                                                {
                                                                    cellValue += ", Aktiv3";
                                                                }
                                                                else if (strFolderName.Contains("Aktiv4"))
                                                                {
                                                                    cellValue += ", Aktiv4";
                                                                }
                                                                else if (strFolderName.Contains("Aktiv5"))
                                                                {
                                                                    cellValue += ", Aktiv5";
                                                                }
                                                            }
                                                            else
                                                            {
                                                                if (strFolderName.Contains("Aktiv1"))
                                                                {
                                                                    cellValue = "Aktiv1";
                                                                }
                                                                else if (strFolderName.Contains("Aktiv2"))
                                                                {
                                                                    cellValue = "Aktiv2";
                                                                }
                                                                else if (strFolderName.Contains("Aktiv3"))
                                                                {
                                                                    cellValue = "Aktiv3";
                                                                }
                                                                else if (strFolderName.Contains("Aktiv4"))
                                                                {
                                                                    cellValue = "Aktiv4";
                                                                }
                                                                else if (strFolderName.Contains("Aktiv5"))
                                                                {
                                                                    cellValue = "Aktiv5";
                                                                }
                                                            }

                                                            AddLogs(LogFilePath + "\\", "\n" + strFolderName + " Local Count: " + count + " DB Count: " + dbCount);
                                                            UserListWorkSheet.Cells[rowIndex, 31].Value = cellValue;
                                                        }
                                                        else
                                                        {
                                                            sw.WriteLine("\n" + strFolderName + " Local Count: " + count + " DB Count: " + dbCount);
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine("Error for user: " + username + " Please check logs at " + LogFilePath);
                                    AddLogs(LogFilePath + "\\", "Username:- " + username + " " + ex.Message + " stacktrace:- " + ex.StackTrace);
                                }
                            }
                            //else
                            //{
                            //    sw.WriteLine("Mismatch: " + strFolderName);
                            //}
                        }
                    }
                }
                con.Close();
                sw.Close();
                excelPackage.Save();
                excelPackage.Dispose();

                Console.WriteLine("Done");
            }
            catch (MySqlException ex)
            {
                Console.WriteLine("Error: {0}", ex.ToString());
                AddLogs(LogFilePath + "\\", "Error: " + ex.ToString());

            }
            catch (IOException)
            {
                Console.WriteLine("\nThe Excel file cannot be accessed if it is open. Please close the excel file and try again");
                AddLogs(LogFilePath + "\\", "The Excel file cannot be accessed if it is open. Please close the excel file and try again");
            }
            catch (ArgumentException ex)
            {
                Console.WriteLine("\nThe path of the excel file is not valid");
                AddLogs(LogFilePath + "\\", "The path of the excel file is not valid");
            }
            catch (InvalidFilePathException ex)
            {
                Console.WriteLine(ex.Message);
                AddLogs(LogFilePath + "\\", ex.Message);
            }
            catch (Exception ex)
            {
                Console.WriteLine("\nError please check logs at " + LogFilePath);
                AddLogs(LogFilePath + "\\", ex.Message + " stacktrace:- " + ex.StackTrace);
            }
        }

        /// <summary>
        /// Gets the count of messages of the particular path for a given user
        /// </summary>
        /// <param name="con">mysql connection string</param>
        /// <param name="name">name of the user</param>
        /// <param name="path">path to search</param>
        /// <returns></returns>
        private string GetDBUserCount(MySqlConnection con, string name, string path)
        {
            //string constring = @"server=192.168.90.61;port=3321;User Id=alr;password=auW+2Dc.sAso;database=emailarchiv_test_final;CharacterSet=utf8;Convert Zero Datetime=True;Allow Zero Datetime=True;AllowUserVariables=True";

            try
            {
                //using (MySqlConnection con = new MySqlConnection(constring))
                //{
                //string query = @"select count(*) from emailarchiv_test_final.emailarchiv_new where path like @Path;";
                string query = @"select count(*), postfach from emailarchiv_test_final.emailarchiv_new where Postfach like @Postfach and path like @Path;";
                //string query = @"select count(*), postfach from nord_lb_email_archiv.emailarchiv_new group by postfach ;";

                string count = null;
                using (MySqlCommand cmd = new MySqlCommand(query, con))
                {
                    cmd.CommandType = CommandType.Text;
                    cmd.Parameters.AddWithValue("@Postfach", name);
                    cmd.Parameters.AddWithValue("@Path", path);
                    //con.Open();
                    object o = cmd.ExecuteScalar();
                    if (o != null)
                    {
                        count = o.ToString();
                        Console.WriteLine(" DB Count : {0}", count);
                    }
                    //con.Close();
                    return count;
                }
                //}
            }
            catch (MySqlException ex)
            {
                Console.WriteLine("Error: {0}", ex.ToString());
                AddLogs(LogFilePath + "\\", "Error: " + ex.ToString());
                return string.Format("Error: {0}", ex.ToString());

            }
        }

        /// <summary>
        /// Generate CSV for Customer which conatins user Id, email, name etc.
        /// </summary>
        private void GenerateCSVCustomerMapping()
        {
            try
            {
                var SourceExcelFilePath = new FileInfo(@"C:\Users\s.poojary\Desktop\New folder\EmailArcheve_FinalList.xlsx");
                var UserNameExcelSheet = new FileInfo(@"C:\Users\s.poojary\Desktop\New folder\NordLB_exchange_User_for_mapping_FINAL.xlsx");
                var EmailIdExcelSheet = new FileInfo(@"C:\Users\s.poojary\Desktop\New folder\Exchange_alle_Postfächer_Aktiv12345_2019-05-02.xlsx");
                var DestinationExcelSheet = new FileInfo(@"C:\Users\s.poojary\Desktop\New folder\csv.xlsx");

                //InitialLog = new StringBuilder();
                //AddLogs(LogFilePath + "\\", "\nSource Excel file path: " + SourceExcelFilePath);
                //AddLogs(LogFilePath + "\\", "\nDestination Excel file path: " + UserNameExcelSheet);

                var sourcePackage = new ExcelPackage(SourceExcelFilePath);
                var UserNamePackage = new ExcelPackage(UserNameExcelSheet);
                var EmailIdPackage = new ExcelPackage(EmailIdExcelSheet);
                var DestinationPackage = new ExcelPackage(DestinationExcelSheet);

                var sourceWorkSheet = UserNamePackage.Workbook.Worksheets["Tabelle1"];
                var userNameWorkSheet = UserNamePackage.Workbook.Worksheets["Tabelle1"];
                var emailIdWorkSheet = EmailIdPackage.Workbook.Worksheets["Postfächer"];
                var destinationWorkSheet = DestinationPackage.Workbook.Worksheets["Sheet1"];

                //iterate the rows
                int destinationRowIndex = 1;
                #region for completing csv(ie generating csv from an incomplete csv)
                /*
                for (int sourceRowIndex = 1; sourceRowIndex <= sourceWorkSheet.Dimension.End.Row; sourceRowIndex++)
                {
                    var sourceCSV = sourceWorkSheet.Cells[sourceRowIndex, 1].Value.ToString().Trim();
                    if (!string.IsNullOrWhiteSpace(sourceCSV))
                    {
                        //split the csv file based on the delimeter
                        var csv = sourceCSV.Split(';');

                        //set the id and email id as empty
                        csv[0] = string.Empty;
                        csv[1] = string.Empty;

                        //get the username and add a comma after first name
                        var arrUserName = csv[3].Split(' ');
                        var username = string.Join(", ", arrUserName);
                        var sanitizedUsername = username.Replace("-", string.Empty).Replace(" ", string.Empty).Replace(",", string.Empty).Replace(".", string.Empty).ToLowerInvariant();

                        bool isUserIdFound = false;
                        for (int userNameRowIndex = 1; userNameRowIndex <= userNameWorkSheet.Dimension.End.Row; userNameRowIndex++)
                        {
                            var IdUserName = userNameWorkSheet.Cells[userNameRowIndex, 1].Value.ToString().Trim();
                            if (!string.IsNullOrWhiteSpace(IdUserName) && !IdUserName.Equals("ArchivName"))
                            {
                                var sanitizedIdUserName = IdUserName.Replace("-", string.Empty).Replace(" ", string.Empty).Replace(",", string.Empty).Replace(".", string.Empty).ToLowerInvariant();
                                if ((sanitizedUsername.ToLower().Trim().Equals(sanitizedIdUserName.ToLower()) || sanitizedUsername.ToLower().Trim().Contains(sanitizedIdUserName.ToLower()) || sanitizedIdUserName.ToLower().Contains(sanitizedUsername.ToLower().Trim())) && userNameWorkSheet.Cells[userNameRowIndex, 2].Value != null)
                                {
                                    isUserIdFound = true;
                                    string userId = userNameWorkSheet.Cells[userNameRowIndex, 2].Value.ToString().Trim();

                                    //if the names match then add the id to the csv array
                                    csv[0] = userId;
                                    //write the csv to the excel sheet
                                    //sourceWorkSheet.Cells[sourceRowIndex, 1].Value = string.Join(";", csv);
                                }
                            }
                        }

                        bool isEmailFound = false;
                        for (int EmailIdRowIndex = 1; EmailIdRowIndex <= emailIdWorkSheet.Dimension.End.Row; EmailIdRowIndex++)
                        {
                            var EmailIdName = emailIdWorkSheet.Cells[EmailIdRowIndex, 9].Value.ToString().Trim();
                            if (!string.IsNullOrWhiteSpace(EmailIdName.ToString()) && !EmailIdName.ToString().Trim().Equals("ArchivName"))
                            {
                                var sanitizedEmailIdName = EmailIdName.Replace("-", string.Empty).Replace(" ", string.Empty).Replace(",", string.Empty).Replace(".", string.Empty).ToLowerInvariant();
                                if ((sanitizedUsername.ToLower().Trim().Equals(sanitizedEmailIdName.ToLower()) || sanitizedUsername.ToLower().Trim().Contains(sanitizedEmailIdName.ToLower()) || sanitizedEmailIdName.ToLower().Contains(sanitizedUsername.ToLower().Trim())) && emailIdWorkSheet.Cells[EmailIdRowIndex, 2].Value != null)
                                {
                                    isEmailFound = true;
                                    string emailId = emailIdWorkSheet.Cells[EmailIdRowIndex, 2].Value.ToString().Trim();

                                    //if the names match then add the email id to the csv array
                                    csv[1] = emailId;
                                    //write the csv to the excel sheet
                                    //sourceWorkSheet.Cells[sourceRowIndex, 1].Value = string.Join(";", csv);
                                }

                            }
                        }

                        var Ids = csv[0].Split(',');
                        if (Ids.Length > 1)
                        {
                            foreach (string id in Ids)
                            {
                                string[] newCsv = csv;
                                newCsv[0] = id.Trim();
                                destinationWorkSheet.Cells[destinationRowIndex, 1].Value = string.Join(";", newCsv);

                                if (!isEmailFound)
                                {
                                    destinationWorkSheet.Cells[destinationRowIndex, 3].Value = "Email ID not found";
                                }

                                destinationRowIndex++;
                            }
                        }
                        else
                        {
                            destinationWorkSheet.Cells[destinationRowIndex, 1].Value = string.Join(";", csv);

                            if (!isUserIdFound)
                            {
                                destinationWorkSheet.Cells[destinationRowIndex, 2].Value = "User Id not found";
                            }

                            if (!isEmailFound)
                            {
                                destinationWorkSheet.Cells[destinationRowIndex, 3].Value = "Email ID not found";
                            }

                            destinationRowIndex++;
                        }

                        //sourceWorkSheet.Cells[sourceRowIndex, 1].Value = string.Join(";", csv);


                        // if (!isUserIdFound)
                        // {
                        //     sourceWorkSheet.Cells[sourceRowIndex, 2].Value = "User Id not found";
                        // }

                        // if (!isEmailFound)
                        // {
                        //     sourceWorkSheet.Cells[sourceRowIndex, 3].Value = "Email ID not found";
                        // }

                        // if (!isUserIdFound && !isEmailFound)
                        // {
                        //     sourceWorkSheet.Cells[sourceRowIndex, 1].Value = string.Join(";", csv);
                        // }

                        Console.WriteLine(string.Join(";", csv));
                        //Console.WriteLine(sourceRowIndex);
                        //JobCount++;
                        //AddLogs(LogFilePath + "\\", "\n" + sourceRowIndex.ToString());
                    }
                }
                */
                #endregion

                for (int sourceRowIndex = 1; sourceRowIndex <= sourceWorkSheet.Dimension.End.Row; sourceRowIndex++)
                {
                    var username = sourceWorkSheet.Cells[sourceRowIndex, 3].Value.ToString().Trim();
                    if (!string.IsNullOrWhiteSpace(username) && !username.Equals("ArchivName"))
                    {
                        //create a new csv array and initialize it
                        string[] csv = new string[4];
                        csv[0] = "";
                        csv[1] = "";
                        csv[2] = "emailarchiv_new";
                        csv[3] = username;
                        //remove special characters from the name
                        var sanitizedUsername = username.Replace("-", string.Empty).Replace(" ", string.Empty).Replace(",", string.Empty).Replace(".", string.Empty).ToLowerInvariant().Trim();

                        bool isUserIdFound = false;
                        for (int userNameRowIndex = 1; userNameRowIndex <= userNameWorkSheet.Dimension.End.Row; userNameRowIndex++)
                        {
                            var IdUserName = userNameWorkSheet.Cells[userNameRowIndex, 3].Value.ToString().Trim();
                            if (!string.IsNullOrWhiteSpace(IdUserName) && !IdUserName.Equals("ArchivName"))
                            {
                                var sanitizedIdUserName = IdUserName.Replace("-", string.Empty).Replace(" ", string.Empty).Replace(",", string.Empty).Replace(".", string.Empty).ToLowerInvariant().Trim();
                                if (sanitizedUsername.Equals(sanitizedIdUserName) && userNameWorkSheet.Cells[userNameRowIndex, 2].Value != null)
                                {
                                    isUserIdFound = true;
                                    string userId = userNameWorkSheet.Cells[userNameRowIndex, 2].Value.ToString().Trim();

                                    //if the names match then add the id to the csv array
                                    csv[0] = userId;
                                    //write the csv to the excel sheet
                                    //destinationWorkSheet.Cells[sourceRowIndex, 1].Value = string.Join(";", csv);
                                    break;
                                }

                            }
                        }

                        bool isEmailFound = false;
                        for (int EmailIdRowIndex = 1; EmailIdRowIndex <= emailIdWorkSheet.Dimension.End.Row; EmailIdRowIndex++)
                        {
                            var EmailIdName = emailIdWorkSheet.Cells[EmailIdRowIndex, 8].Value.ToString().Trim();
                            if (!string.IsNullOrWhiteSpace(EmailIdName.ToString()) && !EmailIdName.ToString().Trim().Equals("ArchivName"))
                            {
                                //remove special characters from the name
                                var sanitizedEmailIdName = EmailIdName.Replace("-", string.Empty).Replace(" ", string.Empty).Replace(",", string.Empty).Replace(".", string.Empty).ToLowerInvariant().Trim();
                                if (sanitizedUsername.Equals(sanitizedEmailIdName) && emailIdWorkSheet.Cells[EmailIdRowIndex, 2].Value != null)
                                {
                                    isEmailFound = true;
                                    string emailId = emailIdWorkSheet.Cells[EmailIdRowIndex, 1].Value.ToString().Trim();

                                    //if the names match then add the email id to the csv array
                                    csv[1] = emailId;
                                    //write the csv to the excel sheet
                                    //destinationWorkSheet.Cells[sourceRowIndex, 1].Value = string.Join(";", csv);
                                    break;
                                }
                            }
                        }

                        //check if the id string in the csv array contains more than 1 ids
                        var Ids = csv[0].Split(',');
                        if (Ids.Length > 1)
                        {
                            foreach (string id in Ids)
                            {
                                string[] newCsv = csv;
                                newCsv[0] = id.Trim();
                                //create a seperate entry for each individual id
                                destinationWorkSheet.Cells[destinationRowIndex, 1].Value = string.Join(";", newCsv);

                                if (!isEmailFound)
                                {
                                    destinationWorkSheet.Cells[destinationRowIndex, 3].Value = "Email ID not found";
                                }

                                destinationRowIndex++;
                            }
                        }
                        else
                        {
                            //else add the csv string in the excel sheet
                            destinationWorkSheet.Cells[destinationRowIndex, 1].Value = string.Join(";", csv);

                            if (!isUserIdFound)
                            {
                                destinationWorkSheet.Cells[destinationRowIndex, 2].Value = "User Id not found";
                            }

                            if (!isEmailFound)
                            {
                                destinationWorkSheet.Cells[destinationRowIndex, 3].Value = "Email ID not found";
                            }

                            destinationRowIndex++;
                        }

                        //destinationWorkSheet.Cells[destinationRowIndex, 1].Value = string.Join(";", csv);

                        //if (!isUserIdFound)
                        //{
                        //    destinationWorkSheet.Cells[destinationRowIndex, 2].Value = "User Id not found";
                        //}

                        //if (!isEmailFound)
                        //{
                        //    destinationWorkSheet.Cells[destinationRowIndex, 3].Value = "Email ID not found";
                        //}

                        //destinationRowIndex++;

                        Console.WriteLine(string.Join(";", csv));

                        //if (!isUserIdFound && !isEmailFound)
                        //{
                        //    destinationWorkSheet.Cells[sourceRowIndex, 1].Value = string.Join(";", csv);
                        //}

                        //Console.WriteLine(sourceRowIndex);
                        //JobCount++;
                        //AddLogs(LogFilePath + "\\", "\n" + sourceRowIndex.ToString());
                    }
                }

                //sourcePackage.Save();
                DestinationPackage.Save();

                sourcePackage.Dispose();
                UserNamePackage.Dispose();
                EmailIdPackage.Dispose();
                DestinationPackage.Dispose();

                Console.WriteLine("Done");
            }
            catch (IOException)
            {
                Console.WriteLine("\nThe Excel file cannot be accessed if it is open. Please close the excel file and try again");
                AddLogs(LogFilePath + "\\", "The Excel file cannot be accessed if it is open. Please close the excel file and try again");
            }
            catch (ArgumentException ex)
            {
                Console.WriteLine("\nThe path of the excel file is not valid");
                AddLogs(LogFilePath + "\\", "The path of the excel file is not valid");
            }
            catch (InvalidFilePathException ex)
            {
                Console.WriteLine(ex.Message);
                AddLogs(LogFilePath + "\\", ex.Message);
            }
            catch (Exception ex)
            {
                Console.WriteLine("\nError please check logs at " + LogFilePath);
                AddLogs(LogFilePath + "\\", ex.Message + " stacktrace:- " + ex.StackTrace);
            }
        }

        /// <summary>
        /// The mismatched users from the Customer csv when compared to the main excel sheet
        /// </summary>
        private void GetCSVMismatchedUsers()
        {
            string aktivUsersFilePath = @"C:\Users\s.poojary\Desktop\Aktiv1Users.txt";
            string userListFilePath = @"C:\Users\s.poojary\Desktop\NewListUsers.txt";
            string mismatchUsersFilePath = @"C:\Users\s.poojary\Desktop\MismatchUsers.txt";
            string mismatchUsersDBFilePath = @"C:\Users\s.poojary\Desktop\MismatchDB.txt";

            Dictionary<string, string> dictAktiv1 = new Dictionary<string, string>();
            Dictionary<string, string> dictUserList = new Dictionary<string, string>();

            var lstAktivUsers = File.ReadLines(aktivUsersFilePath, Encoding.UTF8);
            var lstUserList = File.ReadLines(userListFilePath, Encoding.UTF8);

            var lstSanitizedAktivUsers = lstAktivUsers.Select(x => x.Replace("-", string.Empty).Replace(" ", string.Empty).Replace(",", string.Empty).Replace(".", string.Empty)).ToList();
            var lstSanitizedUsers = lstUserList.Select(x => x.Replace("-", string.Empty).Replace(" ", string.Empty).Replace(",", string.Empty).Replace(".", string.Empty)).ToList();

            for (int i = 0; i < lstAktivUsers.Count(); i++)
            {
                dictAktiv1.Add(lstSanitizedAktivUsers[i], lstAktivUsers.ElementAt(i));
            }

            for (int i = 0; i < lstUserList.Count(); i++)
            {
                dictUserList.Add(lstSanitizedUsers[i], lstUserList.ElementAt(i));
            }
            //}

            var lstMismatchUsers = lstSanitizedUsers.Except(lstSanitizedAktivUsers);
            var lstMismatchUsersDB = lstSanitizedAktivUsers.Except(lstSanitizedUsers);

            File.WriteAllLines(mismatchUsersFilePath, lstMismatchUsers);
            File.WriteAllLines(mismatchUsersDBFilePath, lstMismatchUsersDB);
        }

        /// <summary>
        /// Generate an excel file which contains archive name and status which is copied from main excel sheet
        /// </summary>
        private void GenerateUserListStatusExcel()
        {
            try
            {
                var DestinationExcelSheet = new FileInfo(@"C:\Users\s.poojary\Desktop\MailBoxUsers.xlsx");
                using (var excelPackage = new ExcelPackage(DestinationExcelSheet))
                {
                    var extractedWorkSheet = excelPackage.Workbook.Worksheets["extracted4"];
                    var aktivWorkSheet = excelPackage.Workbook.Worksheets["Aktiv4"];
                    var destinationWorkSheet = excelPackage.Workbook.Worksheets["Aktiv4Users"];

                    int destinationSheetRowCounter = 1;

                    //loop all rows
                    for (int mailboxRowIndex = aktivWorkSheet.Dimension.Start.Row; mailboxRowIndex <= aktivWorkSheet.Dimension.End.Row; mailboxRowIndex++)
                    {
                        var mailboxUsername = aktivWorkSheet.Cells[mailboxRowIndex, 8].Value.ToString().Trim();

                        bool isUserFound = false;
                        for (int extractedRowIndex = extractedWorkSheet.Dimension.Start.Row; extractedRowIndex <= extractedWorkSheet.Dimension.End.Row; extractedRowIndex++)
                        {
                            var extractedUsername = extractedWorkSheet.Cells[extractedRowIndex, 2].Value.ToString().Trim();

                            var sanitizedMailboxName = mailboxUsername.Replace("-", string.Empty).Replace(" ", string.Empty).Replace(",", string.Empty).Replace(".", string.Empty);
                            var sanitizedExtractedName = extractedUsername.Replace("-", string.Empty).Replace(" ", string.Empty).Replace(",", string.Empty).Replace(".", string.Empty);

                            if (sanitizedMailboxName.Contains(sanitizedExtractedName) || sanitizedExtractedName.Contains(sanitizedMailboxName))
                            {
                                //if the names match, then copy its name and status to the destination excel sheet
                                isUserFound = true;
                                destinationWorkSheet.Cells[destinationSheetRowCounter, 1].Value = mailboxUsername;
                                destinationWorkSheet.Cells[destinationSheetRowCounter, 2].Value = extractedWorkSheet.GetValue(extractedRowIndex, 3);

                                destinationSheetRowCounter++;
                            }

                        }

                        Console.WriteLine(mailboxRowIndex);

                        if (!isUserFound)
                        {
                            destinationWorkSheet.Cells[destinationSheetRowCounter, 1].Value = mailboxUsername;
                            destinationWorkSheet.Cells[destinationSheetRowCounter, 2].Value = "Not Found in Extracted Sheet";

                            destinationSheetRowCounter++;
                        }
                    }

                    excelPackage.Save();

                    Console.WriteLine("Done");
                }
            }
            catch (IOException)
            {
                Console.WriteLine("\nThe Excel file cannot be accessed if it is open. Please close the excel file and try again");
                AddLogs(LogFilePath + "\\", "The Excel file cannot be accessed if it is open. Please close the excel file and try again");
            }
            catch (ArgumentException ex)
            {
                Console.WriteLine("\nThe path of the excel file is not valid");
                AddLogs(LogFilePath + "\\", "The path of the excel file is not valid");
            }
            catch (InvalidFilePathException ex)
            {
                Console.WriteLine(ex.Message);
                AddLogs(LogFilePath + "\\", ex.Message);
            }
            catch (Exception ex)
            {
                Console.WriteLine("\nError please check logs at " + LogFilePath);
                AddLogs(LogFilePath + "\\", ex.Message + " stacktrace:- " + ex.StackTrace);
            }
        }

        /// <summary>
        /// Mark the users who are present in the user mapping excel sheet
        /// </summary>
        private void MarkUserMapping()
        {
            try
            {
                var DestinationExcelSheet = new FileInfo(@"C:\Users\s.poojary\Desktop\MailBoxUsersAktiv1-4.xlsx");
                using (var excelPackage = new ExcelPackage(DestinationExcelSheet))
                {
                    var userMappingWorkSheet = excelPackage.Workbook.Worksheets["Tabelle1"];
                    var aktivWorkSheet = excelPackage.Workbook.Worksheets["Aktiv2Users"];

                    //loop all rows
                    for (int aktivRowIndex = aktivWorkSheet.Dimension.Start.Row; aktivRowIndex <= aktivWorkSheet.Dimension.End.Row; aktivRowIndex++)
                    {
                        var aktivUsername = aktivWorkSheet.Cells[aktivRowIndex, 1].Value.ToString().Trim();

                        for (int userMappingRowIndex = userMappingWorkSheet.Dimension.Start.Row; userMappingRowIndex <= userMappingWorkSheet.Dimension.End.Row; userMappingRowIndex++)
                        {
                            var userMappingUsername = userMappingWorkSheet.Cells[userMappingRowIndex, 3].Value.ToString().Trim();

                            var sanitizedAktivName = aktivUsername.Replace("-", string.Empty).Replace(" ", string.Empty).Replace(",", string.Empty).Replace(".", string.Empty);
                            var sanitizedUserMappingName = userMappingUsername.Replace("-", string.Empty).Replace(" ", string.Empty).Replace(",", string.Empty).Replace(".", string.Empty);

                            if (sanitizedAktivName.Contains(sanitizedUserMappingName) || sanitizedUserMappingName.Contains(sanitizedAktivName))
                            {
                                //if the names match, then set the background color of the row to yellow
                                aktivWorkSheet.Row(aktivRowIndex).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                aktivWorkSheet.Row(aktivRowIndex).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                            }

                        }

                        Console.WriteLine(aktivRowIndex);
                    }

                    excelPackage.Save();

                    Console.WriteLine("Done");
                }
            }
            catch (IOException)
            {
                Console.WriteLine("\nThe Excel file cannot be accessed if it is open. Please close the excel file and try again");
                AddLogs(LogFilePath + "\\", "The Excel file cannot be accessed if it is open. Please close the excel file and try again");
            }
            catch (ArgumentException ex)
            {
                Console.WriteLine("\nThe path of the excel file is not valid");
                AddLogs(LogFilePath + "\\", "The path of the excel file is not valid");
            }
            catch (InvalidFilePathException ex)
            {
                Console.WriteLine(ex.Message);
                AddLogs(LogFilePath + "\\", ex.Message);
            }
            catch (Exception ex)
            {
                Console.WriteLine("\nError please check logs at " + LogFilePath);
                AddLogs(LogFilePath + "\\", ex.Message + " stacktrace:- " + ex.StackTrace);
            }
        }

        private void CreateAktivExcelSheet()
        {
            try
            {
                var MailboxExcelFilePath = @"C:\Users\s.poojary\Desktop\Exchange_alle_Postfächer_Aktiv12345_2019-05-02.xlsx";
                var UserIdExcelFilePath = @"C:\Users\s.poojary\Desktop\NordLB_exchange_User_for_mapping_FINAL.xlsx";
                var ExtractedExcelFilePath = @"C:\Users\s.poojary\Desktop\Extracted_Report10-5-19.xlsx";
                var DestinationExcelFilePath = @"C:\Users\s.poojary\Desktop\Aktiv2.xlsx";

                var mailboxExcelPackage = new ExcelPackage(new FileInfo(MailboxExcelFilePath));
                var userIdExcelPackage = new ExcelPackage(new FileInfo(UserIdExcelFilePath));
                var extractedExcelPackage = new ExcelPackage(new FileInfo(ExtractedExcelFilePath));
                var destinationExcelPackage = new ExcelPackage(new FileInfo(DestinationExcelFilePath));

                var mailboxWorkSheet = mailboxExcelPackage.Workbook.Worksheets["Sheet1"];
                var userIdWorkSheet = userIdExcelPackage.Workbook.Worksheets["Tabelle1"];
                var extractedWorkSheet = extractedExcelPackage.Workbook.Worksheets["extracted"];
                var destinationWorkSheet = destinationExcelPackage.Workbook.Worksheets["Sheet1"];

                int destinationRowIndex = 2;
                //iterate the rows
                for (int mailboxRowIndex = 2; mailboxRowIndex <= mailboxWorkSheet.Dimension.End.Row; mailboxRowIndex++)
                {
                    var mailboxUsername = mailboxWorkSheet.Cells[mailboxRowIndex, 8].Value.ToString().Trim();

                    bool isUserFound = false;
                    bool isUserFoundInMapping = false;

                    for (int userIdRowIndex = 2; userIdRowIndex <= userIdWorkSheet.Dimension.End.Row; userIdRowIndex++)
                    {
                        var userIdUsername = userIdWorkSheet.Cells[userIdRowIndex, 3].Value.ToString().Trim();

                        var sanitizedmailboxName = mailboxUsername.Replace("-", string.Empty).Replace(" ", string.Empty).Replace(",", string.Empty).Replace(".", string.Empty).ToLowerInvariant();
                        var sanitizedUserIdName = userIdUsername.Replace("-", string.Empty).Replace(" ", string.Empty).Replace(",", string.Empty).Replace(".", string.Empty).ToLowerInvariant();


                        if (sanitizedmailboxName.Contains(sanitizedUserIdName) || sanitizedUserIdName.Contains(sanitizedmailboxName))
                        {
                            isUserFoundInMapping = true;
                            for (int extractedRowIndex = 2; extractedRowIndex <= extractedWorkSheet.Dimension.End.Row; extractedRowIndex++)
                            {
                                var userMappingUsername = extractedWorkSheet.Cells[extractedRowIndex, 2].Value.ToString().Trim();

                                //var sanitizedmailboxName = mailboxUsername.Replace("-", string.Empty).Replace(" ", string.Empty).Replace(",", string.Empty).Replace(".", string.Empty);
                                var sanitizedExtractedName = userMappingUsername.Replace("-", string.Empty).Replace(" ", string.Empty).Replace(",", string.Empty).Replace(".", string.Empty).ToLowerInvariant();

                                if (sanitizedmailboxName.Contains(sanitizedExtractedName) || sanitizedExtractedName.Contains(sanitizedmailboxName))
                                {
                                    isUserFound = true;
                                    extractedWorkSheet.Cells[extractedRowIndex, 2, extractedRowIndex, 14].Copy(destinationWorkSheet.Cells[destinationRowIndex, 2, destinationRowIndex, 14]);
                                    destinationRowIndex++;

                                }
                            }

                        }
                    }


                    if (!isUserFoundInMapping)
                    {
                        destinationWorkSheet.Cells[destinationRowIndex, 2].Value = mailboxUsername;
                        destinationWorkSheet.Cells[destinationRowIndex, 15].Value = "Not Required";
                        destinationRowIndex++;
                    }


                    if (isUserFoundInMapping && !isUserFound)
                    {
                        destinationWorkSheet.Cells[destinationRowIndex, 2].Value = mailboxUsername;
                        destinationWorkSheet.Cells[destinationRowIndex, 15].Value = "Not Found in extracted";
                        destinationRowIndex++;
                    }

                    Console.WriteLine(mailboxRowIndex);
                }

                destinationExcelPackage.Save();

                mailboxExcelPackage.Dispose();
                extractedExcelPackage.Dispose();
                destinationExcelPackage.Dispose();

                Console.WriteLine("Done");
            }
            catch (IOException)
            {
                Console.WriteLine("\nThe Excel file cannot be accessed if it is open. Please close the excel file and try again");
                AddLogs(LogFilePath + "\\", "The Excel file cannot be accessed if it is open. Please close the excel file and try again");
            }
            catch (ArgumentException ex)
            {
                Console.WriteLine("\nThe path of the excel file is not valid");
                AddLogs(LogFilePath + "\\", "The path of the excel file is not valid");
            }
            catch (InvalidFilePathException ex)
            {
                Console.WriteLine(ex.Message);
                AddLogs(LogFilePath + "\\", ex.Message);
            }
            catch (Exception ex)
            {
                Console.WriteLine("\nError please check logs at " + LogFilePath);
                AddLogs(LogFilePath + "\\", ex.Message + " stacktrace:- " + ex.StackTrace);
            }
        }

        private void GetCSVFromPostfach()
        {
            try
            {
                var SourceExcelFilePath = new FileInfo(@"C:\Users\s.poojary\Desktop\OKcsv.xlsx");
                var UserCountExcelSheet = new FileInfo(@"C:\Users\s.poojary\Desktop\Email Count.xlsx");
                var DestinationExcelSheet = new FileInfo(@"C:\Users\s.poojary\Desktop\FinalCSV_17-5-19.xlsx");

                //InitialLog = new StringBuilder();
                //AddLogs(LogFilePath + "\\", "\nSource Excel file path: " + SourceExcelFilePath);
                //AddLogs(LogFilePath + "\\", "\nDestination Excel file path: " + UserNameExcelSheet);

                var sourcePackage = new ExcelPackage(SourceExcelFilePath);
                var UserCountPackage = new ExcelPackage(UserCountExcelSheet);
                var DestinationPackage = new ExcelPackage(DestinationExcelSheet);

                var sourceWorkSheet = sourcePackage.Workbook.Worksheets["Sheet1"];
                var userCountWorkSheet = UserCountPackage.Workbook.Worksheets["Sheet1"];
                var destinationWorkSheet = DestinationPackage.Workbook.Worksheets["Sheet2"];

                //iterate the rows
                int destinationRowIndex = 1;

                for (int userCountRowIndex = 1; userCountRowIndex <= userCountWorkSheet.Dimension.End.Row; userCountRowIndex++)
                {
                    var IdUserName = userCountWorkSheet.Cells[userCountRowIndex, 2].Value.ToString().Trim();
                    if (!string.IsNullOrWhiteSpace(IdUserName) && !IdUserName.Equals("Postfach"))
                    {
                        var sanitizedIdUserName = IdUserName.Replace("-", string.Empty).Replace(" ", string.Empty).Replace(",", string.Empty).Replace(".", string.Empty).ToLowerInvariant().Trim();

                        bool isUserFound = false;
                        for (int sourceRowIndex = 2; sourceRowIndex <= sourceWorkSheet.Dimension.End.Row; sourceRowIndex++)
                        {
                            var sourceCSV = sourceWorkSheet.Cells[sourceRowIndex, 1].Value.ToString().Trim();
                            if (!string.IsNullOrWhiteSpace(sourceCSV))
                            {
                                //split the csv file based on the delimeter
                                var csv = sourceCSV.Split(';');

                                //get the username and add a comma after first name
                                //var arrUserName = csv[3].Split(' ');
                                //var username = string.Join(", ", arrUserName);
                                var username = csv[3];
                                var sanitizedUsername = username.Replace("-", string.Empty).Replace(" ", string.Empty).Replace(",", string.Empty).Replace(".", string.Empty).ToLowerInvariant().Trim();

                                if (sanitizedIdUserName.Equals(sanitizedUsername))
                                {
                                    isUserFound = true;
                                    destinationWorkSheet.Cells[destinationRowIndex, 1].Value = sourceWorkSheet.Cells[sourceRowIndex, 1].Value;
                                    destinationWorkSheet.Cells[destinationRowIndex, 2].Value = sourceWorkSheet.Cells[sourceRowIndex, 2].Value;
                                    destinationWorkSheet.Cells[destinationRowIndex, 3].Value = sourceWorkSheet.Cells[sourceRowIndex, 3].Value;
                                    destinationWorkSheet.Cells[destinationRowIndex, 4].Value = userCountWorkSheet.Cells[userCountRowIndex, 1].Value;
                                    destinationRowIndex++;
                                }
                            }
                        }
                        if (!isUserFound)
                        {
                            Console.WriteLine("Not Found: " + IdUserName);
                        }
                        else
                        {
                            Console.WriteLine("Found: " + IdUserName);
                        }
                    }
                }

                DestinationPackage.Save();

                sourcePackage.Dispose();
                UserCountPackage.Dispose();
                DestinationPackage.Dispose();

                Console.WriteLine("Done");
            }
            catch (IOException)
            {
                Console.WriteLine("\nThe Excel file cannot be accessed if it is open. Please close the excel file and try again");
                AddLogs(LogFilePath + "\\", "The Excel file cannot be accessed if it is open. Please close the excel file and try again");
            }
            catch (ArgumentException ex)
            {
                Console.WriteLine("\nThe path of the excel file is not valid");
                AddLogs(LogFilePath + "\\", "The path of the excel file is not valid");
            }
            catch (InvalidFilePathException ex)
            {
                Console.WriteLine(ex.Message);
                AddLogs(LogFilePath + "\\", ex.Message);
            }
            catch (Exception ex)
            {
                Console.WriteLine("\nError please check logs at " + LogFilePath);
                AddLogs(LogFilePath + "\\", ex.Message + " stacktrace:- " + ex.StackTrace);
            }
        }

        private void GetFolderNameForUser()
        {
            string ExcelFileName = @"C:\Users\s.poojary\Desktop\UsersFolders.xlsx";
            string FolderName = @"D:\Sachith\PstTest\pstofbothAktiv1andAktiv2";

            try
            {
                var ExcelSheet = new FileInfo(ExcelFileName);
                using (var excelPackage = new ExcelPackage(ExcelSheet))
                {
                    var WorkSheet = excelPackage.Workbook.Worksheets["Sheet1"];

                    //loop all rows
                    for (int rowIndex = 2; rowIndex <= WorkSheet.Dimension.End.Row; rowIndex++)
                    {
                        bool isUserFound = false;
                        var username = WorkSheet.Cells[rowIndex, 1].Value.ToString().Trim();
                        var sanitizedAktivName = username.Replace("-", string.Empty).Replace(" ", string.Empty).Replace(",", string.Empty).Replace(".", string.Empty).ToLowerInvariant();

                        try
                        {
                            //var directory = Directory.GetDirectories(FolderName).ToList();

                            //var a = directory.Where(x =>
                            //{
                            //    var sanitizedFolderName = x.Replace("-", string.Empty).Replace(" ", string.Empty).Replace(",", string.Empty).Replace(".", string.Empty);
                            //    if (sanitizedFolderName.Contains(sanitizedAktivName) || sanitizedAktivName.Contains(sanitizedFolderName))
                            //    {
                            //        Console.WriteLine(rowIndex);
                            //        AddLogs(LogFilePath + "\\", "Found: " + username);
                            //        return true;
                            //    }
                            //    else
                            //    {
                            //        Console.WriteLine("Not Found: " + username);
                            //        AddLogs(LogFilePath + "\\", "Not Found: " + username);
                            //        return false;
                            //    }

                            //}).ToList();

                            foreach (var directory in Directory.GetDirectories(FolderName))
                            {
                                var sanitizedFolderName = directory.Replace("-", string.Empty).Replace(" ", string.Empty).Replace(",", string.Empty).Replace(".", string.Empty).ToLowerInvariant();

                                if ((sanitizedFolderName.Contains(sanitizedAktivName) || sanitizedAktivName.Contains(sanitizedFolderName)) && sanitizedFolderName.Contains("aktiv"))
                                {
                                    string cellText = null;

                                    if (WorkSheet.Cells[rowIndex, 2].Value != null)
                                    {
                                        cellText = WorkSheet.Cells[rowIndex, 2].Value.ToString().Trim();
                                        cellText = "\n" + directory;
                                    }
                                    else
                                    {
                                        cellText = directory;
                                    }

                                    WorkSheet.Cells[rowIndex, 2].Value = cellText;
                                    Console.WriteLine(rowIndex);
                                    AddLogs(LogFilePath + "\\", "Found: " + username);
                                    isUserFound = true;
                                }

                                //var lines = File.ReadAllLines(filePath);
                                //for (var i = 0; i < lines.Length; i += 1)
                                //{
                                //    var line = lines[i];
                                //    line = line.Replace(",", string.Empty);
                                //    //check if the current line contains the user name
                                //    if (directory.Trim().ToLower().Contains(line.Trim().ToLower()))
                                //    {
                                //        //sw.WriteLine(directory);
                                //        AddLogs(LogFilePath + "\\", "\nUser Found: " + line + " - " + directory);
                                //        Console.WriteLine(directory);
                                //    }
                                //}

                            }
                        }
                        catch (System.Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }

                        if (!isUserFound)
                        {
                            Console.WriteLine("Not Found: " + username);
                            AddLogs(LogFilePath + "\\", "Not Found: " + username);
                        }
                    }

                    excelPackage.Save();

                    Console.WriteLine("Done");
                }
            }
            catch (IOException ex)
            {
                Console.WriteLine("\nThe Excel file cannot be accessed if it is open. Please close the excel file and try again");
                AddLogs(LogFilePath + "\\", ex.Message + " stacktrace:- " + ex.StackTrace);
            }
            catch (ArgumentException ex)
            {
                Console.WriteLine("\nThe path of the excel file is not valid");
                AddLogs(LogFilePath + "\\", ex.Message + " stacktrace:- " + ex.StackTrace);
            }
            catch (InvalidFilePathException ex)
            {
                Console.WriteLine(ex.Message);
                AddLogs(LogFilePath + "\\", ex.Message);
            }
            catch (Exception ex)
            {
                Console.WriteLine("\nError please check logs at " + LogFilePath);
                AddLogs(LogFilePath + "\\", ex.Message + " stacktrace:- " + ex.StackTrace);
            }
        }
    }
}