using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Collections.Generic;

namespace PstDataExtractionTools
{
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
            //prog.UpdateExcelForAllAktivUsers();
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
                "10) Exit\n");
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
                    Console.WriteLine("Add Folder name to Excel");
                    prog.InitialLog.AppendLine("\nAdd Folder name to Excel");
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

                Microsoft.Office.Interop.Excel.Application xlApp;
                Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                Microsoft.Office.Interop.Excel.Sheets xlBigSheet;
                Microsoft.Office.Interop.Excel.Range xlSheetRange;

                xlApp = new Microsoft.Office.Interop.Excel.Application();
                //sets whether the excel file will be open during this process
                xlApp.Visible = false;
                //open the excel file
                xlWorkBook = xlApp.Workbooks.Open(ExcelFilePath, 0,
                            false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows,
                             "", true, false, 0, true, false, false);

                //get all the worksheets in the excel  file
                xlBigSheet = xlWorkBook.Worksheets;

                Console.WriteLine("\nEnter excel sheet name");
                var xlSheetName = Console.ReadLine();

                //string x = "Extracted";
                //get the specified worksheet
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlBigSheet.get_Item(xlSheetName);

                xlSheetRange = xlWorkSheet.UsedRange;

                int colCount = xlSheetRange.Columns.Count;
                int rowCount = xlSheetRange.Rows.Count;
                //iterate the rows
                for (int index = 0; index <= rowCount; index++)
                {
                    Microsoft.Office.Interop.Excel.Range cell = xlSheetRange.Cells[index, 2];
                    if (cell.Value2 != null && !string.IsNullOrWhiteSpace(cell.Value2.ToString()) && !cell.Value2.ToString().Trim().Equals("User name"))
                    {
                        //Rename and move files and folder
                        GetAndMoveFoldersAndFiles(Aktiv1FolderPath.Replace(" ", string.Empty), cell.Value2.ToString().Replace(" ", string.Empty));
                        GetAndMoveFoldersAndFiles(Aktiv2FolderPath.Replace(" ", string.Empty), cell.Value2.ToString().Replace(" ", string.Empty));

                        //Search and Move already renamed folders
                        //MoveFolder("Aktiv1", Aktiv1FolderPath.Replace(" ", string.Empty), cell.Value2.ToString().Replace(",", string.Empty));
                        //MoveFolder("Aktiv2", Aktiv2FolderPath.Replace(" ", string.Empty), cell.Value2.ToString().Replace(",", string.Empty));

                        //Check if renamed folders exists for a user
                        //CheckFolders("Aktiv1", Aktiv1FolderPath.Replace(" ", string.Empty), cell.Value2.ToString().Replace(",", string.Empty));
                        //CheckFolders("Aktiv2", Aktiv2FolderPath.Replace(" ", string.Empty), cell.Value2.ToString().Replace(",", string.Empty));
                    }
                }

                //xlWorkBook.Save();

                //this line causes the excel file to get corrupted
                //xlWorkBook.SaveAs(excelFilePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal,
                //        Missing.Value, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive,
                //        Missing.Value, Missing.Value, Missing.Value,
                //        Missing.Value, Missing.Value);

                //cleanup
                xlWorkBook.Close(Missing.Value, Missing.Value, Missing.Value);
                xlWorkBook = null;
                xlApp.Quit();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
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
                    //if (strFolderName.Contains("-") && int.TryParse(strFolderName.Substring(strFolderName.LastIndexOf("-") + 1), out int n))
                    //{
                    //    strFolderName = strFolderName.Substring(0, strFolderName.LastIndexOf("-"));
                    //}

                    /*uncomment if you want to remove 'extern' from file name*/
                    //if folder name contains 'extern', then remove it and its preceding string
                    //if (strFolderName.Contains("-") && "extern".Contains(strFolderName.Substring(strFolderName.LastIndexOf("-") + 1).ToLower().Trim()))
                    //{
                    //    strFolderName = strFolderName.Substring(0, strFolderName.LastIndexOf("-"));
                    //}

                    strFolderName += "-" + currentActivFolder;

                    int folderCounter = 0;
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
                                //var newFileName = file.DirectoryName + "\\" + file.DirectoryName.Substring(file.DirectoryName.LastIndexOf("\\") + 1) + ".pst";
                                var newFileName = file.DirectoryName + "\\" + strFolderName + (folderCounter == 0 ? string.Empty : folderCounter.ToString()) + ".pst";
                                folderCounter++;
                                Directory.Move(file.FullName, newFileName);
                                Console.WriteLine("Renamed File: " + file.FullName + " to " + newFileName);
                                AddLogs(LogFilePath + "\\", "Renamed File: " + file.FullName + " to " + newFileName);
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
            string allUsersFilePath = @"C:\Users\s.poojary\Desktop\FullUserList.txt";
            string aktiv1UsersFilePath = @"C:\Users\s.poojary\Desktop\Aktiv1Users.txt";
            string aktiv2UsersFilePath = @"C:\Users\s.poojary\Desktop\Aktiv2Users.txt";
            string aktiv3UsersFilePath = @"C:\Users\s.poojary\Desktop\Aktiv3Users.txt";
            string aktiv4UsersFilePath = @"C:\Users\s.poojary\Desktop\Aktiv4Users.txt";
            string aktiv5UsersFilePath = @"C:\Users\s.poojary\Desktop\Aktiv5Users.txt";
            string completeUsersFilePath = @"C:\Users\s.poojary\Desktop\CompleteUsers.txt";

            var fullUsersArray = File.ReadAllLines(allUsersFilePath, Encoding.UTF8);
            //var aktiv1UsersArray = File.ReadAllLines(aktiv1UsersFilePath, Encoding.UTF8);
            //var aktiv2UsersArray = File.ReadAllLines(aktiv2UsersFilePath, Encoding.UTF8);
            var aktiv3UsersArray = File.ReadAllLines(aktiv3UsersFilePath, Encoding.UTF8);
            var aktiv4UsersArray = File.ReadAllLines(aktiv4UsersFilePath, Encoding.UTF8);
            var aktiv5UsersArray = File.ReadAllLines(aktiv5UsersFilePath, Encoding.UTF8);

            var lstFullUsers = fullUsersArray.ToList();
            //var lstAktiv1Users = aktiv1UsersArray.ToList();
            //var lstAktiv2Users = aktiv2UsersArray.ToList();
            var lstAktiv3Users = aktiv3UsersArray.ToList();
            var lstAktiv4Users = aktiv4UsersArray.ToList();
            var lstAktiv5Users = aktiv5UsersArray.ToList();

            var lstCompleteUsers = lstFullUsers.Union(lstAktiv3Users).Union(lstAktiv4Users).Union(lstAktiv5Users).ToList();
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
                var SourceExcelFilePath = @"C:\Users\s.poojary\Desktop\Extracted_Report16Apr14.18.xlsx";
                var DestinationExcelFilePath = @"C:\Users\s.poojary\Desktop\Aktiv1-5Users.xlsx";
                //InitialLog = new StringBuilder();
                AddLogs(LogFilePath + "\\", "\nSource Excel file path: " + SourceExcelFilePath);
                AddLogs(LogFilePath + "\\", "\nDestination Excel file path: " + DestinationExcelFilePath);

                //string aktiv1UsersFilePath = @"C:\Users\s.poojary\Desktop\UserList\Aktiv1Users.txt";
                //string aktiv2UsersFilePath = @"C:\Users\s.poojary\Desktop\UserList\Aktiv2Users.txt";
                //string aktiv3UsersFilePath = @"C:\Users\s.poojary\Desktop\UserList\Aktiv3Users.txt";
                //string aktiv4UsersFilePath = @"C:\Users\s.poojary\Desktop\UserList\Aktiv4Users.txt";
                //string aktiv5UsersFilePath = @"C:\Users\s.poojary\Desktop\UserList\Aktiv5Users.txt";
                //string sharedUsersFilePath = @"C:\Users\s.poojary\Desktop\UserList\SharedUsers.txt";

                //var lstAktiv1Users = File.ReadLines(aktiv1UsersFilePath, Encoding.UTF8);
                //var lstAktiv2Users = File.ReadLines(aktiv2UsersFilePath, Encoding.UTF8);
                //var lstAktiv3Users = File.ReadLines(aktiv3UsersFilePath, Encoding.UTF8);
                //var lstAktiv4Users = File.ReadLines(aktiv4UsersFilePath, Encoding.UTF8);
                //var lstAktiv5Users = File.ReadLines(aktiv5UsersFilePath, Encoding.UTF8);
                //var lstSharedUsers = File.ReadLines(sharedUsersFilePath, Encoding.UTF8);

                #region Source excel sheet setup
                Microsoft.Office.Interop.Excel.Application xlSourceApp;
                Microsoft.Office.Interop.Excel.Workbook xlSourceWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet xlSourceWorkSheet;
                Microsoft.Office.Interop.Excel.Sheets xlSourceBigSheet;
                Microsoft.Office.Interop.Excel.Range xlSourceSheetRange;

                xlSourceApp = new Microsoft.Office.Interop.Excel.Application();
                //sets whether the excel file will be open during this process
                xlSourceApp.Visible = false;
                //open the excel file
                xlSourceWorkBook = xlSourceApp.Workbooks.Open(SourceExcelFilePath, 0,
                            false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows,
                             "", true, false, 0, true, false, false);

                //get all the worksheets in the excel  file
                xlSourceBigSheet = xlSourceWorkBook.Worksheets;

                string xlSourceSheetName = "Aktiv5_Users";

                //get the specified worksheet
                xlSourceWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlSourceBigSheet.get_Item(xlSourceSheetName);

                xlSourceSheetRange = xlSourceWorkSheet.UsedRange;

                int sourceColCount = xlSourceSheetRange.Columns.Count;
                int sourceRowCount = xlSourceSheetRange.Rows.Count;
                #endregion

                #region Destination sheet setup
                Microsoft.Office.Interop.Excel.Application xlDestinationApp;
                Microsoft.Office.Interop.Excel.Workbook xlDestinationWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet xlDestinationWorkSheet;
                Microsoft.Office.Interop.Excel.Sheets xlDestinationBigSheet;
                Microsoft.Office.Interop.Excel.Range xlDestinationSheetRange;

                xlDestinationApp = new Microsoft.Office.Interop.Excel.Application();
                //sets whether the excel file will be open during this process
                xlDestinationApp.Visible = false;
                //open the excel file
                xlDestinationWorkBook = xlDestinationApp.Workbooks.Open(DestinationExcelFilePath, 0,
                            false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows,
                             "", true, false, 0, true, false, false);

                //get all the worksheets in the excel  file
                xlDestinationBigSheet = xlDestinationWorkBook.Worksheets;

                string xlDestinationSheetName = "Sheet1";

                //get the specified worksheet
                xlDestinationWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlDestinationBigSheet.get_Item(xlDestinationSheetName);

                xlDestinationSheetRange = xlDestinationWorkSheet.UsedRange;

                int destinationColCount = xlDestinationSheetRange.Columns.Count;
                int destinationRowCount = xlDestinationSheetRange.Rows.Count;
                #endregion

                Dictionary<int, string> dictSourceColoumns = new Dictionary<int, string>();
                Dictionary<int, string> dictDestinationColoumns = new Dictionary<int, string>();

                for (int i = 1; i <= sourceColCount; i++)
                {
                    Microsoft.Office.Interop.Excel.Range colCell = xlSourceSheetRange.Cells[1, i];
                    dictSourceColoumns.Add(i, colCell.Value2.ToString());
                }
                for (int i = 1; i <= destinationColCount; i++)
                {
                    Microsoft.Office.Interop.Excel.Range colCell = xlDestinationSheetRange.Cells[1, i];
                    dictDestinationColoumns.Add(i, colCell.Value2.ToString());
                }

                int nameIndex = 1;
                //iterate the rows
                for (int sourceRowIndex = 1; sourceRowIndex <= sourceRowCount; sourceRowIndex++)
                {
                    Microsoft.Office.Interop.Excel.Range nameCell = xlSourceSheetRange.Cells[sourceRowIndex, 2];
                    if (nameCell.Value2 != null && !string.IsNullOrWhiteSpace(nameCell.Value2.ToString()) && !nameCell.Value2.ToString().ToLower().Trim().Equals("Users"))
                    {
                        string sourceUserName = nameCell.Value2.ToString();

                        for (int destinationRowIndex = nameIndex; destinationRowIndex <= destinationRowCount; destinationRowIndex++)
                        {
                            Microsoft.Office.Interop.Excel.Range cell = xlDestinationSheetRange.Cells[destinationRowIndex, 2];
                            if (cell.Value2 != null && !string.IsNullOrWhiteSpace(cell.Value2.ToString()) && !cell.Value2.ToString().ToLower().Trim().Equals("user name"))
                            {
                                string destinationUserName = cell.Value2.ToString();
                                #region code for updated aktiv1-5 status
                                //if (lstAktiv1Users.Any(str => str.Contains(userName)))
                                //{
                                //    xlDestinationSheetRange.Cells[destinationRowIndex, 3] = "Yes";
                                //}
                                //else
                                //{
                                //    xlDestinationSheetRange.Cells[destinationRowIndex, 3] = "No";
                                //}

                                //if(lstAktiv2Users.Any(str => str.Contains(userName)))
                                //{
                                //    xlDestinationSheetRange.Cells[destinationRowIndex, 4] = "Yes";
                                //}
                                //else
                                //{
                                //    xlDestinationSheetRange.Cells[destinationRowIndex, 4] = "No";
                                //}

                                //if(lstAktiv3Users.Any(str => str.Contains(userName)))
                                //{
                                //    xlDestinationSheetRange.Cells[destinationRowIndex, 5] = "Yes";
                                //}
                                //else
                                //{
                                //    xlDestinationSheetRange.Cells[destinationRowIndex, 5] = "No";
                                //}

                                //if(lstAktiv4Users.Any(str => str.Contains(userName)))
                                //{
                                //    xlDestinationSheetRange.Cells[destinationRowIndex, 6] = "Yes";
                                //}
                                //else
                                //{
                                //    xlDestinationSheetRange.Cells[destinationRowIndex, 6] = "No";
                                //}

                                //if(lstAktiv5Users.Any(str => str.Contains(userName)))
                                //{
                                //    xlDestinationSheetRange.Cells[destinationRowIndex, 7] = "Yes";
                                //}
                                //else
                                //{
                                //    xlDestinationSheetRange.Cells[destinationRowIndex, 7] = "No";
                                //}

                                //if(lstSharedUsers.Any(str => str.Contains(userName)))
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

                                    var matches = dictDestinationColoumns.Values.Intersect(dictSourceColoumns.Values);
                                    foreach (var m in matches)
                                    {
                                        var sourceColoumn = dictSourceColoumns.Where(x => m.Equals(x.Value)).Select(x => x.Key).First();
                                        var destinationColoumn = dictDestinationColoumns.Where(x => m.Equals(x.Value)).Select(x => x.Key).First();


                                        xlDestinationSheetRange.Cells[destinationRowIndex, destinationColoumn] = xlSourceSheetRange[sourceRowIndex, sourceColoumn];


                                    }


                                    //xlDestinationSheetRange.Cells[destinationRowIndex, 9] = xlSourceSheetRange[sourceRowIndex, 3];
                                }

                            }
                        }
                        Console.WriteLine(sourceRowIndex);
                        //AddLogs(LogFilePath + "\\", "\n" + sourceRowIndex.ToString());
                    }
                }




                xlDestinationWorkBook.Save();

                //cleanup
                xlSourceWorkBook.Close(Missing.Value, Missing.Value, Missing.Value);
                xlSourceWorkBook = null;
                xlSourceApp.Quit();

                xlDestinationWorkBook.Close(Missing.Value, Missing.Value, Missing.Value);
                xlDestinationWorkBook = null;
                xlDestinationApp.Quit();

                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

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
    }
}