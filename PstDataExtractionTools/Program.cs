using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Reflection;

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
            prog.InitialLog = new StringBuilder();

            DateTime startTime = DateTime.Now;
            Console.WriteLine("-----------------------------------Start----------------------------------------");
            prog.InitialLog.AppendLine("\n-----------------------------------Start of Log----------------------------------------");

            Console.WriteLine("Please select option \n" +
                "1) Read Excel File and Move & Rename folders \n" +
                "2) Remove .pst from folder name \n" +
                "3) Remove unwanted folders from destination path \n" +
                "4) Rename PST Files\n" +
                "5) Get Mismatch count from log file\n" +
                "6) Search Directory for backup folder\n" +
                "7) Exit");
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

        /*Incomplete*/
        private void WriteToExcel(string excelFilePath)
        {
            //Console.WriteLine("\nEnter path of excel file");
            //ExcelFilePath = Console.ReadLine();
            //InitialLog.AppendLine("\nExcel file path: " + ExcelFilePath);

            //ExcelFilePath = @"D:\Sachith\TestUsers.xlsx";
            //string excelFilePath = System.Web.HttpUtility.HtmlEncode(@"https://avendatagmbh-my.sharepoint.com/:x:/r/personal/s_poojary_avendata_com/_layouts/15/Doc.aspx?sourcedoc=%7BBB1C2FB1-3ABC-4856-B53E-EE150E98F64A%7D&file=Book%201.xlsx&action=editnew&mobileredirect=true&wdNewAndOpenCt=1554865062581&wdPreviousSession=8cb444c0-ac3b-4ff6-a931-b466fe30dc56&wdOrigin=ohpAppStartPages");

            Microsoft.Office.Interop.Excel.Application xlApp;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            Microsoft.Office.Interop.Excel.Sheets xlBigSheet;
            Microsoft.Office.Interop.Excel.Range xlSheetRange;

            xlApp = new Microsoft.Office.Interop.Excel.Application();
            //sets whether the excel file will be open during this process
            xlApp.Visible = false;
            //open the excel file
            xlWorkBook = xlApp.Workbooks.Open(excelFilePath, 0,
                        false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows,
                         "", true, false, 0, true, false, false);

            //get all the worksheets in the excel  file
            xlBigSheet = xlWorkBook.Worksheets;
            string sheetName = "Extracted";
            //get the specified worksheet
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlBigSheet.get_Item(sheetName);

            xlSheetRange = xlWorkSheet.UsedRange;

            int colCount = xlSheetRange.Columns.Count;
            int rowCount = xlSheetRange.Rows.Count;

            //iterate the rows
            for (int rowIndex = 1; rowIndex <= rowCount; rowIndex++)
            {
                Microsoft.Office.Interop.Excel.Range cell = xlSheetRange.Cells[rowIndex, 2];
                if (cell.Value2 != null && !string.IsNullOrWhiteSpace(cell.Value2.ToString()) && !cell.Value2.ToString().Trim().Equals("User name"))
                {
                    Microsoft.Office.Interop.Excel.Range cellAktiv1 = xlSheetRange.Cells[rowIndex, 3];
                    Microsoft.Office.Interop.Excel.Range cellAktiv2 = xlSheetRange.Cells[rowIndex, 4];

                    if (cellAktiv1.Value2 != null && !string.IsNullOrWhiteSpace(cellAktiv1.Value2.ToString()) && cellAktiv1.Value2.ToString().Trim().Equals("Cross-checking"))
                    {

                        xlSheetRange.Cells[rowIndex, 3] = "Done";
                    }
                    if (cellAktiv2.Value2 != null && !string.IsNullOrWhiteSpace(cellAktiv2.Value2.ToString()) && cellAktiv2.Value2.ToString().Trim().Equals("Cross-checking"))
                    {
                        xlSheetRange.Cells[rowIndex, 4] = "Done";
                    }
                }
            }

            xlWorkBook.Save();

            #region test to print excel
            //takes all the cells even null valued cells
            //xlSheetRange = xlWorkSheet.get_Range("A1", "A" + xlWorkSheet.Rows.Count);
            //var values = (System.Array)xlSheetRange.Cells.Value2;

            //foreach (var n in values)
            //{
            //    if (n != null && !string.IsNullOrWhiteSpace(n.ToString()))
            //    {
            //        Console.WriteLine(n.ToString());
            //    }
            //}

            //working
            //for (int rowIndex = 1; rowIndex <= rowCount; rowIndex++)
            //{
            //    for (int colIndex = 1; colIndex <= colCount; colIndex++)
            //    {
            //        Microsoft.Office.Interop.Excel.Range cellPrint = xlSheetRange.Cells[rowIndex, colIndex];
            //        if (cellPrint.Value2 != null && !string.IsNullOrWhiteSpace(cellPrint.Value2.ToString()))
            //        {
            //            Console.WriteLine(cellPrint.Value2.ToString());
            //        }
            //    }
            //}
            #endregion

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

            #region OfficeOpenXml method
            //string filePath = @"D:\Sachith\PstTest\TestUsers.xlsx";

            //// Saves the file via a FileInfo
            //var file = new FileInfo(filePath);

            //// Creates the package and make sure you wrap it in a using statement
            //using (var package = new ExcelPackage(file))
            //{
            //    // Adds a new worksheet to the empty workbook
            //    //OfficeOpenXml.ExcelWorksheet worksheet = package.Workbook.Worksheets["Extracted"];
            //    OfficeOpenXml.Core.ExcelPackage.ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1"];


            //    // Starts to get data from database
            //    for (int row = 1; row < 10; row++)
            //    {
            //        // Writes data from sql database to excel's columns
            //        for (int col = 1; col < 10; col++)
            //        {
            //            worksheet.Cell(row, col).Value = Convert.ToString(row * col);
            //        }// Ends writing data from sql database to excel's columns

            //    }// Ends getting data from database


            //    // Saves new workbook and we are done!
            //    package.Save();
            //}
            #endregion

        }

    }
}
