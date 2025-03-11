using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
using NLog;
using NLog.Config;
using NLog.Targets;

namespace ExcelLinkProcessor
{
    class Program
    {
        // Logger
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        // Configuration
        private static readonly string OverviewFilePath = @"S:\ITD\Kei\overview_file_template\董事會成員定額紀錄 2024.xlsx";
        private static readonly string EventFilesDirectory = @"S:\";

        // Constants
        private const string ProgramDonationSheet = "節目贊助";
        private const string CouponDonationSheet = "購券定額";
        private const string EventRecordSheet = "贊助記錄總表";
        private const string BoardMemberIdentifier = "董事會成員";
        private const string ProgramDonationIdentifier = "節目贊助金額";
        private const string CouponDonationIdentifier = "購券定額金額";

        static void Main(string[] args)
        {
            // Configure NLog
            ConfigureLogging();

            try
            {
                Logger.Info("Excel Link Processor Starting...");
                Console.WriteLine("Excel Link Processor Starting...");

                // List of specific event files to process
                List<string> eventFilesToProcess = new List<string>
                {
                    @"U:\IT staff\Kei\event_file_template\周年慈善晚會_24-25_贊助記錄表.xlsx",
                    @"U:\IT staff\Kei\event_file_template\慈善盆菜宴_2024_贊助紀錄表.xlsx"
                    // Add more event files as needed
                };

                // Validate all files before processing
                if (ValidateFiles(eventFilesToProcess))
                {
                    ProcessFiles(eventFilesToProcess);
                    Logger.Info("Processing completed successfully");
                    Console.WriteLine("Processing completed successfully. See log file for details.");
                }
                else
                {
                    Logger.Error("Validation failed");
                    Console.WriteLine("Validation failed. See log file for details.");
                }
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "Critical error");
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        private static void ConfigureLogging()
        {
            var config = new LoggingConfiguration();
            var fileTarget = new FileTarget
            {
                Name = "file",
                FileName = "processing_log.txt",
                Layout = "${longdate} - ${level:uppercase=true} - ${message}${onexception:${newline}${exception:format=tostring}}"
            };
            
            config.AddRule(LogLevel.Debug, LogLevel.Fatal, fileTarget);
            LogManager.Configuration = config;
        }

        private static bool ValidateFiles(List<string> eventFilePaths)
        {
            bool isValid = true;
            Logger.Info("Starting validation...");

            // Validate overview file
            if (!ValidateOverviewFile())
            {
                isValid = false;
            }

            // Validate each event file
            foreach (string eventFilePath in eventFilePaths)
            {
                if (!ValidateEventFile(eventFilePath))
                {
                    isValid = false;
                }
            }

            Logger.Info($"Validation completed. Result: {(isValid ? "Passed" : "Failed")}");
            return isValid;
        }

        private static bool ValidateOverviewFile()
        {
            Logger.Info($"Validating overview file: {OverviewFilePath}");

            if (!File.Exists(OverviewFilePath))
            {
                Logger.Error($"Overview file not found: {OverviewFilePath}");
                return false;
            }

            bool isValid = true;

            try
            {
                using (var fs = new FileStream(OverviewFilePath, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook workbook = new XSSFWorkbook(fs);

                    // Check for required sheets
                    ISheet programSheet = workbook.GetSheet(ProgramDonationSheet);
                    ISheet couponSheet = workbook.GetSheet(CouponDonationSheet);

                    if (programSheet == null)
                    {
                        Logger.Error($"Missing '{ProgramDonationSheet}' sheet in overview file");
                        isValid = false;
                    }

                    if (couponSheet == null)
                    {
                        Logger.Error($"Missing '{CouponDonationSheet}' sheet in overview file");
                        isValid = false;
                    }

                    // Validate program donation sheet
                    if (programSheet != null)
                    {
                        CellReference boardMemberCell = FindCellWithText(programSheet, BoardMemberIdentifier);
                        
                        if (boardMemberCell == null)
                        {
                            Logger.Error($"'{BoardMemberIdentifier}' cell not found in '{ProgramDonationSheet}' sheet");
                            isValid = false;
                        }
                        else
                        {
                            // Check if column F is available for events
                            if (boardMemberCell.Col > 5) // 0-based index, so 5 = column F
                            {
                                Logger.Error($"'{BoardMemberIdentifier}' cell is not positioned to allow events starting from column F");
                                isValid = false;
                            }
                            else
                            {
                                Logger.Info($"'{ProgramDonationSheet}' sheet: '{BoardMemberIdentifier}' found at {CellReference.ConvertNumToColString(boardMemberCell.Col)}{boardMemberCell.Row + 1}");
                            }
                        }
                    }

                    // Validate coupon donation sheet
                    if (couponSheet != null)
                    {
                        CellReference boardMemberCell = FindCellWithText(couponSheet, BoardMemberIdentifier);
                        
                        if (boardMemberCell == null)
                        {
                            Logger.Error($"'{BoardMemberIdentifier}' cell not found in '{CouponDonationSheet}' sheet");
                            isValid = false;
                        }
                        else
                        {
                            // Check if column G is available for events
                            if (boardMemberCell.Col > 6) // 0-based index, so 6 = column G
                            {
                                Logger.Error($"'{BoardMemberIdentifier}' cell is not positioned to allow events starting from column G");
                                isValid = false;
                            }
                            else
                            {
                                Logger.Info($"'{CouponDonationSheet}' sheet: '{BoardMemberIdentifier}' found at {CellReference.ConvertNumToColString(boardMemberCell.Col)}{boardMemberCell.Row + 1}");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error(ex, $"Error validating overview file");
                isValid = false;
            }

            return isValid;
        }

        private static bool ValidateEventFile(string eventFilePath)
        {
            Logger.Info($"Validating event file: {eventFilePath}");

            if (!File.Exists(eventFilePath))
            {
                Logger.Error($"Event file not found: {eventFilePath}");
                return false;
            }

            bool isValid = true;

            try
            {
                using (var fs = new FileStream(eventFilePath, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook workbook = new XSSFWorkbook(fs);

                    // Check for required sheet
                    ISheet recordSheet = workbook.GetSheet(EventRecordSheet);

                    if (recordSheet == null)
                    {
                        Logger.Error($"Missing '{EventRecordSheet}' sheet in event file: {eventFilePath}");
                        return false;
                    }

                    // Check for required board member cell
                    CellReference boardMemberCell = FindCellWithText(recordSheet, BoardMemberIdentifier);
                    
                    if (boardMemberCell == null)
                    {
                        Logger.Error($"'{BoardMemberIdentifier}' cell not found in event file: {eventFilePath}");
                        isValid = false;
                    }
                    else
                    {
                        Logger.Info($"Event file '{Path.GetFileName(eventFilePath)}': '{BoardMemberIdentifier}' found at {CellReference.ConvertNumToColString(boardMemberCell.Col)}{boardMemberCell.Row + 1}");
                    }

                    // Check for at least one of the donation identifier cells
                    CellReference programDonationCell = FindCellWithText(recordSheet, ProgramDonationIdentifier);
                    CellReference couponDonationCell = FindCellWithText(recordSheet, CouponDonationIdentifier);

                    if (programDonationCell == null && couponDonationCell == null)
                    {
                        Logger.Error($"Neither '{ProgramDonationIdentifier}' nor '{CouponDonationIdentifier}' cell found in event file: {eventFilePath}");
                        isValid = false;
                    }
                    else
                    {
                        // Log which donation identifiers were found
                        if (programDonationCell != null)
                        {
                            Logger.Info($"Event file '{Path.GetFileName(eventFilePath)}': '{ProgramDonationIdentifier}' found at {CellReference.ConvertNumToColString(programDonationCell.Col)}{programDonationCell.Row + 1}");
                        }
                        else
                        {
                            Logger.Warn($"Event file '{Path.GetFileName(eventFilePath)}': '{ProgramDonationIdentifier}' not found, but will continue processing");
                        }

                        if (couponDonationCell != null)
                        {
                            Logger.Info($"Event file '{Path.GetFileName(eventFilePath)}': '{CouponDonationIdentifier}' found at {CellReference.ConvertNumToColString(couponDonationCell.Col)}{couponDonationCell.Row + 1}");
                        }
                        else
                        {
                            Logger.Warn($"Event file '{Path.GetFileName(eventFilePath)}': '{CouponDonationIdentifier}' not found, but will continue processing");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error(ex, $"Error validating event file {eventFilePath}");
                isValid = false;
            }

            return isValid;
        }

        private static void ProcessFiles(List<string> eventFilePaths)
        {
            Logger.Info("Starting file processing...");
            
            try
            {
                // Load overview workbook
                IWorkbook overviewWorkbook;
                using (var fs = new FileStream(OverviewFilePath, FileMode.Open, FileAccess.Read))
                {
                    overviewWorkbook = new XSSFWorkbook(fs);
                }

                // Get overview sheets
                ISheet programSheet = overviewWorkbook.GetSheet(ProgramDonationSheet);
                ISheet couponSheet = overviewWorkbook.GetSheet(CouponDonationSheet);

                // Find board member cells in overview sheets
                CellReference programBoardMemberCell = FindCellWithText(programSheet, BoardMemberIdentifier);
                CellReference couponBoardMemberCell = FindCellWithText(couponSheet, BoardMemberIdentifier);

                // Get board members from overview file
                Dictionary<string, int> programBoardMembers = GetBoardMembers(programSheet, programBoardMemberCell);
                Dictionary<string, int> couponBoardMembers = GetBoardMembers(couponSheet, couponBoardMemberCell);

                // Get existing events in overview file
                Dictionary<string, int> programEvents = GetExistingEvents(programSheet, programBoardMemberCell, 5); // 0-based index, so 5 = column F
                Dictionary<string, int> couponEvents = GetExistingEvents(couponSheet, couponBoardMemberCell, 6);    // 0-based index, so 6 = column G

                // Process each event file
                foreach (string eventFilePath in eventFilePaths)
                {
                    ProcessEventFile(
                        eventFilePath,
                        overviewWorkbook,
                        programSheet,
                        couponSheet,
                        programBoardMemberCell,
                        couponBoardMemberCell,
                        programBoardMembers,
                        couponBoardMembers,
                        programEvents,
                        couponEvents
                    );
                }

                // Save the overview workbook
                using (var fs = new FileStream(OverviewFilePath, FileMode.Create, FileAccess.Write))
                {
                    overviewWorkbook.Write(fs);
                }
                
                Logger.Info("Overview file saved successfully");
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "Error during processing");
                throw;
            }
        }

        private static void ProcessEventFile(
            string eventFilePath,
            IWorkbook overviewWorkbook,
            ISheet programSheet,
            ISheet couponSheet,
            CellReference programBoardMemberCell,
            CellReference couponBoardMemberCell,
            Dictionary<string, int> programBoardMembers,
            Dictionary<string, int> couponBoardMembers,
            Dictionary<string, int> programEvents,
            Dictionary<string, int> couponEvents)
        {
            string eventName = Path.GetFileNameWithoutExtension(eventFilePath);
            Logger.Info($"Processing event file: {eventName}");

            try
            {
                // Load event workbook
                IWorkbook eventWorkbook;
                using (var fs = new FileStream(eventFilePath, FileMode.Open, FileAccess.Read))
                {
                    eventWorkbook = new XSSFWorkbook(fs);
                }

                ISheet eventSheet = eventWorkbook.GetSheet(EventRecordSheet);

                // Find key cells in event file
                CellReference eventBoardMemberCell = FindCellWithText(eventSheet, BoardMemberIdentifier);
                CellReference programDonationCell = FindCellWithText(eventSheet, ProgramDonationIdentifier);
                CellReference couponDonationCell = FindCellWithText(eventSheet, CouponDonationIdentifier);

                // Get board members from event file
                Dictionary<string, int> eventBoardMembers = GetBoardMembers(eventSheet, eventBoardMemberCell);

                // Process program donations if the identifier exists
                if (programDonationCell != null)
                {
                    ProcessDonationType(
                        overviewWorkbook,
                        programSheet,
                        eventSheet,
                        eventFilePath,
                        eventName,
                        programBoardMemberCell,
                        eventBoardMemberCell,
                        programDonationCell,
                        programBoardMembers,
                        eventBoardMembers,
                        programEvents,
                        4, // 0-based index, so 5 = column F
                        "Program Donation"
                    );
                }
                else
                {
                    Logger.Info($"Skipping Program Donation processing for event '{eventName}' as the identifier was not found");
                }

                // Process coupon donations if the identifier exists
                if (couponDonationCell != null)
                {
                    ProcessDonationType(
                        overviewWorkbook,
                        couponSheet,
                        eventSheet,
                        eventFilePath,
                        eventName,
                        couponBoardMemberCell,
                        eventBoardMemberCell,
                        couponDonationCell,
                        couponBoardMembers,
                        eventBoardMembers,
                        couponEvents,
                        5, // 0-based index, so 6 = column G
                        "Coupon Donation"
                    );
                }
                else
                {
                    Logger.Info($"Skipping Coupon Donation processing for event '{eventName}' as the identifier was not found");
                }
            }
            catch (Exception ex)
            {
                Logger.Error(ex, $"Error processing event file {eventName}");
                throw;
            }
        }

        private static void ProcessDonationType(
            IWorkbook overviewWorkbook,
            ISheet overviewSheet,
            ISheet eventSheet,
            string eventFilePath,
            string eventName,
            CellReference overviewBoardMemberCell,
            CellReference eventBoardMemberCell,
            CellReference donationTypeCell,
            Dictionary<string, int> overviewBoardMembers,
            Dictionary<string, int> eventBoardMembers,
            Dictionary<string, int> existingEvents,
            int startColumn,
            string donationType)
        {
            int eventColumn;

            // Check if event exists in overview
            if (existingEvents.ContainsKey(eventName))
            {
                eventColumn = existingEvents[eventName];
                Logger.Info($"{donationType}: Event '{eventName}' found at column {CellReference.ConvertNumToColString(eventColumn)}");
            }
            else
            {
                // Find the rightmost event column
                int lastColumn = startColumn;
                foreach (int col in existingEvents.Values)
                {
                    if (col > lastColumn)
                        lastColumn = col;
                }

                // Add new event column
                eventColumn = lastColumn + 1;
                
                // Add event name to header
                IRow headerRow = overviewSheet.GetRow(overviewBoardMemberCell.Row);
                if (headerRow == null)
                {
                    headerRow = overviewSheet.CreateRow(overviewBoardMemberCell.Row);
                }
                
                ICell headerCell = headerRow.CreateCell(eventColumn);
                headerCell.SetCellValue(eventName);
                
                // Copy cell style from previous column if possible
                if (headerRow.GetCell(lastColumn) != null)
                {
                    ICellStyle sourceStyle = headerRow.GetCell(lastColumn).CellStyle;
                    headerCell.CellStyle = sourceStyle;
                }
                
                existingEvents.Add(eventName, eventColumn);
                Logger.Info($"{donationType}: Added new event '{eventName}' at column {CellReference.ConvertNumToColString(eventColumn)}");
            }

            // Create formula links for each board member
            foreach (var overviewMember in overviewBoardMembers)
            {
                string memberId = GetBoardMemberId(overviewMember.Key);
                
                // Find matching board member in event file
                string matchingEventMember = null;
                foreach (var eventMember in eventBoardMembers)
                {
                    if (GetBoardMemberId(eventMember.Key) == memberId)
                    {
                        matchingEventMember = eventMember.Key;
                        break;
                    }
                }

                if (matchingEventMember != null)
                {
                    // Create formula link
                    int eventMemberRow = eventBoardMembers[matchingEventMember];
                    int overviewMemberRow = overviewMember.Value;
                    
                    // Find donation amount cell in event file
                    int donationColumn = donationTypeCell.Col;
                    
                    // Create external reference formula
                    // Note: NPOI doesn't directly support external references, so we create a string formula
                    string externalRef = $"'{Path.GetDirectoryName(eventFilePath)}\\[{Path.GetFileName(eventFilePath)}]{EventRecordSheet}'!{CellReference.ConvertNumToColString(donationColumn)}{eventMemberRow + 1}";
                    
                    IRow row = overviewSheet.GetRow(overviewMemberRow);
                    if (row == null)
                    {
                        row = overviewSheet.CreateRow(overviewMemberRow);
                    }
                    
                    ICell cell = row.CreateCell(eventColumn);
                    cell.SetCellFormula(externalRef);
                    
                    Logger.Info($"{donationType}: Created link for member {memberId} from {CellReference.ConvertNumToColString(donationColumn)}{eventMemberRow + 1} to {CellReference.ConvertNumToColString(eventColumn)}{overviewMemberRow + 1}");
                }
                else
                {
                    Logger.Warn($"{donationType}: Board member with ID {memberId} not found in event file {eventName}");
                }
            }
        }

        private static Dictionary<string, int> GetBoardMembers(ISheet sheet, CellReference boardMemberCell)
        {
            Dictionary<string, int> boardMembers = new Dictionary<string, int>();
            int expectedMemberCount = 20; // Expected number of board members
            int foundMemberCount = 0;
            
            int row = boardMemberCell.Row + 1;
            int maxRowsToCheck = row + 20; // Check up to 20 rows to find board members
            
            while (row < maxRowsToCheck)
            {
                IRow currentRow = sheet.GetRow(row);
                if (currentRow == null)
                    break;
                
                ICell cell = currentRow.GetCell(boardMemberCell.Col);
                if (cell == null)
                {
                    row++;
                    continue;
                }
                
                string cellValue = cell.ToString().Trim();
                if (string.IsNullOrEmpty(cellValue))
                {
                    row++;
                    continue;
                }
                
                // Check if the cell value matches the expected format: number followed by a dot
                if (Regex.IsMatch(cellValue, @"^\d+\."))
                {
                    boardMembers.Add(cellValue, row);
                    foundMemberCount++;
                    Logger.Debug($"Found board member: {cellValue} at row {row + 1}");
                }
                else
                {
                    // If we've already found some members but this one doesn't match the pattern,
                    // it might indicate we've reached the end of the member list
                    if (foundMemberCount > 0)
                    {
                        Logger.Debug($"Possible end of board member list at row {row + 1}: '{cellValue}'");
                        // Don't break immediately, as there might be valid members after this one
                    }
                }
                
                row++;
            }
            
            // Log warning if we found significantly fewer or more members than expected
            if (foundMemberCount < expectedMemberCount - 5)
            {
                Logger.Warn($"Found only {foundMemberCount} board members, expected around {expectedMemberCount}");
            }
            else if (foundMemberCount > expectedMemberCount + 5)
            {
                Logger.Warn($"Found {foundMemberCount} board members, which is more than the expected {expectedMemberCount}");
            }
            else
            {
                Logger.Info($"Found {foundMemberCount} board members");
            }
            
            return boardMembers;
        }

        private static Dictionary<string, int> GetExistingEvents(ISheet sheet, CellReference boardMemberCell, int startColumn)
        {
            Dictionary<string, int> events = new Dictionary<string, int>();
            
            IRow headerRow = sheet.GetRow(boardMemberCell.Row);
            if (headerRow == null)
                return events;
            
            int column = startColumn;
            while (true)
            {
                ICell cell = headerRow.GetCell(column);
                if (cell == null || string.IsNullOrEmpty(cell.ToString()))
                    break;
                
                string eventName = cell.ToString();
                events.Add(eventName, column);
                column++;
            }
            
            return events;
        }

        private static CellReference FindCellWithText(ISheet sheet, string text)
        {
            for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                IRow row = sheet.GetRow(rowIndex);
                if (row == null) continue;
                
                for (int colIndex = 0; colIndex < row.LastCellNum; colIndex++)
                {
                    ICell cell = row.GetCell(colIndex);
                    if (cell != null && cell.ToString().Replace("\r\n", string.Empty).Replace("\n", string.Empty) == text)
                    {
                        return new CellReference(rowIndex, colIndex);
                    }
                }
            }
            
            return null;
        }

        private static string GetBoardMemberId(string memberName)
        {
            // Extract the identifier number before the first dot
            Match match = Regex.Match(memberName, @"^(\d+)\.");
            return match.Success ? match.Groups[1].Value : memberName;
        }
    }
}
