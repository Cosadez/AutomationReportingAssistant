using System.Collections.Generic;

namespace AutomationReportingAssistant
{
    class ReportingReqs : IReportingReqs
    {
        string[,] currentDocUrlArrayNotFilled, currentDocUrlArrayIncludingFilled,  currentGroupedDocUrlArray, previousDocUrlArray, previousGroupedDocUrlArray;

        public static readonly string sourceSpreadsheetId = "put spreadsheet ID here"; // e.g. "18A-qjVd2aipy7pOBXqRTHNNriuc6USlnFVPVujALDVM"
        public static readonly string targetSpreadsheetId = sourceSpreadsheetId;

        public static readonly string sheetOfAllTests = "Failed"; // "Failed" or "NotExecuted"
        public static readonly string sheetOfTestsGroupedByException = "errors pivot list";
        public static readonly string sheetCopiedFromSource = "Previous report results";

        static readonly List<string> keysList = new (); // { "registration", "login", "deposit", "debit", "e2e" };

        public static readonly string environment = "qa"; // "qa" (by default) or "rc"

        public void FillGoogleReportWithTestResults()
        {
            ReportingMethods.PrepareGoogleSheetApiClient();

            //ReportingMethods.CopySheetFromOneSpreadSheetToOtherAndRename(sourceSpreadsheetId, targetSpreadsheetId);

            currentDocUrlArrayNotFilled = ReportingMethods.GetAllFailedTestsAsArray(targetSpreadsheetId, sheetOfAllTests, false);
            currentDocUrlArrayIncludingFilled = ReportingMethods.GetAllFailedTestsAsArray(targetSpreadsheetId, sheetOfAllTests, true);
            previousDocUrlArray = ReportingMethods.GetAllFailedTestsAsArray(targetSpreadsheetId, sheetOfAllTests, true); // return parameters sheetCopiedFromSource and false later

            currentGroupedDocUrlArray = ReportingMethods.GetGroupedFailedTestsAsArray(targetSpreadsheetId, sheetOfTestsGroupedByException);
            //previousGroupedDocUrlArray = ReportingMethods.GetGroupedFailedTestsAsArray(targetSpreadsheetId, sheetCopiedFromSource);

            //ReportingMethods.CompareCurrentAndPreviousFailedTestsAndFillResults(currentDocUrlArray, currentGroupedDocUrlArray, previousDocUrlArray, previousGroupedDocUrlArray);

            ReportingMethods.RunAllFailedTests(ReportingMethods.GetAllCommonFailedTestsFilteredByKeys(currentDocUrlArrayNotFilled, previousDocUrlArray, keysList));

            //ReportingMethods.FillPivotSheetWithResults(currentGroupedDocUrlArray, currentDocUrlArrayIncludingFilled);
        }
    }
}
