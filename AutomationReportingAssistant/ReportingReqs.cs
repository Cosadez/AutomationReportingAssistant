using System.Collections.Generic;

namespace AutomationReportingAssistant
{
    class ReportingReqs : IReportingReqs
    {
        public static readonly string sourceSpreadsheetId = "put spreadsheet ID here"; // e.g. "18A-qjVd2aipy7pOBXqRTHNNriuc6USlnFVPVujALDVM"
        public static readonly string targetSpreadsheetId = sourceSpreadsheetId;

        public static readonly string allTestsSheetName = "Failed"; // "Failed" or "NotExecuted"

        public static readonly string environment = "qa"; // "qa" (by default) or "rc"

        public static readonly List<string> keysFilterList = new (); // { "registration", "login", "deposit", "debit", "e2e" };

        public static readonly TestRunApproach testRunApproach = TestRunApproach.Parallel; // Parallel or AsyncAwait

        public static bool isRunIncludesTestsWithFilledResults = false; // set it to 'true' if you want to also run tests that already have results in table

        public static bool isReversedTestList = false; // set it to 'true' if you want to run tests from the bottom of list

        public void FillGoogleReportWithTestResults()
        {
            #region HiddenMethods

            ReportingMethods.PrepareGoogleSheetApiClient();

            //ReportingMethods.CopySheetFromOneSpreadSheetToOtherAndRename(sourceSpreadsheetId, targetSpreadsheetId);

            //previousGroupedDocUrlArray = ReportingMethods.GetGroupedFailedTestsAsArray(targetSpreadsheetId, sheetCopiedFromSource);

            //ReportingMethods.CompareCurrentAndPreviousFailedTestsAndFillResults(currentDocUrlArray, currentGroupedDocUrlArray, previousDocUrlArray, previousGroupedDocUrlArray);

            #endregion HiddenMethods

            ReportingMethods.RunAllFailedTests(testRunApproach);

            //ReportingMethods.FillPivotSheetWithResults();
        }
    }
}
