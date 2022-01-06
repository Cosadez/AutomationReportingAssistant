﻿using System.Collections.Generic;

namespace AutomationReportingAssistant
{
    class ReportingReqs : IReportingReqs
    {
        string[,] currentDocUrlArrayNotFilled, currentDocUrlArrayIncludingFilled,  currentGroupedDocUrlArray, previousDocUrlArray, previousGroupedDocUrlArray;

        public static readonly string sourceSpreadsheetId = "1VMU63GN7751EiOFtqW5o7iyUufcBAHhxhexc5s932NQ";
        public static readonly string targetSpreadsheetId = sourceSpreadsheetId;

        public static readonly string sheetOfAllTests = "Failed";
        public static readonly string sheetOfTestsGroupedByException = "errors pivot list";
        public static readonly string sheetCopiedFromSource = "Previous report results";

        static readonly List<string> keysList = new List<string>();// { "registration", "login", "deposit", "debit", "e2e" };

        public void FillGoogleReportWithTestResults()
        {
            ReportingMethods.CopySheetFromOneSpreadSheetToOtherAndRename(sourceSpreadsheetId, targetSpreadsheetId);

            currentDocUrlArrayNotFilled = ReportingMethods.GetAllFailedTestsAsArray(targetSpreadsheetId, sheetOfAllTests, false);
            currentDocUrlArrayIncludingFilled = ReportingMethods.GetAllFailedTestsAsArray(targetSpreadsheetId, sheetOfAllTests, true);
            previousDocUrlArray = ReportingMethods.GetAllFailedTestsAsArray(targetSpreadsheetId, sheetCopiedFromSource, false);

            currentGroupedDocUrlArray = ReportingMethods.GetGroupedFailedTestsAsArray(targetSpreadsheetId, sheetOfTestsGroupedByException);
            //previousGroupedDocUrlArray = ReportingMethods.GetGroupedFailedTestsAsArray(targetSpreadsheetId, sheetCopiedFromSource);

            //ReportingMethods.CompareCurrentAndPreviousFailedTestsAndFillResults(currentDocUrlArray, currentGroupedDocUrlArray, previousDocUrlArray, previousGroupedDocUrlArray);

            ReportingMethods.RunAllFailedTests(ReportingMethods.GetAllCommonFailedTestsFilteredByKeys(currentDocUrlArrayNotFilled, previousDocUrlArray, keysList));

            //ReportingMethods.FillPivotSheetWithResults(currentGroupedDocUrlArray, currentDocUrlArrayIncludingFilled);
        }
    }
}
