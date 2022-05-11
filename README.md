# AutomationReportingAssistant

Steps:

1. Load the project
2. Open **ReportingReqs.cs** file
3. Copy Google Sheet ID e.g. docs.google.com/spreadsheets/d/**1VMU63GN7751EiOFtqW5o7iyUufcBAHhxhexc5s932NQ**/edit#gid=2019175412. You need only **1VMU63GN7751EiOFtqW5o7iyUufcBAHhxhexc5s932NQ** part
4. Paste it to ```public static readonly string sourceSpreadsheetId = "..."```
5. Select sheet name ```public static readonly string sheetOfAllTests = "Failed"``` or ```= "NotExecuted"```, where tests should be run
6. Setup keyword filtering if needed or just leave it as empty List here: ```static readonly List<string> keysList = new (); // { "registration", "login", "deposit", "debit", "e2e" };``` 
7. Make sure correct environment is selected in ```public static readonly string environment = "..."```
8. (Temporary solution)
   - If you want to run failed tests:
     - comment method ```ReportingMethods.FillPivotSheetWithResults(currentGroupedDocUrlArray, currentDocUrlArrayIncludingFilled);```
     - uncomment method ```ReportingMethods.RunAllFailedTests(ReportingMethods.GetAllCommonFailedTestsFilteredByKeys(currentDocUrlArrayNotFilled, previousDocUrlArray, keysList));```  
   - If failed tests have been run and now you want to fill 'errors pivot' tab with test results: 
     - comment method ```ReportingMethods.RunAllFailedTests(ReportingMethods.GetAllCommonFailedTestsFilteredByKeys(currentDocUrlArrayNotFilled, previousDocUrlArray, keysList));```
     - uncomment method ```ReportingMethods.FillPivotSheetWithResults(currentGroupedDocUrlArray, currentDocUrlArrayIncludingFilled);```

**NOTE:** In case you want to rerun tests that have already given results, use currentDocUrlArrayIncludingFilled in ```ReportingMethods.RunAllFailedTests(ReportingMethods.GetAllCommonFailedTestsFilteredByKeys(currentDocUrlArrayIncludingFilled, previousDocUrlArray, keysList));```

If not, use currentDocUrlArrayNotFilled in ```ReportingMethods.RunAllFailedTests(ReportingMethods.GetAllCommonFailedTestsFilteredByKeys(currentDocUrlArrayNotFilled, previousDocUrlArray, keysList));```
