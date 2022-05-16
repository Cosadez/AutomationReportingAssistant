# AutomationReportingAssistant

Steps:

1. Load the project
2. Open **ReportingReqs.cs** file
3. Copy Google Sheet ID e.g. from this URL: docs.google.com/spreadsheets/d/**1VMU63GN7751EiOFtqW5o7iyUufcBAHhxhexc5s932NQ**/edit#gid=2019175412. You need only **1VMU63GN7751EiOFtqW5o7iyUufcBAHhxhexc5s932NQ** part
4. Paste it to ```public static readonly string sourceSpreadsheetId = "..."```
5. Select sheet name in ```public static readonly string allTestsSheetName =```, pick from ```"Failed"``` (default) or ```= "NotExecuted"``` sheets, where tests should be run
6. Set needed environment in ```public static readonly string environment =``` to ```"qa"``` (default) or ```"rc"```
7. Setup keyword filtering if needed or just leave it as empty List here (by default): ```public static readonly List<string> keysFilterList = new (); // { "registration", "login", "deposit", "debit", "e2e" };``` 
8. Select how you want to run tests (```public static readonly TestRunApproach testRunApproach =```):
    - ```TestRunApproach.Parallel``` (default, faster) if you just want to run tests and do not care about run order of tests
    - ```TestRunApproach.AsyncAwait``` (slower but sometimes fits best) if you want to run tests one by one and to control where to start test run: from the top or bottom of tests list
9. Set ```public static bool isRunIncludesTestsWithFilledResults =``` to ```true``` if you want to rerun tests that already have filled results in the table. Or just leave it as ```false``` (default)
10. If you want to run tests from the bottom of the list, select ```public static bool isReversedTestList = true```. If not, leave it as ```false``` (default). **NOTE** that in case of ```TestRunApproach.Parallel``` approach tests will also start from the bottom but not consequently
11. Running options:
   - If you want to run failed tests:
     - uncomment method ```ReportingMethods.RunAllFailedTests(testRunApproach);```  
     - comment method ```//ReportingMethods.FillPivotSheetWithResults();``` (commented by default)
   - If failed tests have been run and now you want to fill 'errors pivot list' tab with test results: 
     - comment method ```//ReportingMethods.RunAllFailedTests(testRunApproach);```
     - uncomment method ```ReportingMethods.FillPivotSheetWithResults();```
  - Or you can uncomment both methods and table will be filled with results after test run was successfully finished
