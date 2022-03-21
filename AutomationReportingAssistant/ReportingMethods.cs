using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using HtmlAgilityPack;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using RestSharp;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;

namespace AutomationReportingAssistant
{
    class ReportingMethods
    {
        static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };

        static readonly string ApplicationName = "AutomationReportingAssistant";

        static SheetsService service;

        internal static string[,] GetAllFailedTestsAsArray(string spreadsheetId, string sheetId, bool includeFilledResults)
        {
            var failedTestsArray = new string[1000, 3];

            var envColumn = "J"; // QA
            //var envColumn = "I"; // RC

            var range = $"A2:{envColumn}1000";

            var valuesResult = ReadEntries(spreadsheetId, sheetId, range);

            if (valuesResult != null && valuesResult.Count > 0)
            {
                // if report = RC
                //for (int i = 0; i < valuesResult.Count; i++)
                //{
                //    if (valuesResult[i].Count <= 8)
                //    {
                //        failedTestsArray[i, 0] = (string)valuesResult[i][0];
                //        failedTestsArray[i, 1] = (string)valuesResult[i][1];
                //        failedTestsArray[i, 2] = (string)valuesResult[i][7];
                //    }
                //}

                // if report = QA

                if (includeFilledResults)
                {
                    for (int i = 0; i < valuesResult.Count; i++)
                    {

                        failedTestsArray[i, 0] = (string)valuesResult[i][0]; // test name
                        failedTestsArray[i, 1] = (string)valuesResult[i][1]; // test parameters
                        failedTestsArray[i, 2] = (string)valuesResult[i][8]; // exception message
                    }
                }
                else
                {
                    for (int i = 0; i < valuesResult.Count; i++)
                    {
                        if (valuesResult[i].Count <= 9)
                        {
                            failedTestsArray[i, 0] = (string)valuesResult[i][0]; // test name
                            failedTestsArray[i, 1] = (string)valuesResult[i][1]; // test parameters
                            failedTestsArray[i, 2] = (string)valuesResult[i][8]; // exception message
                        }
                    }
                }
            }

            return failedTestsArray;
        }

        internal static string[,] GetGroupedFailedTestsAsArray(string spreadsheetId, string sheetId)
        {
            var failedTestsGroupedArray = new string[1000, 3];

            var valuesGroupedResult = ReadEntries(spreadsheetId, sheetId, "A2:D1000");

            if (valuesGroupedResult != null && valuesGroupedResult.Count > 0)
            {
                for (int i = 0; i < valuesGroupedResult.Count; i++)
                {
                    if (valuesGroupedResult[i].Count != 0)
                    {
                        if (!valuesGroupedResult[i][1].Equals(string.Empty))
                        {
                            failedTestsGroupedArray[i, 0] = (string)valuesGroupedResult[i][0]; // exception message
                            failedTestsGroupedArray[i, 1] = (string)valuesGroupedResult[i][1]; // test name
                            failedTestsGroupedArray[i, 2] = (string)valuesGroupedResult[i][2]; // count of several hit for same test
                        }
                    }
                }
            }

            return failedTestsGroupedArray;
        }

        static IList<IList<object>> ReadEntries(string spreadsheetId, string sheetName, string range)
        {
            var rangeForRequest = $"{sheetName}!{range}";
            var request = service.Spreadsheets.Values.Get(spreadsheetId, rangeForRequest);
            var response = request.Execute();
            var values = response.Values;

            return values;
        }

        static void CreateEntry(string spreadsheetId) // creates new row at the bottom of existing ones
        {
            var range = "";// $"{sheetOfTestsGroupedByException}!E:E";
            var valueRange = new ValueRange();

            var objectList = new List<object>() { "Hello", "This" };
            valueRange.Values = new List<IList<object>> { objectList };

            var appendRequest = service.Spreadsheets.Values.Append(valueRange, spreadsheetId, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var appendResponse = appendRequest.Execute();
        }

        static void UpdateEntry(string spreadsheetId, string sheetName, string rowId, string value)
        {
            var envColumn = "J"; // QA
            //var envColumn = "I"; // RC

            var range = $"{sheetName}!{envColumn}{rowId}";
            var valueRange = new ValueRange();

            var objectList = new List<object>() { value };
            valueRange.Values = new List<IList<object>> { objectList };

            var updateRequest = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            var updateResponse = updateRequest.Execute();
        }

        static void DeleteEntry(string spreadsheetId)
        {
            var range = ""; //$"{sheetOfTestsGroupedByException}!E2";
            var requestBody = new ClearValuesRequest();

            var deleteRequest = service.Spreadsheets.Values.Clear(requestBody, spreadsheetId, range);
            var deleteResponse = deleteRequest.Execute();
        }

        public static void CopySheetFromOneSpreadSheetToOtherAndRename(string sSpreadsheetId, string tSpreadsheetId)
        {
            GoogleCredential credential;

            using (var stream = new FileStream("client_secretsAuto.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream).CreateScoped(Scopes);
            }

            service = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName
            });

            var sSpreadSheet = service.Spreadsheets.Get(sSpreadsheetId).Execute();
            var sourceSheetId = sSpreadSheet.Sheets.Where(s => s.Properties.Title == ReportingReqs.sheetOfAllTests).FirstOrDefault().Properties.SheetId;

            var isExistingPreviousSheet = sSpreadSheet.Sheets.Where(s => s.Properties.Title == ReportingReqs.sheetCopiedFromSource).FirstOrDefault() != null;

            if (!isExistingPreviousSheet)
            {
                var requestForCopy = new CopySheetToAnotherSpreadsheetRequest()
                {
                    DestinationSpreadsheetId = ReportingReqs.targetSpreadsheetId
                };

                service.Spreadsheets.Sheets.CopyTo(requestForCopy, ReportingReqs.sourceSpreadsheetId, (int)sourceSheetId).Execute();

                RenameSheet(tSpreadsheetId);
            }
        }

        public static void RenameSheet(string tSpreadsheetId)
        {
            var tSpreadSheet = service.Spreadsheets.Get(tSpreadsheetId).Execute();
            var targetSheetId = tSpreadSheet.Sheets.Where(s => s.Properties.Title.Contains($"{ReportingReqs.sheetOfAllTests} (")).FirstOrDefault().Properties.SheetId;

            BatchUpdateSpreadsheetRequest bussr = new BatchUpdateSpreadsheetRequest();

            var request = new Request()
            {
                UpdateSheetProperties = new UpdateSheetPropertiesRequest
                {
                    Properties = new SheetProperties()
                    {
                        Title = ReportingReqs.sheetCopiedFromSource,
                        SheetId = targetSheetId
                    },
                    Fields = "Title"
                }
            };

            bussr.Requests = new List<Request>();
            bussr.Requests.Add(request);
            service.Spreadsheets.BatchUpdate(bussr, tSpreadsheetId).Execute();
        }

        static void CopyEntry(string tSpreadsheetId, int sourceRowIndex, int targetRowIndex)
        {
            System.Threading.Thread.Sleep(1000);
            var tSpreadSheet = service.Spreadsheets.Get(tSpreadsheetId).Execute();
            var sourceSheetId = tSpreadSheet.Sheets.Where(s => s.Properties.Title == ReportingReqs.sheetOfAllTests).FirstOrDefault().Properties.SheetId;
            var targetSheetId = tSpreadSheet.Sheets.Where(s => s.Properties.Title == ReportingReqs.sheetOfTestsGroupedByException).FirstOrDefault().Properties.SheetId;

            var sourceColumnId = 9; // qa
            //var sourceColumnId = 8; // rc

            var targetColumnId = 3;

            var copyReq = new CopyPasteRequest()
            {
                Source = new GridRange()
                {
                    SheetId = sourceSheetId,
                    StartRowIndex = sourceRowIndex,
                    EndRowIndex = sourceRowIndex + 1, // = (StartRowIndex + 1) if 1 cell is copied
                    StartColumnIndex = sourceColumnId, // always fixed
                    EndColumnIndex = sourceColumnId + 1, // always fixed
                },
                Destination = new GridRange()
                {
                    SheetId = targetSheetId,
                    StartRowIndex = targetRowIndex,
                    EndRowIndex = targetRowIndex + 1, // = (StartRowIndex + 1) if 1 cell is copied
                    StartColumnIndex = targetColumnId, // always fixed
                    EndColumnIndex = targetColumnId + 1, // always fixed
                },
                PasteType = "PASTE_NORMAL",
                PasteOrientation = "NORMAL"
            };

            var copyResource = new BatchUpdateSpreadsheetRequest() { Requests = new List<Request>() };
            var reqCopy = new Request() { CopyPaste = copyReq };
            copyResource.Requests.Add(reqCopy);

            service.Spreadsheets.BatchUpdate(copyResource, tSpreadsheetId).Execute();
        }

        internal static void CompareCurrentAndPreviousFailedTestsAndFillResults(string[,] currentDocUrlArray, string[,] currentGroupedDocUrlArray, string[,] previousDocUrlArray, string[,] previousGroupedDocUrlArray)
        {
            for (int i = 0; i < currentDocUrlArray.Length / 3 - 1; i++)
            {
                for (int j = 0; j < previousDocUrlArray.Length / 3 - 1; j++)
                {
                    if (currentDocUrlArray[i, 0] == previousDocUrlArray[j, 0] && currentDocUrlArray[i, 1] == previousDocUrlArray[j, 1])
                    {
                        for (int k = 0; k < currentGroupedDocUrlArray.Length / 3 - 1; k++)
                        {
                            if (currentDocUrlArray[i, 0] == currentGroupedDocUrlArray[k, 1] && currentGroupedDocUrlArray[k, 1] != null)
                            {
                                for (int l = 0; l < previousGroupedDocUrlArray.Length / 3 - 1; l++)
                                {
                                    if (currentGroupedDocUrlArray[k, 1] == previousGroupedDocUrlArray[l, 1] && previousGroupedDocUrlArray[l, 1] != null)
                                    {
                                        int iz = 0;
                                        CopyEntry(ReportingReqs.targetSpreadsheetId, l + 1, k + 1);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        internal static void RunAllFailedTests(List<string> failedTestsList)
        {
            var processedTestsCounter = 0;
            var failedTestsListCounter = failedTestsList.Count;

            foreach (var failedTest in failedTestsList)//.Take(100))
            {
                var testResult = RunFailedTest(failedTest.Split(":::").FirstOrDefault(), failedTest.Split(":::")[1]);

                if (testResult.Contains("Failed"))
                {
                    var ex1 = testResult.Split(":::")[1].Replace("\r", "").Replace("\n", "").Replace("<string.Empty>", "").Replace("<br/>", "").Replace("<Automation.Api.DataAccessLayer.PlayersLimitsModels.PlayerLimits>", "").Replace("<BtoBet.Gateway.ServiceLayer.Messages.Common.Error>", "").Replace("<empty>", "");

                    var ex2 = failedTest.Split(":::")[3].Replace("\r", "").Replace("\n", "").Replace("<string.Empty>", "").Replace("<br/>", "").Replace("<Automation.Api.DataAccessLayer.PlayersLimitsModels.PlayerLimits>", "").Replace("<BtoBet.Gateway.ServiceLayer.Messages.Common.Error>", "").Replace("<empty>", "");

                    if (ex1 == ex2)
                    {
                        UpdateEntry(ReportingReqs.targetSpreadsheetId, ReportingReqs.sheetOfAllTests, failedTest.Split(":::")[2], $"{testResult.Split(":::")[0]}; Same exception");
                    }
                    else
                    {
                        UpdateEntry(ReportingReqs.targetSpreadsheetId, ReportingReqs.sheetOfAllTests, failedTest.Split(":::")[2], $"{testResult.Split(":::")[0]}; New exception: {testResult.Split(":::")[1]}");
                    }

                    processedTestsCounter++;
                }
                else if (testResult.Contains("Passed") || testResult.Contains("Skipped") || testResult.Contains("Inconclusive"))
                {
                    UpdateEntry(ReportingReqs.targetSpreadsheetId, ReportingReqs.sheetOfAllTests, failedTest.Split(":::")[2], testResult);

                    processedTestsCounter++;
                }
            }

            Console.WriteLine($"Processed tests number: {processedTestsCounter} from total {failedTestsListCounter} tests");
        }

        internal static void FillPivotSheetWithResults(string[,] groupedDocUrlArray, string[,] currentDocUrlArrayIncludingFilled)
        {
            for (int i = 0; i < groupedDocUrlArray.Length / 3; i++)
            {
                var exceptionGrouped = "";
                
                for (int k = i; k >= 0; k--)
                {
                    if (groupedDocUrlArray[k, 0] == null || groupedDocUrlArray[k, 0] == string.Empty)
                    {
                        continue;
                    }
                    else
                    {
                        exceptionGrouped = groupedDocUrlArray[k, 0];
                        break;
                    }
                }

                for (int j = 0; j < currentDocUrlArrayIncludingFilled.Length / 3; j++)
                {
                    if (groupedDocUrlArray[i, 1] != null && groupedDocUrlArray[i, 2] != "" && groupedDocUrlArray[i, 1] == currentDocUrlArrayIncludingFilled[j, 0] && exceptionGrouped == currentDocUrlArrayIncludingFilled[j, 2])
                    {
                        CopyEntry(ReportingReqs.targetSpreadsheetId, j + 1, i + 1);
                    }
                }
            }
        }

        internal static List<string> GetAllCommonFailedTests(string[,] currentDocUrlArray, string[,] previousDocUrlArray)
        {
            var failedTestsList = new List<string>();

            for (int i = 0; i < currentDocUrlArray.Length / 3; i++)
            {
                for (int j = 0; j < previousDocUrlArray.Length / 3; j++)
                {
                    if (currentDocUrlArray[i, 0] == previousDocUrlArray[j, 0] && currentDocUrlArray[i, 1] == previousDocUrlArray[j, 1])
                    {
                        var failedTestTemp = $"{currentDocUrlArray[i, 0]}:::{currentDocUrlArray[i, 1]}:::{i + 2}:::{currentDocUrlArray[i, 2]}";

                        if (!failedTestTemp.StartsWith(":::"))
                        {
                            failedTestsList.Add(failedTestTemp);
                        }
                    }
                }
            }

            return failedTestsList;
        }

        internal static List<string> GetAllCommonFailedTestsFilteredByKeys(string[,] currentDocUrlArray, string[,] previousDocUrlArray, List<string> keysList)
        {
            var allCommonFailedTestsList = GetAllCommonFailedTests(currentDocUrlArray, previousDocUrlArray);
            var allCommonFailedTestsListFilteredByKeys = new List<string>();

            if (keysList.Count == 0)
            {
                foreach (var allCommonFailedTest in allCommonFailedTestsList)
                {
                    if (!allCommonFailedTest.StartsWith("::::::"))
                    {
                        allCommonFailedTestsListFilteredByKeys.Add(allCommonFailedTest);
                    }
                }
            }
            else
            {
                foreach (var allCommonFailedTest in allCommonFailedTestsList)
                {
                    foreach (var key in keysList)
                    {
                        if (allCommonFailedTest.Split(":::").FirstOrDefault().ToLower().Contains(key.ToLower()))
                        {
                            allCommonFailedTestsListFilteredByKeys.Add(allCommonFailedTest);
                        }
                    }
                }
            }

            return allCommonFailedTestsListFilteredByKeys;
        }

        internal static string RunFailedTest(string testName, string parameterName)
        {
            var parameterTestValue = "";
            var response = "";
            var totalCount = "";
            var passedCount = "";
            var failedCount = "";
            var inconclusiveCount = "";
            var skippedCount = "";

            var retryCount = 3;

            var env = "qa"; // or "rc"

            var requestUrl = $"http://specflow.{env}.gamesrv1.com/Home/RunTests";

            var request = new RestRequest(requestUrl, Method.Post);

            for (int i = 0; i < retryCount - 1; i++)
            {
                parameterTestValue = $"tests[]={GetTestFullName(testName)}{parameterName}&connectionId=9f0ea001-646c-44ba-a6d3-8272d7f7f449&testParameters=";

                if (parameterTestValue != string.Empty)
                {
                    break;
                }
            }

            if (parameterTestValue == string.Empty)
            {
                return $"GetTestFullName HTTP response was null after {retryCount} tries in test {testName}{parameterName}";
            }

            request.AddHeader("Content-Length", $"{parameterTestValue.Length}");
            request.AddHeader("Content-Type", "application/x-www-form-urlencoded; charset=UTF-8");
            request.AddParameter("", parameterTestValue, ParameterType.RequestBody);

            for (int i = 0; i < retryCount - 1; i++)
            {
                response = new RestClient(requestUrl).ExecuteAsync(request).Result.Content;

                if (response != null)
                {
                    break;
                }
            }

            if (response == null)
            {
                var retryFailMessage = $"RunFaiIedTest HTTP response was null after {retryCount} tries";

                Console.WriteLine($"{retryFailMessage} in test {testName}{parameterName}");
                return retryFailMessage;
            }

            var source = WebUtility.HtmlDecode(response);
            var html = new HtmlDocument();
            html.LoadHtml(source);

            var root = html.DocumentNode;

            try
            {
                var resultList = root.Descendants("h6").Where(t => t.InnerHtml.Contains("Passed")).Select(t => t.InnerText).ToList().FirstOrDefault();

                if (resultList != null)
                {
                    var resultArray = resultList.ToString().Split(';');

                    var exceptionMessage = root.Descendants("span");

                    var exceptionMessageParsed = "";
                    foreach (var e in exceptionMessage.Skip(1))
                    {
                        exceptionMessageParsed += e.InnerText;
                    }

                    totalCount = resultArray[0].Split(':')[1].Trim();
                    passedCount = resultArray[1].Split(':')[1].Trim();
                    failedCount = resultArray[2].Split(':')[1].Trim();
                    inconclusiveCount = resultArray[3].Split(':')[1].Trim();
                    skippedCount = resultArray[4].Split(':')[1].Trim().Split("   ").FirstOrDefault();

                    if (totalCount == "1")
                    {
                        var returnValue = "";

                        if (passedCount == "1")
                        {
                            returnValue = "Passed";
                        }
                        else if (failedCount == "1")
                        {
                            returnValue = $"Failed:::{exceptionMessageParsed}";
                        }
                        else if (inconclusiveCount == "1")
                        {
                            returnValue = "Inconclusive";
                            Console.WriteLine($"{returnValue}: {testName}{parameterName}");
                        }
                        else if (skippedCount == "1")
                        {
                            returnValue = "Skipped";
                            Console.WriteLine($"{returnValue}: {testName}{parameterName}");
                        }
                        return returnValue;
                    }
                    else
                    {
                        Console.WriteLine($"Total count is not 1 in test {testName}{parameterName}");
                        return $"Total count is not 1 in test {testName}{parameterName}";
                    }
                }
                else
                {
                    Console.WriteLine($"Result list is null in test {testName}{parameterName}");
                    return $"Result list is null in test {testName}{parameterName}";
                }
            }
            catch (NullReferenceException)
            {
                Console.WriteLine($"Unexpected HTML response in test {testName}{parameterName}");
                return $"Unexpected HTML response in test {testName}{parameterName}";
            }
        }

        internal static string GetTestFullName(string testName)
        {
            var fullTestName = "";

            var env = "qa"; // or "rc"

            var requestUrl = $"http://specflow.{env}.gamesrv1.com/Home/GetTree?search={testName.Replace("\\", "")}";
            var response = new RestClient(requestUrl).ExecuteGetAsync(new RestRequest(requestUrl, Method.Get));

            try
            {
                var json = (JObject)JsonConvert.DeserializeObject(response.Result.Content);

                var pathList = json.DescendantsAndSelf().OfType<JProperty>().Where(p => p.Name == "fullName").Select(v => v.Value).ToList();

                foreach (var item in pathList)
                {
                    if (item.ToString().EndsWith(testName))
                    {
                        fullTestName = item.ToString();
                    }
                }
            }
            catch (ArgumentNullException e)
            {
                Console.WriteLine("VPN is probably disabled");
            }
            catch (JsonReaderException)
            {
                if (response.Result.Content.Contains("An error occurred while processing your request"))
                {
                    Console.WriteLine($"HTML response is returned with error instead of JSON for test {testName}");
                }
                else
                {
                    Console.WriteLine($"Could not parse JSON response for test {testName}");
                }
            }

            return fullTestName;
        }
    }
}
