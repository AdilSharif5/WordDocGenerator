using Newtonsoft.Json;
using Microsoft.Office.Interop;
using Newtonsoft.Json.Linq;

namespace WordDocGenerator
{
    class Program
    {
        static void Main()
        {
            /** Not working full JSON */
            string json = @"{
    ""rows"": [
        {
            ""cols"": [
                {
                    ""font"": {},
                    ""bgColor"": """",
                    ""color"": """",
                    ""colSpan"": 3,
                    ""rowSpan"": ""1"",
                    ""textAlign"": """",
                    ""cellContent"": [
                        {
                            ""component"": ""Companion Policy Information"",
                            ""cellType"": ""StaticText"",
							""label"": ""Companion Policy Information"",
                        }
                    ]
                }
            ]
        },
        {
            ""cols"": [
                {
                    ""font"": {},
                    ""bgColor"": """",
                    ""color"": """",
                    ""colSpan"": ""1"",
                    ""rowSpan"": ""1"",
                    ""textAlign"": """",
                    ""cellContent"": [
                        {
                            ""label"": ""<p>There are many variations of passages of Lorem Ipsum available, but the majority have suffered alteration in some form, by injected humour, or randomised words which don't look even slightly believable. If you are going to use a passage of Lorem Ipsum, you need to be sure there isn't anything embarrassing hidden in the middle of text. All the Lorem Ipsum generators on the Internet tend to repeat predefined chunks as necessary, making this the first true generator on the Internet. It uses a dictionary of over 200 Latin words, combined with a handful of model sentence structures, to generate Lorem Ipsum which looks reasonable. The generated Lorem Ipsum is therefore always free from repetition, injected humour, or non-characteristic words etc.</p>"",
                            ""display"": ""block"",
                            ""cellType"": ""StaticText""
							
                        }
                    ]
                },
                {
                    ""font"": {},
                    ""bgColor"": """",
                    ""color"": """",
                    ""cellContent"": [],
                    ""colSpan"": ""1"",
                    ""rowSpan"": ""1"",
                    ""textAlign"": """"
                },
                {
                    ""font"": {},
                    ""bgColor"": """",
                    ""color"": """",
                    ""cellContent"": [
                        {
                            ""label"": ""<p>There are many variations of passages of Lorem Ipsum available, but the majority have suffered alteration in some form, by injected humour, or randomised words which don't look even slightly believable. If you are going to use a passage of Lorem Ipsum, you need to be sure there isn't anything embarrassing hidden in the middle of text. All the Lorem Ipsum generators on the Internet tend to repeat predefined chunks as necessary, making this the first true generator on the Internet. It uses a dictionary of over 200 Latin words, combined with a handful of model sentence structures, to generate Lorem Ipsum which looks reasonable. The generated Lorem Ipsum is therefore always free from repetition, injected humour, or non-characteristic words etc.</p>"",
                            ""display"": ""block"",
                            ""cellType"": ""StaticText""
                        }
                    ],
                    ""colSpan"": ""1"",
                    ""rowSpan"": ""1"",
                    ""textAlign"": """"
                }
            ]
        },
        {
            ""cols"": [
                {
                    ""font"": {},
                    ""bgColor"": """",
                    ""color"": """",
                    ""colSpan"": ""1"",
                    ""rowSpan"": ""1"",
                    ""textAlign"": """",
                    ""cellContent"": [
                        {
                            ""valuesList"": [
                                ""Male"",
                                ""Female"",
                                ""Other""
                            ],
                            ""label"": ""Gender"",
                            ""display"": ""block"",
                            ""align"": ""block"",
                            ""cellType"": ""Radio""
                        }
                    ]
                },
                {
                    ""font"": {},
                    ""bgColor"": """",
                    ""color"": """",
                    ""cellContent"": [
                        {
                            ""valuesList"": [
                                ""Personal Home"",
                                ""Fire"",
                                ""Cooking""
                            ],
                            ""label"": ""LOB"",
                            ""display"": ""block"",
                            ""align"": ""inline-block"",
                            ""cellType"": ""Checkbox""
                        }
                    ],
                    ""colSpan"": ""1"",
                    ""rowSpan"": ""1"",
                    ""textAlign"": """"
                },
                {
                    ""font"": {},
                    ""bgColor"": """",
                    ""color"": """",
                    ""cellContent"": [
                        {
                            ""imageName"": ""whatsapp.png"",
                            ""imageSrc"": ""blob:http://localhost:4200/2b432dee-98d8-44ac-9813-c95da8a017ac""
,
                            ""ImageWidth"": ""150"",
                            ""ImageHeight"": ""150"",
                            ""cellType"": ""Image""
                        }
                    ],
                    ""colSpan"": ""1"",
                    ""rowSpan"": ""1"",
                    ""textAlign"": """"
                }
            ]
        },
        {
            ""cols"": [
                {
                    ""font"": {},
                    ""bgColor"": """",
                    ""color"": """",
                    ""colSpan"": 3,
                    ""rowSpan"": ""1"",
                    ""textAlign"": """",
                    ""cellContent"": [
                        {
                            ""component"": ""Address Information"",
                            ""cellType"": ""StaticText"",
                                                        ""label"": ""Address Information""
                        }
                    ]
                }
            ]
        }
    ],
    ""totalCols"": 3
}";
            /*    Not working full JSON - END    **/
            /** Working short code without tableJson
//            string json = @"{
//    ""rows"": [
//        {
//            ""cols"": [
//                {
//                    ""font"": {},
//                    ""bgColor"": """",
//                    ""color"": """",
//                    ""colSpan"": ""1"",
//                    ""rowSpan"": ""1"",
//                    ""textAlign"": """",
//                    ""cellContent"": [
//                        {
//                            ""label"": ""<p>There are many variations of passages of Lorem Ipsum available, but the majority have suffered alteration in some form, by injected humour, or randomised words which don't look even slightly believable. If you are going to use a passage of Lorem Ipsum, you need to be sure there isn't anything embarrassing hidden in the middle of text. All the Lorem Ipsum generators on the Internet tend to repeat predefined chunks as necessary, making this the first true generator on the Internet. It uses a dictionary of over 200 Latin words, combined with a handful of model sentence structures, to generate Lorem Ipsum which looks reasonable. The generated Lorem Ipsum is therefore always free from repetition, injected humour, or non-characteristic words etc.</p>"",
//                            ""display"": ""block"",
//                            ""cellType"": ""StaticText""
//                        }
//                    ]
//                },
//                {
//                    ""font"": {},
//                    ""bgColor"": """",
//                    ""color"": """",
//                    ""cellContent"": [],
//                    ""colSpan"": ""1"",
//                    ""rowSpan"": ""1"",
//                    ""textAlign"": """"
//                },
//                {
//                    ""font"": {},
//                    ""bgColor"": """",
//                    ""color"": """",
//                    ""cellContent"": [
//                        {
//                            ""label"": ""<p>There are many variations of passages of Lorem Ipsum available, but the majority have suffered alteration in some form, by injected humour, or randomised words which don't look even slightly believable. If you are going to use a passage of Lorem Ipsum, you need to be sure there isn't anything embarrassing hidden in the middle of text. All the Lorem Ipsum generators on the Internet tend to repeat predefined chunks as necessary, making this the first true generator on the Internet. It uses a dictionary of over 200 Latin words, combined with a handful of model sentence structures, to generate Lorem Ipsum which looks reasonable. The generated Lorem Ipsum is therefore always free from repetition, injected humour, or non-characteristic words etc.</p>"",
//                            ""display"": ""block"",
//                            ""cellType"": ""StaticText""
//                        }
//                    ],
//                    ""colSpan"": ""1"",
//                    ""rowSpan"": ""1"",
//                    ""textAlign"": """"
//                }
//            ]
//        },
//        {
//            ""cols"": [
//                {
//                    ""font"": {},
//                    ""bgColor"": """",
//                    ""color"": """",
//                    ""colSpan"": ""1"",
//                    ""rowSpan"": ""1"",
//                    ""textAlign"": """",
//                    ""cellContent"": [
//                        {
//                            ""label"": ""Gender"",
//                            ""display"": ""block"",
//                            ""align"": ""block"",
//                            ""cellType"": ""Radio""
//                        }
//                    ]
//                },
//                {
//                    ""font"": {},
//                    ""bgColor"": """",
//                    ""color"": """",
//                    ""cellContent"": [
//                        {
//                            ""label"": ""LOB"",
//                            ""display"": ""block"",
//                            ""align"": ""inline-block"",
//                            ""cellType"": ""Checkbox""
//                        }
//                    ],
//                    ""colSpan"": ""1"",
//                    ""rowSpan"": ""1"",
//                    ""textAlign"": """"
//                },
//                {
//                    ""font"": {},
//                    ""bgColor"": """",
//                    ""color"": """",
//                    ""cellContent"": [
//                        {
//                            ""imageName"": ""whatsapp.png"",
//                            ""imageSrc"": ""blob:http://localhost:4200/2b432dee-98d8-44ac-9813-c95da8a017ac""
//                            ,
//                            ""ImageWidth"": ""150"",
//                            ""ImageHeight"": ""150"",
//                            ""cellType"": ""Image""
//                        }
//                    ],
//                    ""colSpan"": ""1"",
//                    ""rowSpan"": ""1"",
//                    ""textAlign"": """"
//                }
//            ]
//        }
//    ],
//    ""totalCols"": ""3""
//}";
            */
            // Create a new Word application using late binding
            dynamic? wordApp = Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application")!);

            // Create a new document
            dynamic doc = wordApp!.Documents.Add();

            List<Dictionary<string, object>> DicList = new();
            JObject ParsedJson = JObject.Parse(json);
            Dictionary<string, object> Dictionaryobject = JsonConvert.DeserializeObject<Dictionary<string, object>>(JsonConvert.SerializeObject(ParsedJson))!;
            long rowsCount = (long)ParsedJson["rows"]!.Count();
            long colsCount = (long)Dictionaryobject["totalCols"];

            // Get the array of cell data objects from the ProcessJSON function
            List<Dictionary<string, object>> cellDataList = ProcessJSON(json);
            Console.WriteLine($"rowsCount: {rowsCount}, colsCount: {colsCount}");
            Microsoft.Office.Interop.Word.Table table = doc.Tables.Add(doc.Range(), rowsCount + 1, colsCount + 1);

            // Iterate through each cell data object
            foreach (Dictionary<string, object> cellData in cellDataList)
            {
                // Access properties of the cell data object
                int rowNumber = (int)cellData["rowNumber"];
                int colNumber = (int)cellData["colNumber"];
                string value = cellData.TryGetValue("value", out object? content) ? (string)content : "";
                string cellType = cellData.TryGetValue("cellType", out object? type) ? (string)type : "";

                // Execute your desired function, passing the relevant cell data
                Console.WriteLine($"rowNumber: {rowNumber}, colNumber: {colNumber}, value: {value}, cellType: {cellType}");
                InsertTextInTable(table, rowNumber, colNumber, value); // Replace "MyFunction" with your actual function name
            }

            //Console.WriteLine("{");
            int row = 0;
            //var result = RowsToResponse(json, doc: doc, row: ref row);
            //Console.WriteLine("}");

            // Save the document
            doc.SaveAs2("C:\\Documents\\example.docx");

            // Close Word application
            wordApp.Quit();
        }
        public static List<Dictionary<string, object>> ProcessJSON(string jsonString)
        {
            List<Dictionary<string, object>> result = new();
            JObject parsedJson = JObject.Parse(jsonString);

            // Track row and column indexes
            int rowIndex = 0;
            int colIndex = 0;

            // Iterate through rows
            foreach (var rowItem in parsedJson["rows"].Children())
            {
                // Iterate through columns within the row
                foreach (var colItem in rowItem["cols"].Children())
                {
                    Dictionary<string, object> cellData = new();
                    cellData["rowNumber"] = rowIndex;
                    cellData["colNumber"] = colIndex;

                    // Extract rowSpan and colSpan
                    int? rowSpan = colItem["rowSpan"]?.Value<int>();
                    int? colSpan = colItem["colSpan"]?.Value<int>();
                    cellData["rowSpan"] = rowSpan;
                    cellData["colSpan"] = colSpan;

                    // Access the first element of the cellContent array
                    JToken firstCellContent = colItem["cellContent"]?.FirstOrDefault();

                    if (firstCellContent != null)
                    {
                        // Now you can access properties of the first element
                        string label = firstCellContent["label"]?.ToString() ?? "";
                        string cellType = firstCellContent["cellType"]?.ToString() ?? "";

                        cellData["value"] = label;
                        cellData["cellType"] = cellType;


                        // Handle images (if applicable)
                        if (cellType == "Image")
                        {
                            // Incorporate image information
                            cellData["imageName"] = colItem["cellContent"]?.FirstOrDefault()["imageName"]?.ToString();
                            cellData["imageSrc"] = colItem["cellContent"]?.FirstOrDefault()["imageSrc"]?.ToString();
                            cellData["imageWidth"] = colItem["cellContent"]?.FirstOrDefault()["ImageWidth"]?.ToString();
                            cellData["imageHeight"] = colItem["cellContent"]?.FirstOrDefault()["ImageHeight"]?.ToString();
                        }
                    }

                    result.Add(cellData);

                    // Increment column index, accounting for colSpan
                    colIndex += colSpan ?? 1;
                }

                // Reset column index for the next row
                colIndex = 0;
                rowIndex++;
            }

            return result;
        }

        public static List<Dictionary<string, object>> RowsToResponse(string InputJsonString, dynamic doc, ref int row, int col = 0, Microsoft.Office.Interop.Word.Table? table = null)
        {
            List<Dictionary<string, object>> DicList = new();
            JObject ParsedJson = JObject.Parse(InputJsonString);
            Dictionary<string, object> Dictionaryobject = JsonConvert.DeserializeObject<Dictionary<string, object>>(JsonConvert.SerializeObject(ParsedJson))!;


            if (Dictionaryobject.ContainsKey("rows") && Dictionaryobject.TryGetValue("totalCols", out object? totalColumns))
            {
                try
                {
                    var totalRows = Dictionaryobject.Count;
                    table = doc.Tables.Add(doc.Range(), totalRows, totalColumns);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error: {ex.Message}");
                }
            }

            foreach (KeyValuePair<string, object> keyValuePair in Dictionaryobject.ToArray())
            {
                if (keyValuePair.Key == "cols")
                {
                    row++;
                    //Console.WriteLine($"Updating row at: {keyValuePair.Key}");
                }
                if (keyValuePair.Value is JObject)
                {
                    //Console.WriteLine($"\"{keyValuePair.Key}\": {{");
                    //Console.WriteLine(keyValuePair.Value);
                    RowsToResponse(JsonConvert.SerializeObject(keyValuePair.Value), doc: doc, row: ref row, table: table);
                    //Console.WriteLine("}");
                }
                if (keyValuePair.Key == "valuesList")
                {
                    continue;
                }
                if (keyValuePair.Value is JArray array)
                {
                    //Console.WriteLine($"\"{keyValuePair.Key}\": [");
                    foreach (JObject jObject in array.Cast<JObject>())
                    {
                        //Console.WriteLine($"{{");
                        //Console.WriteLine($"keyValuePair: {keyValuePair.Key}");
                        if (keyValuePair.Key == "cols")
                        {
                            col++;
                        }
                        RowsToResponse(JsonConvert.SerializeObject(jObject), doc: doc, row: ref row, col: col, table: table);
                        //Console.WriteLine("}");
                    }
                    //Console.WriteLine("]");
                }
                if (keyValuePair.Value is string)
                {
                    //Console.WriteLine($"\"{keyValuePair.Key}\": \"{keyValuePair.Value}\"{(isLast ? "" : ",")}");
                    string cellType;
                    int colSpan;
                    int rowSpan;
                    switch (keyValuePair.Key)
                    {
                        case "colSpan":
                            // Console.WriteLine($"colSpan: {keyValuePair.Value}");
                            colSpan = int.Parse(keyValuePair.Value.ToString()!);
                            break;
                        case "rowSpan":
                            // Console.WriteLine($"rowSpan: {keyValuePair.Value}");
                            rowSpan = int.Parse(keyValuePair.Value.ToString()!);
                            break;
                        case "cellType":
                            // Console.WriteLine($"cellType: {keyValuePair.Value}");
                            cellType = keyValuePair.Value.ToString()!;
                            break;
                        case "label":
                            // Console.WriteLine($"label: {keyValuePair.Value}");
                            break;
                        default:
                            break;
                    }
                    if (table != null && keyValuePair.Key == "label")
                    {
                        //Console.WriteLine($"cellType, colSpan, rowSpan, row, col: {cellType}, {colSpan}, {rowSpan}, {row}, {col}");
                        InsertTextInTable(table, row, col, null, keyValuePair);
                    }
                }
            }
            return DicList;
        }
        static void InsertTextInTable(Microsoft.Office.Interop.Word.Table table, int row, int column, string value, KeyValuePair<string, object>? cellContent = null)
        {
            row++;
            column++;
            if (row <= 0 || column <= 0 || row > table.Rows.Count || column > table.Columns.Count)
            {
                Console.WriteLine($"Invalid table position. {row} {row > table.Rows.Count}, {column} {column > table.Columns.Count}");
                return;
            }

            //if (cellContent.Value == "StaticText")
            //{
            //Console.WriteLine($"row: {row}, column: {column}, content: {cellContent?.Value ?? value}, currentRowIndex: {currentRowIndex}, cellRowIndex: {cellRowIndex}");
            table.Cell(row, column).Range.Text = cellContent?.Value.ToString() ?? value;
            //}

            //if (cellContent.Value == "DynamicInput")
            //{
            //    // Get the active window for the document
            //    dynamic window = doc.ActiveWindow;

            //    // Add a text range at the specified position
            //    dynamic range = window.Selection.Range;
            //    table.Cell(row, column).Range.Fields.Add(range, Microsoft.Office.Interop.Word.WdFieldType.wdFieldMergeField, cellContent.Key);
            //}
        }

    }

}
