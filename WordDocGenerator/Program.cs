using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Microsoft.Office.Interop.Word;

namespace WordDocGenerator
{
    class Program
    {
        static void Main()
        {
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
            // Create a new Word application using late binding
            dynamic? wordApp = Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application")!);

            // Create a new document
            dynamic doc = wordApp!.Documents.Add();

            List<Dictionary<string, object>> DicList = new();
            JObject ParsedJson = JObject.Parse(json);
            Dictionary<string, object> Dictionaryobject = JsonConvert.DeserializeObject<Dictionary<string, object>>(JsonConvert.SerializeObject(ParsedJson))!;
            long rowsCount = ParsedJson["rows"]!.Count();
            long colsCount = (long)Dictionaryobject["totalCols"];

            // Get the array of cell data objects from the ProcessJSON function
            List<Dictionary<string, object>> cellDataList = ProcessJSON(json);
            Console.WriteLine($"rowsCount: {rowsCount}, colsCount: {colsCount}");
            Table table = doc.Tables.Add(doc.Range(), rowsCount, colsCount);

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
                InsertTextInTable(table, rowNumber, colNumber, value, cellData); // Replace "MyFunction" with your actual function name
            }
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
            foreach (var rowItem in parsedJson["rows"]!.Children())
            {
                // Iterate through columns within the row
                foreach (var colItem in rowItem["cols"]!.Children())
                {
                    Dictionary<string, object> cellData = new()
                    {
                        ["rowNumber"] = rowIndex + 1,
                        ["colNumber"] = colIndex + 1
                    };

                    // Extract rowSpan and colSpan
                    int? rowSpan = colItem["rowSpan"]?.Value<int>();
                    int? colSpan = colItem["colSpan"]?.Value<int>();
                    cellData["rowSpan"] = rowSpan!;
                    cellData["colSpan"] = colSpan!;

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

            Dictionary<int, List<KeyValuePair<int, int>>> mergedCells = new(); // Track merged cell ranges
            foreach (var cellData in result)
            {
                int rowNumber = (int)cellData["rowNumber"];
                int colNumber = (int)cellData["colNumber"];

                int? colSpan = (int?)cellData["colSpan"] ?? 1;
                int? rowSpan = (int?)cellData["rowSpan"] ?? 1;

                // Add merged range to the dictionary
                if (mergedCells.TryGetValue(rowNumber, out var rangesInRow))
                {
                    rangesInRow.Add(new KeyValuePair<int, int>(colNumber, (colNumber + colSpan - 1) ?? 1));
                }
                else
                {
                    mergedCells.Add(rowNumber, new List<KeyValuePair<int, int>> { new KeyValuePair<int, int>(colNumber, (colNumber + colSpan - 1) ?? 1) });
                }
            }

            // Add mergedCells information to each cellData object
            foreach (var cellData in result)
            {
                int rowNumber = (int)cellData["rowNumber"];
                cellData["mergedCells"] = mergedCells.TryGetValue(rowNumber, out var rangesInRow) ? rangesInRow : null;
            }


            return result;
        }
        static void InsertTextInTable(Table table, int row, int column, string value, Dictionary<string, object>? cellContent = null)
        {
            if (row <= 0 || column <= 0 || row > table.Rows.Count || column > table.Columns.Count)
            {
                Console.WriteLine($"Invalid table position. {row} {row > table.Rows.Count}, {column} {column > table.Columns.Count}");
                return;
            }
            //var mergedRanges = (List<KeyValuePair<int, int>>)cellContent["mergedCells"];
            //foreach (var range in mergedRanges)
            //{
            //Microsoft.Office.Interop.Word.Range cellRange = table.Cell(row, range.Value).Range;  // Get the range of the target cell
            //    //table.Cell(row, range.Key).Merge(table.Cell(row, range.Value));
            //cellRange.Cells.Merge();  // Merge the cells within the range
            //}
            if (cellContent != null && cellContent.ContainsKey("mergedCells"))
            {
                //var mergedRanges = (List<KeyValuePair<int, int>>)cellContent["mergedCells"];
                //foreach (var range in mergedRanges)
                //{
                    // Get the range of the cell containing the merge field
                    var cellRange = table.Cell(row, column).Range;

                    // Adjust the End property of the range to span multiple columns
                    cellRange.SetRange(cellRange.Start, table.Cell(row, column + int.Parse(cellContent["colSpan"].ToString()) - 1).Range.End);

                    // Merge the cells
                    cellRange.Cells.Merge();
                    //table.Cell(row, column).Merge(table.Cell(range.Key, range.Value));
                //}
            }
            Cell cell = table.Cell(row, column);

            // Access the existing paragraph or add a new one
            Paragraph paragraph = cell.Range.Paragraphs.Count > 0
                ? cell.Range.Paragraphs[1]
                : cell.Range.Paragraphs.Add();

            // Set the alignment to center
            paragraph.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

            // Add text to the cell, if needed
            paragraph.Range.Text = value;
            //table.Cell(row, column).Range.Text = value;
        }
    }
}