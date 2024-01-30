using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Microsoft.Office.Interop.Word;
namespace WordDocGenerator
{
    class Program
    {
        static void Main()
        {
            #region
            //            string json = @"{
            //    ""rows"": [
            //        {
            //            ""cols"": [
            //                {
            //                    ""font"": {
            //""family"": ""Karla"",
            //""size"": 24
            //},
            //                    ""bgColor"": ""#32a852"",
            //                    ""color"": ""#fff"",
            //                    ""colSpan"": 3,
            //                    ""rowSpan"": ""1"",
            //                    ""textAlign"": """",
            //                    ""cellContent"": [
            //                        {
            //                            ""component"": ""Companion Policy Information"",
            //                            ""cellType"": ""StaticText"",
            //							""label"": ""Companion Policy Information"",
            //                        }
            //                    ]
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
            //                            ""valuesList"": [
            //                                ""Male"",
            //                                ""Female"",
            //                                ""Other""
            //                            ],
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
            //                            ""valuesList"": [
            //                                ""Personal Home"",
            //                                ""Fire"",
            //                                ""Cooking""
            //                            ],
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
            //,
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
            //        },
            //        {
            //            ""cols"": [
            //                {
            //                    ""font"": {},
            //                    ""bgColor"": """",
            //                    ""color"": """",
            //                    ""colSpan"": 3,
            //                    ""rowSpan"": ""1"",
            //                    ""textAlign"": """",
            //                    ""cellContent"": [
            //                        {
            //                            ""component"": ""Address Information"",
            //                            ""cellType"": ""StaticText"",
            //                                                        ""label"": ""Address Information""
            //                        }
            //                    ]
            //                }
            //            ]
            //        }
            //    ],
            //    ""totalCols"": 3
            //}";
            #endregion
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
                            ""tableJson"": {
                                ""rows"": [
                                    {
                                        ""cols"": [
                                            {
                                                ""font"": {},
                                                ""bgColor"": ""#272872"",
                                                ""color"": ""#fff"",
                                                ""cellContent"": [
                                                    {
                                                        ""cellType"": ""StaticText"",
                                                        ""label"": ""Companion Policy Information""
                                                    }
                                                ],
                                                ""colSpan"": ""4"",
                                                ""rowSpan"": ""1"",
                                                ""textAlign"": ""center""
                                            }
                                        ]
                                    },
                                    {
                                        ""cols"": [
                                            {
                                                ""font"": {},
                                                ""bgColor"": """",
                                                ""cellContent"": [
                                                    {
                                                        ""cellType"": ""DynamicInput"",
                                                        ""label"": ""Participating Insurer""
                                                    }
                                                ],
                                                ""colSpan"": ""1"",
                                                ""rowSpan"": ""1""
                                            },
                                            {
                                                ""font"": {},
                                                ""bgColor"": """",
                                                ""cellContent"": [
                                                    {
                                                        ""cellType"": ""DynamicInput"",
                                                        ""label"": ""Companion Policy Number""
                                                    }
                                                ],
                                                ""colSpan"": ""1"",
                                                ""rowSpan"": ""1""
                                            },
                                            {
                                                ""font"": {},
                                                ""bgColor"": """",
                                                ""cellContent"": [
                                                    {
                                                        ""cellType"": ""DynamicInput"",
                                                        ""label"": ""Dwelling - Coverage A Limit""
                                                    }
                                                ],
                                                ""colSpan"": ""1"",
                                                ""rowSpan"": ""1""
                                            },
                                            {
                                                ""font"": {},
                                                ""bgColor"": """",
                                                ""cellContent"": [
                                                    {
                                                        ""cellType"": ""DynamicInput"",
                                                        ""label"": ""Expiration Date""
                                                    }
                                                ],
                                                ""colSpan"": ""1"",
                                                ""rowSpan"": ""1""
                                            }
                                        ]
                                    }
                                ],
                                ""totalCols"": 4
                            },
                            ""cellType"": ""Component""
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
                            ""imageSrc"": ""blob:
http://localhost:4200/2b432dee-98d8-44ac-9813-c95da8a017ac""
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
                            ""tableJson"": {
                                ""rows"": [
                                    {
                                        ""cols"": [
                                            {
                                                ""font"": {},
                                                ""bgColor"": ""#272872"",
                                                ""color"": ""#fff"",
                                                ""cellContent"": [
                                                    {
                                                        ""cellType"": ""StaticText"",
                                                        ""label"": ""Address Information""
                                                    }
                                                ],
                                                ""colSpan"": ""8"",
                                                ""rowSpan"": ""1"",
                                                ""textAlign"": ""center""
                                            }
                                        ]
                                    },
                                    {
                                        ""cols"": [
                                            {
                                                ""font"": {},
                                                ""bgColor"": """",
                                                ""cellContent"": [
                                                    {
                                                        ""cellType"": ""DynamicInput"",
                                                        ""label"": ""Risk Address - Physical Location of Property - Number and Street Address""
                                                    }
                                                ],
                                                ""colSpan"": ""6"",
                                                ""rowSpan"": ""1""
                                            },
                                            {
                                                ""font"": {},
                                                ""bgColor"": """",
                                                ""cellContent"": [
                                                    {
                                                        ""cellType"": ""DynamicInput"",
                                                        ""label"": ""City""
                                                    }
                                                ],
                                                ""colSpan"": ""1"",
                                                ""rowSpan"": ""1""
                                            },
                                            {
                                                ""font"": {},
                                                ""bgColor"": """",
                                                ""cellContent"": [
                                                    {
                                                        ""cellType"": ""DynamicInput"",
                                                        ""label"": ""State""
                                                    }
                                                ],
                                                ""colSpan"": ""1"",
                                                ""rowSpan"": ""1""
                                            }
                                        ]
                                    }
                                ],
                                ""totalCols"": 8
                            },
                            ""cellType"": ""Component""
                        }
                    ]
                }
            ]
        }
    ],
    ""totalCols"": 3
}";
            // Create a new Word application using late binding
            dynamic? wordApp = Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application"));
            try
            {
                // Create a new document
                dynamic doc = wordApp!.Documents.Add();
                JObject ParsedJson = JObject.Parse(json);
                Dictionary<string, object> Dictionaryobject = JsonConvert.DeserializeObject<Dictionary<string, object>>(JsonConvert.SerializeObject(ParsedJson))!;
                // Get the array of cell data objects from the ProcessJSON function
                List<Dictionary<string, object>> cellDataList = ProcessJSON(json);
                CreateTable(doc, cellDataList);
                // Save the document
                doc.SaveAs2("C:\\Documents\\example.docx");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            finally
            {
                // Close Word application
                wordApp!.Quit();
            }
        }
        public static List<Dictionary<string, object>> ProcessJSON(string jsonString)
        {
            List<Dictionary<string, object>> result = new();
            JObject parsedJson = JObject.Parse(jsonString);
            long rowsCount = parsedJson["rows"]!.Count();
            long colsCount = (long)parsedJson["totalCols"]!;
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
                        ["colNumber"] = colIndex + 1,
                        ["rowsCount"] = rowsCount,
                        ["colsCount"] = colsCount
                    };
                    // Extract needed keys and values
                    int? rowSpan = colItem["rowSpan"]?.Value<int>();
                    int? colSpan = colItem["colSpan"]?.Value<int>();
                    string? textAlign = colItem["textAlign"]?.Value<string>();
                    string? bgColor = colItem["bgColor"]?.Value<string>();
                    string? color = colItem["color"]?.Value<string>();
                    cellData["rowSpan"] = rowSpan!;
                    cellData["colSpan"] = colSpan!;
                    cellData["textAlign"] = textAlign!;
                    cellData["bgColor"] = bgColor!;
                    cellData["color"] = color!;
                    if (colItem.HasValues && colItem["font"]!.Type == JTokenType.Object && colItem["font"]!.Children().Any())
                    {
                        var font = colItem["font"];
                        string? fontFamily = font["family"]?.Value<string>() ?? string.Empty;
                        float fontSize = font["size"]?.Value<float>() ?? 0;
                        cellData["fontFamily"] = fontFamily;
                        cellData["fontSize"] = fontSize;
                    }
                    // Access the first element of the cellContent array
                    JToken? firstCellContent = colItem["cellContent"]?.FirstOrDefault();
                    if (firstCellContent != null)
                    {

                        string label = firstCellContent["label"]?.ToString() ?? "";
                        string cellType = firstCellContent["cellType"]?.ToString() ?? "";
                        cellData["value"] = label;
                        cellData["cellType"] = cellType;

                        // Now you can access properties of the first element
                        if (firstCellContent["tableJson"] != null)
                        {
                            var nestedContent = ProcessJSON(firstCellContent["tableJson"]!.ToString());
                            cellData["table"] = nestedContent;
                        }
                        // Handle images (if applicable)
                        if (cellType == "Image")
                        {
                            // Incorporate image information
                            cellData["imageName"] = colItem["cellContent"]?.FirstOrDefault()!["imageName"]?.ToString()!;
                            cellData["imageSrc"] = colItem["cellContent"]?.FirstOrDefault()!["imageSrc"]?.ToString()!;
                            cellData["imageWidth"] = colItem["cellContent"]?.FirstOrDefault()!["ImageWidth"]?.ToString()!;
                            cellData["imageHeight"] = colItem["cellContent"]?.FirstOrDefault()!["ImageHeight"]?.ToString()!;
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
        public static void CreateTable(dynamic doc, List<Dictionary<string, object>> contentList, Table? parentTable = null, int? row = null, int? col = null)
        {
            long rowsCount = (long)contentList[0]["rowsCount"];
            long colsCount = (long)contentList[0]["colsCount"];
            var table = (parentTable != null) ? doc.Tables.Add(parentTable.Cell(row ?? 1, col ?? 1).Range, rowsCount, colsCount) : doc.Tables.Add(doc.Range(), rowsCount, colsCount);
            if (parentTable != null)
            {
                for (int i = 1; i < rowsCount; i++)
                {
                    _ = table.Rows.Add();
                }
            }
            foreach (Dictionary<string, object> cellData in contentList)
            {
                // Access properties of the cell data object
                int rowNumber = (int)cellData["rowNumber"];
                int colNumber = (int)cellData["colNumber"];
                string value = cellData.TryGetValue("value", out object? content) ? (string)content : "";
                string cellType = cellData.TryGetValue("cellType", out object? type) ? (string)type : "";
                List<Dictionary<string, object>> tableJson = (List<Dictionary<string, object>>)(cellData.TryGetValue("table", out object? tableContent) ? tableContent : null)!;
                if (tableJson != null)
                {
                    CreateTable(doc, tableJson, table, rowNumber, colNumber);
                }
                if (cellType == "Image")
                {
                    InsertImage(table, rowNumber, colNumber, cellData);
                }
                else
                {
                    InsertTextInTable(table, rowNumber, colNumber, value, cellData);
                }
            }
        }
        static void InsertTextInTable(Table table, int row, int column, string value, Dictionary<string, object>? cellContent = null)
        {
            if (row <= 0 || column <= 0 || row > table.Rows.Count || column > table.Columns.Count)
            {
                //Console.WriteLine($"Invalid table position. {row} {row > table.Rows.Count}, {column} {column > table.Columns.Count}");
                return;
            }
            /** maybe can be used to merge rows?
            //var mergedRanges = (List<KeyValuePair<int, int>>)cellContent["mergedCells"];
            //foreach (var range in mergedRanges)
            //{
            //Microsoft.Office.Interop.Word.Range cellRange = table.Cell(row, range.Value).Range;  // Get the range of the target cell
            //    //table.Cell(row, range.Key).Merge(table.Cell(row, range.Value));
            //cellRange.Cells.Merge();  // Merge the cells within the range
            //}
            */
            Cell cell = table.Cell(row, column);
            if (cellContent != null)
            {
                // Merge cell start
                #region
                if (cellContent.ContainsKey("mergedCells"))
                {
                    var cellRange = cell.Range;
                    // Adjust the End property of the range to span multiple columns
                    var rowSpan = (row + (int.Parse(cellContent["rowSpan"].ToString()!) - 1));
                    var colSpan = (column + (int.Parse(cellContent["colSpan"].ToString()!) - 1));
                    cellRange.SetRange(cellRange.Start, table.Cell(rowSpan, colSpan).Range.End);
                    // Merge the cells
                    cellRange.Cells.Merge();
                }
                #endregion
                // Merge cell end
                //styles start
                #region
                Font font = cell.Range.Font;
                string? textAlign = cellContent.TryGetValue("textAlign", out object? content) ? (string)content : "";
                string? bgColor = cellContent.TryGetValue("bgColor", out object? bColor) ? (string)bColor : "";
                string? textColor = cellContent.TryGetValue("color", out object? fontColor) ? (string)fontColor : "";
                string? fontFamily = cellContent.TryGetValue("fontFamily", out object? family) ? (string)family : "";
                float? fontSize = cellContent.TryGetValue("fontSize", out object? size) ? (float)size : null;
                cellContent.TryGetValue("cellType", out object? cellType);
                Paragraph paragraph = AlignCellContent(cell, textAlign, (cellType?.ToString() == "DynamicInput"))!;

                if (bgColor != null && bgColor != "")
                {
                    System.Drawing.Color backgroundColor = HexToColor(bgColor);
                    // Set background color
                    cell.Range.Shading.BackgroundPatternColor = (WdColor)(backgroundColor.R + 0x100 * backgroundColor.G + 0x10000 * backgroundColor.B);
                }
                if (textColor != null && textColor != "")
                {
                    System.Drawing.Color color = HexToColor(textColor);
                    font.Color = (WdColor)(color.R + 0x100 * color.G + 0x10000 * color.B);
                }
                if (fontFamily != "")
                {
                    // Set font family
                    font.Name = fontFamily;
                }
                if (fontSize != null)
                {
                    // Set font size
                    font.Size = (float)fontSize;
                }
                #endregion
                //styles end
                if (cellType?.ToString() == "DynamicInput")
                {
                    AddMergeFieldToCell(cell, value, textAlign);
                }
                else
                {
                    paragraph.Range.Text += value;
                }
            }
            else
            {
                cell.Range.Text = value;
            }
        }
        static void InsertImage(Table table, int row, int column, Dictionary<string, object>? cellContent)
        {
            // Get the current selection or range in the document
            Cell cell = table.Cell(row, column);

            // Insert the image
            InlineShape picture = cell.Range.InlineShapes.AddPicture("");

            // Optionally, you can modify properties of the inserted picture, such as width and height
            picture.Width = 200; // Set the width in points
            picture.Height = 150; // Set the height in points
        }
        static void AddMergeFieldToCell(Cell cell, string fieldName, string textAlign)
        {
            // Duplicate the current cell range
            Microsoft.Office.Interop.Word.Range range = cell.Range.Duplicate;
            AlignCellContent(cell, textAlign, true, range);
            object objType = WdFieldType.wdFieldMergeField;
            object objFieldName = fieldName;

            /** important, don't remove the below code is used for getting rid of "Command not available" as such error! **/
            #region
            // Move one step forward
            range.MoveStart(WdUnits.wdCharacter, 1);
            // Move it back to the cell range, basically reset.
            range.SetRange(range.Start - 1, range.End - 1);
            #endregion
            /** End of error fix! **/
            // Insert the merge field
            range.Fields.Add(range, objType, objFieldName, false);
        }
        static dynamic AlignCellContent(Cell cell, string textAlign = "", bool isDynamicField = false, Microsoft.Office.Interop.Word.Range? range = null)
        {
            // Access the existing paragraph or add a new one
            dynamic paragraphOrRange = (isDynamicField ? range! : ((cell.Range.Paragraphs.Count > 0)
                ? cell.Range.Paragraphs[1]
                : cell.Range.Paragraphs.Add())!);
            switch (textAlign)
            {
                case "left":
                    SetAlignment(paragraphOrRange, WdParagraphAlignment.wdAlignParagraphLeft);
                    break;
                case "right":
                    SetAlignment(paragraphOrRange, WdParagraphAlignment.wdAlignParagraphRight);
                    break;
                case "center":
                    SetAlignment(paragraphOrRange, WdParagraphAlignment.wdAlignParagraphCenter);
                    break;
                default:
                    break;
            }
            return paragraphOrRange;
        }
        static void SetAlignment(dynamic x, WdParagraphAlignment alignment)
        {
            x.Alignment = alignment;
        }
        static System.Drawing.Color HexToColor(string hex)
        {
            hex = hex.TrimStart('#');
            if (hex.Length < 6)
            {
                hex += new string(hex[0], 3);
            }
            int rgb = int.Parse(hex, System.Globalization.NumberStyles.HexNumber);
            byte red = (byte)(rgb >> 16);
            byte green = (byte)(rgb >> 8);
            byte blue = (byte)(rgb);
            return System.Drawing.Color.FromArgb(red, green, blue);
        }
    }
}