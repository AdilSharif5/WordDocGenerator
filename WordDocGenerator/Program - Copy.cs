using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json.Linq;

class Program
{
    static void main()
    {
        string jsonString = @"{
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
                            ""imageSrc"": ""blob:http://localhost:4200/2b432dee-98d8-44ac-9813-c95da8a017ac"",
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
        JToken json = JToken.Parse(jsonString);
        using (WordprocessingDocument wordDocument = WordprocessingDocument.Create("C:\\Documents\\example.docx", WordprocessingDocumentType.Document))
        {
            MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
            mainPart.Document = new Document();
            Body body = new Body();
            mainPart.Document.Append(body);
            CreateTableFromJson(json, body);
            wordDocument.Save();
        }
    }

    static void CreateTableFromJson(JToken json, OpenXmlElement parentElement)
    {
        int totalCols = json["totalCols"].Value<int>();
        Table table = new Table();

        foreach (var row in json["rows"])
        {
            TableRow tableRow = new TableRow();

            foreach (var col in row["cols"])
            {
                TableCell tableCell = new TableCell();

                if (col["cellContent"].HasValues)
                {
                    foreach (var cellContent in col["cellContent"])
                    {
                        switch (cellContent["cellType"].Value<string>())
                        {
                            case "Component":
                                Console.WriteLine($"cellContent: {cellContent["cellType"]}");
                                TableCell nestedTableCell = new();
                                //CreateTableFromJson(cellContent["tableJson"], nestedTableCell);
                                CreateTableFromJson(cellContent["tableJson"], tableCell);
                                //tableCell.AppendChild(nestedTableCell);
                                break;
                            case "StaticText":
                                Console.WriteLine($"cellContent: {cellContent["cellType"]}");
                                tableCell.Append(new Paragraph(new Run(new Text(cellContent["label"].Value<string>()))));
                                break;

                            case "DynamicInput":
                                // Use MergeField for dynamic content
                                string mergeFieldName = cellContent["label"].Value<string>();
                                tableCell.Append(new Paragraph(new Run(new FieldCode(mergeFieldName))));
                                tableCell.Append(new Paragraph(new Run(new Text(cellContent["label"].Value<string>()))));
                                break;
                            default:
                                break;
                        }
                    }
                }

                // Set alignment based on JSON data
                string? textAlign = col["textAlign"]?.Value<string>();
                if (!string.IsNullOrEmpty(textAlign))
                {
                    SetCellAlignment(tableCell, textAlign);
                }

                tableRow.Append(tableCell);
            }

            table.Append(tableRow);
        }

        parentElement.Append(table);
    }
    static void SetCellAlignment(TableCell cell, string alignment)
    {
        TableCellProperties properties = cell.Elements<TableCellProperties>().FirstOrDefault();
        if (properties == null)
        {
            properties = new TableCellProperties();
            cell.AppendChild(properties);
        }

        switch (alignment.ToLower())
        {
            case "center":
                properties.Append(new Justification() { Val = JustificationValues.Center });
                break;
            case "right":
                properties.Append(new Justification() { Val = JustificationValues.Right });
                break;
            // Handle other alignments as needed
            default:
                // Default to left alignment
                properties.Append(new Justification() { Val = JustificationValues.Left });
                break;
        }
        //properties.Append(new Justification() { Val = JustificationValues.Left });
    }


}
