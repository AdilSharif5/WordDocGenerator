using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

internal static class ProgramHelpers
{
    public static List<Dictionary<string, object>> RowsToResponse(string InputJsonString, bool isLast = false)
    {
        List<Dictionary<string, object>> DicList = new();
        JObject ParsedJson = JObject.Parse(InputJsonString);
        Dictionary<string, object> Dictionaryobject = JsonConvert.DeserializeObject<Dictionary<string, object>>(JsonConvert.SerializeObject(ParsedJson))!;

        foreach (KeyValuePair<string, object> keyValuePair in Dictionaryobject.ToArray())
        {
            if (keyValuePair.Value is JObject)
            {
                Console.WriteLine($"\"{keyValuePair.Key}\": {{");
                RowsToResponse(JsonConvert.SerializeObject(keyValuePair.Value), isLast: true);
                Console.WriteLine("}");
            }
            if (keyValuePair.Value is JArray)
            {
                Console.WriteLine($"\"{keyValuePair.Key}\": [");
                foreach (JObject jObject in ((JArray)keyValuePair.Value))
                {
                    Console.WriteLine($"{{");
                    RowsToResponse(JsonConvert.SerializeObject(jObject), isLast: true);
                    Console.WriteLine("}");
                }
                Console.WriteLine("]");
            }
            if (keyValuePair.Value is string)
            {
                Console.WriteLine($"\"{keyValuePair.Key}\": \"{keyValuePair.Value}\"{(isLast ? "" : ",")}");
            }
        }
        return DicList;
    }
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
                    ""colSpan"": ""3"",
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
                                ""totalCols"": ""4""
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
                    ""colSpan"": ""3"",
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
                                ""totalCols"": ""8""
                            },
                            ""cellType"": ""Component""
                        }
                    ]
                }
            ]
        }
    ],
    ""totalCols"": ""3""
}";
        Console.WriteLine("{");
        var result = RowsToResponse(json);
        Console.WriteLine("}");
        Console.WriteLine(result);
    }
}