using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace OpenXmlContentControlDemo
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string templatePath = "Template.docx"; // Ensure this exists
            string outputPath = "Output.docx";

            Console.WriteLine("🔍 Extracting payload...");
            var payload = ExtractRequiredPayload(templatePath);
            var json = JsonConvert.SerializeObject(payload, Formatting.Indented);
            File.WriteAllText("PayloadTemplate.json", json);
            Console.WriteLine(json);

            Console.WriteLine("\n📝 Populating document...");
            PopulateContentControlsFromJson(templatePath, outputPath, json);
            Console.WriteLine($"✅ Output saved to: {outputPath}");
        }

        public static JObject ExtractRequiredPayload(string filePath)
        {
            var payload = new JObject();
            var childTagsInsideRepeaters = new HashSet<string>();
            var repeatingSectionTags = new HashSet<string>();
            var repeatingSectionStructures = new Dictionary<string, List<string>>();

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false))
            {
                var body = wordDoc.MainDocumentPart.Document.Body;
                var sdtElements = body.Descendants<SdtElement>();

                // First pass: identify repeaters
                foreach (var sdt in sdtElements)
                {
                    var tag = sdt.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value;
                    if (string.IsNullOrWhiteSpace(tag)) continue;

                    if (sdt is SdtBlock || sdt is SdtRow)
                    {
                        var innerFields = sdt.Descendants<SdtElement>()
                            .Where(x => x != sdt && x.SdtProperties?.GetFirstChild<Tag>() != null)
                            .Select(x => x.SdtProperties.GetFirstChild<Tag>().Val.Value)
                            .Distinct()
                            .ToList();

                        if (innerFields.Any())
                        {
                            repeatingSectionTags.Add(tag);
                            repeatingSectionStructures[tag] = innerFields;
                            foreach (var childTag in innerFields)
                                childTagsInsideRepeaters.Add(childTag);
                        }
                    }
                }

                // Second pass: build payload
                foreach (var sdt in sdtElements)
                {
                    var tag = sdt.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value;
                    if (string.IsNullOrWhiteSpace(tag)) continue;

                    if (repeatingSectionTags.Contains(tag))
                    {
                        if (!payload.ContainsKey(tag))
                        {
                            var item = new JObject();
                            foreach (var field in repeatingSectionStructures[tag])
                                item[field] = "";
                            payload[tag] = new JArray { item };
                        }
                    }
                    else if (!childTagsInsideRepeaters.Contains(tag))
                    {
                        if (!payload.ContainsKey(tag))
                        {
                            if (sdt.SdtProperties?.GetFirstChild<CheckBox>() != null)
                                payload[tag] = false;
                            else if (sdt.SdtProperties?.GetFirstChild<Date>() != null)
                                payload[tag] = "2025-07-31";
                            else if (sdt.SdtProperties?.GetFirstChild<DropDownList>() != null)
                                payload[tag] = "";
                            else
                                payload[tag] = "";
                        }
                    }
                }
            }

            return payload;
        }

        public static void PopulateContentControlsFromJson(string templatePath, string outputPath, string jsonPayload)
        {
            var payload = JObject.Parse(jsonPayload);
            string tempFile = Path.GetTempFileName();
            File.Copy(templatePath, tempFile, true);

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(tempFile, true))
            {
                var doc = wordDoc.MainDocumentPart.Document;
                var body = doc.Body;
                var sdtElements = body.Descendants<SdtElement>().ToList();

                foreach (var sdt in sdtElements)
                {
                    var tag = sdt.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value;
                    if (string.IsNullOrWhiteSpace(tag) || !payload.ContainsKey(tag)) continue;

                    var token = payload[tag];

                    if (token.Type == JTokenType.String || token.Type == JTokenType.Integer)
                    {
                        foreach (var text in sdt.Descendants<Text>())
                            text.Text = token.ToString();
                    }
                    else if (token.Type == JTokenType.Boolean && sdt.SdtProperties.GetFirstChild<CheckBox>() != null)
                    {
                        var isChecked = token.Value<bool>();
                        var val = isChecked ? "☒" : "☐";

                        foreach (var text in sdt.Descendants<Text>())
                            text.Text = val;
                    }
                    else if (token.Type == JTokenType.String && sdt.SdtProperties?.GetFirstChild<Date>() != null)
                    {
                        foreach (var text in sdt.Descendants<Text>())
                            text.Text = token.ToString();
                    }
                    else if (token.Type == JTokenType.String && sdt.SdtProperties?.GetFirstChild<DropDownList>() != null)
                    {
                        foreach (var text in sdt.Descendants<Text>())
                            text.Text = token.ToString();
                    }
                    else if (token.Type == JTokenType.Array)
                    {
                        var prototype = sdt.CloneNode(true);
                        var parent = sdt.Parent;
                        sdt.Remove();

                        foreach (var obj in token)
                        {
                            var newSdt = (SdtElement)prototype.CloneNode(true);
                            var objFields = (JObject)obj;

                            foreach (var innerSdt in newSdt.Descendants<SdtElement>())
                            {
                                var innerTag = innerSdt.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value;
                                if (!string.IsNullOrWhiteSpace(innerTag) && objFields.ContainsKey(innerTag))
                                {
                                    foreach (var text in innerSdt.Descendants<Text>())
                                        text.Text = objFields[innerTag]?.ToString();
                                }
                            }

                            parent.AppendChild(newSdt);
                        }
                    }
                }

                doc.Save();
            }

            File.Copy(tempFile, outputPath, true);
            File.Delete(tempFile);
        }
    }
}
