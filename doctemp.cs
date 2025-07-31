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
            string templatePath = "Template.docx";   // Ensure this file exists
            string outputPath = "Output.docx";

            Console.WriteLine("🔍 Extracting content control structure...");
            var payload = ExtractRequiredPayload(templatePath);
            var jsonPayload = JsonConvert.SerializeObject(payload, Formatting.Indented);
            Console.WriteLine(jsonPayload);

            File.WriteAllText("PayloadTemplate.json", jsonPayload);

            Console.WriteLine("\n📝 Populating document with placeholder values...");
            PopulateContentControlsFromJson(templatePath, outputPath, jsonPayload);
            Console.WriteLine($"✅ Document generated at: {outputPath}");
        }

        public static JObject ExtractRequiredPayload(string filePath)
        {
            var payload = new JObject();

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false))
            {
                var sdtElements = wordDoc.MainDocumentPart.Document.Body.Descendants<SdtElement>();

                foreach (var sdt in sdtElements)
                {
                    var tag = sdt.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value;
                    if (string.IsNullOrWhiteSpace(tag))
                        continue;

                    // Repeating section logic (SdtBlock with nested SDTs)
                    if (sdt is SdtBlock sb && sb.Descendants<SdtElement>().Any(x => x != sdt && x.SdtProperties?.GetFirstChild<Tag>() != null))
                    {
                        if (!payload.ContainsKey(tag))
                        {
                            var innerFields = sb.Descendants<SdtElement>()
                                .Where(x => x != sdt && x.SdtProperties?.GetFirstChild<Tag>() != null)
                                .Select(x => x.SdtProperties.GetFirstChild<Tag>().Val.Value)
                                .Distinct();

                            var item = new JObject();
                            foreach (var field in innerFields)
                                item[field] = "";

                            payload[tag] = new JArray { item };
                        }
                    }
                    else
                    {
                        // Non-repeating simple field
                        if (!payload.ContainsKey(tag))
                            payload[tag] = "";
                    }
                }
            }

            return payload;
        }

        public static void PopulateContentControlsFromJson(string templatePath, string outputPath, string jsonPayload)
        {
            var payload = JObject.Parse(jsonPayload);
            var tempFile = Path.GetTempFileName();
            File.Copy(templatePath, tempFile, true);

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(tempFile, true))
            {
                var doc = wordDoc.MainDocumentPart.Document;
                var body = doc.Body;
                var sdtBlocks = body.Descendants<SdtBlock>().ToList();

                foreach (var sdt in sdtBlocks)
                {
                    var tag = sdt.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value;
                    if (string.IsNullOrWhiteSpace(tag) || !payload.ContainsKey(tag))
                        continue;

                    var token = payload[tag];

                    if (token.Type == JTokenType.String || token.Type == JTokenType.Integer)
                    {
                        // Single text replacement
                        foreach (var text in sdt.Descendants<Text>())
                            text.Text = token.ToString();
                    }
                    else if (token.Type == JTokenType.Array)
                    {
                        // Repeating section
                        var prototype = sdt.CloneNode(true);
                        var parent = sdt.Parent;
                        sdt.Remove(); // Remove the original block

                        foreach (var obj in token)
                        {
                            var newSdt = (SdtBlock)prototype.CloneNode(true);
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
