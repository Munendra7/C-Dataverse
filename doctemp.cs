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
            string templatePath = "Template.docx";
            string outputPath = "Output.docx";

            Console.WriteLine("Extracting required payload from template...");
            var payload = ExtractRequiredPayload(templatePath);

            string jsonPayload = JsonConvert.SerializeObject(payload, Formatting.Indented);
            Console.WriteLine(jsonPayload);

            // Save payload for reuse
            File.WriteAllText("PayloadTemplate.json", jsonPayload);

            Console.WriteLine("\nPopulating template with sample values...");
            PopulateContentControlsFromJson(templatePath, outputPath, jsonPayload);
            Console.WriteLine($"Done. Output saved to {outputPath}");
        }

        public static JObject ExtractRequiredPayload(string filePath)
        {
            var payload = new JObject();

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false))
            {
                var sdtElements = wordDoc.MainDocumentPart.Document.Body.Descendants<SdtElement>();

                var repeatingGroups = new Dictionary<string, List<SdtElement>>();

                foreach (var sdt in sdtElements)
                {
                    var tag = sdt.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value;
                    if (string.IsNullOrWhiteSpace(tag)) continue;

                    if (sdt.GetType() == typeof(SdtBlock) && sdt.Descendants<SdtElement>().Any(e => e != sdt && e.SdtProperties?.GetFirstChild<Tag>() != null))
                    {
                        // Repeating section (with nested SDTs)
                        if (!repeatingGroups.ContainsKey(tag))
                            repeatingGroups[tag] = new List<SdtElement>();
                        repeatingGroups[tag].Add(sdt);
                    }
                    else
                    {
                        // Single field
                        if (!payload.ContainsKey(tag))
                            payload[tag] = "";
                    }
                }

                // Process repeating content controls
                foreach (var group in repeatingGroups)
                {
                    var tag = group.Key;
                    var array = new JArray();

                    // Pick first instance to understand nested tags
                    var first = group.Value.FirstOrDefault();
                    if (first == null) continue;

                    var innerTags = first.Descendants<SdtElement>()
                        .Where(x => x.SdtProperties?.GetFirstChild<Tag>() != null)
                        .Select(x => x.SdtProperties.GetFirstChild<Tag>().Val.Value)
                        .Distinct();

                    var item = new JObject();
                    foreach (var innerTag in innerTags)
                    {
                        item[innerTag] = "";
                    }

                    array.Add(item);
                    payload[tag] = array;
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

                var sdtBlocks = body.Descendants<SdtBlock>().ToList();

                foreach (var sdt in sdtBlocks)
                {
                    var tag = sdt.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value;
                    if (string.IsNullOrWhiteSpace(tag) || !payload.ContainsKey(tag)) continue;

                    var token = payload[tag];

                    if (token.Type == JTokenType.String || token.Type == JTokenType.Integer)
                    {
                        foreach (var text in sdt.Descendants<Text>())
                            text.Text = token.ToString();
                    }
                    else if (token.Type == JTokenType.Array)
                    {
                        // Repeating block
                        var parent = sdt.Parent;
                        var prototype = sdt.CloneNode(true);
                        var jsonArray = (JArray)token;

                        sdt.Remove(); // remove original

                        foreach (var entry in jsonArray)
                        {
                            var newSdt = (SdtBlock)prototype.CloneNode(true);
                            var entryObject = (JObject)entry;

                            foreach (var innerSdt in newSdt.Descendants<SdtElement>())
                            {
                                var innerTag = innerSdt.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value;
                                if (!string.IsNullOrWhiteSpace(innerTag) && entryObject.ContainsKey(innerTag))
                                {
                                    foreach (var text in innerSdt.Descendants<Text>())
                                        text.Text = entryObject[innerTag]?.ToString();
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
