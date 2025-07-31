using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json.Linq;

namespace OpenXmlContentControlDemo
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string templatePath = "Template.docx"; // Ensure this file exists
            string outputPath = "Output.docx";

            Console.WriteLine("Extracting content controls...");
            var controls = ExtractContentControls(templatePath);
            foreach (var kvp in controls)
            {
                Console.WriteLine($"Tag: {kvp.Key} => Value(s): {string.Join(", ", kvp.Value)}");
            }

            Console.WriteLine("\nPopulating Word with JSON data...");

            string json = @"{
                'FirstName': 'Alice',
                'Bio': 'Full-stack developer.',
                'Skills': ['C#', 'JavaScript', 'SQL']
            }";

            PopulateContentControlsFromJson(templatePath, outputPath, json);
            Console.WriteLine($"Done. New file saved as '{outputPath}'.");
        }

        public static Dictionary<string, List<string>> ExtractContentControls(string filePath)
        {
            var contentControlMap = new Dictionary<string, List<string>>();

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false))
            {
                var sdtElements = wordDoc.MainDocumentPart.Document.Descendants<SdtElement>();

                foreach (var sdt in sdtElements)
                {
                    var tag = sdt.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value;
                    var text = string.Join("", sdt.Descendants<Text>().Select(t => t.Text));

                    if (string.IsNullOrWhiteSpace(tag)) continue;

                    if (!contentControlMap.ContainsKey(tag))
                        contentControlMap[tag] = new List<string>();

                    contentControlMap[tag].Add(text);
                }
            }

            return contentControlMap;
        }

        public static void PopulateContentControlsFromJson(string templatePath, string outputPath, string jsonPayload)
        {
            var payload = JObject.Parse(jsonPayload);
            string tempFile = Path.GetTempFileName();
            File.Copy(templatePath, tempFile, overwrite: true);

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(tempFile, true))
            {
                var doc = wordDoc.MainDocumentPart.Document;

                foreach (var sdt in doc.Descendants<SdtElement>().ToList())
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
                        var parent = sdt.Parent;
                        var prototype = sdt.CloneNode(true);
                        sdt.Remove();

                        foreach (var item in token)
                        {
                            var newSdt = (SdtElement)prototype.CloneNode(true);
                            foreach (var text in newSdt.Descendants<Text>())
                            {
                                text.Text = item.ToString();
                            }
                            parent.AppendChild(newSdt);
                        }
                    }
                }

                doc.Save();
            }

            File.Copy(tempFile, outputPath, overwrite: true);
            File.Delete(tempFile);
        }
    }
}
