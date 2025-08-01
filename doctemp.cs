using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using static System.Net.Mime.MediaTypeNames;
using static System.Runtime.InteropServices.JavaScript.JSType;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace CreateDocFromTemplate
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string templatePath = "C:\\Users\\munendra\\Downloads\\BKU Template - Red.docx";
            string outputPath = "C:\\Users\\munendra\\Downloads\\Output.docx";

            Console.WriteLine("🔍 Extracting payload...");
            var payload = ExtractRequiredPayload(templatePath);
            var json = JsonConvert.SerializeObject(payload, Formatting.Indented);
            File.WriteAllText("PayloadTemplate.json", json);
            Console.WriteLine(json);


            Console.WriteLine("\n📝 Populating document...");
            PopulateContentControlsFromJson(templatePath, outputPath, sampleJson);
            Console.WriteLine($"✅ Output saved to: {outputPath}");
        }

        public static JObject ExtractRequiredPayload(string filePath)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false))
            {
                var body = wordDoc.MainDocumentPart.Document.Body;
                return ExtractPayloadFromSdtElements(body.Descendants<SdtElement>());
            }
        }

        private static JObject ExtractPayloadFromSdtElements(IEnumerable<SdtElement> sdtElements)
        {
            var payload = new JObject();
            var childTagsInsideRepeaters = new HashSet<string>();
            var repeatingSectionTags = new HashSet<string>();
            var repeatingSectionStructures = new Dictionary<string, List<SdtElement>>();

            // Identify repeaters and their children
            foreach (var sdt in sdtElements)
            {
                var tag = sdt.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value;
                if (string.IsNullOrWhiteSpace(tag)) continue;

                if (sdt is SdtBlock || sdt is SdtRow)
                {
                    var innerSdts = sdt.Descendants<SdtElement>()
                        .Where(x => x != sdt && x.SdtProperties?.GetFirstChild<Tag>() != null)
                        .ToList();

                    if (innerSdts.Any())
                    {
                        repeatingSectionTags.Add(tag);
                        repeatingSectionStructures[tag] = innerSdts;
                        foreach (var childSdt in innerSdts)
                            childTagsInsideRepeaters.Add(childSdt.SdtProperties.GetFirstChild<Tag>().Val.Value);
                    }
                }
            }

            // Build payload recursively
            foreach (var sdt in sdtElements)
            {
                var tag = sdt.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value;
                if (string.IsNullOrWhiteSpace(tag)) continue;

                if (repeatingSectionTags.Contains(tag))
                {
                    if (!payload.ContainsKey(tag))
                    {
                        var arr = new JArray();
                        var item = ExtractPayloadFromSdtElements(repeatingSectionStructures[tag]);
                        arr.Add(item);
                        payload[tag] = arr;
                    }
                }
                else if (!childTagsInsideRepeaters.Contains(tag))
                {
                    if (!payload.ContainsKey(tag))
                    {
                        if (sdt.SdtProperties?.GetFirstChild<CheckBox>() != null)
                            payload[tag] = false;
                        else if (sdt.SdtProperties?.GetFirstChild<SdtContentDropDownList>() != null)
                            payload[tag] = "";
                        else
                            payload[tag] = "";
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
                PopulateSdtElements(body, payload);
                doc.Save();
            }

            File.Copy(tempFile, outputPath, true);
            File.Delete(tempFile);
        }

        private static void PopulateSdtElements(OpenXmlElement parent, JObject payload)
        {
            var sdtElements = parent.Descendants<SdtElement>().ToList();

            foreach (var sdt in sdtElements)
            {
                var tag = sdt.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value;
                if (string.IsNullOrWhiteSpace(tag) || !payload.ContainsKey(tag)) continue;

                var token = payload[tag];

                if (token.Type == JTokenType.String || token.Type == JTokenType.Integer)
                {
                    SetSingleText(sdt, token.ToString());
                }
                else if (token.Type == JTokenType.Boolean && sdt.SdtProperties.GetFirstChild<CheckBox>() != null)
                {
                    var isChecked = token.Value<bool>();
                    var val = isChecked ? "☒" : "☐";
                    SetSingleText(sdt, val);
                }
                else if (token.Type == JTokenType.String && sdt.SdtProperties?.GetFirstChild<SdtContentDropDownList>() != null)
                {
                    SetSingleText(sdt, token.ToString());
                }
                else if (token.Type == JTokenType.Array)
                {
                    var prototype = sdt.CloneNode(true);
                    var parentElement = sdt.Parent;
                    sdt.Remove();

                    foreach (var obj in token)
                    {
                        var newSdt = (SdtElement)prototype.CloneNode(true);
                        if (obj is JObject objFields)
                        {
                            // Recursively fill nested SDTs
                            PopulateSdtElements(newSdt, objFields);
                        }
                        parentElement.AppendChild(newSdt);
                    }
                }
            }
        }

        // Helper method to set only the first <Text> and clear the rest
        private static void SetSingleText(SdtElement sdt, string value)
        {
            var textElements = sdt.Descendants<Text>().ToList();
            if (textElements.Count > 0)
            {
                textElements[0].Text = value;
                for (int i = 1; i < textElements.Count; i++)
                    textElements[i].Text = string.Empty;
            }
        }
    }
}
