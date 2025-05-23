using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace OdtPlaceholderReplacer
{
	class Program
	{
		static void Main(string[] args)
		{
			string odtFilePath = null;
			string jsonFilePath = null;
			string outputOdtFilePath = null;
			bool convertToPdf = false;
			bool overrideDest = false;
			string noValueReplacement = null;

			for (int i = 0; i < args.Length; i++)
			{
				switch (args[i])
				{
					case "--template":
						if (i + 1 < args.Length) odtFilePath = args[++i];
						break;
					case "--json":
						if (i + 1 < args.Length) jsonFilePath = args[++i];
						break;
					case "--destfile":
						if (i + 1 < args.Length) outputOdtFilePath = args[++i];
						break;
					case "--pdf":
						convertToPdf = true;
						break;
					case "--overwrite": // renamed from --override
						overrideDest = true;
						break;
					case "--novalue":
						if (i + 1 < args.Length) noValueReplacement = args[++i];
						break;
				}
			}

			// Force output file to end with .odt
			if (!string.IsNullOrEmpty(outputOdtFilePath) && !outputOdtFilePath.EndsWith(".odt", StringComparison.OrdinalIgnoreCase))
			{
				outputOdtFilePath += ".odt";
			}

			if (string.IsNullOrEmpty(odtFilePath) || string.IsNullOrEmpty(jsonFilePath) || string.IsNullOrEmpty(outputOdtFilePath))
			{
				Console.WriteLine("Usage: FillODT --template template.odt --json data.json --destfile document.odt [--pdf] [--overwrite]");
				return;
			}

			// Ensure the template file ends with .odt
			if (!odtFilePath.EndsWith(".odt", StringComparison.OrdinalIgnoreCase))
			{
				Console.WriteLine("Error: --template file must have a .odt extension.");
				return;
			}

			// Check if output file exists and handle override
			if (File.Exists(outputOdtFilePath) && !overrideDest)
			{
				Console.WriteLine($"Error: Destination file '{outputOdtFilePath}' already exists. Use --overwrite to overwrite.");
				return;
			}
			else if (File.Exists(outputOdtFilePath) && overrideDest)
			{
				File.Delete(outputOdtFilePath);
			}

			// Read JSON file
			var placeholders = FlattenPlaceholders(ReadJsonFile(jsonFilePath));

			// If "incomplete": true, append "__" to the destfile name (before .odt)
			if (placeholders.TryGetValue("incomplete", out var incompleteValue)
				&& incompleteValue is JsonElement elem
				&& elem.ValueKind == JsonValueKind.True)
			{
				if (outputOdtFilePath.EndsWith(".odt", StringComparison.OrdinalIgnoreCase))
				{
					outputOdtFilePath = outputOdtFilePath.Substring(0, outputOdtFilePath.Length - 4) + "__.odt";
				}
				else
				{
					outputOdtFilePath += "__";
				}
			}

			// Extract ODT file
			string extractedFolder = ExtractOdt(odtFilePath);

			// Replace placeholders in content.xml
			string contentFilePath = Path.Combine(extractedFolder, "content.xml");
			ReplacePlaceholders(contentFilePath, placeholders, noValueReplacement);

			// Create new ODT file
			CreateOdtFromExtracted(extractedFolder, outputOdtFilePath);

			if (convertToPdf)
			{
				ConvertOdtToPdf(outputOdtFilePath);
				File.Delete(outputOdtFilePath);
				Console.WriteLine($"PDF created and ODT deleted: {Path.ChangeExtension(outputOdtFilePath, ".pdf")}");
			}
			else
			{
				Console.WriteLine($"Output ODT file created at: {outputOdtFilePath}");
			}
		}

		static Dictionary<string, object> ReadJsonFile(string jsonFilePath)
		{
			string json = File.ReadAllText(jsonFilePath);
			return JsonSerializer.Deserialize<Dictionary<string, object>>(json);
		}

		static string ExtractOdt(string odtFilePath)
		{
			string extractPath = Path.Combine(Path.GetTempPath(), Path.GetFileNameWithoutExtension(odtFilePath));
			if (Directory.Exists(extractPath))
			{
				Directory.Delete(extractPath, true);
			}
			ZipFile.ExtractToDirectory(odtFilePath, extractPath);
			return extractPath;
		}

		static void ReplacePlaceholders(string filePath, Dictionary<string, object> placeholders, string noValueReplacement)
		{
			string content = File.ReadAllText(filePath);

			// Handle array placeholders in tables
			foreach (var placeholder in placeholders)
			{
				if (placeholder.Value is JsonElement jsonElement && jsonElement.ValueKind == JsonValueKind.Array)
				{
					// Only match rows with at least one @@arrayName.something
					string rowPattern = $@"(<table:table-row\b[^>]*>[\s\S]*?</table:table-row>)";
					var rowMatches = Regex.Matches(content, rowPattern, RegexOptions.Singleline);

					foreach (Match rowMatch in rowMatches)
					{
						string originalRow = rowMatch.Groups[1].Value;
						// Only process rows that contain the array placeholder
						if (!originalRow.Contains($"@@{placeholder.Key}."))
							continue;

						string rowsToInsert = "";

						// Deserialize array of objects
						var items = JsonSerializer.Deserialize<List<Dictionary<string, JsonElement>>>(jsonElement.GetRawText());

						foreach (var item in items)
						{
							string filledRow = originalRow;

							// Replace all @@arrayName.field placeholders for this item
							foreach (var field in item)
							{
								filledRow = filledRow.Replace($"@@{placeholder.Key}.{field.Key}", field.Value.ToString());
							}

							// Replace all other flattened placeholders in the row
							foreach (var flatPlaceholder in placeholders)
							{
								if (flatPlaceholder.Value is JsonElement flatElem &&
									flatElem.ValueKind != JsonValueKind.Array &&
									flatElem.ValueKind != JsonValueKind.Object)
								{
									filledRow = filledRow.Replace($"@@{flatPlaceholder.Key}", flatElem.ToString());
								}
							}

							rowsToInsert += filledRow;
						}

						// Replace the original row with all filled rows
						content = content.Replace(originalRow, rowsToInsert);
					}
				}
			}

			// Handle simple and flattened placeholders (including nested non-arrays)
			foreach (var placeholder in placeholders)
			{
				if (placeholder.Value is JsonElement jsonElement && jsonElement.ValueKind != JsonValueKind.Array && jsonElement.ValueKind != JsonValueKind.Object)
				{
					string key = placeholder.Key;
					string value = jsonElement.ToString();

					// Check for HTML and convert
					if (!string.IsNullOrEmpty(value) && Regex.IsMatch(value, @"<[^>]+>"))
					{
						value = HtmlToOdtXml(value);
					}

					content = content.Replace($"@@{key}", value);
				}
			}

			if (!string.IsNullOrEmpty(noValueReplacement))
			{
				// Replace any remaining @@placeholder patterns with the noValueReplacement string
				content = Regex.Replace(content, @"@@[a-zA-Z0-9_.]+", noValueReplacement);
			}

			File.WriteAllText(filePath, content);
		}

		static void CreateOdtFromExtracted(string extractedFolder, string outputOdtFilePath)
		{
			// Check if the output file already exists and delete it
			if (File.Exists(outputOdtFilePath))
			{
				File.Delete(outputOdtFilePath);
			}

			// Create the new ODT file from the extracted folder
			ZipFile.CreateFromDirectory(extractedFolder, outputOdtFilePath);
		}

		static Dictionary<string, object> FlattenPlaceholders(Dictionary<string, object> dict, string parentKey = "")
		{
			var flatDict = new Dictionary<string, object>();
			foreach (var kvp in dict)
			{
				string key = string.IsNullOrEmpty(parentKey) ? kvp.Key : $"{parentKey}.{kvp.Key}";
				if (kvp.Value is JsonElement jsonElement)
				{
					if (jsonElement.ValueKind == JsonValueKind.Object)
					{
						var nested = JsonSerializer.Deserialize<Dictionary<string, object>>(jsonElement.GetRawText());
						foreach (var nestedKvp in FlattenPlaceholders(nested, key))
						{
							flatDict[nestedKvp.Key] = nestedKvp.Value;
						}
					}
					else
					{
						flatDict[key] = jsonElement;
					}
				}
				else
				{
					flatDict[key] = kvp.Value;
				}
			}
			return flatDict;
		}

		// Add this method:
		static void ConvertOdtToPdf(string odtPath)
		{
			// Requires LibreOffice installed and in PATH
			string pdfPath = Path.ChangeExtension(odtPath, ".pdf");
			var process = new System.Diagnostics.Process();
			process.StartInfo.FileName = "soffice";
			process.StartInfo.Arguments = $"--headless --convert-to pdf \"{odtPath}\" --outdir \"{Path.GetDirectoryName(odtPath)}\"";
			process.StartInfo.CreateNoWindow = true;
			process.StartInfo.UseShellExecute = false;
			process.Start();
			process.WaitForExit();
		}
		// Converts a limited set of HTML tags to ODT XML equivalents
		static string HtmlToOdtXml(string html)
		{
			// Bold
			html = Regex.Replace(html, @"<b>(.*?)</b>", "<text:span text:style-name=\"Bold\">$1</text:span>", RegexOptions.IgnoreCase);
			html = Regex.Replace(html, @"<strong>(.*?)</strong>", "<text:span text:style-name=\"Bold\">$1</text:span>", RegexOptions.IgnoreCase);

			// Italic
			html = Regex.Replace(html, @"<i>(.*?)</i>", "<text:span text:style-name=\"Italic\">$1</text:span>", RegexOptions.IgnoreCase);
			html = Regex.Replace(html, @"<em>(.*?)</em>", "<text:span text:style-name=\"Italic\">$1</text:span>", RegexOptions.IgnoreCase);

			// Paragraphs
			html = Regex.Replace(html, @"<p>(.*?)</p>", "<text:p>$1</text:p>", RegexOptions.IgnoreCase);

			// Line breaks
			html = Regex.Replace(html, @"<br\s*/?>", "<text:line-break/>", RegexOptions.IgnoreCase);

			// Unordered list
			html = Regex.Replace(html, @"<ul>(.*?)</ul>", "<text:list>$1</text:list>", RegexOptions.IgnoreCase | RegexOptions.Singleline);
			html = Regex.Replace(html, @"<li>(.*?)</li>", "<text:list-item><text:p>$1</text:p></text:list-item>", RegexOptions.IgnoreCase);

			// Remove tags that do NOT start with <text:
			html = Regex.Replace(html, @"<(?!/?text:)[^>]+>", string.Empty);

			return html;
		}
	}
}
