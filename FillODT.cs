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
					case "--output":
						if (i + 1 < args.Length) outputOdtFilePath = args[++i];
						break;
				}
			}

			if (string.IsNullOrEmpty(odtFilePath) || string.IsNullOrEmpty(jsonFilePath) || string.IsNullOrEmpty(outputOdtFilePath))
			{
				Console.WriteLine("Usage: FillODT --template template.odt --json data.json --output document.odt");
				return;
			}

			// Read JSON file
			var placeholders = FlattenPlaceholders(ReadJsonFile(jsonFilePath));

			// Extract ODT file
			string extractedFolder = ExtractOdt(odtFilePath);

			// Replace placeholders in content.xml
			string contentFilePath = Path.Combine(extractedFolder, "content.xml");
			ReplacePlaceholders(contentFilePath, placeholders);

			// Create new ODT file
			CreateOdtFromExtracted(extractedFolder, outputOdtFilePath);

			Console.WriteLine($"Output ODT file created at: {outputOdtFilePath}");
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

		static void ReplacePlaceholders(string filePath, Dictionary<string, object> placeholders)
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
					content = content.Replace($"@@{key}", value);
				}
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
	}
}
