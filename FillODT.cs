using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net.Http;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using QRCoder;

// dotnet publish -c Release -r win-x64 --self-contained true
namespace FillODT {
	class Program {
		static void Main(string[] args) {
			string odtFilePath = null;
			string jsonFilePath = null;
			string outputOdtFilePath = null;
			string xmlFilePath = null;
			bool convertToPdf = false;
			bool overrideDest = false;
			string noValueReplacement = null;
			bool sanitize = false;
			string printConfig = null;

			for (int i = 0; i < args.Length; i++) {
				switch (args[i]) {
					case "--template":
						if (i + 1 < args.Length) odtFilePath = args[++i];
						break;
					case "--json":
						if (i + 1 < args.Length) jsonFilePath = args[++i];
						break;
					case "--xml":
						if (i + 1 < args.Length) xmlFilePath = args[++i];
						break;
					case "--destfile":
						if (i + 1 < args.Length) outputOdtFilePath = args[++i];
						break;
					case "--novalue":
						if (i + 1 < args.Length) noValueReplacement = args[++i];
						break;
					case "--pdf":
						convertToPdf = true;
						break;
					case "--print":
						if (i + 1 < args.Length) printConfig = args[++i];
						convertToPdf = true; // Imply --pdf
						break;
					case "--overwrite":
						overrideDest = true;
						break;
					case "--sanitize":
						sanitize = true;
						break;
				}
			}

			// If --sanitize is present, only --template is required
			if (sanitize) {
				if (string.IsNullOrEmpty(odtFilePath)) {
					Console.WriteLine("Usage: FillODT --template template.odt --sanitize");
					return;
				}
				string sanitizeFolder = ExtractOdt(odtFilePath);
				if (string.IsNullOrEmpty(sanitizeFolder)) {
					Console.WriteLine($"Error accessing ODT template '{odtFilePath}': does not exist or is in use.");
					return;
				}
				SanitizeODT(sanitizeFolder);
				CreateOdtFromExtracted(sanitizeFolder, odtFilePath);
				Console.WriteLine("Sanitized ODT file by removing useless styles.");
				return;
			}

			// Force output file to end with .odt
			if (!string.IsNullOrEmpty(outputOdtFilePath) && !outputOdtFilePath.EndsWith(".odt", StringComparison.OrdinalIgnoreCase))
				outputOdtFilePath += ".odt";

			if (string.IsNullOrEmpty(odtFilePath) || (string.IsNullOrEmpty(jsonFilePath) && string.IsNullOrEmpty(xmlFilePath)) || string.IsNullOrEmpty(outputOdtFilePath)) {
				Console.WriteLine("Usage: FillODT --template template.odt --json data.json|--xml data.xml --destfile document.odt [--pdf] [--overwrite]");
				return;
			}

			// Ensure the template file ends with .odt
			if (!odtFilePath.EndsWith(".odt", StringComparison.OrdinalIgnoreCase)) {
				Console.WriteLine("Error: --template file must have a .odt extension.");
				return;
			}

			// Check if output file exists and handle override
			if (File.Exists(outputOdtFilePath) && !overrideDest) {
				Console.WriteLine($"Error: Destination file '{outputOdtFilePath}' already exists. Use --overwrite to overwrite.");
				return;
			}
			else if (File.Exists(outputOdtFilePath) && overrideDest)
				File.Delete(outputOdtFilePath);

			if (odtFilePath.StartsWith("https://", StringComparison.OrdinalIgnoreCase)) {
				string tempTemplatePath = Path.Combine(Path.GetTempPath(), "template_downloaded.odt");
				using (var httpClient = new HttpClient()) {
					try {
						var odtBytes = httpClient.GetByteArrayAsync(odtFilePath).Result;
						File.WriteAllBytes(tempTemplatePath, odtBytes);
					}
					catch (Exception) {
						Console.WriteLine($"Error downloading template {odtFilePath}");
						return;
					}
				}
				odtFilePath = tempTemplatePath;
			}

			// Download JSON or XML if given as URL
			if (!string.IsNullOrEmpty(jsonFilePath) && jsonFilePath.StartsWith("https://", StringComparison.OrdinalIgnoreCase)) {
				string tempJsonPath = Path.Combine(Path.GetTempPath(), "data_downloaded.json");
				using var httpClient = new HttpClient();
				try {
					var jsonBytes = httpClient.GetByteArrayAsync(jsonFilePath).Result;
					File.WriteAllBytes(tempJsonPath, jsonBytes);
					jsonFilePath = tempJsonPath;
				}
				catch (Exception) {
					Console.WriteLine($"Error downloading JSON {jsonFilePath}");
					return;
				}
			}
			if (!string.IsNullOrEmpty(xmlFilePath) && xmlFilePath.StartsWith("https://", StringComparison.OrdinalIgnoreCase)) {
				string tempXmlPath = Path.Combine(Path.GetTempPath(), "data_downloaded.xml");
				using var httpClient = new HttpClient();
				try {
					var xmlBytes = httpClient.GetByteArrayAsync(xmlFilePath).Result;
					File.WriteAllBytes(tempXmlPath, xmlBytes);
					xmlFilePath = tempXmlPath;
				}
				catch (Exception) {
					Console.WriteLine($"Error downloading XML {xmlFilePath}");
					return;
				}
			}

			Dictionary<string, object> placeholders = null;
			try {
				if (!string.IsNullOrEmpty(jsonFilePath)) {
					placeholders = FlattenPlaceholders(ReadJsonFile(jsonFilePath));
				}
				else if (!string.IsNullOrEmpty(xmlFilePath)) {
					var xmlDict = ReadXmlFile(xmlFilePath);
					string jsonFromXml = JsonSerializer.Serialize(xmlDict);
					placeholders = FlattenPlaceholders(JsonSerializer.Deserialize<Dictionary<string, object>>(jsonFromXml));
				}
				else {
					Console.WriteLine("Error: You must specify either --json or --xml.");
					return;
				}
			}
			catch (Exception) {
				Console.WriteLine($"Error reading or parsing the data file");
				return;
			}

			// If "incomplete": true, append "__" to the destfile name (before .odt)
			if (placeholders.TryGetValue("incomplete", out var incompleteValue)
				&& incompleteValue is JsonElement elem
				&& elem.ValueKind == JsonValueKind.True) {
				if (outputOdtFilePath.EndsWith(".odt", StringComparison.OrdinalIgnoreCase))
					outputOdtFilePath = outputOdtFilePath.Substring(0, outputOdtFilePath.Length - 4) + "__.odt";
				else
					outputOdtFilePath += "__";
			}

			// Extract ODT file
			string extractedFolder = ExtractOdt(odtFilePath);
			if (string.IsNullOrEmpty(extractedFolder)) {
				Console.WriteLine($"Error accessing ODT template '{odtFilePath}': does not exists or is in use.");
				return;
			}

			// Replace placeholders in content.xml
			string contentFilePath = Path.Combine(extractedFolder, "content.xml");
			ReplacePlaceholders(contentFilePath, placeholders, noValueReplacement);

			// Replace placeholders in styles.xml
			string stylesFilePath = Path.Combine(extractedFolder, "styles.xml");
			if (File.Exists(stylesFilePath))
				ReplacePlaceholders(stylesFilePath, placeholders, noValueReplacement);

			// Update manifest with images
			UpdateManifestWithImages(extractedFolder);

			// Create new ODT file
			CreateOdtFromExtracted(extractedFolder, outputOdtFilePath);

			if (convertToPdf) {
				if (ConvertOdtToPdf(outputOdtFilePath)) {
					File.Delete(outputOdtFilePath);
					string pdfPath = Path.ChangeExtension(outputOdtFilePath, ".pdf");

					bool shouldOptimize = false;
					if (!string.IsNullOrEmpty(printConfig) && IsThermalPrinter(printConfig)) {
						shouldOptimize = true;
						Console.WriteLine($"Thermal printer detected: {printConfig}, optimizing PDF.");
					}

					if (shouldOptimize) {
						if (OptimizePdfWithGhostscript(pdfPath))
							Console.WriteLine($"Thermal-optimized PDF ready: {pdfPath}");
						else
							Console.WriteLine("Ghostscript optimization failed.");
					}
					else {
						Console.WriteLine($"Standard PDF ready: {pdfPath}");
					}

					// Print dispatch
					if (!string.IsNullOrEmpty(printConfig)) {
						bool printOk = PrintPdf(pdfPath, printConfig, shouldOptimize);
						Console.WriteLine(printOk
							? $"PDF sent to printer '{printConfig}'."
							: $"Failed to print to '{printConfig}'.");
					}
				}
				else {
					Console.WriteLine($"PDF conversion failed, ODT retained: {outputOdtFilePath}");
				}
			}

		}

		static Dictionary<string, object> ReadJsonFile(string jsonFilePath) {
			string json = File.ReadAllText(jsonFilePath);
			return JsonSerializer.Deserialize<Dictionary<string, object>>(json);
		}

		static Dictionary<string, object> ReadXmlFile(string xmlFilePath) {
			var dict = new Dictionary<string, object>();
			var doc = XDocument.Load(xmlFilePath);

			foreach (var el in doc.Root.Elements()) {
				// Handle repeated elements as arrays
				var siblings = doc.Root.Elements(el.Name).ToList();
				if (siblings.Count > 1) {
					// If there are multiple elements with the same name, treat as array
					if (!dict.ContainsKey(el.Name.LocalName))
						dict[el.Name.LocalName] = new List<Dictionary<string, object>>();

					var list = (List<Dictionary<string, object>>)dict[el.Name.LocalName];
					list.Add(XmlElementToDict(el));
				}
				else if (el.HasElements) {
					// If element has children, recurse
					dict[el.Name.LocalName] = el.Elements().All(e => e.Name == el.Elements().First().Name)
						? el.Elements().Select(XmlElementToDict).ToList()
						: XmlElementToDict(el);
				}
				else {
					dict[el.Name.LocalName] = el.Value;
				}
			}
			return dict;
		}

		static Dictionary<string, object> XmlElementToDict(XElement el) {
			var dict = new Dictionary<string, object>();
			foreach (var child in el.Elements()) {
				var siblings = el.Elements(child.Name).ToList();
				if (siblings.Count > 1) {
					if (!dict.ContainsKey(child.Name.LocalName))
						dict[child.Name.LocalName] = new List<Dictionary<string, object>>();
					var list = (List<Dictionary<string, object>>)dict[child.Name.LocalName];
					list.Add(XmlElementToDict(child));
				}
				else if (child.HasElements) {
					dict[child.Name.LocalName] = XmlElementToDict(child);
				}
				else {
					dict[child.Name.LocalName] = child.Value;
				}
			}
			return dict;
		}

		static string ExtractOdt(string odtFilePath) {
			try {
				string extractPath = Path.Combine(Path.GetTempPath(), Path.GetFileNameWithoutExtension(odtFilePath));
				if (Directory.Exists(extractPath))
					Directory.Delete(extractPath, true);
				ZipFile.ExtractToDirectory(odtFilePath, extractPath);
				return extractPath;
			}
			catch (Exception) {
				return null;
			}
		}

		static void ReplacePlaceholders(string filePath, Dictionary<string, object> placeholders, string noValueReplacement) {
			string content = File.ReadAllText(filePath);

			string picturesDir = Path.Combine(Path.GetDirectoryName(filePath), "Pictures");
			if (!Directory.Exists(picturesDir))
				Directory.CreateDirectory(picturesDir);

			// Remove annotation regions if placeholder does not resolve to true
			content = Regex.Replace(content,
				@"(<office:annotation\b[^>]*>[\s\S]*?@@([a-zA-Z0-9_.]+)[\s\S]*?</office:annotation>)([\s\S]*?)(<office:annotation-end\b[^>]*>)",
				match => {
					string annotation = match.Groups[1].Value;
					string placeholderKey = match.Groups[2].Value;
					string region = match.Groups[3].Value;
					string annotationEnd = match.Groups[4].Value;

					// Try to resolve the placeholder to a boolean true
					bool isTrue = false;
					if (placeholders.TryGetValue(placeholderKey, out var val)) {
						if (val is JsonElement je) {
							isTrue = (je.ValueKind == JsonValueKind.True) ||
								(je.ValueKind == JsonValueKind.Number && je.GetInt32() == 1) ||
								(je.ValueKind == JsonValueKind.String && je.GetString()?.ToLowerInvariant() == "true");
						}
						else if (val is bool b) {
							isTrue = b;
						}
						else if (val is int i) {
							isTrue = i == 1;
						}
						else if (val is string s) {
							isTrue = s.ToLowerInvariant() == "true";
						}
					}
					// If not true, remove the annotation and region
					return isTrue ? match.Value : "@@removeRegion";
				},
				RegexOptions.IgnoreCase | RegexOptions.Singleline);

			// Remove only the table-row that contains @@removeRegion and has no other content
			content = Regex.Replace(
				content,
				@"<table:table-row\b[^>]*>[\s\S]*?<\/table:table-row>",
				m => {
					// Remove if contains @@removeRegion and, aside from that, the inner text is empty/whitespace
					if (m.Value.Contains("@@removeRegion")) {
						// Remove @@removeRegion and check if the rest is empty/whitespace
						var inner = Regex.Replace(m.Value, "@@removeRegion", "", RegexOptions.IgnoreCase);
						// Remove all tags to get inner text
						var innerText = Regex.Replace(inner, "<.*?>", "").Trim();
						if (string.IsNullOrEmpty(innerText))
							return "";
					}
					return m.Value;
				},
				RegexOptions.IgnoreCase
			);
			// Remove all @@removeRegion placeholders left in the content
			content = Regex.Replace(content, "@@removeRegion", "", RegexOptions.IgnoreCase);

			// Process image placeholders [@@placeholderName width height]
			foreach (var placeholder in placeholders) {
				var imagePattern = new Regex($"\\[\\s*@@{placeholder.Key}(?>\\s+(\\d+(?>\\.\\d+)?|\\*))?(?>\\s+(\\d+(?>\\.\\d+)?|\\*))?\\s*\\]", RegexOptions.Compiled);
				foreach (Match image in imagePattern.Matches(content)) {
					string key = placeholder.Key;
					string imagePath = placeholder.Value.ToString();
					string imageFileName = null;
					string destImagePath = null;
					string width = image.Groups[1].Value == "*" ? "" : image.Groups[1].Value;
					string height = image.Groups[2].Value == "*" ? "" : image.Groups[2].Value;

					// Download image if it's an HTTPS URL
					if (imagePath.StartsWith("https://", StringComparison.OrdinalIgnoreCase)) {
						using var httpClient = new HttpClient();
						var imageBytes = httpClient.GetByteArrayAsync(imagePath).Result;
						imageFileName = $"{key}_{Path.GetFileName(new Uri(imagePath).AbsolutePath)}";
						destImagePath = Path.Combine(picturesDir, imageFileName);
						File.WriteAllBytes(destImagePath, imageBytes);
						TryResizeImage(destImagePath);
					}
					else if (imagePath.StartsWith("qrcode://", StringComparison.OrdinalIgnoreCase)) {
						using QRCoder.QRCodeGenerator qrGenerator = new QRCoder.QRCodeGenerator();
						using QRCoder.QRCodeData qrCodeData = qrGenerator.CreateQrCode(imagePath[9..], QRCoder.QRCodeGenerator.ECCLevel.Q);
						using QRCoder.BitmapByteQRCode qrCode = new QRCoder.BitmapByteQRCode(qrCodeData);
						using MemoryStream ms = new(qrCode.GetGraphic(20));
						using Bitmap qrCodeImage = new Bitmap(ms);
						imageFileName = $"{key}_qrcode.png";
						destImagePath = Path.Combine(picturesDir, imageFileName);
						qrCodeImage.Save(destImagePath, System.Drawing.Imaging.ImageFormat.Png);
					}
					else if (File.Exists(imagePath)) {
						imageFileName = Path.GetFileName(imagePath);
						destImagePath = Path.Combine(picturesDir, imageFileName);
						File.Copy(imagePath, destImagePath, true);
						TryResizeImage(destImagePath);
					}

					string imageXml = "";
					if (imageFileName != null && destImagePath != null) {
						// Aspect ratio logic
						if (string.IsNullOrEmpty(width) || string.IsNullOrEmpty(height)) {
							try {
								using var img = Image.FromFile(destImagePath);

								if (string.IsNullOrEmpty(width) && string.IsNullOrEmpty(height))
									height = "1in";

								double aspect = (double)img.Width / img.Height;
								double? widthCm = null, heightCm = null;

								if (!string.IsNullOrEmpty(width))
									widthCm = ParseToCm(width);
								if (!string.IsNullOrEmpty(height))
									heightCm = ParseToCm(height);

								if (widthCm.HasValue && !heightCm.HasValue) {
									heightCm = widthCm / aspect;
									height = $"{heightCm.Value.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture)}cm";
								}
								else if (!widthCm.HasValue && heightCm.HasValue) {
									widthCm = heightCm * aspect;
									width = $"{widthCm.Value.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture)}cm";
								}
							}
							catch (Exception) {
								Console.WriteLine($"Aspect ratio not applicable {destImagePath}");
							}
						}

						string odtImagePath = "Pictures/" + imageFileName;
						string heightAttr = !string.IsNullOrEmpty(height) ? $" svg:height=\"{ConvertToOdtLength(height).Replace(",", ".")}\"" : "";
						string widthAttr = !string.IsNullOrEmpty(width) ? $" svg:width=\"{ConvertToOdtLength(width).Replace(",", ".")}\"" : "";

						imageXml =
							$@"<draw:frame draw:name=""{placeholder.Key.Split('.')[0]}"" text:anchor-type=""as-char"" draw:z-index=""0""{widthAttr}{heightAttr}><draw:image xlink:href=""{odtImagePath}"" xlink:type=""simple"" xlink:show=""embed"" xlink:actuate=""onLoad""/></draw:frame>";
					}
					content = content.Replace(image.Groups[0].Value, imageXml);
				}
			}

			// Handle array placeholders in tables
			foreach (var placeholder in placeholders) {
				if (placeholder.Value is JsonElement jsonElement && jsonElement.ValueKind == JsonValueKind.Array) {
					// Only match rows with at least one @@arrayName.something
					string rowPattern = $@"(<table:table-row\b[^>]*>[\s\S]*?</table:table-row>)";
					var rowMatches = Regex.Matches(content, rowPattern, RegexOptions.Singleline);

					int row = 0;
					foreach (Match rowMatch in rowMatches) {
						string originalRow = rowMatch.Groups[1].Value;
						// Only process rows that contain the array placeholder
						if (!originalRow.Contains($"@@{placeholder.Key}."))
							continue;

						string rowsToInsert = "";

						// Deserialize array of objects
						var items = JsonSerializer.Deserialize<List<Dictionary<string, JsonElement>>>(jsonElement.GetRawText());

						foreach (var item in items) {
							string filledRow = originalRow;

							// Replace image placeholders in the row
							filledRow = ReplaceImagePlaceholders(
								filledRow,
								key => item.TryGetValue(key.Split('.').Last(), out var val) ? val.ToString() : null,
								picturesDir,
								filePath,
								++row
							);

							// Replace all @@arrayName.field placeholders for this item
							foreach (var field in item) {
								string valueToInsert = field.Value.ToString();

								// Escape special XML characters unless it's simple HTML
								if (!string.IsNullOrEmpty(valueToInsert)) {
									if (Regex.IsMatch(valueToInsert, @"<[a-z][\s\S]*>")) { // Basic HTML check
										valueToInsert = HtmlToOdtXml(valueToInsert);
									}
									else {
										valueToInsert = System.Security.SecurityElement.Escape(valueToInsert);
									}
								}
								filledRow = filledRow.Replace($"@@{placeholder.Key}.{field.Key}", valueToInsert);
							}

							// Replace all other flattened placeholders in the row
							foreach (var flatPlaceholder in placeholders) {
								if (flatPlaceholder.Value is JsonElement flatElem &&
									flatElem.ValueKind != JsonValueKind.Array &&
									flatElem.ValueKind != JsonValueKind.Object) {
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
			foreach (var placeholder in placeholders) {
				if (placeholder.Value is JsonElement jsonElement && jsonElement.ValueKind != JsonValueKind.Array && jsonElement.ValueKind != JsonValueKind.Object) {
					string key = placeholder.Key;
					string value = jsonElement.ToString();

					// Check for HTML and convert, otherwise escape for XML
					if (!string.IsNullOrEmpty(value)) {
						if (Regex.IsMatch(value, @"<[a-z][\s\S]*>")) { // Basic HTML check
							value = HtmlToOdtXml(value);
						}
						else {
							value = System.Security.SecurityElement.Escape(value);
						}
					}

					content = content.Replace($"@@{key}", value);
				}
			}

			// Remove any leftover @@image placeholders
			var leftoverImages = new Regex($"\\[\\s*@@[a-zA-Z_0-9.]+(?>\\s+(\\d+(\\.\\d+)?|\\*))?(?>\\s+(\\d+(\\.\\d+)?|\\*))?\\s*\\]", RegexOptions.Compiled);
			leftoverImages.Matches(content).ToList().ForEach(m => {
				content = content.Replace(m.Groups[0].Value, "");
			});

			// Replace any remaining @@placeholder patterns with the noValueReplacement string
			if (!string.IsNullOrEmpty(noValueReplacement))
				content = Regex.Replace(content, @"@@[a-zA-Z0-9_.]+", noValueReplacement);

			// Ensure the destination directory exists before writing
			if (Path.GetDirectoryName(filePath) != "")
				Directory.CreateDirectory(Path.GetDirectoryName(filePath));
			File.WriteAllText(filePath, content);
		}

		static void CreateOdtFromExtracted(string extractedFolder, string outputOdtFilePath) {
			string mimetypePath = Path.Combine(extractedFolder, "mimetype");
			if (!File.Exists(mimetypePath))
				throw new Exception("mimetype file missing in extracted folder.");

			if (File.Exists(outputOdtFilePath))
				File.Delete(outputOdtFilePath);

			// Ensure the destination directory exists before writing
			if (Path.GetDirectoryName(outputOdtFilePath) != "")
				Directory.CreateDirectory(Path.GetDirectoryName(outputOdtFilePath));
			using var zip = new FileStream(outputOdtFilePath, FileMode.Create);
			using var archive = new ZipArchive(zip, ZipArchiveMode.Create);
			// Add mimetype file first, uncompressed
			var mimetypeEntry = archive.CreateEntry("mimetype", CompressionLevel.NoCompression);
			using (var entryStream = mimetypeEntry.Open())
			using (var fileStream = File.OpenRead(mimetypePath))
				fileStream.CopyTo(entryStream);

			// Add all other files and folders, compressed
			foreach (var file in Directory.GetFiles(extractedFolder, "*", SearchOption.AllDirectories)) {
				string relPath = Path.GetRelativePath(extractedFolder, file).Replace("\\", "/");
				if (relPath == "mimetype") continue; // already added

				archive.CreateEntryFromFile(file, relPath, CompressionLevel.Optimal);
			}
		}

		static Dictionary<string, object> FlattenPlaceholders(Dictionary<string, object> dict, string parentKey = "") {
			var flatDict = new Dictionary<string, object>();
			foreach (var kvp in dict) {
				string key = string.IsNullOrEmpty(parentKey) ? kvp.Key : $"{parentKey}.{kvp.Key}";
				if (kvp.Value is JsonElement jsonElement) {
					if (jsonElement.ValueKind == JsonValueKind.Object) {
						var nested = JsonSerializer.Deserialize<Dictionary<string, object>>(jsonElement.GetRawText());
						foreach (var nestedKvp in FlattenPlaceholders(nested, key))
							flatDict[nestedKvp.Key] = nestedKvp.Value;
					}
					else
						flatDict[key] = jsonElement;
				}
				else
					flatDict[key] = kvp.Value;
			}
			return flatDict;
		}

		static bool ConvertOdtToPdf(string odtPath) {
			var process = new System.Diagnostics.Process();
			process.StartInfo.FileName = "soffice";
			process.StartInfo.Arguments = $"--headless --convert-to pdf \"{odtPath}\" --outdir \"{Path.GetDirectoryName(odtPath)}\"";
			process.StartInfo.CreateNoWindow = true;
			process.StartInfo.UseShellExecute = false;
			process.StartInfo.RedirectStandardOutput = true;
			process.StartInfo.RedirectStandardError = true;
			try {
				process.Start();
				string error = process.StandardError.ReadToEnd();
				process.WaitForExit();

				if (process.ExitCode != 0) {
					Console.WriteLine("Error during PDF conversion.");
					if (!string.IsNullOrEmpty(error)) {
						Console.WriteLine($"Details: {error}");
					}
					return false;
				}

				string pdfPath = Path.ChangeExtension(odtPath, ".pdf");
				return File.Exists(pdfPath);
			}
			catch (Exception ex) {
				Console.WriteLine("Error running LibreOffice. Is it installed and in your system's PATH?");
				Console.WriteLine($"Details: {ex.Message}");
				return false;
			}
		}

		// Converts a limited set of HTML tags to ODT XML equivalents
		static string HtmlToOdtXml(string html) {
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

		static string ConvertToOdtLength(string value) {
			if (string.IsNullOrEmpty(value)) return null;
			value = value.Trim().ToLowerInvariant();
			if (value.EndsWith("px")) {
				if (double.TryParse(value.Replace("px", ""), out double px)) {
					double cm = px * 0.0352778;
					return $"{cm:0.###}cm";
				}
			}
			// Allow cm, mm, in, pt as-is
			if (value.EndsWith("cm") || value.EndsWith("mm") || value.EndsWith("in") || value.EndsWith("pt"))
				return value;
			// Default: treat as cm
			if (double.TryParse(value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double val))
				return $"{val}cm";
			return value;
		}

		static void UpdateManifestWithImages(string extractedFolder) {
			string manifestPath = Path.Combine(extractedFolder, "META-INF", "manifest.xml");
			if (!File.Exists(manifestPath)) return;

			XNamespace manifest = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
			var doc = XDocument.Load(manifestPath);

			// Get all image files in Pictures/
			string picturesDir = Path.Combine(extractedFolder, "Pictures");
			if (!Directory.Exists(picturesDir)) return;
			var imageFiles = Directory.GetFiles(picturesDir);

			// Get the root <manifest:manifest> element
			var root = doc.Root;
			foreach (var imagePath in imageFiles) {
				string relPath = "Pictures/" + Path.GetFileName(imagePath);
				string mediaType = GetMediaTypeFromExtension(Path.GetExtension(imagePath));

				// Check if already present
				bool exists = root.Elements(manifest + "file-entry")
					.Any(e => (string)e.Attribute(manifest + "full-path") == relPath);

				if (!exists)
					root.Add(new XElement(manifest + "file-entry",
						new XAttribute(manifest + "media-type", mediaType),
						new XAttribute(manifest + "full-path", relPath)
					));
			}
			doc.Save(manifestPath);
		}

		static string GetMediaTypeFromExtension(string ext) {
			switch (ext.ToLowerInvariant()) {
				case ".png": return "image/png";
				case ".jpg":
				case ".jpeg": return "image/jpeg";
				case ".gif": return "image/gif";
				case ".svg": return "image/svg+xml";
				case ".bmp": return "image/bmp";
				default: return "application/octet-stream";
			}
		}

		static double ParseToCm(string value) {
			value = value.Trim().ToLowerInvariant();
			if (value.EndsWith("cm"))
				return double.Parse(value.Replace("cm", ""), System.Globalization.CultureInfo.InvariantCulture);
			if (value.EndsWith("mm"))
				return double.Parse(value.Replace("mm", ""), System.Globalization.CultureInfo.InvariantCulture) / 10.0;
			if (value.EndsWith("in"))
				return double.Parse(value.Replace("in", ""), System.Globalization.CultureInfo.InvariantCulture) * 2.54;
			if (value.EndsWith("pt"))
				return double.Parse(value.Replace("pt", ""), System.Globalization.CultureInfo.InvariantCulture) * 0.0352778;
			if (value.EndsWith("px"))
				return double.Parse(value.Replace("px", ""), System.Globalization.CultureInfo.InvariantCulture) * 0.0352778;
			// Default: treat as cm
			return double.Parse(value, System.Globalization.CultureInfo.InvariantCulture);
		}

		static Bitmap ResizeImage(Image image, int maxWidth, int maxHeight) {
			double ratioX = (double)maxWidth / image.Width;
			double ratioY = (double)maxHeight / image.Height;
			double ratio = Math.Min(ratioX, ratioY);

			int newWidth = (int)(image.Width * ratio);
			int newHeight = (int)(image.Height * ratio);

			var newImage = new Bitmap(newWidth, newHeight);
			using (var graphics = Graphics.FromImage(newImage)) {
				graphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
				graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
				graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
				graphics.DrawImage(image, 0, 0, newWidth, newHeight);
			}
			return newImage;
		}

		static void TryResizeImage(string imagePath, int maxWidth = 1024, int maxHeight = 1024) {
			if (imagePath.EndsWith(".svg", StringComparison.OrdinalIgnoreCase))
				return;

			try {
				using var fs = new FileStream(imagePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
				using var img = Image.FromStream(fs);

				// Fix orientation based on EXIF data
				const int ExifOrientationId = 0x0112;
				if (img.PropertyIdList.Contains(ExifOrientationId)) {
					var prop = img.GetPropertyItem(ExifOrientationId);
					int orientation = prop.Value[0];
					RotateFlipType rotateFlip = RotateFlipType.RotateNoneFlipNone;
					switch (orientation) {
						case 2: rotateFlip = RotateFlipType.RotateNoneFlipX; break;
						case 3: rotateFlip = RotateFlipType.Rotate180FlipNone; break;
						case 4: rotateFlip = RotateFlipType.Rotate180FlipX; break;
						case 5: rotateFlip = RotateFlipType.Rotate90FlipX; break;
						case 6: rotateFlip = RotateFlipType.Rotate90FlipNone; break;
						case 7: rotateFlip = RotateFlipType.Rotate270FlipX; break;
						case 8: rotateFlip = RotateFlipType.Rotate270FlipNone; break;
					}
					if (rotateFlip != RotateFlipType.RotateNoneFlipNone)
						img.RotateFlip(rotateFlip);
					// Remove orientation property to prevent re-rotation
					img.RemovePropertyItem(ExifOrientationId);
				}

				if (img.Width > maxWidth || img.Height > maxHeight) {
					using var resized = ResizeImage(img, maxWidth, maxHeight);
					resized.Save(imagePath, img.RawFormat);
				}
			}
			catch (Exception ex) {
				Console.WriteLine($"Warning: Could not resize image {imagePath}: {ex.GetType().Name}: {ex.Message}\n{ex.StackTrace}");
			}
		}

		static void SanitizeODT(string extractPath) {
			IdentifyAndRemoveUselessStyles(extractPath, "styles.xml");
			IdentifyAndRemoveUselessStyles(extractPath, "content.xml");
		}
		static void IdentifyAndRemoveUselessStyles(string extractPath, string filename) {
			string filePath = Path.Combine(extractPath, filename);
			string fileContent = File.ReadAllText(filePath);

			// Find all styles and their properties
			var stylePattern = new Regex(@"<style:style style:name=""(\w*?)"" style:family=""text"">(.*?)<\/style:style>", RegexOptions.Singleline);
			var matches = stylePattern.Matches(fileContent);

			// Identify useless styles
			var uselessStyles = new List<string>();
			foreach (Match match in matches) {
				string styleName = match.Groups[1].Value;
				string properties = match.Groups[2].Value;

				// Check if properties are empty or contain only default attributes
				if (Regex.IsMatch(properties, @"<style:text-properties\s*\/>") ||
					!Regex.IsMatch(properties, @"text:font-weight|text:font-size|text:color")) {
					uselessStyles.Add(styleName);
				}
			}

			// Create a pattern to remove spans with useless styles
			string spanPattern = $@"<text:span text:style-name=""({string.Join("|", uselessStyles)})"">(.*?)<\/text:span>";

			File.WriteAllText(filePath, Regex.Replace(fileContent, spanPattern, "$2"));
		}

		static string ReplaceImagePlaceholders(string input, Func<string, string> imagePathResolver, string picturesDir, string filePath, int row) {
			return Regex.Replace(input, @"\[\s*@@([a-zA-Z0-9_.]+)(?>\s+([0-9.*]+))?(?>\s+([0-9.*]+))?\s*\]", match => {
				string key = match.Groups[1].Value;
				string widthStr = match.Groups[2].Value;
				string heightStr = match.Groups[3].Value;

				string imagePath = imagePathResolver(key);
				if (string.IsNullOrEmpty(imagePath))
					return "";

				if (!Directory.Exists(picturesDir))
					Directory.CreateDirectory(picturesDir);

				string imageFileName;
				string destImagePath;

				// Download/copia/genera l'immagine
				if (imagePath.StartsWith("https://", StringComparison.OrdinalIgnoreCase)) {
					using var httpClient = new HttpClient();
					var imageBytes = httpClient.GetByteArrayAsync(imagePath).Result;
					imageFileName = $"{key}_{Path.GetFileName(new Uri(imagePath).AbsolutePath)}";
					destImagePath = Path.Combine(picturesDir, imageFileName);
					File.WriteAllBytes(destImagePath, imageBytes);
					TryResizeImage(destImagePath);
				}
				else if (imagePath.StartsWith("qrcode://", StringComparison.OrdinalIgnoreCase)) {
					using QRCoder.QRCodeGenerator qrGenerator = new QRCoder.QRCodeGenerator();
					using QRCoder.QRCodeData qrCodeData = qrGenerator.CreateQrCode(imagePath[9..], QRCoder.QRCodeGenerator.ECCLevel.Q);
					using QRCoder.BitmapByteQRCode qrCode = new QRCoder.BitmapByteQRCode(qrCodeData);
					using MemoryStream ms = new(qrCode.GetGraphic(20));
					using Bitmap qrCodeImage = new Bitmap(ms);
					imageFileName = $"{key}{row}_qrcode.png";
					destImagePath = Path.Combine(picturesDir, imageFileName);
					qrCodeImage.Save(destImagePath, System.Drawing.Imaging.ImageFormat.Png);
				}
				else if (File.Exists(imagePath)) {
					imageFileName = Path.GetFileName(imagePath);
					destImagePath = Path.Combine(picturesDir, imageFileName);
					File.Copy(imagePath, destImagePath, true);
					TryResizeImage(destImagePath);
				}
				else {
					return "";
				}

				// Calcola width/height in cm
				string width = null, height = null;
				double? widthCm = null, heightCm = null;
				try {
					using var img = Image.FromFile(destImagePath);
					double aspect = (double)img.Width / img.Height;

					if (widthStr != "*" && double.TryParse(widthStr, out double w))
						widthCm = w;
					if (heightStr != "*" && double.TryParse(heightStr, out double h))
						heightCm = h;

					if (widthCm.HasValue && !heightCm.HasValue) {
						heightCm = widthCm / aspect;
					}
					else if (!widthCm.HasValue && heightCm.HasValue) {
						widthCm = heightCm * aspect;
					}
					else if (!widthCm.HasValue && !heightCm.HasValue) {
						widthCm = img.Width * 0.0352778;
						heightCm = img.Height * 0.0352778;
					}

					width = $"{widthCm.Value.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture)}cm";
					height = $"{heightCm.Value.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture)}cm";
				}
				catch (Exception) {
					// fallback: no width/height
				}

				string odtImagePath = "Pictures/" + imageFileName;
				string heightAttr = !string.IsNullOrEmpty(height) ? $" svg:height=\"{ConvertToOdtLength(height)}\"" : "";
				string widthAttr = !string.IsNullOrEmpty(width) ? $" svg:width=\"{ConvertToOdtLength(width)}\"" : "";

				return $@"<draw:frame draw:name=""{key}"" text:anchor-type=""as-char"" draw:z-index=""0""{widthAttr}{heightAttr}><draw:image xlink:href=""{odtImagePath}"" xlink:type=""simple"" xlink:show=""embed"" xlink:actuate=""onLoad""/></draw:frame>";
			});
		}

		static bool OptimizePdfWithGhostscript(string pdfPath) {
			string optimizedPath = Path.Combine(Path.GetDirectoryName(pdfPath),
										Path.GetFileNameWithoutExtension(pdfPath) + "_optimized.pdf");

			var gsArgs = $"-dNOPAUSE -dBATCH -sDEVICE=pdfwrite " +
							 "-dCompatibilityLevel=1.4 -dPDFSETTINGS=/screen " +
							 $"-sOutputFile=\"{optimizedPath}\" \"{pdfPath}\"";

			var process = new Process {
				StartInfo = new ProcessStartInfo {
					FileName = "gs",
					Arguments = gsArgs,
					UseShellExecute = false,
					CreateNoWindow = true,
					RedirectStandardOutput = true,
					RedirectStandardError = true,
				}
			};

			try {
				process.Start();
				string output = process.StandardOutput.ReadToEnd();
				string error = process.StandardError.ReadToEnd();
				process.WaitForExit();

				if (process.ExitCode == 0 && File.Exists(optimizedPath)) {
					File.Delete(pdfPath); // Optionally replace original
					File.Move(optimizedPath, pdfPath);
					return true;
				}

				Console.WriteLine("Ghostscript Error: " + error);
				return false;
			}
			catch (Exception ex) {
				Console.WriteLine("Ghostscript execution failed: " + ex.Message);
				return false;
			}
		}
		static bool IsThermalPrinter(string printerName) {
			try {
				using var searcher = new System.Management.ManagementObjectSearcher("SELECT * FROM Win32_Printer");
				foreach (var printer in searcher.Get()) {
					string name = printer["Name"]?.ToString();
					if (name?.Equals(printerName, StringComparison.OrdinalIgnoreCase) == true) {
						// Heuristics: look at model name for "Zebra", "Bixolon", etc.
						string description = printer["DriverName"]?.ToString() ?? "";
						if (description.ToLowerInvariant().Contains("zebra") ||
							 description.ToLowerInvariant().Contains("bixolon") ||
							 description.ToLowerInvariant().Contains("datamax") ||
							 description.ToLowerInvariant().Contains("sato") ||
							 description.ToLowerInvariant().Contains("tsp")) {
							return true;
						}
					}
				}
			}
			catch (Exception ex) {
				Console.WriteLine("Printer detection failed: " + ex.Message);
			}

			return false;
		}

		static bool PrintPdf(string pdfPath, string printerName, bool isThermal)
		{
			try
			{
				var process = new Process();
				if (isThermal)
				{
					// Ghostscript: send PDF to Windows printer using mswinpr2 device
					process.StartInfo.FileName = "gswin64c.exe"; // or "gs" if in PATH

					// Paper size configuration
					string paperSize = "4x6"; // from printers.json
					int widthPoints = 288, heightPoints = 432; // defaults

					if (Regex.IsMatch(paperSize, @"^(\d+(\.\d+)?)x(\d+(\.\d+)?)$")) {
						var parts = paperSize.Split('x');
						double widthInches = double.Parse(parts[0], System.Globalization.CultureInfo.InvariantCulture);
						double heightInches = double.Parse(parts[1], System.Globalization.CultureInfo.InvariantCulture);
						widthPoints = (int)(widthInches * 72);
						heightPoints = (int)(heightInches * 72);
					}

					// Then build Ghostscript args:
					process.StartInfo.Arguments =
						$"-dPrinted -dNOPAUSE -dBATCH -sDEVICE=mswinpr2 -sPAPERSIZE=custom -dDEVICEWIDTHPOINTS={widthPoints} -dDEVICEHEIGHTPOINTS={heightPoints} -sOutputFile=\"\\\\spool\\{printerName}\" \"{pdfPath}\"";
				}
				else
				{
					// LibreOffice: print PDF to printer
					process.StartInfo.FileName = "soffice";
					process.StartInfo.Arguments = $"--headless --pt \"{printerName}\" \"{pdfPath}\"";
				}
				process.StartInfo.CreateNoWindow = true;
				process.StartInfo.UseShellExecute = false;
				process.StartInfo.RedirectStandardOutput = true;
				process.StartInfo.RedirectStandardError = true;

				process.Start();
				string output = process.StandardOutput.ReadToEnd();
				string error = process.StandardError.ReadToEnd();
				process.WaitForExit();

				if (process.ExitCode != 0)
				{
					Console.WriteLine($"Print error: {error}");
					return false;
				}
				return true;
			}
			catch (Exception ex)
			{
				Console.WriteLine($"Print dispatch failed: {ex.Message}");
				return false;
			}
		}
	}
}