using System;
using System.Collections.Generic;
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
					case "--overwrite":
						overrideDest = true;
						break;
				}
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
				ConvertOdtToPdf(outputOdtFilePath);
				File.Delete(outputOdtFilePath);
				Console.WriteLine($"PDF created and ODT deleted: {Path.ChangeExtension(outputOdtFilePath, ".pdf")}");
			}
			else
				Console.WriteLine($"Output ODT file created at: {outputOdtFilePath}");
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
			string extractPath = Path.Combine(Path.GetTempPath(), Path.GetFileNameWithoutExtension(odtFilePath));
			if (Directory.Exists(extractPath))
				Directory.Delete(extractPath, true);
			ZipFile.ExtractToDirectory(odtFilePath, extractPath);
			return extractPath;
		}

		static void ReplacePlaceholders(string filePath, Dictionary<string, object> placeholders, string noValueReplacement) {
			string content = File.ReadAllText(filePath);

			string picturesDir = Path.Combine(Path.GetDirectoryName(filePath), "Pictures");
			if (!Directory.Exists(picturesDir))
				Directory.CreateDirectory(picturesDir);

			foreach (var placeholder in placeholders) {
				// Only process image placeholders if the placeholder exists in the template
				if (Regex.IsMatch(placeholder.Key, @"^image\d+(\.path)?$", RegexOptions.IgnoreCase)
					&& content.Contains($"@@{placeholder.Key.Split('.')[0]}")) {
					string key = placeholder.Key.Split('.')[0];
					string imagePath = placeholder.Value.ToString();
					string imageFileName;
					string destImagePath;
					string height = null;
					string width = null;

					if (placeholder.Key.Contains('.')) {
						if (placeholders.ContainsKey($"{key}.path"))
							imagePath = placeholders[$"{key}.path"].ToString();
						if (placeholders.ContainsKey($"{key}.height"))
							height = placeholders[$"{key}.height"].ToString();
						if (placeholders.ContainsKey($"{key}.width"))
							width = placeholders[$"{key}.width"].ToString();
					}

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
						TryResizeImage(destImagePath);
					}
					else if (File.Exists(imagePath)) {
						imageFileName = Path.GetFileName(imagePath);
						destImagePath = Path.Combine(picturesDir, imageFileName);
						File.Copy(imagePath, destImagePath, true);
						TryResizeImage(destImagePath);
					}
					else {
						content = content.Replace($"@@{placeholder.Key}", "[Image not found]");
						continue;
					}

					// Aspect ratio logic
					if (string.IsNullOrEmpty(width) || string.IsNullOrEmpty(height)) {
						try {
							using var img = Image.FromFile(destImagePath);
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
					string heightAttr = !string.IsNullOrEmpty(height) ? $" svg:height=\"{ConvertToOdtLength(height)}\"" : "";
					string widthAttr = !string.IsNullOrEmpty(width) ? $" svg:width=\"{ConvertToOdtLength(width)}\"" : "";

					string imageXml =
						$@"<draw:frame draw:name=""{placeholder.Key.Split('.')[0]}"" text:anchor-type=""as-char"" draw:z-index=""0""{widthAttr}{heightAttr}><draw:image xlink:href=""{odtImagePath}"" xlink:type=""simple"" xlink:show=""embed"" xlink:actuate=""onLoad""/></draw:frame>";

					content = content.Replace($"@@{key}", imageXml);
				}
			}

			// Handle array placeholders in tables
			foreach (var placeholder in placeholders) {
				if (placeholder.Value is JsonElement jsonElement && jsonElement.ValueKind == JsonValueKind.Array) {
					// Only match rows with at least one @@arrayName.something
					string rowPattern = $@"(<table:table-row\b[^>]*>[\s\S]*?</table:table-row>)";
					var rowMatches = Regex.Matches(content, rowPattern, RegexOptions.Singleline);

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

							// Replace all @@arrayName.field placeholders for this item
							foreach (var field in item) {
								filledRow = filledRow.Replace($"@@{placeholder.Key}.{field.Key}", field.Value.ToString());
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

					// Check for HTML and convert
					if (!string.IsNullOrEmpty(value) && Regex.IsMatch(value, @"<[^>]+>"))
						value = HtmlToOdtXml(value);

					content = content.Replace($"@@{key}", value);
				}
			}

			// Replace any remaining @@placeholder patterns with the noValueReplacement string
			if (!string.IsNullOrEmpty(noValueReplacement))
				content = Regex.Replace(content, @"@@[a-zA-Z0-9_.]+", noValueReplacement);

			File.WriteAllText(filePath, content);
		}

		static void CreateOdtFromExtracted(string extractedFolder, string outputOdtFilePath) {
			string mimetypePath = Path.Combine(extractedFolder, "mimetype");
			if (!File.Exists(mimetypePath))
				throw new Exception("mimetype file missing in extracted folder.");

			if (File.Exists(outputOdtFilePath))
				File.Delete(outputOdtFilePath);

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

		static void ConvertOdtToPdf(string odtPath) {
			var process = new System.Diagnostics.Process();
			process.StartInfo.FileName = "soffice";
			process.StartInfo.Arguments = $"--headless --convert-to pdf \"{odtPath}\" --outdir \"{Path.GetDirectoryName(odtPath)}\"";
			process.StartInfo.CreateNoWindow = true;
			process.StartInfo.UseShellExecute = false;
			process.StartInfo.RedirectStandardOutput = true;
			process.StartInfo.RedirectStandardError = true;
			process.Start();
			process.WaitForExit();
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
			if (double.TryParse(value, out double val))
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

		static void TryResizeImage(string imagePath, int maxWidth = 2000, int maxHeight = 2000) {
			if (imagePath.EndsWith(".svg", StringComparison.OrdinalIgnoreCase)) return;
			try {
				using var img = Image.FromFile(imagePath);
				if (img.Width > maxWidth || img.Height > maxHeight) {
					using var resized = ResizeImage(img, maxWidth, maxHeight);
					resized.Save(imagePath, img.RawFormat);
				}
			}
			catch (Exception ex) {
				Console.WriteLine($"Warning: Could not resize image {imagePath}: {ex.Message}");
			}
		}
	}
}
