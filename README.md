# FillODT.exe

**FillODT** is a command-line tool that takes an ODT (OpenDocument Text) template containing `@@placeholders` and a JSON file with key-value pairs, then generates a new ODT file with all placeholders replaced by their corresponding values from the JSON.

- Supports simple text, HTML fragments, and images (including QR codes).
- Handles array data for table row expansion.
- Allows image sizing and QR code generation via special placeholder syntax.
- Can output to PDF (requires LibreOffice).
- Includes options for overwriting, empty placeholder replacement, and marking incomplete documents.

This project was developed with the assistance of GitHub Copilot.

## Publish the Project
$ dotnet publish -c Release -r win-x64 --self-contained true

## Run the App
$ FillODT.exe --template template.odt --json data.json --destfile output.odt

## Usage

```sh
FillODT.exe --template <template.odt> --json <data.json> --destfile <output.odt> [options]
```

**Options:**

- `--overwrite` Overwrite the output file if it exists.
- `--pdf` Also generate a PDF (requires LibreOffice).
- `--novalue <text>` Replace any unreplaced placeholders with the given text.

## Examples

### 1. Basic Usage
```sh
FillODT.exe --template template.odt --json data.json --destfile output.odt
```

### 2. JSON Data Structure
```json
{
  "name": "Jane Doe",
  "image1": "qrcode://https://example.com",
  "image2": { "path": "./media/photo.png", "height": "3cm" },
  "jobs": [
    { "company": "Tech Solutions", "position": "Engineer" },
    { "company": "Web Innovations", "position": "Developer" }
  ]
}
```

---

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
