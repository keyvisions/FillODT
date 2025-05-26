# FillODT.exe

**FillODT** is a command-line tool that takes an ODT (OpenDocument Text) template containing `@@placeholders` and a JSON or XML file with key-value pairs, then generates a new ODT file with all placeholders replaced by their corresponding values from the JSON or XML.

- Supports simple text, HTML fragments, and images (including QR codes).
- Handles array data for table row expansion.
- Allows image sizing and QR code generation via special placeholder syntax.
- Can output to PDF (requires LibreOffice).
- Includes options for overwriting, empty placeholder replacement, and marking incomplete documents.

[Download the latest Windows executable](https://github.com/keyvisions/FillODT/releases/latest)

---

## Real-world Usage

**FillODT** is also used as a backend tool in web servers to generate PDFs on demand for business documents such as:

- Shipping labels
- Plates
- Certificates
- Datasheets
- Picking lists
- ...and more

This makes it ideal for automated document workflows, e-commerce, logistics, and manufacturing applications where dynamic PDF generation is required.

---

## Image placeholder syntax

The `@@imageN` placeholders in your ODT template can be replaced with:

- **A local file:**  
  Specify a path to an image file on disk.
  ```json
  "image1": "./media/photo.png"
  ```

- **A URL (https):**  
  Provide an HTTPS URL to an image. The image will be downloaded automatically.
  ```json
  "image2": "https://example.com/image.jpg"
  ```

- **A QR code:**  
  Use the `qrcode://` prefix followed by the text or URL you want encoded. A QR code image will be generated and inserted.
  ```json
  "image3": "qrcode://https://example.com"
  ```

You can also use an object to specify image size:
```json
"image4": { "path": "./media/photo.png", "height": "3cm" }
```

## Run the App
$ FillODT.exe --template template.odt --json data.json --destfile output.odt

## Usage

```sh
FillODT.exe --template <template.odt> --json <data.json> --destfile <output.odt> [options]
FillODT.exe --template <template.odt> --xml <data.xml> --destfile <output.odt> [options]
```

You must specify either a JSON file (`--json data.json`) or an XML file (`--xml data.xml`) as your data source.

**Options:**
- `--overwrite`         Overwrite the output file if it exists.
- `--pdf`             Also generate a PDF (requires LibreOffice).
- `--novalue <text>`     Replace any unreplaced placeholders with the given text.

---

## Example

```sh
FillODT.exe --template template.odt --json data.json --destfile output.odt --pdf
```

See [template.odt](https://github.com/keyvisions/FillODT/blob/master/template.odt), [data.json](https://github.com/keyvisions/FillODT/blob/master/data.json) and [output.pdf](https://github.com/keyvisions/FillODT/blob/master/output.pdf)

---

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

This project was developed with the assistance of GitHub Copilot.
