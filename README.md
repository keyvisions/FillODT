# FillODT.exe

**FillODT** is a command-line tool that takes an ODT (Open Document Text) template containing `@@name` and `[@@image (W|*) (H|*)]` placeholders and a JSON or an XML file with key-value pairs, then generates a new ODT file with all placeholders replaced by their corresponding values as specified in the JSON or XML file.

- Supports simple text, HTML fragments, and images and QRCode generation.
- Handles array data for table row expansion.
- Allows image sizing.
- Can output to PDF (requires LibreOffice).
- Includes options for overwriting, empty placeholder replacement, and sanitizing ODT.
- If the data file contains the `incomplete` key set to `1`, the generated document's filename will have two underscores appended before the extension (for example, `document__.pdf`). This indicates that the document is incomplete.

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
The `[@@image (W|*) (H|*)]` Width and Height are in cm, when an * is used the associated dimension is set based on the aspect ratio of the image (dimensions are preset to 2 *), placeholders in your ODT template can be defined as follows:

- **A local file:**  
  Specify a path to an image file on disk.
  ```json
  "myphoto": "./media/photo.png"
  ```

- **A URL (https):**  
  Provide an HTTPS URL to an image. The image will be downloaded automatically.
  ```json
  "myimage": "https://example.com/image.jpg"
  ```

- **A QR code:**  
  Use the `qrcode://` prefix followed by the text or URL you want encoded. FillODT will generate a QR code image.
  ```json
  "myqrcode": "qrcode://https://example.com"
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

| Option                | Description                                                                                      |
|-----------------------|--------------------------------------------------------------------------------------------------|
| `--overwrite`         | Overwrite the output file if it exists.                                                          |
| `--pdf`               | Also generate a PDF (requires LibreOffice).                                                      |
| `--novalue <text>`    | Replace any unreplaced placeholders with the given text.                                         |
| `--sanitize`          | Remove useless `<text:style>` introduced in the ODT during editing. FillODT should be run in sanitize mode on a ODT template after it has been edited. |

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
