# FillODT.exe

**FillODT** is a command-line utility that processes an ODT (Open Document Text) template containing placeholders like `@@name` and `[@@name (W|*) (H|*)]`. It replaces these placeholders with values from a provided JSON or XML file, generating a new ODT file with all substitutions applied ([see example](#example)).

- Supports text, simple HTML fragments, images and QRCode generation.
- Handles array data for table row expansion.
- Allows image sizing.
- Can output to PDF (requires LibreOffice).
- Includes options for overwriting, empty placeholder replacement, and sanitizing ODT.
- If the data file contains the `incomplete` key set to `1`, the generated document's filename will have two underscores appended before the extension (for example, `document__.pdf`). This indicates that the document is incomplete.

[Download the latest Windows executable](https://github.com/keyvisions/FillODT/releases/latest)

---

## Real-world Usage

**FillODT** is used as a backend tool in web servers to generate PDFs on demand for business documents such as:

- Shipping labels
- Plates
- Certificates
- Datasheets
- Picking lists
- ...and more

This makes it ideal for automated document workflows, e-commerce, logistics, and manufacturing applications where dynamic PDF generation is required.

---

## Image placeholder syntax

To indicate that a placeholder represents an image, wrap the key in square brackets: `[@@name]`. You can control the image size using the syntax `[@@name (W|*) (H|*)]`, where `W` and `H` are the width and height in centimeters. Use `*` to automatically scale the dimension based on the image's aspect ratio. If no size is specified, the default is `[@@name * 2.54]` (height of 2.54 cm, width auto). The value for an image placeholder can be any of the following:

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

You must provide a data source using either a JSON file (`--json data.json`) or an XML file (`--xml data.xml`). Both template and data files can be specified as local paths or remote URLs.

**Options:**

| Option                   | Description                                                                                      |
|--------------------------|--------------------------------------------------------------------------------------------------|
| `--overwrite`            | Overwrite the output file if it exists.                                                          |
| `--pdf`                  | Also generate a PDF (requires LibreOffice).                                                      |
| `--novalue <text>`       | Replace any unreplaced placeholders with the given text.                                         |
| `--sanitize`             | Remove useless `<text:style>` introduced in the ODT during editing. FillODT should be run in sanitize mode on a ODT template after it has been edited. |

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
