# pyreqif

Python library to read, edit, and repackage [ReqIF](https://www.omg.org/spec/ReqIF/) documents (`.reqif` / `.reqifz`), the OMG standard commonly used to exchange requirements between manufacturer and supplier (e.g. automotive lastenheft).
This version is prepared per defect for reqif from VW.
In case you have another OEM file please contact me.

Two classes:

- **`Reqif`**: a single `.reqif` file. Reads each requirement (`SPEC-OBJECT`) with its text, and two configurable fields designed for the supplier workflow: comment (`Kommentar Lieferant M`, XHTML) and status (`Status Lieferant M`, enumerated).
Allows editing them, exporting/importing via Excel (with a validation dropdown for the status, and images embedded in the cell if a resolver is provided), and detects images embedded in each requirement's text.
- **`Reqifz`**: a `.reqifz` (a zip with one or more `.reqif` files plus their attachments). Extracts to a working directory, exposes each `.reqif` as a `Reqif`, and repackages everything back while keeping the attachments intact.

## Installation

```bash
pip install pyreqifz
```

(the package name on PyPI is `pyreqifz` — `pyreqif` was already taken — but the module is still imported as `pyreqif`.)

## Basic usage

```python
from pyreqif import Reqifz

with Reqifz("lastenheft.reqifz") as pack:
    for doc in pack:
        print(doc, "->", len(doc), "requirements")

    doc = pack.get(0)  # or pack.get("file_name.reqif")

    # read
    req = doc.get("_a1b2c3...")
    print(req.text, req.comment, req.status, req.images)

    # edit (None leaves the field unchanged, "" clears it)
    doc.update("_a1b2c3...", comment="Accepted, no changes.", status="akzeptiert")

    # export to Excel, with the first image of each requirement
    # embedded in its cell ("Image" column)
    doc.to_excel("requirements.xlsx", image_resolver=pack.image_path)

    # import back (only touches comment and status)
    doc.update_from_excel("requirements_reviewed.xlsx")

    # repackage with the changes
    pack.save("lastenheft_edited.reqifz")
```

A standalone (uncompressed) `.reqif` is used the same way, without going through `Reqifz` (but then `to_excel()` can't embed images, since it doesn't know where the attachments are — pass your own `image_resolver` if you have them elsewhere):

```python
from pyreqif import Reqif

doc = Reqif("document.reqif")
doc.update("_a1b2c3...", status="Klärungsbedarf")
doc.save("document_edited.reqif")
```

## Attribute names

By default, the standard ReqIF names are used for the text (`ReqIF.Text` / `ReqIF.ChapterName` as fallback) and the field names for this specific supplier workflow (`Kommentar Lieferant M` / `Status Lieferant M`). If your document uses other names, create a subclass:

```python
from pyreqif import Reqif

class MyReqif(Reqif):
    COMMENT_ATTR = "Supplier Comment"
    STATUS_ATTR = "Supplier Status"
```

## Development

```bash
uv venv
uv pip install -e ".[dev]"
uv run pytest
```

## GitHub

https://github.com/ErnestoAvedillo/pyreqif

## Contact

Ernesto Avedillo Carretero — eavedillo@yahoo.es

## License

MIT
