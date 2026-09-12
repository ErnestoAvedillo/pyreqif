# pyreqif

Librería Python para leer, editar y reempaquetar documentos [ReqIF](https://www.omg.org/spec/ReqIF/) (`.reqif` / `.reqifz`), el estándar OMG usado habitualmente para intercambiar requisitos entre fabricante y proveedor (p.ej. lastenheft de automoción).

Dos clases:

- **`Reqif`**: un único fichero `.reqif`. Lee cada requisito (`SPEC-OBJECT`) con su texto, y dos campos configurables pensados para el flujo proveedor: comentario (`Kommentar Lieferant M`, XHTML) y estado (`Status Lieferant M`, enumerado). Permite editarlos, exportar/importar por Excel (con desplegable de validación para el estado) y detecta las imágenes embebidas en el texto de cada requisito.
- **`Reqifz`**: un `.reqifz` (zip con uno o varios `.reqif` más sus adjuntos). Extrae a un directorio de trabajo, expone cada `.reqif` como un `Reqif`, y reempaqueta todo de vuelta conservando los adjuntos intactos.

## Instalación

```bash
pip install pyreqif
```

## Uso básico

```python
from pyreqif import Reqifz

with Reqifz("lastenheft.reqifz") as pack:
    for doc in pack:
        print(doc, "->", len(doc), "requisitos")

    doc = pack.get(0)  # o pack.get("nombre_del_fichero.reqif")

    # leer
    req = doc.get("_a1b2c3...")
    print(req.text, req.comment, req.status, req.images)

    # editar (None deja el campo igual, "" lo vacía)
    doc.update("_a1b2c3...", comment="Aceptado, sin cambios.", status="akzeptiert")

    # exportar/importar por Excel
    doc.to_excel("requisitos.xlsx")
    doc.update_from_excel("requisitos_revisado.xlsx")

    # volver a empaquetar con los cambios
    pack.save("lastenheft_editado.reqifz")
```

Un `.reqif` suelto (sin comprimir) se usa igual, sin pasar por `Reqifz`:

```python
from pyreqif import Reqif

doc = Reqif("documento.reqif")
doc.update("_a1b2c3...", status="Klärungsbedarf")
doc.save("documento_editado.reqif")
```

## Nombres de atributo

Por defecto se usan los nombres estándar de ReqIF para el texto (`ReqIF.Text` / `ReqIF.ChapterName` como respaldo) y los nombres de campo de este flujo concreto de proveedor (`Kommentar Lieferant M` / `Status Lieferant M`). Si tu documento usa otros nombres, crea una subclase:

```python
from pyreqif import Reqif

class MiReqif(Reqif):
    COMMENT_ATTR = "Supplier Comment"
    STATUS_ATTR = "Supplier Status"
```

## Desarrollo

```bash
uv venv
uv pip install -e ".[dev]"
uv run pytest
```

## Licencia

MIT
