"""Reqifz: apertura, edición y reempaquetado de un .reqifz completo."""
from __future__ import annotations

import shutil
import tempfile
import zipfile
from pathlib import Path

from .reqif import Reqif


class Reqifz:
    """Un .reqifz es un zip con uno o varios .reqif más los ficheros
    adjuntos que referencian (imágenes, documentos...). Esta clase lo
    extrae a un directorio de trabajo, expone cada .reqif como un
    `Reqif`, y permite reempaquetarlo todo de vuelta conservando los
    adjuntos intactos.

    Se puede usar como gestor de contexto para limpiar automáticamente
    el directorio de trabajo cuando este se ha creado internamente:

        with Reqifz("documento.reqifz") as pack:
            pack.documents[0].update("_abc123", status="akzeptiert")
            pack.save("documento_editado.reqifz")
    """

    def __init__(self, source, work_dir: str | Path | None = None):
        self._owns_work_dir = work_dir is None
        self.work_dir = Path(work_dir) if work_dir else Path(tempfile.mkdtemp(prefix='pyreqif_'))
        self.work_dir.mkdir(parents=True, exist_ok=True)

        with zipfile.ZipFile(source) as zf:
            zf.extractall(self.work_dir)

        self._document_paths = sorted(
            p.relative_to(self.work_dir) for p in self.work_dir.rglob('*.reqif')
        )
        self.documents: list[Reqif] = [Reqif(self.work_dir / p) for p in self._document_paths]

    # -- acceso a los documentos ---------------------------------------

    def get(self, name_or_index) -> Reqif:
        """Busca un documento por índice, o por nombre (con o sin ruta
        relativa dentro del zip)."""
        if isinstance(name_or_index, int):
            return self.documents[name_or_index]
        for rel_path, doc in zip(self._document_paths, self.documents):
            if rel_path.name == name_or_index or str(rel_path) == name_or_index:
                return doc
        raise KeyError(name_or_index)

    def __iter__(self):
        return iter(self.documents)

    def __len__(self):
        return len(self.documents)

    def document_names(self) -> list[str]:
        return [p.name for p in self._document_paths]

    # -- adjuntos ------------------------------------------------------

    def image_path(self, image_ref: str) -> Path:
        """Resuelve una ruta de imagen (tal como aparece en
        Requirement.images) a una ruta absoluta dentro del directorio
        de trabajo."""
        return self.work_dir / image_ref

    def read_image(self, image_ref: str) -> bytes:
        return self.image_path(image_ref).read_bytes()

    # -- persistencia ----------------------------------------------------

    def save(self, destination):
        """Vuelca los cambios de cada documento y reempaqueta todo el
        directorio de trabajo (documentos + adjuntos) en `destination`
        (ruta u objeto tipo fichero), como .reqifz."""
        for rel_path, doc in zip(self._document_paths, self.documents):
            doc.save(self.work_dir / rel_path)

        with zipfile.ZipFile(destination, 'w', zipfile.ZIP_DEFLATED) as zf:
            for file_path in sorted(self.work_dir.rglob('*')):
                if file_path.is_file():
                    zf.write(file_path, file_path.relative_to(self.work_dir))

    def close(self):
        """Borra el directorio de trabajo, solo si lo creó esta
        instancia (no si se pasó `work_dir` explícitamente)."""
        if self._owns_work_dir and self.work_dir.exists():
            shutil.rmtree(self.work_dir)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        self.close()
