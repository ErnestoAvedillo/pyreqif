"""Reqifz: opening, editing, and repackaging a complete .reqifz."""
from __future__ import annotations

import shutil
import tempfile
import zipfile
from pathlib import Path

from .reqif import Reqif


class Reqifz:
    """A .reqifz is a zip with one or more .reqif files plus the
    attachment files they reference (images, documents...). This class
    extracts it to a working directory, exposes each .reqif as a
    `Reqif`, and allows repackaging everything back while keeping the
    attachments intact.

    Can be used as a context manager to automatically clean up the
    working directory when it was created internally:

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

    # -- document access ------------------------------------------------

    def get(self, name_or_index) -> Reqif:
        """Looks up a document by index, or by name (with or without
        its relative path within the zip)."""
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

    # -- attachments ------------------------------------------------------

    def image_path(self, image_ref: str) -> Path:
        """Resolves an image path (as it appears in
        Requirement.images) to an absolute path within the working
        directory."""
        return self.work_dir / image_ref

    def read_image(self, image_ref: str) -> bytes:
        return self.image_path(image_ref).read_bytes()

    # -- persistence ----------------------------------------------------

    def save(self, destination):
        """Flushes the changes of each document and repackages the
        whole working directory (documents + attachments) into
        `destination` (path or file-like object), as a .reqifz."""
        for rel_path, doc in zip(self._document_paths, self.documents):
            doc.save(self.work_dir / rel_path)

        with zipfile.ZipFile(destination, 'w', zipfile.ZIP_DEFLATED) as zf:
            for file_path in sorted(self.work_dir.rglob('*')):
                if file_path.is_file():
                    zf.write(file_path, file_path.relative_to(self.work_dir))

    def close(self):
        """Deletes the working directory, only if this instance
        created it (not if `work_dir` was passed explicitly)."""
        if self._owns_work_dir and self.work_dir.exists():
            shutil.rmtree(self.work_dir)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        self.close()
