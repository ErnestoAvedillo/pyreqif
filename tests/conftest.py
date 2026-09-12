import zipfile
from pathlib import Path

import pytest

FIXTURES = Path(__file__).parent / 'fixtures'


@pytest.fixture
def sample_reqif_path():
    return FIXTURES / 'sample.reqif'


@pytest.fixture
def sample_reqifz_path(tmp_path):
    """Construye un .reqifz mínimo: el sample.reqif + una imagen falsa
    en attachments/img1.png, sin depender de ningún documento real."""
    zip_path = tmp_path / 'sample.reqifz'
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.write(FIXTURES / 'sample.reqif', 'sample.reqif')
        zf.writestr('attachments/img1.png', b'\x89PNG\r\n\x1a\nfake-image-bytes')
    return zip_path
