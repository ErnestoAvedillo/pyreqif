import io
import zipfile
from pathlib import Path

import pytest
from PIL import Image as PILImage

FIXTURES = Path(__file__).parent / 'fixtures'


@pytest.fixture
def sample_reqif_path():
    return FIXTURES / 'sample.reqif'


def _tiny_png_bytes():
    buffer = io.BytesIO()
    PILImage.new('RGB', (40, 20), color=(200, 30, 30)).save(buffer, format='PNG')
    return buffer.getvalue()


@pytest.fixture
def sample_reqifz_path(tmp_path):
    """Construye un .reqifz mínimo: el sample.reqif + una imagen real
    (generada con Pillow) en attachments/img1.png, sin depender de
    ningún documento real."""
    zip_path = tmp_path / 'sample.reqifz'
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.write(FIXTURES / 'sample.reqif', 'sample.reqif')
        zf.writestr('attachments/img1.png', _tiny_png_bytes())
    return zip_path
