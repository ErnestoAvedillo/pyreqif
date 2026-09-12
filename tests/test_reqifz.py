import zipfile

from pyreqif import Reqifz


def test_extracts_and_parses_documents(sample_reqifz_path):
    with Reqifz(sample_reqifz_path) as pack:
        assert pack.document_names() == ['sample.reqif']
        assert len(pack) == 1

        doc = pack.get(0)
        assert len(doc) == 3
        assert doc.get('_req-2').status == 'akzeptiert'

        # también se puede pedir por nombre de fichero
        assert pack.get('sample.reqif') is doc


def test_image_path_resolves_inside_work_dir(sample_reqifz_path):
    with Reqifz(sample_reqifz_path) as pack:
        doc = pack.get(0)
        req1 = doc.get('_req-1')
        assert req1.images == ['attachments/img1.png']

        image_path = pack.image_path(req1.images[0])
        assert image_path.exists()
        assert image_path.read_bytes().startswith(b'\x89PNG')
        assert pack.read_image(req1.images[0]).startswith(b'\x89PNG')


def test_save_repacks_with_edits_and_keeps_attachments(sample_reqifz_path, tmp_path):
    with Reqifz(sample_reqifz_path) as pack:
        doc = pack.get(0)
        doc.update('_req-1', comment='Revisado por el proveedor', status='akzeptiert')

        out_path = tmp_path / 'edited.reqifz'
        pack.save(out_path)

    with zipfile.ZipFile(sample_reqifz_path) as zf:
        original_names = set(zf.namelist())
    with zipfile.ZipFile(out_path) as zf:
        edited_names = set(zf.namelist())

    assert original_names == edited_names  # mismos ficheros, nada perdido

    with Reqifz(out_path) as reopened:
        req1 = reopened.get(0).get('_req-1')
        assert req1.comment == 'Revisado por el proveedor'
        assert req1.status == 'akzeptiert'


def test_work_dir_is_cleaned_up_when_owned_by_reqifz(sample_reqifz_path):
    with Reqifz(sample_reqifz_path) as pack:
        work_dir = pack.work_dir
        assert work_dir.exists()
    assert not work_dir.exists()


def test_explicit_work_dir_is_not_deleted_on_close(sample_reqifz_path, tmp_path):
    work_dir = tmp_path / 'kept'
    with Reqifz(sample_reqifz_path, work_dir=work_dir) as pack:
        assert pack.work_dir == work_dir
    assert work_dir.exists()
