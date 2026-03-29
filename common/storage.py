import io
import zipfile

from webdav3.client import Client

from common.config import get_webdav_options, get_webdav_upload_dir


def _get_client() -> Client:
    return Client(get_webdav_options())


def ensure_remote_dir(path: str) -> None:
    client = _get_client()
    if not client.check(path):
        client.mkdir(path)


def upload_bytes(buf: io.BytesIO, remote_path: str) -> None:
    client = _get_client()
    buf.seek(0)
    client.upload_to(buf, remote_path)


def upload_zip(buf: io.BytesIO, filename: str, is_wrapped: bool = False) -> list[str]:
    """Upload .zip and extract contents into a group folder.

    Zip name is {GROUP}.zip. Contents are extracted into
    {UPLOAD_DIR}/{GROUP}/ — the group folder is created if it doesn't exist,
    or accumulated into if it does.

    If is_wrapped is True, the archive has a top-level {GROUP}/ folder that
    should be stripped (its contents go directly into the group dir).

    Returns list of uploaded remote paths.
    """
    upload_dir = get_webdav_upload_dir()
    ensure_remote_dir(upload_dir)

    uploaded = []

    # Upload the .zip itself
    zip_remote = f"{upload_dir}{filename}"
    upload_bytes(buf, zip_remote)
    uploaded.append(zip_remote)

    # Group folder: {UPLOAD_DIR}/{GROUP}/
    group_name = filename[:-4]  # strip .zip
    group_dir = f"{upload_dir}{group_name}/"
    ensure_remote_dir(group_dir)

    # Prefix to strip from entry paths when wrapped
    strip_prefix = f"{group_name}/" if is_wrapped else ""

    buf.seek(0)
    with zipfile.ZipFile(buf) as zf:
        for entry in zf.infolist():
            path = entry.filename
            if strip_prefix and not path.startswith(strip_prefix):
                continue
            rel_path = path[len(strip_prefix):]
            if not rel_path:
                continue

            if entry.is_dir():
                ensure_remote_dir(f"{group_dir}{rel_path}")
            else:
                if "/" in rel_path:
                    parent = rel_path.rsplit("/", 1)[0] + "/"
                    ensure_remote_dir(f"{group_dir}{parent}")
                data = io.BytesIO(zf.read(entry.filename))
                entry_remote = f"{group_dir}{rel_path}"
                upload_bytes(data, entry_remote)
                uploaded.append(entry_remote)

    return uploaded


def upload_single_member_zip(buf: io.BytesIO, group_number: str, folder_name: str) -> list[str]:
    """Upload a single student zip into {UPLOAD_DIR}/{GROUP}/{folder_name}/.

    The zip contents are placed directly into the member folder
    (top-level entries from the archive, no nesting by archive structure).

    Returns list of uploaded remote paths.
    """
    upload_dir = get_webdav_upload_dir()
    ensure_remote_dir(upload_dir)

    group_dir = f"{upload_dir}{group_number}/"
    ensure_remote_dir(group_dir)

    member_dir = f"{group_dir}{folder_name}/"
    ensure_remote_dir(member_dir)

    uploaded = []

    buf.seek(0)
    with zipfile.ZipFile(buf) as zf:
        for entry in zf.infolist():
            if entry.is_dir():
                entry_dir = f"{member_dir}{entry.filename}"
                ensure_remote_dir(entry_dir)
            else:
                if "/" in entry.filename:
                    parent = entry.filename.rsplit("/", 1)[0] + "/"
                    ensure_remote_dir(f"{member_dir}{parent}")
                data = io.BytesIO(zf.read(entry.filename))
                entry_remote = f"{member_dir}{entry.filename}"
                upload_bytes(data, entry_remote)
                uploaded.append(entry_remote)

    return uploaded
