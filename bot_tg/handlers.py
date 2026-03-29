import asyncio
import io
import re
import zipfile

from telegram import Update
from telegram.ext import ContextTypes

from common.config import get_group_number_regex
from common.logger import log
from common.storage import upload_zip, upload_single_member_zip

# Russian name: capital letter followed by lowercase Cyrillic
_RU_NAME = r"[А-ЯЁ][а-яё]+"
_FOLDER_RE = re.compile(rf"^({_RU_NAME})_({_RU_NAME})(_({_RU_NAME})?)?$")


def _build_group_zip_regex() -> re.Pattern:
    group = get_group_number_regex()
    return re.compile(rf"^({group})\.zip$")


def _build_member_zip_regex() -> re.Pattern:
    group = get_group_number_regex()
    return re.compile(rf"^({group})_({_RU_NAME})_({_RU_NAME})(_({_RU_NAME})?)?\.zip$")


def _get_top_level_folders(zf: zipfile.ZipFile) -> list[str]:
    folders = set()
    for entry in zf.infolist():
        parts = entry.filename.split("/")
        if len(parts) > 1 and parts[0]:
            folders.add(parts[0])
    return sorted(folders)


def _get_student_folders(zf: zipfile.ZipFile, group_number: str) -> tuple[list[str], bool]:
    top_folders = _get_top_level_folders(zf)

    if len(top_folders) == 1 and top_folders[0] == group_number:
        folders = set()
        for entry in zf.infolist():
            parts = entry.filename.split("/")
            if len(parts) > 2 and parts[0] == group_number and parts[1]:
                folders.add(parts[1])
        return sorted(folders), True

    return top_folders, False


def _validate_group_zip(filename: str, group_number: str, buf: io.BytesIO) -> tuple[bool, bool, list[str]]:
    errors = []
    buf.seek(0)
    try:
        with zipfile.ZipFile(buf) as zf:
            folders, is_wrapped = _get_student_folders(zf, group_number)
            if not folders:
                errors.append("The archive contains no student folders.")
                return False, False, errors

            invalid_folders = [f for f in folders if not _FOLDER_RE.match(f)]
            if invalid_folders:
                errors.append("The following folders have incorrect names:")
                for f in invalid_folders:
                    errors.append(f"  • {f}")
                errors.append("")
                errors.append(
                    "Folder names must match the format:\n"
                    "  {Фамилия}_{Имя}_{Отчество}\n"
                    "Example: Иванов_Иван_Иванович\n"
                    "Middle name (Отчество) can be omitted:\n"
                    "  Иванов_Иван_"
                )
                return False, is_wrapped, errors
    except zipfile.BadZipFile:
        errors.append("The file doesn't appear to be a valid ZIP archive.")
        return False, False, errors

    return True, is_wrapped, []


def _parse_filename(filename: str) -> tuple[str, str | None]:
    m = _build_group_zip_regex().match(filename)
    if m:
        return "group", m.group(1)

    m = _build_member_zip_regex().match(filename)
    if m:
        return "member", m.group(1)

    return "invalid", None


def _member_folder_name(filename: str) -> str:
    name = filename[:-4]
    _, rest = name.split("_", 1)
    return rest


def _user_info(update: Update) -> str:
    user = update.effective_user
    if user:
        return f"@{user.username}" if user.username else f"{user.full_name} (id={user.id})"
    return "unknown"


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    log.info("User %s sent /start", _user_info(update))
    await update.message.reply_text("Hello! I'm your Telegram bot.")


async def echo(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    log.info("User %s sent message: %s", _user_info(update), update.message.text)
    await update.message.reply_text(update.message.text)


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    doc = update.message.document
    if not doc:
        return

    user = _user_info(update)
    filename = doc.file_name or ""

    if not filename.lower().endswith(".zip"):
        log.warning("User %s sent non-zip file: %s", user, filename)
        await update.message.reply_text("Please send a .zip file.")
        return

    log.info("User %s uploaded file: %s (%d bytes)", user, filename, doc.file_size)

    # Download
    file = await doc.get_file()
    buf = io.BytesIO()
    await file.download_to_memory(buf)

    # Log zip contents
    buf.seek(0)
    try:
        with zipfile.ZipFile(buf) as zf:
            entries = zf.infolist()
            total_size = sum(e.file_size for e in entries)
            log.info("ZIP %s: %d items, %d bytes uncompressed", filename, len(entries), total_size)
            for e in entries:
                log.info("  %s (%d bytes)", e.filename, e.file_size)
    except zipfile.BadZipFile:
        log.error("Invalid ZIP file: %s", filename)

    # Determine type
    file_type, group_number = _parse_filename(filename)

    if file_type == "invalid":
        log.warning("Validation failed for %s from %s: filename doesn't match any pattern", filename, user)
        await update.message.reply_text(
            "Validation failed:\n\n"
            f"Archive name \"{filename}\" doesn't match any expected format:\n"
            "  • {Группа}.zip (e.g. 5308.zip) — for a whole group\n"
            "  • {Группа}_{Фамилия}_{Имя}_{Отчество}.zip "
            "(e.g. 5308_Иванов_Иван_Иванович.zip) — for a single student\n\n"
            "Please rename and upload again."
        )
        return

    if file_type == "group":
        valid, is_wrapped, errors = _validate_group_zip(filename, group_number, buf)
        if not valid:
            log.warning("Validation failed for %s from %s: %s", filename, user, "; ".join(errors))
            lines = ["Validation failed:", ""]
            lines.extend(errors)
            lines.append("")
            lines.append("Please fix the issues and upload the archive again.")
            await update.message.reply_text("\n".join(lines))
            return

        try:
            remote_paths = await asyncio.to_thread(upload_zip, buf, filename, is_wrapped)
            log.info("Group archive %s from %s uploaded (%d files)", filename, user, len(remote_paths))
            for rp in remote_paths:
                log.info("  -> %s", rp)
            lines = [
                f"Group archive {filename} accepted!",
                f"Uploaded {len(remote_paths)} file(s) to Yandex Disk.",
            ]
            await update.message.reply_text("\n".join(lines))
        except Exception as e:
            log.error("Upload failed for %s from %s: %s", filename, user, e)
            await update.message.reply_text(f"Upload to Yandex Disk failed: {e}")

    elif file_type == "member":
        folder_name = _member_folder_name(filename)
        try:
            remote_paths = await asyncio.to_thread(
                upload_single_member_zip, buf, group_number, folder_name
            )
            log.info("Student archive %s from %s uploaded to group %s as %s/ (%d files)",
                     filename, user, group_number, folder_name, len(remote_paths))
            for rp in remote_paths:
                log.info("  -> %s", rp)
            lines = [
                f"Student archive {filename} accepted!",
                f"Uploaded to group {group_number} as {folder_name}/",
                f"{len(remote_paths)} file(s) uploaded to Yandex Disk.",
            ]
            await update.message.reply_text("\n".join(lines))
        except Exception as e:
            log.error("Upload failed for %s from %s: %s", filename, user, e)
            await update.message.reply_text(f"Upload to Yandex Disk failed: {e}")
