import os


def get_required_env(name: str) -> str:
    value = os.environ.get(name)
    if not value:
        raise RuntimeError(f"Environment variable {name} is not set")
    return value


def get_telegram_token() -> str:
    return get_required_env("TELEGRAM_BOT_TOKEN")


def get_max_token() -> str:
    return get_required_env("MAX_BOT_TOKEN")


def get_group_number_regex() -> str:
    return os.environ.get("GROUP_NUMBER_REGEX", r"\d{7}-\d{5}")


def get_webdav_options() -> dict:
    return {
        "webdav_hostname": "https://webdav.yandex.ru",
        "webdav_login": get_required_env("YANDEX_WEBDAV_LOGIN"),
        "webdav_password": get_required_env("YANDEX_WEBDAV_TOKEN"),
        "timeout": 300,
    }


def get_webdav_upload_dir() -> str:
    return os.environ.get("YANDEX_WEBDAV_UPLOAD_DIR", "/Практики/")


