import json
import os
import shutil
import sqlite3
import tempfile
import zipfile
from pathlib import Path


_BACKUP_FILES = {"bookmarks.db", "settings.json"}


def create_data_backup(
    database_connection: sqlite3.Connection,
    settings_path: Path,
    archive_path: Path,
) -> None:
    settings = json.loads(settings_path.read_text(encoding="utf-8"))
    if not isinstance(settings, dict):
        raise ValueError("Ayarlar dosyasının kök değeri bir JSON nesnesi olmalı.")

    archive_path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(
        prefix=".doc-network-backup-", dir=archive_path.parent
    ) as temporary_dir:
        temporary_path = Path(temporary_dir)
        database_snapshot = temporary_path / "bookmarks.db"
        temporary_archive = temporary_path / "backup.zip"

        snapshot_connection = sqlite3.connect(database_snapshot)
        try:
            database_connection.backup(snapshot_connection)
        finally:
            snapshot_connection.close()

        with zipfile.ZipFile(
            temporary_archive, "w", compression=zipfile.ZIP_DEFLATED
        ) as archive:
            archive.write(database_snapshot, "bookmarks.db")
            archive.write(settings_path, "settings.json")

        os.replace(temporary_archive, archive_path)


def _validate_database(database_path: Path) -> None:
    connection = sqlite3.connect(
        f"{database_path.resolve().as_uri()}?mode=ro", uri=True
    )
    try:
        integrity = connection.execute("PRAGMA integrity_check").fetchone()
        if integrity is None or integrity[0] != "ok":
            raise ValueError("Yedekteki veritabanı bütünlük kontrolünden geçemedi.")

        tables = {
            row[0]
            for row in connection.execute(
                "SELECT name FROM sqlite_master WHERE type='table'"
            )
        }
        if not {"groups", "bookmarks"}.issubset(tables):
            raise ValueError("Yedekte gerekli veritabanı tabloları bulunamadı.")
    finally:
        connection.close()


def restore_data_backup(
    archive_path: Path,
    destination_dir: Path,
    active_connection: sqlite3.Connection,
) -> None:
    connection_closed = False
    try:
        destination_dir.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(
            prefix=".doc-network-restore-", dir=destination_dir
        ) as temporary_dir:
            temporary_path = Path(temporary_dir)
            staged_database = temporary_path / "bookmarks.db"
            staged_settings = temporary_path / "settings.json"

            with zipfile.ZipFile(archive_path) as archive:
                names = archive.namelist()
                if len(names) != 2 or set(names) != _BACKUP_FILES:
                    raise ValueError(
                        "Yedek, yalnızca bookmarks.db ve settings.json içermeli."
                    )

                for name in _BACKUP_FILES:
                    with archive.open(name) as source, (temporary_path / name).open(
                        "wb"
                    ) as destination:
                        shutil.copyfileobj(source, destination)

            settings = json.loads(staged_settings.read_text(encoding="utf-8"))
            if not isinstance(settings, dict):
                raise ValueError("Yedekteki ayarların kök değeri JSON nesnesi değil.")
            _validate_database(staged_database)

            rollback_database = temporary_path / "previous-bookmarks.db"
            rollback_settings = temporary_path / "previous-settings.json"
            rollback_connection = sqlite3.connect(rollback_database)
            try:
                active_connection.backup(rollback_connection)
            finally:
                rollback_connection.close()
            shutil.copy2(destination_dir / "settings.json", rollback_settings)

            active_connection.close()
            connection_closed = True

            database_path = destination_dir / "bookmarks.db"
            settings_path = destination_dir / "settings.json"
            sidecars = (
                Path(f"{database_path}-wal"),
                Path(f"{database_path}-shm"),
            )

            def clear_sidecars() -> None:
                for sidecar in sidecars:
                    sidecar.unlink(missing_ok=True)

            try:
                clear_sidecars()
                os.replace(staged_database, database_path)
                os.replace(staged_settings, settings_path)
            except Exception as replace_error:
                try:
                    clear_sidecars()
                    os.replace(rollback_database, database_path)
                    os.replace(rollback_settings, settings_path)
                except Exception as rollback_error:
                    raise RuntimeError(
                        f"Yedek geri yüklenemedi ({replace_error}); "
                        f"önceki veriler de geri alınamadı ({rollback_error})."
                    ) from rollback_error
                raise
    finally:
        if not connection_closed:
            active_connection.close()
