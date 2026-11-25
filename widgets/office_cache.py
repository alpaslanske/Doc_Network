import hashlib
from functools import lru_cache
from pathlib import Path


@lru_cache(maxsize=256)
def _content_digest(
    source_path: str,
    size: int,
    modified_ns: int,
    changed_ns: int,
    inode: int,
) -> str:
    digest = hashlib.sha256(source_path.encode("utf-8", "surrogatepass"))
    digest.update(f"\0{size}\0{modified_ns}\0{changed_ns}\0{inode}".encode())
    with Path(source_path).open("rb") as source_file:
        for chunk in iter(lambda: source_file.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def get_office_cache_path(input_path: Path, cache_dir: Path) -> Path:
    source_path = Path(input_path).resolve()
    stat = source_path.stat()
    digest = _content_digest(
        str(source_path),
        stat.st_size,
        stat.st_mtime_ns,
        stat.st_ctime_ns,
        getattr(stat, "st_ino", 0),
    )
    return Path(cache_dir) / f"{digest}.pdf"
