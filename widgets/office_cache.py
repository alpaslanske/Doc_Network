import hashlib
from pathlib import Path


def get_office_cache_path(input_path: Path, cache_dir: Path) -> Path:
    source_path = Path(input_path).resolve()
    digest = hashlib.sha256(str(source_path).encode("utf-8", "surrogatepass"))

    with source_path.open("rb") as source_file:
        for chunk in iter(lambda: source_file.read(1024 * 1024), b""):
            digest.update(chunk)

    return Path(cache_dir) / f"{digest.hexdigest()}.pdf"
