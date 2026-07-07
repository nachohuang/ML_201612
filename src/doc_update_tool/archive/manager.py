"""Manage the incoming/archive/output folder trio for one client.

Each archived version keeps three fixed-name files (bilingual.<ext>, diff_report.xlsx,
email.msg) so that "find the document" never has to guess among files that happen to
share an extension.
"""

from __future__ import annotations

import re
import shutil
from datetime import date
from pathlib import Path

from doc_update_tool.models import IncomingBundle

_DOC_SUFFIXES = {".xlsx", ".docx"}
_VERSION_DIR_RE = re.compile(r"^v(\d+)_")


class ArchiveError(RuntimeError):
    pass


class ArchiveManager:
    def __init__(self, incoming_dir: Path, archive_dir: Path, output_dir: Path):
        self.incoming_dir = Path(incoming_dir)
        self.archive_dir = Path(archive_dir)
        self.output_dir = Path(output_dir)

    def discover_incoming(self) -> IncomingBundle:
        if not self.incoming_dir.exists():
            raise ArchiveError(f"incoming 資料夾不存在: {self.incoming_dir}")

        entries = list(self.incoming_dir.iterdir())
        docs = [p for p in entries if p.suffix.lower() in _DOC_SUFFIXES]
        msgs = [p for p in entries if p.suffix.lower() == ".msg"]

        if len(docs) == 0:
            raise ArchiveError(f"incoming 資料夾內找不到 Word/Excel 附件: {self.incoming_dir}")
        if len(docs) > 1:
            raise ArchiveError(
                f"incoming 資料夾內有多份文件，一次只能處理一份: {[p.name for p in docs]}"
            )
        if len(msgs) > 1:
            raise ArchiveError(
                f"incoming 資料夾內有多份 .msg 信件，一次只能處理一份: {[p.name for p in msgs]}"
            )

        return IncomingBundle(document_path=docs[0], email_path=msgs[0] if msgs else None)

    def existing_versions(self) -> list[int]:
        if not self.archive_dir.exists():
            return []
        versions = []
        for p in self.archive_dir.iterdir():
            if p.is_dir():
                m = _VERSION_DIR_RE.match(p.name)
                if m:
                    versions.append(int(m.group(1)))
        return sorted(versions)

    def is_first_version(self) -> bool:
        return len(self.existing_versions()) == 0

    def next_version_no(self) -> int:
        versions = self.existing_versions()
        return (versions[-1] + 1) if versions else 1

    def latest_version_dir(self) -> Path | None:
        versions = self.existing_versions()
        if not versions:
            return None
        latest = versions[-1]
        for p in self.archive_dir.iterdir():
            if p.is_dir() and p.name.startswith(f"v{latest:02d}_"):
                return p
        return None

    def latest_bilingual_document(self) -> Path | None:
        version_dir = self.latest_version_dir()
        if version_dir is None:
            return None
        for suffix in _DOC_SUFFIXES:
            candidate = version_dir / f"bilingual{suffix}"
            if candidate.exists():
                return candidate
        return None

    def commit_version(
        self,
        version_no: int,
        document_path: Path,
        diff_report_path: Path | None,
        email_path: Path | None,
    ) -> Path:
        version_dir = self.archive_dir / f"v{version_no:02d}_{date.today():%Y%m%d}"
        version_dir.mkdir(parents=True, exist_ok=True)
        self.output_dir.mkdir(parents=True, exist_ok=True)

        bilingual_dest = version_dir / f"bilingual{document_path.suffix.lower()}"
        shutil.copy2(document_path, bilingual_dest)
        shutil.copy2(document_path, self.output_dir / bilingual_dest.name)

        if diff_report_path is not None:
            diff_dest = version_dir / "diff_report.xlsx"
            shutil.copy2(diff_report_path, diff_dest)
            shutil.copy2(diff_report_path, self.output_dir / diff_dest.name)

        if email_path is not None:
            # Kept in archive/ as evidence only — output/ is what gets uploaded to
            # Teams/SharePoint, and the raw customer email doesn't belong there.
            shutil.copy2(email_path, version_dir / "email.msg")

        return version_dir
