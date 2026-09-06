"""Outlook (mail) assessor — walks a user's mailbox folders and runs the scans.

The first non-SharePoint product slice: instead of walking a web/list tree, the
walker enumerates the mailbox's mail-folder tree (``mailFolders`` ->
``childFolders``), builds ``MAIL_FOLDER`` payloads with the folder's counts and
dispatches to the registered scans. Reuses the shared :class:`AssessmentReport`
/ :class:`ScanReport` / issue model and exporters.
"""

from __future__ import annotations

import uuid
from dataclasses import dataclass

from office365.migration.assessment.containers import ScanContainer
from office365.migration.assessment.issue import AssessmentIssue
from office365.migration.assessment.report import AssessmentReport, ScanReport
from office365.migration.assessment.scanners.base import ScanTarget
from office365.migration.outlook.registry import outlook_scan_pairs
from office365.migration.outlook.scanner import OutlookOptions

# Graph mailFolder properties the walker selects
_FOLDER_COLUMNS = ["displayName", "totalItemCount", "unreadItemCount", "childFolderCount"]


@dataclass
class MailFolderData:
    """A MAIL_FOLDER payload handed to the folder scan (plain, testable data)."""

    path: str
    item_count: int | None = 0
    unread_count: int | None = 0
    child_count: int | None = 0


class MailboxAssessor:
    """Assesses a user's mailbox — folder inventory with large-folder flags.

    Example:
        >>> report = MailboxAssessor(client.me).assess()
        >>> print(report.scan_reports["MailFolders"].to_csv())
    """

    def __init__(self, user, options: OutlookOptions | None = None) -> None:
        self._user = user
        self._options = options or OutlookOptions()
        self.folder_count = 0
        self.message_count = 0

    def assess(self, progress=None) -> AssessmentReport:
        """Scan the user's mail-folder tree (eager, Graph reads only)."""
        report = AssessmentReport()
        report.scan_id = str(uuid.uuid4())
        active = outlook_scan_pairs(self._options)
        folder_scans = [scanner for definition, scanner in active if definition.container is ScanContainer.MAIL_FOLDER]

        def _walk(folders, parent_path: str) -> None:
            for folder in folders:
                name = folder.display_name or folder.id
                path = f"{parent_path}/{name}" if parent_path else name
                folder_data = MailFolderData(
                    path=path,
                    item_count=_to_int(folder.total_item_count),
                    unread_count=_to_int(folder.unread_item_count),
                    child_count=_to_int(folder.child_folder_count),
                )
                self.folder_count += 1
                self.message_count += folder_data.item_count or 0
                target = ScanTarget(ScanContainer.MAIL_FOLDER, folder_data, path)
                for scanner in folder_scans:
                    scanner.run(target, report)
                try:
                    children = folder.child_folders.select(_FOLDER_COLUMNS).get().execute_query()
                except Exception as e:  # noqa: BLE001 — unreadable subtree is a warning, not fatal
                    report.issues.append(AssessmentIssue("warning", "access", path, f"folder scan skipped — {e}"))
                    continue
                if children:
                    _walk(children, path)

        try:
            root_folders = self._user.mail_folders.select(_FOLDER_COLUMNS).get().execute_query()
        except Exception as e:  # noqa: BLE001
            report.issues.append(AssessmentIssue("warning", "access", "mailbox", f"mailbox scan skipped — {e}"))
            root_folders = []

        if root_folders:
            _walk(root_folders, "")

        for definition, scanner in active:
            if scanner.records:
                report._scan_reports[scanner.scan_name] = ScanReport(
                    name=scanner.scan_name,
                    container=definition.container,
                    columns=scanner.columns,
                    records=scanner.records,
                )
        return report


def _to_int(value) -> int | None:
    try:
        return int(value) if value is not None else None
    except (TypeError, ValueError):
        return None
