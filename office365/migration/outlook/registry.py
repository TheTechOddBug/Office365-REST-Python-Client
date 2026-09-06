"""Outlook scan registry — the product-scoped ScanDefinitions.

Outlook scans live in their own registry (the shared ``SCANS`` stays
SharePoint-scoped); :class:`~office365.migration.outlook.assessor.MailboxAssessor`
consumes it via :func:`outlook_scan_pairs`.
"""

from __future__ import annotations

from office365.migration.assessment.containers import ScanContainer
from office365.migration.assessment.registry import ScanDefinition
from office365.migration.outlook.scanner import MailboxFolderScan, OutlookOptions

OUTLOOK_SCANS: list[ScanDefinition] = [
    ScanDefinition(
        name="MailFolders",
        scanner=MailboxFolderScan,
        container=ScanContainer.MAIL_FOLDER,
        properties={"large_folder_items": 100_000},
    ),
]


def outlook_scan_pairs(options: OutlookOptions | None = None) -> list[tuple[ScanDefinition, MailboxFolderScan]]:
    """The enabled ``(definition, scanner)`` pairs from :data:`OUTLOOK_SCANS`."""
    options = options or OutlookOptions()
    pairs = []
    for definition in OUTLOOK_SCANS:
        if definition.enabled and definition.name not in options.disabled_scans:
            pairs.append((definition, definition.scanner(options)))
    return pairs
