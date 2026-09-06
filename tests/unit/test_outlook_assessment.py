"""Offline tests for the Outlook mailbox assessor and MailFolders scan."""

from __future__ import annotations

from office365.migration import MailboxAssessor, OutlookOptions


class _FoldersEndpoint:
    """Duck-type of a Graph folder collection (select -> get -> execute)."""

    def __init__(self, items):
        self._items = items
        self._select = None

    def select(self, columns):
        self._select = columns
        return self

    def get(self):
        return self

    def execute_query(self):
        return self._items


class _MailFolder:
    def __init__(self, name, item_count=0, unread_count=0, children=None, folder_id=None):
        self.display_name = name
        self.id = folder_id or name
        self.total_item_count = item_count
        self.unread_item_count = unread_count
        self.child_folder_count = len(children or [])
        self._children = _FoldersEndpoint(children or [])

    @property
    def child_folders(self):
        return self._children


def _root_mailbox(folders):
    class _User:
        mail_folders = _FoldersEndpoint(folders)

    return _User()


def _scan(folders, options=None):
    return MailboxAssessor(_root_mailbox(folders), options or OutlookOptions()).assess()


def test_walker_reports_nested_folders_with_counts():
    inbox = _MailFolder(
        "Inbox",
        item_count=12,
        unread_count=3,
        children=[_MailFolder("Projects", item_count=5, unread_count=1, folder_id="p1")],
    )
    sent = _MailFolder("Sent Items", item_count=4, folder_id="s1")
    report = _scan([inbox, sent])

    scan = report.scan_reports["MailFolders"]
    rows = {row["FolderPath"]: row for row in scan.to_records()}
    assert rows["Inbox"]["ItemCount"] == 12  # noqa: PLR2004
    assert rows["Inbox"]["UnreadItemCount"] == 3  # noqa: PLR2004
    assert rows["Inbox/Projects"]["ItemCount"] == 5  # noqa: PLR2004
    assert rows["Sent Items"]["ItemCount"] == 4  # noqa: PLR2004
    assert report.issues == []


def test_large_folder_is_flagged():
    report = _scan(
        [_MailFolder("Archive", item_count=150_000, folder_id="a1")],
        OutlookOptions(large_folder_items=100_000),
    )
    flags = [i for i in report.issues if i.category == "mail"]
    assert len(flags) == 1
    assert "Archive" in flags[0].location
    assert flags[0].suggestion


def test_disabled_scan_collects_nothing():
    report = _scan([_MailFolder("Inbox", item_count=5, folder_id="i1")], OutlookOptions(disabled_scans={"MailFolders"}))
    assert report.scan_reports == {}
    assert report.issues == []


def test_unreadable_subtree_is_warning_not_fatal():
    broken = _MailFolder("Broken", folder_id="b1")

    class _BrokenEndpoint:
        def select(self, _cols):
            return self

        def get(self):
            return self

        def execute_query(self):
            raise RuntimeError("denied")

    broken._children = _BrokenEndpoint()  # type: ignore[assignment]

    ok = _MailFolder("Inbox", item_count=2, folder_id="i1")
    report = _scan([ok, broken])

    assert report.scan_reports["MailFolders"].records  # healthy folders still reported
    access = [i for i in report.issues if i.category == "access"]
    assert any("Broken" in i.location for i in access)


def test_report_columns_match_record():
    report = _scan([_MailFolder("Inbox", item_count=1, folder_id="i1")])
    header = report.scan_reports["MailFolders"].to_csv().splitlines()[0]
    assert header == "FolderPath,ItemCount,UnreadItemCount,ChildFolderCount"
