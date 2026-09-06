"""Regression: ClientValueCollection-typed dataclass fields must keep their item
type when declared with a default value (server responses like GetSiteDesigns)."""

from __future__ import annotations

import uuid

from office365.runtime.client_value_collection import ClientValueCollection
from office365.sharepoint.sitedesigns.metadata import SiteDesignMetadata


def test_site_design_metadata_parses_non_empty_site_script_ids():
    design = SiteDesignMetadata()
    design.set_property(
        "SiteScriptIds",
        ["07702c07-0485-426f-b710-4704241caad9", "6250ceba-8724-4fb4-8c52-5a89183b9587"],
    )

    assert isinstance(design.SiteScriptIds, ClientValueCollection)
    assert len(design.SiteScriptIds) == 2  # noqa: PLR2004
    assert all(isinstance(item, uuid.UUID) for item in design.SiteScriptIds)


def test_site_design_metadata_parses_empty_collection():
    design = SiteDesignMetadata()
    design.set_property("SiteScriptIds", [])
    assert isinstance(design.SiteScriptIds, ClientValueCollection)
    assert len(design.SiteScriptIds) == 0
