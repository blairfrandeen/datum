from dataclasses import dataclass
from collections import namedtuple
import pytest
import sys

# Mock Missing NXOpen Module
NXOpen = type(sys)("NXOpen")
NXOpen.Session = type(sys)("Session")
NXOpen.Session.GetSession = type(sys)("GetSession")
NXOpen.Session.SetUndoMark = type(sys)("SetUndoMark")
NXOpen.Session.MarkVisibility = type(sys)("MarkVisibility")
NXOpen.Session.MarkVisibility.Visible = type(sys)("Visible")
NXOpen.Session.UpdateManager = type(sys)("UpdateManager")
NXOpen.Session.UpdateManager.DoInterpartUpdate = type(sys)("DoInterpartUpdate")
sys.modules["NXOpen"] = NXOpen

import nx_journals.nx_get_measurements as nxgm


Point = namedtuple("Point", "X Y Z")


@dataclass
class MockFeature:
    Name: str = "Mock Feature Name"
    Suppressed: bool = True


@dataclass
class MockWCS:
    Origin = Point(1, 2, 3)


@dataclass
class MockWorkPart:
    Features = [MockFeature(Name="Mock Feature Name")]
    WCS = MockWCS()
    Name: str = "Mock Work Part"
    FullPath: str = "Mock Work Part/A1"
    PartUnits: int = 1  # millimeters


@dataclass
class MockParts:
    Work = MockWorkPart


@dataclass
class MockSession:
    Parts = MockParts
    ReleaseNumber = 1969  # in the sunshine ;)
    SetUndoMark = lambda *_: None  # type(sys)("SetUndoMark")
    UpdateManager = type(sys)("UpdateManager")
    UpdateManager.DoInterpartUpdate = lambda *_: None


@pytest.fixture
def nxSession():
    yield MockSession()


@pytest.mark.parametrize(
    "feature_name, retval",
    [("Mock Feature Name", MockFeature()), ("Not a feature", None)],
)
def test_find_feature_by_name(feature_name, retval, monkeypatch):
    monkeypatch.setattr("NXOpen.Session.GetSession", lambda: MockSession)
    assert nxgm.find_feature_by_name(feature_name) == retval


def test_get_WCS(nxSession):
    assert nxgm.get_WCS(nxSession)["name"] == "World Coordinate System"
    assert nxgm.get_WCS(nxSession)["expressions"][0]["value"] == {
        "x": 1,
        "y": 2,
        "z": 3,
    }


def test_get_metadata(nxSession, monkeypatch):
    monkeypatch.setattr(nxgm, "nxprint", lambda arg: print(arg))
    md = nxgm.get_metadata(nxSession)["METADATA"]
    assert md["source_version"] == "1969"
    assert md["source_type"] == "NX"
    assert md["part_rev"] == "A1"
    mock_hd_part_name = "Something on your hard drive"
    nxSession.Parts.Work.FullPath = mock_hd_part_name
    nxSession.Parts.Work.Name = mock_hd_part_name
    md = nxgm.get_metadata(nxSession)["METADATA"]
    assert md["part_rev"] is None
    assert md["part_name"] == mock_hd_part_name


# TODO: Full coverage. Need several mock features.
@pytest.mark.xfail
def test_export(nxSession, monkeypatch):
    monkeypatch.setattr(nxgm, "nxprint", lambda arg: print(arg))
    monkeypatch.setattr(nxgm, "write_metadata_db", lambda _: None)
    # TODO: Pass valid JSON file?
    num_feats = nxgm.export_measurements("test str", nxSession)
    assert num_feats == 0
