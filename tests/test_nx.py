from dataclasses import dataclass
import pytest
import sys

# Mock Missing NXOpen Module
NXOpen = type(sys)("NXOpen")
NXOpen.Session = type(sys)("Session")
NXOpen.Session.GetSession = type(sys)("GetSession")
sys.modules["NXOpen"] = NXOpen

import nx_journals.nx_get_measurements as nxgm


@pytest.mark.xfail
def test_export():
    assert 0


@dataclass
class MockFeature:
    Name: str = "MockName"


@dataclass
class MockWorkPart:
    Features = [MockFeature(Name="MockName")]


@dataclass
class MockParts:
    Work = MockWorkPart


@dataclass
class MockSession:
    Parts = MockParts


@pytest.mark.parametrize(
    "feature_name, retval", [("MockName", MockFeature()), ("Not a feature", None)]
)
def test_find_feature_by_name(feature_name, retval, monkeypatch):
    monkeypatch.setattr("NXOpen.Session.GetSession", lambda: MockSession)
    assert nxgm.find_feature_by_name(feature_name) == retval
