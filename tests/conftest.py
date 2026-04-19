from __future__ import annotations

import importlib.util
from pathlib import Path

import pytest

_BUILDER_PATH = Path(__file__).resolve().parent / "fixtures" / "sample_docx_builder.py"
_BUILDER_SPEC = importlib.util.spec_from_file_location(
    "sample_docx_builder",
    _BUILDER_PATH,
)
if _BUILDER_SPEC is None or _BUILDER_SPEC.loader is None:
    raise RuntimeError(f"Unable to load fixture builder from {_BUILDER_PATH}")
_BUILDER_MODULE = importlib.util.module_from_spec(_BUILDER_SPEC)
_BUILDER_SPEC.loader.exec_module(_BUILDER_MODULE)
build_sample_docx = _BUILDER_MODULE.build_sample_docx


@pytest.fixture
def sample_docx_path(tmp_path_factory: pytest.TempPathFactory) -> Path:
    path = tmp_path_factory.mktemp("sample-docx") / "sample.docx"
    return build_sample_docx(path)
