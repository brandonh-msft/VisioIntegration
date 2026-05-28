from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from visio_mcp.server import create_diagram
from visio_mcp.tools import save_tools


def test_save_diagram_forces_drawio_in_container_runtime(monkeypatch, tmp_path):
    create_diagram("Container Diagram")
    monkeypatch.setenv("VISIO_MCP_FORCE_DRAWIO", "1")

    def fail_visio_render(self, state, output_path):
        raise AssertionError("Visio rendering should not run in draw.io-only runtime")

    def render_drawio(self, state, output_path):
        Path(output_path).write_text("<mxfile />", encoding="utf-8")
        return output_path

    monkeypatch.setattr(save_tools.VisioEngine, "render", fail_visio_render)
    monkeypatch.setattr(save_tools.DrawioEngine, "render", render_drawio)

    result = save_tools.save_diagram(
        output_path=str(tmp_path / "container-diagram.vsdx"),
        format="vsdx",
        auto_layout_before_save=False,
    )

    assert result["status"] == "saved"
    assert result["format"] == "drawio"
    assert result["requested_format"] == "vsdx"
    assert result["output_path"].endswith(".drawio")
    assert "draw.io output" in result["message"]


def test_save_diagram_keeps_vsdx_when_not_forced(monkeypatch, tmp_path):
    create_diagram("Native Diagram")
    monkeypatch.delenv("VISIO_MCP_FORCE_DRAWIO", raising=False)

    def render_visio(self, state, output_path):
        Path(output_path).write_text("vsdx", encoding="utf-8")
        return output_path

    def fail_drawio_render(self, state, output_path):
        raise AssertionError("draw.io fallback should not run for a successful vsdx save")

    monkeypatch.setattr(save_tools.VisioEngine, "render", render_visio)
    monkeypatch.setattr(save_tools.DrawioEngine, "render", fail_drawio_render)

    result = save_tools.save_diagram(
        output_path=str(tmp_path / "native-diagram.vsdx"),
        format="vsdx",
        auto_layout_before_save=False,
    )

    assert result["status"] == "saved"
    assert result["format"] == "vsdx"
    assert "requested_format" not in result
    assert result["output_path"].endswith(".vsdx")
