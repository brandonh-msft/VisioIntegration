"""Diagram save/export tools."""

from __future__ import annotations

import json
import logging
import os
from pathlib import Path
from typing import Any

from visio_mcp._state import mcp, _diagram, _layout, _waf, _caf
from visio_mcp.drawio_engine import DrawioEngine
from visio_mcp.visio_engine import VisioEngine

logger = logging.getLogger(__name__)


def _drawio_only_runtime_enabled() -> bool:
    """Return True when the runtime should only emit draw.io output."""
    value = os.getenv("VISIO_MCP_FORCE_DRAWIO", "")
    return value.strip().lower() in {"1", "true", "yes", "on"}

# ═══════════════════════════════════════════════════════════════════
# TOOL: save_diagram
# ═══════════════════════════════════════════════════════════════════

@mcp.tool()
def save_diagram(
    output_path: str,
    format: str = "vsdx",
    stencil_dir: str | None = None,
    icons_root: str | None = None,
    auto_layout_before_save: bool = True,
    layout_strategy: str = "tiered",
) -> dict[str, Any]:
    """Save the current diagram as a Visio .vsdx or draw.io (.drawio) file.

    If Microsoft Visio is installed, uses COM automation for full-fidelity output
    with official Azure SVG icons imported directly. Otherwise, uses python-vsdx
    for basic .vsdx creation. Alternatively, choose 'drawio' format for a
    portable XML file that opens in draw.io desktop, VS Code, or diagrams.net.

    Args:
        output_path: File path for the output file
                     (e.g., 'C:/diagrams/my-architecture.vsdx').
        format: Output format — 'vsdx' (default) or 'drawio'.
        stencil_dir: Optional directory containing Azure Visio stencil files (.vssx).
                     These take priority over SVG icons if both are available.
                     Only used for 'vsdx' format.
        icons_root: Optional directory containing the Azure Public Service Icons
                    (the 'Icons' folder with category subfolders like compute/, networking/).
                    Defaults to the bundled stencils directory. Only used for 'vsdx' format.
        auto_layout_before_save: Whether to auto-layout before saving (default: True).
        layout_strategy: Layout strategy if auto-layout is enabled
                         ('tiered', 'grid', 'grouped').

    Returns:
        Save status, output path, and rendering method used.
    """
    requested_format = format.lower().strip()
    fmt = requested_format
    if fmt not in ("vsdx", "drawio"):
        return {"status": "error", "message": f"Unsupported format '{format}'. Use 'vsdx' or 'drawio'."}

    drawio_only_runtime = _drawio_only_runtime_enabled()
    if drawio_only_runtime and fmt == "vsdx":
        logger.info("Forcing draw.io output because VISIO_MCP_FORCE_DRAWIO is enabled")
        fmt = "drawio"

    # Resolve relative paths to a well-known output directory
    _output_dir = Path(__file__).resolve().parent.parent.parent.parent / "output"
    out = Path(output_path)
    if not out.is_absolute():
        out = _output_dir / out.name  # always land in output/
    if not out.suffix:
        out = out.with_suffix(f".{fmt}")
    # Force correct extension for the chosen format
    expected_ext = f".{fmt}"
    if out.suffix.lower() != expected_ext:
        out = out.with_suffix(expected_ext)
    out.parent.mkdir(parents=True, exist_ok=True)
    output_path = str(out)

    if auto_layout_before_save and not _diagram.state.properties.get("preserve_original_style"):
        # Use stored layout hints if available (from reference architectures)
        layout_hints = getattr(_diagram.state, '_layout_hints', None) or None
        boundary_hints = getattr(_diagram.state, '_boundary_hints', None) or None
        if layout_hints or boundary_hints:
            _layout.auto_layout(
                _diagram.state,
                strategy=layout_strategy,
                layout_hints=layout_hints,
                boundary_hints=boundary_hints,
            )
        else:
            # Auto-layout if resources are unpositioned or clustered at same spot
            positions = [
                (r.position.x, r.position.y)
                for r in _diagram.state.resources.values()
            ]
            unique_positions = set(positions)
            needs_layout = (
                len(positions) == 0
                or all(x <= 0.1 and y <= 0.1 for x, y in positions)
                or (len(positions) > 1 and len(unique_positions) == 1)
            )
            if needs_layout:
                _layout.auto_layout(_diagram.state, strategy=layout_strategy)
            else:
                # Even if positioned, still fit boundaries to enclose their resources
                _layout._fit_boundaries(_diagram.state)

    if fmt == "drawio":
        engine = DrawioEngine()
        rendering_method = "draw.io (mxGraph XML)"
    else:
        engine = VisioEngine(stencil_dir=stencil_dir, icons_root=icons_root)
        from visio_mcp.visio_engine import VISIO_AVAILABLE
        rendering_method = "Visio COM automation (SVG icons)" if VISIO_AVAILABLE else "python-vsdx (basic)"

    try:
        saved_path = engine.render(_diagram.state, output_path)
    except Exception as e:
        if fmt == "vsdx":
            # Auto-fallback to .drawio when .vsdx fails
            logger.warning("Visio render failed (%s), falling back to .drawio", e)
            fallback_path = Path(output_path).with_suffix(".drawio")
            try:
                fallback_engine = DrawioEngine()
                saved_path = fallback_engine.render(_diagram.state, str(fallback_path))
                return {
                    "status": "saved",
                    "output_path": saved_path,
                    "format": "drawio",
                    "rendering_method": "draw.io (mxGraph XML) — fallback from .vsdx failure",
                    "vsdx_error": str(e),
                    "resource_count": len(_diagram.state.resources),
                    "connection_count": len(_diagram.state.connections),
                    "boundary_count": len(_diagram.state.boundaries),
                }
            except Exception as e2:
                return {
                    "status": "error",
                    "message": f"Failed to save diagram: vsdx={e}, drawio={e2}",
                }
        return {
            "status": "error",
            "message": f"Failed to save diagram: {e}",
        }

    return {
        "status": "saved",
        "output_path": saved_path,
        "format": fmt,
        "rendering_method": rendering_method,
        "resource_count": len(_diagram.state.resources),
        "connection_count": len(_diagram.state.connections),
        "boundary_count": len(_diagram.state.boundaries),
        **(
            {
                "requested_format": requested_format,
                "message": "This runtime only supports draw.io output; saved as .drawio.",
            }
            if drawio_only_runtime and requested_format != fmt
            else {}
        ),
    }

