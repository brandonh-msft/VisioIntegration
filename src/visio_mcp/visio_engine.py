"""Visio COM automation engine for rendering Azure architecture diagrams.

Uses win32com.client to automate Microsoft Visio when available.
Imports official Azure SVG icons directly into the Visio document.
Falls back to python-vsdx for basic .vsdx file creation when Visio is not installed.

Visual standards follow Azure Architecture Center conventions:
  - Official SVG icons at 1:1 aspect ratio (no crop/flip/rotate)
  - Product name label below each icon
  - Standard boundary colors per Microsoft diagram palette
  - Numbered workflow step circles on data-flow arrows
  - Dashed lines for management/identity, solid for data flows
"""

from __future__ import annotations

import logging
import os
from pathlib import Path

from .azure_catalog import (
    AZURE_SHAPE_CATALOG,
    BOUNDARY_STYLES,
    CONNECTOR_STYLES,
    SVG_ICON_MAP,
    get_icons_root,
    resolve_svg_path,
)
from .models import DiagramState
from .reference_architectures import AZURE_DIAGRAM_COLORS, MICROSOFT_STANDARDS

logger = logging.getLogger(__name__)

# Visio constants (from Visio type library)
_VIS_STANDARD = 0
_VIS_LANDSCAPE = 2
_VIS_PAGE_WIDTH = 0
_VIS_PAGE_HEIGHT = 1
_VIS_SAVE_AS_VSDX = 60


def _try_import_win32():
    """Attempt to import win32com."""
    try:
        import win32com.client  # noqa: F401
        return True
    except ImportError:
        return False


VISIO_AVAILABLE = _try_import_win32()


class VisioEngine:
    """Renders a DiagramState into a Visio .vsdx file."""

    def __init__(self, stencil_dir: str | None = None, icons_root: str | None = None) -> None:
        self._stencil_dir = stencil_dir
        self._icons_root = Path(icons_root) if icons_root else get_icons_root()
        self._app = None
        self._doc = None
        self._page = None
        self._stencils: dict[str, object] = {}
        self._shapes: dict[str, object] = {}

    # ── Public API ────────────────────────────────────────────────

    def render(self, state: DiagramState, output_path: str) -> str:
        """Render the diagram state to a .vsdx file. Returns the output path."""
        output_path = os.path.abspath(output_path)
        self._source_image = state.properties.get("source_image_path")

        if VISIO_AVAILABLE:
            return self._render_com(state, output_path)
        else:
            return self._render_vsdx(state, output_path)

    def close(self) -> None:
        """Close any open Visio COM instance."""
        if self._app is not None:
            try:
                self._app.Quit()
            except Exception:
                pass
            self._app = None

    def _flip_y(self, y: float) -> float:
        """Convert top-down layout Y to Visio bottom-up Y coordinate."""
        return self._page_height - y

    # ── COM-based rendering (requires Visio) ──────────────────────

    def _render_com(self, state: DiagramState, output_path: str) -> str:
        """Render the diagram via Visio COM automation.

        Creates a new Visio document, sets up page dimensions based on content
        bounding box, draws boundaries (largest first for correct z-order),
        resources (with SVG icon import), and connections (with step circles).
        Kills any orphaned VISIO.EXE processes before starting.

        Args:
            state: The complete diagram state to render.
            output_path: Absolute path for the output .vsdx file.

        Returns:
            The absolute path to the saved .vsdx file.
        """
        import win32com.client

        # Ensure COM is initialized for the current thread
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception:
            pass

        try:
            self._app = win32com.client.Dispatch("Visio.Application")
            self._app.Visible = False
            self._app.AlertResponse = 6  # IDYES — auto-confirm all prompts

            # Create new document
            self._doc = self._app.Documents.Add("")
            self._page = self._app.ActivePage
            self._page.Name = state.name

            # Strip Visio theme
            try:
                self._doc.RemoveTheme()
            except Exception:
                pass

            # ── Calculate content bounding box ────────────────────
            margin = 2.0
            all_coords = []
            for r in state.resources.values():
                all_coords.append((r.position.x, r.position.y))
            for b in state.boundaries.values():
                all_coords.append((b.position.x, b.position.y))
                all_coords.append((b.position.x + b.size.width, b.position.y + b.size.height))

            if all_coords:
                max_x = max(c[0] for c in all_coords) + margin
                max_y = max(c[1] for c in all_coords) + margin
            else:
                max_x = state.page_width
                max_y = state.page_height

            page_w = max(max_x + margin, 11.0)
            page_h = max(max_y + margin, 8.5)

            # Set page size using named cells
            self._page.PageSheet.Cells("PageWidth").FormulaU = f"{page_w} in"
            self._page.PageSheet.Cells("PageHeight").FormulaU = f"{page_h} in"

            # Store for Y-flip
            self._page_height = page_h

            # ── White background page ─────────────────────────────
            try:
                bg_page = self._doc.Pages.Add()
                bg_page.Name = "VisioBG"
                bg_page.Background = -1
                bg_page.PageSheet.Cells("PageWidth").FormulaU = f"{page_w} in"
                bg_page.PageSheet.Cells("PageHeight").FormulaU = f"{page_h} in"
                bg_rect = bg_page.DrawRectangle(0, 0, page_w, page_h)
                bg_rect.Cells("FillPattern").FormulaU = "1"
                bg_rect.Cells("FillForegnd").FormulaU = "RGB(255,255,255)"
                bg_rect.Cells("FillForegndTrans").FormulaU = "0%"
                bg_rect.Cells("LinePattern").FormulaU = "0"
                self._page.BackPage = "VisioBG"
            except Exception as e:
                logger.warning("Could not create background page: %s", e)

            # Enable page-level right-angle routing for all connectors
            try:
                self._page.PageSheet.Cells("RouteStyle").FormulaU = "1"    # Right-angle
                self._page.PageSheet.Cells("PlaceStyle").FormulaU = "1"    # Automatic
            except Exception:
                pass

            # Open stencils
            if self._stencil_dir:
                self._open_stencils(state)

            # Draw boundaries (behind resources) — sorted largest-first
            # so parent containers are drawn first and end up behind children
            sorted_boundaries = sorted(
                state.boundaries.values(),
                key=lambda b: b.size.width * b.size.height,
                reverse=True,
            )
            for boundary in sorted_boundaries:
                try:
                    self._draw_boundary_com(boundary)
                except Exception as e:
                    logger.warning("Skipping boundary %s: %s", boundary.id, e)

            # Draw resources
            for resource in state.resources.values():
                try:
                    self._draw_resource_com(resource)
                except Exception as e:
                    logger.warning("Skipping resource %s: %s", resource.id, e)

            # Draw connections
            for connection in state.connections.values():
                try:
                    self._draw_connection_com(connection)
                except Exception as e:
                    logger.warning("Skipping connection %s: %s", connection.id, e)

            # Title
            try:
                self._add_title_com(state.name)
            except Exception as e:
                logger.warning("Could not add title: %s", e)

            # Save (do NOT call AutoSizeDrawing — page is already properly sized)
            self._doc.SaveAs(output_path)
            logger.info("Diagram saved to %s (COM)", output_path)
            return output_path

        finally:
            self.close()

    def _open_stencils(self, state: DiagramState) -> None:
        """Open .vssx stencil files if stencil_dir is provided."""
        if not self._stencil_dir:
            return
        needed_files = set()
        for res in state.resources.values():
            shape_info = AZURE_SHAPE_CATALOG.get(res.resource_type)
            if shape_info:
                needed_files.add(shape_info.stencil_file)

        for stencil_file in needed_files:
            stencil_path = os.path.join(self._stencil_dir, stencil_file)
            if os.path.exists(stencil_path):
                try:
                    stencil = self._app.Documents.OpenEx(
                        stencil_path, 4  # visOpenDocked
                    )
                    self._stencils[stencil_file] = stencil
                except Exception as e:
                    logger.warning("Could not open stencil %s: %s", stencil_file, e)

    def _draw_resource_com(self, resource) -> None:
        """Draw a single resource shape using COM with named cells.

        Strategy priority:
          0. Preserved original style (from image import)
          1. Visio stencil master (.vssx)
          2. SVG icon import + label below
          3. Colored rectangle fallback
        """
        # Strategy 0: Preserved original style from image import
        if resource.properties.get("preserve_style"):
            self._draw_preserved_resource_com(resource)
            return

        shape_info = AZURE_SHAPE_CATALOG.get(resource.resource_type)
        shape = None
        svg_used = False
        vx = resource.position.x
        vy = self._flip_y(resource.position.y)
        w = resource.size.width
        h = resource.size.height
        icon_size = MICROSOFT_STANDARDS.resource_icon_size  # 0.6"

        # Strategy 1: .vssx stencil master
        if shape_info and shape_info.stencil_file in self._stencils:
            stencil = self._stencils[shape_info.stencil_file]
            try:
                master = stencil.Masters.ItemU(shape_info.stencil_name)
                shape = self._page.Drop(master, vx, vy)
            except Exception:
                pass

        # Strategy 2: SVG icon import
        if shape is None:
            svg_path = resolve_svg_path(resource.resource_type, self._icons_root)
            if svg_path and svg_path.exists():
                icon_y = vy + 0.15  # Shift up to leave room for label below
                icon_shape = self._import_svg_as_shape(
                    str(svg_path), vx, icon_y, icon_size, icon_size
                )
                if icon_shape is not None:
                    # Add text label below the icon
                    label_w = max(w, 1.2)
                    label_h = 0.35
                    label_y = vy - icon_size / 2 - 0.05
                    label = self._page.DrawRectangle(
                        vx - label_w / 2, label_y - label_h,
                        vx + label_w / 2, label_y,
                    )
                    label.Text = resource.display_name
                    try:
                        label.Cells("FillPattern").FormulaU = "0"
                        label.Cells("LinePattern").FormulaU = "0"
                        label.Cells("Char.Size").FormulaU = (
                            f"{MICROSOFT_STANDARDS.label_font_size} pt"
                        )
                        label.Cells("Char.Color").FormulaU = (
                            f"RGB({_hex_to_rgb(AZURE_DIAGRAM_COLORS['label_primary'])})"
                        )
                        label.Cells("Char.Font").FormulaU = "\"Segoe UI\""
                        label.Cells("Para.HorzAlign").FormulaU = "1"
                    except Exception:
                        pass
                    shape = icon_shape
                    svg_used = True

        # Strategy 3: Colored rectangle fallback
        if shape is None:
            shape = self._page.DrawRectangle(
                vx - w / 2, vy - h / 2,
                vx + w / 2, vy + h / 2,
            )
            color = shape_info.icon_color if shape_info else "#0078D4"
            try:
                shape.Cells("FillPattern").FormulaU = "1"
                shape.Cells("FillForegnd").FormulaU = f"RGB({_hex_to_rgb(color)})"
                shape.Cells("FillForegndTrans").FormulaU = "30%"
                shape.Cells("LineColor").FormulaU = f"RGB({_hex_to_rgb(color)})"
                shape.Cells("LineWeight").FormulaU = "0.5 pt"
                shape.Cells("Rounding").FormulaU = "0.05 in"
            except Exception:
                pass

        # Label — stencil/rectangle only; SVG already has a separate label shape
        if not svg_used:
            shape.Text = resource.display_name
            try:
                shape.Cells("Char.Size").FormulaU = (
                    f"{MICROSOFT_STANDARDS.label_font_size} pt"
                )
                shape.Cells("Char.Color").FormulaU = (
                    f"RGB({_hex_to_rgb(AZURE_DIAGRAM_COLORS['label_primary'])})"
                )
                shape.Cells("Char.Font").FormulaU = "\"Segoe UI\""
            except Exception:
                pass

        self._shapes[resource.id] = shape

    def _draw_preserved_resource_com(self, resource) -> None:
        """Draw a resource with its original preserved visual style (COM)."""
        props = resource.properties
        shape_type = props.get("original_shape", "rectangle")
        fill = props.get("fill_color", "#FFFFFF")
        border = props.get("border_color", "#000000")
        text_color = props.get("text_color", "#000000")
        vx = resource.position.x
        vy = self._flip_y(resource.position.y)
        w = resource.size.width
        h = resource.size.height

        if shape_type == "diamond":
            # Draw a diamond by rotating a square 45 degrees
            shape = self._page.DrawRectangle(
                vx - w / 2, vy - h / 2, vx + w / 2, vy + h / 2
            )
            try:
                shape.Cells("Angle").FormulaU = "45 deg"
            except Exception:
                pass
        elif shape_type == "circle":
            shape = self._page.DrawOval(
                vx - w / 2, vy - h / 2, vx + w / 2, vy + h / 2
            )
        elif shape_type in ("cylinder",):
            # Approximate cylinder with an oval on top of a rectangle
            shape = self._page.DrawRectangle(
                vx - w / 2, vy - h / 2, vx + w / 2, vy + h / 2
            )
            try:
                shape.Cells("Rounding").FormulaU = "0.15 in"
            except Exception:
                pass
        elif shape_type == "hexagon":
            shape = self._page.DrawRectangle(
                vx - w / 2, vy - h / 2, vx + w / 2, vy + h / 2
            )
            try:
                shape.Cells("Rounding").FormulaU = "0.1 in"
            except Exception:
                pass
        elif shape_type == "rounded_rectangle":
            shape = self._page.DrawRectangle(
                vx - w / 2, vy - h / 2, vx + w / 2, vy + h / 2
            )
            try:
                shape.Cells("Rounding").FormulaU = "0.1 in"
            except Exception:
                pass
        else:
            # Default: rectangle
            shape = self._page.DrawRectangle(
                vx - w / 2, vy - h / 2, vx + w / 2, vy + h / 2
            )

        # Apply colors
        try:
            shape.Cells("FillPattern").FormulaU = "1"
            shape.Cells("FillForegnd").FormulaU = f"RGB({_hex_to_rgb(fill)})"
            shape.Cells("LineColor").FormulaU = f"RGB({_hex_to_rgb(border)})"
            shape.Cells("LineWeight").FormulaU = "1 pt"
        except Exception:
            pass

        # Label
        shape.Text = resource.display_name
        try:
            shape.Cells("Char.Size").FormulaU = "10 pt"
            shape.Cells("Char.Color").FormulaU = f"RGB({_hex_to_rgb(text_color)})"
            shape.Cells("Char.Font").FormulaU = '"Segoe UI"'
        except Exception:
            pass

        self._shapes[resource.id] = shape

    def _embed_background_image_com(self, page_w: float, page_h: float) -> None:
        """Embed the source image as a background trace on the page (COM)."""
        img_path = self._source_image
        if not img_path or not os.path.isfile(img_path):
            return

        # Import image as a shape centered on the page
        shape = self._page.Import(img_path)
        if shape is None:
            return

        # Position and size the image to fill the page (with margin)
        margin = 1.0
        img_w = page_w - margin * 2
        img_h = page_h - margin * 2

        try:
            shape.Cells("PinX").FormulaU = f"{page_w / 2} in"
            shape.Cells("PinY").FormulaU = f"{page_h / 2} in"
            shape.Cells("Width").FormulaU = f"{img_w} in"
            shape.Cells("Height").FormulaU = f"{img_h} in"
            # Lock the image so it doesn't get accidentally moved
            shape.Cells("LockSelect").FormulaU = "1"
            shape.Cells("LockPosition").FormulaU = "1"
            shape.Cells("LockSize").FormulaU = "1"
            # Make transparent so editable shapes stand out on top
            shape.Cells("FillForegndTrans").FormulaU = "70%"
        except Exception as e:
            logger.warning("Could not configure background image: %s", e)

        # Send to back so it's behind all other shapes
        try:
            shape.SendToBack()
        except Exception:
            pass

    def _import_svg_as_shape(
        self, svg_path: str, x: float, y: float, width: float, height: float
    ) -> object | None:
        """Import an SVG file into the Visio page and position/size it.

        Uses Page.Import for SVG-to-shape conversion, then positions and
        sizes using named cells (not CellsSRC indices).
        """
        try:
            shape = self._page.Import(svg_path)

            # Position and size using named cells
            shape.Cells("PinX").FormulaU = f"{x} in"
            shape.Cells("PinY").FormulaU = f"{y} in"
            shape.Cells("Width").FormulaU = f"{width} in"
            shape.Cells("Height").FormulaU = f"{height} in"

            # Lock aspect ratio to prevent icon distortion
            try:
                shape.Cells("LockAspect").FormulaU = "1"
            except Exception:
                pass

            return shape
        except Exception as e:
            logger.warning("SVG import failed for %s: %s", svg_path, e)
            return None

    # Microsoft Architecture Center boundary styling — per actual published SVG CSS
    # MS uses GRAY/WHITE fills, gray borders, dashed lines, very thin strokes.
    # CSS classes: st2 (#FFF, 79% opacity, #7F7F7F dashed 0.5pt),
    #              st4 (#F2F2F2, #BFBFBF solid 1pt),
    #              st91 (#F2F2F2, 79% opacity, #7F7F7F dashed 0.5pt)
    # type: (fill_hex, fill_trans%, line_hex, line_weight_pt, line_pattern, corner_radius, label_color)
    _BOUNDARY_VISUAL = {
        "subscription":      ("#FFFFFF", 79, "#7F7F7F", 0.5, 2, 0.0, "#5B9BD5"),  # st2: dashed
        "management_group":  ("#F8F8F8", 79, "#7F7F7F", 0.5, 2, 0.0, "#5B9BD5"),  # dashed
        "resource_group":    ("#F2F2F2", 55, "#BFBFBF", 1.0, 1, 0.0, "#5B9BD5"),  # st4: solid
        "vnet":              ("#FFFFFF", 79, "#7F7F7F", 0.5, 2, 0.0, "#5B9BD5"),  # st2: dashed
        "subnet":            ("#F2F2F2", 79, "#7F7F7F", 0.5, 2, 0.0, "#5B9BD5"),  # st91: dashed
        "availability_zone": ("#F2F2F2", 70, "#BFBFBF", 0.5, 2, 0.0, "#5B9BD5"),  # dashed
        "region":            ("#F8F8F8", 79, "#BFBFBF", 0.24,1, 0.0, "#0A609E"),  # st6: solid, dark blue label
        "nsg":               ("#F2F2F2", 70, "#7F7F7F", 0.5, 2, 0.0, "#5B9BD5"),  # dashed
    }
    _DEFAULT_BOUNDARY_VISUAL = ("#F2F2F2", 79, "#7F7F7F", 0.5, 2, 0.0, "#5B9BD5")

    def _draw_boundary_com(self, boundary) -> None:
        """Draw a boundary/container rectangle per Microsoft Architecture Center standards.

        Each boundary type gets a distinct fill color, transparency, line style,
        and label color so they are visually distinguishable at a glance.
        """
        vis = self._BOUNDARY_VISUAL.get(
            boundary.boundary_type, self._DEFAULT_BOUNDARY_VISUAL
        )
        fill_hex, fill_trans, line_hex, line_wt, line_pat, corner_r, label_color = vis

        x1 = boundary.position.x
        x2 = x1 + boundary.size.width
        vy1 = self._flip_y(boundary.position.y + boundary.size.height)
        vy2 = self._flip_y(boundary.position.y)

        shape = self._page.DrawRectangle(x1, vy1, x2, vy2)

        try:
            # ── Fill ──
            shape.Cells("FillPattern").FormulaU = "1"
            shape.Cells("FillForegnd").FormulaU = f"RGB({_hex_to_rgb(fill_hex)})"
            shape.Cells("FillForegndTrans").FormulaU = f"{fill_trans}%"

            # ── Border ──
            shape.Cells("LineColor").FormulaU = f"RGB({_hex_to_rgb(line_hex)})"
            shape.Cells("LineWeight").FormulaU = f"{line_wt} pt"
            shape.Cells("LinePattern").FormulaU = str(line_pat)
            shape.Cells("Rounding").FormulaU = f"{corner_r} in"

            # ── Label: top-left aligned, blue-gray per MS CSS st10 (#5B9BD5 bold) ──
            shape.Cells("VerticalAlign").FormulaU = "0"    # Top
            shape.Cells("Para.HorzAlign").FormulaU = "0"   # Left
            shape.Cells("TxtPinY").FormulaU = "Height - 0.15 in"  # Pin near top
            shape.Cells("TxtLocPinY").FormulaU = "TxtHeight"      # Anchor at top of text
            shape.Cells("LeftMargin").FormulaU = "0.15 in"
            shape.Cells("TopMargin").FormulaU = "0.08 in"
            shape.Cells("Char.Size").FormulaU = f"{MICROSOFT_STANDARDS.boundary_label_font_size} pt"
            shape.Cells("Char.Color").FormulaU = f"RGB({_hex_to_rgb(label_color)})"
            shape.Cells("Char.Style").FormulaU = "1"  # Bold
            # Set font to Segoe UI (MS Architecture Center standard)
            shape.Cells("Char.Font").FormulaU = "\"Segoe UI\""

            shape.SendToBack()
        except Exception as e:
            logger.warning("Could not style boundary %s: %s", boundary.id, e)

        # Set text AFTER styling to avoid style overwrite
        shape.Text = boundary.display_name
        self._shapes[boundary.id] = shape

    def _draw_connection_com(self, connection) -> None:
        """Draw a connector between two shapes using Microsoft Architecture Center conventions.

        Connectors use dynamic glue (BeginX→PinX, EndX→PinX) for automatic
        routing. Falls back to manual endpoint positioning if glue fails,
        and to a simple line as a last resort.

        Includes:
          - Right-angle routing (Visio RouteStyle)
          - Arrow endpoint on target end
          - All-black line color (#000000) per MS Architecture Center CSS
          - Line pattern per connection type (solid/dashed/dotted)
          - Workflow step number circle at midpoint (if label starts with '(N)')

        Args:
            connection: A Connection model instance with source_id, target_id,
                       label, connection_type, and style attributes.
        """
        source_shape = self._shapes.get(connection.source_id)
        target_shape = self._shapes.get(connection.target_id)
        if not source_shape or not target_shape:
            logger.warning(
                "Cannot draw connection %s: missing shape(s)", connection.id
            )
            return

        conn_style = CONNECTOR_STYLES.get(connection.connection_type, CONNECTOR_STYLES["data_flow"])

        # MS Architecture Center: ALL connectors are BLACK (st11/st46: stroke:#000000)
        line_hex = "#000000"

        # Create a proper dynamic connector and glue to source/target shapes.
        # Use named cells (never CellsSRC indices) for reliability with SVG groups.
        connector = None
        glued = False
        try:
            connector = self._page.Drop(
                self._app.ConnectorToolDataObject, 0, 0
            )
            # Glue using named cells — works with SVG-imported group shapes
            connector.Cells("BeginX").GlueTo(source_shape.Cells("PinX"))
            connector.Cells("EndX").GlueTo(target_shape.Cells("PinX"))
            glued = True
        except Exception:
            pass

        # If glue failed, still use a dynamic connector but position endpoints manually
        if not glued:
            try:
                if connector is None:
                    connector = self._page.Drop(
                        self._app.ConnectorToolDataObject, 0, 0
                    )
                sx = source_shape.Cells("PinX").ResultIU
                sy = source_shape.Cells("PinY").ResultIU
                tx = target_shape.Cells("PinX").ResultIU
                ty = target_shape.Cells("PinY").ResultIU
                connector.Cells("BeginX").FormulaU = f"{sx} in"
                connector.Cells("BeginY").FormulaU = f"{sy} in"
                connector.Cells("EndX").FormulaU = f"{tx} in"
                connector.Cells("EndY").FormulaU = f"{ty} in"
            except Exception:
                # Last resort: draw a line (no routing)
                sx = source_shape.Cells("PinX").ResultIU
                sy = source_shape.Cells("PinY").ResultIU
                tx = target_shape.Cells("PinX").ResultIU
                ty = target_shape.Cells("PinY").ResultIU
                connector = self._page.DrawLine(sx, sy, tx, ty)

        # Style per Microsoft conventions
        try:
            line_rgb = _hex_to_rgb(line_hex)
            connector.Cells("LineColor").FormulaU = f"RGB({line_rgb})"
            connector.Cells("LineWeight").FormulaU = f"{MICROSOFT_STANDARDS.connector_width} pt"

            # Arrow endpoint on target end (Architecture Center standard)
            connector.Cells("EndArrow").FormulaU = "4"  # Open arrowhead
            connector.Cells("EndArrowSize").FormulaU = "2"  # Medium

            # Right-angle routing (Architecture Center uses orthogonal connectors)
            try:
                connector.Cells("ShapeRouteStyle").FormulaU = "1"  # Right-angle
                connector.Cells("ConLineRouteExt").FormulaU = "1"  # Right-angle
            except Exception:
                pass

            if connection.connection_type in ("dependency", "reference"):
                connector.Cells("LinePattern").FormulaU = "2"
            elif connection.connection_type == "private_link":
                connector.Cells("LinePattern").FormulaU = "3"
            elif conn_style["pattern"] == "dashed":
                connector.Cells("LinePattern").FormulaU = "2"
            elif conn_style["pattern"] == "dotted":
                connector.Cells("LinePattern").FormulaU = "3"
        except Exception as e:
            logger.warning("Could not style connector %s: %s", connection.id, e)

        # Extract workflow step number from label like "(3) Some label"
        step_num = None
        display_label = connection.label or ""
        import re
        step_match = re.match(r"^\((\d+)\)\s*(.*)", display_label)
        if step_match:
            step_num = int(step_match.group(1))
            display_label = step_match.group(2).strip()

        # Set connector label (stripped of step number prefix)
        if display_label:
            connector.Text = display_label
            try:
                connector.Cells("Char.Size").FormulaU = "7 pt"
                connector.Cells("Char.Color").FormulaU = (
                    f"RGB({_hex_to_rgb(AZURE_DIAGRAM_COLORS['label_secondary'])})"
                )
                connector.Cells("Char.Font").FormulaU = "\"Segoe UI\""
            except Exception:
                pass

        self._shapes[connection.id] = connector

        # Draw numbered workflow step circle at connector midpoint
        if step_num is not None:
            try:
                self._draw_step_circle(connector, step_num)
            except Exception as e:
                logger.warning("Could not draw step circle %d: %s", step_num, e)

    def _draw_step_circle(self, connector, step_num: int) -> None:
        """Draw a numbered workflow step circle at the midpoint of a connector.

        Per Microsoft Architecture Center SVG CSS:
          - st96: fill:#107c10 (GREEN circle, no stroke)
          - st97: fill:#ffffff font-family:Segoe UI font-weight:bold (white number)
        """
        # Get connector midpoint
        try:
            bx = connector.Cells("BeginX").ResultIU
            by = connector.Cells("BeginY").ResultIU
            ex = connector.Cells("EndX").ResultIU
            ey = connector.Cells("EndY").ResultIU
            mx = (bx + ex) / 2
            my = (by + ey) / 2
        except Exception:
            return

        radius = MICROSOFT_STANDARDS.step_number_radius  # 0.2"
        diameter = radius * 2

        circle = self._page.DrawOval(
            mx - radius, my - radius,
            mx + radius, my + radius,
        )
        circle.Text = str(step_num)

        try:
            # GREEN filled circle — MS Architecture Center standard (CSS st96: #107C10)
            circle.Cells("FillPattern").FormulaU = "1"
            circle.Cells("FillForegnd").FormulaU = "RGB(16,124,16)"
            circle.Cells("FillForegndTrans").FormulaU = "0%"

            # No border stroke (MS CSS has no stroke on step circles)
            circle.Cells("LinePattern").FormulaU = "0"

            # White bold number text — Segoe UI (CSS st97)
            circle.Cells("Char.Color").FormulaU = "RGB(255,255,255)"
            circle.Cells("Char.Size").FormulaU = "9 pt"
            circle.Cells("Char.Style").FormulaU = "1"  # Bold
            circle.Cells("Char.Font").FormulaU = "\"Segoe UI\""
            circle.Cells("Para.HorzAlign").FormulaU = "1"  # Center
            circle.Cells("VerticalAlign").FormulaU = "1"  # Middle

            # Ensure circle is on top of connector
            circle.BringToFront()
        except Exception as e:
            logger.warning("Could not style step circle %d: %s", step_num, e)

    def _add_title_com(self, title: str) -> None:
        """Add a title text block at the top of the page."""
        top_y = self._page_height - 0.5
        shape = self._page.DrawRectangle(0.5, top_y - 0.6, 10.0, top_y)
        shape.Text = title
        try:
            shape.Cells("FillForegndTrans").FormulaU = "100%"
            shape.Cells("LinePattern").FormulaU = "0"
            shape.Cells("Char.Size").FormulaU = f"{MICROSOFT_STANDARDS.title_font_size} pt"
            shape.Cells("Char.Style").FormulaU = "1"  # Bold
            shape.Cells("Char.Color").FormulaU = f"RGB({_hex_to_rgb(AZURE_DIAGRAM_COLORS['label_primary'])})"
            shape.Cells("Char.Font").FormulaU = "\"Segoe UI\""
        except Exception:
            pass

    # ── python-vsdx based rendering (no Visio required) ───────────

    def _render_vsdx(self, state: DiagramState, output_path: str) -> str:
        """Render using python-vsdx when Visio is not installed.

        Provides a basic .vsdx output with colored rectangles for boundaries
        and resources. Does not support SVG icon import, step circles, or
        Microsoft Architecture Center visual conventions. Suitable for
        headless/Linux environments.

        Args:
            state: The complete diagram state to render.
            output_path: Absolute path for the output .vsdx file.

        Returns:
            The absolute path to the saved .vsdx file.

        Raises:
            RuntimeError: If neither Visio COM nor python-vsdx is available.
        """
        try:
            import vsdx
        except ImportError:
            raise RuntimeError(
                "Neither Microsoft Visio (COM) nor python-vsdx is available. "
                "Install vsdx (`pip install vsdx`) or use a machine with Visio installed."
            )

        with vsdx.VisioFile() as vis:
            page = vis.pages[0]
            page.name = state.name

            # Track shapes for potential later use
            shape_map: dict[str, object] = {}

            # Draw boundaries as rectangles
            for boundary in state.boundaries.values():
                style = BOUNDARY_STYLES.get(
                    boundary.boundary_type, BOUNDARY_STYLES["resource_group"]
                )
                shape = page.add_shape(
                    x=boundary.position.x + boundary.size.width / 2,
                    y=boundary.position.y + boundary.size.height / 2,
                    width=boundary.size.width,
                    height=boundary.size.height,
                    text=boundary.display_name,
                )
                if shape is not None:
                    shape_map[boundary.id] = shape
                    # Apply fill color
                    try:
                        shape.set_cell_value("FillForegnd", style["fill_color"])
                        shape.set_cell_value("LineColor", style["line_color"])
                    except Exception:
                        pass

            # Draw resources as shapes
            for resource in state.resources.values():
                preserved = resource.properties.get("preserve_style")
                if preserved:
                    fill = resource.properties.get("fill_color", "#FFFFFF")
                    border = resource.properties.get("border_color", "#000000")
                    label = resource.display_name
                else:
                    shape_info = AZURE_SHAPE_CATALOG.get(resource.resource_type)
                    label = resource.display_name
                    if shape_info:
                        label = f"{shape_info.display_name}\n{resource.display_name}"

                shape = page.add_shape(
                    x=resource.position.x,
                    y=resource.position.y,
                    width=resource.size.width,
                    height=resource.size.height,
                    text=label,
                )
                if shape is not None:
                    shape_map[resource.id] = shape
                    if preserved:
                        try:
                            shape.set_cell_value("FillForegnd", fill)
                            shape.set_cell_value("LineColor", border)
                        except Exception:
                            pass
                    elif shape_info:
                        try:
                            shape.set_cell_value("FillForegnd", shape_info.icon_color)
                        except Exception:
                            pass

            # Draw connections as lines
            for connection in state.connections.values():
                src = state.resources.get(connection.source_id)
                tgt = state.resources.get(connection.target_id)
                if src and tgt:
                    try:
                        connector = page.add_connect(
                            shape_a=shape_map.get(connection.source_id),
                            shape_b=shape_map.get(connection.target_id),
                            text=connection.label,
                        )
                    except Exception:
                        # Fallback: draw a line
                        try:
                            page.add_shape(
                                x=(src.position.x + tgt.position.x) / 2,
                                y=(src.position.y + tgt.position.y) / 2,
                                width=0.1,
                                height=0.1,
                                text=connection.label or "→",
                            )
                        except Exception:
                            pass

            vis.save_vsdx(output_path)
            logger.info("Diagram saved to %s (python-vsdx)", output_path)
            return output_path


# ── Utilities ─────────────────────────────────────────────────────

def _hex_to_rgb(hex_color: str) -> str:
    """Convert '#RRGGBB' to 'R,G,B' string."""
    hex_color = hex_color.lstrip("#")
    r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
    return f"{r},{g},{b}"
