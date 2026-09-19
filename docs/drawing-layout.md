# Drawing and Normal-font metadata

The maintained parser is the single owner of XLSX package relationships and
DrawingML coordinates. Renderers must not reopen the ZIP to repair missing fields.

`SheetImage.anchor` now preserves the container `kind` (`twoCell`, `oneCell`,
`absolute`), an explicitly saved `editAs`, exact EMU `extent` and (for absolute
placement) EMU `position`. Existing `from`/`to` markers and offsets remain intact.
Absent metadata remains optional for older caller-created models. `width` and
`height` are compatibility dimensions at 96 DPI, without premature integer rounding.
One inch is 914400 EMU; DPI, display zoom, hidden-axis layout, and browser coordinate
conversion belong to the renderer, not the parser.

A fixed-size placement must not be stretched just because it has a second saved
cell marker. A normal two-cell placement still derives its current displayed size
from the cell markers. The writer preserves container kind, edit behavior, marker
offsets and exact extents; cloning and worker structured cloning retain them too.
This change does not claim full support for grouped-shape transforms or image effects.

`Workbook.defaultFont` follows the built-in Normal style's reference chain:
`cellStyles[builtinId=0].xfId -> cellStyleXfs[xfId].fontId -> fonts[fontId]`.
It is not inferred from a cell or from a screenshot. Invalid references fall back
to the first font, as before. View, page setup, print area and page-break metadata
continue to come from the existing parser.

Run `npm run test:drawing-layout` for the post-build parser/writer/clone regression
checks. Fixtures are generated in memory and contain no customer data. A `prepare`
lifecycle script builds the package when installed through a pinned Git dependency;
consumers do not need a separately installed pnpm executable to run that build.
