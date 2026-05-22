"""
AutoCAD MCP Server — controla AutoCAD desde Claude via COM/ActiveX.
Requiere Windows + AutoCAD abierto + pywin32 instalado.

Capacidades:
- Crear y modificar geometría (líneas, círculos, arcos, polilíneas, etc.)
- Editar textos: contenido, estilo, altura, rotación, alineación
- Buscar y reemplazar texto en todo el dibujo
- Análisis del dibujo: estadísticas por capa, bbox, conteos
- Cálculos de superficie: por entidad, por capa, sumatorias, regiones
- Identificación de boundaries cerrados y cómputos
- Operaciones booleanas con regiones (unión, resta, intersección)
"""

import sys
import math
from contextlib import contextmanager
from typing import Any
from mcp.server.fastmcp import FastMCP

try:
    import win32com.client as win32
    import pythoncom
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

mcp = FastMCP("autocad-mcp")

# Cache de la conexión COM a AutoCAD para evitar reconectar en cada tool call.
_ACAD_CACHE: dict[str, Any] = {"app": None}


# ============================================================================
# Helpers internos
# ============================================================================

# Mapeo de tipos ObjectName de AutoCAD a nombres amigables
ENTITY_TYPE_NAMES = {
    "AcDbLine": "Línea",
    "AcDbCircle": "Círculo",
    "AcDbArc": "Arco",
    "AcDbEllipse": "Elipse",
    "AcDbPolyline": "Polilínea",
    "AcDb2dPolyline": "Polilínea 2D",
    "AcDb3dPolyline": "Polilínea 3D",
    "AcDbSpline": "Spline",
    "AcDbText": "Texto",
    "AcDbMText": "MText",
    "AcDbAttributeDefinition": "Definición de Atributo",
    "AcDbAttribute": "Atributo",
    "AcDbBlockReference": "Bloque",
    "AcDbHatch": "Hatch",
    "AcDbRegion": "Región",
    "AcDbSolid": "Sólido",
    "AcDb3dSolid": "Sólido 3D",
    "AcDbRotatedDimension": "Cota Lineal",
    "AcDbAlignedDimension": "Cota Alineada",
    "AcDbRadialDimension": "Cota Radial",
    "AcDbDiametricDimension": "Cota Diametral",
    "AcDb2LineAngularDimension": "Cota Angular",
    "AcDbLeader": "Directriz",
    "AcDbMLeader": "Directriz Múltiple",
    "AcDbViewport": "Viewport",
}

# Unidades del dibujo (LUNITS)
UNITS_NAMES = {
    1: "Científico",
    2: "Decimal",
    3: "Ingeniería (pies/pulgadas)",
    4: "Arquitectónico (pies/pulgadas)",
    5: "Fraccional",
}

# Códigos ACI comunes
ACI_COLORS = {
    1: "Rojo", 2: "Amarillo", 3: "Verde", 4: "Cian",
    5: "Azul", 6: "Magenta", 7: "Blanco/Negro",
    8: "Gris oscuro", 9: "Gris claro", 256: "ByLayer", 0: "ByBlock",
}


def _get_acad():
    """Conecta a AutoCAD reutilizando la conexión COM cacheada cuando es posible."""
    if not WIN32_AVAILABLE:
        raise RuntimeError("pywin32 no instalado. Ejecuta: pip install pywin32")
    cached = _ACAD_CACHE.get("app")
    if cached is not None:
        try:
            _ = cached.Version  # ping para verificar que el objeto sigue vivo
            return cached
        except Exception:
            _ACAD_CACHE["app"] = None
    try:
        app = win32.GetActiveObject("AutoCAD.Application")
        _ACAD_CACHE["app"] = app
        return app
    except Exception:
        raise RuntimeError("AutoCAD no está abierto. Abrí AutoCAD primero.")


def _space(acad, in_paper: bool | None = None):
    """
    Devuelve el espacio en el que trabajar:
    - in_paper=True  → PaperSpace
    - in_paper=False → ModelSpace
    - in_paper=None  → el espacio activo según ActiveSpace del documento
    """
    doc = _active_doc(acad)
    if in_paper is None:
        # ActiveSpace: 0=Paper, 1=Model
        return doc.PaperSpace if doc.ActiveSpace == 0 else doc.ModelSpace
    return doc.PaperSpace if in_paper else doc.ModelSpace


@contextmanager
def _fast_batch(acad, regen_at_end: bool = True):
    """
    Context manager para operaciones en lote. Desactiva el redibujado/echo de
    comandos mientras se hacen muchas modificaciones y restaura al salir.

    Reduce drásticamente el tiempo de creación de muchas entidades.
    """
    doc = _active_doc(acad)
    prev_cmdecho = None
    try:
        try:
            prev_cmdecho = doc.GetVariable("CMDECHO")
            doc.SetVariable("CMDECHO", 0)
        except Exception:
            pass
        # Congelar la actualización gráfica
        try:
            acad.Application.Update()  # flush previo
        except Exception:
            pass
        yield
    finally:
        try:
            if prev_cmdecho is not None:
                doc.SetVariable("CMDECHO", prev_cmdecho)
        except Exception:
            pass
        if regen_at_end:
            try:
                # 1 = AllViewports
                doc.Regen(1)
            except Exception:
                pass


def _active_doc(acad):
    doc = acad.ActiveDocument
    if doc is None:
        raise RuntimeError("No hay ningún dibujo abierto en AutoCAD.")
    return doc


def _model_space(acad):
    return _active_doc(acad).ModelSpace


def _point(x: float, y: float, z: float = 0.0):
    return win32.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (float(x), float(y), float(z)))


def _double_array(values: list[float]):
    return win32.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [float(v) for v in values])


def _friendly_type(object_name: str) -> str:
    return ENTITY_TYPE_NAMES.get(object_name, object_name)


def _entity_brief(entity) -> dict:
    """Resumen breve de una entidad."""
    return {
        "handle": entity.Handle,
        "type": entity.ObjectName,
        "type_friendly": _friendly_type(entity.ObjectName),
        "layer": entity.Layer,
    }


def _entity_detail(entity) -> dict:
    """Detalles completos de una entidad según su tipo."""
    info: dict[str, Any] = _entity_brief(entity)
    try:
        info["visible"] = entity.Visible
        info["color"] = entity.color
    except Exception:
        pass

    name = entity.ObjectName

    if name == "AcDbLine":
        sp, ep = entity.StartPoint, entity.EndPoint
        info["start"] = {"x": sp[0], "y": sp[1], "z": sp[2]}
        info["end"] = {"x": ep[0], "y": ep[1], "z": ep[2]}
        info["length"] = entity.Length
        info["angle_deg"] = math.degrees(entity.Angle)

    elif name == "AcDbCircle":
        c = entity.Center
        info["center"] = {"x": c[0], "y": c[1], "z": c[2]}
        info["radius"] = entity.Radius
        info["area"] = entity.Area
        info["circumference"] = entity.Circumference

    elif name == "AcDbArc":
        c = entity.Center
        info["center"] = {"x": c[0], "y": c[1], "z": c[2]}
        info["radius"] = entity.Radius
        info["start_angle_deg"] = math.degrees(entity.StartAngle)
        info["end_angle_deg"] = math.degrees(entity.EndAngle)
        info["arc_length"] = entity.ArcLength

    elif name == "AcDbEllipse":
        c = entity.Center
        info["center"] = {"x": c[0], "y": c[1], "z": c[2]}
        try:
            info["major_radius"] = entity.MajorRadius
            info["minor_radius"] = entity.MinorRadius
            info["area"] = entity.Area
        except Exception:
            pass

    elif name in ("AcDbPolyline", "AcDb2dPolyline", "AcDb3dPolyline"):
        try:
            info["closed"] = bool(entity.Closed)
        except Exception:
            pass
        try:
            info["length"] = entity.Length
        except Exception:
            pass
        try:
            if info.get("closed"):
                info["area"] = entity.Area
        except Exception:
            pass
        try:
            coords = list(entity.Coordinates)
            stride = 2 if name == "AcDbPolyline" else 3
            pts = []
            for i in range(0, len(coords), stride):
                pt = {"x": coords[i], "y": coords[i+1]}
                if stride == 3:
                    pt["z"] = coords[i+2]
                pts.append(pt)
            info["points"] = pts
        except Exception:
            pass

    elif name == "AcDbText":
        info["text"] = entity.TextString
        ip = entity.InsertionPoint
        info["insertion"] = {"x": ip[0], "y": ip[1], "z": ip[2]}
        info["height"] = entity.Height
        info["rotation_deg"] = math.degrees(entity.Rotation)
        try:
            info["style"] = entity.StyleName
        except Exception:
            pass
        try:
            info["alignment"] = entity.Alignment
        except Exception:
            pass

    elif name == "AcDbMText":
        info["text"] = entity.TextString
        ip = entity.InsertionPoint
        info["insertion"] = {"x": ip[0], "y": ip[1], "z": ip[2]}
        info["height"] = entity.Height
        info["width"] = entity.Width
        info["rotation_deg"] = math.degrees(entity.Rotation)
        try:
            info["style"] = entity.StyleName
        except Exception:
            pass

    elif name == "AcDbBlockReference":
        ip = entity.InsertionPoint
        info["block_name"] = entity.Name
        info["insertion"] = {"x": ip[0], "y": ip[1], "z": ip[2]}
        info["rotation_deg"] = math.degrees(entity.Rotation)
        info["scale"] = {
            "x": entity.XScaleFactor,
            "y": entity.YScaleFactor,
            "z": entity.ZScaleFactor,
        }
        try:
            attrs = []
            if entity.HasAttributes:
                for att in entity.GetAttributes():
                    attrs.append({"tag": att.TagString, "value": att.TextString})
            info["attributes"] = attrs
        except Exception:
            pass

    elif name == "AcDbHatch":
        try:
            info["area"] = entity.Area
            info["pattern"] = entity.PatternName
        except Exception:
            pass

    elif name == "AcDbRegion":
        try:
            info["area"] = entity.Area
            info["perimeter"] = entity.Perimeter
        except Exception:
            pass

    elif name in ("AcDbRotatedDimension", "AcDbAlignedDimension",
                  "AcDbRadialDimension", "AcDbDiametricDimension"):
        try:
            info["measurement"] = entity.Measurement
            info["text_override"] = entity.TextOverride
        except Exception:
            pass

    # Bounding box genérico
    try:
        bb_min, bb_max = entity.GetBoundingBox()
        info["bbox"] = {
            "min": {"x": bb_min[0], "y": bb_min[1], "z": bb_min[2]},
            "max": {"x": bb_max[0], "y": bb_max[1], "z": bb_max[2]},
        }
    except Exception:
        pass

    return info


def _safe_area(entity) -> float | None:
    """Intenta leer .Area de una entidad de forma segura."""
    try:
        return float(entity.Area)
    except Exception:
        return None


# ============================================================================
# Prompts — guías que Claude usa para interpretar pedidos
# ============================================================================

@mcp.prompt()
def autocad_context() -> str:
    """Contexto general para que Claude interprete pedidos de CAD."""
    return """Estás conectado a una sesión activa de AutoCAD via MCP. Cuando el usuario
te pida modificar el dibujo, seguí estos principios:

1. **Identificá primero, modificá después.** Antes de cambiar nada, usá
   `get_drawing_info` y/o `analyze_drawing` para entender qué hay. Si el
   usuario habla de "los muros", listá entidades de la capa MUROS (o similares)
   con `list_entities` antes de actuar.

2. **Manejá unidades.** Pedí o inferí las unidades del dibujo con
   `get_drawing_info`. Si el dibujo está en milímetros y el usuario pide
   "una línea de 3 metros", convertilo a 3000.

3. **Capas correctas.** Antes de crear geometría, verificá si existe la capa
   apropiada con `list_layers`. Si no existe, creala con `create_layer`.
   No mezcles geometría en la capa "0" salvo que el usuario lo pida.

4. **Editar textos.** Para encontrar texto usá `find_texts_containing`,
   y para reemplazar usá `edit_text_by_handle` o `find_and_replace_text`.
   Nunca borres y recrees un texto si podés editarlo en su lugar.

5. **Cómputos de superficie.** Para sacar áreas:
   - De una entidad cerrada concreta: `calculate_area_of_entity`.
   - Sumatoria por capa (ej: "área total de losas"): `calculate_area_by_layer`.
   - Boundary cerrado a partir de un punto interno: `calculate_boundary_area_at_point`.
   - Múltiples handles: `sum_areas`.

6. **Layouts y escalas.** Cuando el usuario pida "generar un layout" o
   "escalar a 1:50":
   - Trabajá con `list_layouts`, `create_layout`, `set_active_layout`,
     `setup_layout_page` (para fijar tamaño de hoja A4/A3/A1).
   - Insertá viewports con `create_viewport` indicando layout, centro,
     tamaño y `scale_ratio` ('1:50', '1:100', etc.) directamente —
     evitá fijar escala con SendCommand.
   - Para reajustar una escala existente usá `set_viewport_scale`.
   - Para que el modelo quepa entero usá `fit_viewport_to_extents`.
   - `lock_viewport` evita que se descalibre la escala.

7. **Rendimiento.** Cuando vayas a crear/modificar muchas entidades:
   - Llamá `set_performance_mode(True)` al empezar y `False` al terminar.
   - Usá las versiones batch: `create_lines_batch`, `create_polylines_batch`,
     `create_texts_batch`, `transform_batch`. Una sola llamada con 200
     elementos es 50x más rápida que 200 llamadas individuales.

8. **Confirmá cambios destructivos.** Si vas a borrar muchas entidades o
   modificar varias capas, mostrá primero qué encontraste y confirmá.

9. **Devolvé contexto al usuario.** Después de modificar, decí qué handles
   creaste o tocaste y en qué capa, así el usuario puede deshacer si quiere.

10. **No uses `run_autocad_command` salvo último recurso.** SendCommand puede
    quedar esperando input. Preferí las herramientas tipadas."""


@mcp.prompt()
def computar_superficies(descripcion: str = "todas las áreas") -> str:
    """Plantilla para guiar el cómputo de superficies."""
    return f"""El usuario pide computar superficies: "{descripcion}".

Procedé así:
1. Llamá a `analyze_drawing` para ver qué capas y entidades hay.
2. Identificá qué capas representan superficies a computar (ej: LOSAS, MUROS,
   SUPERFICIES, AREAS).
3. Para cada capa relevante, llamá a `calculate_area_by_layer`.
4. Si el usuario indica un punto interno de una habitación o región sin
   polilínea cerrada explícita, usá `calculate_boundary_area_at_point`.
5. Presentá los resultados en una tabla con: descripción, capa, cantidad de
   entidades, área individual (si aplica), área total. Aclarar las unidades
   (las del dibujo elevadas al cuadrado).
6. Si las unidades del dibujo son mm, ofrecé también el total en m²
   (dividir por 1.000.000)."""


@mcp.prompt()
def armar_layout(descripcion: str = "") -> str:
    """Plantilla para armar layouts (presentaciones) con escala."""
    return f"""El usuario quiere armar un layout: "{descripcion}".

Procedé en este orden, sin atajos:

1. **Activá performance mode**: `set_performance_mode(True)`.

2. **Verificá unidades**: `get_drawing_info`. Las hojas se piden en mm
   (A4 = 210x297, A3 = 297x420, A2 = 420x594, A1 = 594x841).

3. **Listá layouts existentes**: `list_layouts`. Reutilizá si ya hay uno
   con el nombre pedido; si no, creá con `create_layout`.

4. **Configurá la hoja**: `setup_layout_page` con el ancho/alto del formato
   solicitado y `plot_unit='mm'`.

5. **Calculá los extents del modelo**: leelos de `get_drawing_info` para
   saber qué tamaño tiene el dibujo y elegir la escala correcta.

6. **Insertá viewports**: para cada vista, llamá `create_viewport` con
   `layout`, `center_x/y` (en coordenadas de papel),
   `width`/`height` del viewport, `scale_ratio` (ej '1:50') y
   opcionalmente `model_target_x/y` (centro del área del modelo a mostrar).

7. **Bloqueá los viewports** importantes con `lock_viewport(handle, True)`
   para que no se descalibre la escala al hacer zoom.

8. **Agregá rotulación**: usá `set_active_layout` para asegurarte de estar
   en el layout, después `create_text` o `create_mtext` con coordenadas
   de papel para el cajetín/título.

9. **Desactivá performance mode**: `set_performance_mode(False)`.

10. **Reportá**: layout creado, viewports con handles y escalas, hoja
    configurada, total de tiempo aproximado. Sugerí los próximos pasos
    (impresión, exportar PDF) si aplica."""


@mcp.prompt()
def editar_texto(busqueda: str = "") -> str:
    """Plantilla para edición de textos."""
    return f"""El usuario pide editar texto{f' relacionado con "{busqueda}"' if busqueda else ''}.

Procedé así:
1. Si hay una búsqueda, usá `find_texts_containing` con la palabra clave.
2. Mostrale al usuario los handles, contenidos y capas encontrados.
3. Confirmá antes de modificar masivamente.
4. Para edición individual: `edit_text_by_handle`.
5. Para reemplazo global con patrón: `find_and_replace_text`.
6. Después de editar, listá los handles modificados con el contenido nuevo."""


# ============================================================================
# Estado y documentos
# ============================================================================

@mcp.tool()
def get_autocad_status() -> dict:
    """Verifica si AutoCAD está corriendo y devuelve info básica."""
    try:
        acad = _get_acad()
        doc = acad.ActiveDocument
        docs = [d.Name for d in acad.Documents]
        return {
            "running": True,
            "version": acad.Version,
            "active_document": doc.Name if doc else None,
            "active_document_path": doc.FullName if doc else None,
            "open_documents": docs,
        }
    except Exception as exc:
        return {"running": False, "error": str(exc)}


@mcp.tool()
def open_drawing(path: str) -> dict:
    """Abre un archivo DWG en AutoCAD."""
    acad = _get_acad()
    doc = acad.Documents.Open(path)
    return {"opened": doc.Name, "path": doc.FullName}


@mcp.tool()
def save_drawing(path: str = "") -> dict:
    """Guarda el dibujo activo. Si se indica path, hace 'Save As'."""
    acad = _get_acad()
    doc = _active_doc(acad)
    if path:
        doc.SaveAs(path)
    else:
        doc.Save()
    return {"saved": True, "path": doc.FullName}


@mcp.tool()
def close_drawing(save: bool = True) -> dict:
    """Cierra el dibujo activo. save=True guarda antes."""
    acad = _get_acad()
    doc = _active_doc(acad)
    name = doc.Name
    doc.Close(save)
    return {"closed": name}


@mcp.tool()
def switch_drawing(name: str) -> dict:
    """Cambia el dibujo activo. name = nombre del archivo (con o sin extensión)."""
    acad = _get_acad()
    for doc in acad.Documents:
        if name.lower() in doc.Name.lower():
            doc.Activate()
            return {"active": doc.Name}
    raise RuntimeError(f"No se encontró un dibujo abierto con nombre '{name}'.")


# ============================================================================
# Identificación y análisis del dibujo
# ============================================================================

@mcp.tool()
def get_drawing_info() -> dict:
    """Metadatos completos del dibujo activo: nombre, ruta, unidades, límites, bbox."""
    acad = _get_acad()
    doc = _active_doc(acad)
    db = doc.Database
    limits_min = db.Limmin
    limits_max = db.Limmax

    # Extents reales
    try:
        ext_min = db.Extmin
        ext_max = db.Extmax
        extents = {
            "min": {"x": ext_min[0], "y": ext_min[1]},
            "max": {"x": ext_max[0], "y": ext_max[1]},
            "width": ext_max[0] - ext_min[0],
            "height": ext_max[1] - ext_min[1],
        }
    except Exception:
        extents = None

    return {
        "name": doc.Name,
        "path": doc.FullName,
        "saved": doc.Saved,
        "units_code": db.Lunits,
        "units_name": UNITS_NAMES.get(db.Lunits, "Desconocido"),
        "limits": {
            "min": {"x": limits_min[0], "y": limits_min[1]},
            "max": {"x": limits_max[0], "y": limits_max[1]},
        },
        "extents": extents,
        "entity_count": doc.ModelSpace.Count,
        "layer_count": doc.Layers.Count,
        "block_count": doc.Blocks.Count,
        "active_layer": doc.ActiveLayer.Name,
    }


@mcp.tool()
def analyze_drawing(include_handles: bool = False) -> dict:
    """
    Analiza el dibujo activo: conteo por tipo, conteo por capa, bbox global.
    Ideal para que Claude tenga una visión general antes de actuar.

    Args:
        include_handles: True para incluir handles agrupados (puede ser largo).
    """
    acad = _get_acad()
    doc = _active_doc(acad)
    ms = doc.ModelSpace

    by_type: dict[str, int] = {}
    by_layer: dict[str, dict[str, Any]] = {}
    handles_by_layer: dict[str, list[str]] = {}

    for entity in ms:
        try:
            t = entity.ObjectName
            layer = entity.Layer
            by_type[t] = by_type.get(t, 0) + 1
            if layer not in by_layer:
                by_layer[layer] = {"count": 0, "types": {}}
            by_layer[layer]["count"] += 1
            by_layer[layer]["types"][t] = by_layer[layer]["types"].get(t, 0) + 1
            if include_handles:
                handles_by_layer.setdefault(layer, []).append(entity.Handle)
        except Exception:
            continue

    # Convertir a tipos amigables
    by_type_friendly = {
        _friendly_type(k): v for k, v in by_type.items()
    }

    result = {
        "document": doc.Name,
        "total_entities": ms.Count,
        "by_type": by_type_friendly,
        "by_type_raw": by_type,
        "by_layer": by_layer,
        "layers_total": doc.Layers.Count,
    }
    if include_handles:
        result["handles_by_layer"] = handles_by_layer
    return result


@mcp.tool()
def get_bounding_box(handle: str = "") -> dict:
    """
    Bounding box de una entidad (por handle) o de todo el dibujo si handle="".
    """
    acad = _get_acad()
    doc = _active_doc(acad)
    if handle:
        e = doc.HandleToObject(handle)
        bb_min, bb_max = e.GetBoundingBox()
        return {
            "handle": handle,
            "min": {"x": bb_min[0], "y": bb_min[1], "z": bb_min[2]},
            "max": {"x": bb_max[0], "y": bb_max[1], "z": bb_max[2]},
            "width": bb_max[0] - bb_min[0],
            "height": bb_max[1] - bb_min[1],
        }
    db = doc.Database
    ext_min, ext_max = db.Extmin, db.Extmax
    return {
        "scope": "drawing",
        "min": {"x": ext_min[0], "y": ext_min[1]},
        "max": {"x": ext_max[0], "y": ext_max[1]},
        "width": ext_max[0] - ext_min[0],
        "height": ext_max[1] - ext_min[1],
    }


@mcp.tool()
def list_blocks() -> list[dict]:
    """Lista las definiciones de bloque del dibujo (excluye espacios modelo/papel)."""
    acad = _get_acad()
    doc = _active_doc(acad)
    result = []
    for block in doc.Blocks:
        if block.IsLayout:
            continue
        try:
            result.append({
                "name": block.Name,
                "entity_count": block.Count,
                "is_xref": block.IsXRef,
            })
        except Exception:
            pass
    return result


# ============================================================================
# Capas
# ============================================================================

@mcp.tool()
def list_layers() -> list[dict]:
    """Lista todas las capas del dibujo con estado y color."""
    acad = _get_acad()
    doc = _active_doc(acad)
    layers = []
    for layer in doc.Layers:
        layers.append({
            "name": layer.Name,
            "on": layer.LayerOn,
            "frozen": layer.Freeze,
            "locked": layer.Lock,
            "color": layer.color,
            "color_name": ACI_COLORS.get(layer.color, str(layer.color)),
            "linetype": layer.Linetype,
            "lineweight": layer.Lineweight,
        })
    return layers


@mcp.tool()
def create_layer(name: str, color: int = 7) -> dict:
    """Crea una capa nueva con un color ACI."""
    acad = _get_acad()
    doc = _active_doc(acad)
    layer = doc.Layers.Add(name)
    layer.color = color
    return {"created": name, "color": color}


@mcp.tool()
def set_active_layer(name: str) -> dict:
    """Cambia la capa activa."""
    acad = _get_acad()
    doc = _active_doc(acad)
    doc.ActiveLayer = doc.Layers.Item(name)
    return {"active_layer": name}


@mcp.tool()
def set_layer_state(name: str, on: bool | None = None,
                    frozen: bool | None = None,
                    locked: bool | None = None) -> dict:
    """Cambia el estado on/frozen/locked de una capa."""
    acad = _get_acad()
    doc = _active_doc(acad)
    layer = doc.Layers.Item(name)
    if on is not None:
        layer.LayerOn = on
    if frozen is not None:
        layer.Freeze = frozen
    if locked is not None:
        layer.Lock = locked
    return {
        "layer": name,
        "on": layer.LayerOn,
        "frozen": layer.Freeze,
        "locked": layer.Lock,
    }


# ============================================================================
# Crear geometría
# ============================================================================

@mcp.tool()
def create_line(x1: float, y1: float, x2: float, y2: float,
                z1: float = 0.0, z2: float = 0.0, layer: str = "") -> dict:
    """Crea una línea entre dos puntos."""
    acad = _get_acad()
    ms = _model_space(acad)
    line = ms.AddLine(_point(x1, y1, z1), _point(x2, y2, z2))
    if layer:
        line.Layer = layer
    return _entity_detail(line)


@mcp.tool()
def create_circle(cx: float, cy: float, radius: float,
                  cz: float = 0.0, layer: str = "") -> dict:
    """Crea un círculo dado centro y radio."""
    acad = _get_acad()
    ms = _model_space(acad)
    c = ms.AddCircle(_point(cx, cy, cz), radius)
    if layer:
        c.Layer = layer
    return _entity_detail(c)


@mcp.tool()
def create_arc(cx: float, cy: float, radius: float,
               start_angle_deg: float, end_angle_deg: float,
               cz: float = 0.0, layer: str = "") -> dict:
    """Crea un arco por centro, radio y ángulos en grados."""
    acad = _get_acad()
    ms = _model_space(acad)
    arc = ms.AddArc(
        _point(cx, cy, cz),
        radius,
        math.radians(start_angle_deg),
        math.radians(end_angle_deg),
    )
    if layer:
        arc.Layer = layer
    return _entity_detail(arc)


@mcp.tool()
def create_rectangle(x1: float, y1: float, x2: float, y2: float,
                     layer: str = "") -> dict:
    """Crea un rectángulo como polilínea cerrada entre dos esquinas opuestas."""
    acad = _get_acad()
    ms = _model_space(acad)
    pts = _double_array([x1, y1, x2, y1, x2, y2, x1, y2])
    pline = ms.AddLightWeightPolyline(pts)
    pline.Closed = True
    if layer:
        pline.Layer = layer
    return _entity_detail(pline)


@mcp.tool()
def create_polyline(points: list[list[float]], closed: bool = False,
                    layer: str = "") -> dict:
    """Polilínea 2D dada una lista de puntos [[x,y], ...]."""
    acad = _get_acad()
    ms = _model_space(acad)
    flat = []
    for p in points:
        flat.extend([float(p[0]), float(p[1])])
    pline = ms.AddLightWeightPolyline(_double_array(flat))
    pline.Closed = closed
    if layer:
        pline.Layer = layer
    return _entity_detail(pline)


@mcp.tool()
def create_ellipse(cx: float, cy: float, major_x: float, major_y: float,
                   ratio: float, layer: str = "") -> dict:
    """
    Crea una elipse.

    Args:
        cx, cy: Centro.
        major_x, major_y: Vector del semieje mayor (desde el centro).
        ratio: Relación semieje menor / mayor (entre 0 y 1).
    """
    acad = _get_acad()
    ms = _model_space(acad)
    ell = ms.AddEllipse(_point(cx, cy, 0), _point(major_x, major_y, 0), ratio)
    if layer:
        ell.Layer = layer
    return _entity_detail(ell)


@mcp.tool()
def insert_block(block_name: str, x: float, y: float,
                 scale_x: float = 1.0, scale_y: float = 1.0, scale_z: float = 1.0,
                 rotation_deg: float = 0.0, layer: str = "") -> dict:
    """Inserta un bloque existente."""
    acad = _get_acad()
    ms = _model_space(acad)
    ref = ms.InsertBlock(
        _point(x, y, 0), block_name,
        scale_x, scale_y, scale_z, math.radians(rotation_deg),
    )
    if layer:
        ref.Layer = layer
    return _entity_detail(ref)


# ============================================================================
# Textos — creación
# ============================================================================

@mcp.tool()
def create_text(text: str, x: float, y: float, height: float = 2.5,
                rotation_deg: float = 0.0, z: float = 0.0,
                style: str = "", layer: str = "") -> dict:
    """Crea un texto de una línea."""
    acad = _get_acad()
    ms = _model_space(acad)
    txt = ms.AddText(text, _point(x, y, z), height)
    if rotation_deg:
        txt.Rotation = math.radians(rotation_deg)
    if style:
        txt.StyleName = style
    if layer:
        txt.Layer = layer
    return _entity_detail(txt)


@mcp.tool()
def create_mtext(text: str, x: float, y: float, width: float = 100.0,
                 height: float = 2.5, rotation_deg: float = 0.0,
                 style: str = "", layer: str = "") -> dict:
    """Crea un MText (multilínea). Usá \\P para saltos de línea."""
    acad = _get_acad()
    ms = _model_space(acad)
    mt = ms.AddMText(_point(x, y, 0), width, text)
    mt.Height = height
    if rotation_deg:
        mt.Rotation = math.radians(rotation_deg)
    if style:
        mt.StyleName = style
    if layer:
        mt.Layer = layer
    return _entity_detail(mt)


# ============================================================================
# Textos — edición avanzada
# ============================================================================

@mcp.tool()
def list_all_texts(layer_filter: str = "", include_blocks: bool = False) -> list[dict]:
    """
    Lista todos los textos (Text y MText) del modelo con contenido y posición.

    Args:
        layer_filter: Si se indica, filtra por capa.
        include_blocks: Si True, también busca atributos dentro de bloques.
    """
    acad = _get_acad()
    ms = _model_space(acad)
    result = []
    for entity in ms:
        try:
            if entity.ObjectName in ("AcDbText", "AcDbMText"):
                if layer_filter and entity.Layer != layer_filter:
                    continue
                ip = entity.InsertionPoint
                result.append({
                    "handle": entity.Handle,
                    "type": _friendly_type(entity.ObjectName),
                    "layer": entity.Layer,
                    "text": entity.TextString,
                    "x": ip[0], "y": ip[1],
                    "height": entity.Height,
                    "rotation_deg": math.degrees(entity.Rotation),
                })
            elif include_blocks and entity.ObjectName == "AcDbBlockReference":
                if entity.HasAttributes:
                    for att in entity.GetAttributes():
                        if layer_filter and entity.Layer != layer_filter:
                            continue
                        result.append({
                            "handle": att.Handle,
                            "type": "Atributo de bloque",
                            "block_name": entity.Name,
                            "block_handle": entity.Handle,
                            "layer": entity.Layer,
                            "tag": att.TagString,
                            "text": att.TextString,
                        })
        except Exception:
            continue
    return result


@mcp.tool()
def find_texts_containing(query: str, case_sensitive: bool = False,
                          include_blocks: bool = True) -> list[dict]:
    """
    Busca textos que contengan una subcadena.

    Args:
        query: Texto a buscar.
        case_sensitive: Distinguir mayúsculas/minúsculas.
        include_blocks: Buscar también en atributos de bloques.
    """
    needle = query if case_sensitive else query.lower()
    matches = []
    for t in list_all_texts(include_blocks=include_blocks):
        content = t.get("text", "") or ""
        haystack = content if case_sensitive else content.lower()
        if needle in haystack:
            matches.append(t)
    return matches


@mcp.tool()
def edit_text_by_handle(handle: str, new_text: str) -> dict:
    """
    Reemplaza el contenido de un texto (Text, MText o atributo) identificado por su handle.
    """
    acad = _get_acad()
    doc = _active_doc(acad)
    entity = doc.HandleToObject(handle)
    name = entity.ObjectName
    if name not in ("AcDbText", "AcDbMText", "AcDbAttribute", "AcDbAttributeDefinition"):
        raise RuntimeError(f"La entidad {handle} no es un texto editable (es {name}).")
    old = entity.TextString
    entity.TextString = new_text
    return {
        "handle": handle,
        "type": _friendly_type(name),
        "old_text": old,
        "new_text": new_text,
    }


@mcp.tool()
def find_and_replace_text(search: str, replace: str, case_sensitive: bool = False,
                          layer_filter: str = "", include_blocks: bool = True,
                          dry_run: bool = False) -> dict:
    """
    Reemplaza texto en todo el modelo.

    Args:
        search: Texto a buscar.
        replace: Texto de reemplazo.
        case_sensitive: Distinguir mayúsculas.
        layer_filter: Limita el reemplazo a una capa.
        include_blocks: Incluir atributos de bloques.
        dry_run: Si True, solo lista los matches sin modificar.
    """
    acad = _get_acad()
    doc = _active_doc(acad)
    ms = _model_space(acad)

    changed = []
    needle_cmp = search if case_sensitive else search.lower()

    def _do_replace(s: str) -> str:
        if case_sensitive:
            return s.replace(search, replace)
        # Case-insensitive replace preservando capitalización del reemplazo
        out, i = [], 0
        sl = s.lower()
        nl = len(search)
        while i < len(s):
            if sl[i:i+nl] == needle_cmp:
                out.append(replace)
                i += nl
            else:
                out.append(s[i])
                i += 1
        return "".join(out)

    for entity in ms:
        try:
            if layer_filter and entity.Layer != layer_filter:
                continue
            if entity.ObjectName in ("AcDbText", "AcDbMText"):
                old = entity.TextString
                cmp = old if case_sensitive else old.lower()
                if needle_cmp in cmp:
                    new = _do_replace(old)
                    if not dry_run:
                        entity.TextString = new
                    changed.append({
                        "handle": entity.Handle,
                        "type": _friendly_type(entity.ObjectName),
                        "layer": entity.Layer,
                        "old": old,
                        "new": new,
                    })
            elif include_blocks and entity.ObjectName == "AcDbBlockReference":
                if entity.HasAttributes:
                    for att in entity.GetAttributes():
                        old = att.TextString
                        cmp = old if case_sensitive else old.lower()
                        if needle_cmp in cmp:
                            new = _do_replace(old)
                            if not dry_run:
                                att.TextString = new
                            changed.append({
                                "handle": att.Handle,
                                "type": "Atributo",
                                "block": entity.Name,
                                "tag": att.TagString,
                                "old": old,
                                "new": new,
                            })
        except Exception:
            continue

    return {
        "dry_run": dry_run,
        "total_changed": len(changed),
        "changes": changed,
    }


@mcp.tool()
def set_text_properties(handle: str, height: float | None = None,
                        rotation_deg: float | None = None,
                        style: str | None = None,
                        width: float | None = None,
                        alignment_point: list[float] | None = None) -> dict:
    """
    Modifica propiedades de un texto sin tocar su contenido.

    Args:
        handle: Handle del texto.
        height: Nueva altura (None = sin cambio).
        rotation_deg: Nueva rotación en grados (None = sin cambio).
        style: Nuevo nombre de estilo de texto (None = sin cambio).
        width: Nuevo ancho del cuadro (solo MText).
        alignment_point: [x, y] nueva posición de inserción.
    """
    acad = _get_acad()
    doc = _active_doc(acad)
    entity = doc.HandleToObject(handle)
    if entity.ObjectName not in ("AcDbText", "AcDbMText"):
        raise RuntimeError(f"La entidad {handle} no es texto.")
    if height is not None:
        entity.Height = height
    if rotation_deg is not None:
        entity.Rotation = math.radians(rotation_deg)
    if style is not None:
        entity.StyleName = style
    if width is not None and entity.ObjectName == "AcDbMText":
        entity.Width = width
    if alignment_point is not None:
        entity.InsertionPoint = _point(alignment_point[0], alignment_point[1])
    return _entity_detail(entity)


@mcp.tool()
def list_text_styles() -> list[dict]:
    """Lista los estilos de texto definidos en el dibujo."""
    acad = _get_acad()
    doc = _active_doc(acad)
    styles = []
    for s in doc.TextStyles:
        try:
            styles.append({
                "name": s.Name,
                "font_file": s.fontFile,
                "height": s.Height,
            })
        except Exception:
            styles.append({"name": s.Name})
    return styles


# ============================================================================
# Modificación de entidades (genérica)
# ============================================================================

@mcp.tool()
def list_entities(layer_filter: str = "", type_filter: str = "",
                  limit: int = 200) -> list[dict]:
    """
    Lista entidades del modelo con detalles según su tipo.

    Args:
        layer_filter: Filtra por capa (vacío = todas).
        type_filter: Filtra por ObjectName (ej: 'AcDbLine') o nombre amigable
                     ('Línea', 'Círculo'). Vacío = todas.
        limit: Máximo de entidades a devolver.
    """
    acad = _get_acad()
    ms = _model_space(acad)
    result = []
    tf = type_filter.lower()
    for entity in ms:
        if len(result) >= limit:
            break
        try:
            if layer_filter and entity.Layer != layer_filter:
                continue
            if type_filter:
                if (entity.ObjectName.lower() != tf
                        and _friendly_type(entity.ObjectName).lower() != tf):
                    continue
            result.append(_entity_detail(entity))
        except Exception:
            continue
    return result


@mcp.tool()
def get_entity_by_handle(handle: str) -> dict:
    """Devuelve los detalles completos de una entidad por handle."""
    acad = _get_acad()
    doc = _active_doc(acad)
    return _entity_detail(doc.HandleToObject(handle))


@mcp.tool()
def move_entity(handle: str, dx: float, dy: float, dz: float = 0.0) -> dict:
    """Mueve una entidad por un desplazamiento."""
    acad = _get_acad()
    doc = _active_doc(acad)
    entity = doc.HandleToObject(handle)
    entity.Move(_point(0, 0, 0), _point(dx, dy, dz))
    return {"handle": handle, "moved_by": {"dx": dx, "dy": dy, "dz": dz}}


@mcp.tool()
def copy_entity(handle: str, dx: float, dy: float, dz: float = 0.0) -> dict:
    """Copia una entidad desplazándola."""
    acad = _get_acad()
    doc = _active_doc(acad)
    entity = doc.HandleToObject(handle)
    new_e = entity.Copy()
    new_e.Move(_point(0, 0, 0), _point(dx, dy, dz))
    return _entity_detail(new_e)


@mcp.tool()
def scale_entity(handle: str, base_x: float, base_y: float, factor: float) -> dict:
    """Escala una entidad respecto a un punto base."""
    acad = _get_acad()
    doc = _active_doc(acad)
    entity = doc.HandleToObject(handle)
    entity.ScaleEntity(_point(base_x, base_y, 0), factor)
    return {"handle": handle, "scale_factor": factor}


@mcp.tool()
def rotate_entity(handle: str, base_x: float, base_y: float, angle_deg: float) -> dict:
    """Rota una entidad respecto a un punto base."""
    acad = _get_acad()
    doc = _active_doc(acad)
    entity = doc.HandleToObject(handle)
    entity.Rotate(_point(base_x, base_y, 0), math.radians(angle_deg))
    return {"handle": handle, "rotated_deg": angle_deg}


@mcp.tool()
def mirror_entity(handle: str, x1: float, y1: float, x2: float, y2: float,
                  keep_original: bool = True) -> dict:
    """Espeja una entidad respecto al eje definido por dos puntos."""
    acad = _get_acad()
    doc = _active_doc(acad)
    entity = doc.HandleToObject(handle)
    mirrored = entity.Mirror(_point(x1, y1, 0), _point(x2, y2, 0))
    if not keep_original:
        entity.Delete()
    return _entity_detail(mirrored)


@mcp.tool()
def offset_entity(handle: str, distance: float) -> list[dict]:
    """Desplaza paralelamente una entidad (offset). Devuelve los nuevos objetos."""
    acad = _get_acad()
    doc = _active_doc(acad)
    entity = doc.HandleToObject(handle)
    result = entity.Offset(distance)
    out = []
    try:
        for r in result:
            out.append(_entity_detail(r))
    except Exception:
        out.append(_entity_detail(result))
    return out


@mcp.tool()
def change_entity_layer(handle: str, layer: str) -> dict:
    """Cambia la capa de una entidad."""
    acad = _get_acad()
    doc = _active_doc(acad)
    entity = doc.HandleToObject(handle)
    entity.Layer = layer
    return {"handle": handle, "new_layer": layer}


@mcp.tool()
def change_entity_color(handle: str, color: int) -> dict:
    """Cambia el color ACI de una entidad."""
    acad = _get_acad()
    doc = _active_doc(acad)
    entity = doc.HandleToObject(handle)
    entity.color = color
    return {"handle": handle, "new_color": color}


@mcp.tool()
def delete_entity(handle: str) -> dict:
    """Elimina una entidad."""
    acad = _get_acad()
    doc = _active_doc(acad)
    entity = doc.HandleToObject(handle)
    entity.Delete()
    return {"deleted": handle}


@mcp.tool()
def delete_entities(handles: list[str]) -> dict:
    """Elimina varias entidades por sus handles."""
    acad = _get_acad()
    doc = _active_doc(acad)
    deleted, errors = [], []
    for h in handles:
        try:
            doc.HandleToObject(h).Delete()
            deleted.append(h)
        except Exception as exc:
            errors.append({"handle": h, "error": str(exc)})
    return {"deleted": deleted, "errors": errors}


# ============================================================================
# Cómputos de superficie y perímetro
# ============================================================================

@mcp.tool()
def calculate_area_of_entity(handle: str) -> dict:
    """
    Calcula el área de una entidad cerrada (círculo, elipse, polilínea cerrada,
    región, hatch). Si la entidad no es cerrada o no tiene área, devuelve None.
    """
    acad = _get_acad()
    doc = _active_doc(acad)
    entity = doc.HandleToObject(handle)
    area = _safe_area(entity)
    perimeter = None
    try:
        if entity.ObjectName == "AcDbRegion":
            perimeter = entity.Perimeter
        elif entity.ObjectName == "AcDbCircle":
            perimeter = entity.Circumference
        elif entity.ObjectName in ("AcDbPolyline", "AcDb2dPolyline"):
            perimeter = entity.Length
    except Exception:
        pass
    return {
        "handle": handle,
        "type": _friendly_type(entity.ObjectName),
        "layer": entity.Layer,
        "area": area,
        "perimeter": perimeter,
    }


@mcp.tool()
def sum_areas(handles: list[str]) -> dict:
    """
    Suma el área de varias entidades por handle.
    Útil cuando el usuario seleccionó manualmente o ya filtraste handles.
    """
    acad = _get_acad()
    doc = _active_doc(acad)
    items, total, skipped = [], 0.0, []
    for h in handles:
        try:
            e = doc.HandleToObject(h)
            a = _safe_area(e)
            if a is None:
                skipped.append({"handle": h, "type": _friendly_type(e.ObjectName),
                                "reason": "sin área"})
                continue
            items.append({
                "handle": h,
                "type": _friendly_type(e.ObjectName),
                "layer": e.Layer,
                "area": a,
            })
            total += a
        except Exception as exc:
            skipped.append({"handle": h, "error": str(exc)})
    return {"items": items, "skipped": skipped, "total_area": total,
            "count": len(items)}


@mcp.tool()
def calculate_area_by_layer(layer: str,
                            only_closed: bool = True) -> dict:
    """
    Suma el área de todas las entidades de una capa que tengan área.

    Args:
        layer: Nombre de la capa.
        only_closed: Si True, salta entidades sin área (no cerradas).
    """
    acad = _get_acad()
    ms = _model_space(acad)
    items, total = [], 0.0
    skipped = 0
    for entity in ms:
        try:
            if entity.Layer != layer:
                continue
            a = _safe_area(entity)
            if a is None:
                if only_closed:
                    skipped += 1
                    continue
                a = 0.0
            items.append({
                "handle": entity.Handle,
                "type": _friendly_type(entity.ObjectName),
                "area": a,
            })
            total += a
        except Exception:
            skipped += 1
    return {
        "layer": layer,
        "count_with_area": len(items),
        "skipped": skipped,
        "items": items,
        "total_area": total,
    }


@mcp.tool()
def calculate_areas_all_layers(only_closed: bool = True) -> dict:
    """
    Cómputo de superficies agrupado por capa para todo el dibujo.
    Devuelve un resumen útil para informes.
    """
    acad = _get_acad()
    ms = _model_space(acad)
    by_layer: dict[str, dict[str, Any]] = {}

    for entity in ms:
        try:
            layer = entity.Layer
            a = _safe_area(entity)
            if a is None and only_closed:
                continue
            slot = by_layer.setdefault(layer, {
                "count": 0, "total_area": 0.0, "types": {}
            })
            slot["count"] += 1
            if a is not None:
                slot["total_area"] += a
            t = _friendly_type(entity.ObjectName)
            slot["types"][t] = slot["types"].get(t, 0) + 1
        except Exception:
            continue

    grand_total = sum(v["total_area"] for v in by_layer.values())
    return {
        "by_layer": by_layer,
        "grand_total_area": grand_total,
        "units_note": "Las áreas están en (unidad del dibujo)^2",
    }


@mcp.tool()
def calculate_boundary_area_at_point(x: float, y: float,
                                     keep_boundary: bool = False) -> dict:
    """
    Identifica el contorno cerrado que rodea a un punto y calcula su área.
    Usa el comando -BOUNDARY de AutoCAD para construir una polilínea cerrada,
    lee su área y opcionalmente la borra.

    Args:
        x, y: Punto interior del contorno cerrado.
        keep_boundary: True para conservar la polilínea generada; False la borra.
    """
    acad = _get_acad()
    doc = _active_doc(acad)
    ms = doc.ModelSpace

    # Snapshot de handles antes
    before = set()
    for e in ms:
        try:
            before.add(e.Handle)
        except Exception:
            pass

    # -BOUNDARY: A=advanced settings, sin diálogo, con polilínea
    # Comando: _-BOUNDARY  <enter>  (point) <enter>
    cmd = f"_-BOUNDARY \n{x},{y}\n\n"
    doc.SendCommand(cmd)

    new_handles = []
    for e in ms:
        try:
            if e.Handle not in before:
                new_handles.append(e.Handle)
        except Exception:
            pass

    if not new_handles:
        raise RuntimeError(
            f"No se pudo construir un contorno cerrado en ({x},{y}). "
            "Verificá que el punto esté dentro de un área cerrada visible."
        )

    results = []
    for h in new_handles:
        e = doc.HandleToObject(h)
        a = _safe_area(e)
        info = {
            "handle": h,
            "type": _friendly_type(e.ObjectName),
            "area": a,
        }
        try:
            info["perimeter"] = e.Length
        except Exception:
            pass
        if not keep_boundary:
            e.Delete()
            info["deleted"] = True
        results.append(info)

    total = sum((r.get("area") or 0.0) for r in results)
    return {"point": {"x": x, "y": y},
            "boundaries_created": len(results),
            "results": results,
            "total_area": total}


@mcp.tool()
def create_region_from_polyline(handle: str, delete_source: bool = False) -> dict:
    """
    Convierte una polilínea cerrada en Región (AcDbRegion), útil para
    operaciones booleanas y cómputos exactos.
    """
    acad = _get_acad()
    doc = _active_doc(acad)
    ms = doc.ModelSpace
    entity = doc.HandleToObject(handle)
    # AddRegion espera un array de curvas
    curves = win32.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [entity])
    regions = ms.AddRegion(curves)
    out = []
    for r in regions:
        out.append(_entity_detail(r))
    if delete_source:
        entity.Delete()
    return {"regions": out, "count": len(out)}


@mcp.tool()
def region_boolean(handle_a: str, handle_b: str,
                   operation: str = "union") -> dict:
    """
    Operación booleana entre dos regiones. La región A absorbe a B.

    Args:
        handle_a, handle_b: Handles de dos regiones.
        operation: 'union', 'subtract' o 'intersect'.
    """
    acad = _get_acad()
    doc = _active_doc(acad)
    op_map = {"union": 0, "intersect": 1, "subtract": 2}
    if operation not in op_map:
        raise ValueError(f"operation debe ser uno de {list(op_map)}")
    a = doc.HandleToObject(handle_a)
    b = doc.HandleToObject(handle_b)
    a.Boolean(op_map[operation], b)
    return {"result_handle": a.Handle, "type": "Región", "area": _safe_area(a)}


# ============================================================================
# Cotas y dimensiones
# ============================================================================

@mcp.tool()
def create_dimension_aligned(x1: float, y1: float, x2: float, y2: float,
                             tx: float, ty: float, layer: str = "") -> dict:
    """
    Cota alineada entre dos puntos.

    Args:
        x1,y1: Primer punto de extensión.
        x2,y2: Segundo punto de extensión.
        tx,ty: Punto de ubicación del texto de cota.
    """
    acad = _get_acad()
    ms = _model_space(acad)
    dim = ms.AddDimAligned(_point(x1, y1, 0), _point(x2, y2, 0), _point(tx, ty, 0))
    if layer:
        dim.Layer = layer
    return _entity_detail(dim)


@mcp.tool()
def create_dimension_linear(x1: float, y1: float, x2: float, y2: float,
                            tx: float, ty: float, rotation_deg: float = 0.0,
                            layer: str = "") -> dict:
    """Cota lineal (horizontal/vertical/rotada)."""
    acad = _get_acad()
    ms = _model_space(acad)
    dim = ms.AddDimRotated(
        _point(x1, y1, 0), _point(x2, y2, 0), _point(tx, ty, 0),
        math.radians(rotation_deg),
    )
    if layer:
        dim.Layer = layer
    return _entity_detail(dim)


# ============================================================================
# Vista y comandos
# ============================================================================

@mcp.tool()
def zoom_extents() -> dict:
    """Zoom para mostrar todo el dibujo."""
    _get_acad().ZoomExtents()
    return {"zoomed": "extents"}


@mcp.tool()
def zoom_window(x1: float, y1: float, x2: float, y2: float) -> dict:
    """Zoom a una ventana rectangular."""
    _get_acad().ZoomWindow(_point(x1, y1, 0), _point(x2, y2, 0))
    return {"zoomed": "window"}


@mcp.tool()
def zoom_to_entity(handle: str, margin: float = 1.2) -> dict:
    """
    Zoom centrado en una entidad concreta, con un margen multiplicativo.
    """
    acad = _get_acad()
    doc = _active_doc(acad)
    entity = doc.HandleToObject(handle)
    bb_min, bb_max = entity.GetBoundingBox()
    cx = (bb_min[0] + bb_max[0]) / 2
    cy = (bb_min[1] + bb_max[1]) / 2
    half_w = (bb_max[0] - bb_min[0]) / 2 * margin
    half_h = (bb_max[1] - bb_min[1]) / 2 * margin
    half = max(half_w, half_h, 1.0)
    acad.ZoomWindow(_point(cx - half, cy - half, 0), _point(cx + half, cy + half, 0))
    return {"zoomed_to": handle}


@mcp.tool()
def run_autocad_command(command: str) -> dict:
    """
    Envía un comando directo a AutoCAD. Usar con cuidado: el comando debe
    incluir los \\n necesarios para no quedar esperando input.
    """
    acad = _get_acad()
    _active_doc(acad).SendCommand(command)
    return {"command_sent": command}


@mcp.tool()
def get_selection_set(name: str = "MCP_SEL") -> list[dict]:
    """
    Pide al usuario que seleccione entidades en AutoCAD y devuelve sus detalles.
    Bloquea hasta que el usuario termina la selección.
    """
    acad = _get_acad()
    doc = _active_doc(acad)
    try:
        doc.SelectionSets.Item(name).Delete()
    except Exception:
        pass
    ss = doc.SelectionSets.Add(name)
    ss.SelectOnScreen()
    return [_entity_detail(e) for e in ss]


# ============================================================================
# Layouts (presentaciones / paper space)
# ============================================================================

# Escalas estándar de viewport: relación = papel / modelo
STANDARD_VIEWPORT_SCALES = {
    "1:1": 1.0,
    "1:2": 1/2,
    "1:5": 1/5,
    "1:10": 1/10,
    "1:20": 1/20,
    "1:25": 1/25,
    "1:50": 1/50,
    "1:75": 1/75,
    "1:100": 1/100,
    "1:125": 1/125,
    "1:200": 1/200,
    "1:250": 1/250,
    "1:500": 1/500,
    "1:1000": 1/1000,
    "2:1": 2.0,
    "5:1": 5.0,
    "10:1": 10.0,
}


@mcp.tool()
def list_layouts() -> list[dict]:
    """
    Lista todos los layouts (presentaciones) del dibujo, incluyendo Model.
    """
    acad = _get_acad()
    doc = _active_doc(acad)
    layouts = []
    for layout in doc.Layouts:
        try:
            layouts.append({
                "name": layout.Name,
                "tab_order": layout.TabOrder,
                "is_active": (layout.Name == doc.ActiveLayout.Name),
                "block_name": layout.Block.Name,
                "entity_count": layout.Block.Count,
            })
        except Exception:
            pass
    return sorted(layouts, key=lambda l: l.get("tab_order", 0))


@mcp.tool()
def set_active_layout(name: str) -> dict:
    """
    Cambia el layout activo. Usar 'Model' para volver a espacio modelo.

    Args:
        name: Nombre del layout (ej: 'Layout1', 'Plano A1', 'Model').
    """
    acad = _get_acad()
    doc = _active_doc(acad)
    for layout in doc.Layouts:
        if layout.Name.lower() == name.lower():
            doc.ActiveLayout = layout
            return {"active_layout": layout.Name}
    raise RuntimeError(f"No existe el layout '{name}'.")


@mcp.tool()
def create_layout(name: str, activate: bool = True) -> dict:
    """
    Crea un layout nuevo (presentación).

    Args:
        name: Nombre del layout.
        activate: True para activarlo automáticamente.
    """
    acad = _get_acad()
    doc = _active_doc(acad)
    layout = doc.Layouts.Add(name)
    if activate:
        doc.ActiveLayout = layout
    return {
        "created": name,
        "block_name": layout.Block.Name,
        "active": activate,
    }


@mcp.tool()
def delete_layout(name: str) -> dict:
    """Borra un layout. No se puede borrar 'Model'."""
    acad = _get_acad()
    doc = _active_doc(acad)
    if name.lower() == "model":
        raise RuntimeError("No se puede borrar el espacio Model.")
    for layout in doc.Layouts:
        if layout.Name.lower() == name.lower():
            layout.Delete()
            return {"deleted": name}
    raise RuntimeError(f"No existe el layout '{name}'.")


@mcp.tool()
def setup_layout_page(name: str, paper_width: float, paper_height: float,
                      plot_unit: str = "mm") -> dict:
    """
    Configura el tamaño de hoja de un layout.

    Args:
        name: Nombre del layout.
        paper_width: Ancho del papel.
        paper_height: Alto del papel.
        plot_unit: 'mm' o 'inches'.
    """
    acad = _get_acad()
    doc = _active_doc(acad)
    layout = None
    for l in doc.Layouts:
        if l.Name.lower() == name.lower():
            layout = l
            break
    if layout is None:
        raise RuntimeError(f"No existe el layout '{name}'.")
    # 0=inches, 1=mm
    layout.PaperUnits = 1 if plot_unit.lower() == "mm" else 0
    try:
        layout.SetCustomScale(1, 1)
    except Exception:
        pass
    try:
        layout.CanonicalMediaName = "User Defined"
    except Exception:
        pass
    try:
        layout.SetCustomPaperSize(paper_width, paper_height)
    except Exception as exc:
        return {"layout": name, "warning": f"No se pudo fijar tamaño custom: {exc}"}
    return {
        "layout": name,
        "paper_width": paper_width,
        "paper_height": paper_height,
        "unit": plot_unit,
    }


# ============================================================================
# Viewports (en paper space)
# ============================================================================

@mcp.tool()
def create_viewport(layout: str, center_x: float, center_y: float,
                    width: float, height: float,
                    scale_ratio: str = "",
                    model_target_x: float | None = None,
                    model_target_y: float | None = None) -> dict:
    """
    Crea un viewport rectangular en un layout y opcionalmente le fija escala.

    Args:
        layout: Nombre del layout (debe existir y estar en paper space).
        center_x, center_y: Centro del viewport en coordenadas de papel.
        width, height: Ancho y alto del viewport en unidades de papel.
        scale_ratio: Escala estándar como '1:50', '1:100', '2:1', etc.
                     Vacío = no fijar escala.
        model_target_x, model_target_y: Punto del modelo a centrar dentro
                     del viewport. Vacío = no cambiar el target.
    """
    acad = _get_acad()
    doc = _active_doc(acad)

    # Activar el layout solicitado
    found = None
    for l in doc.Layouts:
        if l.Name.lower() == layout.lower():
            found = l
            break
    if found is None:
        raise RuntimeError(f"Layout '{layout}' no existe.")
    doc.ActiveLayout = found

    pspace = doc.PaperSpace
    vp = pspace.AddPViewport(_point(center_x, center_y, 0), width, height)

    # Encender el viewport (mostrar el modelo dentro)
    try:
        doc.MSpace = True
        vp.Display(True)
    except Exception:
        pass
    finally:
        try:
            doc.MSpace = False
        except Exception:
            pass

    info: dict[str, Any] = {
        "handle": vp.Handle,
        "layout": layout,
        "center": {"x": center_x, "y": center_y},
        "size": {"width": width, "height": height},
    }

    # Fijar escala custom
    if scale_ratio:
        ratio = STANDARD_VIEWPORT_SCALES.get(scale_ratio)
        if ratio is None:
            # Parseo manual "A:B"
            if ":" in scale_ratio:
                a, b = scale_ratio.split(":")
                ratio = float(a) / float(b)
            else:
                raise ValueError(f"Escala inválida: {scale_ratio}")
        try:
            vp.CustomScale = ratio
            info["scale"] = scale_ratio
            info["custom_scale"] = ratio
        except Exception as exc:
            info["scale_error"] = str(exc)

    # Centrar el viewport en un punto del modelo
    if model_target_x is not None and model_target_y is not None:
        try:
            vp.ViewCenter = _point(model_target_x, model_target_y, 0)
            info["target"] = {"x": model_target_x, "y": model_target_y}
        except Exception as exc:
            info["target_error"] = str(exc)

    return info


@mcp.tool()
def list_viewports(layout: str = "") -> list[dict]:
    """
    Lista los viewports de un layout (o del layout activo si no se indica).
    No incluye el viewport "general" (#1) que envuelve toda la presentación.
    """
    acad = _get_acad()
    doc = _active_doc(acad)
    if layout:
        for l in doc.Layouts:
            if l.Name.lower() == layout.lower():
                doc.ActiveLayout = l
                break
    result = []
    for entity in doc.PaperSpace:
        if entity.ObjectName == "AcDbViewport":
            try:
                # El viewport 1 (con number=1) es el contenedor del layout
                if entity.Number == 1:
                    continue
            except Exception:
                pass
            try:
                c = entity.Center
                result.append({
                    "handle": entity.Handle,
                    "center": {"x": c[0], "y": c[1]},
                    "width": entity.Width,
                    "height": entity.Height,
                    "custom_scale": entity.CustomScale,
                    "scale_label": _scale_label(entity.CustomScale),
                    "on": entity.ViewportOn,
                    "locked": entity.DisplayLocked,
                })
            except Exception:
                pass
    return result


def _scale_label(ratio: float) -> str:
    """Convierte un factor de escala a 'A:B' aproximado."""
    if ratio <= 0:
        return str(ratio)
    if ratio >= 1:
        return f"{round(ratio)}:1"
    inv = 1 / ratio
    # Redondear a la escala estándar más cercana si está dentro de 1%
    for label, std in STANDARD_VIEWPORT_SCALES.items():
        if std == 0:
            continue
        if abs(std - ratio) / std < 0.01:
            return label
    return f"1:{round(inv)}"


@mcp.tool()
def set_viewport_scale(handle: str, scale_ratio: str) -> dict:
    """
    Fija la escala de un viewport por su handle.

    Args:
        handle: Handle del viewport (lo da list_viewports).
        scale_ratio: '1:50', '1:100', '2:1', etc.
    """
    acad = _get_acad()
    doc = _active_doc(acad)
    vp = doc.HandleToObject(handle)
    if vp.ObjectName != "AcDbViewport":
        raise RuntimeError(f"{handle} no es un viewport (es {vp.ObjectName}).")

    ratio = STANDARD_VIEWPORT_SCALES.get(scale_ratio)
    if ratio is None:
        if ":" in scale_ratio:
            a, b = scale_ratio.split(":")
            ratio = float(a) / float(b)
        else:
            raise ValueError(f"Escala inválida: {scale_ratio}")
    vp.CustomScale = ratio
    return {
        "handle": handle,
        "scale": scale_ratio,
        "custom_scale": ratio,
    }


@mcp.tool()
def lock_viewport(handle: str, locked: bool = True) -> dict:
    """Bloquea o desbloquea la visualización de un viewport para que no se altere su escala."""
    acad = _get_acad()
    doc = _active_doc(acad)
    vp = doc.HandleToObject(handle)
    vp.DisplayLocked = locked
    return {"handle": handle, "locked": locked}


@mcp.tool()
def fit_viewport_to_extents(handle: str, margin: float = 1.05) -> dict:
    """
    Ajusta el viewport para mostrar los extents del modelo.
    No fija una escala estándar — usalo cuando querés que se vea "todo".

    Args:
        handle: Handle del viewport.
        margin: Factor multiplicativo del tamaño visible (1.05 = 5% extra).
    """
    acad = _get_acad()
    doc = _active_doc(acad)
    vp = doc.HandleToObject(handle)
    db = doc.Database
    ext_min, ext_max = db.Extmin, db.Extmax
    cx = (ext_min[0] + ext_max[0]) / 2
    cy = (ext_min[1] + ext_max[1]) / 2
    model_w = (ext_max[0] - ext_min[0]) * margin
    model_h = (ext_max[1] - ext_min[1]) * margin
    # Escala que hace caber el modelo dentro del viewport
    paper_w = vp.Width
    paper_h = vp.Height
    scale = min(paper_w / max(model_w, 1e-9), paper_h / max(model_h, 1e-9))
    vp.CustomScale = scale
    vp.ViewCenter = _point(cx, cy, 0)
    return {
        "handle": handle,
        "custom_scale": scale,
        "scale_label": _scale_label(scale),
        "centered_on": {"x": cx, "y": cy},
    }


# ============================================================================
# Operaciones batch (lotes de geometría — mucho más rápidas)
# ============================================================================

@mcp.tool()
def create_lines_batch(lines: list[list[float]], layer: str = "") -> dict:
    """
    Crea varias líneas en un solo bloqueo de pantalla.

    Args:
        lines: Lista de líneas [[x1,y1,x2,y2], ...] o [[x1,y1,z1,x2,y2,z2], ...].
        layer: Capa para todas las líneas (vacío = capa activa).
    """
    acad = _get_acad()
    ms = _model_space(acad)
    handles = []
    with _fast_batch(acad):
        for ln in lines:
            if len(ln) == 4:
                x1, y1, x2, y2 = ln
                z1 = z2 = 0.0
            elif len(ln) == 6:
                x1, y1, z1, x2, y2, z2 = ln
            else:
                continue
            line = ms.AddLine(_point(x1, y1, z1), _point(x2, y2, z2))
            if layer:
                line.Layer = layer
            handles.append(line.Handle)
    return {"created": len(handles), "handles": handles}


@mcp.tool()
def create_polylines_batch(polylines: list[dict]) -> dict:
    """
    Crea varias polilíneas en lote.

    Args:
        polylines: Lista de objetos {"points": [[x,y],...], "closed": bool, "layer": str}
    """
    acad = _get_acad()
    ms = _model_space(acad)
    handles = []
    with _fast_batch(acad):
        for pl in polylines:
            pts = pl.get("points") or []
            flat = []
            for p in pts:
                flat.extend([float(p[0]), float(p[1])])
            if len(flat) < 4:
                continue
            pline = ms.AddLightWeightPolyline(_double_array(flat))
            if pl.get("closed"):
                pline.Closed = True
            if pl.get("layer"):
                pline.Layer = pl["layer"]
            handles.append(pline.Handle)
    return {"created": len(handles), "handles": handles}


@mcp.tool()
def create_texts_batch(texts: list[dict]) -> dict:
    """
    Crea varios textos en lote.

    Args:
        texts: Lista de {"text": str, "x": float, "y": float,
                         "height": float, "rotation_deg": float,
                         "style": str, "layer": str}
    """
    acad = _get_acad()
    ms = _model_space(acad)
    handles = []
    with _fast_batch(acad):
        for t in texts:
            txt = ms.AddText(
                t["text"],
                _point(t["x"], t["y"], 0),
                float(t.get("height", 2.5)),
            )
            if t.get("rotation_deg"):
                txt.Rotation = math.radians(float(t["rotation_deg"]))
            if t.get("style"):
                txt.StyleName = t["style"]
            if t.get("layer"):
                txt.Layer = t["layer"]
            handles.append(txt.Handle)
    return {"created": len(handles), "handles": handles}


@mcp.tool()
def transform_batch(handles: list[str],
                    dx: float = 0.0, dy: float = 0.0,
                    scale_factor: float | None = None,
                    base_x: float = 0.0, base_y: float = 0.0,
                    rotation_deg: float | None = None) -> dict:
    """
    Aplica desplazamiento, escala y/o rotación a varias entidades en una sola pasada.

    Args:
        handles: Lista de handles.
        dx, dy: Desplazamiento (opcional).
        scale_factor: Factor de escala respecto a (base_x, base_y).
        base_x, base_y: Punto base para escala y rotación.
        rotation_deg: Rotación en grados respecto a (base_x, base_y).
    """
    acad = _get_acad()
    doc = _active_doc(acad)
    affected, errors = [], []
    base = _point(base_x, base_y, 0)
    with _fast_batch(acad):
        for h in handles:
            try:
                e = doc.HandleToObject(h)
                if dx or dy:
                    e.Move(_point(0, 0, 0), _point(dx, dy, 0))
                if scale_factor is not None:
                    e.ScaleEntity(base, float(scale_factor))
                if rotation_deg is not None:
                    e.Rotate(base, math.radians(float(rotation_deg)))
                affected.append(h)
            except Exception as exc:
                errors.append({"handle": h, "error": str(exc)})
    return {"affected": len(affected), "handles": affected, "errors": errors}


# ============================================================================
# Control de performance
# ============================================================================

@mcp.tool()
def set_performance_mode(enabled: bool = True) -> dict:
    """
    Activa/desactiva variables de AutoCAD que mejoran el rendimiento durante
    sesiones largas:
    - CMDECHO=0: no imprime cada comando en la línea de comandos
    - REGENMODE=0: evita regeneraciones automáticas frecuentes
    - HIGHLIGHT=0: no resalta entidades al seleccionarlas

    Llamá a esta tool con enabled=False al final para restaurar defaults.
    """
    acad = _get_acad()
    doc = _active_doc(acad)
    if enabled:
        doc.SetVariable("CMDECHO", 0)
        try:
            doc.SetVariable("REGENMODE", 0)
        except Exception:
            pass
        try:
            doc.SetVariable("HIGHLIGHT", 0)
        except Exception:
            pass
        return {"performance_mode": "ON"}
    else:
        doc.SetVariable("CMDECHO", 1)
        try:
            doc.SetVariable("REGENMODE", 1)
        except Exception:
            pass
        try:
            doc.SetVariable("HIGHLIGHT", 1)
        except Exception:
            pass
        try:
            doc.Regen(1)
        except Exception:
            pass
        return {"performance_mode": "OFF"}


@mcp.tool()
def regen_drawing() -> dict:
    """Fuerza una regeneración del dibujo (útil después de un batch grande)."""
    acad = _get_acad()
    doc = _active_doc(acad)
    doc.Regen(1)
    return {"regenerated": True}


# ============================================================================
# Entry point
# ============================================================================

if __name__ == "__main__":
    mcp.run(transport="stdio")
