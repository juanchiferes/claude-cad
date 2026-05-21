"""
AutoCAD MCP Server — controla AutoCAD desde Claude via COM/ActiveX.
Requiere Windows + AutoCAD abierto + pywin32 instalado.
"""

import sys
import json
import math
from typing import Any
from mcp.server.fastmcp import FastMCP

# win32com solo disponible en Windows
try:
    import win32com.client as win32
    import pythoncom
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

mcp = FastMCP("autocad-mcp")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _get_acad():
    """Obtiene la instancia activa de AutoCAD."""
    if not WIN32_AVAILABLE:
        raise RuntimeError(
            "pywin32 no instalado. Ejecuta: pip install pywin32"
        )
    try:
        acad = win32.GetActiveObject("AutoCAD.Application")
        return acad
    except Exception:
        raise RuntimeError(
            "AutoCAD no está abierto. Abrí AutoCAD primero."
        )


def _active_doc(acad):
    doc = acad.ActiveDocument
    if doc is None:
        raise RuntimeError("No hay ningún dibujo abierto en AutoCAD.")
    return doc


def _model_space(acad):
    return _active_doc(acad).ModelSpace


def _point(x: float, y: float, z: float = 0.0):
    """Crea un array de punto compatible con COM."""
    pt = win32.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))
    return pt


def _flatten(entities):
    """Serializa una lista de entidades a dicts."""
    result = []
    for e in entities:
        try:
            result.append({
                "handle": e.Handle,
                "type": e.ObjectName,
                "layer": e.Layer,
            })
        except Exception:
            pass
    return result


# ---------------------------------------------------------------------------
# Herramientas — Estado y documentos
# ---------------------------------------------------------------------------

@mcp.tool()
def get_autocad_status() -> dict:
    """Verifica si AutoCAD está corriendo y devuelve info del documento activo."""
    try:
        acad = _get_acad()
        doc = acad.ActiveDocument
        return {
            "running": True,
            "version": acad.Version,
            "document": doc.Name if doc else None,
            "saved": doc.Saved if doc else None,
            "path": doc.FullName if doc else None,
        }
    except Exception as exc:
        return {"running": False, "error": str(exc)}


@mcp.tool()
def open_drawing(path: str) -> dict:
    """
    Abre un archivo DWG en AutoCAD.

    Args:
        path: Ruta completa al archivo .dwg (ej: C:\\proyectos\\plano.dwg)
    """
    acad = _get_acad()
    doc = acad.Documents.Open(path)
    return {"opened": doc.Name, "path": doc.FullName}


@mcp.tool()
def save_drawing(path: str = "") -> dict:
    """
    Guarda el dibujo activo.

    Args:
        path: Si se indica, guarda como (Save As) en esa ruta. Vacío = guarda en el mismo archivo.
    """
    acad = _get_acad()
    doc = _active_doc(acad)
    if path:
        doc.SaveAs(path)
    else:
        doc.Save()
    return {"saved": True, "path": doc.FullName}


@mcp.tool()
def close_drawing(save: bool = True) -> dict:
    """
    Cierra el dibujo activo.

    Args:
        save: True para guardar antes de cerrar.
    """
    acad = _get_acad()
    doc = _active_doc(acad)
    name = doc.Name
    doc.Close(save)
    return {"closed": name}


@mcp.tool()
def get_drawing_info() -> dict:
    """Devuelve metadatos del dibujo activo: nombre, ruta, unidades, límites."""
    acad = _get_acad()
    doc = _active_doc(acad)
    db = doc.Database
    limits_min = db.Limmin
    limits_max = db.Limmax
    return {
        "name": doc.Name,
        "path": doc.FullName,
        "saved": doc.Saved,
        "units": db.Lunits,
        "limits_min": {"x": limits_min[0], "y": limits_min[1]},
        "limits_max": {"x": limits_max[0], "y": limits_max[1]},
        "entity_count": doc.ModelSpace.Count,
    }


# ---------------------------------------------------------------------------
# Herramientas — Capas
# ---------------------------------------------------------------------------

@mcp.tool()
def list_layers() -> list[dict]:
    """Lista todas las capas del dibujo activo."""
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
        })
    return layers


@mcp.tool()
def create_layer(name: str, color: int = 7) -> dict:
    """
    Crea una capa nueva.

    Args:
        name: Nombre de la capa.
        color: Número de color ACI (1=rojo, 2=amarillo, 3=verde, 7=blanco).
    """
    acad = _get_acad()
    doc = _active_doc(acad)
    layer = doc.Layers.Add(name)
    layer.color = color
    return {"created": name, "color": color}


@mcp.tool()
def set_active_layer(name: str) -> dict:
    """
    Cambia la capa activa.

    Args:
        name: Nombre de la capa existente.
    """
    acad = _get_acad()
    doc = _active_doc(acad)
    doc.ActiveLayer = doc.Layers.Item(name)
    return {"active_layer": name}


# ---------------------------------------------------------------------------
# Herramientas — Creación de geometría
# ---------------------------------------------------------------------------

@mcp.tool()
def create_line(
    x1: float, y1: float, x2: float, y2: float,
    z1: float = 0.0, z2: float = 0.0,
    layer: str = ""
) -> dict:
    """
    Crea una línea en el espacio modelo.

    Args:
        x1, y1, z1: Punto de inicio.
        x2, y2, z2: Punto de fin.
        layer: Capa donde se crea (vacío = capa activa).
    """
    acad = _get_acad()
    ms = _model_space(acad)
    line = ms.AddLine(_point(x1, y1, z1), _point(x2, y2, z2))
    if layer:
        line.Layer = layer
    return {"handle": line.Handle, "type": "Line", "layer": line.Layer}


@mcp.tool()
def create_circle(
    cx: float, cy: float, radius: float,
    cz: float = 0.0, layer: str = ""
) -> dict:
    """
    Crea un círculo en el espacio modelo.

    Args:
        cx, cy, cz: Centro del círculo.
        radius: Radio.
        layer: Capa (vacío = capa activa).
    """
    acad = _get_acad()
    ms = _model_space(acad)
    circle = ms.AddCircle(_point(cx, cy, cz), radius)
    if layer:
        circle.Layer = layer
    return {"handle": circle.Handle, "type": "Circle", "center": {"x": cx, "y": cy}, "radius": radius, "layer": circle.Layer}


@mcp.tool()
def create_arc(
    cx: float, cy: float, radius: float,
    start_angle_deg: float, end_angle_deg: float,
    cz: float = 0.0, layer: str = ""
) -> dict:
    """
    Crea un arco en el espacio modelo.

    Args:
        cx, cy, cz: Centro del arco.
        radius: Radio.
        start_angle_deg: Ángulo inicial en grados.
        end_angle_deg: Ángulo final en grados.
        layer: Capa (vacío = capa activa).
    """
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
    return {"handle": arc.Handle, "type": "Arc", "layer": arc.Layer}


@mcp.tool()
def create_rectangle(
    x1: float, y1: float, x2: float, y2: float,
    layer: str = ""
) -> dict:
    """
    Crea un rectángulo (polilínea cerrada) dados dos esquinas opuestas.

    Args:
        x1, y1: Esquina inferior izquierda.
        x2, y2: Esquina superior derecha.
        layer: Capa (vacío = capa activa).
    """
    acad = _get_acad()
    ms = _model_space(acad)
    pts = win32.VARIANT(
        pythoncom.VT_ARRAY | pythoncom.VT_R8,
        [x1, y1, 0, x2, y1, 0, x2, y2, 0, x1, y2, 0, x1, y1, 0],
    )
    pline = ms.Add3DPoly(pts)
    pline.Closed = True
    if layer:
        pline.Layer = layer
    return {"handle": pline.Handle, "type": "Rectangle/3DPoly", "layer": pline.Layer}


@mcp.tool()
def create_polyline(
    points: list[list[float]],
    closed: bool = False,
    layer: str = ""
) -> dict:
    """
    Crea una polilínea 2D con los puntos dados.

    Args:
        points: Lista de puntos [[x1,y1], [x2,y2], ...].
        closed: True para cerrar la polilínea.
        layer: Capa (vacío = capa activa).
    """
    acad = _get_acad()
    ms = _model_space(acad)
    flat = []
    for p in points:
        flat.extend([float(p[0]), float(p[1])])
    pts = win32.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, flat)
    pline = ms.AddLightWeightPolyline(pts)
    pline.Closed = closed
    if layer:
        pline.Layer = layer
    return {"handle": pline.Handle, "type": "LWPolyline", "closed": closed, "layer": pline.Layer}


@mcp.tool()
def create_text(
    text: str, x: float, y: float,
    height: float = 2.5, z: float = 0.0,
    layer: str = ""
) -> dict:
    """
    Crea un texto en el espacio modelo.

    Args:
        text: Contenido del texto.
        x, y, z: Punto de inserción.
        height: Altura del texto.
        layer: Capa (vacío = capa activa).
    """
    acad = _get_acad()
    ms = _model_space(acad)
    txt = ms.AddText(text, _point(x, y, z), height)
    if layer:
        txt.Layer = layer
    return {"handle": txt.Handle, "type": "Text", "content": text, "layer": txt.Layer}


@mcp.tool()
def create_mtext(
    text: str, x: float, y: float,
    width: float = 100.0, height: float = 2.5,
    layer: str = ""
) -> dict:
    """
    Crea un texto multilínea (MText).

    Args:
        text: Contenido (soporta \\P para salto de línea).
        x, y: Punto de inserción.
        width: Ancho del cuadro de texto.
        height: Altura del carácter.
        layer: Capa (vacío = capa activa).
    """
    acad = _get_acad()
    ms = _model_space(acad)
    mt = ms.AddMText(_point(x, y, 0), width, text)
    mt.Height = height
    if layer:
        mt.Layer = layer
    return {"handle": mt.Handle, "type": "MText", "layer": mt.Layer}


@mcp.tool()
def insert_block(
    block_name: str, x: float, y: float,
    scale_x: float = 1.0, scale_y: float = 1.0, scale_z: float = 1.0,
    rotation_deg: float = 0.0, layer: str = ""
) -> dict:
    """
    Inserta un bloque existente en el dibujo.

    Args:
        block_name: Nombre del bloque definido en el dibujo.
        x, y: Punto de inserción.
        scale_x, scale_y, scale_z: Escala en cada eje.
        rotation_deg: Rotación en grados.
        layer: Capa (vacío = capa activa).
    """
    acad = _get_acad()
    ms = _model_space(acad)
    ref = ms.InsertBlock(
        _point(x, y, 0),
        block_name,
        scale_x, scale_y, scale_z,
        math.radians(rotation_deg),
    )
    if layer:
        ref.Layer = layer
    return {"handle": ref.Handle, "type": "BlockReference", "block": block_name, "layer": ref.Layer}


# ---------------------------------------------------------------------------
# Herramientas — Consulta y modificación de entidades
# ---------------------------------------------------------------------------

@mcp.tool()
def list_entities(layer_filter: str = "") -> list[dict]:
    """
    Lista entidades del espacio modelo.

    Args:
        layer_filter: Si se indica, filtra por nombre de capa.
    """
    acad = _get_acad()
    ms = _model_space(acad)
    result = []
    for entity in ms:
        try:
            if layer_filter and entity.Layer != layer_filter:
                continue
            info: dict[str, Any] = {
                "handle": entity.Handle,
                "type": entity.ObjectName,
                "layer": entity.Layer,
            }
            name = entity.ObjectName
            if name == "AcDbLine":
                sp = entity.StartPoint
                ep = entity.EndPoint
                info["start"] = {"x": sp[0], "y": sp[1], "z": sp[2]}
                info["end"] = {"x": ep[0], "y": ep[1], "z": ep[2]}
                info["length"] = entity.Length
            elif name == "AcDbCircle":
                c = entity.Center
                info["center"] = {"x": c[0], "y": c[1]}
                info["radius"] = entity.Radius
            elif name == "AcDbArc":
                c = entity.Center
                info["center"] = {"x": c[0], "y": c[1]}
                info["radius"] = entity.Radius
                info["start_angle_deg"] = math.degrees(entity.StartAngle)
                info["end_angle_deg"] = math.degrees(entity.EndAngle)
            elif name in ("AcDbText", "AcDbMText"):
                info["text"] = entity.TextString
            result.append(info)
        except Exception:
            pass
    return result


@mcp.tool()
def get_entity_by_handle(handle: str) -> dict:
    """
    Obtiene los detalles de una entidad por su handle (ID interno).

    Args:
        handle: Handle de la entidad (se obtiene con list_entities).
    """
    acad = _get_acad()
    doc = _active_doc(acad)
    entity = doc.HandleToObject(handle)
    info: dict[str, Any] = {
        "handle": entity.Handle,
        "type": entity.ObjectName,
        "layer": entity.Layer,
        "visible": entity.Visible,
    }
    name = entity.ObjectName
    if name == "AcDbLine":
        sp = entity.StartPoint
        ep = entity.EndPoint
        info["start"] = {"x": sp[0], "y": sp[1], "z": sp[2]}
        info["end"] = {"x": ep[0], "y": ep[1], "z": ep[2]}
        info["length"] = entity.Length
    elif name == "AcDbCircle":
        c = entity.Center
        info["center"] = {"x": c[0], "y": c[1]}
        info["radius"] = entity.Radius
    elif name in ("AcDbText", "AcDbMText"):
        info["text"] = entity.TextString
    return info


@mcp.tool()
def move_entity(handle: str, dx: float, dy: float, dz: float = 0.0) -> dict:
    """
    Mueve una entidad el desplazamiento indicado.

    Args:
        handle: Handle de la entidad.
        dx, dy, dz: Desplazamiento.
    """
    acad = _get_acad()
    doc = _active_doc(acad)
    entity = doc.HandleToObject(handle)
    entity.Move(_point(0, 0, 0), _point(dx, dy, dz))
    return {"handle": handle, "moved_by": {"dx": dx, "dy": dy, "dz": dz}}


@mcp.tool()
def scale_entity(handle: str, base_x: float, base_y: float, factor: float) -> dict:
    """
    Escala una entidad respecto a un punto base.

    Args:
        handle: Handle de la entidad.
        base_x, base_y: Punto base de la escala.
        factor: Factor de escala.
    """
    acad = _get_acad()
    doc = _active_doc(acad)
    entity = doc.HandleToObject(handle)
    entity.ScaleEntity(_point(base_x, base_y, 0), factor)
    return {"handle": handle, "scale_factor": factor}


@mcp.tool()
def rotate_entity(handle: str, base_x: float, base_y: float, angle_deg: float) -> dict:
    """
    Rota una entidad respecto a un punto base.

    Args:
        handle: Handle de la entidad.
        base_x, base_y: Punto de rotación.
        angle_deg: Ángulo de rotación en grados (antihorario).
    """
    acad = _get_acad()
    doc = _active_doc(acad)
    entity = doc.HandleToObject(handle)
    entity.Rotate(_point(base_x, base_y, 0), math.radians(angle_deg))
    return {"handle": handle, "rotated_deg": angle_deg}


@mcp.tool()
def change_entity_layer(handle: str, layer: str) -> dict:
    """
    Cambia la capa de una entidad.

    Args:
        handle: Handle de la entidad.
        layer: Nombre de la capa destino.
    """
    acad = _get_acad()
    doc = _active_doc(acad)
    entity = doc.HandleToObject(handle)
    entity.Layer = layer
    return {"handle": handle, "new_layer": layer}


@mcp.tool()
def change_entity_color(handle: str, color: int) -> dict:
    """
    Cambia el color de una entidad (número ACI).

    Args:
        handle: Handle de la entidad.
        color: Color ACI (0=ByBlock, 1=rojo, 2=amarillo, 3=verde, 256=ByLayer).
    """
    acad = _get_acad()
    doc = _active_doc(acad)
    entity = doc.HandleToObject(handle)
    entity.color = color
    return {"handle": handle, "new_color": color}


@mcp.tool()
def delete_entity(handle: str) -> dict:
    """
    Elimina una entidad del dibujo.

    Args:
        handle: Handle de la entidad a eliminar.
    """
    acad = _get_acad()
    doc = _active_doc(acad)
    entity = doc.HandleToObject(handle)
    entity.Delete()
    return {"deleted": handle}


# ---------------------------------------------------------------------------
# Herramientas — Vista y comandos
# ---------------------------------------------------------------------------

@mcp.tool()
def zoom_extents() -> dict:
    """Ajusta la vista para mostrar todas las entidades del dibujo."""
    acad = _get_acad()
    acad.ZoomExtents()
    return {"zoomed": "extents"}


@mcp.tool()
def zoom_window(x1: float, y1: float, x2: float, y2: float) -> dict:
    """
    Hace zoom a una ventana rectangular.

    Args:
        x1, y1: Esquina inferior izquierda.
        x2, y2: Esquina superior derecha.
    """
    acad = _get_acad()
    acad.ZoomWindow(_point(x1, y1, 0), _point(x2, y2, 0))
    return {"zoomed": "window"}


@mcp.tool()
def run_autocad_command(command: str) -> dict:
    """
    Envía un comando directamente a la línea de comandos de AutoCAD.
    Usar con cuidado — comandos interactivos pueden quedar esperando input.

    Args:
        command: Comando AutoCAD (ej: "_ZOOM E ").
                 Terminar con espacio para confirmar.
    """
    acad = _get_acad()
    doc = _active_doc(acad)
    doc.SendCommand(command)
    return {"command_sent": command}


@mcp.tool()
def get_selection_set(name: str = "SS1") -> list[dict]:
    """
    Devuelve las entidades actualmente seleccionadas en AutoCAD.

    Args:
        name: Nombre interno del selection set.
    """
    acad = _get_acad()
    doc = _active_doc(acad)
    try:
        ss = doc.SelectionSets.Item(name)
        ss.Delete()
    except Exception:
        pass
    ss = doc.SelectionSets.Add(name)
    ss.SelectOnScreen()
    return _flatten(ss)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    mcp.run(transport="stdio")
