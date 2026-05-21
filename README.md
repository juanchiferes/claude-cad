# claude-cad

Servidor MCP que conecta Claude con AutoCAD vía COM/ActiveX.
Funciona en **Windows** con AutoCAD instalado.

---

## Requisitos

- Windows 10/11
- AutoCAD (cualquier versión moderna con COM habilitado)
- Python 3.10+

---

## Instalación

```bash
git clone https://github.com/juanchiferes/claude-cad.git
cd claude-cad
pip install -r requirements.txt
```

---

## Configurar Claude Code

Agregá el servidor en tu configuración de Claude Code.
El archivo se encuentra en `%APPDATA%\Claude\claude_desktop_config.json`
(o `~/.claude.json` si usás la CLI):

```json
{
  "mcpServers": {
    "autocad": {
      "command": "python",
      "args": ["C:\\ruta\\a\\claude-cad\\server.py"],
      "description": "Controla AutoCAD desde Claude"
    }
  }
}
```

Reemplazá `C:\\ruta\\a\\claude-cad\\server.py` con la ruta real donde clonaste el repo.

---

## Uso

1. Abrí AutoCAD con un dibujo cargado.
2. Iniciá Claude Code (CLI o app de escritorio).
3. Claude tendrá acceso a las herramientas de AutoCAD automáticamente.

Ejemplos de lo que podés pedirle a Claude:

- *"Creá una línea desde (0,0) hasta (100,50)"*
- *"Listá todas las entidades en la capa MUROS"*
- *"Creá un círculo de radio 25 centrado en (50,50)"*
- *"Mové la entidad con handle 3A5 10 unidades hacia la derecha"*
- *"Guardá el dibujo"*

---

## Herramientas disponibles

### Estado y documentos
| Herramienta | Descripción |
|---|---|
| `get_autocad_status` | Verifica si AutoCAD está corriendo |
| `open_drawing` | Abre un archivo DWG |
| `save_drawing` | Guarda el dibujo activo |
| `close_drawing` | Cierra el dibujo |
| `get_drawing_info` | Info del dibujo: nombre, unidades, límites |

### Capas
| Herramienta | Descripción |
|---|---|
| `list_layers` | Lista todas las capas |
| `create_layer` | Crea una capa nueva |
| `set_active_layer` | Cambia la capa activa |

### Crear geometría
| Herramienta | Descripción |
|---|---|
| `create_line` | Línea entre dos puntos |
| `create_circle` | Círculo por centro y radio |
| `create_arc` | Arco por centro, radio y ángulos |
| `create_rectangle` | Rectángulo por dos esquinas |
| `create_polyline` | Polilínea 2D con N puntos |
| `create_text` | Texto simple |
| `create_mtext` | Texto multilínea |
| `insert_block` | Insertar bloque existente |

### Modificar entidades
| Herramienta | Descripción |
|---|---|
| `list_entities` | Lista entidades (filtrable por capa) |
| `get_entity_by_handle` | Detalle de una entidad por handle |
| `move_entity` | Mueve por desplazamiento |
| `scale_entity` | Escala respecto a punto base |
| `rotate_entity` | Rota respecto a punto base |
| `change_entity_layer` | Cambia la capa de una entidad |
| `change_entity_color` | Cambia el color (número ACI) |
| `delete_entity` | Elimina una entidad |

### Vista y comandos
| Herramienta | Descripción |
|---|---|
| `zoom_extents` | Zoom a todas las entidades |
| `zoom_window` | Zoom a ventana rectangular |
| `run_autocad_command` | Envía comando directo a AutoCAD |
| `get_selection_set` | Obtiene entidades seleccionadas |
