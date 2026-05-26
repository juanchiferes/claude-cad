# claude-cad

Servidor MCP que conecta Claude con AutoCAD vía COM/ActiveX.
Diseñado para que Claude entienda pedidos en lenguaje natural, identifique el
dibujo, edite textos con precisión y compute superficies.

Funciona en **Windows** con AutoCAD instalado.

---

## Requisitos

- Windows 10/11
- AutoCAD (con COM/ActiveX habilitado — viene por defecto)
- Python 3.10+

## Instalación

```bash
git clone https://github.com/juanchiferes/claude-cad.git
cd claude-cad
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## Configurar Claude Desktop

Edita `%APPDATA%\Claude\claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "autocad": {
      "command": "C:\\Users\\Usuario\\claude-cad\\.venv\\Scripts\\python.exe",
      "args": ["C:\\Users\\Usuario\\claude-cad\\server.py"]
    }
  }
}
```

> Reemplazá `Usuario` con tu nombre de usuario de Windows.

## Uso

1. Abrí AutoCAD con un dibujo.
2. Iniciá Claude Code.
3. Claude tendrá los tools y los prompts disponibles.

Ejemplos de pedidos que Claude puede manejar:

- *"¿Qué hay en este dibujo? Resumime por capas."*
- *"En todos los textos que digan 'OFICINA' cambialos por 'DESPACHO'."*
- *"Computame las superficies de la capa LOSAS y dame el total en m²."*
- *"Hacé click en el ambiente principal y decime cuánto mide."* (después
  podés indicarle un punto interior)
- *"Creá los muros del perímetro: 10×6 m en la capa MUROS, color rojo."*
- *"Cambiale la altura a todos los textos de la capa ROTULO a 5."*

---

## Cómo interpreta los pedidos

El servidor expone **prompts** (templates) que Claude usa automáticamente para
descomponer pedidos comunes:

| Prompt | Cuándo se usa |
|---|---|
| `autocad_context` | Guía general: identificar antes de modificar, respetar unidades, capas correctas |
| `computar_superficies` | Para cómputos de áreas: ordena el flujo por capas / boundary / totales |
| `editar_texto` | Para búsqueda y edición segura de textos con confirmación previa |

---

## Herramientas

### Estado y documentos
| Tool | Descripción |
|---|---|
| `get_autocad_status` | Verifica AutoCAD y lista documentos abiertos |
| `open_drawing` | Abre un DWG |
| `save_drawing` | Guarda (o Save As) |
| `close_drawing` | Cierra el dibujo activo |
| `switch_drawing` | Cambia el dibujo activo por nombre |

### Identificación y análisis
| Tool | Descripción |
|---|---|
| `get_drawing_info` | Nombre, ruta, unidades, límites, extents, conteos |
| `analyze_drawing` | Conteo por tipo y por capa, opcionalmente con handles |
| `get_bounding_box` | BBox de una entidad o del dibujo entero |
| `list_blocks` | Definiciones de bloque (excluye layouts) |

### Capas
| Tool | Descripción |
|---|---|
| `list_layers` | Lista con estado, color, linetype, lineweight |
| `create_layer` | Crea una capa |
| `set_active_layer` | Cambia la capa activa |
| `set_layer_state` | On/frozen/locked |

### Crear geometría
| Tool | Descripción |
|---|---|
| `create_line`, `create_circle`, `create_arc` | Primitivas básicas |
| `create_rectangle`, `create_polyline`, `create_ellipse` | Formas compuestas |
| `insert_block` | Inserta un bloque existente |

### Textos
| Tool | Descripción |
|---|---|
| `create_text`, `create_mtext` | Crear texto y MText |
| `list_all_texts` | Lista todos los textos (incluye atributos de bloques opcional) |
| `find_texts_containing` | Búsqueda por subcadena |
| `edit_text_by_handle` | Reemplaza el contenido de un texto |
| `find_and_replace_text` | Buscar y reemplazar global con `dry_run` para previsualizar |
| `set_text_properties` | Cambia altura, rotación, estilo, ancho, posición |
| `list_text_styles` | Lista los estilos de texto del dibujo |

### Modificación genérica
| Tool | Descripción |
|---|---|
| `list_entities` | Filtrable por capa y/o tipo, devuelve detalles ricos |
| `get_entity_by_handle` | Detalle completo |
| `move_entity`, `copy_entity` | Mover y copiar |
| `scale_entity`, `rotate_entity`, `mirror_entity`, `offset_entity` | Transformaciones |
| `change_entity_layer`, `change_entity_color` | Cambios de propiedades |
| `delete_entity`, `delete_entities` | Borrado individual y por lotes |

### Cómputo de superficies
| Tool | Descripción |
|---|---|
| `calculate_area_of_entity` | Área y perímetro de una entidad cerrada |
| `sum_areas` | Suma de áreas de varios handles |
| `calculate_area_by_layer` | Suma total por capa |
| `calculate_areas_all_layers` | Resumen por capas con totales (ideal para informes) |
| `calculate_boundary_area_at_point` | Construye boundary cerrado desde un punto interior y mide |
| `create_region_from_polyline` | Convierte polilínea cerrada → Región |
| `region_boolean` | Unión / Resta / Intersección de regiones |

### Cotas
| Tool | Descripción |
|---|---|
| `create_dimension_aligned` | Cota alineada |
| `create_dimension_linear` | Cota lineal con rotación opcional |

### Vista y comandos
| Tool | Descripción |
|---|---|
| `zoom_extents`, `zoom_window`, `zoom_to_entity` | Navegación de vista |
| `run_autocad_command` | Comando directo (usar como último recurso) |
| `get_selection_set` | Selección interactiva del usuario en AutoCAD |

---

## Notas

- Las áreas se devuelven en las unidades del dibujo elevadas al cuadrado.
  Si tu dibujo está en mm, dividí por `1_000_000` para obtener m².
- Los handles son identificadores estables: pedile a Claude que te los muestre
  y los podés usar para deshacer manualmente o referenciar después.
- `calculate_boundary_area_at_point` usa `-BOUNDARY` internamente: el punto
  debe estar dentro de un área visualmente cerrada por líneas/polilíneas.
