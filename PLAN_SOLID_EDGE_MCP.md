# Plan: bolígrafo de clic (Diseño 3) en Solid Edge con SolidEdge-MCP

Este plan lo ejecuta una sesión de Claude Code en la PC con Windows que tiene Solid Edge.
Las medidas salen de `medidas.json`, que está en esta misma carpeta; no inventes valores. Si una cota
no está en el JSON, pregunta antes de modelar.

Alcance: solo el Diseño 3 (clic con resorte), que tiene 7 piezas. Los diseños 1 y 2 quedan fuera.

```
Practica_Boligrafo/
├── medidas.json
├── PLAN_SOLID_EDGE_MCP.md      ← este archivo
└── D3_clic/
    ├── d3_cartucho.par
    ├── d3_cuerpo_inferior.par
    ├── d3_cuerpo_superior.par
    ├── d3_pulsador.par
    ├── d3_leva.par
    ├── d3_resorte.par
    ├── d3_clip.par
    └── capturas/                ← PNG para el Word
```

---

## 0. Instalación del MCP (una sola vez)

Requisitos: Windows 10/11, Solid Edge 2025 o 2026 instalado, con licencia y abierto al menos una vez
(así registra el servidor COM), Python 3.11+ y `uv`.

```powershell
winget install astral-sh.uv            # si no está uv
git clone https://github.com/tylerwagler/SolidEdge-MCP C:\tools\SolidEdge-MCP
cd C:\tools\SolidEdge-MCP
uv sync --all-extras
claude mcp add solidedge -- uv --directory C:/tools/SolidEdge-MCP run solidedge-mcp
```

Reinicia Claude Code y comprueba con `/mcp` que `solidedge` aparece conectado.
Si no arranca, ejecuta `uv --directory C:/tools/SolidEdge-MCP run solidedge-mcp` en una terminal y lee
lo que imprime en stderr. Para ver los tracebacks, define `SOLIDEDGE_MCP_DEBUG=1`.

**Ruta del escritorio:** en Windows el escritorio suele estar redirigido a OneDrive. Obtén la ruta real así:

```powershell
[Environment]::GetFolderPath('Desktop')
```

Usa esa ruta con `/` en todas las llamadas, por ejemplo `C:/Users/<usuario>/OneDrive/Desktop/Practica_Boligrafo/D3_clic/d3_leva.par`.

---

## 1. Reglas del MCP que más trampas ponen (leer antes de empezar)

1. **Las unidades son METROS**: cada valor en mm del JSON se divide entre 1000 (por ejemplo, 9.6 mm = 0.0096).
   Este es el error más probable. Los ángulos van en grados.
2. **Los planos son 1-based**: 1=Top (XY), 2=Right (YZ), 3=Front (XZ). Las caras, aristas y features son 0-based.
   En `sketch_constraint` los elementos son 1-based: `[["line", 1]]`.
3. **Los nombres de las herramientas no son los de la guía interna.** La guía `solidedge://guide/workflows`
   menciona `create_sketch`, `draw_line`, `close_sketch`, pero las herramientas reales son:
   - `manage_sketch(action="create", plane="Front")`
   - `draw(shape="line", x1, y1, x2, y2)`, `draw(shape="circle", center_x, center_y, radius)`
   - `manage_sketch(action="set_axis", x1, y1, x2, y2)`, que define el eje de revolución y **debe llamarse ANTES de cerrar**
   - `manage_sketch(action="close")`
   - `create_revolve`, `create_extrude`, `create_helix`, `create_ref_plane`
   - `create_document(type="part")`, `save_document(method="save", file_path=..., overwrite=False)`
4. **Nunca le pases a Solid Edge una ruta que ya existe.** Abre un diálogo modal ("¿sobrescribir?") que
   bloquea el servidor; el cliente acaba diciendo `Connection closed`. El MCP rechaza la llamada si no pones
   `overwrite=true`. Úsalo solo si de verdad quieres reemplazar el archivo.
5. **Cada resultado es un dict.** Si trae la clave `"error"`, detente y corrige antes de seguir. No encadenes
   llamadas a ciegas.
6. **`close` puede devolver `validation_code` distinto de 0** (por ejemplo -103) aunque el perfil esté bien,
   porque une solos los extremos coincidentes. Lo que cuenta es que la revolución diga `status: "created"`.
   El MCP verifica la geometría (caras y volumen) y reporta error si no se creó nada.
7. **Cosas que NO funcionan por COM, así que no las intentes:**
   - `create_thread` sobre un cilindro existente (E_INVALIDARG). Las roscas se representan como un escalón
     de diámetro (la "espiga" del cuerpo inferior) y en el Word se anota la rosca.
   - Shell / pared delgada por selección de cara. Por eso los tubos se dibujan **ya huecos** en el medio perfil.
   - Patrones y simetrías a nivel de ensamble.
   - Métodos `by_keypoint` de extrude y revolve.
8. **No hay herramienta de "Cota inteligente" dentro del boceto de pieza.** El MCP dibuja la geometría
   en las coordenadas exactas del JSON, pero no coloca cotas en el boceto. Las cotas se ponen así:
   - **Opción A (recomendada para las capturas del Word):** al terminar cada pieza, el usuario abre el
     boceto en Solid Edge (clic derecho en el boceto, "Editar perfil") y pone las cotas a mano con
     **Cota inteligente**. Como la geometría ya mide exacto, solo hay que hacer clic en cada línea.
   - **Opción B (automática):** crear un plano (`create_document(type="draft")` con
     `add_drawing_view(type="part", orientation="Front")`) y acotar con `add_2d_dimension` o
     `add_dimension_annotation`. Las coordenadas de esas herramientas son de la HOJA, no de la pieza, y
     tienen escala, así que es más laborioso.
9. **El plano Front tiene la normal invertida** (apunta a −Y). En una revolución de 360° no importa.
   En extrusiones sobre Front usa `direction="Symmetric"`.
10. **Las coordenadas del boceto en Front:** `x` es el radio (distancia al eje) y `y` la posición axial.
    El eje de revolución siempre va de `(0, 0)` a `(0, largo)`.

---

## 2. Procedimiento base por pieza (revolución del medio perfil)

Para cada pieza que tiene `perfil_mm` en el JSON:

```
1. manage_connection(action="connect")                  # solo la primera vez o si Solid Edge se reinició
2. create_document(type="part")
3. manage_sketch(action="create", plane="Front")
4. Por cada par de puntos consecutivos del perfil (y el último con el primero):
     draw(shape="line", x1=P[i].x/1000, y1=P[i].y/1000, x2=P[i+1].x/1000, y2=P[i+1].y/1000)
5. Relaciones: sketch_constraint(type="geometric", constraint_type="Vertical"|"Horizontal",
     elements=[["line", n]]) en cada línea que sea vertical u horizontal (n es 1-based, en el orden en que se dibujó).
     Las líneas inclinadas (conos) no llevan esta relación.
6. manage_sketch(action="set_axis", x1=0, y1=0, x2=0, y2=<largo>/1000)
7. manage_sketch(action="close")                          # revisar has_revolution_axis = true
8. create_revolve(method="finite", angle=360)             # si falla, probar method="full"
9. camera_control(action="set_orientation", view="Iso") y luego camera_control(action="zoom_fit")
10. save_document(method="save", file_path=".../D3_clic/d3_<pieza>.par")
11. export_file(format="image", file_path=".../D3_clic/capturas/d3_<pieza>_3d.png", width=1600, height=1200)
12. Verificar: leer el recurso solidedge://mass-properties y comparar el volumen con el esperado
    (ver la tabla de la sección 4). Si difiere más de un 5 %, algo quedó mal dibujado.
```

Antes de cada pieza, cierra la anterior con `close_document` para que el siguiente `create_document` no
se confunda de documento activo.

Si una línea del perfil toca el eje (`x = 0`), esa línea va sobre el eje y está bien: la pieza es maciza
en ese tramo.

---

## 3. Instrucciones por pieza (orden recomendado)

Todas las coordenadas están en **mm**; divídelas entre 1000 al llamar al MCP.

### 3.1 Cartucho (`d3_cartucho.par`), ISO 12757 tipo G2
- Perfil: `cartucho.perfil_mm` = (0,0) (1.27,0) (1.27,6.2) (2.9,23.2) (3.0,23.2) (3.0,98.1) (0,98.1)
- Eje: (0,0) a (0,98.1). Revolución de 360°.
- Es una pieza maciza (el perfil toca el eje).
- Opcional: `create_round(method="on_face", ...)` en la punta para simular la bola. No es necesario.

### 3.2 Cuerpo inferior (`d3_cuerpo_inferior.par`)
- Perfil: (1.5,0) (2.25,0) (4.8,14) (4.8,52) (4.0,52) (4.0,60) (3.3,60) (3.3,14) (1.5,2)
- Eje: (0,0) a (0,60).
- Pieza hueca: el perfil no toca el eje y de ahí sale el tubo.
- Las líneas (2.25,0)→(4.8,14) y (3.3,14)→(1.5,2) son los conos; no les pongas Vertical/Horizontal.
- La espiga Ø8 × 8 mm representa la rosca de unión. **No uses `create_thread`.**

### 3.3 Cuerpo superior (`d3_cuerpo_superior.par`)
- Perfil: (4.0,0) (4.8,0) (4.8,68) (3.1,68) (3.1,64) (3.5,64) (3.5,8) (4.0,8)
- Eje: (0,0) a (0,68).
- Abajo: alojamiento Ø8 × 8 para la espiga. Arriba: orificio Ø6.2 por donde sale el pulsador.

### 3.4 Pulsador (`d3_pulsador.par`)
- Perfil: (0,0) (3.4,0) (3.4,2) (3.0,2) (3.0,20) (0,20)
- Eje: (0,0) a (0,20).
- Opcional: `create_round(method="all_edges", radius=0.0005)` para redondear el botón.

### 3.5 Leva (`d3_leva.par`), dos operaciones
1. Cilindro base: perfil (0,0) (3.3,0) (3.3,10) (0,10), eje (0,0) a (0,10), revolución de 360°.
2. Dientes (corona de 8 dientes en diente de sierra, como vista superior aparte):
   - Antes de crear el plano, comprueba dónde quedó la cara superior: lee `solidedge://features` o usa
     `query_body` y mira el rango. Si el eje de la pieza quedó en el Z global, la cara superior está en Z = 10 mm.
   - `create_ref_plane(method="offset", parent_plane_index=1, distance=0.010)`, que devuelve `new_plane_index` (normalmente 4).
   - `manage_sketch(action="create_on_plane", plane_index=<new_plane_index>)`
   - Dibuja el polígono cerrado de 24 vértices (ya en **metros**) con `draw(shape="line")` entre puntos consecutivos y cierra el último con el primero:
     ```
     [[0.0033,0.0],[0.003228,0.000686],[0.002068,0.001736],
      [0.002333,0.002333],[0.001797,0.002768],[0.000235,0.00269],
      [0.0,0.0033],[-0.000686,0.003228],[-0.001736,0.002068],
      [-0.002333,0.002333],[-0.002768,0.001797],[-0.00269,0.000235],
      [-0.0033,0.0],[-0.003228,-0.000686],[-0.002068,-0.001736],
      [-0.002333,-0.002333],[-0.001797,-0.002768],[-0.000235,-0.00269],
      [0.0,-0.0033],[0.000686,-0.003228],[0.001736,-0.002068],
      [0.002333,-0.002333],[0.002768,-0.001797],[0.00269,-0.000235]]
     ```
     Así se generaron: 8 dientes cada 45°; por diente, punta en R 3.3 a 0° y 12°, y raíz en R 2.7 a 40°.
   - `manage_sketch(action="close")`, luego `create_extrude(method="finite", distance=0.0015, direction="Normal")`.
     Si los dientes salen hacia dentro de la leva (el volumen no sube), deshaz con `undo_redo` y repite con `direction="Reverse"`.
   - Si el plano offset no queda en la cara correcta, prueba `distance` negativa o `normal_side="Reverse"`.

### 3.6 Resorte (`d3_resorte.par`), helicoidal
- Diámetro exterior 4.5, alambre 0.4, largo libre 18, 10 espiras, paso 1.8.
- Radio medio = (4.5 − 0.4) / 2 = 2.05 mm.
```
create_document(type="part")
manage_sketch(action="create", plane="Front")
draw(shape="circle", center_x=0.00205, center_y=0.0, radius=0.0002)     # sección del alambre
manage_sketch(action="set_axis", x1=0, y1=0, x2=0, y2=0.018)
manage_sketch(action="close")
create_helix(method="finite", pitch=0.0018, height=0.018, direction="Right")
```
- Esta secuencia (círculo en Front + `set_axis` + `create_helix finite`) está verificada contra Solid Edge 2026 en el `LIVE_SWEEP` del repo.
- Si el helicoide falla, el plan B es dibujar en 2D un rectángulo de 4.5 × 18 (como dice el plan original) y anotar los datos del resorte.

### 3.7 Clip (`d3_clip.par`), extrusión (no revolución)
- Perfil lateral: (0,0) (0.8,0) (0.8,45) (−2.0,45) (−2.0,44.2) (0,44.2)
- Boceto en Front, **sin** `set_axis`.
- `create_extrude(method="finite", distance=0.004, direction="Symmetric")`: 4 mm de ancho, simétrico por la peculiaridad de la normal de Front.
- Opcional: `create_round(method="all_edges", radius=0.0002)`.

---

## 4. Volúmenes esperados para verificar (aproximados, en mm³)

Calculados con el teorema de Pappus sobre los perfiles del JSON (clip: área × 4 mm; resorte: sección del alambre × largo de la hélice). Un error de unidades (mm en vez de m) da
volúmenes 10⁹ veces mayores, así que salta a la vista.

| Pieza | Volumen esperado aprox. |
|---|---|
| cartucho | ≈ 2 390 mm³ |
| cuerpo_inferior | ≈ 1 910 mm³ |
| cuerpo_superior | ≈ 2 240 mm³ |
| pulsador | ≈ 580 mm³ |
| leva (cilindro, sin dientes) | ≈ 340 mm³ (con dientes ≈ 385 mm³) |
| resorte | ≈ 16 mm³ |
| clip | ≈ 150 mm³ |

`solidedge://mass-properties` devuelve m³: 1 mm³ = 1e-9 m³.

---

## 5. Opcional: ensamble para comprobar que encaja

Solo si sobra tiempo:
```
create_document(type="assembly")
add_assembly_component(file_path=".../d3_cuerpo_inferior.par", x=0, y=0, z=0)
add_assembly_component(file_path=".../d3_cuerpo_superior.par", x=0, y=0, z=0.052)
add_assembly_component(file_path=".../d3_cartucho.par",        x=0, y=0, z=0.002)
...
query_component(property="interference")
```
Las posiciones Z dependen de cómo quedó orientado el eje en cada pieza; verifica con el rango del cuerpo
antes de colocarlas. No uses patrones ni simetrías de ensamble porque dan E_ACCESSDENIED.

---

## 6. Entregables y cierre

1. Siete archivos `.par` en `D3_clic/` con los nombres de la convención.
2. Capturas en `D3_clic/capturas/`: la vista 3D (generada por el MCP) y el boceto acotado (tomado a mano
   después de poner las cotas inteligentes, opción A del punto 1.8).
3. Comprobar que cada `.par` coincide con `medidas.json` (sección 4 y el rango de la pieza).
4. En el Word, en la tabla de medidas, poner una columna **Fuente** con el `nivel` de cada dato
   (normativo / comercial / estimado) y citar:
   - ISO 12757-1:2017, Tabla 3 (cartucho G2)
   - Ficha Parker Jotter (largo 129 mm, Ø 9.6 mm, 15 g)
   - Listado comercial de resortes (alambre 0.4, Ø ext 4.5, largo 18)

Las URL están en `medidas.json` → `fuentes`.
