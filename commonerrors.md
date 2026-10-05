# commonerrors.md — memoria de errores comunes de este proyecto

> Este archivo registra problemas donde una sesión se atascó y cómo lo solucionó,
> para no repetir los mismos tropezones. **Léelo antes de empezar una tarea nueva.**
> Añadir entradas concisas al final de una tarea (síntoma, causa, solución) — nunca a mitad de tarea.
> Formato: bullet categorizado, lectura rápida. Aún no hay entradas.

- **fpdf2 — esquinas redondeadas**: no existe `pdf.rounded_rect()` (aunque la doc parece sugerirlo). Usar `pdf.rect(x, y, w, h, style="DF", round_corners=True, corner_radius=r)`.
- **fpdf2 — nombre de import**: el paquete pip es `fpdf2` pero se importa como `import fpdf` (no `import fpdf2`).
- **Matemática de limpieza (atajo sutil)**: si la clave de búsqueda se des-acenta (`unicodedata` NFKD → ascii), las claves del diccionario canónico también deben estar sin acentos — si no, "Plomería" (con acento en el dict, des-acentada en la búsqueda) caía a `None` y descartaba 47/47 filas de esa categoría. Síntoma: filas válidas desaparecen sin log de descarte.
- **Alturas de imagen en fpdf2**: con `savefig(bbox_inches="tight")` cada PNG sale con dimensiones distintas y `pdf.image(w=88)` calcula alturas impredecibles → solapos. Solución: `fig.tight_layout()` + `savefig` SIN `bbox_inches="tight"` → ratio determinista (9.5×4.0 → 182×76,6 mm; 5.0×3.6 → 88×63,4 mm a dpi 150).
- **Formato español en matplotlib**: `autopct="%1.1f %%"` saca punto decimal; usar una lambda: `autopct=lambda p: f"{p:.1f}".replace(".", ",") + " %"`.
- **Fuentes de sistema en Windows**: Consolas se llama `consola.ttf` (`consolab.ttf` la bold), no `consolas.ttf`. Fallback multiplataforma: la DejaVu Sans (Mono) que trae matplotlib.
- **pandas 3.0**: `read_excel` devuelve columnas de texto con dtype `str` (no `object`); no afecta a nada del pipeline actual, pero `isinstance(v, str)` sí funciona para detectar fechas escritas como texto.