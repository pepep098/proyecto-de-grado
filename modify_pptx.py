
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
import copy

PPTX_PATH = r"c:\Users\pipeh\Downloads\ppi\EBNSP-90 A-BLACK.pptx"
OUTPUT_PATH = r"c:\Users\pipeh\Downloads\ppi\EBNSP-90 A-BLACK.pptx"

prs = Presentation(PPTX_PATH)

print(f"Total de diapositivas: {len(prs.slides)}")

# Función para obtener el tamaño de fuente efectivo de un run
def get_effective_font_size(run, paragraph, placeholder):
    """Obtiene el tamaño de fuente efectivo, buscando en la jerarquía"""
    if run.font.size:
        return run.font.size
    if paragraph.runs and paragraph.font.size:
        return paragraph.font.size
    return None

# =========================================================
# DIAPOSITIVA 1: Pintar todas las letras de NEGRO
# =========================================================
slide1 = prs.slides[0]
print("\n--- DIAPOSITIVA 1: Cambiando color de texto a negro ---")

for shape in slide1.shapes:
    if shape.has_text_frame:
        print(f"  Shape: '{shape.name}' - Texto: '{shape.text_frame.text[:50]}'")
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.color.rgb = RGBColor(0, 0, 0)  # Negro
                print(f"    Run: '{run.text[:30]}' -> color negro")

# =========================================================
# DIAPOSITIVAS 2-10: Aumentar tamaño de letra
# =========================================================
SIZE_INCREASE = 4  # Aumentar en 4 puntos (ajusta si necesitas más)
MIN_SIZE_PT = 18   # Tamaño mínimo resultado

print("\n--- DIAPOSITIVAS 2-10: Aumentando tamaño de fuente ---")

for slide_index in range(1, min(10, len(prs.slides))):  # índice 1-9 = diapositivas 2-10
    slide = prs.slides[slide_index]
    print(f"\n  Diapositiva {slide_index + 1}:")
    
    for shape in slide.shapes:
        if shape.has_text_frame:
            print(f"    Shape: '{shape.name}'")
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    current_size = run.font.size
                    if current_size is not None:
                        current_pt = current_size.pt
                        new_pt = current_pt + SIZE_INCREASE
                        run.font.size = Pt(new_pt)
                        print(f"      Run '{run.text[:20]}': {current_pt}pt -> {new_pt}pt")
                    else:
                        # Si no tiene tamaño definido, establecer uno base aumentado
                        # Intentamos obtener el tamaño del párrafo o del placeholder
                        para_size = paragraph.font.size
                        if para_size:
                            new_pt = para_size.pt + SIZE_INCREASE
                            run.font.size = Pt(new_pt)
                            print(f"      Run '{run.text[:20]}': heredado {para_size.pt}pt -> {new_pt}pt")
                        else:
                            # Tamaño no definido, asignar tamaño por defecto aumentado
                            run.font.size = Pt(MIN_SIZE_PT)
                            print(f"      Run '{run.text[:20]}': sin tamaño -> {MIN_SIZE_PT}pt")
                
                # También ajustar el tamaño del párrafo si tiene uno definido
                if paragraph.font.size:
                    old_pt = paragraph.font.size.pt
                    new_pt = old_pt + SIZE_INCREASE
                    paragraph.font.size = Pt(new_pt)
                    print(f"    Párrafo con tamaño {old_pt}pt -> {new_pt}pt")

prs.save(OUTPUT_PATH)
print(f"\n✅ Presentación guardada en: {OUTPUT_PATH}")
