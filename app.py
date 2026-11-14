import os
import json
from google import genai
import pypdf
from docx import Document

# 1. Configurar cliente
api_key = os.getenv("GEMINI_API_KEY")
if not api_key:
    raise RuntimeError("Configura la variable de entorno GEMINI_API_KEY antes de ejecutar el programa.")

client = genai.Client(api_key=api_key)

def extraer_texto_archivo(ruta_archivo):
    """Detecta la extensión y extrae todo el texto legible."""
    if not os.path.exists(ruta_archivo):
        raise FileNotFoundError(f"No se encontró el archivo: {ruta_archivo}")

    _, extension = os.path.splitext(ruta_archivo.lower())

    # Leer archivos de texto plano (.txt)
    if extension == ".txt":
        with open(ruta_archivo, "r", encoding="utf-8", errors="ignore") as f:
            return f.read()

    # Leer archivos PDF (.pdf)
    elif extension == ".pdf":
        lector = pypdf.PdfReader(ruta_archivo)
        texto = ""
        for pagina in lector.pages:
            texto_pag = pagina.extract_text()
            if texto_pag:
                texto += texto_pag + "\n"
        return texto

    # Leer archivos de Word (.docx)
    elif extension == ".docx":
        doc = Document(ruta_archivo)
        parrafos = [p.text for p in doc.paragraphs if p.text.strip()]
        return "\n".join(parrafos)

    else:
        raise ValueError("Formato no soportado. Usa archivos .pdf, .docx o .txt")

def procesar_con_ia(texto_documento):
    """Envía el texto al modelo para resumir y extraer datos estructurados genéricos."""
    
    # Prompt dinámico: se adapta al tipo de documento analizado
    prompt = f"""
Eres un sistema experto en análisis y extracción de información.
Analiza el siguiente texto (que puede ser un contrato, informe, artículo científico, noticia, acta, etc.)
y devuelve EXCLUSIVAMENTE un objeto JSON válido con la siguiente estructura:

{{
  "tipo_documento": "Tipo de archivo identificado (ej. Contrato, Artículo, Noticia, Minuta)",
  "tema_principal": "Tema central del texto",
  "resumen": "Resumen ejecutivo claro y comprensible de 2 a 4 oraciones.",
  "datos_estructurados": {{
    "fechas_relevantes": ["Lista de fechas mencionadas con su contexto"],
    "entidades_y_personas": ["Lista de personas, empresas u organizaciones clave"],
    "cifras_y_metricas": ["Valores numéricos, presupuestos, porcentajes o cantidades clave"],
    "puntos_o_acuerdos_clave": ["Lista de las 3 a 5 conclusiones, cláusulas o hallazgos principales"]
  }}
}}

Texto a analizar:
{texto_documento}
"""

    print("Analizando archivo con el modelo...")
    interaction = client.interactions.create(
        model="gemini-3.8-flash",
        input=prompt
    )

    # Extraer texto de la respuesta
    if hasattr(interaction, "output_text"):
        salida_texto = interaction.output_text
    elif hasattr(interaction, "output"):
        salida_texto = str(interaction.output)
    else:
        salida_texto = str(interaction)

    # Limpiar formato markdown si viene incluido
    salida_texto = salida_texto.strip()
    if "```json" in salida_texto:
        salida_texto = salida_texto.split("```json")[1].split("```")[0].strip()
    elif "```" in salida_texto:
        salida_texto = salida_texto.split("```")[1].split("```")[0].strip()

    return json.loads(salida_texto)

# --- EJECUCIÓN PRINCIPAL ---
if __name__ == "__main__":
    # Coloca el archivo que quieras probar en la misma carpeta del proyecto
    # y cambia el nombre aquí (puede ser .pdf, .docx o .txt):
    archivo_a_probar = "ejemplo.pdf" 

    try:
        texto = extraer_texto_archivo(archivo_a_probar)
        
        if not texto.strip():
            print("El archivo está vacío o no contiene texto editable (podría ser un escaneo de imagen).")
        else:
            resultado = procesar_con_ia(texto)

            print("\n================== TIPO Y TEMA ==================")
            print(f"Tipo: {resultado.get('tipo_documento')}")
            print(f"Tema: {resultado.get('tema_principal')}")

            print("\n================== RESUMEN ==================")
            print(resultado.get("resumen"))

            print("\n============== DATOS EXTRAÍDOS (JSON) ==============")
            print(json.dumps(resultado.get("datos_estructurados"), indent=4, ensure_ascii=False))

    except FileNotFoundError:
        print(f"Asegúrate de colocar un archivo llamado '{archivo_a_probar}' en la carpeta del proyecto para probar.")
    except Exception as e:
        print(f"Error durante el procesamiento: {e}")