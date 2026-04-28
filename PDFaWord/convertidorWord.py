import sys
import os

# ---------------------------------------------------------------------------
# Verificación de dependencias
# ---------------------------------------------------------------------------

def check_dependencies():
    missing = []
    try:
        import tkinter
    except ImportError:
        missing.append("tkinter (incluido en Python estándar; reinstala Python si falta)")

    try:
        from pdf2docx import Converter
    except ImportError:
        missing.append("pdf2docx  →  pip install pdf2docx")

    if missing:
        print("=" * 60)
        print("ERROR: Faltan las siguientes dependencias:")
        print("=" * 60)
        for m in missing:
            print(f"  · {m}")
        print("=" * 60)
        input("\nPulsa Enter para salir...")
        sys.exit(1)


# ---------------------------------------------------------------------------
# Diálogos
# ---------------------------------------------------------------------------

def init_tk():
    import tkinter as tk
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    return root


def ask_pdf_file():
    from tkinter import filedialog
    root = init_tk()
    path = filedialog.askopenfilename(
        title="Selecciona el archivo PDF",
        filetypes=[("Archivos PDF", "*.pdf"), ("Todos", "*.*")],
    )
    root.destroy()
    return path.strip() if path else None


def ask_output_name():
    name = input("Nombre del archivo de salida (sin extensión): ").strip()
    if not name:
        return None
    if not name.lower().endswith(".docx"):
        name += ".docx"
    return name


def ask_output_folder():
    from tkinter import filedialog
    root = init_tk()
    folder = filedialog.askdirectory(title="Selecciona la carpeta de destino")
    root.destroy()
    return folder.strip() if folder else None


def show_result(success, message):
    try:
        from tkinter import messagebox
        root = init_tk()
        if success:
            messagebox.showinfo("Conversión completada", message)
        else:
            messagebox.showerror("Error en la conversión", message)
        root.destroy()
    except Exception:
        print(message)


# ---------------------------------------------------------------------------
# Conversión
# ---------------------------------------------------------------------------

def convert_pdf_to_docx(pdf_path, docx_path):
    from pdf2docx import Converter

    print(f"\nConvirtiendo: {pdf_path}")
    print(f"Destino:      {docx_path}")
    print("Por favor, espera...\n")

    cv = Converter(pdf_path)
    cv.convert(docx_path, start=0, end=None)
    cv.close()


# ---------------------------------------------------------------------------
# Validaciones
# ---------------------------------------------------------------------------

def validate_pdf(path):
    if not os.path.isfile(path):
        return False, f"No se encontró el archivo: {path}"
    if not path.lower().endswith(".pdf"):
        return False, "El archivo seleccionado no tiene extensión .pdf"
    return True, ""


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    print("=" * 60)
    print("        Conversor PDF → Word (.docx)")
    print("=" * 60)

    check_dependencies()

    # 1. Seleccionar PDF
    print("\nPaso 1: Selecciona el archivo PDF de entrada...")
    pdf_path = ask_pdf_file()
    if not pdf_path:
        print("Operación cancelada: no se seleccionó ningún archivo.")
        input("\nPulsa Enter para salir...")
        sys.exit(0)

    valid, err = validate_pdf(pdf_path)
    if not valid:
        print(f"Error: {err}")
        input("\nPulsa Enter para salir...")
        sys.exit(1)

    print(f"  ✓ PDF seleccionado: {pdf_path}")

    # 2. Nombre del archivo de salida
    print("\nPaso 2: Indica el nombre del archivo Word de salida.")
    default_name = os.path.splitext(os.path.basename(pdf_path))[0] + ".docx"
    print(f"  (Deja en blanco para usar: {default_name})")
    output_name = ask_output_name()
    if not output_name:
        output_name = default_name
    print(f"  ✓ Nombre de salida: {output_name}")

    # 3. Carpeta de destino
    print("\nPaso 3: Selecciona la carpeta donde guardar el archivo Word...")
    output_folder = ask_output_folder()
    if not output_folder:
        print("Operación cancelada: no se seleccionó carpeta de destino.")
        input("\nPulsa Enter para salir...")
        sys.exit(0)

    print(f"  ✓ Carpeta destino: {output_folder}")

    # 4. Ruta final
    docx_path = os.path.join(output_folder, output_name)

    # 5. Convertir
    try:
        convert_pdf_to_docx(pdf_path, docx_path)
    except Exception as e:
        msg = f"Error durante la conversión:\n{e}"
        print(f"\n{msg}")
        show_result(False, msg)
        input("\nPulsa Enter para salir...")
        sys.exit(1)

    # 6. Éxito
    msg = f"Conversión completada con éxito.\n\nArchivo guardado en:\n{docx_path}"
    print("=" * 60)
    print("✓ CONVERSIÓN COMPLETADA CON ÉXITO")
    print(f"  Archivo guardado en: {docx_path}")
    print("=" * 60)
    show_result(True, msg)

    input("\nPulsa Enter para salir...")


if __name__ == "__main__":
    main()