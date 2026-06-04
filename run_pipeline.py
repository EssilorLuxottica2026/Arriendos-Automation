from pathlib import Path
import sys
import logging

# Configurar logs para mostrarlos en la consola
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)]
)

# Agregar la carpeta 'src' al path de Python para poder importar 'lease_accounting'
BASE_DIR = Path(__file__).resolve().parent
SRC_DIR = BASE_DIR / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

try:
    from lease_accounting.pipeline.processor import LeaseAccountingPipeline, PipelineError
except ImportError as exc:
    print(f"Error al importar el pipeline. Asegúrate de estar en el directorio correcto. Detalle: {exc}")
    sys.exit(1)


def main():
    print("=" * 60)
    print("INICIANDO EJECUCIÓN MANUAL DEL PIPELINE DE ARRIENDOS")
    print("=" * 60)

    INPUT_DIR = BASE_DIR / "data" / "input"
    OUTPUT_DIR = BASE_DIR / "data" / "output"

    # Verificar que exista la carpeta de entrada
    if not INPUT_DIR.exists():
        print(f"ERROR: No se encuentra la carpeta de entrada en: {INPUT_DIR.resolve()}")
        sys.exit(1)

    # 1. Definir rutas a los archivos obligatorios
    invoices_path = INPUT_DIR / "inovices_file.xlsx"
    control_path = INPUT_DIR / "CONTROL_ARRI_ADMON.xlsx"
    contracts_path = INPUT_DIR / "Contratos_con_condiciones.xlsx"
    prorateo_path = INPUT_DIR / "PRORATEO.xlsx"
    distribution_path = INPUT_DIR / "Cuadro_de_distribucion.xls"

    # Verificar existencia de archivos obligatorios
    missing_required = []
    for name, path in [
        ("Facturas (inovices_file.xlsx)", invoices_path),
        ("Control (CONTROL_ARRI_ADMON.xlsx)", control_path),
        ("Contratos (Contratos_con_condiciones.xlsx)", contracts_path),
        ("Prorrateo (PRORATEO.xlsx)", prorateo_path),
        ("Distribución (Cuadro_de_distribucion.xls)", distribution_path)
    ]:
        if not path.exists():
            missing_required.append(name)

    if missing_required:
        print("ERROR: Faltan archivos obligatorios en la carpeta 'data/input/':")
        for item in missing_required:
            print(f"  - {item}")
        sys.exit(1)

    # 2. Definir rutas a los archivos opcionales (si existen)
    history_path = INPUT_DIR / "FACTURAS_CONTABILIZADAS.xlsx"
    if not history_path.exists():
        print("Aviso: No se detectó archivo de historial (FACTURAS_CONTABILIZADAS.xlsx). Se omitirá.")
        history_path = None

    macro_template_path = INPUT_DIR / "Arriendos_Macro.xlsm"
    if not macro_template_path.exists():
        print("Aviso: No se detectó plantilla de macro (Arriendos_Macro.xlsm). Se omitirá.")
        macro_template_path = None

    # 3. Detectar soportes XML y PDF dinámicamente
    support_paths = []
    for ext in ["*.xml", "*.pdf"]:
        support_paths.extend(INPUT_DIR.glob(ext))
    
    if support_paths:
        print(f"Soportes físicos detectados en 'data/input/': {len(support_paths)} archivo(s)")
        for p in support_paths:
            print(f"  - {p.name}")
    else:
        print("Aviso: No se detectaron soportes XML o PDF en 'data/input/'.")

    # Asegurar que exista la carpeta de salida
    OUTPUT_DIR.mkdir(exist_ok=True)

    print("\nProcesando...")
    try:
        pipeline = LeaseAccountingPipeline(output_dir=OUTPUT_DIR)
        
        # Ejecutar el pipeline (el período se infiere automáticamente si se pasa None)
        result = pipeline.run(
            invoices_path=invoices_path,
            control_path=control_path,
            contracts_path=contracts_path,
            prorateo_path=prorateo_path,
            distribution_path=distribution_path,
            history_path=history_path,
            support_paths=support_paths if support_paths else None,
            macro_template_path=macro_template_path,
            period=None
        )

        print("\n" + "=" * 60)
        print("¡PROCESAMIENTO COMPLETADO EXITOSAMENTE!")
        print("=" * 60)
        print(f"Filas generadas en el resumen: {result.get('output_rows')}")
        print(f"Filas de cabeceras generadas: {result.get('header_rows')}")
        print(f"Filas de validación: {result.get('validation_rows')}")
        print("\nArchivos generados en 'data/output/':")
        print(f"  - Excel Principal: {result.get('output_excel')}")
        print(f"  - CSV de salida: {result.get('output_csv')}")
        print(f"  - CSV de cabeceras: {result.get('header_csv')}")
        print(f"  - Reporte de validaciones: {result.get('validation_excel')}")
        print(f"  - Insumos normalizados: {result.get('normalized_invoices_excel')}")
        
        print("\nLogs detallados del procesamiento:")
        for log_line in result.get("logs", []):
            print(f"  > {log_line}")

    except PipelineError as exc:
        print(f"\nERROR CONTROLADO DEL PIPELINE:\n{exc}")
    except Exception as exc:
        print(f"\nERROR INESPERADO AL PROCESAR:\n{exc}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
