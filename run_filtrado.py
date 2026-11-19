# -*- coding: utf-8 -*-
"""
CLI: genera horarios por sede y por grupo de semanas, y un ZIP por sede.

    python run_filtrado.py [carpeta_de_entrada]

Edite `SEDES_FILTRO` para filtrar sedes específicas (`None` procesa todas las
sedes detectadas en los archivos). Los rangos de semanas se detectan
automáticamente desde la columna `SEMANAS`.
Resultado: `Horarios_<SEDE>/` y `Horarios_<SEDE>.zip` en la carpeta indicada.
"""
import os
import sys

from horarios.pipeline import generar_por_grupos_semanas

# None = procesar todas las sedes detectadas en los Excel de entrada.
# Para filtrar sedes específicas, use p. ej.: SEDES_FILTRO = {"AME"}
SEDES_FILTRO: set[str] | None = None

WORKDIR = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))


def _imprimir_progreso(fraccion: float, mensaje: str) -> None:
    sedes_lbl = sorted(SEDES_FILTRO) if SEDES_FILTRO else "Todas"
    print(f"\n{'=' * 60}\n  [{fraccion:4.0%}] {mensaje}  |  Sedes: {sedes_lbl}\n{'=' * 60}")


def main() -> None:
    zips = generar_por_grupos_semanas(WORKDIR, SEDES_FILTRO, grupos=None, on_progress=_imprimir_progreso)
    for z in zips:
        print(f"  {os.path.basename(z)} generado ({os.path.getsize(z) // 1024} KB)")
    print("\nProceso completo.")


if __name__ == "__main__":
    main()
