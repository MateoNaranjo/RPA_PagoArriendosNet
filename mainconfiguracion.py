# mainconfiguracion.py

# ⚠ IMPORTANTE:
# Esto ejecuta TODO el despliegue y carga configuración
import HU.HU00_DespliegueAmbiente  # noqa: F401
from config.init_config import in_config


def main():
    print("=== INICIO PRUEBA CONFIG ===")

    print(f'Prueba = {in_config("Prueba")}')

    print("=== FIN PRUEBA ===")


if __name__ == "__main__":
    main()
