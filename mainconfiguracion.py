import HU.HU00_DespliegueAmbiente
from HU.HU06_ValidacionPresupuesto import HU01_Prueba 
from repositorios.excel import Excel as ExcelRepo
from funciones.EnvioCorreos import EmailCorreos

def main():
    #HU01_Prueba()

    email = EmailCorreos()
    email.EnviarCorreoCod(1)
                



if __name__ == "__main__":
    main()
