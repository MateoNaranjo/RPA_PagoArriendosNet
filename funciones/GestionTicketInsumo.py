from config.database import Database
from repositorios.ticketInsumo import GestionTicketInsumo

class TicketInsumoService:

    MAX_REINTENTOS = 3

    def __init__(self, maquina: str):
        self.maquina = maquina

    def iniciar(self, codigo: str):

        ticket = GestionTicketInsumo.obtener_por_codigo(codigo)

        if not ticket:
            GestionTicketInsumo.crear(codigo, self.maquina)

        GestionTicketInsumo.actualizar_estado(
            codigo,
            estado="EN_PROCESO",
            observaciones="Inicio del procesamiento"
        )

    def finalizar(self, codigo: str):
        GestionTicketInsumo.actualizar_estado(
            codigo,
            estado="FINALIZADO",
            observaciones="Proceso finalizado correctamente",
            finalizar=True
        )

    def error(self, codigo: str, mensaje_error: str):
        ticket = GestionTicketInsumo.obtener_por_codigo(codigo)

        if not ticket:
            raise ValueError("Ticket no encontrado para manejar error")

        reintentos = ticket["numeroreintentos"] + 1
        estado = "REINTENTO" if reintentos < self.MAX_REINTENTOS else "ERROR"

        GestionTicketInsumo.actualizar_estado(
            codigo,
            estado=estado,
            observaciones=mensaje_error,
            incrementar_reintento=True
        )
