import os
import json
import logging
import gspread

logger = logging.getLogger(__name__)

class SheetsService:
    SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")

    def __init__(self, spreadsheet_id: str = None):
        self.spreadsheet_id = spreadsheet_id or self.SPREADSHEET_ID
        self.client = self._autenticar()
        
        try:
            self.spreadsheet = self.client.open_by_key(self.spreadsheet_id)
            
            self.sheet = self.spreadsheet.worksheet("Diário de Estudos")
            
            logger.info("Conectado com sucesso via ID na aba 'Diário de Estudos'!")
        except Exception as e:
            logger.error(f"Erro ao abrir a planilha pelo ID: {e}")
            raise e

    def _autenticar(self):
        creds_env = os.getenv("GOOGLE_CREDENTIALS_JSON")
        if creds_env:
            return gspread.service_account_from_dict(json.loads(creds_env))
        return gspread.service_account(filename="credentials.json")

    def adicionar_sessao_estudo(self, data_str, categoria, materia, tempo_horas, questoes, acertos, taxa, observacao) -> bool:
        try:
            nova_linha = [
                str(data_str),
                str(categoria),
                str(materia),
                tempo_horas,
                questoes,
                acertos,
                taxa,
                str(observacao)
            ]
            self.sheet.append_row(nova_linha, value_input_option="USER_ENTERED")
            logger.info(f"Sessão gravada com sucesso: {materia}")
            return True
        except Exception as e:
            logger.error(f"Erro ao salvar na planilha: {e}")
            return False