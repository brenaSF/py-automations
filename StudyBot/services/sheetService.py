import os
import json
import logging
import gspread
import dotenv

dotenv.load_dotenv()

logger = logging.getLogger(__name__)


class SheetsService:
    SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")

    def __init__(self, spreadsheet_id: str = None):
        self.spreadsheet_id = spreadsheet_id or self.SPREADSHEET_ID
        self.client = None
        self.sheet = None
        self._conectar()

    def _conectar(self):
        try:
            self.client = self._autenticar()
            if self.spreadsheet_id:
                self.spreadsheet = self.client.open_by_key(self.spreadsheet_id)
                self.sheet = self.spreadsheet.worksheet("Diário de Estudos")
                logger.info("Conectado com sucesso via ID na aba 'Diário de Estudos'!")
            else:
                logger.error("SPREADSHEET_ID não foi informado nas variáveis de ambiente.")
        except Exception as e:
            logger.error(f"Erro ao conectar ao Google Sheets: {e}")

    def _autenticar(self):
        creds_env = os.getenv("GOOGLE_CREDENTIALS_JSON")
        
        if creds_env:
            try:
                creds_cleaned = creds_env.strip("'\"")
                creds_dict = json.loads(creds_cleaned)
                return gspread.service_account_from_dict(creds_dict)
            except json.JSONDecodeError as e:
                logger.error("A variável GOOGLE_CREDENTIALS_JSON não é um JSON válido.")
                raise e
        
        if os.path.exists("credentials.json"):
            return gspread.service_account(filename="credentials.json")
            
        raise FileNotFoundError(
            "Credenciais não encontradas. Configure 'GOOGLE_CREDENTIALS_JSON' ou adicione 'credentials.json'."
        )

    def adicionar_sessao_estudo(self, data_str, categoria, materia, tempo_horas, questoes, acertos, taxa, observacao) -> bool:
        if not self.sheet:
            self._conectar()
            if not self.sheet:
                logger.error("Aba do Google Sheets não está acessível.")
                return False

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