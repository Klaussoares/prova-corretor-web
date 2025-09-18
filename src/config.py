import os
from datetime import datetime

# Configurações de segurança
EMAILS_AUTORIZADOS = [
    'klausoares@hotmail.com',
    'natalli.plens@eaportal.org',
    'teste@teste.com'  # Para testes
]

# Configurações da API Gemini
GEMINI_API_KEY = os.environ.get('GEMINI_API_KEY')

# Configurações de logging
LOG_DIR = 'logs'
LOG_FILE = os.path.join(LOG_DIR, 'execucoes.log')

# Configurações de upload
MAX_FILE_SIZE = 50 * 1024 * 1024  # 50MB
ALLOWED_EXCEL_EXTENSIONS = ['.xlsx', '.xls']
ALLOWED_PDF_EXTENSIONS = ['.pdf']

# Coordenadas padrão para cortes das imagens (baseadas no código original)
BOX_NOME = (766, 359, 1512, 498)
BOX_MODELO = (1722, 918, 1941, 1025)
BOX_RESPOSTA = (815, 609, 1807, 824)

# Configurações de limpeza automática
CLEANUP_DELAY_SECONDS = 60  # Tempo para manter arquivos temporários

def is_email_authorized(email):
    """Verifica se o e-mail está na lista de autorizados."""
    return email.lower() in [e.lower() for e in EMAILS_AUTORIZADOS]

def get_log_entry(email, action, details=None, error=None):
    """Cria uma entrada de log formatada."""
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    entry = f"[{timestamp}] E-mail: {email} | Ação: {action}"
    
    if details:
        entry += f" | Detalhes: {details}"
    
    if error:
        entry += f" | Erro: {error}"
    
    return entry

