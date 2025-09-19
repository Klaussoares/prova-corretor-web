import os
import io
import json
import datetime as dt
from datetime import datetime
import tempfile
import shutil
import logging
from werkzeug.utils import secure_filename
from flask import Blueprint, request, jsonify, send_file
from openpyxl import Workbook, load_workbook
from flask_cors import cross_origin
from src.config import (
    EMAILS_AUTORIZADOS,
    GEMINI_API_KEY,
    COORDS_NOVO,
    COORDS_ANTIGO,
    CLEANUP_DELAY_SECONDS,
    LOG_DIR,
    LOG_FILE,
    is_email_authorized,
    get_log_entry
)

import google.generativeai as genai

from pypdf import PdfReader, PdfWriter

if GEMINI_API_KEY:
    genai.configure(api_key=GEMINI_API_KEY)
    model = genai.GenerativeModel("gemini-1.5-flash")
else:
    model = None

prova_bp = Blueprint('prova', __name__)

# Imports condicionais para evitar erros no deploy
convert_from_path = None
Image = None
cv2 = None
np = None
pd = None

try:
    if GEMINI_API_KEY:
        import google.generativeai as genai
        genai.configure(api_key=GEMINI_API_KEY)
        model = genai.GenerativeModel(
            model_name="gemini-1.5-flash",
            system_instruction="Você é excelente ajudante para corrigir rapidamente provas de alunos."
        )
except ImportError:
    logging.warning("Google Generative AI não disponível no ambiente de deploy")

try:
    from pdf2image import convert_from_path
    from PIL import Image
    import cv2
    import numpy as np
    import pandas as pd
except ImportError:
    logging.warning("Dependências de processamento de imagem não disponíveis no ambiente de deploy")

# Configuração de logging
os.makedirs(LOG_DIR, exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler()
    ]
)

def log_action(email, action, details=None, error=None):
    """Registra uma ação no log."""
    log_entry = get_log_entry(email, action, details, error)

    # Log no arquivo
    with open(LOG_FILE, 'a', encoding='utf-8') as f:
        f.write(log_entry + '\n')

    # Log no console
    if error:
        logging.error(log_entry)
    else:
        logging.info(log_entry)


def crop_pil(img, box):
    """Corta uma imagem PIL usando as coordenadas fornecidas."""
    return img.crop(box)


def preprocessar_para_ia(imagem_pil):
    """Preprocessa a imagem para melhor reconhecimento pela IA."""
    img_cv = cv2.cvtColor(np.array(imagem_pil), cv2.COLOR_RGB2BGR)
    img_gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)
    return Image.fromarray(cv2.cvtColor(img_gray, cv2.COLOR_GRAY2RGB))


def carregar_gabaritos_excel(excel_file):
    """Carrega os gabaritos do arquivo Excel fornecido."""
    gabaritos = {}
    try:
        wb_gab = load_workbook(excel_file)
        for sheet_name in wb_gab.sheetnames:
            if "modelo" in sheet_name.lower():
                numero = ''.join(filter(str.isdigit, sheet_name))
                ws_gab = wb_gab[sheet_name]
                respostas = [
                    str(r[1]).strip().upper()
                    for r in ws_gab.iter_rows(min_row=2, max_col=2, values_only=True)
                    if r[1] and str(r[1]).strip().upper() in ['A', 'B', 'C', 'D', 'E']
                ]
                gabaritos[numero] = respostas
    except Exception as e:
        logging.error(f"Erro ao ler gabarito: {e}")
        return {}

    return gabaritos


def corrigir_prova(nome, modelo, respostas, gabaritos):
    """Corrige uma prova individual baseada no gabarito."""
    try:
        gabarito = gabaritos.get(str(modelo))
        if not gabarito:
            logging.warning(f"Modelo {modelo} não encontrado no gabarito.")
            return 0

        respostas_aluno = respostas.replace("Respostas:", "").strip().split(",")
        respostas_aluno = [r.strip().upper() for r in respostas_aluno][:len(gabarito)]

        acertos = sum(1 for a, b in zip(gabarito, respostas_aluno) if a == b)
        return acertos

    except Exception as e:
        logging.error(f"Erro ao corrigir prova de {nome} (modelo {modelo}): {e}")
        return 0


def split_pdf(pdf_path, temp_dir, pages_per_chunk=3):  # <<< AQUI ESTÁ A ÚNICA MUDANÇA >>>
    """
    Quebra um PDF grande em arquivos menores.
    Retorna uma lista de caminhos para os arquivos criados.
    """
    input_pdf = PdfReader(pdf_path)
    total_pages = len(input_pdf.pages)
    chunked_pdf_paths = []

    for start_page in range(0, total_pages, pages_per_chunk):
        writer = PdfWriter()
        end_page = min(start_page + pages_per_chunk, total_pages)

        for i in range(start_page, end_page):
            writer.add_page(input_pdf.pages[i])

        chunk_filename = f"chunk_{start_page // pages_per_chunk + 1}.pdf"
        chunk_path = os.path.join(temp_dir, chunk_filename)

        with open(chunk_path, "wb") as output_file:
            writer.write(output_file)

        chunked_pdf_paths.append(chunk_path)

    return chunked_pdf_paths


@prova_bp.route('/verificar-email', methods=['POST'])
@cross_origin()
def verificar_email():
    """Verifica se o e-mail está autorizado."""
    try:
        data = request.get_json()
        email = data.get('email', '').strip().lower()

        if is_email_authorized(email):
            log_action(email, "LOGIN_AUTORIZADO", "Acesso concedido ao sistema")
            return jsonify({'autorizado': True, 'mensagem': 'Acesso autorizado'})
        else:
            log_action(email, "LOGIN_NEGADO", "Tentativa de acesso não autorizado")
            return jsonify({'autorizado': False, 'mensagem': 'Acesso não autorizado'})

    except Exception as e:
        log_action("DESCONHECIDO", "ERRO_LOGIN", error=str(e))
        return jsonify({'erro': 'Erro interno do servidor'}), 500


@prova_bp.route('/processar', methods=['POST'])
@cross_origin()
def processar_provas():
    """Processa os arquivos Excel e PDF enviados."""
    email = request.form.get('email', '').strip()

    try:
        # Verificar se a API Gemini está configurada
        if not GEMINI_API_KEY or not model:
            log_action(email, "ERRO_PROCESSAMENTO", error="API Gemini não configurada ou não disponível")
            return jsonify({'erro': 'API Gemini não configurada ou não disponível no ambiente atual'}), 500

        # Verificar se as dependências de processamento estão disponíveis
        if not convert_from_path or not Image:
            log_action(email, "ERRO_DEPENDENCIAS", error="Dependências críticas não disponíveis")
            return jsonify({'erro': 'Dependências críticas não disponíveis no ambiente atual'}), 500

        # Verificar se os arquivos foram enviados
        if 'excel' not in request.files or 'pdf' not in request.files:
            log_action(email, "ERRO_UPLOAD", error="Arquivos Excel e PDF são obrigatórios")
            return jsonify({'erro': 'Arquivos Excel e PDF são obrigatórios'}), 400

        excel_file = request.files['excel']
        pdf_file = request.files['pdf']

        # Escolher o conjunto de coordenadas
        tipo_prova = request.form.get("tipo_prova", "novo").strip().lower()
        if tipo_prova == "antigo":
            COORDS = COORDS_ANTIGO
        else:
            COORDS = COORDS_NOVO

        # Verificar se o e-mail está autorizado
        if not is_email_authorized(email):
            log_action(email, "ACESSO_NEGADO", error="E-mail não autorizado para processamento")
            return jsonify({'erro': 'E-mail não autorizado'}), 403

        # Log do início do processamento
        log_action(email, "INICIO_PROCESSAMENTO",
                   f"Arquivos: {excel_file.filename}, {pdf_file.filename}")

        # Criar diretório temporário
        temp_dir = tempfile.mkdtemp()

        try:
            # Salvar arquivos temporariamente
            excel_path = os.path.join(temp_dir, secure_filename(excel_file.filename))
            pdf_path = os.path.join(temp_dir, secure_filename(pdf_file.filename))

            excel_file.save(excel_path)
            pdf_file.save(pdf_path)

            # Carregar gabaritos
            gabaritos = carregar_gabaritos_excel(excel_path)
            if not gabaritos:
                log_action(email, "ERRO_GABARITO",
                           error="Não foi possível carregar os gabaritos do arquivo Excel")
                return jsonify({'erro': 'Não foi possível carregar os gabaritos do arquivo Excel'}), 400

            log_action(email, "GABARITOS_CARREGADOS",
                       f"Modelos encontrados: {list(gabaritos.keys())}")
            
            try:
                log_action(email, "QUEBRANDO_PDF", "Iniciando a divisão do PDF em partes menores.")
                chunked_pdf_paths = split_pdf(pdf_path, temp_dir, pages_per_chunk=3)
                log_action(email, "PDF_QUEBRADO", f"Total de arquivos menores criados: {len(chunked_pdf_paths)}")
            except Exception as e:
                log_action(email, "ERRO_PDF_SPLIT", error=f"Erro ao quebrar o PDF: {str(e)}")
                return jsonify({'erro': 'Erro ao processar o arquivo PDF'}), 400

            # Lista para armazenar os resultados de todas as provas
            resultados_finais = []
            provas_processadas = 0

            # Processar cada pequeno PDF
            for i, chunk_path in enumerate(chunked_pdf_paths):
                try:
                    log_action(email, "PROCESSANDO_CHUNK", f"Processando arquivo {i + 1} de {len(chunked_pdf_paths)}")
                    
                    paginas = convert_from_path(chunk_path, dpi=200)
                    
                    # Processar cada página dentro do chunk
                    for j, pagina in enumerate(paginas):
                        try:
                            provas_processadas += 1
                            
                            # Cortes das imagens usando coordenadas
                            img_nome = crop_pil(pagina, COORDS["BOX_NOME"])
                            img_nome = preprocessar_para_ia(img_nome)
                            img_modelo = crop_pil(pagina, COORDS["BOX_MODELO"])
                            img_resposta = crop_pil(pagina, COORDS["BOX_RESPOSTA"])

                            # Extrair nome
                            prompt_nome = (
                                "Qual o nome completo do aluno nesta imagem? Mostre apenas o que está escrito."
                                "não considere hifens nem pontuações, apenas letras normais"
                                "Apresente o nome sempre com as iniciais maiúsculas e as demais minúsculas."
                                "NUNCA, EM HIPÓTESE ALGUMA, ESCREVA ALGO ALÉM DO NOME DO ALUNO! NUNCA!"
                                "Se possível, verifique letras que podem ser confundidas, como 'u' e 'v'. "
                                "Nomes como Kavana não existem, é Kauana"
                            )
                            resposta_nome = model.generate_content([prompt_nome, img_nome])
                            nome_texto = resposta_nome.text.strip()
                            
                            # Extrair modelo
                            prompt_modelo = (
                                "Qual é o modelo do gabarito nesta imagem (Modelo 1, Modelo 2, etc)? "
                                "Apresente apenas o numero do modelo, exemplo: '1' ou '2' ou '3'"
                            )
                            resposta_modelo = model.generate_content([prompt_modelo, img_modelo])
                            modelo_texto = resposta_modelo.text.strip()
                            
                            # Extrair respostas
                            prompt_resposta = (
                                "Liste as alternativas marcadas no cartão-resposta desta imagem.\n"
                                "Considere apenas A, B, C, D ou E.\n"
                                "Se houver duas alternativas por questão, mostre como A/B.\n"
                                "Formato: 'Respostas: A, B, C...'\n"
                                "Se nenhuma estiver marcada, responda 'Vazia'."
                                "Assinale a letra que está claramente marcada com X, cruz ou rasura visível."
                            )
                            resposta_resposta = model.generate_content([prompt_resposta, img_resposta])
                            respostas_texto = resposta_resposta.text.strip()

                            # Corrigir prova
                            acertos = corrigir_prova(nome_texto, modelo_texto, respostas_texto, gabaritos)
                            
                            # Adicionar o resultado à lista
                            resultados_finais.append([nome_texto, modelo_texto, acertos])
                            
                            log_action(email, "PROVA_CORRIGIDA",
                                       f"Prova {provas_processadas} - Nome: {nome_texto}, Modelo: {modelo_texto}, Nota: {acertos}")
                                       
                        except Exception as e:
                            log_action(email, "ERRO_PAGINA", error=f"Erro ao processar prova na página {provas_processadas}: {str(e)}")
                            continue # Pula para a próxima página do chunk

                except Exception as e:
                    log_action(email, "ERRO_CHUNK", error=f"Erro ao processar chunk {i + 1}: {str(e)}")
                    continue # Pula para o próximo chunk

            # Criar workbook e escrever todos os resultados de uma vez
            wb = Workbook()
            ws = wb.active
            ws.title = "Resultados"
            ws.append(["Nome", "Modelo", "Nota"])

            for resultado in resultados_finais:
                ws.append(resultado)

            # Salvar arquivo de resultado
            resultado_path = os.path.join(temp_dir, 'resultado_provas.xlsx')
            wb.save(resultado_path)

            log_action(email, "PROCESSAMENTO_CONCLUIDO",
                       f"Total de provas processadas: {provas_processadas}, Arquivo gerado: resultado_provas.xlsx")
            

            # Retornar arquivo para download
            return send_file(
                resultado_path,
                as_attachment=True,
                download_name='resultado_provas.xlsx',
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

        finally:
            # Limpar arquivos temporários após um delay
            def cleanup():
                try:
                    shutil.rmtree(temp_dir)
                    log_action(email, "LIMPEZA_CONCLUIDA", "Arquivos temporários removidos")
                except Exception as e:
                    log_action(email, "ERRO_LIMPEZA", error=f"Erro ao remover arquivos temporários: {str(e)}")

            # Em produção, você pode usar um job scheduler para isso
            import threading
            timer = threading.Timer(CLEANUP_DELAY_SECONDS, cleanup)
            timer.start()

    except Exception as e:
        log_action(email, "ERRO_GERAL", error=f"Erro geral no processamento: {str(e)}")
        return jsonify({'erro': 'Erro interno do servidor'}), 500


@prova_bp.route('/status', methods=['GET'])
@cross_origin()
def status():
    """Retorna o status da API."""
    return jsonify({
        'status': 'online',
        'gemini_configurado': bool(GEMINI_API_KEY and model),
        'emails_autorizados_count': len(EMAILS_AUTORIZADOS),
        'timestamp': datetime.now().isoformat()
    })


@prova_bp.route('/logs', methods=['GET'])
@cross_origin()
def get_logs():
    """Retorna os últimos logs do sistema (apenas para administradores)."""
    try:
        if not os.path.exists(LOG_FILE):
            return jsonify({'logs': []})

        with open(LOG_FILE, 'r', encoding='utf-8') as f:
            lines = f.readlines()

        # Retornar apenas as últimas 100 linhas
        recent_logs = lines[-100:] if len(lines) > 100 else lines

        return jsonify({
            'logs': [line.strip() for line in recent_logs],
            'total_lines': len(lines)
        })

    except Exception as e:
        return jsonify({'erro': f'Erro ao ler logs: {str(e)}'}), 500
