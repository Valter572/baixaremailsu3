"""
Extrator Automatico - v36
=========================
Base: v35 (correcao do criterio "pasta sem documentos").
O que mudou na v36:
  1) CAUSA DO ERRO: no criterio automatico o programa caia de "pasta sem
     documentos" para "sem NF" e depois para "todos", mudando o significado
     da lista sem mudar o texto. Agora o automatico usa SOMENTE "pasta sem
     documentos" e so aceita "sem NF" quando a coluna P tem dados reais.
     Nunca cai para "todos".
  2) A leitura de pasta ignora o limite de profundidade e usa o prefixo de
     caminho longo do Windows, que era a causa silenciosa de pastas cheias
     aparecerem vazias (caminho acima de 260 caracteres).
  3) Se alguma subpasta nao puder ser lida, a pasta NAO e considerada vazia.
     O programa deixa de acusar falta de documento por engano.
  4) Arquivos de sistema (desktop.ini, thumbs.db) e temporarios de tamanho
     zero deixam de contar como documento.
  5) Novo botao "Abrir pasta" e duplo clique na lista: abre a pasta do
     contratado no Explorer e informa quantos arquivos o programa viu.
Colunas da liquidacao: A(1)=termo | B(2)=DRE | U(21)=nome+termo | O(15)=dia
                       P(16)=NF | Q(17)=valor
Colunas da aba Dados: D(4)=termo | R(18)=e-mail
"""
import tkinter as tk
from tkinter import filedialog, ttk
from pathlib import Path
import zipfile
import time
import shutil
import threading
import json
import os
import sys
import re
import queue
import unicodedata
import datetime
from PIL import Image, ImageSequence
import win32com.client
try:
    import pythoncom
except Exception:
    pythoncom = None

# =========================
# PONTE PARA AMBIENTE SEM CONSOLE
# =========================
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w", encoding="utf-8")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w", encoding="utf-8")

# =========================
# PASTA DE TRABALHO
# =========================
def _pasta_gravavel(base):
    try:
        base.mkdir(parents=True, exist_ok=True)
        teste = base / ".teste_escrita"
        teste.write_text("x", encoding="utf-8")
        teste.unlink()
        return True
    except Exception:
        return False

if getattr(sys, "frozen", False):
    _base_exe = Path(sys.executable).resolve().parent
else:
    _base_exe = Path(__file__).resolve().parent

if _pasta_gravavel(_base_exe):
    BASE_DIR = _base_exe
else:
    BASE_DIR = Path(os.environ.get("APPDATA") or str(Path.home())) / "ExtratorAutomatico"
    BASE_DIR.mkdir(parents=True, exist_ok=True)

CONFIG_PATH = BASE_DIR / "extrator_config.json"
LOG_PATH = BASE_DIR / "extrator_log.txt"

# =========================
# CONFIG
# =========================
SETOR_PADRAO = "DTP/TEG/Fin"
SAUDACAO_AUTO = "Automática"
OPCOES_SAUDACAO = [SAUDACAO_AUTO, "Bom dia", "Boa tarde", "Boa noite"]
OPCOES_CRITERIO = ("auto", "pasta", "nf", "todos")
NOME_CRITERIO = {
    "auto": "automático",
    "pasta": "pasta sem documentos",
    "nf": "sem NF na planilha",
    "todos": "todos da planilha",
}
TODAS_DRES = "(todas)"
CONTA_PADRAO = "(conta padrão do Outlook)"

ORIENTACAO_VARIOS = (
    "Este e-mail foi enviado a vários contratados. Localize na relação abaixo "
    "o termo contratual e o nome correspondentes ao seu contrato e também o valor."
)
COR_GRIFO = "#FFFF00"

CONFIG_PADRAO = {
    "pasta_principal": "",
    "pasta_download": "",
    "pasta_destino": "",
    "caminho_planilha": "",
    "nome_aba": "LIQUIDACAO",
    "col_termo": 1,
    "col_dre": 2,
    "col_alt": 21,
    "col_dia": 15,
    "col_nf": 16,
    "col_conferencia": 17,
    "profundidade_busca": 2,
    "salvar_apos_gravar": False,
    "acao": "mover",
    "intervalo_ms": 2000,
    "espera_limpeza_s": 2.5,
    "incluir_subpastas_download": True,
    "ignorar_arquivos_antigos": True,
    "departamento": SETOR_PADRAO,
    "assunto_cobranca": "Documentos pendentes – Termo Contratual",
    "criterio_cobranca": "auto",
    "saudacao_cobranca": SAUDACAO_AUTO,
    "assinar_setor": False,
    "remetente_email": "",
    "fonte_email": "excel",
    "nome_aba_email": "Dados",
    "col_email_termo": 4,
    "col_email_email": 18,
    "caminho_access": "",
    "mapa_access": {},
}

def carregar_config():
    cfg = dict(CONFIG_PADRAO)
    try:
        if CONFIG_PATH.exists():
            cfg.update(json.loads(CONFIG_PATH.read_text(encoding="utf-8")))
    except Exception:
        pass
    antigos = {
        "",
        "Departamento",
        "Departamento de Administracao",
        "Departamento de Administração",
    }
    if str(cfg.get("departamento", "")).strip() in antigos:
        cfg["departamento"] = SETOR_PADRAO
    if cfg.get("saudacao_cobranca") not in OPCOES_SAUDACAO:
        cfg["saudacao_cobranca"] = SAUDACAO_AUTO
    if not isinstance(cfg.get("mapa_access"), dict):
        cfg["mapa_access"] = {}
    if (cfg.get("fonte_email") or "").lower() not in ("excel", "access"):
        cfg["fonte_email"] = "excel"
    if cfg.get("criterio_cobranca") not in OPCOES_CRITERIO:
        cfg["criterio_cobranca"] = "auto"
    return cfg

cfg = carregar_config()

def salvar_config():
    try:
        CONFIG_PATH.write_text(
            json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except Exception:
        pass

EXT_IMAGEM = {".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".webp"}
EXT_IGNORAR = {".crdownload", ".part", ".tmp", ".download"}
TENTATIVAS_POR_CICLO = 3
MAX_CICLOS_Q = 8
INTERVALO_TENTATIVA = 0.35
INTERVALO_CICLO_MS = 2200
LIMITE_LINHAS_LEITURA = 200000
LIMITE_LINHAS_ACESS = 300000
LIMITE_LISTA_PENDENTES = 1500
TOLERANCIA_ANTIGOS_S = 0.5
INICIO_PROGRAMA = time.time()
LIMITE_LISTA_EMAIL = 30
DRE_VALIDADE_S = 120.0

# =========================
# ESTADO
# =========================
pasta_principal = None
pasta_download = None
pasta_destino_global = None
monitorando = False
movidos = 0
pastas_vazias = 0
candidatos_atuais = 0
recontar = False
ultimo_movimento = 0.0
snapshot_campos = None
cache_pastas = {}
indice_chaves = {}
resultados_atuais = []
estado = {}
fila = queue.Queue()
fila_q = queue.Queue()
geracao_q = 0
pasta_q_pendente = None
linha_atual = 0
chave_atual = ""
valor_q_atual = ""
timer = None
msg_label = None
acao_label = None
contrato_label = None
preview_label = None
confere_frame = None
confere_titulo = None
confere_label = None
confere_info = None
# --- DRE (coluna B) ---
cache_dre = {}
cache_dre_indice = {}
cache_dre_ts = 0.0
dre_em_leitura = False
# --- e-mails (Excel "Dados" ou Access) ---
cache_email_chave = {}
cache_email_seq = {}
cache_email_nome = {}
cache_termo_nome = {}
cache_email_info = ""
cache_email_diag = []
cache_email_amostra = []
ultimo_erro_email = ""
# --- liquidação (aba principal) ---
cache_base_info = ""
cache_base_diag = []
cache_base_amostra = []
# --- cobrança ---
janela_cob = None
lista_cob = None
texto_cob = None
status_cob = None
label_hora_cob = None
label_fonte_cob = None
var_dre_cob = None
combo_dre_cob = None
label_dre_cob = None
var_remetente_cob = None
combo_remetente_cob = None
label_remetente_cob = None
pendentes_cob = []
var_criterio_cob = None
entrada_setor_cob = None
entrada_assunto_cob = None
var_saudacao_cob = None
var_assinar_cob = None
entrada_dest_cob = None
# --- carga automática da lista ---
fila_cob = queue.Queue()
cob_carregando = False
cob_tentativas = 0
cob_poll_ativo = False
# --- pastas vazias ---
janela_vazias = None
lista_vazias = None
status_vazias = None
entrada_filtro_vazias = None
vazias_atuais = []
# --- mapeamento Access ---
janela_map = None
cb_map_tabela = None
cb_map_email = None
cb_map_chave = None
cb_map_nome = None
lbl_map_status = None

# =========================
# RETORNO NA TELA
# =========================
def log(msg):
    linha = f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {msg}"
    try:
        print(linha)
    except Exception:
        pass
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(linha + "\n")
    except Exception:
        pass
    try:
        janela.after(0, lambda: acao_label.config(text=msg[:110]))
    except Exception:
        pass

def aviso(texto, cor="#c62828"):
    try:
        janela.after(0, lambda: msg_label.config(text=texto, fg=cor))
    except Exception:
        pass

def msg_cob(texto, cor="#555555"):
    try:
        if status_cob is not None:
            status_cob.config(text=texto, fg=cor)
            return
    except Exception:
        pass
    aviso(texto, cor)

def painel(aguardando=0):
    txt = (f"Vazias: {pastas_vazias}   |   Movidos: {movidos}"
           f"   |   Em Downloads: {candidatos_atuais}"
           f"   |   Aguardando: {aguardando}"
           f"   |   Monitor: {'ON' if monitorando else 'OFF'}")
    try:
        janela.after(0, lambda: estado_label.config(text=txt))
    except Exception:
        pass

def resumir(nome, n=44):
    return nome if len(nome) <= n else nome[:n - 1] + "…"

def set_entry(entry, texto, manter_estado=True):
    try:
        st = entry.cget("state")
        entry.config(state="normal")
        entry.delete(0, tk.END)
        entry.insert(0, texto)
        entry.config(state=st if manter_estado else "normal")
    except Exception:
        pass

# =========================
# SAUDAÇÃO POR HORÁRIO
# =========================
def saudacao_por_hora(agora=None):
    """Bom dia 5h-11h59 | Boa tarde 12h-17h59 | Boa noite 18h-4h59."""
    h = (agora or datetime.datetime.now()).hour
    if 5 <= h < 12:
        return "bom dia"
    if 12 <= h < 18:
        return "boa tarde"
    return "boa noite"

def resolver_saudacao(escolha):
    e = (escolha or "").strip().lower()
    if e in ("bom dia", "boa tarde", "boa noite"):
        return e
    return saudacao_por_hora()

def saudacao_escolhida():
    try:
        return var_saudacao_cob.get()
    except Exception:
        return SAUDACAO_AUTO

def assinar_escolhido():
    try:
        return bool(var_assinar_cob.get())
    except Exception:
        return False

def dre_escolhida():
    """Devolve a DRE selecionada no combo, ou "" quando for todas."""
    try:
        valor = (var_dre_cob.get() or "").strip()
    except Exception:
        valor = TODAS_DRES
    if not valor or normalizar(valor) == normalizar(TODAS_DRES):
        return ""
    return valor

def remetente_escolhido():
    """Devolve a conta/caixa escolhida, ou "" para a conta padrão."""
    try:
        v = (var_remetente_cob.get() or "").strip()
    except Exception:
        return ""
    if not v or v.startswith("("):
        return ""
    return v

# =========================
# BLOCO DA COLUNA Q
# =========================
def mostrar_q(valor, extra=""):
    def _aplicar():
        try:
            if valor == "…":
                fundo, cor = "#e3f2fd", "#1565c0"
            elif valor and valor != "—":
                fundo, cor = "#fff8e1", "#b26a00"
            else:
                fundo, cor = "#f5f5f5", "#777777"
            confere_frame.config(bg=fundo, highlightbackground=cor)
            confere_titulo.config(bg=fundo, fg=cor)
            confere_label.config(text=valor or "—", fg=cor, bg=fundo)
            confere_info.config(text=extra, fg=cor, bg=fundo)
        except Exception:
            pass
    try:
        janela.after(0, _aplicar)
    except Exception:
        pass

def habilitar_campos(habilitado):
    estado_entry = "normal" if habilitado else "disabled"
    for e in (entrada_nf, entrada_dia):
        try:
            e.config(state="normal")
            e.delete(0, tk.END)
            e.config(state=estado_entry)
        except Exception:
            pass
    if habilitado:
        try:
            entrada_nf.focus_set()
        except Exception:
            pass

# =========================
# TEXTO / CONTRATO
# =========================
def texto_celula(v):
    if v is None:
        return ""
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else str(v)
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime("%Y/%m/%d")
    return str(v).strip()

def digitos(t):
    return re.sub(r"\D", "", "" if t is None else str(t))

_RE_CONTRATO = re.compile(r"(\d{4})\s*[-/._ ]\s*(\d[\d.\s]*)")

def chave_contrato(texto):
    t = texto_celula(texto)
    if not t:
        return ""
    for m in _RE_CONTRATO.finditer(t):
        num = digitos(m.group(2))
        if len(num) >= 4:
            return m.group(1) + num
    puro = digitos(t)
    if len(puro) >= 9 and puro == t.strip():
        return puro
    return ""

def termo_do_nome(nome):
    m = _RE_CONTRATO.search(str(nome))
    if not m:
        return ""
    return f"{m.group(1)}/{digitos(m.group(2))}"

def normalizar(t):
    if t is None:
        return ""
    t = unicodedata.normalize("NFKD", str(t))
    t = "".join(c for c in t if not unicodedata.combining(c))
    return t.lower().strip()

def pasta_destino_atual():
    if pasta_destino_global and pasta_destino_global.is_dir():
        return pasta_destino_global
    if cfg.get("pasta_destino") and Path(cfg["pasta_destino"]).is_dir():
        return Path(cfg["pasta_destino"])
    return None

# ---------- CHAVES TOLERANTES DE TERMO ----------
def _seq_limpa(seq):
    d = digitos(seq)
    return d.lstrip("0") or d

def chaves_exatas(texto):
    t = texto_celula(texto)
    out = set()
    if not t:
        return out
    for m in _RE_CONTRATO.finditer(t):
        ano, seq = m.group(1), m.group(2)
        dseq = digitos(seq)
        if len(dseq) >= 3:
            out.add(ano + _seq_limpa(dseq))
    d = digitos(t)
    if len(d) >= 9:
        out.add(d)
        if len(d) >= 12:
            out.add(d[:4] + _seq_limpa(d[4:]))
    return out

def chaves_seq(texto):
    t = texto_celula(texto)
    out = set()
    if not t:
        return out
    for m in _RE_CONTRATO.finditer(t):
        s = _seq_limpa(m.group(2))
        if len(s) >= 4:
            out.add(s)
    d = digitos(t)
    if 4 <= len(d) <= 10:
        out.add(_seq_limpa(d))
    return out

def nome_base(nome):
    t = texto_celula(nome)
    t = _RE_CONTRATO.sub(" ", t)
    t = re.sub(r"[\-_/]{2,}", " ", t)
    t = t.strip(" -–—/")
    return normalizar(t)

# =========================
# EXCEL
# =========================
def abas_do_wb(wb):
    nomes = []
    try:
        for sh in wb.Worksheets:
            try:
                nomes.append(str(sh.Name))
            except Exception:
                pass
    except Exception:
        pass
    return nomes

def _listar_sheets(wb):
    abas = []
    try:
        for sh in wb.Worksheets:
            abas.append(sh)
    except Exception:
        pass
    return abas

def obter_aba(wb):
    """
    Aba da liquidação, tolerante a acento e caixa.
    NUNCA devolve a aba usada para os e-mails.
    Devolve None quando não encontra.
    """
    alvo = normalizar(cfg.get("nome_aba") or "LIQUIDACAO")
    evitar = normalizar(cfg.get("nome_aba_email") or "")
    abas = _listar_sheets(wb)
    if not abas:
        return None

    for sh in abas:
        try:
            n = normalizar(sh.Name)
        except Exception:
            continue
        if n and n == alvo and n != evitar:
            return sh

    for sh in abas:
        try:
            n = normalizar(sh.Name)
        except Exception:
            continue
        if alvo and alvo in n and n != evitar:
            return sh

    for sh in abas:
        try:
            n = normalizar(sh.Name)
        except Exception:
            continue
        if n and n in alvo and n != evitar:
            return sh

    return None

def obter_aba_por_nome(wb, nome):
    alvo = normalizar(nome)
    abas = _listar_sheets(wb)
    for sh in abas:
        try:
            if normalizar(sh.Name) == alvo:
                return sh
        except Exception:
            continue
    for sh in abas:
        try:
            if alvo and alvo in normalizar(sh.Name):
                return sh
        except Exception:
            continue
    return None

def obter_wb(excel):
    alvo = Path(cfg["caminho_planilha"])
    for w in excel.Workbooks:
        try:
            if Path(str(w.FullName)) == alvo:
                return w
        except Exception:
            continue
    for w in excel.Workbooks:
        try:
            if Path(str(w.FullName)).name.lower() == alvo.name.lower():
                return w
        except Exception:
            continue
    if not alvo.exists():
        raise FileNotFoundError(str(alvo))
    return excel.Workbooks.Open(str(alvo))

def esperar_excel_pronto(excel, limite_s=1.2):
    fim = time.time() + limite_s
    while time.time() < fim:
        try:
            if excel.Ready:
                return True
        except Exception:
            pass
        time.sleep(0.1)
    return False

def normalizar_matriz(v):
    if v is None:
        return []
    if not isinstance(v, tuple):
        return [[v]]
    if len(v) == 0:
        return []
    if not isinstance(v[0], tuple):
        return [list(v)]
    return [list(r) if isinstance(r, tuple) else [r] for r in v]

def colunas_relevantes():
    cols = []
    for chave in ("col_termo", "col_dre", "col_alt", "col_conferencia",
                  "col_dia", "col_nf"):
        try:
            c = int(cfg.get(chave) or 0)
        except Exception:
            c = 0
        if c > 0:
            cols.append(c)
    return cols or [1]

def localizar_por_bloco(ws, chave):
    if not chave:
        return 0, "contrato vazio"
    try:
        ur = ws.UsedRange
        ult = int(ur.Row) + int(ur.Rows.Count) - 1
    except Exception:
        ult = 0
    if ult <= 0:
        try:
            ult = int(ws.UsedRange.SpecialCells(11).Row)
        except Exception:
            ult = 1
    ult = min(max(ult, 1), LIMITE_LINHAS_LEITURA)
    col_max = max(colunas_relevantes())
    try:
        bruto = ws.Range(ws.Cells(1, 1), ws.Cells(ult, col_max)).Value
    except Exception as e:
        return 0, f"leitura falhou ({e})"
    linhas = normalizar_matriz(bruto)
    if not linhas:
        return 0, "planilha vazia"
    ca = int(cfg["col_termo"]) - 1
    cu = int(cfg.get("col_alt") or 0) - 1
    for i, linha in enumerate(linhas):
        if 0 <= ca < len(linha) and chave_contrato(linha[ca]) == chave:
            return i + 1, "coluna A"
    if cu >= 0:
        achados = [i + 1 for i, linha in enumerate(linhas)
                   if cu < len(linha) and chave_contrato(linha[cu]) == chave]
        if len(achados) == 1:
            return achados[0], "coluna U"
        if len(achados) > 1:
            return 0, f"{len(achados)} linhas na coluna U com o mesmo contrato"
    parecidos = []
    for i, linha in enumerate(linhas):
        for idx in (ca, cu):
            if idx < 0 or idx >= len(linha):
                continue
            d = digitos(linha[idx])
            if len(d) >= 9 and (chave in d or d in chave):
                parecidos.append(i + 1)
                break
    parecidos = sorted(set(parecidos))
    if len(parecidos) == 1:
        return parecidos[0], "aproximado"
    if len(parecidos) > 1:
        return 0, f"{len(parecidos)} linhas parecidas"
    return 0, "contrato não encontrado"

def ler_valor_q(ws, linha):
    col_q = int(cfg.get("col_conferencia", 17))
    valor = ""
    try:
        valor = texto_celula(ws.Cells(linha, col_q).Value)
    except Exception:
        valor = ""
    if not valor:
        try:
            valor = str(ws.Cells(linha, col_q).Text or "").strip()
        except Exception:
            valor = ""
    return valor

# =========================
# E-MAILS — ABA "DADOS" DA PLANILHA
# =========================
_RE_EMAIL = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")

ALVOS_EMAIL = ("email", "e-mail", "e_mail", "mail", "correio", "eletronico")
ALVOS_CHAVE = ("termo_de_adesao", "termodeadesao", "termo de adesao", "termo",
               "adesao", "contrato", "instrumento", "processo", "empenho",
               "protocolo", "numero")
ALVOS_NOME = ("razao_social", "razaosocial", "razao social", "contratado",
              "fornecedor", "credor", "fantasia", "empresa", "nome")

def emails_validos(bruto):
    t = texto_celula(bruto)
    if not t:
        return []
    t = re.sub(r"(?i)mailto:", "", t)
    achados = []
    for m in _RE_EMAIL.finditer(t):
        e = m.group(0).strip(" .,;:#").lower()
        if e and e not in achados:
            achados.append(e)
    return achados

def letra_para_numero(txt):
    t = texto_celula(txt).upper()
    if not t:
        return 0
    if t.isdigit():
        n = int(t)
        return n if n > 0 else 0
    n = 0
    for c in t:
        if "A" <= c <= "Z":
            n = n * 26 + (ord(c) - 64)
        else:
            return 0
    return n

def numero_para_letra(n):
    try:
        n = int(n)
    except Exception:
        return "?"
    if n <= 0:
        return "?"
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

def ler_emails_excel():
    global cache_email_chave, cache_email_seq, cache_email_nome
    global cache_termo_nome, cache_email_info, cache_email_diag, cache_email_amostra

    if not cfg.get("caminho_planilha"):
        raise RuntimeError("planilha não mapeada — clique em 'Planilha' na tela principal")

    aba_nome = (cfg.get("nome_aba_email") or "Dados").strip()
    col_t = int(cfg.get("col_email_termo") or 4)
    col_e = int(cfg.get("col_email_email") or 18)

    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        raise RuntimeError("o Excel não está aberto. Abra o arquivo Liquidacao.")

    esperar_excel_pronto(excel)
    try:
        wb = obter_wb(excel)
    except Exception as e:
        raise RuntimeError(f"não consegui localizar a planilha ({e})")

    ws = obter_aba_por_nome(wb, aba_nome)
    nomes_abas = abas_do_wb(wb)
    if ws is None:
        raise RuntimeError(f'a aba "{aba_nome}" não existe em {wb.Name}. '
                           f'Abas disponíveis: {", ".join(nomes_abas)}')

    try:
        ur = ws.UsedRange
        ult = int(ur.Row) + int(ur.Rows.Count) - 1
    except Exception:
        ult = 1
    ult = min(max(ult, 1), LIMITE_LINHAS_LEITURA)
    col_max = max(col_t, col_e)
    try:
        bruto = ws.Range(ws.Cells(1, 1), ws.Cells(ult, col_max)).Value
    except Exception as e:
        raise RuntimeError(f"erro ao ler a aba {aba_nome} ({e})")

    linhas = normalizar_matriz(bruto)
    it, ie = col_t - 1, col_e - 1

    chave_map, seq_map = {}, {}
    amostra = []
    total = com_email = sem_email = sem_termo = 0

    for linha in linhas:
        t = texto_celula(linha[it]) if 0 <= it < len(linha) else ""
        bruto_email = linha[ie] if 0 <= ie < len(linha) else ""
        if not t and not texto_celula(bruto_email):
            continue
        total += 1
        emails = emails_validos(bruto_email)
        if not emails:
            sem_email += 1
            continue
        com_email += 1
        exatas = chaves_exatas(t)
        seqs = chaves_seq(t)
        if not exatas and not seqs:
            sem_termo += 1
        for k in exatas:
            chave_map.setdefault(k, set()).update(emails)
        for k in seqs:
            seq_map.setdefault(k, set()).update(emails)
        if len(amostra) < 8:
            amostra.append((t or "(termo vazio)", "; ".join(emails)))

    cache_email_chave = chave_map
    cache_email_seq = seq_map
    cache_email_nome = {}
    cache_termo_nome = {}
    cache_email_amostra = amostra
    cache_email_info = (f'{wb.Name} · aba "{aba_nome}" · '
                        f'{numero_para_letra(col_t)}/{numero_para_letra(col_e)} · '
                        f'{com_email}/{total} com e-mail')
    cache_email_diag = [
        f'Fonte: Excel — aba "{aba_nome}" de {wb.Name}',
        f'Termo: coluna {numero_para_letra(col_t)} ({col_t})',
        f'E-mail: coluna {numero_para_letra(col_e)} ({col_e})',
        f'Linhas com conteúdo: {total}',
        f'Com e-mail válido: {com_email}',
        f'Sem e-mail na coluna: {sem_email}',
        f'Com e-mail mas sem termo reconhecido: {sem_termo}',
        f'Índices: {len(chave_map)} termo(s) e {len(seq_map)} sequência(s)',
    ]

    if not chave_map and not seq_map:
        raise RuntimeError(
            f'a aba "{aba_nome}" foi lida ({total} linha(s)), mas nenhum par '
            f'termo/e-mail válido apareceu nas colunas '
            f'{numero_para_letra(col_t)} e {numero_para_letra(col_e)}. '
            f'Confira as colunas em "Fonte".'
        )

    log(f"E-mails: {cache_email_info}")
    return cache_email_info

# =========================
# E-MAILS — ACCESS (ALTERNATIVA)
# =========================
def _abrir_access(caminho):
    erros = []
    for provider in ("Microsoft.ACE.OLEDB.12.0", "Microsoft.Jet.OLEDB.4.0"):
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            try:
                conn.Mode = 1
            except Exception:
                pass
            conn.Open(f"Provider={provider};Data Source={caminho};"
                      f"Persist Security Info=False;")
            return conn, provider, ""
        except Exception as e:
            erros.append(f"{provider}: {str(e)[:180]}")
    return None, "", " | ".join(erros)

def _tabelas_access(conn):
    tabelas, consultas = [], []
    try:
        rs = conn.OpenSchema(20)
        while not rs.EOF:
            try:
                tipo = str(rs.Fields.Item("TABLE_TYPE").Value).upper()
                nome = str(rs.Fields.Item("TABLE_NAME").Value)
            except Exception:
                tipo, nome = "", ""
            if nome:
                if "TABLE" in tipo and not nome.upper().startswith("MSYS"):
                    tabelas.append(nome)
                elif "VIEW" in tipo and not nome.upper().startswith("MSYS"):
                    consultas.append(nome)
            rs.MoveNext()
        rs.Close()
    except Exception as e:
        log(f"Access: falha ao listar tabelas ({e})")
    return tabelas + consultas

def _colunas_schema(conn, tabela=None):
    saida = []
    try:
        if tabela:
            rs = conn.OpenSchema(4, (None, None, str(tabela), None))
        else:
            rs = conn.OpenSchema(4)
        while not rs.EOF:
            try:
                nome_tab = str(rs.Fields.Item("TABLE_NAME").Value)
                nome_col = str(rs.Fields.Item("COLUMN_NAME").Value)
            except Exception:
                nome_tab, nome_col = "", ""
            if nome_col:
                if tabela is None:
                    saida.append((nome_tab, nome_col))
                elif nome_tab.lower() == str(tabela).lower():
                    saida.append(nome_col)
            rs.MoveNext()
        rs.Close()
    except Exception as e:
        log(f"Access: OpenSchema(4) falhou para {tabela} ({e})")
    return saida

def _campos_da_tabela(conn, tabela):
    campos = _colunas_schema(conn, tabela)
    if campos:
        return campos
    try:
        rs = conn.Execute(f"SELECT TOP 1 * FROM [{tabela}]")
        for i in range(rs.Fields.Count):
            campos.append(str(rs.Fields.Item(i).Name))
        rs.Close()
        if campos:
            return campos
    except Exception as e:
        log(f"Access: TOP 1 falhou em {tabela} ({str(e)[:120]})")
    try:
        rs = conn.Execute(f"SELECT * FROM [{tabela}] WHERE 1=0")
        for i in range(rs.Fields.Count):
            campos.append(str(rs.Fields.Item(i).Name))
        rs.Close()
    except Exception as e:
        log(f"Access: WHERE 1=0 falhou em {tabela} ({str(e)[:120]})")
    return campos

def _contar_preenchidos(conn, tabela, campo):
    try:
        rs = conn.Execute(f"SELECT COUNT(*) FROM [{tabela}] "
                          f"WHERE [{campo}] IS NOT NULL")
        n = -1
        if not rs.EOF:
            n = int(rs.Fields.Item(0).Value or 0)
        rs.Close()
        return n
    except Exception:
        return -1

def _score_preenchido(n):
    if n > 0:
        return n
    if n == -1:
        return 1
    return 0

def _candidatos_campo(campos, alvos):
    achados = []
    for alvo in alvos:
        for nome in campos:
            if alvo in normalizar(nome) and nome not in achados:
                achados.append(nome)
    return achados

def _melhor_campo(conn, tabela, campos, alvos):
    cands = _candidatos_campo(campos, alvos)
    if not cands:
        return ""
    melhor, melhor_score = cands[0], -1
    for c in cands:
        s = _score_preenchido(_contar_preenchidos(conn, tabela, c))
        if s > melhor_score:
            melhor, melhor_score = c, s
    return melhor

def _campos_email_tabela(conn, tabela, campos, limite=3):
    cands = _candidatos_campo(campos, ALVOS_EMAIL)
    pares = []
    for c in cands:
        n = _contar_preenchidos(conn, tabela, c)
        if n == 0:
            continue
        pares.append((_score_preenchido(n), c))
    pares.sort(key=lambda par: par[0], reverse=True)
    return [c for _, c in pares[:limite]]

def _escapar(campo):
    return "[" + str(campo).replace("]", "]]") + "]"

def ler_access():
    global cache_email_chave, cache_email_seq, cache_email_nome
    global cache_termo_nome, cache_email_info, cache_email_diag, cache_email_amostra

    caminho = cfg.get("caminho_access", "")
    if not caminho:
        raise RuntimeError("banco Access não selecionado")
    if not Path(caminho).exists():
        raise RuntimeError("arquivo Access não encontrado no caminho salvo")

    conn, provider, erro = _abrir_access(caminho)
    if conn is None:
        raise RuntimeError(
            "não consegui abrir o Access. Instale o Microsoft Access Database "
            "Engine compatível com a arquitetura do Python "
            f"({(sys.maxsize > 2**32 and '64') or '32'} bits). Detalhe: {erro}"
        )

    chave_map, seq_map, nome_map, termo_nome = {}, {}, {}, {}
    diag = [f"Provedor: {provider}"]

    try:
        tabelas = _tabelas_access(conn)
        diag.append(f"Tabelas/consultas: {len(tabelas)}")
        if not tabelas:
            raise RuntimeError("nenhuma tabela ou consulta no banco")

        for t in tabelas:
            campos = _campos_da_tabela(conn, t)
            if not campos:
                diag.append(f"{t}: CAMPOS NAO LIDOS")
                continue

            f_emails = _campos_email_tabela(conn, t, campos)
            f_chave = _melhor_campo(conn, t, campos, ALVOS_CHAVE)
            f_nome = _melhor_campo(conn, t, campos, ALVOS_NOME)

            ov = (cfg.get("mapa_access") or {}).get(t) or {}
            if ov.get("email") and ov["email"] in campos:
                f_emails = [ov["email"]]
            if ov.get("chave") and ov["chave"] in campos:
                f_chave = ov["chave"]
            if ov.get("nome") and ov["nome"] in campos:
                f_nome = ov["nome"]

            if not f_emails and not f_chave and not f_nome:
                diag.append(f"{t}: {len(campos)} campos · sem campo útil "
                            f"(campos: {', '.join(campos[:8])})")
                continue

            sel_cols = []
            for c in list(f_emails or []) + [f_chave, f_nome]:
                if c and c not in sel_cols:
                    sel_cols.append(c)
            if not sel_cols:
                continue

            sql = "SELECT " + ", ".join(_escapar(c) for c in sel_cols) + f" FROM [{t}]"
            try:
                rs = conn.Execute(sql)
            except Exception as e:
                diag.append(f"{t}: erro ao ler ({str(e)[:120]})")
                continue

            idx = {c: i for i, c in enumerate(sel_cols)}

            def valor_col(registro, campo):
                if not campo:
                    return ""
                i = idx.get(campo)
                if i is None:
                    return ""
                try:
                    return registro.Fields.Item(i).Value
                except Exception:
                    return ""

            total = com_email = com_chave = 0
            while not rs.EOF and total < LIMITE_LINHAS_ACESS:
                total += 1
                try:
                    emails = []
                    for c in (f_emails or []):
                        for e in emails_validos(valor_col(rs, c)):
                            if e not in emails:
                                emails.append(e)
                    chaves_e, chaves_s = set(), set()
                    if f_chave:
                        bruto_chave = valor_col(rs, f_chave)
                        chaves_e = chaves_exatas(bruto_chave)
                        chaves_s = chaves_seq(bruto_chave)
                    nm = nome_base(valor_col(rs, f_nome)) if f_nome else ""

                    if chaves_e or chaves_s:
                        com_chave += 1
                    if emails:
                        com_email += 1
                        for k in chaves_e:
                            chave_map.setdefault(k, set()).update(emails)
                        for k in chaves_s:
                            seq_map.setdefault(k, set()).update(emails)
                        if nm:
                            nome_map.setdefault(nm, set()).update(emails)
                    if nm and (chaves_e or chaves_s):
                        for k in (chaves_e | chaves_s):
                            termo_nome.setdefault(k, set()).add(nm)
                except Exception:
                    pass
                rs.MoveNext()
            try:
                rs.Close()
            except Exception:
                pass

            diag.append(f"{t}: {total} linhas · {com_email} com e-mail · "
                        f"{com_chave} com termo · email={f_emails or '-'} · "
                        f"termo={f_chave or '-'} · nome={f_nome or '-'}")
    finally:
        try:
            conn.Close()
        except Exception:
            pass

    cache_email_chave = chave_map
    cache_email_seq = seq_map
    cache_email_nome = nome_map
    cache_termo_nome = termo_nome
    cache_email_amostra = []

    if not chave_map and not nome_map and not seq_map:
        cache_email_diag = diag + ["", "Nenhum e-mail indexado no Access. "
                                       "Use a fonte Excel (aba Dados)."]
        raise RuntimeError("o banco abriu, mas nenhum e-mail foi encontrado")

    cache_email_info = (f"{Path(caminho).name} · {len(chave_map)} termos · "
                        f"{len(nome_map)} nomes · {len(seq_map)} sequências")
    diag.append("")
    diag.append(f"Índices: {len(chave_map)} termo(s), {len(seq_map)} sequência(s), "
                f"{len(nome_map)} nome(s)")
    cache_email_diag = diag
    log(f"Access: {cache_email_info}")
    return cache_email_info

def procurar_termo_no_banco(termo):
    caminho = cfg.get("caminho_access", "")
    alvo = texto_celula(termo)
    if not caminho:
        raise RuntimeError("banco Access não selecionado")
    if not alvo:
        raise RuntimeError("digite um termo para procurar")
    conn, provider, erro = _abrir_access(caminho)
    if conn is None:
        raise RuntimeError(erro)
    like = alvo.replace("'", "''")
    achados = []
    try:
        for t in _tabelas_access(conn):
            campos = _campos_da_tabela(conn, t)
            if not campos:
                continue
            cols = campos[:40]
            cond = " OR ".join(f"[{c}] LIKE '%{like}%'" for c in cols)
            sql = ("SELECT TOP 5 " + ", ".join(_escapar(c) for c in cols)
                   + f" FROM [{t}] WHERE " + cond)
            try:
                rs = conn.Execute(sql)
            except Exception:
                continue
            while not rs.EOF and len(achados) < 40:
                pares = []
                for i, c in enumerate(cols):
                    try:
                        v = rs.Fields.Item(i).Value
                    except Exception:
                        v = ""
                    if v is None:
                        continue
                    txt = texto_celula(v)
                    if txt:
                        pares.append(f"{c}={txt!r}")
                if pares:
                    achados.append((t, " | ".join(pares)))
                rs.MoveNext()
            try:
                rs.Close()
            except Exception:
                pass
    finally:
        try:
            conn.Close()
        except Exception:
            pass
    return achados

# =========================
# CARGA DE E-MAILS
# =========================
def carregar_emails_core():
    global ultimo_erro_email, cache_email_info
    fonte = (cfg.get("fonte_email") or "excel").lower()
    try:
        if fonte == "access":
            info = ler_access()
        else:
            info = ler_emails_excel()
        ultimo_erro_email = ""
        return True, info
    except Exception as e:
        ultimo_erro_email = str(e)
        cache_email_chave.clear()
        cache_email_seq.clear()
        cache_email_nome.clear()
        cache_termo_nome.clear()
        cache_email_amostra.clear()
        cache_email_info = ""
        log(f"E-mails: erro ({e})")
        return False, str(e)

def atualizar_label_fonte():
    if label_fonte_cob is None:
        return
    try:
        if (cfg.get("fonte_email") or "excel").lower() == "access":
            nome = ""
            if cfg.get("caminho_access"):
                nome = Path(cfg["caminho_access"]).name
            texto = f"Access · {nome or '(banco não escolhido)'}"
            cor = "#2e7d32" if nome else "#b26a00"
        else:
            texto = (f'Excel · aba "{cfg.get("nome_aba_email", "Dados")}" · '
                     f'termo {numero_para_letra(cfg.get("col_email_termo", 4))} · '
                     f'e-mail {numero_para_letra(cfg.get("col_email_email", 18))}')
            cor = "#2e7d32"
        if ultimo_erro_email:
            label_fonte_cob.config(text=f"falha: {ultimo_erro_email[:100]}",
                                   fg="#c62828")
        else:
            label_fonte_cob.config(text=texto, fg=cor)
    except Exception:
        pass

def carregar_emails(avisar=True):
    ok, info = carregar_emails_core()
    atualizar_label_fonte()
    if avisar:
        if ok:
            msg_cob(f"E-mails carregados: {info}", "#2e7d32")
        else:
            msg_cob(f"E-mails: {info}", "#c62828")
    return ok

def buscar_emails(registro):
    fontes = [registro.get("termo") or "", registro.get("chave") or "",
              registro.get("nome") or ""]

    exatas, seqs = set(), set()
    for t in fontes:
        exatas |= chaves_exatas(t)
        seqs |= chaves_seq(t)

    for k in exatas:
        if k in cache_email_chave:
            return sorted(cache_email_chave[k])

    for k in seqs:
        if k in cache_email_seq:
            lista = cache_email_seq[k]
            if 0 < len(lista) <= 3:
                return sorted(lista)

    nomes = set()
    for k in (exatas | seqs):
        nomes |= cache_termo_nome.get(k, set())
    for nm in nomes:
        if nm in cache_email_nome:
            return sorted(cache_email_nome[nm])

    nm = nome_base(registro.get("nome") or "")
    if nm and nm in cache_email_nome:
        return sorted(cache_email_nome[nm])
    if nm and len(nm) >= 10:
        for k, v in cache_email_nome.items():
            if len(k) >= 10 and (nm in k or k in nm) and len(v) <= 3:
                return sorted(v)

    for k in seqs:
        if len(k) < 5:
            continue
        for ref, v in cache_email_seq.items():
            if len(ref) < 5:
                continue
            if ref.endswith(k) or k.endswith(ref):
                if 0 < len(v) <= 3:
                    return sorted(v)
    return []

def tracar_vinculo(registro):
    linhas = []
    termo = registro.get("termo") or ""
    nome = registro.get("nome") or ""
    chave = registro.get("chave") or ""
    linhas.append(f"DRE        : {registro.get('dre') or '(vazia)'}")
    linhas.append(f"Contratado : {nome}")
    linhas.append(f"Termo (A)  : {termo}")
    linhas.append(f"Valor (Q)  : {registro.get('valor') or '(vazio)'}")

    exatas, seqs = set(), set()
    for t in (termo, chave, nome):
        exatas |= chaves_exatas(t)
        seqs |= chaves_seq(t)

    linhas.append(f"Chaves exatas : {sorted(exatas) or '(nenhuma)'}")
    linhas.append(f"Sequências    : {sorted(seqs) or '(nenhuma)'}")

    achou = []
    for k in sorted(exatas):
        if k in cache_email_chave:
            achou.append(f"nível 1 (chave exata {k}) -> {sorted(cache_email_chave[k])}")
    for k in sorted(seqs):
        if k in cache_email_seq:
            achou.append(f"nível 2 (sequência {k}) -> {sorted(cache_email_seq[k])}")

    nm = nome_base(nome)
    if nm and nm in cache_email_nome:
        achou.append(f"nível 3 (nome {nm}) -> {sorted(cache_email_nome[nm])}")

    if achou:
        linhas.extend("  OK: " + a for a in achou)
    else:
        linhas.append("  NADA: nenhum vínculo encontrado para este contratado.")
        if cache_email_chave:
            linhas.append(f"  Amostra de termos indexados: "
                          f"{list(cache_email_chave.keys())[:8]}")
        if cache_email_amostra:
            linhas.append("  Amostra lida da aba: " +
                          "; ".join(f"{t} -> {e}" for t, e in cache_email_amostra[:3]))
    return linhas

# =========================
# OUTLOOK — CONTAS E ENVIO
# =========================
def listar_contas_outlook():
    """Contas configuradas no perfil do Outlook."""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        ns = outlook.GetNamespace("MAPI")
        contas = []
        for i in range(1, ns.Accounts.Count + 1):
            try:
                acc = ns.Accounts.Item(i)
                contas.append({
                    "nome": str(acc.DisplayName),
                    "smtp": str(acc.SmtpAddress or ""),
                })
            except Exception:
                continue
        return contas
    except Exception as e:
        log(f"Outlook: não consegui listar as contas ({e})")
        return []

def _achar_conta(ns, remetente):
    """Conta do perfil que casa com o que o usuário escolheu."""
    alvo = normalizar(remetente)
    if not alvo:
        return None
    for i in range(1, ns.Accounts.Count + 1):
        try:
            acc = ns.Accounts.Item(i)
            if (normalizar(acc.SmtpAddress) == alvo
                    or normalizar(acc.DisplayName) == alvo):
                return acc
        except Exception:
            continue
    return None

def _pasta_da_caixa(ns, remetente):
    """Pasta Caixa de Entrada da caixa compartilhada, se estiver no perfil."""
    alvo = normalizar(remetente)
    if not alvo:
        return None
    try:
        for i in range(1, ns.Folders.Count + 1):
            try:
                raiz = ns.Folders.Item(i)
                if normalizar(raiz.Name) != alvo:
                    continue
                try:
                    return raiz.Folders["Caixa de Entrada"]
                except Exception:
                    try:
                        return raiz.Folders["Inbox"]
                    except Exception:
                        return raiz
            except Exception:
                continue
    except Exception:
        pass
    return None

def criar_rascunho_outlook(assunto, corpo, destinatarios=None, html=None,
                           remetente="", modo="auto"):
    """
    Cria o rascunho. Devolve (ok, detalhe).
    modo:
      auto           tenta conta do perfil, depois caixa compartilhada, depois padrão
      conta          força SendUsingAccount
      compartilhada  força criação do item na pasta da caixa
      em_nome        força SentOnBehalfOfName (delegação)
    """
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        ns = outlook.GetNamespace("MAPI")
        remetente = (remetente or "").strip()
        mail = None
        usado = "conta padrão do Outlook"

        if remetente and modo in ("auto", "compartilhada"):
            pasta = _pasta_da_caixa(ns, remetente)
            if pasta is not None:
                try:
                    mail = pasta.Items.Add("IPM.Note")
                    usado = f"caixa compartilhada {remetente}"
                except Exception:
                    mail = None

        if mail is None and remetente and modo in ("auto", "conta"):
            acc = _achar_conta(ns, remetente)
            if acc is not None:
                mail = outlook.CreateItem(0)
                try:
                    mail.SendUsingAccount = acc
                    usado = f"conta {acc.SmtpAddress or acc.DisplayName}"
                except Exception:
                    usado = "conta padrão do Outlook"

        if mail is None and remetente and modo == "em_nome":
            mail = outlook.CreateItem(0)
            try:
                mail.SentOnBehalfOfName = remetente
                usado = f"em nome de {remetente}"
            except Exception:
                usado = "conta padrão do Outlook"

        if mail is None:
            mail = outlook.CreateItem(0)
            if remetente and modo == "auto":
                usado = f"conta padrão (não achei {remetente})"
            else:
                usado = "conta padrão do Outlook"

        mail.Subject = assunto
        if html:
            try:
                mail.HTMLBody = html
            except Exception:
                mail.Body = corpo
        else:
            mail.Body = corpo
        if destinatarios:
            try:
                mail.To = "; ".join(destinatarios)
            except Exception:
                pass
        mail.Display()
        return True, usado
    except Exception as e:
        return False, str(e)

def atualizar_combo_remetente(avisar=False):
    """Preenche a lista de contas e mantém a escolha salva."""
    if combo_remetente_cob is None or var_remetente_cob is None:
        return
    contas = listar_contas_outlook()
    valores = [CONTA_PADRAO]
    for c in contas:
        v = c["smtp"] or c["nome"]
        if v and v not in valores:
            valores.append(v)
    try:
        combo_remetente_cob.config(values=valores)
        salvo = (cfg.get("remetente_email") or "").strip()
        if salvo:
            var_remetente_cob.set(salvo)
        else:
            var_remetente_cob.set(CONTA_PADRAO)
        if label_remetente_cob is not None:
            if contas:
                label_remetente_cob.config(
                    text=f"{len(contas)} conta(s) no perfil", fg="#2e7d32")
            else:
                label_remetente_cob.config(
                    text="nenhuma conta lida do Outlook", fg="#b26a00")
    except Exception:
        pass
    if avisar:
        if contas:
            msg_cob("Contas lidas do Outlook: "
                    + ", ".join(c["smtp"] or c["nome"] for c in contas), "#2e7d32")
        else:
            msg_cob("Não consegui ler as contas do Outlook. "
                    "Digite o endereço da caixa manualmente.", "#b26a00")

# =========================
# JANELAS DE CONFIGURAÇÃO
# =========================
def abrir_config_emails():
    jan = tk.Toplevel(janela_cob or janela)
    jan.title("Fontes: e-mails e liquidação")
    jan.geometry("560x560")
    jan.transient(janela_cob or janela)

    tk.Label(jan, text="De onde vêm os e-mails:",
             font=("Arial", 9, "bold"), anchor="w").pack(
        fill="x", padx=12, pady=(10, 2))

    var_fonte = tk.StringVar(
        value="Access (banco)" if (cfg.get("fonte_email") or "").lower() == "access"
        else "Excel (aba e colunas)")

    tk.Radiobutton(jan, text="Excel — aba da própria planilha", font=("Arial", 9),
                   variable=var_fonte,
                   value="Excel (aba e colunas)").pack(anchor="w", padx=12)
    tk.Radiobutton(jan, text="Access — banco de dados", font=("Arial", 9),
                   variable=var_fonte,
                   value="Access (banco)").pack(anchor="w", padx=12)

    quadro = tk.Frame(jan)
    quadro.pack(fill="x", padx=12, pady=(10, 0))

    tk.Label(quadro, text="Aba dos e-mails", font=("Arial", 8, "bold"),
             anchor="w").grid(row=0, column=0, sticky="w", pady=3)
    e_aba = tk.Entry(quadro, font=("Arial", 10), width=20)
    e_aba.insert(0, cfg.get("nome_aba_email", "Dados"))
    e_aba.grid(row=0, column=1, sticky="w", padx=8)

    tk.Label(quadro, text="Coluna do termo", font=("Arial", 8, "bold"),
             anchor="w").grid(row=1, column=0, sticky="w", pady=3)
    e_col_t = tk.Entry(quadro, font=("Arial", 10), width=8)
    e_col_t.insert(0, numero_para_letra(cfg.get("col_email_termo", 4)))
    e_col_t.grid(row=1, column=1, sticky="w", padx=8)
    tk.Label(quadro, text='letra ("D") ou número ("4")', font=("Arial", 7),
             fg="#777777").grid(row=1, column=2, sticky="w")

    tk.Label(quadro, text="Coluna do e-mail", font=("Arial", 8, "bold"),
             anchor="w").grid(row=2, column=0, sticky="w", pady=3)
    e_col_e = tk.Entry(quadro, font=("Arial", 10), width=8)
    e_col_e.insert(0, numero_para_letra(cfg.get("col_email_email", 18)))
    e_col_e.grid(row=2, column=1, sticky="w", padx=8)
    tk.Label(quadro, text='letra ("R") ou número ("18")', font=("Arial", 7),
             fg="#777777").grid(row=2, column=2, sticky="w")

    # ---------- LIQUIDAÇÃO ----------
    tk.Label(jan, text="Planilha da liquidação (termo, nome e valor):",
             font=("Arial", 9, "bold"), anchor="w").pack(
        fill="x", padx=12, pady=(14, 2))

    quadro2 = tk.Frame(jan)
    quadro2.pack(fill="x", padx=12)

    tk.Label(quadro2, text="Aba da liquidação", font=("Arial", 8, "bold"),
             anchor="w").grid(row=0, column=0, sticky="w", pady=3)
    e_aba_liq = tk.Entry(quadro2, font=("Arial", 10), width=20)
    e_aba_liq.insert(0, cfg.get("nome_aba", "LIQUIDACAO"))
    e_aba_liq.grid(row=0, column=1, sticky="w", padx=8)
    tk.Label(quadro2, text="acento e caixa são ignorados", font=("Arial", 7),
             fg="#777777").grid(row=0, column=2, sticky="w")

    tk.Label(quadro2, text="Coluna do valor", font=("Arial", 8, "bold"),
             anchor="w").grid(row=1, column=0, sticky="w", pady=3)
    e_col_v = tk.Entry(quadro2, font=("Arial", 10), width=8)
    e_col_v.insert(0, numero_para_letra(cfg.get("col_conferencia", 17)))
    e_col_v.grid(row=1, column=1, sticky="w", padx=8)
    tk.Label(quadro2, text='"Q" ou "17"', font=("Arial", 7),
             fg="#777777").grid(row=1, column=2, sticky="w")

    tk.Label(quadro2, text="Coluna da DRE", font=("Arial", 8, "bold"),
             anchor="w").grid(row=2, column=0, sticky="w", pady=3)
    e_col_d = tk.Entry(quadro2, font=("Arial", 10), width=8)
    e_col_d.insert(0, numero_para_letra(cfg.get("col_dre", 2)))
    e_col_d.grid(row=2, column=1, sticky="w", padx=8)
    tk.Label(quadro2, text='"B" ou "2"', font=("Arial", 7),
             fg="#777777").grid(row=2, column=2, sticky="w")

    lbl_abas = tk.Label(jan, text="", font=("Arial", 7), fg="#555555",
                        anchor="w", wraplength=520, justify="left")
    lbl_abas.pack(fill="x", padx=12, pady=(8, 0))

    def _listar_abas():
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
            wb = obter_wb(excel)
            nomes = abas_do_wb(wb)
            atual = obter_aba(wb)
            usado = atual.Name if atual is not None else "(não encontrada)"
            lbl_abas.config(
                text=f'Abas em {wb.Name}: {", ".join(nomes)}\n'
                     f'Aba da liquidação em uso: {usado}',
                fg="#555555")
        except Exception as e:
            lbl_abas.config(text=f"Não consegui listar as abas ({e})", fg="#b26a00")

    linha_abas = tk.Frame(jan)
    linha_abas.pack(fill="x", padx=12, pady=(6, 0))
    tk.Button(linha_abas, text="Ver abas do arquivo", font=("Arial", 8),
              padx=6, command=_listar_abas).pack(side="left")

    linha_ac = tk.Frame(jan)
    linha_ac.pack(fill="x", padx=12, pady=(10, 0))
    lbl_ac = tk.Label(linha_ac, text="", font=("Arial", 7), fg="#555555",
                      anchor="w", wraplength=350, justify="left")
    lbl_ac.pack(side="left", fill="x", expand=True)

    def _atualizar_lbl_ac():
        p = cfg.get("caminho_access") or ""
        try:
            lbl_ac.config(text=f"Banco: {Path(p).name}" if p else "Banco: (não escolhido)")
        except Exception:
            pass

    tk.Button(linha_ac, text="Escolher banco Access…", font=("Arial", 8), padx=6,
              command=lambda: (escolher_access(), _atualizar_lbl_ac())
              ).pack(side="left", padx=4)
    tk.Button(linha_ac, text="Campos do Access…", font=("Arial", 8), padx=6,
              command=abrir_config_access).pack(side="left", padx=4)

    def _aplicar():
        fonte = "access" if var_fonte.get().startswith("Access") else "excel"
        ct = letra_para_numero(e_col_t.get())
        ce = letra_para_numero(e_col_e.get())
        cv = letra_para_numero(e_col_v.get())
        cdre = letra_para_numero(e_col_d.get())
        if fonte == "excel" and (ct <= 0 or ce <= 0):
            lbl_abas.config(text="Coluna de e-mail inválida. Use a letra (D) ou o número (4).",
                            fg="#c62828")
            return
        if cv <= 0 or cdre <= 0:
            lbl_abas.config(text="Coluna do valor ou da DRE inválida.",
                            fg="#c62828")
            return
        cfg["fonte_email"] = fonte
        cfg["nome_aba_email"] = (e_aba.get().strip() or "Dados")
        cfg["col_email_termo"] = ct or 4
        cfg["col_email_email"] = ce or 18
        cfg["nome_aba"] = (e_aba_liq.get().strip() or "LIQUIDACAO")
        cfg["col_conferencia"] = cv
        cfg["col_dre"] = cdre
        salvar_config()
        try:
            jan.destroy()
        except Exception:
            pass
        carregar_emails()
        carregar_pendencias_async()

    botoes = tk.Frame(jan)
    botoes.pack(fill="x", padx=12, pady=12)
    tk.Button(botoes, text="Aplicar", command=_aplicar, width=12,
              bg="#2e7d32", fg="white").pack(side="left")
    tk.Button(botoes, text="Cancelar", width=10,
              command=jan.destroy).pack(side="left", padx=6)

    _atualizar_lbl_ac()
    _listar_abas()

def mostrar_diagnostico_emails():
    jan = tk.Toplevel(janela_cob or janela)
    jan.title("Diagnóstico de e-mails e valores")
    jan.geometry("780x620")
    jan.transient(janela_cob or janela)

    txt = tk.Text(jan, font=("Consolas", 9), wrap="none")
    barra_y = tk.Scrollbar(jan, orient="vertical", command=txt.yview)
    barra_x = tk.Scrollbar(jan, orient="horizontal", command=txt.xview)
    txt.configure(yscrollcommand=barra_y.set, xscrollcommand=barra_x.set)
    barra_y.pack(side="right", fill="y")
    barra_x.pack(side="bottom", fill="x")
    txt.pack(fill="both", expand=True, padx=10, pady=(10, 0))

    linhas = ["CONFIGURAÇÃO ATUAL"]
    if (cfg.get("fonte_email") or "excel").lower() == "access":
        linhas.append(f"  Fonte: Access — {cfg.get('caminho_access') or '(nenhum)'}")
    else:
        linhas.append(f'  E-mails: aba "{cfg.get("nome_aba_email", "Dados")}" · '
                      f'termo {numero_para_letra(cfg.get("col_email_termo", 4))} · '
                      f'e-mail {numero_para_letra(cfg.get("col_email_email", 18))}')
    linhas.append(f'  Liquidação: aba "{cfg.get("nome_aba", "LIQUIDACAO")}" · '
                  f'valor {numero_para_letra(cfg.get("col_conferencia", 17))} · '
                  f'DRE {numero_para_letra(cfg.get("col_dre", 2))}')
    linhas.append(f"  Planilha: {cfg.get('caminho_planilha') or '(não mapeada)'}")
    linhas.append(f"  Filtro de DRE: {dre_escolhida() or TODAS_DRES}")
    linhas.append(f"  Remetente: {remetente_escolhido() or 'conta padrão do Outlook'}")
    linhas.append("")

    if ultimo_erro_email:
        linhas += ["ERRO NOS E-MAILS", f"  {ultimo_erro_email}", ""]

    linhas.append("CONTAS DO OUTLOOK")
    contas = listar_contas_outlook()
    if contas:
        for c in contas:
            linhas.append(f"  {c['smtp'] or '(sem smtp)'}  ·  {c['nome']}")
    else:
        linhas.append("  nenhuma conta lida do perfil")
    linhas.append("")

    linhas.append("LIQUIDAÇÃO (aba principal)")
    linhas.extend("  " + d for d in (cache_base_diag or ["(não lida)"]))
    if cache_base_amostra:
        linhas.append("  Amostra de valores lidos na coluna do valor:")
        for t, v in cache_base_amostra[:6]:
            linhas.append(f"    {t}  ->  {v}")
    linhas.append("")

    linhas.append("E-MAILS")
    linhas.extend("  " + d for d in (cache_email_diag or ["(fonte não carregada)"]))
    linhas.append("")

    if cache_email_amostra:
        linhas.append("AMOSTRA DE E-MAILS (termo -> e-mail)")
        for t, e in cache_email_amostra:
            linhas.append(f"  {t}  ->  {e}")
        linhas.append("")

    selecionados = registros_selecionados()
    if selecionados:
        linhas.append("RASTREAMENTO")
        for r in selecionados[:10]:
            linhas.extend(tracar_vinculo(r))
            linhas.append("")
        if len(selecionados) > 10:
            linhas.append(f"… e mais {len(selecionados) - 10} selecionado(s).")
    else:
        linhas.append("Selecione um ou mais contratados na lista para rastrear.")

    txt.insert("1.0", "\n".join(linhas))
    txt.config(state="disabled")

    botoes = tk.Frame(jan)
    botoes.pack(fill="x", padx=10, pady=8)
    tk.Button(botoes, text="Recarregar", width=12,
              command=lambda: (carregar_emails(), carregar_pendencias_async(),
                               jan.destroy())).pack(side="left")
    tk.Button(botoes, text="Fechar", width=10,
              command=jan.destroy).pack(side="left", padx=6)

def escolher_access():
    a = escolher_arquivo("Selecione o banco Access",
                         [("Access", "*.accdb *.mdb"), ("Todos", "*.*")])
    if not a:
        return
    cfg["caminho_access"] = a
    cfg["mapa_access"] = {}
    cfg["fonte_email"] = "access"
    salvar_config()
    carregar_emails()
    carregar_pendencias_async()

def abrir_config_access():
    global janela_map, cb_map_tabela, cb_map_email
    global cb_map_chave, cb_map_nome, lbl_map_status

    caminho = cfg.get("caminho_access", "")
    if not caminho:
        msg_cob("Escolha o banco Access primeiro.", "#b26a00")
        return
    conn, provider, erro = _abrir_access(caminho)
    if conn is None:
        msg_cob(f"Access: {erro}", "#c62828")
        return
    try:
        tabelas = _tabelas_access(conn)
    finally:
        try:
            conn.Close()
        except Exception:
            pass
    if not tabelas:
        msg_cob("Nenhuma tabela ou consulta encontrada no banco.", "#c62828")
        return

    if janela_map is not None:
        try:
            janela_map.destroy()
        except Exception:
            pass
    janela_map = tk.Toplevel(janela_cob or janela)
    janela_map.title("Campos do banco Access")
    janela_map.geometry("520x470")
    janela_map.transient(janela_cob or janela)

    tk.Label(janela_map,
             text="O padrão é (automático). Escolha manualmente só se a detecção errar.",
             font=("Arial", 8), fg="#555555", anchor="w", justify="left",
             wraplength=490).pack(fill="x", padx=12, pady=(10, 6))

    tk.Label(janela_map, text="Tabela / consulta",
             font=("Arial", 8, "bold"), anchor="w").pack(fill="x", padx=12)
    cb_map_tabela = ttk.Combobox(janela_map, values=tabelas, state="readonly",
                                 font=("Arial", 9))
    cb_map_tabela.set(tabelas[0])
    cb_map_tabela.pack(fill="x", padx=12)

    def _linha(rotulo):
        tk.Label(janela_map, text=rotulo, font=("Arial", 8, "bold"),
                 anchor="w").pack(fill="x", padx=12, pady=(8, 0))
        cb = ttk.Combobox(janela_map, state="readonly", font=("Arial", 9))
        cb.pack(fill="x", padx=12)
        return cb

    cb_map_email = _linha("Campo com o e-mail")
    cb_map_chave = _linha("Campo com o termo / contrato")
    cb_map_nome = _linha("Campo com o nome do contratado")

    lbl_map_status = tk.Label(janela_map, text="", font=("Arial", 7),
                              fg="#555555", anchor="w", wraplength=490,
                              justify="left")
    lbl_map_status.pack(fill="x", padx=12, pady=(8, 0))

    def _carregar_campos(event=None):
        t = cb_map_tabela.get()
        lbl_map_status.config(text=f"Lendo campos de {t}…", fg="#1565c0")
        try:
            janela_map.update_idletasks()
        except Exception:
            pass
        try:
            conn2, _p, e2 = _abrir_access(caminho)
            if conn2 is None:
                raise RuntimeError(e2)
            try:
                lista = _campos_da_tabela(conn2, t)
                if not lista:
                    raise RuntimeError("não consegui ler os campos desta tabela")
                for cb, alvos in ((cb_map_email, ALVOS_EMAIL),
                                  (cb_map_chave, ALVOS_CHAVE),
                                  (cb_map_nome, ALVOS_NOME)):
                    cb.config(values=["(automático)"] + lista)
                    atual = cb.get()
                    if atual not in ["(automático)"] + lista:
                        sug = _melhor_campo(conn2, t, lista, alvos)
                        cb.set(sug or "(automático)")
            finally:
                conn2.Close()
        except Exception as e:
            for cb in (cb_map_email, cb_map_chave, cb_map_nome):
                cb.config(values=["(automático)"])
                cb.set("(automático)")
            lbl_map_status.config(text=f"Erro: {e}", fg="#c62828")
            return
        lbl_map_status.config(
            text=f"{len(lista)} campo(s) em {t}: {', '.join(lista[:14])}"
                 + (" …" if len(lista) > 14 else ""),
            fg="#555555")

    cb_map_tabela.bind("<<ComboboxSelected>>", _carregar_campos)

    def _aplicar():
        t = cb_map_tabela.get()
        entrada = {}
        for chave, cb in (("email", cb_map_email), ("chave", cb_map_chave),
                          ("nome", cb_map_nome)):
            v = (cb.get() or "").strip()
            if v and v != "(automático)":
                entrada[chave] = v
        mapa = dict(cfg.get("mapa_access") or {})
        if entrada:
            mapa[t] = entrada
        else:
            mapa.pop(t, None)
        cfg["mapa_access"] = mapa
        salvar_config()
        try:
            janela_map.destroy()
        except Exception:
            pass
        carregar_emails()
        carregar_pendencias_async()

    botoes = tk.Frame(janela_map)
    botoes.pack(fill="x", padx=12, pady=10)
    tk.Button(botoes, text="Aplicar", command=_aplicar, width=12,
              bg="#2e7d32", fg="white").pack(side="left")
    tk.Button(botoes, text="Limpar tudo", width=11,
              command=lambda: (
                  cfg.__setitem__("mapa_access", {}), salvar_config(),
                  janela_map.destroy(), carregar_emails(),
                  carregar_pendencias_async())
              ).pack(side="left", padx=6)
    tk.Button(botoes, text="Cancelar", width=10,
              command=lambda: janela_map.destroy()).pack(side="left", padx=6)

    _carregar_campos()

# =========================
# DRE (COLUNA B)
# =========================
def ler_dres(force=False):
    global cache_dre, cache_dre_indice, cache_dre_ts
    if not force and cache_dre and (time.time() - cache_dre_ts) < DRE_VALIDADE_S:
        return True
    if not cfg.get("caminho_planilha"):
        return False
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        return False
    try:
        esperar_excel_pronto(excel)
        wb = obter_wb(excel)
        ws = obter_aba(wb)
    except Exception as e:
        log(f"DRE: planilha indisponível ({e})")
        return False
    if ws is None:
        log(f'DRE: aba "{cfg.get("nome_aba")}" não encontrada')
        return False
    try:
        ur = ws.UsedRange
        ult = int(ur.Row) + int(ur.Rows.Count) - 1
    except Exception:
        ult = 1
    ult = min(max(ult, 1), LIMITE_LINHAS_LEITURA)
    col_max = max(colunas_relevantes())
    try:
        bruto = ws.Range(ws.Cells(1, 1), ws.Cells(ult, col_max)).Value
    except Exception as e:
        log(f"DRE: leitura falhou ({e})")
        return False
    linhas = normalizar_matriz(bruto)
    ca = int(cfg["col_termo"]) - 1
    cu = int(cfg.get("col_alt") or 0) - 1
    cd = int(cfg.get("col_dre", 2)) - 1
    dre_map = {}
    for linha in linhas:
        if not (0 <= cd < len(linha)):
            continue
        sigla = texto_celula(linha[cd])
        if not sigla:
            continue
        chave = ""
        if 0 <= ca < len(linha):
            chave = chave_contrato(linha[ca])
        if not chave and 0 <= cu < len(linha):
            chave = chave_contrato(linha[cu])
        if not chave or chave in dre_map:
            continue
        dre_map[chave] = sigla
    indice = {}
    for chave, sigla in dre_map.items():
        indice.setdefault(normalizar(sigla), set()).add(chave)
    cache_dre = dre_map
    cache_dre_indice = indice
    cache_dre_ts = time.time()
    log(f"DRE: {len(dre_map)} contrato(s) indexado(s) na coluna B")
    return True

def _ler_dres_thread(force):
    if pythoncom is not None:
        try:
            pythoncom.CoInitialize()
        except Exception:
            pass
    try:
        return ler_dres(force)
    except Exception as e:
        log(f"DRE: erro na leitura ({e})")
        return False
    finally:
        if pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

def atualizar_dre(force=True):
    global dre_em_leitura
    if dre_em_leitura:
        return
    dre_em_leitura = True

    def _tarefa():
        global dre_em_leitura
        ok = False
        try:
            ok = _ler_dres_thread(force)
        finally:
            dre_em_leitura = False

        def _fim():
            if ok:
                aviso(f"DRE (coluna B): {len(cache_dre)} contrato(s) carregado(s).",
                      "#2e7d32")
                try:
                    if entrada_busca.get().strip():
                        pesquisar()
                except Exception:
                    pass
            else:
                aviso("Não consegui ler a coluna B. Abra a planilha e use F6.",
                      "#b26a00")
        try:
            janela.after(0, _fim)
        except Exception:
            pass

    threading.Thread(target=_tarefa, daemon=True).start()

def variantes_dre(termo_norm):
    variantes = {termo_norm}
    for pref in ("dre ", "dre-", "dre:", "dre."):
        if termo_norm.startswith(pref):
            resto = termo_norm[len(pref):].strip()
            if resto:
                variantes.add(resto)
    return variantes

def dres_correspondentes(termo_norm):
    if not termo_norm or not cache_dre_indice:
        return {}
    variantes = variantes_dre(termo_norm)
    exatos, parciais = {}, {}
    for sigla_norm, chaves in cache_dre_indice.items():
        if sigla_norm in variantes:
            exatos[sigla_norm] = chaves
            continue
        for v in variantes:
            if v and (v in sigla_norm or sigla_norm in v):
                parciais[sigla_norm] = chaves
                break
    return exatos or parciais

def pastas_da_chave(chave):
    return indice_chaves.get(chave, [])

# ---------- THREAD DE CONSULTA DA COLUNA Q ----------
def _ler_q_thread(pasta, geracao, ciclo):
    chave = chave_contrato(pasta.name)
    if not chave:
        return {"tipo": "final", "geracao": geracao, "ciclo": ciclo,
                "status": "sem_chave", "valor": "", "linha": 0,
                "msg": "Não identifiquei o número do contrato na pasta"}
    msg_final = "Linha não localizada"
    for tent in range(1, TENTATIVAS_POR_CICLO + 1):
        if geracao != geracao_q:
            return {"tipo": "cancelado", "geracao": geracao}
        fila_q.put({"tipo": "progresso", "geracao": geracao,
                    "tent": tent, "total": TENTATIVAS_POR_CICLO})
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
        except Exception:
            msg_final = "Excel fechado — abra a planilha"
            time.sleep(INTERVALO_TENTATIVA)
            continue
        try:
            esperar_excel_pronto(excel)
            wb = obter_wb(excel)
            ws = obter_aba(wb)
        except Exception as e:
            msg_final = f"Planilha indisponível ({e})"
            time.sleep(INTERVALO_TENTATIVA)
            continue
        if ws is None:
            msg_final = (f'aba "{cfg.get("nome_aba")}" não encontrada — '
                         f'ajuste em Fonte')
            time.sleep(INTERVALO_TENTATIVA)
            continue
        try:
            linha, motivo = localizar_por_bloco(ws, chave)
        except Exception as e:
            msg_final = f"Erro ao ler a planilha ({e})"
            time.sleep(INTERVALO_TENTATIVA)
            continue
        if linha:
            valor = ler_valor_q(ws, linha)
            return {"tipo": "final", "geracao": geracao, "ciclo": ciclo,
                    "status": "ok" if valor else "vazio",
                    "valor": valor, "linha": linha, "msg": f"Linha {linha}"}
        msg_final = f"Linha não localizada ({motivo})"
        time.sleep(INTERVALO_TENTATIVA)
    return {"tipo": "final", "geracao": geracao, "ciclo": ciclo,
            "status": "nao_encontrado", "valor": "", "linha": 0,
            "msg": msg_final}

def thread_ler_q(pasta, geracao, ciclo):
    resultado = {"tipo": "final", "geracao": geracao, "ciclo": ciclo,
                 "status": "nao_encontrado", "valor": "", "linha": 0,
                 "msg": "Planilha indisponível"}
    if pythoncom is not None:
        try:
            pythoncom.CoInitialize()
        except Exception:
            pass
    try:
        resultado = _ler_q_thread(pasta, geracao, ciclo)
    except Exception as e:
        resultado["msg"] = f"Erro na leitura ({e})"
    finally:
        if pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
    fila_q.put(resultado)

def disparar_ciclo_q(pasta, geracao, ciclo):
    if geracao != geracao_q or not pasta:
        return
    threading.Thread(target=thread_ler_q,
                     args=(pasta, geracao, ciclo), daemon=True).start()

def iniciar_consulta_q(pasta=None):
    global geracao_q, pasta_q_pendente, linha_atual, chave_atual, valor_q_atual
    pasta_q_pendente = pasta if pasta is not None else pasta_destino_atual()
    geracao_q += 1
    linha_atual = 0
    chave_atual = ""
    valor_q_atual = ""
    if not pasta_q_pendente:
        mostrar_q("—", "Selecione o contrato para ver o valor na coluna Q")
        habilitar_campos(False)
        return
    habilitar_campos(False)
    mostrar_q("…", "consultando a planilha…")
    chave_atual = chave_contrato(pasta_q_pendente.name)
    disparar_ciclo_q(pasta_q_pendente, geracao_q, 1)

def tratar_msg_q(m):
    global linha_atual, chave_atual, valor_q_atual
    if m.get("geracao") != geracao_q:
        return
    if m.get("tipo") == "progresso":
        mostrar_q("…", f"consultando a planilha ({m['tent']}/{m['total']})")
        return
    if m.get("tipo") == "cancelado":
        return
    status = m.get("status")
    ciclo = int(m.get("ciclo", 1))
    if status in ("ok", "vazio"):
        linha = int(m.get("linha", 0))
        valor = m.get("valor", "")
        dest = pasta_q_pendente or pasta_destino_atual()
        linha_atual = linha
        chave_atual = chave_contrato(dest.name) if dest else ""
        valor_q_atual = valor
        if status == "ok":
            mostrar_q(valor, f"Linha {linha} · coluna Q")
            aviso("Confira o valor acima antes de digitar a NF.", "#2e7d32")
            log(f"Coluna Q: {valor} (linha {linha})")
        else:
            mostrar_q("—", f"Linha {linha} · a coluna Q está vazia")
            aviso("Linha localizada, mas a coluna Q está vazia.", "#b26a00")
            log(f"Coluna Q vazia na linha {linha}")
        habilitar_campos(True)
        return
    if status == "sem_chave":
        mostrar_q("—", m.get("msg", ""))
        habilitar_campos(False)
        return
    if ciclo < MAX_CICLOS_Q:
        mostrar_q("…", f"{m.get('msg', '')} · tentando de novo ({ciclo}/{MAX_CICLOS_Q})")
        janela.after(
            INTERVALO_CICLO_MS,
            lambda: disparar_ciclo_q(pasta_q_pendente, geracao_q, ciclo + 1),
        )
        return
    mostrar_q("—", m.get("msg", "Linha não localizada") + " · clique para tentar de novo")
    aviso("Não consegui ler a linha. Verifique o contrato nas colunas A ou U.", "#c62828")
    habilitar_campos(True)
    log(f"Coluna Q não carregada: {m.get('msg')}")

def poll_q():
    try:
        while True:
            tratar_msg_q(fila_q.get_nowait())
    except queue.Empty:
        pass
    except Exception:
        pass
    try:
        janela.after(120, poll_q)
    except Exception:
        pass

# =========================
# RÓTULOS
# =========================
def dre_da_pasta(pasta):
    if not pasta:
        return ""
    return cache_dre.get(chave_contrato(pasta.name), "")

def atualizar_contrato_label():
    dest = pasta_destino_atual()
    if dest:
        termo = termo_do_nome(dest.name)
        dre = dre_da_pasta(dest)
        partes = [p for p in (termo, f"DRE {dre}" if dre else "") if p]
        extra = f"  ({' · '.join(partes)})" if partes else ""
        texto, cor = f"Contrato ativo: {resumir(dest.name, 36)}{extra}", "#2e7d32"
    else:
        texto, cor = "Sem contrato ativo — o arquivo aguarda em Downloads", "#e65100"
    try:
        janela.after(0, lambda: contrato_label.config(text=texto, fg=cor))
    except Exception:
        pass

def rotulo_resultado(r):
    dre = r.get("dre") or ""
    if dre:
        return f"{dre} · {resumir(r['pasta'].name, 40)}"
    return r["pasta"].name

def atualizar_preview():
    if not resultados_atuais:
        preview_label.config(text="", fg="#777777")
        return
    sel = lista.curselection()
    idx = sel[0] if sel else 0
    if idx >= len(resultados_atuais):
        idx = 0
    r = resultados_atuais[idx]
    total = len(resultados_atuais)
    via = "DRE" if r.get("via") == "dre" else "contrato"
    if total == 1:
        preview_label.config(
            text=f"Enter aplica ({via}): {resumir(r['pasta'].name)}", fg="#2e7d32")
    else:
        preview_label.config(
            text=f"Enter aplica ({idx + 1}/{total} · {via}): "
                 f"{resumir(r['pasta'].name)}", fg="#1565c0")

# =========================
# DIÁLOGOS
# =========================
def _dialogo(func):
    try:
        janela.attributes("-topmost", False)
    except Exception:
        pass
    try:
        return func()
    finally:
        try:
            janela.attributes("-topmost", True)
            janela.lift()
        except Exception:
            pass

def escolher_pasta(titulo):
    return _dialogo(lambda: filedialog.askdirectory(parent=janela, title=titulo))

def escolher_arquivo(titulo, tipos):
    return _dialogo(lambda: filedialog.askopenfilename(
        parent=janela, title=titulo, filetypes=tipos))

# =========================
# PASTAS VAZIAS (VERIFICAÇÃO ROBUSTA)
# =========================
LIXO_ARQUIVO = {"desktop.ini", "thumbs.db", "ehthumbs.db", ".ds_store"}

def _caminho_longo(p):
    """Prefixo de caminho longo do Windows: lê acima de 260 caracteres."""
    t = str(p)
    if os.name != "nt":
        return t
    if t.startswith("\\\\?\\"):
        return t
    if t.startswith("\\\\"):
        return "\\\\?\\UNC\\" + t[2:]
    return "\\\\?\\" + t

def _arquivos_da_pasta(pasta):
    """
    Lista os arquivos reais da pasta em TODOS os níveis de subpasta.
    Devolve (lista_de_arquivos, quantidade_de_pastas_ilegiveis).
    """
    achados = []
    erros = 0
    pilha = [Path(pasta)]
    while pilha:
        atual = pilha.pop()
        try:
            it = os.scandir(_caminho_longo(atual))
        except (OSError, ValueError):
            erros += 1
            continue
        try:
            for e in it:
                try:
                    nome = e.name
                    if nome.lower() in LIXO_ARQUIVO:
                        continue
                    if e.is_dir(follow_symlinks=False):
                        pilha.append(Path(atual) / nome)
                        continue
                    if not e.is_file(follow_symlinks=False):
                        continue
                    if nome.startswith("~$"):
                        continue
                    baixo = nome.lower()
                    if baixo.endswith((".tmp", ".crdownload", ".part")):
                        try:
                            if e.stat().st_size == 0:
                                continue
                        except OSError:
                            continue
                    achados.append(Path(atual) / nome)
                except OSError:
                    continue
        finally:
            try:
                it.close()
            except Exception:
                pass
    return achados, erros

def pasta_sem_documentos(pasta, profundidade=None):
    """
    True SOMENTE quando a pasta foi lida por inteiro e não tem nenhum arquivo.
    Sem limite de profundidade. Se alguma subpasta não puder ser lida,
    devolve False, para nunca acusar falta de documento por engano.
    """
    achados, erros = _arquivos_da_pasta(pasta)
    if achados:
        return False
    if erros:
        return False
    return True

def contar_pastas_vazias():
    total = 0
    for p in list(cache_pastas.values()):
        if pasta_sem_documentos(p):
            total += 1
    return total

def recontar_vazias():
    global pastas_vazias
    pastas_vazias = contar_pastas_vazias()
    return pastas_vazias

def pastas_vazias_lista():
    """Pastas indexadas sem nenhum arquivo, nem dentro das subpastas."""
    if not pasta_principal or not pasta_principal.is_dir():
        raise RuntimeError("mapeie a pasta principal (botão 'Principal')")
    if not cache_pastas:
        criar_cache(silencioso=True)
    if not cache_pastas:
        raise RuntimeError(
            f'nenhuma subpasta encontrada em {pasta_principal}. '
            f'Use o botão "Principal" para reindexar.')

    base = {}
    if cfg.get("caminho_planilha"):
        try:
            base = ler_base_planilha()
        except Exception:
            base = {}

    achados = []
    for p in cache_pastas.values():
        if not pasta_sem_documentos(p):
            continue
        chave = chave_contrato(p.name)
        dados = base.get(chave, {}) if chave else {}
        achados.append({
            "pasta": p,
            "chave": chave,
            "dre": dados.get("dre", "") or cache_dre.get(chave, ""),
            "nome": dados.get("nome", "") or p.name,
            "termo": dados.get("termo") or termo_do_nome(p.name) or chave,
            "valor": dados.get("valor", ""),
        })
    achados.sort(key=lambda r: (normalizar(r.get("dre") or ""),
                                normalizar(r["pasta"].name)))
    return achados

def atualizar_vazias():
    global vazias_atuais, pastas_vazias
    if lista_vazias is None:
        return
    lista_vazias.delete(0, tk.END)
    vazias_atuais = []
    status_vazias.config(text="Lendo as pastas…", fg="#1565c0")
    try:
        janela_vazias.update_idletasks()
    except Exception:
        pass
    try:
        todas = pastas_vazias_lista()
    except Exception as e:
        status_vazias.config(text=str(e), fg="#c62828")
        return

    filtro = ""
    try:
        filtro = normalizar(entrada_filtro_vazias.get().strip())
    except Exception:
        filtro = ""

    for r in todas:
        if filtro:
            alvo = normalizar(f"{r['pasta'].name} {r['nome']} "
                              f"{r.get('termo') or ''} {r.get('dre') or ''}")
            if filtro not in alvo:
                continue
        vazias_atuais.append(r)
        linha = (f"{resumir(r.get('dre') or '-', 8):<8} | "
                 f"{resumir(r['nome'], 26):<26} | "
                 f"{(r.get('termo') or '-'):<13} | "
                 f"{resumir(r['pasta'].name, 28)}")
        lista_vazias.insert(tk.END, linha)

    pastas_vazias = len(todas)
    painel()

    total_idx = len(cache_pastas)
    if filtro:
        status_vazias.config(
            text=f"{len(vazias_atuais)} de {len(todas)} pasta(s) vazia(s) · "
                 f"{total_idx} indexada(s) · filtro ativo", fg="#b26a00")
    elif vazias_atuais:
        status_vazias.config(
            text=f"{len(vazias_atuais)} pasta(s) sem nenhum documento · "
                 f"{total_idx} indexada(s)", fg="#b26a00")
    else:
        status_vazias.config(
            text=f"Nenhuma pasta vazia · {total_idx} indexada(s)", fg="#2e7d32")

def abrir_pasta_selecionada():
    if lista_vazias is None:
        return
    sel = lista_vazias.curselection()
    if not sel or sel[0] >= len(vazias_atuais):
        status_vazias.config(text="Selecione uma pasta na lista.", fg="#b26a00")
        return
    pasta = vazias_atuais[sel[0]]["pasta"]
    try:
        os.startfile(str(pasta))
    except Exception as e:
        status_vazias.config(text=f"Não consegui abrir: {e}", fg="#c62828")

def copiar_vazias():
    if lista_vazias is None:
        return
    sel = lista_vazias.curselection()
    if sel:
        alvo = [vazias_atuais[i] for i in sel if i < len(vazias_atuais)]
    else:
        alvo = list(vazias_atuais)
    if not alvo:
        status_vazias.config(text="Nada para copiar.", fg="#b26a00")
        return
    linhas = ["DRE\tTermo\tContratado\tValor\tPasta\tCaminho"]
    for r in alvo:
        linhas.append("\t".join([
            r.get("dre") or "",
            r.get("termo") or "",
            r.get("nome") or "",
            formatar_moeda(r.get("valor")),
            r["pasta"].name,
            str(r["pasta"]),
        ]))
    try:
        janela.clipboard_clear()
        janela.clipboard_append("\n".join(linhas))
        status_vazias.config(
            text=f"{len(alvo)} linha(s) copiada(s) — cole direto no Excel.",
            fg="#2e7d32")
    except Exception as e:
        status_vazias.config(text=f"Não consegui copiar: {e}", fg="#c62828")

def fechar_vazias():
    global janela_vazias
    try:
        janela_vazias.destroy()
    except Exception:
        pass
    janela_vazias = None

def abrir_pastas_vazias():
    global janela_vazias, lista_vazias, status_vazias, entrada_filtro_vazias
    if janela_vazias is not None:
        try:
            if janela_vazias.winfo_exists():
                janela_vazias.deiconify()
                janela_vazias.lift()
                atualizar_vazias()
                return
        except Exception:
            pass
        janela_vazias = None

    janela_vazias = tk.Toplevel(janela)
    janela_vazias.title("Pastas sem documentos")
    janela_vazias.geometry("780x580")
    janela_vazias.transient(janela)

    tk.Label(janela_vazias, text="Pastas sem nenhum documento",
             font=("Arial", 11, "bold"), anchor="w").pack(
        fill="x", padx=10, pady=(10, 2))
    tk.Label(janela_vazias,
             text="Considera vazia a pasta que não tem nenhum arquivo, "
                  "em nenhum nível de subpasta. Pasta que não puder ser lida "
                  "não entra na lista. Duplo clique abre no Explorer.",
             font=("Arial", 8), fg="#555555", anchor="w", justify="left",
             wraplength=750).pack(fill="x", padx=10)

    filtro = tk.Frame(janela_vazias)
    filtro.pack(fill="x", padx=10, pady=(8, 0))
    tk.Label(filtro, text="Filtrar:", font=("Arial", 8, "bold")).pack(side="left")
    entrada_filtro_vazias = tk.Entry(filtro, font=("Arial", 9))
    entrada_filtro_vazias.pack(side="left", fill="x", expand=True, padx=(6, 4))
    entrada_filtro_vazias.bind("<KeyRelease>", lambda e: atualizar_vazias())
    tk.Button(filtro, text="Limpar", font=("Arial", 7), padx=4, pady=0,
              command=lambda: (entrada_filtro_vazias.delete(0, tk.END),
                               atualizar_vazias())).pack(side="left")

    tk.Label(janela_vazias, text="DRE | contratado | termo | pasta:",
             font=("Arial", 8, "bold"), anchor="w").pack(
        fill="x", padx=10, pady=(8, 0))

    quadro = tk.Frame(janela_vazias)
    quadro.pack(fill="both", expand=True, padx=10)
    lista_vazias = tk.Listbox(quadro, font=("Consolas", 8),
                              selectmode=tk.EXTENDED, activestyle="dotbox")
    barra = tk.Scrollbar(quadro, orient="vertical", command=lista_vazias.yview)
    lista_vazias.configure(yscrollcommand=barra.set)
    barra.pack(side="right", fill="y")
    lista_vazias.pack(side="left", fill="both", expand=True)
    lista_vazias.bind("<Double-Button-1>",
                      lambda e: abrir_pasta_selecionada())

    botoes = tk.Frame(janela_vazias)
    botoes.pack(fill="x", padx=10, pady=8)
    tk.Button(botoes, text="Copiar lista", command=copiar_vazias,
              width=13).pack(side="left", padx=2)
    tk.Button(botoes, text="Abrir no Explorer", command=abrir_pasta_selecionada,
              width=17).pack(side="left", padx=2)
    tk.Button(botoes, text="Atualizar", command=atualizar_vazias,
              width=10).pack(side="left", padx=2)
    tk.Button(botoes, text="Fechar", command=fechar_vazias,
              width=8).pack(side="right", padx=2)

    status_vazias = tk.Label(janela_vazias, text="", anchor="w",
                             font=("Arial", 8), fg="#555555",
                             wraplength=750, justify="left")
    status_vazias.pack(fill="x", padx=10, pady=(0, 8))

    janela_vazias.protocol("WM_DELETE_WINDOW", fechar_vazias)
    atualizar_vazias()

# =========================
# CACHE DE PASTAS
# =========================
def criar_cache(silencioso=False):
    global indice_chaves
    cache_pastas.clear()
    indice_chaves = {}
    if not pasta_principal or not pasta_principal.is_dir():
        if not silencioso:
            log("Pasta principal não mapeada.")
        return
    base = pasta_principal
    prof = int(cfg.get("profundidade_busca", 2))
    for raiz, dirs, _ in os.walk(base):
        try:
            nivel = len(Path(raiz).relative_to(base).parts)
        except ValueError:
            nivel = 0
        if nivel >= prof:
            dirs[:] = []
            continue
        dirs[:] = [d for d in dirs if not d.startswith(".")
                   and d not in ("$RECYCLE.BIN", "System Volume Information")]
        for d in dirs:
            p = Path(raiz) / d
            cache_pastas[str(p)] = p
            chave = chave_contrato(p.name)
            if chave:
                indice_chaves.setdefault(chave, []).append(p)
    if not silencioso:
        log(f"{len(cache_pastas)} pastas indexadas | vazias: {recontar_vazias()}")
        painel()

# =========================
# MAPEAMENTOS
# =========================
def mapear_principal():
    global pasta_principal
    e = escolher_pasta("Pasta principal (contém as subpastas)")
    if not e:
        return
    pasta_principal = Path(e)
    cfg["pasta_principal"] = str(pasta_principal)
    salvar_config()
    criar_cache()
    carregar_pendencias_async()

def mapear_download():
    global pasta_download
    e = escolher_pasta("Pasta de Downloads")
    if not e:
        return
    pasta_download = Path(e)
    cfg["pasta_download"] = str(pasta_download)
    salvar_config()
    estado.clear()
    log(f"Downloads: {pasta_download}")
    if not monitorando:
        iniciar()

def mapear_planilha():
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
        wb = excel.ActiveWorkbook
        if wb:
            full = str(wb.FullName)
            if full.lower().startswith(("http://", "https://")):
                aviso("Planilha aberta pela nuvem. Abra pela pasta do OneDrive.")
            else:
                cfg["caminho_planilha"] = full
                salvar_config()
                ws = obter_aba(wb)
                usado = ws.Name if ws is not None else "(aba não encontrada)"
                aviso(f"Planilha: {wb.Name} · liquidação: {usado}", "#2e7d32")
                atualizar_dre(True)
                if pasta_destino_atual():
                    iniciar_consulta_q()
                carregar_pendencias_async()
                return
    except Exception:
        pass
    a = escolher_arquivo("Selecione a planilha",
                         [("Excel", "*.xlsx *.xlsm *.xls")])
    if a:
        cfg["caminho_planilha"] = a
        salvar_config()
        aviso(f"Planilha: {Path(a).name}", "#2e7d32")
        atualizar_dre(True)
        if pasta_destino_atual():
            iniciar_consulta_q()
        carregar_pendencias_async()

# =========================
# GRAVAÇÃO NA PLANILHA
# =========================
def para_celula(t):
    t = str(t).strip()
    if re.fullmatch(r"\d{1,9}", t) and not t.startswith("0"):
        return int(t)
    return t

def salvar_excel(event=None):
    global linha_atual, chave_atual, valor_q_atual
    if str(entrada_nf.cget("state")) == "disabled":
        aviso("Aguarde o valor da coluna Q aparecer para liberar a digitação.")
        return
    dest = pasta_destino_atual()
    if not dest:
        aviso("Selecione o contrato antes de gravar.")
        return
    nf = entrada_nf.get().strip()
    dia = entrada_dia.get().strip()
    if not nf or not dia:
        aviso("Preencha a NF e o Dia.")
        return
    if not cfg["caminho_planilha"]:
        aviso("Planilha não mapeada. Clique em 'Planilha'.")
        return
    chave = chave_contrato(dest.name)
    if not chave:
        aviso("Não achei o número do contrato no nome da pasta.")
        return
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        aviso("O Excel não está aberto.")
        return
    try:
        esperar_excel_pronto(excel)
        wb = obter_wb(excel)
        ws = obter_aba(wb)
        if ws is None:
            raise RuntimeError(f'aba "{cfg.get("nome_aba")}" não encontrada')
    except Exception as e:
        aviso("Não consegui abrir a aba da liquidação.")
        log(f"Planilha: {e}")
        return
    if chave_atual == chave and linha_atual:
        linha, motivo = linha_atual, "linha já carregada"
    else:
        try:
            linha, motivo = localizar_por_bloco(ws, chave)
        except Exception as e:
            aviso("Erro ao ler a planilha.")
            log(f"Erro ao localizar: {e}")
            return
    if not linha:
        aviso(f"Não gravado: {motivo}")
        log(f"Não gravado: {motivo}")
        mostrar_q("—", f"Linha não localizada: {motivo} · clique para tentar de novo")
        return
    valor_q = ler_valor_q(ws, linha)
    if valor_q:
        mostrar_q(valor_q, f"Linha {linha} · coluna Q")
    valor_q_atual = valor_q
    try:
        col_dia = int(cfg["col_dia"])
        col_nf = int(cfg["col_nf"])
        ws.Cells(linha, col_dia).Value = para_celula(dia)
        ws.Cells(linha, col_nf).Value = para_celula(nf)
        try:
            excel.Calculate()
        except Exception:
            pass
        if cfg.get("salvar_apos_gravar"):
            try:
                wb.Save()
            except Exception:
                pass
        lido_d = texto_celula(ws.Cells(linha, col_dia).Value)
        lido_n = texto_celula(ws.Cells(linha, col_nf).Value)
        if digitos(lido_d) == digitos(dia) and digitos(lido_n) == digitos(nf):
            aviso(f"Gravado na linha {linha} — Dia {lido_d} / NF {lido_n}", "#2e7d32")
            log(f"GRAVADO linha {linha} | Dia {lido_d} | NF {lido_n}")
            entrada_nf.delete(0, tk.END)
            entrada_dia.delete(0, tk.END)
            entrada_nf.focus_set()
        else:
            aviso(f"Verificação falhou: Dia '{lido_d}' NF '{lido_n}'")
            log(f"Verificação falhou: Dia '{lido_d}' NF '{lido_n}'")
    except Exception as e:
        aviso("Erro ao gravar na planilha.")
        log(f"Erro ao gravar: {e}")

# =========================
# LIMPEZA
# =========================
def valores_campos():
    try:
        return (entrada_nf.get().strip(), entrada_dia.get().strip(),
                entrada_busca.get().strip())
    except Exception:
        return ("", "", "")

def capturar_snapshot():
    global snapshot_campos
    if snapshot_campos is None:
        snapshot_campos = valores_campos()

def resetar_contrato():
    global pasta_destino_global, snapshot_campos, ultimo_movimento
    global geracao_q, pasta_q_pendente, linha_atual, chave_atual, valor_q_atual
    pasta_destino_global = None
    cfg["pasta_destino"] = ""
    salvar_config()
    snapshot_campos = None
    ultimo_movimento = 0.0
    geracao_q += 1
    pasta_q_pendente = None
    linha_atual = 0
    chave_atual = ""
    valor_q_atual = ""
    entrada_busca.delete(0, tk.END)
    lista.delete(0, tk.END)
    resultados_atuais.clear()
    preview_label.config(text="")
    habilitar_campos(False)
    atualizar_contrato_label()
    painel()

def limpar_tudo():
    resetar_contrato()
    mostrar_q("—", "Selecione o contrato para ver o valor na coluna Q")
    aviso("Arquivo arquivado. Campos limpos — informe o próximo contrato.", "#1565c0")
    log("Campos limpos. Pronto para o próximo contrato.")
    entrada_busca.focus_set()

def limpar_contrato(event=None):
    resetar_contrato()
    mostrar_q("—", "Selecione o contrato para ver o valor na coluna Q")
    aviso("Contrato desmarcado. O próximo arquivo aguardará em Downloads.", "#e65100")
    log("Contrato desmarcado.")
    entrada_busca.focus_set()

# =========================
# MOVIMENTAÇÃO
# =========================
def mover(caminho, destino_dir):
    global movidos, ultimo_movimento, recontar
    destino = destino_dir / caminho.name
    n = 1
    while destino.exists():
        destino = destino_dir / f"{caminho.stem}_{n}{caminho.suffix}"
        n += 1
    if cfg.get("acao") == "copiar":
        shutil.copy2(str(caminho), str(destino))
    else:
        shutil.move(str(caminho), str(destino))
    if destino.suffix.lower() in EXT_IMAGEM:
        converter_pdf(destino)
    movidos += 1
    ultimo_movimento = time.time()
    recontar = True
    log(f"MOVIDO: {caminho.name} -> {destino_dir.name}")

def _alvo_seguro(base, membro):
    try:
        alvo = (base / membro).resolve()
        alvo.relative_to(base.resolve())
        return alvo
    except Exception:
        return None

def extrair_zip(arquivo, destino_dir):
    global movidos, ultimo_movimento, recontar
    with zipfile.ZipFile(arquivo, "r") as z:
        base = destino_dir.resolve()
        for membro in z.namelist():
            if _alvo_seguro(base, membro) is None:
                raise ValueError("ZIP com caminho inválido")
        z.extractall(destino_dir)
    arquivo.unlink()
    movidos += 1
    ultimo_movimento = time.time()
    recontar = True
    log(f"ZIP extraído em {destino_dir.name}")

def converter_pdf(arquivo):
    try:
        pdf = arquivo.with_suffix(".pdf")
        n = 1
        while pdf.exists():
            pdf = arquivo.with_name(f"{arquivo.stem}_{n}.pdf")
            n += 1
        with Image.open(arquivo) as img:
            quadros = []
            for quadro in ImageSequence.Iterator(img):
                try:
                    quadros.append(quadro.convert("RGB"))
                except Exception:
                    continue
        if not quadros:
            return
        if len(quadros) > 1:
            quadros[0].save(pdf, save_all=True, append_images=quadros[1:])
        else:
            quadros[0].save(pdf)
        arquivo.unlink()
    except Exception as e:
        log(f"Erro ao converter imagem: {e}")

def worker():
    while True:
        item = fila.get()
        if item is None:
            fila.task_done()
            continue
        caminho, destino_dir = item
        info = estado.get(str(caminho))
        try:
            if caminho.suffix.lower() == ".zip":
                extrair_zip(caminho, destino_dir)
            else:
                mover(caminho, destino_dir)
            estado.pop(str(caminho), None)
        except PermissionError:
            if info is not None:
                info["tentativas"] = info.get("tentativas", 0) + 1
                if info["tentativas"] >= 5:
                    info["status"] = "erro"
                    log(f"Arquivo em uso, desisti: {caminho.name}")
                else:
                    info["status"] = "novo"
                    info["estavel"] = 0
        except Exception as e:
            if info is not None:
                info["tentativas"] = info.get("tentativas", 0) + 1
                if info["tentativas"] >= 5:
                    info["status"] = "erro"
                else:
                    info["status"] = "novo"
                    info["estavel"] = 0
            log(f"Erro em {caminho.name}: {e}")
        finally:
            fila.task_done()

# =========================
# VARREDURA
# =========================
def listar_candidatos():
    achados = []
    if not pasta_download or not pasta_download.is_dir():
        return achados
    try:
        if not cfg.get("incluir_subpastas_download", True):
            return [p for p in pasta_download.iterdir() if p.is_file()]
        for raiz, dirs, nomes in os.walk(pasta_download):
            try:
                nivel = len(Path(raiz).relative_to(pasta_download).parts)
            except ValueError:
                nivel = 0
            if nivel >= 2:
                dirs[:] = []
            dirs[:] = [d for d in dirs if not d.startswith(".")]
            for n in nomes:
                achados.append(Path(raiz) / n)
    except OSError as e:
        log(f"Erro ao ler Downloads: {e}")
    return achados

def varrer():
    global candidatos_atuais
    if not pasta_download or not pasta_download.is_dir():
        candidatos_atuais = 0
        return 0
    candidatos = listar_candidatos()
    for k in list(estado.keys()):
        if not Path(k).exists():
            estado.pop(k, None)
    destino = pasta_destino_atual()
    aguardando = 0
    total = 0
    ignorar_antigos = bool(cfg.get("ignorar_arquivos_antigos", True))
    for f in candidatos:
        try:
            if not f.is_file():
                continue
        except OSError:
            continue
        if f.name.startswith("~$") or f.suffix.lower() in EXT_IGNORAR:
            continue
        try:
            st = f.stat()
        except OSError:
            continue
        tam = st.st_size
        mtime = st.st_mtime
        if tam <= 0:
            continue
        if ignorar_antigos and mtime < INICIO_PROGRAMA - TOLERANCIA_ANTIGOS_S:
            continue
        total += 1
        k = str(f)
        info = estado.get(k)
        if info is None:
            estado[k] = {"size": tam, "mtime": mtime, "estavel": 0,
                         "status": "novo", "tentativas": 0}
            continue
        if info.get("status") in ("fila", "movido", "erro"):
            continue
        if info.get("size") == tam and info.get("mtime") == mtime:
            info["estavel"] = info.get("estavel", 0) + 1
        else:
            info["size"] = tam
            info["mtime"] = mtime
            info["estavel"] = 0
            info["status"] = "novo"
            continue
        if info["estavel"] < 1:
            continue
        if destino is None:
            info["status"] = "aguardando"
            aguardando += 1
            if not info.get("avisado"):
                log(f"Aguardando contrato: {f.name}")
                info["avisado"] = True
            continue
        capturar_snapshot()
        info["status"] = "fila"
        info.pop("avisado", None)
        fila.put((f, destino))
    candidatos_atuais = total
    return aguardando

def nada_pendente():
    if not fila.empty():
        return False
    for info in estado.values():
        if info.get("status") in ("novo", "fila", "aguardando"):
            return False
    return True

def verificar_limpeza():
    global ultimo_movimento, snapshot_campos
    if not ultimo_movimento:
        return
    espera = float(cfg.get("espera_limpeza_s", 2.5))
    if time.time() - ultimo_movimento < espera:
        return
    if not nada_pendente():
        ultimo_movimento = time.time()
        return
    ultimo_movimento = 0.0
    atual = valores_campos()
    if snapshot_campos is not None and atual != snapshot_campos and any(atual):
        log("Campos mantidos: você já iniciou o próximo registro.")
        snapshot_campos = None
        return
    snapshot_campos = None
    if pasta_destino_atual() is None:
        return
    limpar_tudo()

def ciclo():
    global timer, recontar
    if not monitorando:
        return
    try:
        aguardando = varrer()
        if recontar:
            recontar = False
            recontar_vazias()
        painel(aguardando)
        verificar_limpeza()
    except Exception as e:
        log(f"Erro na varredura: {e}")
    finally:
        if monitorando:
            try:
                timer = janela.after(int(cfg.get("intervalo_ms", 2000)), ciclo)
            except Exception:
                timer = None

# =========================
# INICIAR / PARAR
# =========================
def iniciar():
    global monitorando, candidatos_atuais
    if not pasta_download or not pasta_download.is_dir():
        aviso("Mapeie a pasta de Downloads.")
        return
    if monitorando:
        return
    estado.clear()
    candidatos_atuais = 0
    monitorando = True
    botao_toggle.config(text="Parar", bg="#c62828")
    log(f"Monitorando: {pasta_download}")
    if cfg.get("ignorar_arquivos_antigos", True):
        log("Arquivos anteriores à abertura serão ignorados.")
    else:
        log("Tudo em Downloads será movido, inclusive arquivos antigos.")
    painel()
    ciclo()

def parar():
    global monitorando, timer
    monitorando = False
    if timer:
        try:
            janela.after_cancel(timer)
        except Exception:
            pass
        timer = None
    botao_toggle.config(text="Iniciar", bg="#2e7d32")
    log("Monitoramento parado.")
    painel()

def alternar():
    parar() if monitorando else iniciar()

# =========================
# BUSCA E SELEÇÃO
# =========================
def pesquisar(event=None):
    global resultados_atuais
    txt = entrada_busca.get().strip()
    termo, termo_d = normalizar(txt), digitos(txt)
    lista.delete(0, tk.END)
    resultados_atuais = []
    if not termo:
        atualizar_preview()
        return
    if not cache_pastas:
        criar_cache()
    if not cache_dre and not dre_em_leitura:
        atualizar_dre(False)

    vistos = {}

    for p in cache_pastas.values():
        nome = normalizar(p.name)
        bate = termo in nome
        if not bate and len(termo_d) >= 5 and termo_d in digitos(p.name):
            bate = True
        if bate:
            chave = chave_contrato(p.name)
            vistos[str(p)] = {
                "pasta": p,
                "chave": chave,
                "dre": cache_dre.get(chave, ""),
                "via": "contrato",
            }

    chaves_dre = set()
    for chaves in dres_correspondentes(termo).values():
        chaves_dre |= chaves
    for chave in chaves_dre:
        for p in pastas_da_chave(chave):
            if str(p) in vistos:
                continue
            vistos[str(p)] = {
                "pasta": p,
                "chave": chave,
                "dre": cache_dre.get(chave, ""),
                "via": "DRE",
            }

    resultados_atuais = sorted(
        vistos.values(),
        key=lambda r: (normalizar(r.get("dre") or ""), normalizar(r["pasta"].name))
    )
    for r in resultados_atuais:
        lista.insert(tk.END, rotulo_resultado(r))

    if not resultados_atuais:
        if chaves_dre:
            aviso(f"{len(chaves_dre)} contrato(s) dessa DRE, mas sem pasta indexada. "
                  f"Confira a pasta principal.", "#b26a00")
        elif not cache_dre and cfg.get("caminho_planilha"):
            aviso("DRE ainda não carregada. Use F6 depois de abrir a planilha.",
                  "#b26a00")
        else:
            aviso("Nenhuma pasta encontrada para esse termo.")
    elif chaves_dre:
        aviso(f"DRE: {len(resultados_atuais)} contrato(s) com pasta. "
              f"Confirme na lista abaixo.", "#2e7d32")
    atualizar_preview()

def mover_selecao(delta):
    if not resultados_atuais:
        return
    sel = lista.curselection()
    idx = sel[0] if sel else (-1 if delta > 0 else 0)
    idx = max(0, min(len(resultados_atuais) - 1, idx + delta))
    lista.selection_clear(0, tk.END)
    lista.selection_set(idx)
    lista.activate(idx)
    lista.see(idx)
    atualizar_preview()

def tecla_baixo(event=None):
    mover_selecao(1)
    return "break"

def tecla_cima(event=None):
    mover_selecao(-1)
    return "break"

def confirmar(event=None):
    if not resultados_atuais:
        aviso("Digite parte do contrato, o nome ou a sigla de DRE.")
        return
    sel = lista.curselection()
    idx = sel[0] if sel else 0
    if idx >= len(resultados_atuais):
        idx = 0
    escolher_contrato(resultados_atuais[idx]["pasta"])

def clicar_lista(event=None):
    sel = lista.curselection()
    if not sel or sel[0] >= len(resultados_atuais):
        return
    escolher_contrato(resultados_atuais[sel[0]]["pasta"])

def escolher_contrato(pasta):
    global pasta_destino_global, ultimo_movimento, snapshot_campos
    pasta_destino_global = pasta
    cfg["pasta_destino"] = str(pasta)
    salvar_config()
    ultimo_movimento = 0.0
    snapshot_campos = None
    termo = termo_do_nome(pasta.name)
    dre = dre_da_pasta(pasta)
    atualizar_contrato_label()
    extra = f" · DRE {dre}" if dre else ""
    aviso(f"Contrato: {termo or pasta.name}{extra} — lendo a coluna Q…", "#1565c0")
    log(f"Contrato ativo: {pasta.name}{extra}")
    entrada_busca.delete(0, tk.END)
    lista.delete(0, tk.END)
    resultados_atuais.clear()
    preview_label.config(text="")
    iniciar_consulta_q(pasta)
    if pasta_download:
        try:
            painel(varrer())
        except Exception as e:
            log(f"Erro ao varrer: {e}")

def recarregar():
    if not pasta_principal:
        aviso("Mapeie a pasta principal.")
        return
    criar_cache()
    carregar_pendencias_async()

# =========================
# COBRANÇA DE DOCUMENTOS
# =========================
def formatar_moeda(valor):
    t = texto_celula(valor)
    if not t:
        return "a confirmar"
    limpo = re.sub(r"[^\d,.\-]", "", t)
    if not re.search(r"\d", limpo):
        return t
    num = None
    try:
        if "," in limpo:
            limpo = limpo.replace(".", "").replace(",", ".")
        num = float(limpo)
    except Exception:
        num = None
    if num is None:
        return t if "R$" in t.upper() else f"R$ {t}"
    s = f"{num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {s}"

def ler_base_planilha():
    """Lê a aba da liquidação. Falha com mensagem clara quando não acha a aba."""
    global cache_base_info, cache_base_diag, cache_base_amostra

    if not cfg.get("caminho_planilha"):
        raise RuntimeError("planilha não mapeada — clique em 'Planilha'")
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        raise RuntimeError("o Excel não está aberto")

    esperar_excel_pronto(excel)
    try:
        wb = obter_wb(excel)
    except Exception as e:
        raise RuntimeError(f"não consegui localizar a planilha ({e})")

    ws = obter_aba(wb)
    if ws is None:
        raise RuntimeError(
            f'aba "{cfg.get("nome_aba")}" não encontrada em {wb.Name}. '
            f'Abas do arquivo: {", ".join(abas_do_wb(wb))}. '
            f'Ajuste em "Fonte".'
        )

    col_max = max(colunas_relevantes())
    try:
        ur = ws.UsedRange
        ult = int(ur.Row) + int(ur.Rows.Count) - 1
    except Exception:
        ult = 1
    ult = min(max(ult, 1), LIMITE_LINHAS_LEITURA)
    try:
        bruto = ws.Range(ws.Cells(1, 1), ws.Cells(ult, col_max)).Value
    except Exception as e:
        raise RuntimeError(f"erro ao ler a aba {ws.Name} ({e})")

    linhas = normalizar_matriz(bruto)
    ca = int(cfg["col_termo"]) - 1
    cu = int(cfg.get("col_alt") or 0) - 1
    cq = int(cfg.get("col_conferencia", 17)) - 1
    cn = int(cfg.get("col_nf") or 0) - 1
    cd = int(cfg.get("col_dre", 2)) - 1

    base = {}
    com_valor = 0
    for i, linha in enumerate(linhas):
        chave = ""
        if 0 <= ca < len(linha):
            chave = chave_contrato(linha[ca])
        if not chave and 0 <= cu < len(linha):
            chave = chave_contrato(linha[cu])
        if not chave or chave in base:
            continue
        valor = texto_celula(linha[cq]) if 0 <= cq < len(linha) else ""
        if valor:
            com_valor += 1
        base[chave] = {
            "linha": i + 1,
            "termo": texto_celula(linha[ca]) if 0 <= ca < len(linha) else "",
            "dre": texto_celula(linha[cd]) if 0 <= cd < len(linha) else "",
            "nome": texto_celula(linha[cu]) if 0 <= cu < len(linha) else "",
            "valor": valor,
            "nf": texto_celula(linha[cn]) if 0 <= cn < len(linha) else "",
        }

    vazios = [d for d in base.values() if not d.get("valor")]
    if vazios:
        col_q = int(cfg.get("col_conferencia", 17))
        for dados in vazios[:80]:
            ln = dados.get("linha") or 0
            if ln <= 0:
                continue
            try:
                txt = str(ws.Cells(ln, col_q).Text or "").strip()
            except Exception:
                txt = ""
            if txt:
                dados["valor"] = txt
                com_valor += 1

    amostra = []
    for chave, dados in list(base.items())[:8]:
        amostra.append((dados.get("termo") or chave,
                        dados.get("valor") or "(vazio)"))

    cache_base_amostra = amostra
    cache_base_info = (f'aba "{ws.Name}" · {len(base)} contrato(s) · '
                       f'{com_valor} com valor em '
                       f'{numero_para_letra(cfg.get("col_conferencia", 17))}')
    cache_base_diag = [
        f'Aba lida: "{ws.Name}"',
        f'Linhas com conteúdo: {len(linhas)}',
        f'Contratos reconhecidos: {len(base)}',
        f'Com valor preenchido: {com_valor}',
        f'Sem valor: {len(base) - com_valor}',
        f'Com NF preenchida: {sum(1 for d in base.values() if d.get("nf"))}',
        f'Coluna do valor: {numero_para_letra(cfg.get("col_conferencia", 17))} '
        f'({cfg.get("col_conferencia", 17)})',
        f'Coluna do termo: {numero_para_letra(cfg.get("col_termo", 1))}',
        f'Coluna do nome: {numero_para_letra(cfg.get("col_alt", 21))}',
        f'Coluna da DRE: {numero_para_letra(cfg.get("col_dre", 2))}',
        f'DREs distintas: {len(dres_disponiveis(base))}',
    ]

    if not base:
        amostra_a = []
        for linha in linhas[:8]:
            v = texto_celula(linha[ca]) if 0 <= ca < len(linha) else ""
            if v:
                amostra_a.append(v)
        raise RuntimeError(
            f'a aba "{ws.Name}" foi lida ({len(linhas)} linha(s)), mas nenhum '
            f'contrato foi reconhecido nas colunas '
            f'{numero_para_letra(cfg.get("col_termo", 1))} e '
            f'{numero_para_letra(cfg.get("col_alt", 21))}. '
            f'Amostra da coluna A: {amostra_a or "(vazia)"}'
        )

    log(f"Liquidação: {cache_base_info}")
    return base

def dres_disponiveis(base):
    """Siglas de DRE existentes na coluna B, sem duplicidade e ordenadas."""
    vistos = {}
    try:
        for dados in base.values():
            s = texto_celula(dados.get("dre") or "")
            if not s:
                continue
            k = normalizar(s)
            if k and k not in vistos:
                vistos[k] = s
    except Exception:
        pass
    return [vistos[k] for k in sorted(vistos)]

def montar_pendentes(criterio, base=None, dre_filtro=""):
    if base is None:
        base = ler_base_planilha()
    achados = []

    if criterio == "pasta":
        if not cache_pastas and pasta_principal:
            criar_cache(silencioso=True)
        for p in cache_pastas.values():
            chave = chave_contrato(p.name)
            if not chave:
                continue
            if not pasta_sem_documentos(p):
                continue
            dados = base.get(chave, {})
            nome = dados.get("nome") or p.name
            achados.append({
                "chave": chave,
                "nome": nome,
                "termo": dados.get("termo") or termo_do_nome(p.name) or chave,
                "dre": dados.get("dre", ""),
                "valor": dados.get("valor", ""),
                "linha": dados.get("linha", 0),
            })

    elif criterio == "nf":
        for chave, dados in base.items():
            if dados.get("nf"):
                continue
            nome = dados.get("nome") or dados.get("termo") or chave
            achados.append({
                "chave": chave,
                "nome": nome,
                "termo": dados.get("termo") or termo_do_nome(nome) or chave,
                "dre": dados.get("dre", ""),
                "valor": dados.get("valor", ""),
                "linha": dados.get("linha", 0),
            })

    else:  # todos
        for chave, dados in base.items():
            nome = dados.get("nome") or dados.get("termo") or chave
            achados.append({
                "chave": chave,
                "nome": nome,
                "termo": dados.get("termo") or termo_do_nome(nome) or chave,
                "dre": dados.get("dre", ""),
                "valor": dados.get("valor", ""),
                "linha": dados.get("linha", 0),
            })

    alvo = texto_celula(dre_filtro)
    if alvo and normalizar(alvo) != normalizar(TODAS_DRES):
        alvo_n = normalizar(alvo)
        achados = [r for r in achados if normalizar(r.get("dre") or "") == alvo_n]

    achados.sort(key=lambda r: (normalizar(r.get("dre") or ""), normalizar(r["nome"])))
    if len(achados) > LIMITE_LISTA_PENDENTES:
        achados = achados[:LIMITE_LISTA_PENDENTES]
    for r in achados:
        r["emails"] = buscar_emails(r)
    return achados

def _nf_em_uso(base):
    """True quando a coluna de NF tem pelo menos uma linha preenchida."""
    try:
        for dados in base.values():
            if dados.get("nf"):
                return True
    except Exception:
        pass
    return False

def escolher_lista(base, desejado, dre_filtro=""):
    """
    O modo automático só usa 'pasta sem documentos'. Só desce para 'sem NF'
    se a coluna de NF estiver realmente em uso na planilha. Nunca cai para
    'todos', porque isso listaria contratados que não têm pendência.
    """
    if desejado == "auto":
        ordem = ["pasta"]
        if _nf_em_uso(base):
            ordem.append("nf")
    else:
        ordem = [desejado]

    for i, c in enumerate(ordem):
        try:
            lista = montar_pendentes(c, base, dre_filtro)
        except Exception as e:
            log(f"Lista ({c}): erro ({e})")
            continue
        if lista:
            return c, lista, (i > 0)
    return (ordem[0] if ordem else desejado), [], False

# ---------- CARGA AUTOMÁTICA DA LISTA ----------
def carregar_pendencias_async():
    global cob_carregando, cob_tentativas
    if lista_cob is None:
        return
    if cob_carregando:
        return
    cob_carregando = True
    cob_tentativas = 0
    try:
        lista_cob.delete(0, tk.END)
        pendentes_cob.clear()
        texto_cob.delete("1.0", tk.END)
        entrada_dest_cob.delete(0, tk.END)
    except Exception:
        pass
    dre = dre_escolhida()
    msg_cob("Carregando a lista de pendentes…"
            + (f" (DRE {dre})" if dre else ""), "#1565c0")
    desejado = "auto"
    try:
        desejado = var_criterio_cob.get()
    except Exception:
        desejado = cfg.get("criterio_cobranca", "auto")
    threading.Thread(target=_thread_pendencias, args=(desejado, dre),
                     daemon=True).start()

def _thread_pendencias(desejado, dre_filtro=""):
    if pythoncom is not None:
        try:
            pythoncom.CoInitialize()
        except Exception:
            pass
    resultado = {"ok": False, "erro": "", "criterio": "", "ajustado": False,
                 "lista": [], "base_n": 0, "base_info": "",
                 "dres": [], "dre": dre_filtro or "",
                 "email_ok": False, "email_info": ""}
    try:
        email_ok, email_info = carregar_emails_core()
        resultado["email_ok"] = email_ok
        resultado["email_info"] = email_info

        if not cache_pastas and pasta_principal:
            criar_cache(silencioso=True)

        base = ler_base_planilha()
        resultado["base_n"] = len(base)
        resultado["base_info"] = cache_base_info
        resultado["dres"] = dres_disponiveis(base)

        criterio, lista, ajustado = escolher_lista(base, desejado, dre_filtro)
        resultado["criterio"] = criterio
        resultado["ajustado"] = ajustado
        resultado["lista"] = lista
        resultado["ok"] = True
    except Exception as e:
        resultado["erro"] = str(e)
    finally:
        if pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
    fila_cob.put(resultado)

def poll_pendencias():
    global cob_poll_ativo
    try:
        while True:
            aplicar_resultado_pendencias(fila_cob.get_nowait())
    except queue.Empty:
        pass
    except Exception:
        pass
    try:
        janela.after(150, poll_pendencias)
    except Exception:
        cob_poll_ativo = False

def atualizar_combo_dres(dres, selecionada=""):
    """Preenche o combo de DRE sem disparar recarga."""
    if combo_dre_cob is None or var_dre_cob is None:
        return
    try:
        valores = [TODAS_DRES] + list(dres or [])
        combo_dre_cob.config(values=valores)
        atual = selecionada or dre_escolhida() or TODAS_DRES
        if normalizar(atual) not in [normalizar(v) for v in valores]:
            atual = TODAS_DRES
        var_dre_cob.set(atual)
        if label_dre_cob is not None:
            label_dre_cob.config(text=f"{max(0, len(valores) - 1)} DRE(s)")
    except Exception:
        pass

def aplicar_resultado_pendencias(r):
    global cob_carregando, cob_tentativas, pendentes_cob
    cob_carregando = False
    if lista_cob is None:
        return

    if not r.get("ok"):
        cob_tentativas += 1
        if cob_tentativas <= 3:
            msg_cob(f"Lendo a planilha… tentativa {cob_tentativas} de 3 "
                    f"({r.get('erro')})", "#b26a00")
            try:
                janela_cob.after(1500, carregar_pendencias_async)
            except Exception:
                pass
            return
        msg_cob(f"Não consegui carregar a lista: {r.get('erro')}", "#c62828")
        return

    cob_tentativas = 0
    lista = r.get("lista") or []
    criterio = r.get("criterio") or ""
    ajustado = bool(r.get("ajustado"))
    dre_usada = r.get("dre") or ""

    atualizar_combo_dres(r.get("dres") or [], dre_usada)

    pendentes_cob = list(lista)

    try:
        lista_cob.delete(0, tk.END)
        for item in pendentes_cob:
            lista_cob.insert(tk.END, linha_pendente(item))
    except Exception:
        return

    atualizar_label_fonte()

    sem_email = sum(1 for x in pendentes_cob if not x.get("emails"))
    sem_valor = sum(1 for x in pendentes_cob
                    if formatar_moeda(x.get("valor")) == "a confirmar")

    if not pendentes_cob:
        if not cfg.get("caminho_planilha"):
            msg_cob("Mapeie a planilha na tela principal (botão 'Planilha').",
                    "#c62828")
        elif criterio == "pasta":
            msg_cob("Nenhuma pasta sem documentos"
                    + (f' na DRE "{dre_usada}"' if dre_usada else "")
                    + f" · {len(cache_pastas)} pasta(s) indexada(s): "
                      f"todos os contratados com pasta têm arquivo.", "#2e7d32")
        elif dre_usada:
            msg_cob(f'Nenhum pendente na DRE "{dre_usada}" com esse critério. '
                    f'Escolha outra DRE ou "(todas)".', "#b26a00")
        else:
            msg_cob("Nenhum contratado encontrado. Confira aba e colunas em 'Fonte'.",
                    "#b26a00")
    else:
        partes = [f"{len(pendentes_cob)} na lista"]
        if dre_usada:
            partes.append(f"DRE: {dre_usada}")
        rotulo_crit = NOME_CRITERIO.get(criterio, criterio or "—")
        partes.append(("critério ajustado para " if ajustado else "critério: ")
                      + rotulo_crit)
        if r.get("base_info"):
            partes.append(r["base_info"])
        if sem_valor:
            partes.append(f"{sem_valor} sem valor na coluna Q")
        if cache_email_info:
            partes.append(cache_email_info)
        if sem_email:
            partes.append(f"{sem_email} sem e-mail")
        if not r.get("email_ok"):
            partes.append("e-mails: " + (r.get("email_info") or "falha")[:80])
        cor = "#c62828" if not r.get("email_ok") else (
            "#b26a00" if (sem_email or sem_valor) else "#2e7d32")
        msg_cob(" · ".join(partes), cor)

    if criterio and criterio in OPCOES_CRITERIO:
        try:
            var_criterio_cob.set(criterio)
            salvar_campos_cobranca()
        except Exception:
            pass

    log(f"Cobrança: {len(pendentes_cob)} item(ns) · critério {criterio}"
        + (f" · DRE {dre_usada}" if dre_usada else " · todas as DREs")
        + (f" · {sem_valor} sem valor" if sem_valor else "")
        + (f" · {sem_email} sem e-mail" if sem_email else ""))
    atualizar_preview_cobranca()

def atualizar_cobranca():
    carregar_pendencias_async()

def remetente_abertura(setor):
    d = (setor or SETOR_PADRAO).strip()
    minus = d.lower()
    for pref in ("o ", "a ", "os ", "as "):
        if minus.startswith(pref):
            return d
    return f"O {d}"

def bloco_contratado(r):
    return (
        f"Termo contratual: {r['termo']}\n"
        f"Contratado: {r['nome']}\n"
        f"Valor a ser pago: {formatar_moeda(r['valor'])}"
    )

def texto_cobranca(registros, setor, saudacao=None, assinar=False):
    """Texto puro. Com vários destinatários, a orientação vem ACIMA da relação."""
    if not registros:
        return ""
    setor = (setor or SETOR_PADRAO).strip()
    abertura = remetente_abertura(setor)
    cumprimento = resolver_saudacao(saudacao)

    lista = registros[:LIMITE_LISTA_EMAIL]
    varios = len(registros) > 1

    partes = [
        f"Prezado(os) (as), {cumprimento}!",
        f"{abertura} informa que está aguardando os documentos necessários para o "
        f"prosseguimento da liquidação e do pagamento.",
    ]

    if varios:
        partes.append(ORIENTACAO_VARIOS)
        corpo = ("Relação de contratados com documentos pendentes:\n\n"
                 + "\n\n".join(bloco_contratado(r) for r in lista))
        if len(registros) > len(lista):
            corpo += (f"\n\n… e mais {len(registros) - len(lista)} "
                      f"contratado(s) pendente(s).")
        partes.append(corpo)
    else:
        partes.append(bloco_contratado(lista[0]))

    partes.append("Solicitamos o envio dos documentos pendentes com a maior "
                  "brevidade possível.")
    partes.append("Atenciosamente,")
    if assinar:
        partes.append(setor)
    return "\n\n".join(partes)

def _esc_html(t):
    return (texto_celula(t).replace("&", "&amp;")
            .replace("<", "&lt;").replace(">", "&gt;"))

def html_cobranca(registros, setor, saudacao=None, assinar=False):
    """Versão HTML do e-mail. Único formato que aceita o grifo amarelo."""
    if not registros:
        return ""
    setor = (setor or SETOR_PADRAO).strip()
    abertura = remetente_abertura(setor)
    cumprimento = resolver_saudacao(saudacao)

    lista = registros[:LIMITE_LISTA_EMAIL]
    varios = len(registros) > 1
    p = 'style="margin:0 0 12px 0;"'

    partes = [
        '<html><head><meta charset="utf-8"></head>',
        '<body style="font-family:Calibri,Arial,sans-serif;font-size:11pt;'
        'color:#000000;">',
        f'<p {p}>Prezado(os) (as), {_esc_html(cumprimento)}!</p>',
        f'<p {p}>{_esc_html(abertura)} informa que está aguardando os documentos '
        f'necessários para o prosseguimento da liquidação e do pagamento.</p>',
    ]

    if varios:
        partes.append(
            f'<p {p}><span style="background-color:{COR_GRIFO};'
            f'mso-highlight:yellow;">{_esc_html(ORIENTACAO_VARIOS)}</span></p>'
        )
        partes.append(f'<p {p}><b>Relação de contratados com documentos '
                      f'pendentes:</b></p>')
        for r in lista:
            partes.append(
                f'<p {p}>Termo contratual: {_esc_html(r["termo"])}<br>'
                f'Contratado: {_esc_html(r["nome"])}<br>'
                f'Valor a ser pago: {_esc_html(formatar_moeda(r["valor"]))}</p>'
            )
        if len(registros) > len(lista):
            partes.append(
                f'<p {p}>… e mais {len(registros) - len(lista)} '
                f'contratado(s) pendente(s).</p>'
            )
    else:
        r = lista[0]
        partes.append(
            f'<p {p}>Termo contratual: {_esc_html(r["termo"])}<br>'
            f'Contratado: {_esc_html(r["nome"])}<br>'
            f'Valor a ser pago: {_esc_html(formatar_moeda(r["valor"]))}</p>'
        )

    partes.append(f'<p {p}>Solicitamos o envio dos documentos pendentes com a '
                  f'maior brevidade possível.</p>')
    partes.append(f'<p {p}>Atenciosamente,</p>')
    if assinar:
        partes.append(f'<p {p}>{_esc_html(setor)}</p>')
    partes.append('</body></html>')
    return "\n".join(partes)

def salvar_campos_cobranca():
    try:
        cfg["departamento"] = entrada_setor_cob.get().strip()
        cfg["assunto_cobranca"] = entrada_assunto_cob.get().strip()
        cfg["criterio_cobranca"] = var_criterio_cob.get()
        cfg["saudacao_cobranca"] = saudacao_escolhida()
        cfg["assinar_setor"] = assinar_escolhido()
        cfg["remetente_email"] = remetente_escolhido()
        salvar_config()
    except Exception:
        pass

def abrir_pasta_pendente():
    """Abre a pasta do contratado selecionado e informa o que o programa viu."""
    if lista_cob is None:
        return
    sel = lista_cob.curselection()
    if not sel or sel[0] >= len(pendentes_cob):
        msg_cob("Selecione um contratado na lista para abrir a pasta.", "#b26a00")
        return
    r = pendentes_cob[sel[0]]
    chave = r.get("chave") or ""
    destino = None
    for p in cache_pastas.values():
        if chave and chave_contrato(p.name) == chave:
            destino = p
            break
    if destino is None:
        msg_cob("Este contratado não tem pasta indexada. "
                "Use 'Principal' para reindexar.", "#b26a00")
        return

    achados, erros = _arquivos_da_pasta(destino)
    try:
        os.startfile(str(destino))
    except Exception as e:
        msg_cob(f"Não consegui abrir: {e}", "#c62828")
        return

    if achados:
        amostra = ", ".join(p.name for p in achados[:3])
        msg_cob(f"{len(achados)} arquivo(s) em {destino.name} · ex.: {amostra}",
                "#2e7d32")
    else:
        msg_cob(f"Nenhum arquivo em {destino.name}"
                + (f" · {erros} subpasta(s) ilegível(is)" if erros else ""),
                "#b26a00")

def registros_selecionados():
    if lista_cob is None:
        return []
    sel = lista_cob.curselection()
    if not sel:
        return list(pendentes_cob)
    return [pendentes_cob[i] for i in sel if i < len(pendentes_cob)]

def emails_selecionados(lista_registros=None):
    registros = lista_registros if lista_registros is not None else registros_selecionados()
    vistos = []
    for r in registros:
        for e in (r.get("emails") or []):
            if e not in vistos:
                vistos.append(e)
    return vistos

def atualizar_destinatarios():
    if entrada_dest_cob is None:
        return
    emails = emails_selecionados()
    set_entry(entrada_dest_cob, "; ".join(emails))

def copiar_destinatarios():
    if entrada_dest_cob is None:
        return
    texto = entrada_dest_cob.get().strip()
    if not texto:
        msg_cob("Nenhum e-mail selecionado para copiar.", "#b26a00")
        return
    try:
        janela.clipboard_clear()
        janela.clipboard_append(texto)
        msg_cob("Destinatários copiados. Cole no campo Para do Outlook.", "#2e7d32")
    except Exception as e:
        msg_cob(f"Não consegui copiar: {e}", "#c62828")

def atualizar_preview_cobranca():
    if texto_cob is None:
        return
    try:
        setor = entrada_setor_cob.get().strip() or SETOR_PADRAO
    except Exception:
        setor = SETOR_PADRAO
    corpo = texto_cobranca(registros_selecionados(), setor,
                           saudacao_escolhida(), assinar_escolhido())
    texto_cob.delete("1.0", tk.END)
    texto_cob.insert("1.0", corpo)
    try:
        if ORIENTACAO_VARIOS in corpo:
            pos = texto_cob.search(ORIENTACAO_VARIOS, "1.0", tk.END)
            if pos:
                texto_cob.tag_add("grifo", pos,
                                  f"{pos}+{len(ORIENTACAO_VARIOS)}c")
    except Exception:
        pass
    atualizar_destinatarios()
    try:
        agora = datetime.datetime.now()
        label_hora_cob.config(
            text=f"agora: {saudacao_por_hora(agora)} · {agora:%H:%M}")
    except Exception:
        pass

def linha_pendente(r):
    emails = r.get("emails") or []
    email_txt = "; ".join(emails) if emails else "— sem e-mail —"
    return (f"{resumir(r.get('dre') or '-', 8):<8} | "
            f"{resumir(r['nome'], 24):<24} | "
            f"{r['termo']:<13} | "
            f"{formatar_moeda(r['valor']):<14} | "
            f"{resumir(email_txt, 40)}")

def copiar_cobranca():
    if texto_cob is None:
        return
    corpo = texto_cob.get("1.0", tk.END).strip()
    if not corpo:
        return
    try:
        janela.clipboard_clear()
        janela.clipboard_append(corpo)
        msg_cob("Texto copiado (sem cor — a área de transferência é texto puro).",
                "#2e7d32")
    except Exception as e:
        msg_cob(f"Não consegui copiar: {e}", "#c62828")

def rascunho_cobranca():
    if texto_cob is None:
        return
    registros = registros_selecionados()
    if not registros:
        msg_cob("Nada para gerar: nenhum pendente carregado.", "#b26a00")
        return
    try:
        setor = entrada_setor_cob.get().strip() or SETOR_PADRAO
    except Exception:
        setor = SETOR_PADRAO
    corpo = texto_cobranca(registros, setor, saudacao_escolhida(),
                           assinar_escolhido())
    html = html_cobranca(registros, setor, saudacao_escolhida(),
                         assinar_escolhido())
    atualizar_preview_cobranca()
    destinatarios = []
    if entrada_dest_cob is not None:
        destinatarios = emails_validos(entrada_dest_cob.get())
    assunto = entrada_assunto_cob.get().strip() or "Documentos pendentes"
    remetente = remetente_escolhido()
    ok, detalhe = criar_rascunho_outlook(assunto, corpo, destinatarios, html,
                                         remetente, "auto")
    if ok:
        extra = (f"{len(destinatarios)} destinatário(s) preenchido(s)."
                 if destinatarios else "Sem destinatário — use o campo na tela.")
        msg_cob(f"Rascunho criado em {detalhe}. {extra}", "#2e7d32")
        log(f"Cobrança: rascunho criado em {detalhe} "
            f"({len(destinatarios)} destinatário(s))")
    else:
        copiar_cobranca()
        msg_cob(f"Outlook indisponível ({detalhe}). Texto copiado — cole no seu e-mail.",
                "#b26a00")
        log(f"Cobrança: Outlook indisponível ({detalhe})")

def fechar_cobranca():
    global janela_cob
    salvar_campos_cobranca()
    try:
        janela_cob.destroy()
    except Exception:
        pass
    janela_cob = None

def montar_janela_cobranca():
    global janela_cob, lista_cob, texto_cob, status_cob, label_hora_cob
    global label_fonte_cob, var_criterio_cob, entrada_setor_cob
    global entrada_assunto_cob, var_saudacao_cob, var_assinar_cob
    global entrada_dest_cob, var_dre_cob, combo_dre_cob, label_dre_cob
    global var_remetente_cob, combo_remetente_cob, label_remetente_cob

    janela_cob = tk.Toplevel(janela)
    janela_cob.title("Cobrança de documentos")
    janela_cob.geometry("880x870")
    janela_cob.transient(janela)

    tk.Label(janela_cob, text="Cobrança de documentos pendentes",
             font=("Arial", 11, "bold"), anchor="w").pack(
        fill="x", padx=10, pady=(10, 2))

    linha = tk.Frame(janela_cob)
    linha.pack(fill="x", padx=10)
    tk.Label(linha, text="Critério:", font=("Arial", 8, "bold")).pack(side="left")
    var_criterio_cob = tk.StringVar(value=cfg.get("criterio_cobranca", "auto"))
    for valor, rotulo in (("auto", "Automático"),
                          ("pasta", "Pasta sem documentos"),
                          ("nf", "Sem NF"),
                          ("todos", "Todos")):
        tk.Radiobutton(linha, text=rotulo, variable=var_criterio_cob,
                       value=valor, font=("Arial", 8),
                       command=carregar_pendencias_async).pack(side="left", padx=(6, 0))

    # ---------- FILTRO POR DRE ----------
    dre_frame = tk.Frame(janela_cob)
    dre_frame.pack(fill="x", padx=10, pady=(6, 0))
    tk.Label(dre_frame, text="DRE:", font=("Arial", 8, "bold"),
             fg="#1565c0").pack(side="left")
    var_dre_cob = tk.StringVar(value=TODAS_DRES)
    combo_dre_cob = ttk.Combobox(dre_frame, textvariable=var_dre_cob,
                                 state="readonly", font=("Arial", 9),
                                 width=30, values=[TODAS_DRES])
    combo_dre_cob.pack(side="left", padx=(6, 6))
    combo_dre_cob.bind("<<ComboboxSelected>>",
                       lambda e: carregar_pendencias_async())
    tk.Button(dre_frame, text="Atualizar DREs", font=("Arial", 7), padx=4, pady=0,
              command=carregar_pendencias_async).pack(side="left")
    label_dre_cob = tk.Label(dre_frame, text="0 DRE(s)", font=("Arial", 7),
                             fg="#777777")
    label_dre_cob.pack(side="right")

    # ---------- REMETENTE ----------
    rem_frame = tk.Frame(janela_cob)
    rem_frame.pack(fill="x", padx=10, pady=(6, 0))
    tk.Label(rem_frame, text="Enviar por:", font=("Arial", 8, "bold"),
             fg="#1565c0").pack(side="left")
    var_remetente_cob = tk.StringVar(value=CONTA_PADRAO)
    combo_remetente_cob = ttk.Combobox(rem_frame, textvariable=var_remetente_cob,
                                       state="normal", font=("Arial", 9),
                                       width=34, values=[CONTA_PADRAO])
    combo_remetente_cob.pack(side="left", padx=(6, 4))
    tk.Button(rem_frame, text="Contas", font=("Arial", 7), padx=4, pady=0,
              command=lambda: atualizar_combo_remetente(True)).pack(side="left")
    label_remetente_cob = tk.Label(rem_frame, text="", font=("Arial", 7),
                                   fg="#777777")
    label_remetente_cob.pack(side="right")

    tk.Label(janela_cob,
             text='Deixe em "(conta padrão do Outlook)" para enviar pela conta '
                  'principal. Para caixa compartilhada, digite o nome exato que '
                  'aparece no Outlook.',
             font=("Arial", 7), fg="#777777", anchor="w", justify="left",
             wraplength=840).pack(fill="x", padx=10)

    setor_frame = tk.Frame(janela_cob)
    setor_frame.pack(fill="x", padx=10, pady=(6, 0))
    tk.Label(setor_frame, text="Setor / remetente:",
             font=("Arial", 8, "bold")).pack(side="left")
    entrada_setor_cob = tk.Entry(setor_frame, font=("Arial", 9))
    entrada_setor_cob.insert(0, cfg.get("departamento", SETOR_PADRAO))
    entrada_setor_cob.pack(side="left", fill="x", expand=True, padx=(6, 0))
    entrada_setor_cob.bind("<KeyRelease>", lambda e: atualizar_preview_cobranca())

    saud = tk.Frame(janela_cob)
    saud.pack(fill="x", padx=10, pady=(4, 0))
    tk.Label(saud, text="Saudação:", font=("Arial", 8, "bold")).pack(side="left")
    var_saudacao_cob = tk.StringVar(value=cfg.get("saudacao_cobranca", SAUDACAO_AUTO))
    menu_saud = tk.OptionMenu(saud, var_saudacao_cob, *OPCOES_SAUDACAO,
                              command=lambda v: atualizar_preview_cobranca())
    menu_saud.config(font=("Arial", 8), width=11)
    menu_saud.pack(side="left", padx=(6, 10))
    var_assinar_cob = tk.BooleanVar(value=bool(cfg.get("assinar_setor", False)))
    tk.Checkbutton(saud, text="Assinar com o setor",
                   variable=var_assinar_cob, font=("Arial", 8),
                   command=atualizar_preview_cobranca).pack(side="left")
    label_hora_cob = tk.Label(saud, text="", font=("Arial", 7), fg="#777777")
    label_hora_cob.pack(side="right")

    ass = tk.Frame(janela_cob)
    ass.pack(fill="x", padx=10, pady=(4, 0))
    tk.Label(ass, text="Assunto:", font=("Arial", 8, "bold")).pack(side="left")
    entrada_assunto_cob = tk.Entry(ass, font=("Arial", 9))
    entrada_assunto_cob.insert(
        0, cfg.get("assunto_cobranca", "Documentos pendentes – Termo Contratual"))
    entrada_assunto_cob.pack(side="left", fill="x", expand=True, padx=(6, 0))

    fonte_frame = tk.Frame(janela_cob)
    fonte_frame.pack(fill="x", padx=10, pady=(6, 0))
    tk.Label(fonte_frame, text="E-mails:", font=("Arial", 8, "bold")).pack(side="left")
    label_fonte_cob = tk.Label(fonte_frame, text="não carregado",
                               font=("Arial", 7), fg="#b26a00", anchor="w")
    label_fonte_cob.pack(side="left", fill="x", expand=True, padx=(6, 4))
    tk.Button(fonte_frame, text="Fonte", font=("Arial", 7), padx=4, pady=0,
              command=abrir_config_emails).pack(side="left", padx=1)
    tk.Button(fonte_frame, text="Diagnóstico", font=("Arial", 7), padx=4, pady=0,
              command=mostrar_diagnostico_emails).pack(side="left", padx=1)

    dest = tk.Frame(janela_cob)
    dest.pack(fill="x", padx=10, pady=(6, 0))
    tk.Label(dest, text="Destinatários:", font=("Arial", 8, "bold"),
             fg="#1565c0").pack(side="left")
    entrada_dest_cob = tk.Entry(dest, font=("Consolas", 9), fg="#1565c0")
    entrada_dest_cob.pack(side="left", fill="x", expand=True, padx=(6, 4))
    tk.Button(dest, text="Copiar destinatários", font=("Arial", 7), padx=4, pady=0,
              command=copiar_destinatarios).pack(side="left")

    cab_lista = tk.Frame(janela_cob)
    cab_lista.pack(fill="x", padx=10, pady=(8, 0))
    tk.Label(cab_lista,
             text="Pendentes  (DRE | nome | termo | valor | e-mail)  ·  "
                  "duplo clique abre a pasta:",
             font=("Arial", 8, "bold"), anchor="w").pack(side="left")
    tk.Button(cab_lista, text="Carregar lista", font=("Arial", 7), padx=6, pady=0,
              bg="#1565c0", fg="white",
              command=carregar_pendencias_async).pack(side="right")

    lista_cob = tk.Listbox(janela_cob, height=8, font=("Consolas", 8),
                           selectmode=tk.EXTENDED, activestyle="dotbox")
    lista_cob.pack(fill="x", padx=10)
    lista_cob.bind("<<ListboxSelect>>",
                   lambda e: atualizar_preview_cobranca())
    lista_cob.bind("<Double-Button-1>",
                   lambda e: abrir_pasta_pendente())

    tk.Label(janela_cob, text="Texto do e-mail  (o grifo amarelo aparece no Outlook):",
             font=("Arial", 8, "bold"), anchor="w").pack(
        fill="x", padx=10, pady=(8, 0))
    texto_cob = tk.Text(janela_cob, height=9, font=("Consolas", 9), wrap="word")
    texto_cob.tag_configure("grifo", background=COR_GRIFO)
    texto_cob.pack(fill="both", expand=True, padx=10)

    botoes = tk.Frame(janela_cob)
    botoes.pack(fill="x", padx=10, pady=8)
    tk.Button(botoes, text="Copiar texto", command=copiar_cobranca,
              width=13).pack(side="left", padx=2)
    tk.Button(botoes, text="Abrir pasta", command=abrir_pasta_pendente,
              width=11).pack(side="left", padx=2)
    tk.Button(botoes, text="Rascunho no Outlook", command=rascunho_cobranca,
              width=19, bg="#1565c0", fg="white").pack(side="left", padx=2)
    tk.Button(botoes, text="Fechar", command=fechar_cobranca,
              width=8).pack(side="right", padx=2)

    status_cob = tk.Label(janela_cob, text="", anchor="w", font=("Arial", 8),
                          fg="#555555", wraplength=850, justify="left")
    status_cob.pack(fill="x", padx=10, pady=(0, 8))

    janela_cob.protocol("WM_DELETE_WINDOW", fechar_cobranca)
    atualizar_combo_remetente(False)

def abrir_cobranca():
    global janela_cob
    if janela_cob is not None:
        try:
            if janela_cob.winfo_exists():
                janela_cob.deiconify()
                janela_cob.lift()
                atualizar_label_fonte()
                atualizar_combo_remetente(False)
                carregar_pendencias_async()
                return
        except Exception:
            pass
        janela_cob = None
    montar_janela_cobranca()
    atualizar_label_fonte()
    try:
        janela_cob.after(80, carregar_pendencias_async)
    except Exception:
        carregar_pendencias_async()

# =========================
# JANELA
# =========================
janela = tk.Tk()
janela.title("Extrator Automático")
janela.geometry("470x565")
janela.resizable(False, False)
janela.attributes("-topmost", True)
janela.bind("<Escape>", limpar_contrato)
janela.bind("<F5>", lambda e: iniciar_consulta_q())
janela.bind("<F6>", lambda e: atualizar_dre(True))

topo = tk.Frame(janela)
topo.pack(fill="x", padx=8, pady=(8, 2))
tk.Button(topo, text="Principal", command=mapear_principal,
          width=10).pack(side="left", padx=2)
tk.Button(topo, text="Downloads", command=mapear_download,
          width=11).pack(side="left", padx=2)
tk.Button(topo, text="Planilha", command=mapear_planilha,
          width=10).pack(side="left", padx=2)
botao_toggle = tk.Button(topo, text="Iniciar", command=alternar,
                         bg="#2e7d32", fg="white", width=8)
botao_toggle.pack(side="left", padx=2)

linha_cobranca = tk.Frame(janela)
linha_cobranca.pack(fill="x", padx=8, pady=(4, 0))
tk.Button(linha_cobranca, text="Cobrar documentos pendentes",
          command=abrir_cobranca, bg="#1565c0", fg="white",
          font=("Arial", 9, "bold")).pack(side="left", fill="x", expand=True)
tk.Button(linha_cobranca, text="Pastas vazias",
          command=abrir_pastas_vazias,
          font=("Arial", 8), padx=6).pack(side="left", padx=(4, 0))

estado_label = tk.Label(janela, text="", anchor="w", font=("Consolas", 8),
                        bg="#f0f0f0", relief="groove", bd=1)
estado_label.pack(fill="x", padx=8, pady=3)

cab_busca = tk.Frame(janela)
cab_busca.pack(fill="x", padx=8)
tk.Label(cab_busca, text="Contrato, nome ou sigla de DRE (coluna B)",
         font=("Arial", 8, "bold"), anchor="w").pack(side="left")
tk.Button(cab_busca, text="DRE F6", command=lambda: atualizar_dre(True),
          font=("Arial", 7), padx=4, pady=0).pack(side="right")

entrada_busca = tk.Entry(janela, justify="center", font=("Arial", 11))
entrada_busca.pack(fill="x", padx=8)
entrada_busca.bind("<KeyRelease>", pesquisar)
entrada_busca.bind("<Return>", confirmar)
entrada_busca.bind("<Down>", tecla_baixo)
entrada_busca.bind("<Up>", tecla_cima)

lista = tk.Listbox(janela, height=3, font=("Arial", 8), activestyle="dotbox")
lista.pack(fill="x", padx=8)
lista.bind("<ButtonRelease-1>", clicar_lista)

preview_label = tk.Label(janela, text="", anchor="w", font=("Arial", 8),
                         fg="#777777")
preview_label.pack(fill="x", padx=8)

contrato_label = tk.Label(janela, text="", anchor="w", justify="left",
                          font=("Arial", 8, "bold"), fg="#2e7d32")
contrato_label.pack(fill="x", padx=8, pady=(3, 0))

# ---------- COLUNA Q ----------
confere_frame = tk.Frame(janela, bg="#f5f5f5", relief="solid", bd=2)
confere_frame.pack(fill="x", padx=8, pady=(5, 0))
confere_titulo = tk.Label(confere_frame, text="VALOR DO CONTRATO (COLUNA Q)",
                          font=("Arial", 7, "bold"), bg="#f5f5f5",
                          fg="#777777", anchor="w")
confere_titulo.pack(fill="x", padx=6, pady=(3, 0))
confere_label = tk.Label(confere_frame, text="—", font=("Arial", 17, "bold"),
                         bg="#f5f5f5", fg="#777777", anchor="center")
confere_label.pack(fill="x")
confere_info = tk.Label(confere_frame,
                        text="Selecione o contrato para ver o valor na coluna Q",
                        font=("Arial", 7), bg="#f5f5f5",
                        fg="#777777", anchor="w")
confere_info.pack(fill="x", padx=6, pady=(0, 3))
for widget in (confere_frame, confere_titulo, confere_label, confere_info):
    widget.bind("<Button-1>", lambda e: iniciar_consulta_q())

campos = tk.Frame(janela)
campos.pack(fill="x", padx=8, pady=(6, 0))
tk.Label(campos, text="NF", font=("Arial", 9, "bold")).grid(row=0, column=0)
entrada_nf = tk.Entry(campos, justify="center", width=18, font=("Arial", 11))
entrada_nf.grid(row=0, column=1, padx=(4, 10))
entrada_nf.bind("<Return>", lambda e: entrada_dia.focus_set())
tk.Label(campos, text="Dia", font=("Arial", 9, "bold")).grid(row=0, column=2)
entrada_dia = tk.Entry(campos, justify="center", width=6, font=("Arial", 11))
entrada_dia.grid(row=0, column=3, padx=(4, 10))
entrada_dia.bind("<Return>", salvar_excel)
tk.Label(campos, text="Enter salva", font=("Arial", 8, "italic"),
         fg="#2e7d32").grid(row=0, column=4)

msg_label = tk.Label(janela, text="", anchor="w", justify="left",
                     font=("Arial", 8), fg="#555555", wraplength=450)
msg_label.pack(fill="x", padx=8, pady=(4, 0))

acao_label = tk.Label(janela, text="", anchor="w", font=("Consolas", 7),
                      fg="#888888")
acao_label.pack(fill="x", padx=8, pady=(2, 6))

# COM na thread da interface
if pythoncom is not None:
    try:
        pythoncom.CoInitialize()
    except Exception:
        pass

# restaura configuração
if cfg["pasta_principal"] and Path(cfg["pasta_principal"]).is_dir():
    pasta_principal = Path(cfg["pasta_principal"])
    criar_cache()
if cfg["pasta_download"] and Path(cfg["pasta_download"]).is_dir():
    pasta_download = Path(cfg["pasta_download"])

restaurou_contrato = False
if cfg.get("pasta_destino") and Path(cfg["pasta_destino"]).is_dir():
    pasta_destino_global = Path(cfg["pasta_destino"])
    restaurou_contrato = True
    log(f"Contrato restaurado: {pasta_destino_global.name}")

threading.Thread(target=worker, daemon=True).start()
atualizar_contrato_label()
painel()
habilitar_campos(False)
aviso("Busque por contrato, nome ou sigla de DRE (coluna B).", "#555555")
janela.after(150, poll_q)
janela.after(200, poll_pendencias)
if cfg.get("caminho_planilha"):
    janela.after(300, lambda: atualizar_dre(True))
if restaurou_contrato:
    janela.after(500, lambda: iniciar_consulta_q(pasta_destino_global))
if pasta_download:
    janela.after(400, iniciar)
entrada_busca.focus_set()
janela.mainloop()
