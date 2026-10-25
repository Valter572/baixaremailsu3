"""
Extrator Automatico - v19
=========================
Base: v18 (correcao da movimentacao de arquivos).

O que mudou na v19 (necessario para rodar como .exe):
  1) Havia DUAS declaracoes de CONFIG_PATH. A segunda, logo abaixo da
     correcao, sobrescrevia a primeira e apontava para a pasta temporaria
     do PyInstaller. Agora existe uma unica declaracao.
  2) BASE_DIR usa a pasta do .exe. Se ela nao for gravavel (ex.: a pasta
     Program Files), cai para %APPDATA%\\ExtratorAutomatico.
  3) print() nao derruba mais o programa quando compilado com --noconsole.
  4) log() tambem grava em extrator_log.txt, para diagnostico sem console.
  5) CoInitialize na thread da interface: sem isso o GetActiveObject do
     Excel pode falhar de forma intermitente.
  6) Extracao de ZIP com verificacao de caminho mais segura.
  7) Conversao imagem -> PDF usa os quadros ja convertidos.

Colunas: A(1)=termo | U(21)=nome+termo | O(15)=dia | P(16)=NF | Q(17)=valor
"""
import tkinter as tk
from tkinter import filedialog
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
CONFIG_PADRAO = {
    "pasta_principal": "",
    "pasta_download": "",
    "pasta_destino": "",
    "caminho_planilha": "",
    "nome_aba": "LIQUIDACAO",
    "col_termo": 1,
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
}


def carregar_config():
    cfg = dict(CONFIG_PADRAO)
    try:
        if CONFIG_PATH.exists():
            cfg.update(json.loads(CONFIG_PATH.read_text(encoding="utf-8")))
    except Exception:
        pass
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
TOLERANCIA_ANTIGOS_S = 0.5
INICIO_PROGRAMA = time.time()

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

# =========================
# EXCEL
# =========================
def obter_aba(wb):
    try:
        return wb.Worksheets(cfg["nome_aba"])
    except Exception:
        return wb.ActiveSheet


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
    for chave in ("col_termo", "col_alt", "col_conferencia", "col_dia", "col_nf"):
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
            mostrar_q("—", f"Linha {linha} · coluna Q está vazia")
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
def atualizar_contrato_label():
    dest = pasta_destino_atual()
    if dest:
        termo = termo_do_nome(dest.name)
        extra = f"  ({termo})" if termo else ""
        texto, cor = f"Contrato ativo: {resumir(dest.name, 40)}{extra}", "#2e7d32"
    else:
        texto, cor = "Sem contrato ativo — o arquivo aguarda em Downloads", "#e65100"
    try:
        janela.after(0, lambda: contrato_label.config(text=texto, fg=cor))
    except Exception:
        pass


def atualizar_preview():
    if not resultados_atuais:
        preview_label.config(text="", fg="#777777")
        return
    sel = lista.curselection()
    idx = sel[0] if sel else 0
    if idx >= len(resultados_atuais):
        idx = 0
    nome = resultados_atuais[idx].name
    total = len(resultados_atuais)
    if total == 1:
        preview_label.config(text=f"Enter aplica: {resumir(nome)}", fg="#2e7d32")
    else:
        preview_label.config(
            text=f"Enter aplica ({idx + 1}/{total}): {resumir(nome)}", fg="#1565c0")

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
# PASTAS VAZIAS
# =========================
def contar_pastas_vazias():
    total = 0
    for p in list(cache_pastas.values()):
        try:
            vazia = True
            with os.scandir(p) as it:
                for e in it:
                    if e.name.startswith("."):
                        continue
                    try:
                        if e.is_file():
                            vazia = False
                            break
                    except OSError:
                        continue
            if vazia:
                total += 1
        except OSError:
            continue
    return total


def recontar_vazias():
    global pastas_vazias
    pastas_vazias = contar_pastas_vazias()
    return pastas_vazias

# =========================
# CACHE DE PASTAS
# =========================
def criar_cache():
    cache_pastas.clear()
    if not pasta_principal or not pasta_principal.is_dir():
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
                aviso(f"Planilha: {wb.Name}", "#2e7d32")
                if pasta_destino_atual():
                    iniciar_consulta_q()
                return
    except Exception:
        pass
    a = escolher_arquivo("Selecione a planilha",
                         [("Excel", "*.xlsx *.xlsm *.xls")])
    if a:
        cfg["caminho_planilha"] = a
        salvar_config()
        aviso(f"Planilha: {Path(a).name}", "#2e7d32")
        if pasta_destino_atual():
            iniciar_consulta_q()

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
    except Exception as e:
        aviso("Não consegui abrir a planilha.")
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
    """Arquivos de Downloads, incluindo subpastas (até 2 níveis)."""
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
    """Retorna quantos arquivos estão aguardando contrato."""
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
    achados = []
    for p in cache_pastas.values():
        nome = normalizar(p.name)
        bate = termo in nome
        if not bate and len(termo_d) >= 5 and termo_d in digitos(p.name):
            bate = True
        if bate:
            achados.append(p)
    achados.sort(key=lambda x: x.name)
    resultados_atuais = achados
    for p in achados:
        lista.insert(tk.END, p.name)
    if not achados:
        aviso("Nenhuma pasta encontrada para esse termo.")
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
        aviso("Digite parte do contrato para localizar a pasta.")
        return
    sel = lista.curselection()
    idx = sel[0] if sel else 0
    if idx >= len(resultados_atuais):
        idx = 0
    escolher_contrato(resultados_atuais[idx])


def clicar_lista(event=None):
    sel = lista.curselection()
    if not sel or sel[0] >= len(resultados_atuais):
        return
    escolher_contrato(resultados_atuais[sel[0]])


def escolher_contrato(pasta):
    global pasta_destino_global, ultimo_movimento, snapshot_campos
    pasta_destino_global = pasta
    cfg["pasta_destino"] = str(pasta)
    salvar_config()
    ultimo_movimento = 0.0
    snapshot_campos = None
    termo = termo_do_nome(pasta.name)
    atualizar_contrato_label()
    aviso(f"Contrato: {termo or pasta.name} — lendo a coluna Q…", "#1565c0")
    log(f"Contrato ativo: {pasta.name}")
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

# =========================
# JANELA
# =========================
janela = tk.Tk()
janela.title("Extrator Automático")
janela.geometry("470x470")
janela.resizable(False, False)
janela.attributes("-topmost", True)
janela.bind("<Escape>", limpar_contrato)
janela.bind("<F5>", lambda e: iniciar_consulta_q())

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

estado_label = tk.Label(janela, text="", anchor="w", font=("Consolas", 8),
                        bg="#f0f0f0", relief="groove", bd=1)
estado_label.pack(fill="x", padx=8, pady=3)

tk.Label(janela, text="Contrato (parte do número ou do nome)",
         font=("Arial", 8, "bold"), anchor="w").pack(fill="x", padx=8)
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

# COM na thread da interface: necessario para GetActiveObject do Excel
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
aviso("Selecione o contrato para liberar a NF e o Dia.", "#555555")
janela.after(150, poll_q)
if restaurou_contrato:
    janela.after(500, lambda: iniciar_consulta_q(pasta_destino_global))
if pasta_download:
    janela.after(400, iniciar)
entrada_busca.focus_set()
janela.mainloop()
