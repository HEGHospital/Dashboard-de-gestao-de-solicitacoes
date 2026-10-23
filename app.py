import streamlit as st
import pandas as pd
import requests
import copy
import threading
import time
from datetime import datetime

st.set_page_config(page_title="Dashboard de Gestão de Solicitações", layout="wide")

# CSS para os cards de status (fonte maior, cores por status)
st.markdown("""
<style>
.status-card {
    border-radius: 10px;
    padding: 18px 12px;
    text-align: center;
}
.status-card .icone { font-size: 30px; line-height: 1; }
.status-card .rotulo {
    font-size: 16px;
    font-weight: 600;
    margin-top: 6px;
    color: #333;
}
.status-card .valor {
    font-size: 34px;
    font-weight: 800;
    margin-top: 2px;
}
</style>
""", unsafe_allow_html=True)

# ------------------------------------------------------------------
# CONFIGURAÇÃO DA API
# ------------------------------------------------------------------
API_URL = "https://heg.sysepa.com.br/epa/api/api/os"
API_TOKEN = "EPA@39Pd3ha29g"

HEADERS = {
    "Accept": "application/json",
    "Authorization": f"Bearer {API_TOKEN}",
}

# Sessão HTTP reaproveitada entre as chamadas — evita o custo de abrir uma
# conexão TCP/TLS nova a cada requisição (pequeno ganho, mas ajuda).
SESSAO_HTTP = requests.Session()
SESSAO_HTTP.headers.update(HEADERS)

# A cada quantos segundos a thread em segundo plano busca dados novos na API
INTERVALO_ATUALIZACAO_SEGUNDOS = 60

# ------------------------------------------------------------------
# DE-PARA (código -> descrição)
# ------------------------------------------------------------------
# Cada entrada é uma tabela de-para: um arquivo Excel com uma coluna de
# código e uma coluna de descrição, aplicada a uma ou mais colunas da
# tabela principal. Para adicionar um novo de-para no futuro (ex: status,
# equipamento), basta incluir mais uma entrada neste dicionário.
MAPAS_DE_PARA = {
    "usuarios": {
        "arquivo": "EPA - Usuários.xlsx",  # deixe este arquivo na mesma pasta do app.py
        "coluna_codigo": "Código",
        "coluna_descricao": "Nome",
        "colunas_tabela_principal": [
            "responsavel",
            "cliente",
            "usuario_solicitante",
            "usuario_conclusao",
            "ultimo_encaminhamento",
            "usuario_feedback",
        ],
    },
    "unidade_gerencial": {
        "arquivo": "EPA - Unidade Gerencial.xlsx",  # deixe este arquivo na mesma pasta do app.py
        "coluna_codigo": "Código",
        "coluna_descricao": "Unidade",
        "colunas_tabela_principal": [
            "unidade_gerencial",
            "unidade_gerencial_executora",
        ],
    },
}

# ------------------------------------------------------------------
# RELAÇÕES (with) — descrições que já vêm dentro do JSON da API
# ------------------------------------------------------------------
# Algumas colunas, em vez de precisarem de um Excel de-para, já têm a
# descrição disponível na própria API através do parâmetro "with"
# (?with=r_os_tipo). Nesse caso a API devolve, junto de cada registro, um
# objeto aninhado com a descrição pronta — não precisa de planilha.
# Mapeie aqui: nome do campo de código na tabela principal -> nome da
# relação que a API devolve (a chave do objeto aninhado).
RELACOES_WITH = {
    "os_tipo_servico": "r_os_tipo_servico",
}

# Relações pedidas via with que NÃO são um objeto único (descrição de um
# código), e sim uma LISTA dentro de cada solicitação — ex: uma solicitação
# pode ter vários encaminhamentos. Essas viram uma tabela própria (df_enc),
# nunca uma coluna da tabela principal (senão duplicaria os dados da
# solicitação a cada encaminhamento).
WITH_LISTAS = ["encaminhamentos"]

# ------------------------------------------------------------------
# STATUS — ícone e cor de cada card
# ------------------------------------------------------------------
STATUS_CONFIG = {
    "Em Planejamento": {"icone": "📋", "cor": "#3B82F6"},
    "Em Andamento": {"icone": "🚧", "cor": "#F59E0B"},
    "Suspenso": {"icone": "⏸️", "cor": "#6B7280"},
    "Concluido": {"icone": "✅", "cor": "#10B981"},
    "Cancelado": {"icone": "❌", "cor": "#EF4444"},
}

# Só esses status entram na fila de encaminhamentos (quem está com a solicitação agora)
STATUS_FILA_ABERTA = ("Em Andamento", "Em Planejamento")

# Colunas mostradas na tabela detalhada (agrupada por solicitante)
COLUNAS_TABELA_DETALHADA = [
    "codigo",
    "titulo",
    "unidade_gerencial",
    "data_inclusao",
    "observacao",
    "status",
]

# Colunas mostradas na tabela de encaminhamentos (agrupada por usuario_para)
COLUNAS_TABELA_ENCAMINHAMENTOS = [
    "os",
    "motivo_encaminhamento",
    "data_encaminhamento",
    "classificacao",
    "tempo_de_espera",
]


def _eh_excluido(registro: dict) -> bool:
    """True quando a solicitação tem exclusão lógica (campo 'excluido' com
    valor 'verdadeiro' em qualquer formato: 1, '1', True, 'sim'...)."""
    valor = registro.get("excluido")
    if valor is None:
        return False
    if isinstance(valor, bool):
        return valor
    try:
        return float(valor) != 0
    except (TypeError, ValueError):
        return str(valor).strip().lower() in ("1", "true", "sim")


def codigo_como_texto(serie: pd.Series) -> pd.Series:
    """Converte códigos em texto ('843', e não '843.0' — o que acontece quando
    a coluna tem vazios e o pandas a transforma em decimal). Nulos continuam nulos."""
    texto = serie.astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    return texto.where(serie.notna())


def formatar_duracao(espera) -> str:
    """Timedelta -> texto curto: '3d 4h', '5h 20min' ou '12min'."""
    if pd.isna(espera):
        return "—"
    segundos = espera.total_seconds()
    if segundos < -60:  # tolerância de 1 min pra diferença de relógio entre servidores
        return "⚠️ data futura"
    minutos_totais = max(int(segundos // 60), 0)
    dias, resto = divmod(minutos_totais, 1440)
    horas, minutos = divmod(resto, 60)
    if dias:
        return f"{dias}d {horas}h"
    if horas:
        return f"{horas}h {minutos}min"
    return f"{minutos}min"


def extrair_descricoes_relacoes(dados: list) -> list:
    """Substitui, em cada registro (ainda como dict, antes de virar DataFrame),
    o valor do campo de código pela descrição vinda do objeto aninhado
    trazido pelo parâmetro with. Remove o objeto aninhado depois de usá-lo,
    para não sobrar uma coluna com dicionário dentro da tabela."""
    for registro in dados:
        for campo_codigo, chave_relacao in RELACOES_WITH.items():
            relacao = registro.get(chave_relacao)
            if isinstance(relacao, dict) and relacao.get("descricao") is not None:
                registro[campo_codigo] = relacao["descricao"]
            registro.pop(chave_relacao, None)
    return dados


def montar_df_encaminhamentos(dados: list) -> pd.DataFrame:
    """1 linha por encaminhamento. Uma solicitação pode ter vários — por isso
    fica em tabela separada: se ficasse na tabela principal, os dados da
    solicitação (título, status, etc.) seriam duplicados a cada encaminhamento.

    A ligação com a tabela principal é:  df["codigo"] == df_enc["os_pai"]

    Colunas extras que servem de apoio:
      os_pai  -> código da solicitação que continha o encaminhamento
      e_atual -> True no encaminhamento MAIS RECENTE de cada solicitação
                 (é ele que diz com quem a solicitação está agora)
    """
    linhas = []
    for registro in dados:
        os_pai = registro.get("codigo")
        lista = registro.get("encaminhamentos") or []
        for posicao, enc in enumerate(lista):
            if not isinstance(enc, dict):
                continue
            classificacao = enc.get("classificacao")
            if isinstance(classificacao, dict):
                classificacao = classificacao.get("descricao")
            os_enc = enc.get("os")
            linhas.append({
                "os": os_pai if os_enc is None else os_enc,
                "os_pai": os_pai,
                "codigo_encaminhamento": enc.get("codigo"),
                "usuario_de": enc.get("usuario_de"),
                "usuario_para": enc.get("usuario_para"),
                "motivo_encaminhamento": enc.get("motivo_encaminhamento"),
                "data_encaminhamento": enc.get("data_encaminhamento"),
                "classificacao": classificacao,
                "posicao_na_api": posicao,
            })

    colunas = ["os", "os_pai", "codigo_encaminhamento", "usuario_de", "usuario_para",
               "motivo_encaminhamento", "data_encaminhamento", "classificacao", "posicao_na_api"]
    df_enc = pd.DataFrame(linhas, columns=colunas)
    df_enc["data_encaminhamento"] = pd.to_datetime(df_enc["data_encaminhamento"], errors="coerce")

    # Não confiamos só na ordem que a API manda: ordenamos pela data (datas
    # vazias ficam como as mais antigas, pra nunca virarem "o atual").
    df_enc = df_enc.sort_values(
        ["os_pai", "data_encaminhamento", "codigo_encaminhamento"],
        na_position="first", kind="stable",
    )
    df_enc["e_atual"] = df_enc.groupby("os_pai", dropna=False).cumcount(ascending=False).eq(0)
    df_enc = df_enc.reset_index(drop=True)

    # De-para de usuario_de / usuario_para — mesmo Excel de usuários dos
    # outros campos (responsavel, cliente, usuario_solicitante...).
    config = MAPAS_DE_PARA["usuarios"]
    try:
        mapa_usuarios = carregar_mapa(config["arquivo"], config["coluna_codigo"], config["coluna_descricao"])
    except FileNotFoundError:
        mapa_usuarios = {}
        st.warning(f"Arquivo de-para não encontrado: {config['arquivo']}")
    for col in ("usuario_de", "usuario_para"):
        codigos = codigo_como_texto(df_enc[col])
        df_enc[col] = codigos.map(mapa_usuarios).fillna(codigos)

    return df_enc


@st.cache_data(ttl=3600, show_spinner=False)  # de-para muda pouco: cache de 1h
def carregar_mapa(arquivo: str, coluna_codigo: str, coluna_descricao: str) -> dict:
    """Lê um Excel de-para e devolve um dicionário {código (str): descrição}."""
    df_mapa = pd.read_excel(arquivo)
    return dict(zip(df_mapa[coluna_codigo].astype(str), df_mapa[coluna_descricao]))


def aplicar_de_paras(df: pd.DataFrame) -> pd.DataFrame:
    """Substitui os códigos pelas descrições em todas as colunas configuradas
    em MAPAS_DE_PARA. Se um código não for encontrado no de-para, mantém o
    valor original (não some nenhum registro)."""
    df = df.copy()
    for config in MAPAS_DE_PARA.values():
        try:
            mapa = carregar_mapa(
                config["arquivo"], config["coluna_codigo"], config["coluna_descricao"]
            )
        except FileNotFoundError:
            st.warning(f"Arquivo de-para não encontrado: {config['arquivo']}")
            continue

        for col in config["colunas_tabela_principal"]:
            if col in df.columns:
                codigos_como_texto = df[col].astype(str)
                # fillna usa a versão em texto do código original, não o valor
                # bruto — assim a coluna nunca fica com int e str misturados
                df[col] = codigos_como_texto.map(mapa).fillna(codigos_como_texto)
    return df


def _buscar_dados_api() -> tuple:
    """Faz a chamada real na API. Devolve (dados_brutos, df, df_enc, meta).
    Pode demorar (a API às vezes é lenta) — por isso ela só roda dentro da
    thread de segundo plano, nunca no caminho que desenha a tela."""
    params = {}
    nomes_with = list(RELACOES_WITH.values()) + WITH_LISTAS + ["encaminhamentos.classificacao"]
    if nomes_with:
        params["with"] = ",".join(nomes_with)

    resp = SESSAO_HTTP.get(API_URL, params=params, timeout=30)
    resp.raise_for_status()
    resposta_bruta = resp.json()

    # A API pode devolver uma lista direta, ou um dict com a lista
    # dentro de alguma chave (ex: "data", "items", "registros"...).
    # `meta` guarda o resto do dict (ex: total/paginação), se houver.
    dados = resposta_bruta
    meta = {}
    if isinstance(dados, dict):
        for chave in ["data", "resultado", "items", "os", "registros", "result"]:
            if chave in dados and isinstance(dados[chave], list):
                meta = {k: v for k, v in dados.items() if k != chave}
                dados = dados[chave]
                break

    # Exclusão lógica: a API devolve solicitações com excluido=1 (com motivo
    # em excluido_motivo, ex: "Documento duplicado") que não deveriam entrar
    # em nenhuma contagem, tabela ou fila de encaminhamentos. Filtramos ANTES
    # de montar as duas tabelas, pra remover também os encaminhamentos de
    # solicitações excluídas.
    linhas_brutas = len(dados)
    dados = [r for r in dados if not _eh_excluido(r)]
    meta["linhas_brutas"] = linhas_brutas
    meta["linhas_excluidas_removidas"] = linhas_brutas - len(dados)

    # Tabela de encaminhamentos — lê a chave "encaminhamentos" de cada
    # registro ANTES de tirarmos ela da tabela principal, logo abaixo.
    df_enc = montar_df_encaminhamentos(dados)

    registros_principais = [
        {k: v for k, v in registro.items() if k not in WITH_LISTAS}
        for registro in dados
    ]
    dados_processados = extrair_descricoes_relacoes(copy.deepcopy(registros_principais))
    df = pd.DataFrame(dados_processados)
    df = aplicar_de_paras(df)

    meta["total_encaminhamentos"] = len(df_enc)
    return dados, df, df_enc, meta


@st.cache_resource  # roda só 1 vez, e o resultado é compartilhado por TODOS os usuários do app
def obter_cache_compartilhado():
    """Cria o estado com os últimos dados válidos e dispara a thread que
    atualiza esse estado em segundo plano, a cada INTERVALO_ATUALIZACAO_SEGUNDOS.

    A tela nunca chama a API diretamente: ela só lê `estado`, protegido por
    `lock`. Isso desacopla a velocidade da API da velocidade da interface —
    mesmo que a API demore 20-30s pra responder, quem está olhando o
    dashboard não fica esperando, só vê os últimos dados prontos."""
    estado = {
        "bruto": [],
        "df": pd.DataFrame(),
        "df_enc": pd.DataFrame(),
        "meta": {},
        "ultima_atualizacao": None,
        "erro": None,
    }
    lock = threading.Lock()

    def atualizar_agora():
        try:
            bruto, df, df_enc, meta = _buscar_dados_api()
            with lock:
                estado["bruto"] = bruto
                estado["df"] = df
                estado["df_enc"] = df_enc
                estado["meta"] = meta
                estado["ultima_atualizacao"] = datetime.now()
                estado["erro"] = None
        except Exception as e:
            with lock:
                estado["erro"] = str(e)

    def loop_segundo_plano():
        while True:
            atualizar_agora()
            time.sleep(INTERVALO_ATUALIZACAO_SEGUNDOS)

    threading.Thread(target=loop_segundo_plano, daemon=True).start()
    return estado, lock, atualizar_agora


def _ler_estado() -> dict:
    """Foto instantânea (cópia) do que está na memória — nunca chama a API."""
    estado, lock, _ = obter_cache_compartilhado()
    with lock:
        return {
            "df": estado["df"].copy(),
            "df_enc": estado["df_enc"].copy(),
            "bruto": list(estado["bruto"]),
            "meta": dict(estado["meta"]),
            "ultima_atualizacao": estado["ultima_atualizacao"],
            "erro": estado["erro"],
        }


def _cabecalho(snap: dict) -> bool:
    """Mostra a hora da última atualização e os erros. Devolve False quando
    ainda não há dados prontos para desenhar a aba."""
    if snap["ultima_atualizacao"]:
        st.caption(f"🔄 Última atualização: {snap['ultima_atualizacao'].strftime('%H:%M:%S')}")
    else:
        st.caption("🔄 Buscando os dados pela primeira vez... (pode levar alguns segundos)")

    if snap["erro"]:
        st.warning(
            f"⚠️ A última tentativa de atualização falhou: {snap['erro']} — "
            f"mostrando os últimos dados que deram certo, se houver."
        )

    if snap["df"].empty:
        st.info("Ainda não há dados carregados. Aguarde a primeira atualização ou use o botão na barra lateral.")
        return False
    return True


def _render_filtros(df: pd.DataFrame, prefixo: str):
    """Desenha os 4 filtros e devolve o DataFrame já filtrado. `prefixo`
    deixa as chaves dos widgets únicas em cada aba."""
    st.subheader("Filtros")
    col_f1, col_f2, col_f3, col_f4 = st.columns(4)

    def opcoes(coluna):
        return sorted(df[coluna].dropna().astype(str).unique()) if coluna in df.columns else []

    filtro_status = col_f1.multiselect("Status", opcoes("status"), key=f"{prefixo}_status")
    filtro_tipo = col_f2.multiselect("Tipo de Solicitação", opcoes("os_tipo_servico"), key=f"{prefixo}_tipo")
    filtro_solicitante = col_f3.multiselect("Solicitante", opcoes("usuario_solicitante"), key=f"{prefixo}_solicitante")
    filtro_codigo = col_f4.text_input("Código da solicitação", key=f"{prefixo}_codigo")

    df_filtrado = df.copy()
    if filtro_status:
        df_filtrado = df_filtrado[df_filtrado["status"].isin(filtro_status)]
    if filtro_tipo:
        df_filtrado = df_filtrado[df_filtrado["os_tipo_servico"].isin(filtro_tipo)]
    if filtro_solicitante:
        df_filtrado = df_filtrado[df_filtrado["usuario_solicitante"].isin(filtro_solicitante)]
    if filtro_codigo and "codigo" in df_filtrado.columns:
        df_filtrado = df_filtrado[
            df_filtrado["codigo"].astype(str).str.contains(filtro_codigo.strip(), case=False, na=False)
        ]
    return df_filtrado


def _render_cards_status(df_filtrado: pd.DataFrame):
    """Um card por status, com a contagem do filtro atual."""
    st.caption(f"Total de solicitações no filtro atual: **{len(df_filtrado)}**")
    colunas_cards = st.columns(len(STATUS_CONFIG))

    for coluna, (status, cfg) in zip(colunas_cards, STATUS_CONFIG.items()):
        contagem = int((df_filtrado["status"] == status).sum()) if "status" in df_filtrado.columns else 0
        coluna.markdown(
            f"""
            <div class="status-card" style="background-color:{cfg['cor']}15; border-left:4px solid {cfg['cor']};">
                <div class="icone">{cfg['icone']}</div>
                <div class="rotulo">{status}</div>
                <div class="valor" style="color:{cfg['cor']};">{contagem}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )


@st.fragment(run_every=5)  # barato: só lê o estado em memória, não chama a API
def painel():
    snap = _ler_estado()
    if not _cabecalho(snap):
        return
    df = snap["df"]

    # --------------------------------------------------------------
    # Debug: ver o JSON bruto (útil pra conferir o que a API realmente
    # devolve em campos como r_os_tipo_servico ou encaminhamentos, antes
    # de qualquer processamento)
    # --------------------------------------------------------------
    with st.expander("🔍 Debug: ver JSON bruto do 1º registro"):
        if snap["bruto"]:
            st.json(snap["bruto"][2026])
        else:
            st.write("Nenhum registro carregado ainda.")

    excluidos_removidos = snap["meta"].get("linhas_excluidas_removidas", 0)
    if excluidos_removidos:
        st.caption(
            f"🗑️ {excluidos_removidos} solicitações com exclusão lógica (`excluido=1`) "
            f"foram filtradas antes de qualquer contagem."
        )

    df_filtrado = _render_filtros(df, prefixo="sol")

    st.divider()
    _render_cards_status(df_filtrado)
    st.divider()

    # --------------------------------------------------------------
    # Tabela detalhada — agrupada por solicitante
    # --------------------------------------------------------------
    st.subheader("Dados detalhados por solicitante")

    colunas_disponiveis = [c for c in COLUNAS_TABELA_DETALHADA if c in df_filtrado.columns]

    if df_filtrado.empty:
        st.info("Nenhum registro encontrado com os filtros selecionados.")
    elif "usuario_solicitante" in df_filtrado.columns:
        grupos = sorted(df_filtrado.groupby("usuario_solicitante"), key=lambda g: len(g[1]), reverse=True)
        for solicitante, grupo in grupos:
            with st.expander(f"👤 {solicitante}  ·  {len(grupo)} solicitação(ões)"):
                st.dataframe(grupo[colunas_disponiveis], use_container_width=True, hide_index=True)
    else:
        st.dataframe(df_filtrado[colunas_disponiveis], use_container_width=True, hide_index=True)


@st.fragment(run_every=5)  # barato: só lê o estado em memória, não chama a API
def painel_encaminhamentos():
    snap = _ler_estado()
    if not _cabecalho(snap):
        return
    df, df_enc = snap["df"], snap["df_enc"]

    df_filtrado = _render_filtros(df, prefixo="enc")

    st.divider()
    _render_cards_status(df_filtrado)
    st.divider()

    # --------------------------------------------------------------
    # Tabela detalhada — agrupada por usuário de destino (usuario_para)
    # --------------------------------------------------------------
    st.subheader("Dados detalhados por usuário de destino")
    st.caption(
        f"Considera só solicitações com status {' ou '.join(STATUS_FILA_ABERTA)}, "
        "olhando o encaminhamento mais recente de cada uma (quem está com ela agora). "
        "Tempo de espera = agora − data do encaminhamento."
    )

    if df_enc.empty:
        st.info("Nenhum encaminhamento carregado ainda (confira o debug na aba Solicitações).")
        return

    if "status" in df_filtrado.columns:
        abertas = df_filtrado.loc[df_filtrado["status"].isin(STATUS_FILA_ABERTA), "codigo"]
    else:
        abertas = df_filtrado["codigo"]
    codigos_abertos = set(codigo_como_texto(abertas).dropna())

    fila = df_enc[df_enc["e_atual"] & codigo_como_texto(df_enc["os_pai"]).isin(codigos_abertos)].copy()

    if fila.empty:
        st.info("Nenhum encaminhamento pendente para os filtros selecionados.")
        return

    fila["usuario_para"] = fila["usuario_para"].fillna("(sem usuário de destino)")
    agora = pd.Timestamp.now()
    bruta = agora - fila["data_encaminhamento"]
    fila["tempo_de_espera"] = bruta.map(formatar_duracao)  # texto mostra '⚠️ data futura' / '—'
    # `espera` (usada em ordenação, média e máximo) ignora datas futuras e vazias
    fila["espera"] = bruta.mask(bruta < pd.Timedelta(minutes=-1)).clip(lower=pd.Timedelta(0))

    st.caption(
        f"**{len(fila)}** solicitações em fila com **{fila['usuario_para'].nunique()}** pessoa(s)"
        f"  ·  espera média **{formatar_duracao(fila['espera'].mean())}**"
        f"  ·  maior espera **{formatar_duracao(fila['espera'].max())}**"
    )

    grupos = sorted(fila.groupby("usuario_para"), key=lambda g: len(g[1]), reverse=True)
    for usuario, grupo in grupos:
        grupo = grupo.sort_values("espera", ascending=False)  # quem espera há mais tempo, primeiro
        with st.expander(f"👤 {usuario}  ·  {len(grupo)} solicitação(ões)"):
            st.caption(
                f"Maior espera: **{formatar_duracao(grupo['espera'].max())}**"
                f"  ·  espera média: **{formatar_duracao(grupo['espera'].mean())}**"
            )
            st.dataframe(grupo[COLUNAS_TABELA_ENCAMINHAMENTOS], use_container_width=True, hide_index=True)


# ------------------------------------------------------------------
# LAYOUT PRINCIPAL
# ------------------------------------------------------------------
st.title("📊 Dashboard de gestão de solicitações")

with st.sidebar:
    st.header("Configurações")
    st.write(
        f"Os dados são buscados em segundo plano a cada **{INTERVALO_ATUALIZACAO_SEGUNDOS} segundos**. "
        "A tela em si atualiza a exibição a cada poucos segundos, mas sem travar esperando a API."
    )
    if st.button("🔄 Forçar atualização agora"):
        _, _, atualizar_agora = obter_cache_compartilhado()
        with st.spinner("Buscando dados mais recentes..."):
            atualizar_agora()
        st.rerun()

aba_solicitacoes, aba_encaminhamentos = st.tabs(["📋 Solicitações", "📬 Encaminhamentos"])
with aba_solicitacoes:
    painel()
with aba_encaminhamentos:
    painel_encaminhamentos()