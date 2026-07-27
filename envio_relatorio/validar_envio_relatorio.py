#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Validador pré-envio do Relatório Semanal +TOP.

Executa verificações independentes sobre os cálculos do envio_relatorio e
bloqueia o envio quando encontra inconsistências críticas.

Uso standalone (apenas gera relatório de validação):
    python3 validar_envio_relatorio.py

Uso integrado:
    Importado e chamado por gerar_relatorio_semanal.py antes do envio do email.
"""

import logging
import re
import sys
import unicodedata
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd

# ---------------------------------------------------------------------------
# CONSTANTES (espelho de gerar_relatorio_semanal.py para independência)
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "bases"
OUTPUT_DIR = BASE_DIR / "relatorios_gerados"
LOG_DIR = BASE_DIR / "logs"

REVENDAS_EXCLUIR = {"EAI", "Whirlpool", "Novo Mundo", "ELETROMÓVEIS MARTINELLO"}
STATUS_CADASTRO_EXCLUIR = {"Inativo", "Bloqueado", "Reprovado", "Aguardando Aprovação"}

BANCO_PARTICIPANTES_PATH = Path(
    "/home/thamiresvieira/projetos/validacoes_precadastro/PROD_WHP_Participante_Banco_v2_16062026.xlsx"
)

PRAZO_PRE_CADASTRO_INATIVO = 90

META_CADASTRO = 85.0
META_TREINAMENTOS = 70.0
META_ACEITES = 70.0

NOME_REVENDA_EXIBICAO = {
    "CDA": "Casas da Água",
    "Imperio": "Império",
}

FILIAIS_B = {
    "Becker B": "Becker",
    "Bemol B": "Bemol",
    "Colombo B": "Colombo",
    "Imperio B": "Imperio",
    "Koerich B": "Koerich",
    "Laser Eletro B": "Laser Eletro",
    "Multiloja B": "Multiloja",
    "Nosso Lar B": "Nosso Lar",
    "Ramsons B": "Ramsons",
    "Sipolatti B": "Sipolatti",
    "Taqi B": "Taqi",
    "Tele Rio B": "Tele Rio",
    "Armazem Mateus B": "Armazem Mateus",
}

MAPA_REVENDA_PRINCIPAL = {
    "CASAS DA AGUA": "Casas da Água",
    "CDA": "Casas da Água",
    "ELETRO MATEUS": "Armazem Mateus",
    "ARMAZEM MATEUS": "Armazem Mateus",
    "MERCADOMOVEIS VAREJO": "MM Varejo",
    "MERCADOMOVEIS": "MM Atacado",
    "RAMSONS": "Ramsons",
    "SOLAR COMERCIO E AGROINDUSTRIA LTDA": "Solar",
    "SOLAR MOVEIS E ELETROS": "Solar Magazine",
    "LOJA SOLAR": "Solar",
    "GAZIN ATACADO": "Gazin Atacado",
    "GAZIN ONLINE": "Gazin Online",
    "GAZIN": "Gazin Varejo",
    "LOJAS GUAIBIM": "Guaibim",
    "GUAIBIM": "Guaibim",
    "HAVAN": "Havan",
    "JMAHFUZ": "Jmahfuz",
    "IMPERIO": "Imperio",
    "LASER ELETRO": "Laser Eletro",
    "NOSSO LAR LOJAS DE DEPTOS LTDA": "Nosso Lar",
    "MILLENA MOVEIS": "Millena",
    "BECKER": "Becker",
    "BEMOL": "Bemol",
    "COLOMBO": "Colombo",
    "Colombo": "Colombo",
    "FORMOSA": "Formosa",
    "ESTRELA": "Estrela",
    "ANGELONI": "Angeloni",
    "SIPOLATTI": "Sipolatti",
    "TAQI": "Taqi",
    "TELE RIO": "Tele Rio",
    "ZEMA": "Zema",
    "ZENIR": "Zenir",
}


# ---------------------------------------------------------------------------
# FUNÇÕES AUXILIARES (copiadas de gerar_relatorio_semanal.py)
# ---------------------------------------------------------------------------
def _normalizar_texto(texto: str) -> str:
    if pd.isna(texto):
        return ""
    texto = str(texto).strip().upper()
    texto = unicodedata.normalize("NFKD", texto).encode("ASCII", "ignore").decode("ASCII")
    return texto


def limpar_cpf(cpf):
    """Normaliza CPF/CNPJ para texto com 11 dígitos."""
    if pd.isna(cpf):
        return None
    try:
        if isinstance(cpf, float):
            cpf = int(cpf)
    except Exception:
        pass
    s = str(cpf).strip().replace("'", "").replace(".", "").replace("-", "").replace("/", "").replace(" ", "")
    if s.endswith(".0"):
        s = s[:-2]
    return s.zfill(11)


def normalizar_revenda(grupo):
    """Agrupa filiais 'B' na revenda principal."""
    g = str(grupo).strip()
    return FILIAIS_B.get(g, g)


def normalizar_revenda_hierarquia(nome_revenda):
    """Normaliza nome de revenda vindo das hierarquias para a revenda principal."""
    if pd.isna(nome_revenda):
        return nome_revenda
    nome = str(nome_revenda).strip()
    nome = nome.upper()
    nome = re.sub(r"\s+", " ", nome)
    nome = (
        nome.replace("Á", "A")
        .replace("É", "E")
        .replace("Í", "I")
        .replace("Ó", "O")
        .replace("Ú", "U")
        .replace("Â", "A")
        .replace("Ê", "E")
        .replace("Ô", "O")
        .replace("Ã", "A")
        .replace("Ç", "C")
    )
    nome_upper = nome

    if nome_upper in MAPA_REVENDA_PRINCIPAL:
        return MAPA_REVENDA_PRINCIPAL[nome_upper]

    for chave, principal in sorted(MAPA_REVENDA_PRINCIPAL.items(), key=lambda x: -len(x[0])):
        chave_limpa = re.sub(r"\s+", " ", chave.upper())
        if nome_upper.startswith(chave_limpa):
            return principal

    for chave, principal in sorted(MAPA_REVENDA_PRINCIPAL.items(), key=lambda x: -len(x[0])):
        chave_limpa = re.sub(r"\s+", " ", chave.upper())
        if re.search(r"\b" + re.escape(chave_limpa) + r"\b", nome_upper):
            return principal

    return nome_revenda.strip() if isinstance(nome_revenda, str) else nome_revenda


def nome_revenda_exibicao(revenda):
    """Retorna nome amigavel da revenda para exibicao."""
    rev = str(revenda).strip() if not pd.isna(revenda) else ""
    return NOME_REVENDA_EXIBICAO.get(rev, rev)


def regional_curta(regional):
    """Remove prefixo 'CONTA ' da regional."""
    r = str(regional).strip()
    return r.replace("CONTA ", "") if r.startswith("CONTA ") else r


def mapeamento_revenda_regional(df_cad):
    """Cria mapeamento revenda -> regional a partir do cadastro base."""
    mapa = {}
    if "grupo" in df_cad.columns and "regional" in df_cad.columns:
        for _, row in df_cad.dropna(subset=["grupo", "regional"]).iterrows():
            rev = str(row["grupo"]).strip()
            reg = str(row["regional"]).strip()
            if rev and reg and rev.lower() != "nan" and reg.lower() != "nan":
                mapa[rev] = reg
    return mapa


def regional_por_revenda(revenda, mapa_regional):
    """Busca regional no mapa de forma case-insensitiva."""
    if pd.isna(revenda):
        return None
    rev = str(revenda).strip()
    if not rev or rev.lower() == "nan":
        return None
    if rev in mapa_regional:
        return mapa_regional[rev]
    rev_upper = rev.upper()
    for k, v in mapa_regional.items():
        if str(k).strip().upper() == rev_upper:
            return v
    rev_lower = rev.lower().replace("á", "a").replace("ã", "a").replace("ç", "c")
    if rev_lower == "casas da agua" and "CDA" in mapa_regional:
        return mapa_regional["CDA"]
    return None


# ---------------------------------------------------------------------------
# MODELOS DE RESULTADO
# ---------------------------------------------------------------------------
@dataclass
class ResultadoValidacao:
    regra: str
    tipo: str  # OK, ALERTA, CRITICO
    mensagem: str
    esperado: Optional[float] = None
    obtido: Optional[float] = None
    detalhes: Dict = field(default_factory=dict)


# ---------------------------------------------------------------------------
# FUNÇÕES DE CÁLCULO INDEPENDENTES
# ---------------------------------------------------------------------------
def reimplementar_status_cadastro(
    df_cad: pd.DataFrame, banco_path: Path = BANCO_PARTICIPANTES_PATH, prazo_dias: int = PRAZO_PRE_CADASTRO_INATIVO
) -> Dict[str, str]:
    """
    Reaplica de forma independente:
      - exclusão de revendas teste
      - regra dos 90 dias usando a base do banco
      - remoção de status operacionais excluídos
    Retorna dict cpf_limp -> status atualizado.
    """
    df = df_cad.copy()

    # Exclui revendas de teste
    if "grupo" in df.columns:
        df = df[~df["grupo"].isin(REVENDAS_EXCLUIR)].copy()

    # Regra dos 90 dias
    if banco_path.exists() and "status" in df.columns:
        df_banco = pd.read_excel(banco_path)
        df_banco["cpf_limp"] = df_banco["Cpf"].apply(limpar_cpf)
        df_banco["DataInclusao_dt"] = pd.to_datetime(df_banco["DataInclusao"], errors="coerce")
        banco_clean = df_banco[["cpf_limp", "DataInclusao_dt"]].drop_duplicates(subset=["cpf_limp"], keep="first")
        df = df.merge(banco_clean, on="cpf_limp", how="left")
        dias_desde_inclusao = (pd.Timestamp("today").normalize() - df["DataInclusao_dt"]).dt.days
        mask_pre_inativo = (
            (df["status"] == "Pré-Cadastrado")
            & (df["DataInclusao_dt"].notna())
            & (dias_desde_inclusao >= prazo_dias)
        )
        if mask_pre_inativo.sum() > 0:
            df.loc[mask_pre_inativo, "status"] = "Inativo"
        df = df.drop(columns=["DataInclusao_dt"], errors="ignore")

    # Remove status operacionais excluídos
    if "status" in df.columns:
        df = df[~df["status"].isin(STATUS_CADASTRO_EXCLUIR)].copy()

    if "cpf_limp" not in df.columns or "status" not in df.columns:
        return {}

    return (
        df.dropna(subset=["cpf_limp", "status"])
        .drop_duplicates(subset=["cpf_limp"], keep="first")
        .set_index("cpf_limp")["status"]
        .to_dict()
    )


def calcular_ativos_independentemente(
    df_cad: pd.DataFrame,
    df_hier: pd.DataFrame,
    df_ferias: Optional[pd.DataFrame],
    status_atualizado: Dict[str, str],
) -> pd.DataFrame:
    """
    Replica a lógica de calcular_cadastros:
      - base = hierarquia ativa (DESLIGADO=NÃO), já deduplicada
      - férias como inativos
      - status vindo do dict atualizado
    Retorna DataFrame resumo por revenda com total, ativos, inativos, pre_cadastro.
    """
    if df_hier is None or df_hier.empty:
        return pd.DataFrame()

    base = df_hier[["cpf_limp", "revenda"]].drop_duplicates(keep="first").copy()
    base["em_ferias"] = False

    if df_ferias is not None and not df_ferias.empty:
        ferias = df_ferias[["cpf_limp", "revenda"]].drop_duplicates(keep="first").copy()
        ferias["em_ferias"] = True
        base = pd.concat([base, ferias], ignore_index=True)
        base = base.drop_duplicates(subset=["cpf_limp"], keep="first")

    base["em_ferias"] = base["em_ferias"].fillna(False)
    base["status"] = base["cpf_limp"].map(status_atualizado)

    def _classificar(row):
        if row["em_ferias"]:
            return "inativo"
        status = str(row["status"]).strip() if pd.notna(row["status"]) else ""
        if status == "Ativo":
            return "ativo"
        if status in {"Inativo", "Bloqueado"}:
            return "inativo"
        return "pre_cadastro"

    base["tipo_status"] = base.apply(_classificar, axis=1)

    def _agg(grupo_df):
        total = grupo_df["cpf_limp"].nunique()
        ativos = (grupo_df["tipo_status"] == "ativo").sum()
        inativos = (grupo_df["tipo_status"] == "inativo").sum()
        pre_cadastro = total - ativos - inativos
        return pd.Series(
            {
                "total": total,
                "ativos": ativos,
                "inativos": inativos,
                "pre_cadastro": pre_cadastro,
            }
        )

    return base.groupby("revenda").apply(_agg).reset_index()


# ---------------------------------------------------------------------------
# CLASSE VALIDADOR
# ---------------------------------------------------------------------------
class ValidadorEnvioRelatorio:
    def __init__(
        self,
        cadastro_df: pd.DataFrame,
        hierarquia_df: Optional[pd.DataFrame],
        ferias_df: Optional[pd.DataFrame],
        treinamentos_df: pd.DataFrame,
        aceites_df: pd.DataFrame,
        cad_reg: pd.DataFrame,
        cad_rev: pd.DataFrame,
        trein_reg: pd.DataFrame,
        trein_rev: pd.DataFrame,
        aceite_reg: pd.DataFrame,
        aceite_rev: pd.DataFrame,
        base_trein: pd.DataFrame,
        base_aceite: pd.DataFrame,
        mes_referencia: str,
        mes_aceite: str,
        cursos_info: Optional[Dict] = None,
        output_dir: Path = OUTPUT_DIR,
        log_dir: Path = LOG_DIR,
        logger_validador: Optional[logging.Logger] = None,
    ):
        self.cadastro_df = cadastro_df
        self.hierarquia_df = hierarquia_df
        self.ferias_df = ferias_df
        self.treinamentos_df = treinamentos_df
        self.aceites_df = aceites_df
        self.cad_reg = cad_reg
        self.cad_rev = cad_rev
        self.trein_reg = trein_reg
        self.trein_rev = trein_rev
        self.aceite_reg = aceite_reg
        self.aceite_rev = aceite_rev
        self.base_trein = base_trein
        self.base_aceite = base_aceite
        self.mes_referencia = mes_referencia
        self.mes_aceite = mes_aceite
        self.cursos_info = cursos_info or {}
        self.output_dir = output_dir
        self.log_dir = log_dir
        self.resultados: List[ResultadoValidacao] = []
        self.logger = logger_validador or logging.getLogger(__name__)

        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.log_dir.mkdir(parents=True, exist_ok=True)

    def executar(self, gerar_excel: bool = True) -> Tuple[bool, dict]:
        self.resultados = []

        self._validar_base_nao_vazia()
        self._validar_desligados_na_base_ativa()
        self._validar_ferias_como_inativos()
        self._validar_ativos_alinhados()
        self._validar_duplicidades_hierarquia()
        self._validar_metas()
        self._validar_anomalias_totais()

        criticos = [r for r in self.resultados if r.tipo == "CRITICO"]
        alertas = [r for r in self.resultados if r.tipo == "ALERTA"]
        ok = len(criticos) == 0

        relatorio = {
            "ok": ok,
            "criticos": criticos,
            "alertas": alertas,
            "todos": self.resultados,
            "caminho_excel": None,
            "timestamp": datetime.now().isoformat(),
        }

        if gerar_excel:
            relatorio["caminho_excel"] = self._exportar_excel(relatorio)

        return ok, relatorio

    # -----------------------------------------------------------------------
    # REGRAS INDIVIDUAIS
    # -----------------------------------------------------------------------
    def _validar_base_nao_vazia(self):
        critico = False
        if self.hierarquia_df is None or self.hierarquia_df.empty:
            self.resultados.append(
                ResultadoValidacao(
                    regra="Base de cálculo não vazia",
                    tipo="CRITICO",
                    mensagem="DataFrame de hierarquia está vazio.",
                )
            )
            critico = True
        if self.cad_reg is None or self.cad_reg.empty:
            self.resultados.append(
                ResultadoValidacao(
                    regra="Base de cálculo não vazia",
                    tipo="CRITICO",
                    mensagem="DataFrame cad_reg está vazio.",
                )
            )
            critico = True
        if not critico:
            self.resultados.append(
                ResultadoValidacao(
                    regra="Base de cálculo não vazia",
                    tipo="OK",
                    mensagem="Hierarquia e cadastro regional possuem dados.",
                )
            )

    def _validar_desligados_na_base_ativa(self):
        if self.hierarquia_df is None or self.hierarquia_df.empty:
            return

        # A hierarquia já deve ter sido filtrada para DESLIGADO=NÃO no script principal.
        # Esta validação detecta se algum registro com outro valor ainda permaneceu.
        col_desligado = None
        for c in self.hierarquia_df.columns:
            if _normalizar_texto(c) == "DESLIGADO":
                col_desligado = c
                break

        if col_desligado is None:
            self.resultados.append(
                ResultadoValidacao(
                    regra="DESLIGADO na base ativa",
                    tipo="ALERTA",
                    mensagem="Coluna DESLIGADO não encontrada na hierarquia; validação não aplicada.",
                )
            )
            return

        deslig_norm = self.hierarquia_df[col_desligado].fillna("").astype(str).str.strip().str.upper()
        invalidos = self.hierarquia_df[~deslig_norm.isin(["NAO", "NÃO"])]

        if not invalidos.empty:
            self.resultados.append(
                ResultadoValidacao(
                    regra="DESLIGADO na base ativa",
                    tipo="CRITICO",
                    mensagem=f"{invalidos['cpf_limp'].nunique()} CPF(s) na hierarquia ativa possuem DESLIGADO diferente de NÃO.",
                    obtido=invalidos["cpf_limp"].nunique(),
                    esperado=0,
                    detalhes={"revendas_afetadas": sorted(invalidos["revenda"].dropna().unique().tolist())},
                )
            )
        else:
            self.resultados.append(
                ResultadoValidacao(
                    regra="DESLIGADO na base ativa",
                    tipo="OK",
                    mensagem="Todos os CPFs na hierarquia ativa possuem DESLIGADO=NÃO.",
                )
            )

    def _validar_ferias_como_inativos(self):
        if self.ferias_df is None or self.ferias_df.empty:
            self.resultados.append(
                ResultadoValidacao(
                    regra="Férias como inativos",
                    tipo="OK",
                    mensagem="Nenhum CPF em férias na hierarquia.",
                )
            )
            return

        if self.hierarquia_df is None or self.hierarquia_df.empty or self.cad_rev is None or self.cad_rev.empty:
            self.resultados.append(
                ResultadoValidacao(
                    regra="Férias como inativos",
                    tipo="ALERTA",
                    mensagem="Não foi possível verificar férias: dados insuficientes.",
                )
            )
            return

        # Reimplementa a lógica de férias do envio_relatorio:
        # CPFs em férias que NÃO estão na base ativa são adicionados como inativos.
        cpfs_hier_ativa = set(self.hierarquia_df["cpf_limp"].dropna().unique())
        cpfs_ferias = set(self.ferias_df["cpf_limp"].dropna().unique())
        ferias_efetivas = cpfs_ferias - cpfs_hier_ativa

        if not ferias_efetivas:
            self.resultados.append(
                ResultadoValidacao(
                    regra="Férias como inativos",
                    tipo="OK",
                    mensagem="Todos os CPFs em férias também estão na base ativa; nenhum férias adicional como inativo.",
                )
            )
            return

        # Calcula o total de inativos esperado por revenda
        status_atualizado = reimplementar_status_cadastro(self.cadastro_df)
        ferias_df = self.ferias_df.copy()
        ferias_df = ferias_df[ferias_df["cpf_limp"].isin(ferias_efetivas)].drop_duplicates(subset=["cpf_limp"], keep="first")

        # Classifica cada férias efetiva
        def _classificar(row):
            status = str(status_atualizado.get(row["cpf_limp"], "")).strip()
            if status == "Ativo":
                return "ativo"  # Não deveria acontecer para férias, mas mantém a regra
            if status in {"Inativo", "Bloqueado"}:
                return "inativo"
            return "pre_cadastro"

        ferias_df["tipo_status"] = ferias_df.apply(_classificar, axis=1)
        esperado_inativos_ferias = (ferias_df["tipo_status"] == "inativo").sum()

        # Compara com o relatório: inativos totais devem incluir as férias efetivas inativas
        col_inativos = None
        for c in ["inativos", "Inativos"]:
            if c in self.cad_rev.columns:
                col_inativos = c
                break

        if col_inativos is None:
            self.resultados.append(
                ResultadoValidacao(
                    regra="Férias como inativos",
                    tipo="ALERTA",
                    mensagem="Coluna de inativos não encontrada em cad_rev; validação não aplicada.",
                )
            )
            return

        total_inativos_rel = int(self.cad_rev[col_inativos].sum())

        # O total de inativos do relatório deve ser pelo menos o número de férias efetivas inativas
        # (pois pode haver outros inativos também).
        if total_inativos_rel < esperado_inativos_ferias:
            self.resultados.append(
                ResultadoValidacao(
                    regra="Férias como inativos",
                    tipo="CRITICO",
                    mensagem=(
                        f"Férias efetivas que deveriam ser inativas: {esperado_inativos_ferias}, "
                        f"mas relatório reporta apenas {total_inativos_rel} inativos no total."
                    ),
                    esperado=f"≥ {esperado_inativos_ferias}",
                    obtido=total_inativos_rel,
                )
            )
        else:
            self.resultados.append(
                ResultadoValidacao(
                    regra="Férias como inativos",
                    tipo="OK",
                    mensagem=(
                        f"{len(ferias_efetivas)} CPF(s) em férias (não presentes na base ativa) "
                        f"tratados como inativos no relatório."
                    ),
                )
            )

    def _validar_ativos_alinhados(self):
        if self.hierarquia_df is None or self.hierarquia_df.empty:
            return

        status_atualizado = reimplementar_status_cadastro(self.cadastro_df)
        indep = calcular_ativos_independentemente(
            self.cadastro_df, self.hierarquia_df, self.ferias_df, status_atualizado
        )

        if indep.empty or self.cad_rev is None or self.cad_rev.empty:
            self.resultados.append(
                ResultadoValidacao(
                    regra="Ativos alinhados (relatório vs independente)",
                    tipo="ALERTA",
                    mensagem="Não foi possível comparar ativos: DataFrame vazio.",
                )
            )
            return

        total_indep = int(indep["ativos"].sum())

        # Identifica a coluna de ativos no cad_rev
        col_ativos = None
        for c in ["ativos", "Ativos no +TOP", "Ativos"]:
            if c in self.cad_rev.columns:
                col_ativos = c
                break

        if col_ativos is None:
            self.resultados.append(
                ResultadoValidacao(
                    regra="Ativos alinhados (relatório vs independente)",
                    tipo="ALERTA",
                    mensagem=f"Coluna de ativos não encontrada em cad_rev. Colunas: {list(self.cad_rev.columns)}",
                )
            )
            return

        total_rel = int(self.cad_rev[col_ativos].sum())

        if total_indep != total_rel:
            self.resultados.append(
                ResultadoValidacao(
                    regra="Ativos alinhados (relatório vs independente)",
                    tipo="CRITICO",
                    mensagem="Divergência no total de ativos entre relatório e cálculo independente.",
                    esperado=total_indep,
                    obtido=total_rel,
                    detalhes={"diferenca": total_rel - total_indep},
                )
            )
        else:
            self.resultados.append(
                ResultadoValidacao(
                    regra="Ativos alinhados (relatório vs independente)",
                    tipo="OK",
                    mensagem=f"Total de ativos alinhado: {total_rel:,}.",
                    esperado=total_indep,
                    obtido=total_rel,
                )
            )

    def _validar_duplicidades_hierarquia(self):
        if self.hierarquia_df is None or self.hierarquia_df.empty:
            return

        total = self.hierarquia_df["cpf_limp"].nunique()
        duplicados = self.hierarquia_df[self.hierarquia_df.duplicated(subset=["cpf_limp"], keep=False)]
        qtd_dup = duplicados["cpf_limp"].nunique()

        if qtd_dup == 0:
            self.resultados.append(
                ResultadoValidacao(
                    regra="Duplicidades de CPF na hierarquia",
                    tipo="OK",
                    mensagem="Nenhum CPF duplicado na hierarquia.",
                )
            )
            return

        pct = (qtd_dup / total * 100) if total else 0
        tipo = "CRITICO" if pct > 0.5 else "ALERTA"
        self.resultados.append(
            ResultadoValidacao(
                regra="Duplicidades de CPF na hierarquia",
                tipo=tipo,
                mensagem=f"{qtd_dup} CPF(s) duplicados na hierarquia ({pct:.2f}% da base).",
                obtido=qtd_dup,
                esperado=0,
                detalhes={"percentual": pct},
            )
        )

    def _validar_metas(self):
        def _pct(parte, total):
            return (parte / total * 100) if total else 0

        # Cadastro
        if self.cad_reg is not None and not self.cad_reg.empty:
            col_total = "total" if "total" in self.cad_reg.columns else "Total de participantes"
            col_ativos = "ativos" if "ativos" in self.cad_reg.columns else "Ativos no +TOP"
            total = self.cad_reg[col_total].sum()
            ativos = self.cad_reg[col_ativos].sum()
            pct = _pct(ativos, total)
            if pct < META_CADASTRO:
                self.resultados.append(
                    ResultadoValidacao(
                        regra="Meta de cadastro",
                        tipo="ALERTA",
                        mensagem=f"% de cadastro ({pct:.1f}%) abaixo da meta de {META_CADASTRO:.0f}%.",
                        esperado=META_CADASTRO,
                        obtido=pct,
                    )
                )
            else:
                self.resultados.append(
                    ResultadoValidacao(
                        regra="Meta de cadastro",
                        tipo="OK",
                        mensagem=f"% de cadastro ({pct:.1f}%) atinge a meta.",
                        esperado=META_CADASTRO,
                        obtido=pct,
                    )
                )

        # Treinamentos
        if self.trein_reg is not None and not self.trein_reg.empty:
            col_total = "total_ativos" if "total_ativos" in self.trein_reg.columns else "total"
            col_realizado = "realizaram" if "realizaram" in self.trein_reg.columns else "realizado"
            total = self.trein_reg[col_total].sum() if col_total in self.trein_reg.columns else 0
            realizado = self.trein_reg[col_realizado].sum() if col_realizado in self.trein_reg.columns else 0
            pct = _pct(realizado, total)
            if pct < META_TREINAMENTOS:
                self.resultados.append(
                    ResultadoValidacao(
                        regra="Meta de treinamentos",
                        tipo="ALERTA",
                        mensagem=f"% de treinamentos ({pct:.1f}%) abaixo da meta de {META_TREINAMENTOS:.0f}%.",
                        esperado=META_TREINAMENTOS,
                        obtido=pct,
                    )
                )
            else:
                self.resultados.append(
                    ResultadoValidacao(
                        regra="Meta de treinamentos",
                        tipo="OK",
                        mensagem=f"% de treinamentos ({pct:.1f}%) atinge a meta.",
                        esperado=META_TREINAMENTOS,
                        obtido=pct,
                    )
                )

        # Aceites
        if self.aceite_reg is not None and not self.aceite_reg.empty:
            col_total = "total_ativos" if "total_ativos" in self.aceite_reg.columns else "total"
            total = self.aceite_reg[col_total].sum() if col_total in self.aceite_reg.columns else 0
            aceitaram = self.aceite_reg["aceitaram"].sum() if "aceitaram" in self.aceite_reg.columns else 0
            pct = _pct(aceitaram, total)
            if pct < META_ACEITES:
                self.resultados.append(
                    ResultadoValidacao(
                        regra="Meta de aceites",
                        tipo="ALERTA",
                        mensagem=f"% de aceites ({pct:.1f}%) abaixo da meta de {META_ACEITES:.0f}%.",
                        esperado=META_ACEITES,
                        obtido=pct,
                    )
                )
            else:
                self.resultados.append(
                    ResultadoValidacao(
                        regra="Meta de aceites",
                        tipo="OK",
                        mensagem=f"% de aceites ({pct:.1f}%) atinge a meta.",
                        esperado=META_ACEITES,
                        obtido=pct,
                    )
                )

    def _validar_anomalias_totais(self):
        if self.aceite_reg is None or self.aceite_reg.empty:
            return

        col_total = "total" if "total" in self.aceite_reg.columns else "Total de participantes"
        col_aceite = "aceitaram" if "aceitaram" in self.aceite_reg.columns else "Aceitaram"

        for _, row in self.aceite_reg.iterrows():
            total = row.get(col_total, 0)
            aceite = row.get(col_aceite, 0)
            regional = row.get("regional_curta", row.get("regional", "N/A"))
            if total and aceite == 0:
                self.resultados.append(
                    ResultadoValidacao(
                        regra="Regional/Revenda sem aceite",
                        tipo="ALERTA",
                        mensagem=f"Regional {regional} possui {int(total)} participantes mas 0 aceites.",
                        obtido=0,
                        esperado=">= 1",
                    )
                )

    # -----------------------------------------------------------------------
    # EXPORTAÇÃO
    # -----------------------------------------------------------------------
    def _exportar_excel(self, relatorio: dict) -> Optional[Path]:
        try:
            from openpyxl import Workbook
            from openpyxl.styles import Font, PatternFill, Alignment
        except ImportError:
            self.logger.warning("openpyxl não disponível; Excel de validação não gerado.")
            return None

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        caminho = self.output_dir / f"validacao_envio_relatorio_{timestamp}.xlsx"

        with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
            # Aba Resumo
            linhas = []
            for r in self.resultados:
                linhas.append(
                    {
                        "Regra": r.regra,
                        "Tipo": r.tipo,
                        "Mensagem": r.mensagem,
                        "Esperado": r.esperado,
                        "Obtido": r.obtido,
                    }
                )
            df_resumo = pd.DataFrame(linhas)
            if not df_resumo.empty:
                df_resumo.to_excel(writer, sheet_name="Resumo", index=False)
                ws = writer.sheets["Resumo"]
                for cell in ws[1]:
                    cell.font = Font(bold=True, color="FFFFFF")
                    cell.fill = PatternFill(start_color="EF4E22", end_color="EF4E22", fill_type="solid")
                    cell.alignment = Alignment(horizontal="center")
                for col in ws.columns:
                    ws.column_dimensions[col[0].column_letter].width = min(max(len(str(cell.value)) for cell in col) + 2, 60)

            # Aba Ativos comparativo
            status_atualizado = reimplementar_status_cadastro(self.cadastro_df)
            indep = calcular_ativos_independentemente(
                self.cadastro_df, self.hierarquia_df, self.ferias_df, status_atualizado
            )
            if not indep.empty and self.cad_rev is not None and not self.cad_rev.empty:
                col_ativos = None
                for c in ["ativos", "Ativos no +TOP", "Ativos"]:
                    if c in self.cad_rev.columns:
                        col_ativos = c
                        break
                if col_ativos:
                    comp = indep.merge(
                        self.cad_rev[["revenda", col_ativos]].rename(columns={col_ativos: "ativos_relatorio"}),
                        on="revenda",
                        how="outer",
                    ).fillna(0)
                    comp["diferenca"] = comp["ativos"] - comp["ativos_relatorio"]
                    comp.to_excel(writer, sheet_name="Ativos_Comparativo", index=False)

            # Aba Desligados inválidos
            if self.hierarquia_df is not None and not self.hierarquia_df.empty:
                col_desligado = None
                for c in self.hierarquia_df.columns:
                    if _normalizar_texto(c) == "DESLIGADO":
                        col_desligado = c
                        break
                if col_desligado:
                    deslig_norm = self.hierarquia_df[col_desligado].fillna("").astype(str).str.strip().str.upper()
                    invalidos = self.hierarquia_df[~deslig_norm.isin(["NAO", "NÃO"])]
                    if not invalidos.empty:
                        invalidos.to_excel(writer, sheet_name="Desligados_Invalidos", index=False)

            # Aba Duplicidades
            if self.hierarquia_df is not None and not self.hierarquia_df.empty:
                duplicados = self.hierarquia_df[self.hierarquia_df.duplicated(subset=["cpf_limp"], keep=False)]
                if not duplicados.empty:
                    duplicados.to_excel(writer, sheet_name="Duplicidades_Hierarquia", index=False)

        self.logger.info(f"Relatório de validação salvo em: {caminho}")
        return caminho


# ---------------------------------------------------------------------------
# ENTRYPOINT STANDALONE (PARA TESTES)
# ---------------------------------------------------------------------------
def main():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)-8s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    logger = logging.getLogger(__name__)

    # Importa as funções do script principal apenas no modo standalone,
    # evitando dependência circular quando importado por ele.
    sys.path.insert(0, str(BASE_DIR))
    from gerar_relatorio_semanal import carregar_bases, calcular_cadastros, calcular_treinamentos, calcular_aceites

    logger.info("Executando validador em modo standalone...")
    bases = carregar_bases()

    df_cad = bases["cadastro"]
    df_trein = bases["treinamentos"]
    df_aceite = bases["aceites"]
    ano_mes = bases["mes_referencia"]

    cad_reg, cad_rev = calcular_cadastros(
        df_cad, bases.get("hierarquia"), bases.get("ferias_hier"), bases.get("status_completo")
    )
    trein_reg, trein_rev, trein_por_curso, cursos_info, base_trein = calcular_treinamentos(
        df_trein, df_cad, ano_mes, bases.get("hierarquia")
    )
    aceite_reg, aceite_rev, mes_aceite_ref, base_aceite = calcular_aceites(
        df_aceite, df_cad, ano_mes, bases["aba_aceite"], bases["usou_ultima_aba"], df_hier=bases.get("hierarquia")
    )

    validador = ValidadorEnvioRelatorio(
        cadastro_df=df_cad,
        hierarquia_df=bases.get("hierarquia"),
        ferias_df=bases.get("ferias_hier"),
        treinamentos_df=df_trein,
        aceites_df=df_aceite,
        cad_reg=cad_reg,
        cad_rev=cad_rev,
        trein_reg=trein_reg,
        trein_rev=trein_rev,
        aceite_reg=aceite_reg,
        aceite_rev=aceite_rev,
        base_trein=base_trein,
        base_aceite=base_aceite,
        mes_referencia=ano_mes,
        mes_aceite=mes_aceite_ref,
        cursos_info=cursos_info,
        logger_validador=logger,
    )

    ok, relatorio = validador.executar(gerar_excel=True)
    logger.info(f"Resultado da validação: {'OK' if ok else 'FALHA'}")
    for r in relatorio["criticos"]:
        logger.error(f"[CRITICO] {r.regra}: {r.mensagem}")
    for r in relatorio["alertas"]:
        logger.warning(f"[ALERTA] {r.regra}: {r.mensagem}")
    for r in relatorio["todos"]:
        if r.tipo == "OK":
            logger.info(f"[OK] {r.regra}: {r.mensagem}")

    if relatorio["caminho_excel"]:
        logger.info(f"Excel de validação: {relatorio['caminho_excel']}")

    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
