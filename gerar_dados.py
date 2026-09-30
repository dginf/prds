"""Gera dados/prd.json para o Painel de Indicadores dos PRDs a partir da planilha-base.

Uso:
    python gerar_dados.py Base_Indicadores_PRDs_Series_Observadas_2026.xlsx

A planilha deve conter as abas Indicadores, Metas_anuais, Series_observadas,
Cobertura_calculo, Qualidade, Metodologia_series e Fontes.
"""
import json
import math
import re
import sys
import unicodedata
from pathlib import Path

import pandas as pd

PLANOS = {
    "PRDA": {"nome": "Plano Regional de Desenvolvimento da Amazônia 2024-2027",
             "sigla_sup": "SUDAM", "regiao": "Amazônia Legal", "cor": "#2e7d32"},
    "PRDCO": {"nome": "Plano Regional de Desenvolvimento do Centro-Oeste 2024-2027",
              "sigla_sup": "SUDECO", "regiao": "Centro-Oeste", "cor": "#b6790a"},
    "PRDNE": {"nome": "Plano Regional de Desenvolvimento do Nordeste 2024-2027",
              "sigla_sup": "SUDENE", "regiao": "Nordeste", "cor": "#1565c0"},
}

# Nomes de eixo padronizados por código (o texto varia entre os planos)
EIXO_PADRAO = {
    "1": "Desenvolvimento produtivo",
    "2": "Ciência, tecnologia e inovação",
    "3": "Infraestrutura econômica e urbana",
    "4": "Meio ambiente",
    "5": "Capacidades governativas",
    "6": "Desenvolvimento social",
    "7": "Educação",
}


def nn(v):
    """Converte NaN/NaT em None e numpy em tipos nativos."""
    if v is None:
        return None
    if isinstance(v, float) and math.isnan(v):
        return None
    if hasattr(v, "item"):
        v = v.item()
        if isinstance(v, float) and math.isnan(v):
            return None
    if isinstance(v, str):
        v = re.sub(r"\s+", " ", v).strip()
        return v or None
    return v


def limpa_eixo(txt):
    if txt is None:
        return None
    txt = re.sub(r"\s+\d+$", "", txt.strip())  # remove nº de página colado ("... urbana 148")
    return txt[0].upper() + txt[1:].lower() if txt else txt


def cod_eixo(v):
    v = nn(v)
    if v is None:
        return "M"
    return str(int(float(v)))


def chave_periodo(p):
    """Ordenação cronológica de rótulos como 2022, 2022T4, 2024-07, 2022T2-T4, 2022-recalculado."""
    p = str(p)
    ano = int(p[:4])
    resto = p[4:]
    frac = 0.0
    m = re.match(r"T(\d)", resto)
    if m:
        frac = int(m.group(1)) / 4.0 - 0.01
    m = re.match(r"-(\d{2})$", resto)
    if m:
        frac = int(m.group(1)) / 12.0
    if "recalc" in resto:
        frac = 0.001
    if "T2-T4" in resto:
        frac = 0.0
    return ano + frac


def rotulo_periodo(p):
    p = str(p)
    if "recalc" in p:
        return p[:4] + " (recalc.)"
    return p


def origem(status):
    s = (status or "").lower()
    if "proxy" in s or "limite superior" in s:
        return "Proxy / estimativa"
    if s.startswith("publicado no plano"):
        return "Linha de base do plano"
    if s.startswith("publicado"):
        return "Publicado pela fonte"
    return "Calculado"


def sinal_polaridade(pol, meta_txt):
    pol = (pol or "").lower()
    if "maior" in pol:
        return 1, False
    if "menor" in pol:
        return -1, False
    # PRDNE 5.1.x: polaridade não explicitada, mas a meta é "Atingir X%" acima do nível atual
    if meta_txt and meta_txt.lower().startswith("atingir"):
        return 1, True
    return 0, False


def meta_relativa_prda(txt):
    """'Aumento de 30%' -> 30 ; 'Redução de 40%' -> -40 ; '... pontos percentuais' -> (8, 'p.p.')."""
    if not txt:
        return None, None
    m = re.search(r"(-?\d+[\.,]?\d*)", txt)
    if not m:
        return None, None
    v = float(m.group(1).replace(",", "."))
    if re.search(r"redu|diminu", txt, re.I):
        v = -abs(v)
    tipo = "p.p." if "pontos percentuais" in txt.lower() else "%"
    return v, tipo


def main(xlsx):
    rd = lambda aba, h: pd.read_excel(xlsx, aba, header=h)
    I = rd("Indicadores", 3)
    M = rd("Metas_anuais", 3)
    S = rd("Series_observadas", 3)
    C = rd("Cobertura_calculo", 3)
    Q = rd("Qualidade", 3)
    MT = pd.read_excel(xlsx, "Metodologia_series", header=None)
    FO = rd("Fontes", 2)
    RES = pd.read_excel(xlsx, "Resumo", header=None)

    atualizacao = None
    m = re.search(r"(\d{4}-\d{2}-\d{2})", str(RES.iloc[1, 0]))
    if m:
        atualizacao = m.group(1)

    cob = {r.registro_id: r for r in C.itertuples(index=False)}
    ressalvas_q = {}
    for r in Q.itertuples(index=False):
        ressalvas_q.setdefault(r.plano, []).append(r.achado_id)

    indicadores = []
    for r in I.itertuples(index=False):
        rid = r.registro_id
        c = cob.get(rid)
        ce = cod_eixo(r.codigo_eixo)
        eixo_orig = limpa_eixo(nn(r.eixo)) if ce != "M" else "Macroindicadores"
        eixo_pad = EIXO_PADRAO.get(ce, "Macroindicadores")

        # metas anuais
        mm = M[M.registro_id == rid].sort_values("ano")
        metas = {}
        for x in mm.itertuples(index=False):
            metas[str(int(x.ano))] = {"v": nn(x.valor_meta), "t": nn(x.texto_meta)}

        # série observada
        ss = S[S.registro_id == rid].copy()
        ss["k"] = ss["periodo"].astype(str).map(chave_periodo)
        ss = ss.sort_values("k")
        serie = []
        for x in ss.itertuples(index=False):
            serie.append({
                "p": rotulo_periodo(x.periodo), "k": round(float(x.k), 4), "a": int(x.ano),
                "v": nn(x.valor), "u": nn(x.unidade), "tv": nn(x.tipo_valor),
                "st": nn(x.status_calculo), "o": origem(nn(x.status_calculo)),
                "cmp": nn(x.comparabilidade), "met": nn(x.metodo), "f": nn(x.fonte),
                "url": nn(x.url_fonte), "dx": str(nn(x.data_extracao) or "")[:10], "obs": nn(x.observacao),
            })

        ult = serie[-1] if serie else None
        pol, pol_inferida = sinal_polaridade(nn(r.polaridade), nn(r.meta_quadrienal_texto))
        lb_val = nn(r.linha_base_valor)
        lb_ano = nn(r.linha_base_ano)
        lb_ano = int(lb_ano) if lb_ano is not None else None
        v_ult = nn(c.valor_mais_recente) if c is not None else (ult["v"] if ult else None)
        meta27 = nn(c.meta_2027) if c is not None else None
        dif = nn(c.diferenca_para_meta_2027) if c is not None else None
        direcao = nn(c.direcao_desde_base) if c is not None else None
        var_base = nn(c.variacao_absoluta_desde_base) if c is not None else None

        # --- avaliação frente à meta ---
        aval = {"situacao": "Sem avaliação", "progresso": None, "modo": None, "motivo": None,
                "var_rel": None, "meta_rel": None, "base_calc": None}
        cmp_ult = (ult or {}).get("cmp") or ""
        if r.plano == "PRDA":
            meta_rel, tipo_rel = meta_relativa_prda(nn(r.meta_quadrienal_texto))
            aval["meta_rel"] = meta_rel
            aval["modo"] = "relativa"
            base_obs = next((s for s in serie if s["a"] == lb_ano and s["v"] is not None), None)
            if meta_rel is None:
                aval["motivo"] = "Meta quadrienal sem valor numérico."
            elif base_obs is None:
                aval["motivo"] = (f"O plano não publica o valor da linha de base ({lb_ano}) e a série "
                                  "observada não cobre o ano-base; a variação acumulada não pode ser calculada.")
            elif base_obs is ult:
                aval["motivo"] = "Só há observação no ano-base."
            else:
                vb, vu = base_obs["v"], ult["v"]
                var = (vu - vb) if tipo_rel == "p.p." else (vu / vb - 1) * 100
                aval.update(var_rel=round(var, 2), base_calc={"p": base_obs["p"], "v": vb, "o": base_obs["o"]},
                            unid_rel=tipo_rel)
                prog = var / meta_rel if meta_rel else None
                aval["progresso"] = round(prog, 4) if prog is not None else None
                if pol * (var - meta_rel) >= 0:
                    aval["situacao"] = "Meta alcançada"
                elif pol * var > 0:
                    aval["situacao"] = "Avançando"
                else:
                    aval["situacao"] = "Em retrocesso"
                aval["motivo"] = ("Variação acumulada calculada pelo painel entre o valor observado no ano-base "
                                  f"({base_obs['p']}) e o período mais recente, comparada à meta quadrienal relativa.")
        else:
            aval["modo"] = "nivel"
            meta_ref, ano_ref = meta27, 2027
            if meta27 is None and nn(r.meta_2030) is not None:
                # IDH do PRDCO: única meta numérica é 2030
                meta_ref, ano_ref = nn(r.meta_2030), 2030
                dif = round(v_ult - meta_ref, 6) if v_ult is not None else None
            aval["ano_meta"] = ano_ref
            aval["meta_ref"] = meta_ref
            comparavel = meta_ref is not None and dif is not None and not cmp_ult.lower().startswith("não comparável")
            so_base = ult is not None and lb_ano is not None and ult["a"] <= lb_ano
            if comparavel and so_base:
                comparavel = False
                aval["motivo"] = ("Só há observação do próprio ano-base; ainda não existe dado posterior "
                                  "para avaliar a trajetória.")
            elif not comparavel:
                if meta_ref is None:
                    aval["motivo"] = "O plano não define meta numérica para 2027."
                else:
                    aval["motivo"] = ("A observação mais recente não é comparável em nível à meta "
                                      "(ver comparabilidade); diferença para a meta não calculada.")
            elif pol == 0:
                aval["motivo"] = "Polaridade não explicitada."
            else:
                if pol * (v_ult - meta_ref) >= 0:
                    aval["situacao"] = "Meta alcançada"
                elif var_base is not None and abs(var_base) < 1e-9:
                    aval["situacao"] = "Estável"
                elif direcao == "Alinhada":
                    aval["situacao"] = "Avançando"
                elif direcao == "Não alinhada":
                    aval["situacao"] = "Em retrocesso"
                else:
                    aval["situacao"] = "Aquém da meta, sem linha de base"
                if lb_val is not None and meta_ref != lb_val:
                    aval["progresso"] = round((v_ult - lb_val) / (meta_ref - lb_val), 4)
                if ano_ref == 2030:
                    aval["motivo"] = "O plano fixa apenas a meta de 2030 (0,800) para este macroindicador."

        ressalva = bool(ult) and (
            ult["o"] == "Proxy / estimativa" or "ressalva" in cmp_ult.lower() or "não comparável" in cmp_ult.lower()
            or "divergência" in cmp_ult.lower() or "proxy" in cmp_ult.lower())

        indicadores.append({
            "id": rid, "plano": r.plano, "nivel": nn(r.nivel),
            "eixo_cod": ce, "eixo": eixo_pad, "eixo_plano": eixo_orig,
            "prog_cod": nn(str(r.codigo_programa)) if nn(r.codigo_programa) is not None else None,
            "programa": nn(r.programa), "objetivo": nn(r.objetivo),
            "cod": nn(r.codigo_indicador), "cod_fonte": nn(r.codigo_indicador_fonte),
            "nome": nn(r.indicador), "recorte": nn(r.recorte),
            "un": nn(r.unidade_indicador), "un_meta": nn(r.unidade_meta),
            "pol": nn(r.polaridade), "pol_s": pol, "pol_inferida": pol_inferida,
            "fonte_prev": nn(r.fonte_indicador), "periodicidade": nn(r.periodicidade),
            "tipo_meta": nn(r.tipo_meta),
            "lb_ano": lb_ano, "lb_val": lb_val, "lb_txt": nn(r.linha_base_texto), "lb_status": nn(r.status_linha_base),
            "meta_txt": nn(r.meta_quadrienal_texto), "meta2030": nn(r.meta_2030), "meta27": meta27,
            "metas": metas, "n_acoes": nn(r.qtd_acoes_associadas), "acoes": nn(r.acoes_estrategicas_ids),
            "obs_q": nn(r.observacoes_qualidade), "pag": nn(r.pagina_pdf), "url_doc": nn(r.url_oficial),
            "n_obs": len(serie), "p_ini": serie[0]["p"] if serie else None, "p_ult": ult["p"] if ult else None,
            "a_ult": ult["a"] if ult else None,
            "v_ult": v_ult, "o_ult": ult["o"] if ult else None, "cmp_ult": cmp_ult or None,
            "var_base": var_base, "direcao": direcao, "dif_meta": dif if r.plano != "PRDA" else None,
            "limitacao": nn(c.limitacao) if c is not None else None,
            "fonte_ult": nn(c.fonte_mais_recente) if c is not None else None,
            "url_ult": nn(c.url_fonte_mais_recente) if c is not None else None,
            "ressalva": ressalva, "aval": aval, "serie": serie,
        })

    qualidade = [{k: nn(v) for k, v in r._asdict().items()} for r in Q.itertuples(index=False)]

    # Metodologia: bloco de decisões e bloco de defasagem das fontes
    decisoes, defasagem, modo = [], [], "dec"
    for row in MT.itertuples(index=False):
        a, b, c3 = nn(row[0]), nn(row[1]), nn(row[2])
        if a is None or a in ("tema", "Metodologia das séries observadas"):
            continue
        if a == "Defasagem das principais fontes":
            modo = "def"
            continue
        if a == "Fonte" and b == "Indicadores atualizados":
            continue
        (decisoes if modo == "dec" else defasagem).append([a, b, c3])

    fontes = []
    for r in FO.itertuples(index=False):
        d = {k: nn(v) for k, v in r._asdict().items()}
        d.pop("arquivo_local", None)  # caminho local do autor; não publicar
        if d.get("plano") or d.get("documento"):
            fontes.append(d)

    out = {
        "meta": {"atualizacao": atualizacao, "gerado_de": Path(xlsx).name},
        "planos": PLANOS, "eixos": EIXO_PADRAO,
        "ind": indicadores, "qualidade": qualidade,
        "metodologia": decisoes, "defasagem": defasagem, "fontes": fontes,
    }
    Path("dados").mkdir(exist_ok=True)
    with open("dados/prd.json", "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, separators=(",", ":"))
    print(f"dados/prd.json: {len(indicadores)} registros, "
          f"{sum(len(i['serie']) for i in indicadores)} observações")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    main(sys.argv[1])
