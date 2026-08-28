#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
IBnet — Design pass "Clean".
Unifica tokens, afina a tipografia e dessatura a paleta em todos os módulos.
Idempotente: rodar duas vezes não muda mais nada.
"""
import re, sys, os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

FILES = [
    "index.html",
    "comercial/index.html",
    "financeiro/index.html",
    "suporte/index.html",
    "cac/index.html",
    "cac/ativacao.html",
    "logs/index.html",
    "ponto/index.html",
]

ACC = "#d9542e"          # laranja IBnet, versão calma
ACC_H = "#c4491f"        # estado hover/pressed
ACC_RGB = "217,84,46"
PH = "__ACCENT__"        # placeholders p/ proteger o accent das trocas globais
PHH = "__ACCENTH__"

# ── 1. escala tipográfica (só reduz; >=16px) ───────────────────────────────
SIZE_MAP = {
    16: 15, 17: 16, 18: 16, 19: 17, 20: 17, 21: 18, 22: 18, 23: 19, 24: 19,
    25: 20, 26: 20, 27: 21, 28: 22, 30: 22, 32: 23, 34: 24, 36: 26, 38: 27,
    40: 28, 42: 30, 44: 31, 48: 33,
}

# ── 2. paleta dessaturada (referência: módulo Financeiro) ──────────────────
COLORS = [
    # neutros escuros
    ("#0f1117", "#0f1115"), ("#1a1d27", "#161922"), ("#2a2d3e", "#242936"),
    ("#0d1018", "#0c0f15"),
    ("#e2e8f0", "#dfe3ea"), ("#94a3b8", "#8b94a7"),
    ("#1e293b", "#2b3444"), ("#64748b", "#78839a"), ("#f1f5fb", "#f4f6fa"),
    # status — verde
    ("#22c55e", "#6bbf8a"), ("#16a34a", "#4f9e6d"),
    ("rgba(34,197,94,", "rgba(107,191,138,"),
    # status — vermelho
    ("#ef4444", "#cf8080"), ("#dc2626", "#b96a6a"),
    ("rgba(239,68,68,", "rgba(207,128,128,"),
    # status — amarelo / laranja
    ("#f59e0b", "#c6a667"), ("#eab308", "#c6a667"), ("#d97706", "#b58a4a"),
    ("rgba(245,158,11,", "rgba(198,166,103,"), ("rgba(234,179,8,", "rgba(198,166,103,"),
    ("#f97316", "#d08a5a"), ("rgba(249,115,22,", "rgba(208,138,90,"),
    # status — azul / roxo
    ("#3b82f6", "#6b8cce"), ("rgba(59,130,246,", "rgba(107,140,206,"),
    ("#8b5cf6", "#9a8fce"), ("rgba(139,92,246,", "rgba(154,143,206,"),
]

# ── 3. declarações que viram o accent da marca ─────────────────────────────
ACCENT_DECLS = [
    "--accent:#3b82f6", "--accent: #3b82f6",
    "--accent:#E8390E", "--accent: #E8390E",
    "--accent:#e8390e", "--accent: #e8390e",
    "--brand:#E8390E", "--brand: #E8390E",
    "--accent:#6b8cce", "--accent: #6b8cce",
]
# declarações que viram o tom de hover (mais escuro que o accent)
ACCENT_H_DECLS = [
    "--accent2:#CC2200", "--accent2: #CC2200",
    "--accent-hover:#5a7bbd", "--accent-hover: #5a7bbd",
]


def scale_shadow_alpha(text):
    """Suaviza sombras: reduz a opacidade dentro de declarações box-shadow."""
    def soften(decl):
        def a(m):
            v = float(m.group(1))
            return "rgba(0,0,0,%s)" % (("%.2f" % max(0.10, v * 0.55)).rstrip("0").rstrip("."))
        return re.sub(r"rgba\(0,\s*0,\s*0,\s*(0?\.\d+)\)", a, decl)
    return re.sub(r"box-shadow\s*:[^;}\"']+", lambda m: soften(m.group(0)), text)


MARKER = "<!-- ibnet-design:clean-v2 -->"

# ── Etapa 2: cores saturadas residuais (gráficos, badges, KPIs) ────────────
# Só troca de cor -> idempotente (nenhum destino aparece como origem).
RESIDUAL = [
    ("#6366f1", "#8085c9"), ("#4f46e5", "#6b6fbd"), ("#7c3aed", "#8f7bc4"),
    ("#a855f7", "#a181cc"), ("#06b6d4", "#5aa3b5"), ("#0ea5e9", "#5f9cc4"),
    ("#ec4899", "#c47a9b"), ("#f43f5e", "#cf7f8c"), ("#14b8a6", "#5aa89c"),
    ("#10b981", "#5faa8b"), ("#84cc16", "#98b263"), ("#b91c1c", "#a35a5a"),
    ("#991b1b", "#8f5555"), ("#1d4ed8", "#5a78b5"), ("#15803d", "#4d8a63"),
    ("#166534", "#467a58"), ("#92400e", "#8a6540"), ("#a16207", "#97803f"),
    ("#c2410c", "#b06a45"),
    # formas rgba() equivalentes (tints/borders) das mesmas cores
    ("rgba(37,99,235,", "rgba(217,84,46,"),      # azul antigo -> accent da marca
    ("rgba(204,34,0,", "rgba(217,84,46,"),
    ("rgba(99,102,241,", "rgba(128,133,201,"),
    ("rgba(168,85,247,", "rgba(161,129,204,"),
    ("rgba(126,34,206,", "rgba(143,123,196,"),
    ("rgba(236,72,153,", "rgba(196,122,155,"),
    ("rgba(6,182,212,", "rgba(90,163,181,"),
    ("rgba(20,184,166,", "rgba(90,168,156,"),
    ("rgba(16,185,129,", "rgba(95,170,139,"),
    ("rgba(22,163,74,", "rgba(79,158,109,"),
    ("rgba(220,38,38,", "rgba(185,106,106,"),
    ("rgba(185,28,28,", "rgba(163,90,90,"),
    ("rgba(217,119,6,", "rgba(181,138,74,"),
    ("rgba(161,98,7,", "rgba(151,128,63,"),
    ("rgba(30,41,59,", "rgba(43,52,68,"),
    ("rgba(100,116,139,", "rgba(120,131,154,"),
    ("rgba(42,45,62,", "rgba(36,41,54,"),
    ("rgba(15,17,23,", "rgba(15,17,21,"),
    ("rgba(241,245,251,", "rgba(244,246,250,"),
]


def residuals(text):
    for a, b in RESIDUAL:
        text = text.replace(a, b).replace(a.upper(), b)
    return text


def transform(src, path):
    out = src

    # accent -> placeholder (protege das trocas globais de cor)
    for d in ACCENT_DECLS:
        name, _, _ = d.partition(":")
        sep = ": " if ": " in d else ":"
        out = out.replace(d, name + sep + PH)
    for d in ACCENT_H_DECLS:
        name, _, _ = d.partition(":")
        sep = ": " if ": " in d else ":"
        out = out.replace(d, name + sep + PHH)

    # E8390E é a marca em vários lugares hardcoded
    for h in ("#E8390E", "#e8390e"):
        out = out.replace(h, PH)
    out = out.replace("rgba(232,57,14,", "rgba(%s," % ACC_RGB)

    # tints azuis que na verdade eram "accent" (portal + módulos ex-azuis)
    if path in ("index.html", "comercial/index.html", "suporte/index.html"):
        out = out.replace("rgba(59,130,246,", "rgba(%s," % ACC_RGB)

    # ponto: tints neutros (hover de linha, nota, linha de total) ficam neutros
    if path == "ponto/index.html":
        out = out.replace("background:rgba(59,130,246,.08);border:1px solid var(--border)",
                          "background:rgba(255,255,255,.03);border:1px solid var(--border)")
        out = out.replace("tbody tr:hover td{background:rgba(59,130,246,.05)}",
                          "tbody tr:hover td{background:rgba(255,255,255,.03)}")
        out = out.replace(".tot-row td{background:rgba(59,130,246,.07)",
                          ".tot-row td{background:rgba(255,255,255,.04)")

    # tipografia mais fina: teto de peso 600
    out = re.sub(r"font-weight:(\s*)(650|700|800|900)\b", r"font-weight:\g<1>600", out)

    # tipografia menor
    def shrink(m):
        n = int(m.group(2))
        return "font-size:%s%dpx" % (m.group(1), SIZE_MAP.get(n, n))
    out = re.sub(r"font-size:(\s*)(\d+)px", shrink, out)

    # paleta
    for a, b in COLORS:
        out = out.replace(a, b)

    # raio único
    out = re.sub(r"--radius:\s*1[0-9]px", "--radius: 10px", out)

    # gradiente da marca, mais discreto
    out = out.replace("linear-gradient(135deg,#CC2200,#FF5500)",
                      "linear-gradient(135deg,#b8431f,#d9542e)")
    out = out.replace("#CC2200", "#b8431f").replace("#FF5500", "#d9542e")

    # bug: botão de login laranja com hover azul
    out = out.replace(".btn-login:hover{background:#2563eb}",
                      ".btn-login:hover{background:#c4491f}")
    out = out.replace("#2563eb", "#c4491f")

    # sombras suaves
    out = scale_shadow_alpha(out)

    # devolve o accent
    out = out.replace(PHH, ACC_H).replace(PH, ACC)

    # marca o arquivo para que a passagem não seja reaplicada (não é idempotente:
    # a escala de fonte e a suavização de sombra reduziriam de novo a cada run)
    out = out.replace("</head>", "  %s\n</head>" % MARKER, 1)
    return out


def main():
    total = 0
    for rel in FILES:
        p = os.path.join(ROOT, rel)
        with open(p, encoding="utf-8") as fh:
            src = fh.read()
        if MARKER in src:
            new = residuals(src)          # etapa 2 é sempre segura de reaplicar
        else:
            new = residuals(transform(src, rel))
        if new != src:
            with open(p, "w", encoding="utf-8") as fh:
                fh.write(new)
            diff = sum(1 for a, b in zip(src.splitlines(), new.splitlines()) if a != b)
            print("  %-24s %d linhas alteradas" % (rel, diff))
            total += 1
        else:
            print("  %-24s sem mudanças" % rel)
    print("\n%d arquivo(s) atualizados." % total)


if __name__ == "__main__":
    main()
