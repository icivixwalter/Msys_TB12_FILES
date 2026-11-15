# -*- coding: utf-8 -*-
# ==================================================================================================
#  md2html.py - converte i file .md di una cartella in pagine .html con lo stesso stile
#  USO      : python md2html.py "<cartella con i .md>"     (lo lancia GENERA_HTML.bat)
#  SUPPORTA : # titoli, paragrafi, elenchi - e 1., tabelle | a | b |, blocchi ``` ```, `codice`,
#             **grassetto**, *corsivo*, [testo](link), righe ---, celle con \| per il carattere |
#  NESSUNA LIBRERIA ESTERNA: basta Python 3.
#  INDICE
#  #01# IMPOSTAZIONI = stile della pagina
#  #02# TESTO_IN_RIGA = escape HTML, codice, grassetto, corsivo, collegamenti
#  #03# BLOCCHI = conversione riga per riga (titoli, elenchi, tabelle, codice, paragrafi)
#  #04# PAGINA = modello HTML con menu delle schede
#  #05# CICLO = tutti i .md della cartella -> .html
# ==================================================================================================
import html, os, re, sys

# --- #01# IMPOSTAZIONI = stile della pagina (tema chiaro e scuro) --------------------------------
CSS = """
:root{--bg:#f4f5f7;--surface:#fff;--surface2:#eceff3;--fg:#1a202b;--muted:#5b6474;--line:#d6dbe3;--accent:#962c3b;--code:#eef1f5}
@media (prefers-color-scheme:dark){:root{--bg:#11151c;--surface:#181e27;--surface2:#202834;--fg:#e2e6ed;--muted:#98a2b3;--line:#2b3340;--accent:#e57c8b;--code:#202834;color-scheme:dark}}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--fg);font:15px/1.6 "Segoe UI",system-ui,sans-serif}
.wrap{max-width:1000px;margin:0 auto;padding:28px max(16px,env(safe-area-inset-left)) 64px}
nav{display:flex;flex-wrap:wrap;gap:6px 16px;font-size:13px;padding-bottom:10px;border-bottom:1px solid var(--line);margin-bottom:24px}
nav a{color:var(--muted);text-decoration:none} nav a.att{color:var(--accent);font-weight:600}
h1{font-size:30px;line-height:1.2;margin:8px 0 12px} h2{font-size:21px;margin:34px 0 10px;padding-top:12px;border-top:1px solid var(--line)}
h3{font-size:17px;margin:22px 0 8px} h1,h2,h3{text-wrap:balance}
p{margin:0 0 12px;max-width:80ch} ul,ol{margin:0 0 12px;padding-left:22px} li{margin:3px 0}
a{color:var(--accent)}
code{font-family:Consolas,"Cascadia Mono",monospace;font-size:.88em;background:var(--code);padding:1px 5px;border-radius:3px;overflow-wrap:anywhere}
pre{background:var(--surface);border:1px solid var(--line);border-radius:6px;padding:12px 14px;overflow-x:auto;font:13px/1.5 Consolas,"Cascadia Mono",monospace}
pre code{background:none;padding:0;font-size:inherit}
.tbl{overflow-x:auto;border:1px solid var(--line);border-radius:6px;background:var(--surface);margin:0 0 14px}
table{border-collapse:collapse;width:100%;font-size:14px} th,td{text-align:left;vertical-align:top;padding:8px 11px;border-bottom:1px solid var(--line)}
th{background:var(--surface2);font-size:12px;text-transform:uppercase;letter-spacing:.05em;color:var(--muted)} tr:last-child td{border-bottom:0}
hr{border:0;border-top:1px solid var(--line);margin:24px 0}
footer{margin-top:40px;font-size:12px;color:var(--muted)}
"""

# --- #02# TESTO_IN_RIGA ----------------------------------------------------------------------------
def in_riga(t):
    parti = re.split(r'(`[^`]*`)', t)                    # il codice in riga non si tocca
    out = []
    for p in parti:
        if p.startswith('`') and p.endswith('`') and len(p) >= 2:
            out.append('<code>' + html.escape(p[1:-1].replace('\\|', '|')) + '</code>')
        else:
            s = html.escape(p.replace('\\|', '|'))
            s = re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', s)
            s = re.sub(r'(?<!\*)\*(?!\s)(.+?)(?<!\s)\*(?!\*)', r'<em>\1</em>', s)
            s = re.sub(r'\[([^\]]+)\]\(([^)]+)\)', r'<a href="\2">\1</a>', s)
            out.append(s)
    return ''.join(out)

def celle(riga):
    riga = riga.strip().strip('|')
    return [c.strip() for c in re.split(r'(?<!\\)\|', riga)]

def slug(t):
    return re.sub(r'[^a-z0-9]+', '-', t.lower()).strip('-')

# --- #03# BLOCCHI ----------------------------------------------------------------------------------
def converti(md):
    righe = md.replace('\r\n', '\n').split('\n')
    out, i, titolo = [], 0, ''
    while i < len(righe):
        r = righe[i]
        # #03.01# blocco di codice
        if r.strip().startswith('```'):
            i += 1; buf = []
            while i < len(righe) and not righe[i].strip().startswith('```'):
                buf.append(righe[i]); i += 1
            out.append('<pre><code>' + html.escape('\n'.join(buf)) + '</code></pre>'); i += 1; continue
        # #03.02# titoli
        m = re.match(r'^(#{1,4})\s+(.*)$', r)
        if m:
            n = len(m.group(1)); t = m.group(2).strip()
            if n == 1 and not titolo: titolo = re.sub(r'[`*]', '', t)
            out.append('<h%d id="%s">%s</h%d>' % (n, slug(t), in_riga(t), n)); i += 1; continue
        # #03.03# riga orizzontale
        if re.match(r'^\s*(-{3,}|\*{3,})\s*$', r):
            out.append('<hr>'); i += 1; continue
        # #03.04# tabella: riga | ... | seguita da | --- |
        if r.strip().startswith('|') and i + 1 < len(righe) and re.match(r'^\s*\|?\s*:?-{3,}', righe[i + 1]):
            intest = celle(r); i += 2; corpo = []
            while i < len(righe) and righe[i].strip().startswith('|'):
                corpo.append(celle(righe[i])); i += 1
            h = '<div class="tbl"><table><thead><tr>' + ''.join('<th>%s</th>' % in_riga(c) for c in intest) + '</tr></thead><tbody>'
            for c in corpo:
                h += '<tr>' + ''.join('<td>%s</td>' % in_riga(x) for x in c) + '</tr>'
            out.append(h + '</tbody></table></div>'); continue
        # #03.05# elenchi puntati e numerati
        m = re.match(r'^\s*([-*]|\d+\.)\s+(.*)$', r)
        if m:
            num = m.group(1)[0].isdigit(); tag = 'ol' if num else 'ul'; voci = []
            while i < len(righe):
                m2 = re.match(r'^\s*([-*]|\d+\.)\s+(.*)$', righe[i])
                if m2 and (m2.group(1)[0].isdigit()) == num:
                    voci.append(m2.group(2)); i += 1
                elif voci and righe[i].startswith('  ') and righe[i].strip():
                    voci[-1] += ' ' + righe[i].strip(); i += 1
                else:
                    break
            out.append('<%s>%s</%s>' % (tag, ''.join('<li>%s</li>' % in_riga(v) for v in voci), tag)); continue
        # #03.06# riga vuota
        if not r.strip():
            i += 1; continue
        # #03.07# paragrafo: righe consecutive
        buf = [r.strip()]; i += 1
        while i < len(righe) and righe[i].strip() and not re.match(r'^(#{1,4}\s|\s*[-*]\s|\s*\d+\.\s|\s*\||\s*```)', righe[i]):
            buf.append(righe[i].strip()); i += 1
        out.append('<p>' + in_riga(' '.join(buf)) + '</p>')
    return titolo, '\n'.join(out)

# --- #04# PAGINA -----------------------------------------------------------------------------------
def pagina(titolo, corpo, menu, attuale):
    voci = ''.join('<a href="%s"%s>%s</a>' % (f, ' class="att"' if f == attuale else '', html.escape(t)) for f, t in menu)
    return ('<!doctype html>\n<html lang="it">\n<head>\n<meta charset="utf-8">\n'
            '<meta name="viewport" content="width=device-width, initial-scale=1">\n'
            '<title>%s</title>\n<style>%s</style>\n</head>\n<body>\n<div class="wrap">\n<nav>%s</nav>\n%s\n'
            '<footer>Pagina generata da STRUMENTI\\md2html.py a partire dal file .md con lo stesso nome: per modificarla cambia il .md e lancia GENERA_HTML.bat.</footer>\n'
            '</div>\n</body>\n</html>\n') % (html.escape(titolo), CSS, voci, corpo)

# --- #05# CICLO ------------------------------------------------------------------------------------
def main():
    cart = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(os.path.abspath(__file__)), '..')
    cart = os.path.abspath(cart)
    files = sorted(f for f in os.listdir(cart) if f.lower().endswith('.md'))
    if not files:
        print('Nessun file .md in', cart); return 1
    pagine = []
    for f in files:
        with open(os.path.join(cart, f), encoding='utf-8-sig') as h:
            t, c = converti(h.read())
        pagine.append((f, t or f[:-3], c))
    menu = [(f[:-3] + '.html', t.replace('MSys_TB12_', '').replace('_', ' ') if len(t) > 28 else t) for f, t, c in pagine]
    for f, t, c in pagine:
        dest = os.path.join(cart, f[:-3] + '.html')
        with open(dest, 'w', encoding='utf-8', newline='\r\n') as h:
            h.write(pagina(t, c, menu, f[:-3] + '.html'))
        print('creato:', dest)
    return 0

if __name__ == '__main__':
    sys.exit(main())
