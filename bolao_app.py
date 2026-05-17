#!/usr/bin/env python3
"""
Bolão Copa do Mundo 2026 — Turim MFO
Dashboard v5 · Paleta Turim/Tori · Sidebar nativa
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from openpyxl import load_workbook
import os, glob, re, base64
from collections import OrderedDict
from datetime import date
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

if "APP_PASSWORD" in st.secrets:
    _PWD = st.secrets["APP_PASSWORD"]
else:
    _PWD = os.getenv("APP_PASSWORD")

# ══════════════════════════════════════════════════════════════════════
# PAGE CONFIG — sidebar nativa funciona, 3 pontinhos aparecem
# ══════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Bolão Copa 2026 · Turim MFO",
    page_icon="⚽",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ══════════════════════════════════════════════════════════════════════
# PALETA TURIM (extraída do PPTX corporativo)
# Primary: #0D2B40 (navy escuro), #123A56 (navy médio)
# Accent:  #D6B864 (gold), #DC884A (laranja), #B2584E (vermelho)
# Tori:    #0D8587 (teal)
# Text:    #4D4D4F / #FFFFFF
# Fontes:  Effra, Gotham (fallback: Inter)
# ══════════════════════════════════════════════════════════════════════

# CSS — only styles custom HTML divs, inherits text for dark/light compat
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:ital,wght@0,300;0,400;0,600;0,700;0,800;0,900;1,400&display=swap');

/* Fonte global sem quebrar sidebar */
html, body, [class*="css"] {
    font-family: 'Inter', 'Effra', 'Gotham Book', sans-serif !important;
}

/* Remove padding top excessivo */
.block-container { padding-top: 1.2rem !important; max-width: 1400px; }

/* ── Hero banner */
.hero {
    background: linear-gradient(135deg, #0D2B40 0%, #123A56 50%, #0D8587 100%);
    border-radius: 16px; padding: 28px 40px; margin-bottom: 20px;
    position: relative; overflow: hidden;
    box-shadow: 0 6px 28px rgba(13,43,64,.35);
}
.hero::after {
    content: ""; position: absolute; right: -20px; top: -20px;
    width: 200px; height: 200px; border-radius: 50%;
    background: rgba(214,184,100,.12);
}
.hero-title {
    font-size: 1.85rem; font-weight: 900; color: #FFFFFF !important;
    margin: 0; letter-spacing: -.5px;
}
.hero-sub { font-size: .88rem; color: rgba(255,255,255,.72) !important; margin-top: 5px; }
.hero-badge {
    display: inline-block; background: rgba(214,184,100,.25);
    border: 1px solid rgba(214,184,100,.5); border-radius: 20px;
    padding: 3px 12px; font-size: .75rem; color: #D6B864 !important; margin-top: 10px;
}

/* ── Metric mini-card */
.mc {
    border: 1px solid rgba(128,128,128,.15);
    border-radius: 12px; padding: 14px 16px; text-align: center;
}
.mc-v { font-size: 1.7rem; font-weight: 900; color: #D6B864 !important; line-height: 1; }
.mc-l { font-size: .63rem; color: inherit; opacity: .6; text-transform: uppercase;
         letter-spacing: 1px; margin-top: 4px; }

/* ── Ranking card */
.rc {
    border-radius: 11px; padding: 12px 16px; margin-bottom: 6px;
    display: flex; align-items: center; gap: 12px;
    border: 1px solid rgba(128,128,128,.15);
}
.rc1 { border-left: 4px solid #D6B864; }
.rc2 { border-left: 4px solid #7F7F7F; }
.rc3 { border-left: 4px solid #DC884A; }
.rcN { border-left: 4px solid #123A56; }
.rc-name { font-size: .94rem; font-weight: 700; flex: 1; }
.rc-sub  { font-size: .67rem; opacity: .6; margin-top: 2px; }
.rc-pts  { font-size: 1.55rem; font-weight: 900; color: #D6B864 !important; }
.rc-pl   { font-size: .62rem; color: #5C5F62; text-align: right; }
.bar-bg  { height: 3px; border-radius: 3px; background: rgba(18,58,86,.12); margin-top: 5px; }
.bar-fg  { height: 3px; border-radius: 3px;
           background: linear-gradient(90deg, #123A56, #D6B864); }

/* ── Group table */
.gb {
    border: 1px solid rgba(128,128,128,.12); border-radius: 11px;
    padding: 11px 12px; margin-bottom: 9px;
}
.gb-hdr { font-weight: 800; font-size: .88rem; margin-bottom: 6px; }
.gt { width: 100%; border-collapse: collapse; font-size: .76rem; }
.gt th { font-weight: 700; text-align: center; padding: 4px 4px;
          background: #0D2B40; color: #FFFFFF !important;
          border-bottom: 2px solid rgba(214,184,100,.3); }
.gt td { padding: 4px 4px; text-align: center; color: inherit; }
.gt td.nm { text-align: left; font-weight: 600; white-space: nowrap; }
.row-q1 { background: rgba(13,133,135,.14); }
.row-q2 { background: rgba(13,133,135,.07); }
.row-q3 { background: rgba(220,136,74,.12); }
.row-q4 { background: rgba(178,88,78,.10); }
.dot { display: inline-block; width: 7px; height: 7px; border-radius: 50%;
       margin-right: 4px; vertical-align: middle; }

/* ── Score badges */
.b5 { background: #14532D; color: #86EFAC; border-radius: 4px; padding: 1px 7px; font-size: .7rem; font-weight: 700; }
.b3 { background: #0D4040; color: #5EEAD4; border-radius: 4px; padding: 1px 7px; font-size: .7rem; font-weight: 700; }
.b2 { background: #7C2D12; color: #FED7AA; border-radius: 4px; padding: 1px 7px; font-size: .7rem; font-weight: 700; }
.b0 { background: #7F1D1D; color: #FECACA; border-radius: 4px; padding: 1px 7px; font-size: .7rem; font-weight: 700; }
.bN { border-radius: 4px; padding: 1px 7px; font-size: .7rem; opacity: .5; }

/* ── Match row (group detail) */
.mr {
    border: 1px solid rgba(128,128,128,.10); border-radius: 8px;
    padding: 6px 10px; margin-bottom: 3px;
    display: flex; align-items: center; gap: 8px; font-size: .78rem;
}
.mr-t { flex: 1; font-weight: 600; }
.mr-s { font-size: .73rem; color: #5C5F62; white-space: nowrap; }

/* ── MM table */
.mm-wrap { overflow-x: auto; }
.mm-tbl { border-collapse: collapse; font-size: .78rem; width: 100%; min-width: 500px; }
.mm-tbl th {
    background: #0D2B40; color: #FFFFFF !important;
    padding: 7px 9px; text-align: center; border: 1px solid rgba(255,255,255,.1);
    white-space: nowrap; font-weight: 700;
}
.mm-tbl th.nm { text-align: left; min-width: 110px; }
.mm-tbl td { padding: 6px 9px; text-align: center; white-space: nowrap;
             color: inherit; border: 1px solid rgba(128,128,128,.10); }
.mm-tbl td.nm { text-align: left; font-weight: 600; }
.mm-tbl tr:nth-child(even) td { background: rgba(18,58,86,.04); }
.mm-sub { font-size: .66rem; opacity: .6; line-height: 1.3; }

/* ── MM detail row */
.mmr {
    border: 1px solid rgba(128,128,128,.10); border-radius: 8px;
    padding: 7px 12px; margin-bottom: 4px; font-size: .80rem;
}
.mmr-id { font-weight: 700; color: #0D8587 !important; display: inline-block; min-width: 50px; }

/* ── Bonus card */
.bc {
    border: 1px solid rgba(128,128,128,.12); border-radius: 11px;
    padding: 16px; text-align: center;
}
.bc-lbl { font-size: .63rem; color: #5C5F62 !important; text-transform: uppercase;
           letter-spacing: 1.1px; margin-bottom: 5px; }
.bc-ico  { font-size: 1.5rem; }
.bc-bet  { font-size: .88rem; font-weight: 600; margin: 4px 0 2px; }
.bc-real { font-size: .76rem; opacity: .6; }
.bc-pts  { font-size: 1.35rem; font-weight: 900; color: #D6B864 !important; margin-top: 6px; }

/* ── Section header */
.sh {
    font-size: 1.02rem; font-weight: 800;
    border-bottom: 2px solid #0D2B40; padding-bottom: 5px; margin: 14px 0 10px;
}

/* ── Sidebar logo container */
.sb-logo { text-align: center; padding: 8px 0 4px; }
.sb-divider { border: none; border-top: 1px solid rgba(18,58,86,.15); margin: 8px 0; }

/* ── Mascote strip */
.mascote-strip {
    background: linear-gradient(135deg, #0D2B40, #123A56);
    border-radius: 12px; padding: 10px; text-align: center; margin-bottom: 14px;
}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════
# HELPERS — logo como base64
# ══════════════════════════════════════════════════════════════════════
def img_to_b64(path):
    try:
        with open(path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except:
        return ""

# Paths dos assets (ao lado do script ou em subpasta assets/)
SCRIPT_DIR = Path(__file__).parent
def find_asset(name):
    for p in [SCRIPT_DIR/name, SCRIPT_DIR/"assets"/name, Path(name)]:
        if p.exists(): return str(p)
    return ""

LOGO_TURIM_BRANCA = find_asset("logo_turim_branca.png")
LOGO_TURIM_AZUL   = find_asset("logo_turim_azul.png")
LOGO_TORI         = find_asset("logo_tori.png")
MASCOTES          = find_asset("mascotes.png")

# ══════════════════════════════════════════════════════════════════════
# LOGIN GATE
# ══════════════════════════════════════════════════════════════════════
#_PWD = os.getenv("APP_PASSWORD")
if "auth" not in st.session_state:
    st.session_state.auth = False

if not st.session_state.auth:
    st.markdown("""<style>
      section[data-testid="stSidebar"]{display:none!important;}
      #MainMenu,footer,header{visibility:hidden;}
    </style>""", unsafe_allow_html=True)
    _, mcol, _ = st.columns([1,2,1])
    with mcol:
        st.markdown("<div style='height:30px'></div>", unsafe_allow_html=True)
        if MASCOTES:
            st.image(MASCOTES, width='stretch')
        st.markdown(
            "<div style='text-align:center;margin:10px 0 24px'>"
            "<div style='font-size:1.75rem;font-weight:900'>⚽ Bolão Copa 2026</div>"
            "<div style='font-size:.9rem;opacity:.6;margin-top:4px'>Turim MFO · Acesso Restrito</div>"
            "</div>", unsafe_allow_html=True)
        pwd = st.text_input("🔑 Senha", type="password",
                            placeholder="Digite a senha do bolão...")
        if st.button("Entrar →", type="primary", width='stretch'):
            if pwd == _PWD:
                st.session_state.auth = True
                st.rerun()
            else:
                st.error("Senha incorreta. Tente novamente.")
    st.stop()

# ══════════════════════════════════════════════════════════════════════
# CONSTANTES
# ══════════════════════════════════════════════════════════════════════
FLAGS={
    'Mexico':'🇲🇽','Africa do Sul':'🇿🇦','Coreia do Sul':'🇰🇷','Tchequia':'🇨🇿',
    'Canada':'🇨🇦','Bosnia e Herzegovina':'🇧🇦','Qatar':'🇶🇦','Suica':'🇨🇭',
    'Brasil':'🇧🇷','Marrocos':'🇲🇦','Haiti':'🇭🇹','Escocia':'🏴󠁧󠁢󠁳󠁣󠁴󠁿',
    'Estados Unidos':'🇺🇸','Paraguai':'🇵🇾','Australia':'🇦🇺','Turquia':'🇹🇷',
    'Alemanha':'🇩🇪','Curacao':'🇨🇼','Costa do Marfim':'🇨🇮','Equador':'🇪🇨',
    'Holanda':'🇳🇱','Japao':'🇯🇵','Suecia':'🇸🇪','Tunisia':'🇹🇳',
    'Belgica':'🇧🇪','Egito':'🇪🇬','Ira':'🇮🇷','Nova Zelandia':'🇳🇿',
    'Espanha':'🇪🇸','Cabo Verde':'🇨🇻','Arabia Saudita':'🇸🇦','Uruguai':'🇺🇾',
    'Franca':'🇫🇷','Senegal':'🇸🇳','Iraque':'🇮🇶','Noruega':'🇳🇴',
    'Argentina':'🇦🇷','Argelia':'🇩🇿','Austria':'🇦🇹','Jordania':'🇯🇴',
    'Portugal':'🇵🇹','RD Congo':'🇨🇩','Uzbequistao':'🇺🇿','Colombia':'🇨🇴',
    'Inglaterra':'🏴󠁧󠁢󠁥󠁮󠁧󠁿','Croacia':'🇭🇷','Gana':'🇬🇭','Panama':'🇵🇦',
}
def F(t): return FLAGS.get(str(t),'🌍') if t and str(t) not in ('?','') else ''

GRP_COLORS={
    'A':'#123A56','B':'#0D2B40','C':'#0D8587','D':'#1a6b6d',
    'E':'#B2584E','F':'#8B3D35','G':'#DC884A','H':'#a86028',
    'I':'#D6B864','J':'#b89640','K':'#0D2B40','L':'#123A56',
}
FIFA_RANKINGS={
    'Mexico':16,'Africa do Sul':63,'Coreia do Sul':23,'Tchequia':38,
    'Canada':44,'Bosnia e Herzegovina':55,'Qatar':38,'Suica':20,
    'Brasil':5,'Marrocos':14,'Haiti':86,'Escocia':27,
    'Estados Unidos':11,'Paraguai':51,'Australia':23,'Turquia':36,
    'Alemanha':12,'Curacao':83,'Costa do Marfim':41,'Equador':39,
    'Holanda':8,'Japao':15,'Suecia':26,'Tunisia':30,
    'Belgica':3,'Egito':36,'Ira':25,'Nova Zelandia':98,
    'Espanha':4,'Cabo Verde':88,'Arabia Saudita':57,'Uruguai':17,
    'Franca':2,'Senegal':21,'Iraque':65,'Noruega':28,
    'Argentina':1,'Argelia':32,'Austria':24,'Jordania':86,
    'Portugal':6,'RD Congo':57,'Uzbequistao':74,'Colombia':18,
    'Inglaterra':9,'Croacia':10,'Gana':52,'Panama':72,
}
GROUPS_DATA=OrderedDict([
    ('A',['Mexico','Africa do Sul','Coreia do Sul','Tchequia']),
    ('B',['Canada','Bosnia e Herzegovina','Qatar','Suica']),
    ('C',['Brasil','Marrocos','Haiti','Escocia']),
    ('D',['Estados Unidos','Paraguai','Australia','Turquia']),
    ('E',['Alemanha','Curacao','Costa do Marfim','Equador']),
    ('F',['Holanda','Japao','Suecia','Tunisia']),
    ('G',['Belgica','Egito','Ira','Nova Zelandia']),
    ('H',['Espanha','Cabo Verde','Arabia Saudita','Uruguai']),
    ('I',['Franca','Senegal','Iraque','Noruega']),
    ('J',['Argentina','Argelia','Austria','Jordania']),
    ('K',['Portugal','RD Congo','Uzbequistao','Colombia']),
    ('L',['Inglaterra','Croacia','Gana','Panama']),
])
GL=list(GROUPS_DATA.keys())
GROUP_FIXTURES=[
    (date(2026,6,11),'A','Mexico','Africa do Sul'),(date(2026,6,11),'A','Coreia do Sul','Tchequia'),
    (date(2026,6,18),'A','Tchequia','Africa do Sul'),(date(2026,6,18),'A','Mexico','Coreia do Sul'),
    (date(2026,6,24),'A','Tchequia','Mexico'),(date(2026,6,24),'A','Africa do Sul','Coreia do Sul'),
    (date(2026,6,12),'B','Canada','Bosnia e Herzegovina'),(date(2026,6,13),'B','Qatar','Suica'),
    (date(2026,6,18),'B','Suica','Bosnia e Herzegovina'),(date(2026,6,18),'B','Canada','Qatar'),
    (date(2026,6,24),'B','Suica','Canada'),(date(2026,6,24),'B','Bosnia e Herzegovina','Qatar'),
    (date(2026,6,13),'C','Brasil','Marrocos'),(date(2026,6,13),'C','Haiti','Escocia'),
    (date(2026,6,19),'C','Escocia','Marrocos'),(date(2026,6,19),'C','Brasil','Haiti'),
    (date(2026,6,24),'C','Escocia','Brasil'),(date(2026,6,24),'C','Marrocos','Haiti'),
    (date(2026,6,12),'D','Estados Unidos','Paraguai'),(date(2026,6,13),'D','Australia','Turquia'),
    (date(2026,6,19),'D','Estados Unidos','Australia'),(date(2026,6,19),'D','Turquia','Paraguai'),
    (date(2026,6,25),'D','Turquia','Estados Unidos'),(date(2026,6,25),'D','Paraguai','Australia'),
    (date(2026,6,14),'E','Alemanha','Curacao'),(date(2026,6,14),'E','Costa do Marfim','Equador'),
    (date(2026,6,20),'E','Alemanha','Costa do Marfim'),(date(2026,6,20),'E','Equador','Curacao'),
    (date(2026,6,25),'E','Equador','Alemanha'),(date(2026,6,25),'E','Curacao','Costa do Marfim'),
    (date(2026,6,14),'F','Holanda','Japao'),(date(2026,6,14),'F','Suecia','Tunisia'),
    (date(2026,6,20),'F','Holanda','Suecia'),(date(2026,6,20),'F','Tunisia','Japao'),
    (date(2026,6,25),'F','Japao','Suecia'),(date(2026,6,25),'F','Tunisia','Holanda'),
    (date(2026,6,15),'G','Belgica','Egito'),(date(2026,6,15),'G','Ira','Nova Zelandia'),
    (date(2026,6,21),'G','Belgica','Ira'),(date(2026,6,21),'G','Nova Zelandia','Egito'),
    (date(2026,6,26),'G','Egito','Ira'),(date(2026,6,26),'G','Nova Zelandia','Belgica'),
    (date(2026,6,15),'H','Espanha','Cabo Verde'),(date(2026,6,15),'H','Arabia Saudita','Uruguai'),
    (date(2026,6,21),'H','Espanha','Arabia Saudita'),(date(2026,6,21),'H','Uruguai','Cabo Verde'),
    (date(2026,6,26),'H','Cabo Verde','Arabia Saudita'),(date(2026,6,26),'H','Uruguai','Espanha'),
    (date(2026,6,16),'I','Franca','Senegal'),(date(2026,6,16),'I','Iraque','Noruega'),
    (date(2026,6,22),'I','Franca','Iraque'),(date(2026,6,22),'I','Noruega','Senegal'),
    (date(2026,6,26),'I','Noruega','Franca'),(date(2026,6,26),'I','Senegal','Iraque'),
    (date(2026,6,16),'J','Argentina','Argelia'),(date(2026,6,16),'J','Austria','Jordania'),
    (date(2026,6,22),'J','Argentina','Austria'),(date(2026,6,22),'J','Jordania','Argelia'),
    (date(2026,6,27),'J','Argelia','Austria'),(date(2026,6,27),'J','Jordania','Argentina'),
    (date(2026,6,17),'K','Portugal','RD Congo'),(date(2026,6,17),'K','Uzbequistao','Colombia'),
    (date(2026,6,23),'K','Portugal','Uzbequistao'),(date(2026,6,23),'K','Colombia','RD Congo'),
    (date(2026,6,27),'K','Colombia','Portugal'),(date(2026,6,27),'K','RD Congo','Uzbequistao'),
    (date(2026,6,17),'L','Inglaterra','Croacia'),(date(2026,6,17),'L','Gana','Panama'),
    (date(2026,6,23),'L','Inglaterra','Gana'),(date(2026,6,23),'L','Panama','Croacia'),
    (date(2026,6,27),'L','Panama','Inglaterra'),(date(2026,6,27),'L','Croacia','Gana'),
]
NG=len(GROUP_FIXTURES)
KO=[
    ('R32','M73',('r2','A'),('r2','B')),('R32','M74',('w1','E'),('t3',3)),
    ('R32','M75',('w1','F'),('r2','C')),('R32','M76',('w1','C'),('r2','F')),
    ('R32','M77',('w1','I'),('t3',5)),('R32','M78',('r2','E'),('r2','I')),
    ('R32','M79',('w1','A'),('t3',0)),('R32','M80',('w1','L'),('t3',7)),
    ('R32','M81',('w1','D'),('t3',2)),('R32','M82',('w1','G'),('t3',4)),
    ('R32','M83',('r2','K'),('r2','L')),('R32','M84',('w1','H'),('r2','J')),
    ('R32','M85',('w1','B'),('t3',1)),('R32','M86',('w1','J'),('r2','H')),
    ('R32','M87',('w1','K'),('t3',6)),('R32','M88',('r2','D'),('r2','G')),
    ('Oitavas','J89',('win',1),('win',4)),('Oitavas','J90',('win',0),('win',2)),
    ('Oitavas','J91',('win',3),('win',5)),('Oitavas','J92',('win',6),('win',7)),
    ('Oitavas','J93',('win',10),('win',11)),('Oitavas','J94',('win',8),('win',9)),
    ('Oitavas','J95',('win',13),('win',15)),('Oitavas','J96',('win',12),('win',14)),
    ('Quartas','J97',('win',16),('win',17)),('Quartas','J98',('win',20),('win',21)),
    ('Quartas','J99',('win',18),('win',19)),('Quartas','J100',('win',22),('win',23)),
    ('Semi','J101',('win',24),('win',25)),('Semi','J102',('win',26),('win',27)),
    ('3o Lugar','J103',('lose',28),('lose',29)),('Final','J104',('win',28),('win',29)),
]
NKO=len(KO)
KO_PTS={'R32':3,'Oitavas':5,'Quartas':8,'Semi':12,'3o Lugar':10,'Final':25}
MEDALS={1:'🥇',2:'🥈',3:'🥉'}
RND_ICO={'R32':'🔵','Oitavas':'🟢','Quartas':'🟡','Semi':'🔴','3o Lugar':'🟠','Final':'⭐'}

# ══════════════════════════════════════════════════════════════════════
# PONTUAÇÃO
# ══════════════════════════════════════════════════════════════════════
def sg(b1,b2,r1,r2):
    try: b1,b2,r1,r2=int(b1),int(b2),int(r1),int(r2)
    except: return 0
    if b1==r1 and b2==r2: return 5
    br=(b1>b2)-(b1<b2); rr=(r1>r2)-(r1<r2)
    if br==rr: return 3 if (b1-b2)==(r1-r2) else 2
    return 0

def calc_st(data):
    st={g:{t:{'pts':0,'played':0,'w':0,'d':0,'l':0,'gf':0,'ga':0} for t in ts}
        for g,ts in GROUPS_DATA.items()}
    for m,(_,g,t1,t2) in enumerate(GROUP_FIXTURES):
        if m not in data or data[m] is None: continue
        s1,s2=data[m]
        if s1 is None or s2 is None: continue
        try: s1,s2=int(s1),int(s2)
        except: continue
        for t,gf,ga in [(t1,s1,s2),(t2,s2,s1)]:
            st[g][t]['played']+=1; st[g][t]['gf']+=gf; st[g][t]['ga']+=ga
        if s1>s2:   st[g][t1]['pts']+=3;st[g][t1]['w']+=1;st[g][t2]['l']+=1
        elif s2>s1: st[g][t2]['pts']+=3;st[g][t2]['w']+=1;st[g][t1]['l']+=1
        else:
            st[g][t1]['pts']+=1;st[g][t1]['d']+=1
            st[g][t2]['pts']+=1;st[g][t2]['d']+=1
    return st

def sort_st(st):
    return {g:sorted(d.items(),key=lambda x:(-x[1]['pts'],-(x[1]['gf']-x[1]['ga']),-x[1]['gf'],FIFA_RANKINGS.get(x[0],150)))
            for g,d in st.items()}

def build_r32(ss,t495):
    g1={g:l[0][0] for g,l in ss.items() if l}
    g2={g:l[1][0] for g,l in ss.items() if len(l)>1}
    g3={g:l[2][0] for g,l in ss.items() if len(l)>2}
    thirds=sorted(
        [(g,-d['pts'],-(d['gf']-d['ga']),-d['gf'],FIFA_RANKINGS.get(g3.get(g,''),150))
         for g,l in ss.items() if len(l)>2 for _,d in [l[2]]],key=lambda x:x[1:])
    top8=[t[0] for t in thirds[:8]]; key=''.join(sorted(top8))
    asgn=t495.get(key,['?']*8)
    def gt3(i):
        a=str(asgn[i]) if i<len(asgn) else '?'
        return g3.get(a[1],'?') if len(a)>=2 else '?'
    def res(sp):
        if sp[0]=='w1': return g1.get(sp[1],'?')
        if sp[0]=='r2': return g2.get(sp[1],'?')
        if sp[0]=='t3': return gt3(sp[1])
        return '?'
    return {m:(res(KO[m][2]),res(KO[m][3])) for m in range(16)},g1,g2,g3,key

def build_bracket(r32,kor):
    tn=dict(r32); w={}
    def wof(m):
        if m not in kor: return None
        t1,t2=tn.get(m,(None,None)); a,b,pen=kor[m]
        if a is None or b is None: return None
        if a>b: return t1
        if b>a: return t2
        return t1 if pen=='S1' else (t2 if pen=='S2' else None)
    def lof(m):
        t1,t2=tn.get(m,(None,None)); wm=wof(m)
        return t2 if wm==t1 else (t1 if wm==t2 else None)
    for m in range(16): w[m]=wof(m)
    for i,(a,b) in enumerate([(1,4),(0,2),(3,5),(6,7),(10,11),(8,9),(13,15),(12,14)]):
        m=16+i; tn[m]=(w.get(a),w.get(b)); w[m]=wof(m)
    for i,(a,b) in enumerate([(16,17),(20,21),(18,19),(22,23)]):
        m=24+i; tn[m]=(w.get(a),w.get(b)); w[m]=wof(m)
    for i,(a,b) in enumerate([(24,25),(26,27)]):
        m=28+i; tn[m]=(w.get(a),w.get(b)); w[m]=wof(m)
    tn[30]=(lof(28),lof(29)); w[30]=wof(30)
    tn[31]=(w.get(28),w.get(29)); w[31]=wof(31)
    return w,tn

def sim_bet(gb,xm,t495):
    r32,*_=build_r32(sort_st(calc_st(gb)),t495)
    picks={}; bn=dict(r32)
    for m in range(16):
        t1,t2=r32.get(m,('?','?')); x=xm.get(m)
        if x=='t1': picks[m]=t1
        elif x=='t2': picks[m]=t2
    def nr(struct,base):
        for i,(p1,p2) in enumerate(struct):
            m=base+i; t1=picks.get(p1,'?'); t2=picks.get(p2,'?')
            bn[m]=(t1,t2); x=xm.get(m)
            if x=='t1': picks[m]=t1
            elif x=='t2': picks[m]=t2
    nr([(1,4),(0,2),(3,5),(6,7),(10,11),(8,9),(13,15),(12,14)],16)
    nr([(16,17),(20,21),(18,19),(22,23)],24)
    nr([(24,25),(26,27)],28)
    for mi in [30,31]:
        def gh(sp):
            ac,src=sp; t1,t2=bn.get(src,('?','?')); wp=picks.get(src,'?')
            return wp if ac=='win' else (t2 if wp==t1 else t1 if wp==t2 else '?')
        t1=gh(KO[mi][2]); t2=gh(KO[mi][3]); bn[mi]=(t1,t2)
        x=xm.get(mi)
        if x=='t1': picks[mi]=t1
        elif x=='t2': picks[mi]=t2
    return picks,bn

def score_all(gb,xm,bb,gr,mmr,br,t495):
    gdet={}; gt=0
    for m,(r1,r2) in gr.items():
        bv=gb.get(m)
        if bv is None: gdet[m]=None; continue
        p=sg(bv[0],bv[1],r1,r2); gdet[m]=p; gt+=p
    ap=mp=0
    if br and bb:
        ar,mr_=br; ab,mb=bb
        if ab and ar and str(ab).strip().lower()==str(ar).strip().lower(): ap=10
        if mb and mr_ and str(mb).strip().lower()==str(mr_).strip().lower(): mp=5
    rw,rn={},{}
    if gr:
        r32r,*_=build_r32(sort_st(calc_st(gr)),t495)
        rw,rn=build_bracket(r32r,mmr)
    bp,_=sim_bet(gb,xm,t495)

    # Vencedores reais agrupados por rodada
    # Lógica v7c: pontua se a seleção escolhida avançou em QUALQUER
    # jogo daquela rodada, independente do confronto específico previsto.
    _rnd_midxs: dict = {}
    for _m, (_rnd, *_rest) in enumerate(KO):
        _rnd_midxs.setdefault(_rnd, []).append(_m)
    _round_winners: dict = {
        _rnd: {rw.get(_m) for _m in _midxs} - {None}
        for _rnd, _midxs in _rnd_midxs.items()
    }

    mdet={}; mt=0
    for m in range(NKO):
        rnd  = KO[m][0]
        pv   = KO_PTS[rnd]
        pick = bp.get(m)
        winners_this_rnd = _round_winners.get(rnd, set())
        if not winners_this_rnd:
            # Nenhum jogo da rodada finalizado → aguardando
            mdet[m] = None
        elif pick and pick != '?' and pick in winners_this_rnd:
            # Seleção escolhida avançou em qualquer jogo da rodada ✓
            mdet[m] = pv; mt += pv
        else:
            # Rodada iniciada mas seleção não avançou (ou não escolheu)
            mdet[m] = 0
    return dict(total=gt+ap+mp+mt,grupos=gt,bonus=ap+mp,mm=mt,
                art_pts=ap,mg_pts=mp,gdet=gdet,mdet=mdet)

# ══════════════════════════════════════════════════════════════════════
# FILE LOADING
# ══════════════════════════════════════════════════════════════════════
@st.cache_data(show_spinner=False)
def load_t495(path):
    try:
        wb=load_workbook(path,data_only=True,read_only=True)
        if 'T495' not in wb.sheetnames: return {}
        ws=wb['T495']; t={}
        for row in ws.iter_rows(min_row=3,values_only=True):
            if row[0] and isinstance(row[0],str) and len(row[0])==8:
                t[row[0]]=[str(v) if v else '?' for v in row[1:9]]
        return t
    except: return {}

@st.cache_data(show_spinner=False)
def load_gab(path):
    try:
        wb=load_workbook(path,data_only=True,read_only=True); sh=wb.sheetnames
        gr={}
        if 'Jogos - Grupos' in sh:
            ws=wb['Jogos - Grupos']
            for m in range(72):
                r1=ws.cell(row=3+m,column=6).value; r2=ws.cell(row=3+m,column=7).value
                if r1 is not None and r2 is not None:
                    try: gr[m]=(int(r1),int(r2))
                    except: pass
        mmr={}
        if 'Mata-Mata - Jogos' in sh:
            ws=wb['Mata-Mata - Jogos']
            for m in range(32):
                g1v=ws.cell(row=4+m,column=6).value; g2v=ws.cell(row=4+m,column=7).value
                pen=ws.cell(row=4+m,column=8).value
                if g1v is not None and g2v is not None:
                    try: mmr[m]=(int(g1v),int(g2v),pen)
                    except: pass
        br=None
        if 'Apostas - Bonus' in sh:
            ws=wb['Apostas - Bonus']
            a=ws.cell(row=4,column=2).value; mg=ws.cell(row=4,column=3).value
            if a or mg: br=(a,mg)
        return gr,mmr,br
    except Exception as e:
        st.error(f"Erro gabarito: {e}"); return {},{},None

@st.cache_data(show_spinner=False)
def load_part(path):
    try:
        wb=load_workbook(path,data_only=True,read_only=True); sh=wb.sheetnames
        gb={}
        if 'Apostas - Grupos' in sh:
            ws=wb['Apostas - Grupos']
            for m in range(72):
                v1=ws.cell(row=5,column=2+m*2).value; v2=ws.cell(row=5,column=3+m*2).value
                if v1 is not None and v2 is not None:
                    try: gb[m]=(int(v1),int(v2))
                    except: pass
        bb=(None,None)
        if 'Apostas - Bonus' in sh:
            ws=wb['Apostas - Bonus']
            bb=(ws.cell(row=5,column=2).value,ws.cell(row=5,column=3).value)
        xm={}
        if 'Apostas - Mata-Mata' in sh:
            ws=wb['Apostas - Mata-Mata']
            for m in range(32):
                v1=ws.cell(row=6,column=2+m*2).value; v2=ws.cell(row=6,column=3+m*2).value
                if v1 and str(v1).strip().upper()=='X': xm[m]='t1'
                elif v2 and str(v2).strip().upper()=='X': xm[m]='t2'
        return gb,bb,xm
    except Exception as e:
        st.warning(f"Erro: {e}"); return {},(None,None),{}

def detect(folder):
    files=sorted(glob.glob(os.path.join(folder,"Bolao_Copa2026_TurimMFO_*.xlsx")))
    gabs,parts=[],[]
    for f in files:
        base=os.path.basename(f)
        nm=re.sub(r'(?i)^Bolao_Copa2026_TurimMFO_','',base).replace('.xlsx','').replace('_',' ').strip()
        if any(k in base.lower() for k in ['master','gabarito','admin','resultado']): gabs.append((nm,f))
        else: parts.append((nm,f))
    return gabs,parts

# ══════════════════════════════════════════════════════════════════════
# SIDEBAR  — nativa, sem CSS override
# ══════════════════════════════════════════════════════════════════════
with st.sidebar:
    # Logos centrados
    _, lc, _ = st.columns([1,3,1])
    with lc:
        if LOGO_TURIM_AZUL: st.image(LOGO_TURIM_AZUL, width=130)
        else: st.markdown("### 🐂 TURIM")
    _, lc2, _ = st.columns([1,5,1])
    with lc2:
        if LOGO_TORI: st.image(LOGO_TORI, width=160)
    st.markdown("---")
    st.markdown("### ⚙️ Configuração")
    folder = st.text_input("📁 Pasta", value=os.getcwd(), label_visibility="collapsed",
                           placeholder="Caminho da pasta com os arquivos...")
    st.caption(f"📁 {folder[:40]}..." if len(folder) > 40 else f"📁 {folder}")

    gabs, parts = detect(folder)

    st.markdown("---")
    if not gabs:
        st.warning("⚠️ Sem gabarito.\nNomeie:\n`Bolao_Copa2026_TurimMFO_Master.xlsx`")
        gab_path = None
    else:
        gsel = st.selectbox("📋 Gabarito", [n for n,_ in gabs])
        gab_path = next(p for n,p in gabs if n==gsel)
        st.success(f"✅ {gsel}")

    st.markdown("---")
    if parts:
        st.markdown(f"**👥 {len(parts)} participante(s):**")
        for nm,_ in parts:
            st.markdown(f"  · {nm}")
    else:
        st.warning("⚠️ Sem participantes.")

    st.markdown("---")
    if st.button("🔄 Recarregar dados", width='stretch'):
        st.cache_data.clear()
        st.rerun()

    # Mascotes no rodapé da sidebar
    if MASCOTES:
        st.markdown("---")
        st.image(MASCOTES, caption="🐂 Rumo ao Hexa! 🏆", width='stretch')

    st.markdown("---")
    st.caption("Bolão Copa 2026 · Turim MFO · v5.0")

# ══════════════════════════════════════════════════════════════════════
# HERO HEADER
# ══════════════════════════════════════════════════════════════════════
st.markdown("""
<div class="hero">
  <div style="display:flex;align-items:center;gap:16px">
    <div style="font-size:2.5rem">⚽</div>
    <div>
      <div class="hero-title">Bolão Copa do Mundo 2026</div>
      <div class="hero-sub">Dashboard Oficial · Turim MFO</div>
      <div class="hero-badge">🇨🇦 Canadá &nbsp;·&nbsp; 🇲🇽 México &nbsp;·&nbsp; 🇺🇸 EUA &nbsp;·&nbsp; Junho–Julho 2026</div>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

if not gab_path or not parts:
    st.info("👈 Configure a pasta na barra lateral com o arquivo Master e os arquivos dos participantes.")
    st.stop()

# ── Carregar dados
with st.spinner("Carregando dados..."):
    t495  = load_t495(gab_path)
    gr,mmr,br = load_gab(gab_path)
    bettors = []
    for nm,fp in parts:
        gb,bb,xm = load_part(fp)
        sc = score_all(gb,xm,bb,gr,mmr,br,t495)
        bettors.append((nm,gb,bb,xm,sc))

bettors.sort(key=lambda x: -x[4]['total'])
maxp = max((b[4]['total'] for b in bettors), default=1) or 1

rw,rn = {},{}
if gr:
    r32r,*_ = build_r32(sort_st(calc_st(gr)), t495)
    rw,rn   = build_bracket(r32r, mmr)

rnd_ms = OrderedDict()
for m,(rnd,*_) in enumerate(KO): rnd_ms.setdefault(rnd,[]).append(m)

# ── Métricas
c5 = st.columns(5)
for col,(val,lbl) in zip(c5,[
    (f"{len(gr)}/72","Jogos Grupos"), (f"{len(mmr)}/32","Jogos MM"),
    (str(sum(a+b for a,b in gr.values())),"Gols"),
    (str(bettors[0][4]['total']) if bettors else "0","Líder (pts)"),
    (str(len(bettors)),"Participantes"),
]):
    col.markdown(
        f'<div class="mc"><div class="mc-v">{val}</div>'
        f'<div class="mc-l">{lbl}</div></div>',
        unsafe_allow_html=True)

st.markdown("")

# ══════════════════════════════════════════════════════════════════════
# TABS
# ══════════════════════════════════════════════════════════════════════
T1,T2,T3,T4 = st.tabs(["🏆 Ranking","⚽ Grupos","🗓️ Mata-Mata","👤 Por Apostador"])

# ── TAB 1: RANKING ───────────────────────────────────────────────────
with T1:
    cA,cB = st.columns([3,2], gap="large")
    with cA:
        st.markdown('<div class="sh">🏆 Classificação Geral</div>', unsafe_allow_html=True)
        for pos,(nm,_,_,_,sc) in enumerate(bettors,1):
            cls  = {1:'rc1',2:'rc2',3:'rc3'}.get(pos,'rcN')
            medal = MEDALS.get(pos, f"{pos}°")
            pct   = int(sc['total']/maxp*100)
            st.markdown(f"""<div class="rc {cls}">
              <div style="font-size:1.35rem;width:32px;text-align:center;flex-shrink:0">{medal}</div>
              <div style="flex:1;min-width:0">
                <div class="rc-name">{nm}</div>
                <div class="rc-sub">Grupos {sc['grupos']} · Bônus {sc['bonus']} · MM {sc['mm']}</div>
                <div class="bar-bg"><div class="bar-fg" style="width:{pct}%"></div></div>
              </div>
              <div style="text-align:right;flex-shrink:0">
                <div class="rc-pts">{sc['total']}</div>
                <div class="rc-pl">pts</div>
              </div>
            </div>""", unsafe_allow_html=True)

    with cB:
        st.markdown('<div class="sh">📊 Pontuação por Fase</div>', unsafe_allow_html=True)
        fig = go.Figure()
        for lbl,vals,color in [
            ('Grupos', [b[4]['grupos'] for b in bettors], '#123A56'),
            ('Bônus',  [b[4]['bonus']  for b in bettors], '#0D8587'),
            ('MM',     [b[4]['mm']     for b in bettors], '#B2584E'),
        ]:
            fig.add_trace(go.Bar(
                name=lbl, y=[b[0] for b in bettors], x=vals, orientation='h',
                marker_color=color, text=vals, textposition='inside',
                insidetextanchor='middle', textfont=dict(size=10, color='#fff')))
        fig.update_layout(
            barmode='stack', height=250, margin=dict(l=0,r=5,t=5,b=5),
            paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
            legend=dict(orientation='h', y=1.06, x=0, font=dict(size=10)),
            xaxis=dict(gridcolor='rgba(0,0,0,.08)', tickfont=dict(size=10)),
            yaxis=dict(tickfont=dict(size=10)))
        st.plotly_chart(fig, width='stretch', config={'displayModeBar':False})

        if bettors:
            nm0,_,_,_,sc0 = bettors[0]
            st.markdown(f"""<div class="mc" style="text-align:left;padding:14px 18px">
              <div style="font-size:.65rem;color:#5C5F62;text-transform:uppercase;letter-spacing:1px">Líder atual</div>
              <div style="font-size:1.35rem;font-weight:800;margin:3px 0">🥇 {nm0}</div>
              <div class="mc-v" style="font-size:1.9rem;text-align:left;color:#D6B864">{sc0['total']} pts</div>
              <div style="font-size:.73rem;color:#5C5F62;margin-top:3px">
                G {sc0['grupos']} · B {sc0['bonus']} · MM {sc0['mm']}</div>
            </div>""", unsafe_allow_html=True)

    # ── EVOLUÇÃO TEMPORAL ────────────────────────────────────────────────
    st.markdown('<div class="sh">📈 Evolução de Pontuação</div>', unsafe_allow_html=True)

    # ── KO dates (R32→Final) ──────────────────────────────────────────────
    KO_DATES = [
        # R32 — 28 Jun–3 Jul
        date(2026,6,28),date(2026,6,28),date(2026,6,28),date(2026,6,28),
        date(2026,6,29),date(2026,6,29),date(2026,6,29),date(2026,6,29),
        date(2026,6,30),date(2026,6,30),date(2026,7,1), date(2026,7,1),
        date(2026,7,2), date(2026,7,2), date(2026,7,3), date(2026,7,3),
        # Oitavas — 4–7 Jul
        date(2026,7,4), date(2026,7,4), date(2026,7,5), date(2026,7,5),
        date(2026,7,6), date(2026,7,6), date(2026,7,7), date(2026,7,7),
        # Quartas — 9–11 Jul
        date(2026,7,9), date(2026,7,10),date(2026,7,11),date(2026,7,11),
        # Semis — 14–15 Jul
        date(2026,7,14),date(2026,7,15),
        # 3o Lugar, Final
        date(2026,7,18),date(2026,7,19),
    ]

    # Phase zones for background highlighting
    PHASE_ZONES = [
        ("Grupos",   date(2026,6,11), date(2026,6,27), "rgba(18,58,86,.12)"),
        ("R32",      date(2026,6,28), date(2026,7,3),  "rgba(13,133,135,.12)"),
        ("Oitavas",  date(2026,7,4),  date(2026,7,7),  "rgba(214,184,100,.10)"),
        ("Quartas",  date(2026,7,9),  date(2026,7,11), "rgba(220,136,74,.12)"),
        ("Semis",    date(2026,7,14), date(2026,7,15), "rgba(178,88,78,.12)"),
        ("3°/Final", date(2026,7,18), date(2026,7,19), "rgba(107,56,143,.12)"),
    ]

    CHART_COLORS = [
        '#D6B864','#0D8587','#B2584E','#2563EB',
        '#DC884A','#7B3FA0','#22C55E','#F97316',
        '#EF4444','#8B5CF6','#06B6D4','#EC4899',
    ]

    # ── Build timeline: {bettor: {date: pts_that_day}} ───────────────────
    # Group matches → date from GROUP_FIXTURES
    # KO matches    → date from KO_DATES
    # Bonus         → injected at Copa kick-off date (Jun 11)
    COPA_START = date(2026, 6, 11)

    all_names = [b[0] for b in bettors]

    # Participant filter
    # Align button height with multiselect using a hidden label trick
    fc1, fc2 = st.columns([1, 5])
    with fc1:
        # Hidden label so button aligns with multiselect (same label height)
        st.markdown(
            "<label style='font-size:.875rem;opacity:0;display:block'>_</label>",
            unsafe_allow_html=True,
        )
        if st.button("👥 Todos", use_container_width=True):
            st.session_state["sel_bettors"] = all_names

    with fc2:
        if "sel_bettors" not in st.session_state:
            st.session_state["sel_bettors"] = all_names
        valid_sel = [n for n in st.session_state["sel_bettors"] if n in all_names]
        if not valid_sel:
            valid_sel = all_names
        chosen = st.multiselect(
            "Participantes:",
            options=all_names,
            default=valid_sel,
            key="ms_bettors",
            placeholder="Digite ou selecione participantes...",
        )
        st.session_state["sel_bettors"] = chosen if chosen else all_names

    active_bettors = [b for b in bettors if b[0] in st.session_state["sel_bettors"]]
    if not active_bettors:
        active_bettors = bettors

    # Chart mode
    chart_mode = st.radio(
        "Modo:",
        ["📈 Evolução Acumulada", "📊 Pontuação por Fase"],
        horizontal=True,
        label_visibility="collapsed",
    )

    if chart_mode == "📈 Evolução Acumulada":

        # Build sorted list of all match dates (group + KO)
        all_match_dates = sorted(set(
            [gf[0] for gf in GROUP_FIXTURES] + KO_DATES
        ))

        fig_evo = go.Figure()

        # Phase zone fills with top-aligned label annotations
        ZONE_FONT_COLOR = "rgba(200,220,240,.70)"
        for ph_name, d0, d1, color in PHASE_ZONES:
            fig_evo.add_vrect(
                x0=str(d0), x1=str(d1),
                fillcolor=color, layer="below",
                line=dict(width=1, color=color, dash="dot"),
            )
            # Place label at mid-x of zone using add_annotation
            # We'll use xref="x" with the first date of each zone as anchor
            fig_evo.add_annotation(
                x=str(d0), xref="x", y=1.0, yref="paper",
                text=f"<b>{ph_name}</b>",
                showarrow=False, xanchor="left", yanchor="top",
                font=dict(size=9, color=ZONE_FONT_COLOR),
                bgcolor="rgba(0,0,0,0)", borderpad=3,
            )

        for idx, (nm, gb, _, xm, sc) in enumerate(active_bettors):
            # Points per date
            pts_by_date: dict = {COPA_START: sc["bonus"]}  # Bônus on kick-off
            for m, (d_, g, t1, t2) in enumerate(GROUP_FIXTURES):
                p = sc["gdet"].get(m)
                if p is not None and p > 0:
                    pts_by_date[d_] = pts_by_date.get(d_, 0) + p
            for m, ko_date in enumerate(KO_DATES):
                p = sc["mdet"].get(m)
                if p is not None and p > 0:
                    pts_by_date[ko_date] = pts_by_date.get(ko_date, 0) + p

            # Build cumulative series over all match dates
            dates_x, cumul_y = [], []
            running = 0
            for d in all_match_dates:
                running += pts_by_date.get(d, 0)
                dates_x.append(str(d))
                cumul_y.append(running)

            color = CHART_COLORS[idx % len(CHART_COLORS)]
            first = nm.split()[0] if " " in nm else nm
            fig_evo.add_trace(go.Scatter(
                x=dates_x, y=cumul_y,
                mode="lines+markers",
                name=first,
                line=dict(color=color, width=2.5),
                marker=dict(size=6, color=color,
                            line=dict(color="white", width=1.5)),
                hovertemplate=f"<b>{nm}</b><br>%{{x}}: <b>%{{y}} pts</b><extra></extra>",
            ))

        fig_evo.update_layout(
            height=400, margin=dict(l=0, r=10, t=36, b=10),
            paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
            legend=dict(orientation="h", y=1.08, x=0,
                        font=dict(size=11), bgcolor="rgba(0,0,0,0)"),
            xaxis=dict(
                type="category",
                gridcolor="rgba(128,128,128,.10)",
                tickfont=dict(size=10),
                tickangle=-35,
                showline=True, linecolor="rgba(128,128,128,.2)",
                title=dict(text="Data do jogo", font=dict(size=11)),
            ),
            yaxis=dict(
                gridcolor="rgba(128,128,128,.12)",
                tickfont=dict(size=11),
                title=dict(text="Pts acumulados", font=dict(size=11)),
                rangemode="tozero",
            ),
            hovermode="x unified",
        )
        # Phase legend strip
        legend_html = (
            '<div style="display:flex;flex-wrap:wrap;align-items:center;'
            'gap:0;margin:-4px 0 6px;padding:6px 10px;'
            'border-radius:8px;background:rgba(128,128,128,.06)">' +
            f'<span style="display:inline-flex;align-items:center;gap:5px;margin-right:14px;font-size:.78rem;font-weight:600;white-space:nowrap"><span style="display:inline-block;width:13px;height:13px;border-radius:3px;background:#123A56;opacity:.75;flex-shrink:0"></span>Grupos</span>'f'<span style="display:inline-flex;align-items:center;gap:5px;margin-right:14px;font-size:.78rem;font-weight:600;white-space:nowrap"><span style="display:inline-block;width:13px;height:13px;border-radius:3px;background:#0D8587;opacity:.75;flex-shrink:0"></span>R32</span>'f'<span style="display:inline-flex;align-items:center;gap:5px;margin-right:14px;font-size:.78rem;font-weight:600;white-space:nowrap"><span style="display:inline-block;width:13px;height:13px;border-radius:3px;background:#D6B864;opacity:.75;flex-shrink:0"></span>Oitavas</span>'f'<span style="display:inline-flex;align-items:center;gap:5px;margin-right:14px;font-size:.78rem;font-weight:600;white-space:nowrap"><span style="display:inline-block;width:13px;height:13px;border-radius:3px;background:#DC884A;opacity:.75;flex-shrink:0"></span>Quartas</span>'f'<span style="display:inline-flex;align-items:center;gap:5px;margin-right:14px;font-size:.78rem;font-weight:600;white-space:nowrap"><span style="display:inline-block;width:13px;height:13px;border-radius:3px;background:#B2584E;opacity:.75;flex-shrink:0"></span>Semis</span>'f'<span style="display:inline-flex;align-items:center;gap:5px;margin-right:14px;font-size:.78rem;font-weight:600;white-space:nowrap"><span style="display:inline-block;width:13px;height:13px;border-radius:3px;background:#7B3FA0;opacity:.75;flex-shrink:0"></span>3°/Final</span>' +
            '</div>'
        )
        st.markdown(legend_html, unsafe_allow_html=True)
        st.plotly_chart(fig_evo, use_container_width=True, config={"displayModeBar": False})

    else:  # Pontuação por fase
        PHASES_ORDER = ["Grupos","Bônus","R32","Oitavas","Quartas","Semi","3o Lugar","Final"]

        def phase_pts(sc, phase):
            if phase == "Grupos": return sc["grupos"]
            if phase == "Bônus":  return sc["bonus"]
            midxs = [m for m,(rnd,*_) in enumerate(KO) if rnd==phase]
            return sum(sc["mdet"].get(m,0) or 0 for m in midxs)

        sel_phase = st.select_slider(
            "Fase:", options=PHASES_ORDER, value="Grupos",
            label_visibility="collapsed",
        )
        phase_data = sorted(
            [(nm, phase_pts(sc, sel_phase)) for nm,_,_,_,sc in active_bettors],
            key=lambda x: -x[1],
        )
        bar_names  = [p[0].split()[0] if " " in p[0] else p[0] for p in phase_data]
        bar_vals   = [p[1] for p in phase_data]
        bar_colors = [CHART_COLORS[i % len(CHART_COLORS)] for i in range(len(phase_data))]

        fig_bar = go.Figure(go.Bar(
            x=bar_names, y=bar_vals, text=bar_vals,
            textposition="outside",
            marker_color=bar_colors,
            hovertext=[p[0] for p in phase_data],
            hovertemplate="<b>%{hovertext}</b><br>%{y} pts<extra></extra>",
        ))
        fig_bar.update_layout(
            height=340, margin=dict(l=0, r=10, t=40, b=10),
            paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
            title=dict(text=f"Pontuação — fase: <b>{sel_phase}</b>",
                       font=dict(size=13), x=0),
            xaxis=dict(gridcolor="rgba(128,128,128,.10)", tickfont=dict(size=11)),
            yaxis=dict(gridcolor="rgba(128,128,128,.12)", tickfont=dict(size=11),
                       rangemode="tozero"),
        )
        st.plotly_chart(fig_bar, use_container_width=True, config={"displayModeBar": False})


# ── TAB 2: GRUPOS ────────────────────────────────────────────────────
with T2:
    st.markdown('<div class="sh">⚽ Classificação nos Grupos</div>', unsafe_allow_html=True)
    mode = st.radio("Baseado em:", ["📋 Resultados Reais", "👤 Apostas de Participante"],
                    horizontal=True, label_visibility="collapsed")
    if mode.startswith("📋"):
        data_src = sort_st(calc_st(gr)) if gr else {
            g:[(t,{'pts':0,'played':0,'w':0,'d':0,'l':0,'gf':0,'ga':0}) for t in ts]
            for g,ts in GROUPS_DATA.items()}
        if not gr: st.info("Nenhum resultado preenchido no gabarito ainda.")
    else:
        sel_b = st.selectbox("Apostador", [b[0] for b in bettors], key='grp_pick')
        bd    = next(b for b in bettors if b[0]==sel_b)
        data_src = sort_st(calc_st(bd[1]))

    DOT   = {1:'#22C55E',2:'#86EFAC',3:'#FB923C',4:'#F87171'}
    RCLS  = {1:'row-q1',2:'row-q2',3:'row-q3',4:'row-q4'}
    for row_g in [GL[i:i+3] for i in range(0,12,3)]:
        cs = st.columns(3, gap="small")
        for c,grp in zip(cs,row_g):
            with c:
                color = GRP_COLORS.get(grp,'#123A56')
                rows  = ""
                for pos,(team,d) in enumerate(data_src.get(grp,[]),1):
                    gd  = d['gf']-d['ga']; pl = d['played']
                    pct = int(d['pts']/(pl*3)*100) if pl else 0
                    gdc = '#22C55E' if gd>0 else '#EF4444' if gd<0 else 'inherit'
                    gds = ('+' if gd>0 else '')+str(gd)
                    rows += f"""<tr>
                      <td><span class="dot" style="background:{DOT.get(pos,'#888')}"></span>{pos}</td>
                      <td class="nm">{F(team)} {team}</td>
                      <td><b>{d['pts']}</b></td><td>{pl}</td>
                      <td>{d['w']}</td><td>{d['d']}</td><td>{d['l']}</td>
                      <td>{d['gf']}</td><td>{d['ga']}</td>
                      <td style="color:{gdc}">{gds}</td><td>{pct}%</td>
                    </tr>"""
                st.markdown(f"""<div class="gb">
                  <div class="gb-hdr" style="color:{color}">Grupo {grp}</div>
                  <table class="gt">
                    <tr><th>#</th><th style="text-align:left">Seleção</th>
                        <th>P</th><th>J</th><th>V</th><th>E</th><th>D</th>
                        <th>GP</th><th>GC</th><th>SG</th><th>%</th></tr>
                    {rows}
                  </table>
                </div>""", unsafe_allow_html=True)

# ── TAB 3: MATA-MATA ─────────────────────────────────────────────────
with T3:
    st.markdown('<div class="sh">🗓️ Mata-Mata — Seleções por Fase</div>', unsafe_allow_html=True)

    # ── Pre-compute ─────────────────────────────────────────────────────
    # Real winners per round
    real_by_rnd: dict = {
        rnd: {rw.get(m) for m in midxs} - {None}
        for rnd, midxs in rnd_ms.items()
    }

    # Each bettor's picks per round  {rnd: set_of_teams}
    def picks_by_round(gb, xm, t495_):
        bp, _ = sim_bet(gb, xm, t495_)
        out: dict = {}
        for m, (rnd, *_) in enumerate(KO):
            p = bp.get(m)
            if p and p != "?":
                out.setdefault(rnd, set()).add(p)
        return out

    btr = [(nm, picks_by_round(gb, xm, t495), sc)
           for nm, gb, _, xm, sc in bettors]

    # ── Round-by-round sections ─────────────────────────────────────────
    for rnd, mlist in rnd_ms.items():
        pv       = KO_PTS[rnd]
        real_set = real_by_rnd.get(rnd, set())
        started  = bool(real_set)
        ico      = RND_ICO.get(rnd, "⚪")
        max_pts  = pv * len(mlist)

        with st.expander(f"{ico} {rnd} — {pv} pts por seleção que avançar", expanded=(rnd=="R32")):

            # ── Row 1: Who really advanced ───────────────────────────
            if started:
                real_pills = "".join(
                    f'<span class="pill pill-real">{F(t)} {t}</span>'
                    for t in sorted(real_set, key=str)
                )
                st.markdown(
                    f'<div style="margin-bottom:12px">'
                    f'<div class="rnd-section-lbl">✅ Avançaram ({len(real_set)})</div>'
                    f'<div>{real_pills}</div></div>',
                    unsafe_allow_html=True,
                )
            else:
                st.markdown(
                    '<div style="opacity:.5;font-size:.82rem;margin-bottom:12px">⏳ Rodada ainda não iniciou.</div>',
                    unsafe_allow_html=True,
                )

            st.divider()

            # ── Row 2: Per-bettor comparison ─────────────────────────
            # Collect all teams that appear (picks + real)
            all_teams = set(real_set)
            for _, prnd, _ in btr:
                all_teams |= prnd.get(rnd, set())
            all_teams -= {"?", None}

            # Sort: real winners first (alphabetically), then misses
            sorted_teams = (
                sorted(all_teams & real_set, key=str) +
                sorted(all_teams - real_set,  key=str)
            )

            # Header row: apostadores
            hcols = st.columns([3] + [1]*len(btr))
            hcols[0].markdown(
                '<div style="font-size:.72rem;font-weight:700;opacity:.6;padding:2px 0">SELEÇÃO</div>',
                unsafe_allow_html=True,
            )
            for ci, (nm, _, _sc) in enumerate(btr, 1):
                first = nm.split()[0]
                hcols[ci].markdown(
                    f'<div style="text-align:center;font-size:.72rem;font-weight:700;opacity:.6">{first}</div>',
                    unsafe_allow_html=True,
                )

            # Data rows: one per team
            for team in sorted_teams:
                in_real = team in real_set
                dcols   = st.columns([3] + [1]*len(btr))

                # Team cell
                if in_real:
                    team_html = f'<span style="font-weight:700;color:#0D8587;font-size:.82rem">{F(team)} {team}</span>'
                else:
                    team_html = f'<span style="font-size:.82rem;opacity:.55">{F(team)} {team}</span>'
                dcols[0].markdown(team_html, unsafe_allow_html=True)

                # Apostador cells
                for ci, (nm, prnd, sc) in enumerate(btr, 1):
                    picked = team in prnd.get(rnd, set())
                    if picked and in_real and started:
                        icon_c = "✅"
                    elif picked and not started:
                        icon_c = "⏳"
                    elif picked:          # picked but didn't advance
                        icon_c = "❌"
                    elif in_real and started:  # advanced but not picked
                        icon_c = "🔵"
                    else:
                        icon_c = "·"
                    dcols[ci].markdown(
                        f'<div style="text-align:center;font-size:.88rem">{icon_c}</div>',
                        unsafe_allow_html=True,
                    )

            st.divider()

            # ── Row 3: Score strip per bettor ─────────────────────────
            strip_cols = st.columns([3] + [1]*len(btr))
            strip_cols[0].markdown(
                '<div style="font-size:.72rem;font-weight:700;opacity:.6">ACERTOS / PTS</div>',
                unsafe_allow_html=True,
            )
            for ci, (nm, prnd, sc) in enumerate(btr, 1):
                my = prnd.get(rnd, set())
                ok = len(my & real_set) if started else "—"
                pts_rnd = sum(sc["mdet"].get(m, 0) or 0 for m in mlist)
                strip_cols[ci].markdown(
                    f'<div style="text-align:center">'
                    f'<div style="font-size:.9rem;font-weight:800;color:#D6B864">{pts_rnd}</div>'
                    f'<div style="font-size:.65rem;opacity:.55">{ok}/{len(my)}</div>'
                    f'</div>',
                    unsafe_allow_html=True,
                )

    # ── Overall MM ranking chart ──────────────────────────────────────────
    st.markdown('<div class="sh">📊 Ranking Mata-Mata</div>', unsafe_allow_html=True)
    mmdf = pd.DataFrame([(b[0], b[4]["mm"]) for b in bettors], columns=["Apostador","Pts"])
    fig2 = px.bar(mmdf, x="Apostador", y="Pts", text="Pts",
                  color="Pts", color_continuous_scale=["#B2584E","#DC884A","#D6B864"])
    fig2.update_layout(
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=0,r=0,t=5,b=5), coloraxis_showscale=False, height=230,
        xaxis=dict(gridcolor="rgba(128,128,128,.12)"),
        yaxis=dict(gridcolor="rgba(128,128,128,.12)"),
    )
    fig2.update_traces(textposition="outside")
    st.plotly_chart(fig2, use_container_width=True, config={"displayModeBar":False})

# ── TAB 4: POR APOSTADOR ─────────────────────────────────────────────
with T4:
    sel = st.selectbox("👤 Apostador", [b[0] for b in bettors])
    bnm,bgb,bbb,bxm,bsc = next(b for b in bettors if b[0]==sel)
    pos   = next(i for i,(nm,*_) in enumerate(bettors,1) if nm==bnm)
    medal = MEDALS.get(pos, f"{pos}°")

    st.markdown(f"""
    <div class="hero" style="padding:18px 26px;margin-bottom:12px">
      <span style="font-size:1.5rem">{medal}</span>
      <span class="hero-title" style="font-size:1.45rem;margin-left:10px">{bnm}</span>
      <div class="hero-sub" style="margin-top:6px">
        Total: <b>{bsc['total']} pts</b> &nbsp;·&nbsp;
        Grupos: <b>{bsc['grupos']}</b> &nbsp;·&nbsp;
        Bônus: <b>{bsc['bonus']}</b> &nbsp;·&nbsp;
        Mata-Mata: <b>{bsc['mm']}</b>
      </div>
    </div>""", unsafe_allow_html=True)

    # Bônus
    st.markdown('<div class="sh">🎁 Bônus Pré-Copa</div>', unsafe_allow_html=True)
    ab,mb = bbb; ar,mr_ = (br or (None,None))
    for col,(lbl,bv,rv,pts) in zip(st.columns(2),[
        ("⚽ Artilheiro", ab, ar, bsc['art_pts']),
        ("🏅 Melhor Campanha", mb, mr_, bsc['mg_pts']),
    ]):
        ok  = bv and rv and str(bv).strip().lower()==str(rv).strip().lower()
        ico = "✅" if ok else ("❌" if rv else "⏳")
        bvs = f"{F(bv)} {bv}" if bv else "—"
        rvs = f"{F(rv)} {rv}" if rv else "Aguardando"
        col.markdown(f"""<div class="bc">
          <div class="bc-lbl">{lbl}</div>
          <div class="bc-ico">{ico}</div>
          <div class="bc-bet">Apostou: {bvs}</div>
          <div class="bc-real">Real: {rvs}</div>
          <div class="bc-pts">{pts} pts</div>
        </div>""", unsafe_allow_html=True)

    # Grupos
    st.markdown('<div class="sh">⚽ Fase de Grupos — Jogo a Jogo</div>', unsafe_allow_html=True)
    for row_g in [GL[i:i+3] for i in range(0,12,3)]:
        gcols = st.columns(3)
        for gcol,grp in zip(gcols,row_g):
            with gcol:
                color = GRP_COLORS.get(grp,'#123A56')
                st.markdown(
                    f'<div style="font-weight:800;color:{color};font-size:.87rem;margin-bottom:5px">Grupo {grp}</div>',
                    unsafe_allow_html=True)
                html = ""
                for m,(_,g,t1,t2) in enumerate(GROUP_FIXTURES):
                    if g!=grp: continue
                    pts  = bsc['gdet'].get(m); bet = bgb.get(m); real = gr.get(m)
                    bs   = f"{bet[0]}–{bet[1]}" if bet else "—"
                    rs   = f"{real[0]}–{real[1]}" if real else "⏳"
                    bdg  = ('<span class="b5">5</span>' if pts==5 else
                            '<span class="b3">3</span>' if pts==3 else
                            '<span class="b2">2</span>' if pts==2 else
                            '<span class="b0">0</span>' if pts==0 else
                            '<span class="bN">–</span>')
                    html += f"""<div class="mr">
                      <div class="mr-t">{F(t1)} {t1}<br>{F(t2)} {t2}</div>
                      <span class="mr-s">{bs} → {rs}</span>
                      {bdg}
                    </div>"""
                st.markdown(html, unsafe_allow_html=True)

    # Mata-mata
    # Mata-mata por apostador
    # ── Mata-Mata: seleções por fase ──────────────────────────────────────
    st.markdown('<div class="sh">🏆 Mata-Mata — Seleções por Fase</div>', unsafe_allow_html=True)

    bpicks_b, bnames_b = sim_bet(bgb, bxm, t495)
    b_prnd: dict = {}
    for m, (rnd, *_) in enumerate(KO):
        p = bpicks_b.get(m)
        if p and p != "?":
            b_prnd.setdefault(rnd, set()).add(p)

    # Reuse real_by_rnd from T3 scope (same module)

    for rnd, mlist in rnd_ms.items():
        pv         = KO_PTS[rnd]
        real_set   = real_by_rnd.get(rnd, set())
        my_picks   = b_prnd.get(rnd, set())
        started    = bool(real_set)
        pts_earned = sum(bsc["mdet"].get(m,0) or 0 for m in mlist)
        correct    = len(my_picks & real_set) if started else 0
        missed     = my_picks - real_set         # picked but didn't advance
        not_picked = real_set - my_picks         # advanced but not picked
        ico        = RND_ICO.get(rnd,"⚪")

        pts_label = (f"+{pts_earned} pts" if pts_earned > 0 else
                     ("⏳ aguardando" if not started else "0 pts"))

        with st.expander(
            f"{ico} {rnd}  ({pv} pts/seleção) — {pts_label}",
            expanded=(rnd in ["R32","Oitavas"]),
        ):
            col_p, col_r = st.columns(2, gap="large")

            with col_p:
                st.markdown(
                    f'<div class="rnd-section-lbl">🎯 Apostou que avançariam ({len(my_picks)})</div>',
                    unsafe_allow_html=True,
                )
                if my_picks:
                    html_p = ""
                    for team in sorted(my_picks, key=str):
                        in_real = team in real_set
                        if not started:
                            cls, ico_t = "pill-wait", "⏳"
                        elif in_real:
                            cls, ico_t = "pill-hit",  "✅"
                        else:
                            cls, ico_t = "pill-miss", "❌"
                        html_p += f'<span class="pill {cls}">{ico_t} {F(team)} {team}</span>'
                    st.markdown(f'<div style="line-height:2">{html_p}</div>', unsafe_allow_html=True)
                else:
                    st.markdown('<div style="opacity:.45;font-size:.82rem">Nenhuma seleção escolhida.</div>',
                                unsafe_allow_html=True)

            with col_r:
                st.markdown(
                    f'<div class="rnd-section-lbl">✅ Avançaram de verdade ({len(real_set)})</div>',
                    unsafe_allow_html=True,
                )
                if real_set:
                    html_r = ""
                    for team in sorted(real_set, key=str):
                        picked_it = team in my_picks
                        cls_r  = "pill-hit"  if picked_it else "pill-real"
                        ico_r  = "🎯" if picked_it else ""
                        html_r += f'<span class="pill {cls_r}">{ico_r} {F(team)} {team}</span>'
                    st.markdown(f'<div style="line-height:2">{html_r}</div>', unsafe_allow_html=True)
                else:
                    st.markdown('<div style="opacity:.45;font-size:.82rem">Aguardando resultados...</div>',
                                unsafe_allow_html=True)

            # Score strip
            if started:
                miss_str = (", ".join(f"{F(t)}{t}" for t in sorted(missed, key=str))
                            if missed else "—")
                unexp_str = (", ".join(f"{F(t)}{t}" for t in sorted(not_picked, key=str))
                             if not_picked else "—")
                st.markdown(
                    f'<div class="score-strip">'
                    f'<span>✅ <b style="color:#22C55E">{correct}</b> acertos</span>'
                    f'<span>❌ <b style="color:#EF4444">{len(missed)}</b> erros</span>'
                    f'<span style="opacity:.6">Não apostou em: {unexp_str}</span>'
                    f'<span style="margin-left:auto;font-weight:800;color:#D6B864">{pts_earned} pts</span>'
                    f'</div>',
                    unsafe_allow_html=True,
                )


# ── Footer
st.markdown("---")
ft_cols = st.columns([1,3,1])
with ft_cols[0]:
    if LOGO_TORI:
        st.image(LOGO_TORI, width=100)
with ft_cols[1]:
    st.markdown(
        "<div style='text-align:center;color:#5C5F62;font-size:.8rem;padding:.5rem 0'>"
        "Bolão Copa do Mundo 2026 · Turim MFO &nbsp;·&nbsp; "
        "Desenvolvido com ❤️ &nbsp;·&nbsp; 🐂 Rumo ao Hexa!</div>",
        unsafe_allow_html=True)
with ft_cols[2]:
    if LOGO_TURIM_AZUL:
        st.image(LOGO_TURIM_AZUL, width=100)