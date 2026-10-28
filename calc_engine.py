# calc_engine.py
# Motor de cálculo y gráficos — sin dependencias de Streamlit ni Colab

import pandas as pd
import numpy as np
import os
import io
import itertools
import warnings
from datetime import date
import statsmodels.formula.api as smf
from scipy import stats
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D  # noqa
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

warnings.filterwarnings('default')

VIA_1, VIA_2, VIA_3 = 'via1', 'via2', 'via3'
VIA_INFO = {
    VIA_1: {'nombre': '2^k con replicas puras', 'emoji': 'A'},
    VIA_2: {'nombre': '2^k con bloques + centro', 'emoji': 'B'},
    VIA_3: {'nombre': '2^k no replicado (Lenth)', 'emoji': 'C'},
}

# ---------- Lenth ----------
def calcular_lenth(effects, df_resid=None):
    vals = np.array([e['efecto'] for e in effects.values()])
    abs_eff = np.abs(vals); m = len(abs_eff)
    if m == 0:
        return {'m': 0, 'PSE': 0, 'ME': 0, 'SME': 0, 'df_t': m}
    s0 = 1.5 * np.median(abs_eff)
    sub = abs_eff[abs_eff < 2.5 * s0]
    s1 = 1.5 * np.median(sub) if len(sub) > 0 else s0
    PSE = max(s1, 1e-9)
    df_t = df_resid if (df_resid is not None and df_resid >= 1) else m
    ME = stats.t.ppf(0.975, df_t) * PSE
    gamma = 0.5 * (1 + 0.95**(1.0/m))
    SME = stats.t.ppf(gamma, df_t) * PSE
    return {'m': m, 'PSE': PSE, 'ME': ME, 'SME': SME, 'df_t': df_t}

# ---------- Preparacion ----------
def preparar_datos(raw, factores, respuestas):
    tmp = raw.copy()
    cg = '__G__'; cy = '__Y__'
    while cg in tmp.columns: cg += '_'
    while cy in tmp.columns: cy += '_'
    df = pd.melt(tmp, id_vars=factores, value_vars=respuestas,
                 var_name=cg, value_name=cy).rename(columns={cg:'Grupo', cy:'Y'})
    df['Curvatura'] = df[factores].apply(
        lambda r: 1 if all(v==0 for v in r) else 0, axis=1)
    for f in factores: df[f] = pd.to_numeric(df[f], errors='raise')
    return df

def diagnosticar(raw, factores, respuestas):
    k = len(factores); n = len(raw); n_r = len(respuestas)
    if k < 2: return None, "Se requieren 2 o mas factores."
    for f in factores:
        vals = set(pd.to_numeric(raw[f], errors='coerce').dropna().unique().tolist())
        if not vals.issubset({-1,0,1}) or not {-1,1}.issubset(vals):
            return None, f"Factor {f} mal codificado: {vals}."
    centros = int(raw[factores].eq(0).all(axis=1).sum())
    esq = raw.loc[~raw[factores].eq(0).all(axis=1), factores]
    esp = set(itertools.product([-1,1], repeat=k))
    obs = [tuple(int(v) for v in row) for row in esq.to_numpy()]
    falt = esp - set(obs)
    if falt: return None, f"Faltan {len(falt)} esquinas."
    if centros > 0 and n == 2**k + 1 and centros == 1 and n_r >= 2:
        return VIA_2, f"{2**k} esquinas + 1 centro / {n_r} bloques."
    if centros == 0 and n == 2**k and n_r >= 2:
        return VIA_1, f"{2**k} esquinas / {n_r} replicas."
    if centros == 0 and n == 2**k and n_r == 1:
        return VIA_3, f"{2**k} esquinas / sin replicas."
    if centros > 0 and n_r >= 2:
        return VIA_2, f"{n} filas / {centros} centro(s) / {n_r} columnas."
    return None, f"No clasificable (n={n}, k={k}, centros={centros}, resp={n_r})."

# ---------- Calculos base ----------
def _calc_effects(df, factores, Nf):
    k = len(factores); esq = df[df['Curvatura']==0].copy()
    out = {}
    for mask in range(1, 2**k):
        inv = [factores[i] for i in range(k) if mask & (1<<i)]
        s = esq[inv].prod(axis=1)
        c = (esq['Y']*s).sum()
        out['*'.join(inv)] = {'contraste': c, 'SS': c**2/Nf,
                              'efecto': c/(Nf/2), 'coef': c/Nf, 'inv': inv}
    return out

def _calc_main_effects(df, factores):
    esq = df[df['Curvatura']==0]
    return {f: {'low': esq[esq[f]==-1]['Y'].mean(),
                'high': esq[esq[f]==1]['Y'].mean(),
                'eff': esq[esq[f]==1]['Y'].mean() - esq[esq[f]==-1]['Y'].mean()}
            for f in factores}

def _calc_interactions(df, factores):
    esq = df[df['Curvatura']==0]; k = len(factores); out = {}
    for i in range(k):
        for j in range(i+1, k):
            f1, f2 = factores[i], factores[j]
            out[(f1,f2)] = {
                'mll': esq[(esq[f1]==-1)&(esq[f2]==-1)]['Y'].mean(),
                'mhl': esq[(esq[f1]== 1)&(esq[f2]==-1)]['Y'].mean(),
                'mlh': esq[(esq[f1]==-1)&(esq[f2]== 1)]['Y'].mean(),
                'mhh': esq[(esq[f1]== 1)&(esq[f2]== 1)]['Y'].mean()}
    return out

# ---------- Via 1 con soporte para terminos ----------
def calcular_via1_con_terminos(df, factores, terms=None):
    k = len(factores); n_r = df['Grupo'].nunique(); N = len(df); Nf = n_r*2**k
    G = df['Y'].sum(); y = G/N
    SS_t = (df['Y']**2).sum() - G**2/N
    ef_all = _calc_effects(df, factores, Nf)
    if terms is None:
        terms = list(ef_all.keys())
    ef = {t: ef_all[t] for t in terms if t in ef_all}
    if len(ef) == 0:
        raise ValueError("Debes incluir al menos un termino en el modelo.")
    SS_m = sum(e['SS'] for e in ef.values())
    SS_r = SS_t - SS_m
    df_r = N - 1 - len(ef)
    if df_r < 1: raise ValueError(f"df_resid={df_r}<1. Demasiados terminos.")
    MS_r = SS_r/df_r
    for e in ef.values():
        e['MS']=e['SS']; e['F']=e['MS']/MS_r; e['p']=stats.f.sf(e['F'],1,df_r)
    rows = [(n,e['SS'],1,e['MS'],e['F'],e['p']) for n,e in ef.items()]
    rows.append(('Residual', SS_r, df_r, MS_r, np.nan, np.nan))
    anova = pd.DataFrame(rows, columns=['Fuente','SS','GL','MS','F','p'])
    anova['F_crit'] = [stats.f.ppf(0.95,gl,df_r) if pd.notna(gl) and gl>0 else np.nan
                       for gl in anova['GL']]
    anova['Sig'] = ['OK' if (pd.notna(p) and p<0.05) else ('-' if pd.isna(p) else 'NO')
                    for p in anova['p']]
    R2 = 1 - SS_r/SS_t; R2a = 1 - (SS_r/df_r)/(SS_t/(N-1))
    R = dict(via=VIA_1, k=k, factores=factores, N=N, Nf=Nf, n_replicas=n_r,
             n_bloques=0, tiene_bloques=False, tiene_centro=False,
             y_bar=y, SS_total=SS_t, SS_resid=SS_r, SS_model=SS_m,
             effects=ef, all_effects=ef_all, terms_included=list(terms),
             anova=anova, MS_resid=MS_r, df_resid=df_r,
             R2=R2, R2_adj=R2a,
             main_effects=_calc_main_effects(df, factores),
             interactions=_calc_interactions(df, factores),
             t_crit=stats.t.ppf(0.975, df_r))
    R['lenth'] = calcular_lenth(ef, df_resid=df_r)
    return R

# ---------- Via 2 con soporte para terminos ----------
def calcular_via2_con_terminos(df, factores, terms=None):
    k = len(factores); n_b = df['Grupo'].nunique(); N = len(df)
    Nf = n_b*2**k; Nc = N - Nf; G = df['Y'].sum(); y = G/N
    esq = df[df['Curvatura']==0]; cen = df[df['Curvatura']==1]
    SS_t = (df['Y']**2).sum() - G**2/N
    bt = df.groupby('Grupo')['Y'].sum(); sz = df.groupby('Grupo').size()
    SS_b = (bt**2/sz).sum() - G**2/N
    df_b = n_b - 1
    ef_all = _calc_effects(df, factores, Nf)
    if terms is None:
        terms = list(ef_all.keys())
    ef = {t: ef_all[t] for t in terms if t in ef_all}
    yf = esq['Y'].mean(); yc = cen['Y'].mean()
    SS_c = Nf*Nc/(Nf+Nc)*(yf-yc)**2
    SS_m = SS_b + sum(e['SS'] for e in ef.values()) + SS_c
    SS_r = SS_t - SS_m
    df_r = N - 1 - (df_b + len(ef) + 1)
    if df_r < 1: raise ValueError(f"df_resid={df_r}<1")
    MS_r = SS_r/df_r; MS_b = SS_b/df_b; F_b = MS_b/MS_r
    p_b = stats.f.sf(F_b, df_b, df_r)
    MS_c = SS_c; F_c = MS_c/MS_r; p_c = stats.f.sf(F_c, 1, df_r)
    for e in ef.values():
        e['MS']=e['SS']; e['F']=e['MS']/MS_r; e['p']=stats.f.sf(e['F'],1,df_r)
    rows = [('Bloques', SS_b, df_b, MS_b, F_b, p_b)]
    for n_, e in ef.items(): rows.append((n_, e['SS'], 1, e['MS'], e['F'], e['p']))
    rows.append(('Curvatura', SS_c, 1, MS_c, F_c, p_c))
    rows.append(('Residual', SS_r, df_r, MS_r, np.nan, np.nan))
    anova = pd.DataFrame(rows, columns=['Fuente','SS','GL','MS','F','p'])
    anova['F_crit'] = [stats.f.ppf(0.95,gl,df_r) if pd.notna(gl) and gl>0 else np.nan
                       for gl in anova['GL']]
    anova['Sig'] = ['OK' if (pd.notna(p) and p<0.05) else ('-' if pd.isna(p) else 'NO')
                    for p in anova['p']]
    R2 = 1 - SS_r/SS_t; R2a = 1 - (SS_r/df_r)/(SS_t/(N-1))
    R = dict(via=VIA_2, k=k, factores=factores, N=N, Nf=Nf, Nc=Nc,
             n_bloques=n_b, n_replicas=n_b, tiene_bloques=True, tiene_centro=True,
             y_bar=y, SS_total=SS_t, SS_blocks=SS_b, SS_curv=SS_c,
             SS_resid=SS_r, SS_model=SS_m, effects=ef, all_effects=ef_all,
             terms_included=list(terms), anova=anova,
             MS_resid=MS_r, MS_blocks=MS_b, MS_curv=MS_c,
             F_blocks=F_b, p_blocks=p_b, F_curv=F_c, p_curv=p_c,
             df_blocks=df_b, df_resid=df_r, yf_bar=yf, yc_bar=yc,
             R2=R2, R2_adj=R2a,
             main_effects=_calc_main_effects(df, factores),
             interactions=_calc_interactions(df, factores),
             t_crit=stats.t.ppf(0.975, df_r))
    R['lenth'] = calcular_lenth(ef, df_resid=df_r)
    return R

# ---------- Via 3 ----------
def calcular_via3(df, factores):
    k = len(factores); N = len(df); G = df['Y'].sum(); y = G/N
    SS_t = (df['Y']**2).sum() - G**2/N
    ef = _calc_effects(df, factores, Nf=N)
    L = calcular_lenth(ef, df_resid=0)
    for n_, e in ef.items(): e['activo'] = abs(e['efecto']) > L['ME']
    act = [n_ for n_, e in ef.items() if e['activo']]
    ina = [n_ for n_, e in ef.items() if not e['activo']]
    SS_p = sum(ef[n_]['SS'] for n_ in ina); df_p = len(ina)
    MS_p = SS_p/df_p if df_p > 0 else np.nan
    for n_, e in ef.items():
        e['MS'] = e['SS']
        if e['activo'] and df_p >= 1 and MS_p > 0:
            e['F'] = e['MS']/MS_p; e['p'] = stats.f.sf(e['F'], 1, df_p)
        else: e['F'] = np.nan; e['p'] = np.nan
    rows = [(n_,e['SS'],1,e['MS'],e['F'],e['p']) for n_,e in ef.items() if e['activo']]
    if df_p > 0: rows.append(('Error (pooled)', SS_p, df_p, MS_p, np.nan, np.nan))
    anova = pd.DataFrame(rows, columns=['Fuente','SS','GL','MS','F','p'])
    if df_p > 0:
        anova['F_crit'] = [stats.f.ppf(0.95,gl,df_p) if pd.notna(gl) and gl>0 else np.nan
                           for gl in anova['GL']]
    else: anova['F_crit'] = np.nan
    anova['Sig'] = ['OK' if (pd.notna(p) and p<0.05) else
                    ('EXP' if pd.isna(p) and n_!='Error (pooled)' else '-')
                    for p, n_ in zip(anova['p'], anova['Fuente'])]
    SS_m = sum(ef[n_]['SS'] for n_ in act)
    R2 = SS_m/SS_t if SS_t>0 else 0
    R = dict(via=VIA_3, k=k, factores=factores, N=N, Nf=N,
             n_bloques=0, n_replicas=1, tiene_bloques=False, tiene_centro=False,
             y_bar=y, SS_total=SS_t, SS_resid=np.nan, SS_model=SS_m,
             effects=ef, all_effects=ef, terms_included=act,
             anova=anova, activos=act, inactivos=ina,
             SS_pooled=SS_p, df_pooled=df_p, MS_pooled=MS_p,
             df_resid=0, R2=R2, R2_adj=None,
             main_effects=_calc_main_effects(df, factores),
             interactions=_calc_interactions(df, factores),
             t_crit=stats.t.ppf(0.975, max(1, L['m'])))
    R['lenth'] = L
    return R

# ---------- Graficos ----------
def _gen_pareto(R):
    terms = list(R['effects'].keys())
    if R['via'] == VIA_3:
        tvals = [abs(R['effects'][n]['efecto'])/max(R['lenth']['PSE'],1e-9) for n in R['effects']]
    else:
        tvals = [abs(R['effects'][n]['F']**0.5) if pd.notna(R['effects'][n].get('F'))
                 else abs(R['effects'][n]['efecto'])/max(R['lenth']['PSE'],1e-9)
                 for n in R['effects']]
    pvals = [R['effects'][n]['p'] if pd.notna(R['effects'][n].get('p'))
             else (0.01 if R['effects'][n].get('activo') else 0.5) for n in R['effects']]
    if R['tiene_centro']:
        terms.append('Curvatura'); tvals.append(abs(R['F_curv']**0.5)); pvals.append(R['p_curv'])
    order = np.argsort(tvals)
    fig, ax = plt.subplots(figsize=(10, max(4, 0.35*len(terms)+2)))
    colors = ['#d9534f' if p<0.05 else '#7fb3d5' for p in np.array(pvals)[order]]
    ax.barh([terms[i] for i in order], [tvals[i] for i in order], color=colors, edgecolor='black')
    if R['via'] == VIA_3:
        umb = R['lenth']['ME']/max(R['lenth']['PSE'],1e-9); lab = f'ME/PSE = {umb:.2f} (Lenth)'
    else:
        umb = R['t_crit']; lab = f'|t| critico = {umb:.2f}'
    ax.axvline(umb, color='red', linestyle='--', label=lab)
    ylo, yhi = ax.get_ylim(); ax.set_ylim(ylo, yhi+0.8)
    ax.text(umb, yhi+0.10, f'  {lab}  ', color='red', fontsize=10, fontweight='bold',
            ha='center', va='bottom',
            bbox=dict(boxstyle='round,pad=0.3', facecolor='white', edgecolor='red', alpha=0.95))
    ax.set_xlabel('|t| / |efecto| estandarizado')
    ax.set_title('Diagrama de Pareto' + (' (exploratorio)' if R['via']==VIA_3 else ''))
    ax.legend(loc='lower right'); plt.tight_layout()
    fn = 'pareto.png'; plt.savefig(fn, bbox_inches='tight'); plt.close(); return fn

def _gen_seminormal(R):
    L = R['lenth']; names = list(R['effects'].keys())
    ae = np.abs([R['effects'][n]['efecto'] for n in names]); m = len(ae)
    o = np.argsort(ae); a = ae[o]; ns = [names[i] for i in o]
    pp = (np.arange(1, m+1)-0.5)/m
    q = stats.norm.ppf((1+pp)/2)
    fig, ax = plt.subplots(figsize=(8.5, 5))
    ax.scatter(q, a, color='steelblue', s=90, edgecolor='black', zorder=3)
    for qi, yi, ni in zip(q, a, ns):
        ax.annotate(ni, (qi, yi), xytext=(5,3), textcoords='offset points', fontsize=8)
    nm = a < L['ME']
    sl, ic = np.polyfit(q[nm] if nm.sum()>=2 else q, a[nm] if nm.sum()>=2 else a, 1)
    lx = np.array([q.min()-0.1, q.max()+0.1])
    ax.plot(lx, sl*lx+ic, 'k--', label='Referencia')
    ax.axhline(L['ME'], color='red', linestyle='--', label=f'ME={L["ME"]:.3f}')
    ax.set_xlabel('Cuantiles half-normal'); ax.set_ylabel('|Efecto|')
    ax.set_title(f'Semi-normal (Daniel) PSE={L["PSE"]:.3f}')
    ax.legend(fontsize=8); ax.grid(alpha=0.3); plt.tight_layout()
    fn = 'seminormal.png'; plt.savefig(fn, bbox_inches='tight'); plt.close(); return fn

def _gen_main_effects(R, spec):
    nf = len(R['factores'])
    # EL TRUCO ESTÁ AQUÍ: sharey=True
    fig, axes = plt.subplots(1, nf, figsize=(4*nf, 4), sharey=True)
    
    if nf == 1: axes = [axes]
    
    for i, (ax, f) in enumerate(zip(axes, R['factores'])):
        me = R['main_effects'][f]
        ax.plot([-1,1], [me['low'], me['high']], 'o-', markersize=10)
        ax.set_xticks([-1,1]); ax.set_xticklabels(['-1','+1'])
        ax.set_xlabel(f)
        
        # Solo poner el título del Eje Y en el primer gráfico
        if i == 0:
            ax.set_ylabel('Media de Y')
            
        ax.set_title(f'Efecto principal: {f}'); ax.grid(alpha=0.3)
        
        if spec:
            ax.axhspan(spec['target']-spec['tol'], spec['target']+spec['tol'],
                       alpha=0.1, color='green')
                       
    plt.tight_layout()
    fn = 'efectos_principales.png'; plt.savefig(fn, bbox_inches='tight'); plt.close(); return fn


def _gen_interactions(R, custom_order=None):
    es_fb = False
    sig = [n for n in R['effects'] if len(n.split('*'))==2]
    
    # NUEVO: Si el usuario definió un orden personalizado, lo usamos
    if custom_order is not None and len(custom_order) > 0:
        sig = [n for n in custom_order if n in R['effects']]
        
    if not sig: return []
    files = []
    chunk_size = 3 # Máximo 3 gráficos por fila

    # Procesar en bloques de 3
    for i in range(0, len(sig), chunk_size):
        chunk = sig[i:i+chunk_size]
        
        # Crear figura para el bloque actual
        fig, axes = plt.subplots(1, len(chunk), figsize=(5*len(chunk), 4))
        if len(chunk) == 1: 
            axes = [axes]
            
        for ax, name in zip(axes, chunk):
            # CORRECCIÓN PARA INTERACCIONES DE 2, 3 O 4 VÍAS
            parts = name.split('*')
            if len(parts) == 2:
                f1, f2 = parts
                it = R['interactions'][(f1,f2)]
                ax.plot([-1,1], [it['mll'], it['mhl']], 'o-', label=f'{f2}=-1')
                ax.plot([-1,1], [it['mlh'], it['mhh']], 's--', label=f'{f2}=+1')
                ax.set_xticks([-1,1]); ax.set_xticklabels([f'{f1}=-1', f'{f1}=+1'])
                ax.set_xlabel(f1); ax.set_ylabel('Media de Y')
                suf = ' (exploratoria)' if R['via'] == VIA_3 else (' (NO significativa)' if es_fb else '')
                ax.set_title(f'Interaccion {name}{suf}'); ax.legend(); ax.grid(alpha=0.3)
            
        plt.tight_layout()
        # Guardar la imagen del bloque (ej. interacciones_1.png, interacciones_2.png)
        fn_bloque = f'interacciones_grupo_{i//chunk_size + 1}.png'
        plt.savefig(fn_bloque, bbox_inches='tight'); plt.close()
        files.append(fn_bloque)

    return files


def _gen_residuales(df, R):
    terms = R.get('terms_included', list(R['effects'].keys()))
    parts = [f'Q("{f}")' for f in R['factores'] if f in terms]
    formula = 'Q("Y") ~ ' + ' + '.join(parts) if parts else 'Q("Y") ~ 1'
    if R['tiene_bloques']: formula = 'Q("Y") ~ Q("Grupo") + ' + formula.split('~')[1].strip()
    if R['tiene_centro']: formula += ' + Q("Curvatura")'
    mfit = smf.ols(formula, data=df).fit()
    df2 = df.copy(); df2['fit'] = mfit.fittedvalues; df2['res'] = mfit.resid
    df2 = df2.reset_index(drop=True); df2['orden'] = range(1, len(df2)+1)
    fig, axes = plt.subplots(1, 3, figsize=(18, 4))
    stats.probplot(df2['res'], dist='norm', plot=axes[0])
    axes[0].set_title('(a) Normalidad de residuales')
    axes[1].scatter(df2['fit'], df2['res'], alpha=0.7, edgecolor='k')
    axes[1].axhline(0, color='red', linestyle='--'); axes[1].grid(alpha=0.3)
    axes[1].set_title('(b) Residuales vs Ajustados')
    axes[2].scatter(df2['orden'], df2['res'], alpha=0.7, edgecolor='k')
    axes[2].axhline(0, color='red', linestyle='--'); axes[2].grid(alpha=0.3)
    axes[2].set_title('(c) Residuales vs Orden')
    plt.tight_layout()
    fn = 'residuales.png'; plt.savefig(fn, bbox_inches='tight'); plt.close(); return fn

def _gen_cubo(R):
    k = R['k']
    if k not in (2,3): return None
    fn = 'cubo.png'
    if k == 3:
        fig = plt.figure(figsize=(9,7)); ax = fig.add_subplot(111, projection='3d')
        v = [(-1,-1,-1,'1'),(1,-1,-1,'2'),(-1,1,-1,'3'),(1,1,-1,'4'),
             (-1,-1,1,'5'),(1,-1,1,'6'),(-1,1,1,'7'),(1,1,1,'8')]
        e = [(0,1),(0,2),(0,4),(1,3),(1,5),(2,3),(2,6),(3,7),(4,5),(4,6),(5,7),(6,7)]
        for i,j in e: ax.plot([v[i][0],v[j][0]],[v[i][1],v[j][1]],[v[i][2],v[j][2]],'b-',lw=1.5)
        for x,y,z,l in v:
            ax.scatter(x,y,z,color='red',s=100,edgecolor='black',zorder=5)
            ax.text(x*1.15,y*1.15,z*1.15,l,fontsize=12,fontweight='bold')
        ax.scatter(0,0,0,color='green',s=200,marker='*',edgecolor='black',zorder=5)
        ax.set_xlabel(R['factores'][0]); ax.set_ylabel(R['factores'][1]); ax.set_zlabel(R['factores'][2])
        ax.set_xticks([-1,0,1]); ax.set_yticks([-1,0,1]); ax.set_zticks([-1,0,1])
        ax.set_title('Cubo 2^3 + centro')
    else:
        fig, ax = plt.subplots(figsize=(7,7))
        v = [(-1,-1,'1'),(1,-1,'2'),(-1,1,'3'),(1,1,'4')]
        for a,b in [(0,1),(0,2),(1,3),(2,3)]: ax.plot([v[a][0],v[b][0]],[v[a][1],v[b][1]],'b-',lw=1.5)
        for x,y,l in v:
            ax.scatter(x,y,color='red',s=120,edgecolor='black',zorder=5)
            ax.text(x*1.12,y*1.12,l,fontsize=13,fontweight='bold',ha='center')
        ax.scatter(0,0,color='green',s=250,marker='*',edgecolor='black',zorder=5)
        ax.set_xlabel(R['factores'][0]); ax.set_ylabel(R['factores'][1])
        ax.set_xticks([-1,0,1]); ax.set_yticks([-1,0,1])
        ax.set_xlim(-1.3,1.3); ax.set_ylim(-1.3,1.3)
        ax.set_title('Diseno 2^2 + centro'); ax.grid(alpha=0.3)
    plt.tight_layout(); plt.savefig(fn, dpi=150, bbox_inches='tight'); plt.close(); return fn

def _gen_curvatura(R):
    if not R['tiene_centro']: return None
    yf, yc = R['yf_bar'], R['yc_bar']; d = yf - yc
    if abs(d) < 1e-9: return None
    fn = 'curvatura.png'
    x = np.linspace(-1,1,100); yc2 = d*x**2 + yc
    fig, ax = plt.subplots(figsize=(8,5))
    ax.plot(x, np.full_like(x, yf), 'k--', lw=2, label=f'Lineal: {yf:.2f}')
    ax.plot(x, yc2, 'b-', lw=2, label='Centro-esquinas')
    ax.scatter([-1,1], [yf,yf], color='red', s=120, edgecolor='black', zorder=5,
               label=f'Esquinas ({yf:.2f})')
    ax.scatter([0], [yc], color='green', s=120, edgecolor='black', zorder=5,
               label=f'Centro ({yc:.2f})')
    ax.annotate('', xy=(0,yf), xytext=(0,yc),
                arrowprops=dict(arrowstyle='<->', color='purple', lw=2))
    ax.text(0.06, (yf+yc)/2, f'{d:+.2f}\n(curvatura)', color='purple', fontsize=11)
    ax.set_xlabel('Desalineacion'); ax.set_ylabel('Y')
    ax.set_title('Curvatura'); ax.set_xticks([-1,0,1])
    ax.set_xticklabels(['-1','0','+1'])
    ax.legend(fontsize=9); ax.grid(alpha=0.3)
    plt.tight_layout(); plt.savefig(fn, dpi=150, bbox_inches='tight'); plt.close(); return fn

def _gen_spec(df, R, spec):
    cond = df.groupby(R['factores'])['Y'].mean().reset_index()
    lo = spec['target'] - spec['tol']; hi = spec['target'] + spec['tol']
    
    fig, ax = plt.subplots(figsize=(14, 6))
    cols = ['green' if lo<=v<=hi else 'red' for v in cond['Y']]
    bars = ax.bar(range(len(cond)), cond['Y'], color=cols, edgecolor='black')
    
    labels = []
    for _, row in cond.iterrows():
        partes = []
        for f in R['factores']:
            v = row[f]
            if v == -1: partes.append(f"{f}: -1")
            elif v == 1: partes.append(f"{f}: +1")
            else: partes.append(f"{f}: 0")
        labels.append("\n".join(partes))
    
    ax.axhline(spec['target'], color='green', linestyle='--', label=f"Objetivo {spec['target']}")
    ax.axhspan(lo, hi, alpha=0.15, color='green', label=f"Spec [{lo}, {hi}]")
    
    ax.set_xticks(range(len(cond)))
    ax.set_xticklabels(labels, rotation=45, ha='right', fontsize=9)
    ax.set_ylabel('Media de Y'); ax.set_title('Medias por condicion vs Spec')
    ax.legend(loc='upper right')
    
    for i, v in enumerate(cond['Y']):
        ax.text(i, v + 0.3, f'{v:.2f}', ha='center', fontsize=8)
    
    plt.tight_layout()
    fn = 'especificacion.png'; plt.savefig(fn, bbox_inches='tight'); plt.close(); return fn

# ---------- Ecuacion ----------
def _ecuacion_str(R):
    if R['via'] == VIA_3:
        sig = [n for n,e in R['effects'].items() if e.get('activo')]
    else:
        sig = [n for n,e in R['effects'].items() if e.get('p',1)<0.05]
    if not sig: return None
    return f"Y = {R['y_bar']:.4f} " + " ".join(
        f"{R['effects'][n]['coef']:+.4f}*{n.replace('*','*')}" for n in sig)

# ---------- Funciones de Reporte y Balance ----------

def generar_tabla_regresion(df, R):
    """Genera la tabla de coeficientes estilo Excel y el resumen de ajuste global."""
    terms = R.get('terms_included', list(R['effects'].keys()))
    
    # 1. Construir la fórmula para statsmodels
    parts = []
    for f in R['factores']:
        if f in terms:
            parts.append(f'Q("{f}")')
    
    for t in terms:
        if '*' in t:
            factors_in_term = t.split('*')
            term_formula = ':'.join(f'Q("{f}")' for f in factors_in_term)
            parts.append(term_formula)
            
    formula = 'Q("Y") ~ ' + ' + '.join(parts) if parts else 'Q("Y") ~ 1'
    
    if R.get('tiene_bloques'): 
        formula = 'Q("Y") ~ Q("Grupo") + ' + formula.split('~')[1].strip()
    if R.get('tiene_centro'): 
        formula += ' + Q("Curvatura")'

    # 2. Ajustar el modelo
    mfit = smf.ols(formula, data=df).fit()
    
    # 3. Tabla de Coeficientes
    coef_df = pd.DataFrame({
        'Término': mfit.params.index,
        'Coef': mfit.params.values,
        'Error Est.': mfit.bse.values,
        't stat': mfit.tvalues.values,
        'Valor p': mfit.pvalues.values,
    })
    
    # Limpiar nombres (quitar Q(" ") y cambiar : por *)
    coef_df['Término'] = coef_df['Término'].str.replace('Q("', '').str.replace('")', '').str.replace(':', ' * ')
    
    # Intervalos de confianza
    conf_int = mfit.conf_int()
    coef_df['Inferior 95%'] = conf_int[0].values
    coef_df['Superior 95%'] = conf_int[1].values
    
    # VIF (Varianza inflada)
    coef_df['VIF'] = 1.0 
    
    # 4. Resumen de Ajuste Global
    fit_df = pd.DataFrame({
        'Métrica': ['Multiple R', 'R Square', 'Adjusted R Square', 'Standard Error', 'Observations', 'AIC', 'BIC (SBC)'],
        'Valor': [
            np.sqrt(mfit.rsquared), 
            mfit.rsquared, 
            mfit.rsquared_adj, 
            np.sqrt(mfit.mse_resid), 
            len(df), 
            mfit.aic, 
            mfit.bic
        ]
    })
    
    return coef_df, fit_df


def generar_texto_pareto(R):
    """Genera el texto explicativo del Pareto calculando los grados de libertad dinámicamente."""
    df_error = R['df_resid']
    t_crit = R['t_crit']
    
    total_terms = len(R['all_effects'])
    terms_usados = len(R['terms_included'])
    terminos_agrupados = total_terms - terms_usados
    
    df_total = R['N'] - 1
    df_puro = df_total - total_terms
    
    texto = f"Para validar visualmente la significancia de los efectos, se utiliza el estadístico t de cada factor. Este se compara contra un valor crítico de t derivado de los {df_error} grados de libertad del Error Puro del ANOVA. Con un nivel de significancia α = 0.05, el umbral es t ≈ {t_crit:.2f}. "
    
    if terminos_agrupados > 0:
        texto += f"\n\nNota metodológica: El modelo inicial evaluó {total_terms} términos ({R['k']} factores principales y {total_terms - R['k']} interacciones). Para aumentar la sensibilidad de la prueba, las {terminos_agrupados} interacciones que resultaron no significativas fueron eliminadas y agrupadas dentro del término del error (proceso de pooling). Por lo tanto, el Residual final cuenta con {df_error} grados de libertad: los {df_puro} correspondientes a las {R['N']} observaciones menos los {total_terms} términos iniciales, más los {terminos_agrupados} grados de libertad de las interacciones agrupadas. Esto permite obtener un valor crítico de t más preciso para el diagrama de Pareto."
    
    return texto


def generar_tabla_balance_neto(R, objetivo='max'):
    """Genera la tabla de balance neto adaptada al objetivo (maximizar o minimizar)."""
    y_bar = R['y_bar']
    
    # Filtrar solo los términos significativos
    sig_terms = [n for n, e in R['effects'].items() if (e.get('p', 1) < 0.05) or (R['via'] == VIA_3 and e.get('activo'))]
    
    rows = []
    rows.append({'Concepto (Factor / Interacción)': 'Base del modelo (Intercepto)', 
                 'Efecto (Puntos)': round(y_bar, 2), 
                 'Impacto porcentual (%)': 100.00, 
                 'Tipo de Efecto': 'Punto de partida'})
    
    main_effects = []
    interactions = []
    
    for name in sig_terms:
        e = R['effects'][name]
        eff = e['efecto']
        pct = (eff / y_bar) * 100
        
        # Lógica de beneficio según el objetivo
        if objetivo == 'max':
            es_beneficio = eff > 0
            tipo = 'Beneficio' if es_beneficio else 'Penalización'
        else:
            es_beneficio = eff < 0 # Si busco minimizar, un efecto negativo es bueno
            tipo = 'Beneficio (Reduce)' if es_beneficio else 'Penalización (Aumenta)'
        
        row_data = {
            'Concepto (Factor / Interacción)': name.replace('*', ' x '),
            'Efecto (Puntos)': round(eff, 2),
            'Impacto porcentual (%)': round(pct, 2),
            'Tipo de Efecto': tipo
        }
        
        if '*' not in name:
            main_effects.append(row_data)
        else:
            interactions.append(row_data)
            
    sum_main = sum(r['Efecto (Puntos)'] for r in main_effects)
    pct_main = sum(r['Impacto porcentual (%)'] for r in main_effects)
    sum_int = sum(r['Efecto (Puntos)'] for r in interactions)
    pct_int = sum(r['Impacto porcentual (%)'] for r in interactions)
    
    for r in main_effects: rows.append(r)
    if main_effects:
        rows.append({'Concepto (Factor / Interacción)': 'Suma de Efectos Principales', 
                     'Efecto (Puntos)': round(sum_main, 2), 
                     'Impacto porcentual (%)': round(pct_main, 2), 
                     'Tipo de Efecto': 'Totales'})
                     
    for r in interactions: rows.append(r)
    if interactions:
        rows.append({'Concepto (Factor / Interacción)': 'Suma de Interacciones', 
                     'Efecto (Puntos)': round(sum_int, 2), 
                     'Impacto porcentual (%)': round(pct_int, 2), 
                     'Tipo de Efecto': 'Totales'})
                     
    final_val = y_bar + sum_main + sum_int
    final_pct = 100.00 + pct_main + pct_int
    rows.append({'Concepto (Factor / Interacción)': 'Resultado Final Estimado', 
                 'Efecto (Puntos)': round(final_val, 2), 
                 'Impacto porcentual (%)': round(final_pct, 2), 
                 'Tipo de Efecto': 'Balance Neto'})
                 
    return pd.DataFrame(rows)