# app.py
# Interfaz Streamlit con flujo tipo Minitab

import streamlit as st
import pandas as pd
import numpy as np
import os
import io
import zipfile
from datetime import date
from docx import Document

import calc_engine as ce

from docx.shared import Inches

def _insertar_tabla_word(doc, df_tabla):
    """Convierte un DataFrame de Pandas en una tabla bonita de Word."""
    cols = list(df_tabla.columns)
    t = doc.add_table(rows=1, cols=len(cols))
    t.style = 'Table Grid'
    hdr_cells = t.rows[0].cells
    for i, col_name in enumerate(cols):
        hdr_cells[i].text = str(col_name)
        for paragraph in hdr_cells[i].paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
                run.font.size = Pt(9)
    
    for _, row in df_tabla.iterrows():
        row_cells = t.add_row().cells
        for i, col_name in enumerate(cols):
            val = row[col_name]
            if pd.isna(val):
                row_cells[i].text = "—"
            elif isinstance(val, float):
                row_cells[i].text = f"{val:.4f}"
            else:
                row_cells[i].text = str(val)
            for paragraph in row_cells[i].paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(9)
    doc.add_paragraph() # Espacio después de la tabla

st.set_page_config(page_title="DOE 2^k - Modo Minitab", layout="wide")

for k, v in {
    'stage': 'upload', 'raw': None, 'df': None, 'via': None,
    'factores': None, 'respuesta': None, 'spec': None,
    'R_full': None, 'R_reduced': None, 'selected_terms': None,
    'nombre': '', 'estudiante': '', 'curso': '',
    'excel_bytes': None, 'word_bytes': None, 'figs_dict': None,
}.items():
    st.session_state.setdefault(k, v)

# ---------- Deteccion inteligente de columnas ----------
def _detectar_columnas(raw):
    factores, respuestas, descartar = [], [], []
    id_names = {'run','corrida','corrida_id','id','n_corrida','numero_corrida',
                'obs','observacion','index','unnamed: 0'}
    resp_exact = {'y', 'response', 'respuesta', 'resp'}
    resp_hints = ['rep', 'replica', 'block', 'bloque']

    for c in raw.columns:
        nombre_lower = str(c).strip().lower()
        vals = raw[c].dropna().unique().tolist()
        n_unicos = len(vals)

        if nombre_lower in id_names or n_unicos <= 1:
            descartar.append(c)
            continue

        es_resp_hint = (nombre_lower in resp_exact or any(h in nombre_lower for h in resp_hints))
        if es_resp_hint and pd.api.types.is_numeric_dtype(raw[c]):
            respuestas.append(c)
            continue

        if n_unicos == 2:
            factores.append(c)
            continue

        if n_unicos == 3:
            try:
                vals_num = [float(v) for v in vals]
                if set(vals_num) == {-1.0, 0.0, 1.0}:
                    factores.append(c)
                    continue
            except (ValueError, TypeError):
                pass

        if pd.api.types.is_numeric_dtype(raw[c]):
            respuestas.append(c)
            continue

        factores.append(c)
    return factores, respuestas, descartar


# ---------- Sidebar ----------
with st.sidebar:
    st.header("Configuracion")
    if st.button("Reiniciar"):
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.rerun()

    uploaded = st.file_uploader("Archivo Excel o CSV", type=['xlsx','csv'])
    if uploaded:
        try:
            if uploaded.name.endswith('.xlsx'):
                raw = pd.read_excel(uploaded)
            else:
                raw = pd.read_csv(uploaded)
            raw.columns = raw.columns.astype(str).str.strip()
            st.session_state.raw = raw
            st.success(f"OK: {raw.shape[0]} filas x {raw.shape[1]} cols")
        except Exception as e:
            st.error(f"Error: {e}")

    if st.session_state.raw is not None:
        raw = st.session_state.raw
        st.markdown("### Columnas")
        fact_cands, resp_cands, _ = _detectar_columnas(raw)

        cols = st.multiselect("Factores", raw.columns.tolist(), default=fact_cands)
        resp = st.multiselect("Respuesta(s)", raw.columns.tolist(), default=resp_cands)

        # --- PANEL PARA RENOMBRAR ---
        mapping_nombres = {}
        if len(cols) > 0:
            with st.expander("✏️ Renombrar factores (opcional)"):
                st.markdown("Si tus columnas se llaman A, B, C, D, puedes ponerles su nombre real aquí:")
                for c in cols:
                    nuevo = st.text_input(f"Nombre para '{c}':", value=c, key=f"rename_{c}")
                    if nuevo.strip() != "" and nuevo.strip() != c:
                        mapping_nombres[c] = nuevo.strip()
        
        # --- PANEL PARA CONFIGURAR NIVELES -1 / +1 ---
        mapeos_niveles = {}
        if len(cols) > 0:
            with st.expander("⚙️ Configurar Niveles (-1 / +1)", expanded=False):
                st.markdown("Elige qué nivel corresponde al **bajo (-1)** y cuál al **alto (+1)**.")
                for c in cols:
                    col_real = mapping_nombres.get(c, c)
                    vals = list(dict.fromkeys(raw[col_real].dropna().tolist()))
                    
                    if set(vals).issubset({-1,0,1,-1.0,0.0,1.0}):
                        continue
                        
                    if len(vals) == 2:
                        v_bajo = st.selectbox(f"'{col_real}' - Nivel BAJO (-1):", vals, index=0, key=f"low_{c}")
                        v_alto = [v for v in vals if v != v_bajo][0]
                        st.caption(f"→ Nivel ALTO (+1): **{v_alto}**")
                        mapeos_niveles[col_real] = {v_bajo: -1, v_alto: 1}
                        
                    elif len(vals) == 3:
                        v_bajo = st.selectbox(f"'{col_real}' - Nivel BAJO (-1):", vals, index=0, key=f"low_{c}")
                        v_alto = st.selectbox(f"'{col_real}' - Nivel ALTO (+1):", [v for v in vals if v != v_bajo], index=0, key=f"high_{c}")
                        v_mid = [v for v in vals if v != v_bajo and v != v_alto][0]
                        st.caption(f"→ Centro (0): **{v_mid}**")
                        mapeos_niveles[col_real] = {v_bajo: -1, v_mid: 0, v_alto: 1}
            st.markdown("---")

        via = st.radio("Via",
                       options=[ce.VIA_1, ce.VIA_2, ce.VIA_3],
                       format_func=lambda v: f"{ce.VIA_INFO[v]['nombre']}")
        objetivo = st.radio("Objetivo del análisis:", ["Maximizar la respuesta", "Minimizar la respuesta"], index=0)

        use_spec = st.checkbox("Especificacion (target +/- tol)")
        spec = None
        if use_spec:
            c1, c2 = st.columns(2)
            target = c1.number_input("Target", value=24.0)
            tol = c2.number_input("Tolerancia", value=4.0)
            spec = {'target': target, 'tol': tol}

        st.markdown("### Contexto")
        nombre = st.text_input("Proceso", "Analisis 2^k")
        estudiante = st.text_input("Estudiante", "")
        curso = st.text_input("Curso", "Asignacion")

        if st.button("Procesar", type='primary'):
            try:
                # 1. Aplicar el renombrado
                if mapping_nombres:
                    raw = raw.rename(columns=mapping_nombres)
                    cols = [mapping_nombres.get(c, c) for c in cols]

                rp = raw.copy()
                
                # 2. Aplicar los niveles manuales del estudiante
                for f in cols:
                    if f in mapeos_niveles:
                        rp[f] = rp[f].map(mapeos_niveles[f])
                    elif set(rp[f].dropna().unique()).issubset({-1,0,1,-1.0,0.0,1.0}):
                        continue
                    else:
                        vals = list(dict.fromkeys(rp[f].dropna().tolist()))
                        if len(vals) == 2:
                            lo, hi = min(vals), max(vals)
                            rp[f] = rp[f].map({lo: -1, hi: 1})
                
                # 3. Convertir respuestas a numérico y preparar datos
                for c in resp:
                    rp[c] = pd.to_numeric(rp[c], errors='raise')
                df = ce.preparar_datos(rp, cols, resp)

                # 4. Guardar en el estado de la sesión y avanzar al Paso 1
                st.session_state.df = df
                st.session_state.via = via
                st.session_state.objetivo = 'max' if objetivo == "Maximizar la respuesta" else 'min'
                st.session_state.factores = cols
                st.session_state.respuesta = resp
                st.session_state.spec = spec
                st.session_state.nombre = nombre
                st.session_state.estudiante = estudiante
                st.session_state.curso = curso
                st.session_state.stage = 'design'
                st.session_state.R_full = None
                st.session_state.R_reduced = None
                st.session_state.selected_terms = None
                st.session_state.excel_bytes = None
                st.session_state.word_bytes = None
                st.session_state.figs_dict = None
                st.rerun()
            except Exception as e:
                st.error(f"Error: {e}")

# ---------- Main ----------
if st.session_state.stage == 'upload':
    st.title("DOE 2^k - Modo Minitab")
    st.markdown("""
    ### Como funciona

    Esta app reproduce el flujo de Minitab para disenos factoriales 2^k:

    1. **Diseno** - revisa la estructura
    2. **Modelo completo** - ANOVA con todos los terminos, Pareto, semi-normal
    3. **Refinar modelo** - TU decides que terminos quitar (aqui esta el esfuerzo)
    4. **Diagnostico** - residuales, efectos principales, interacciones
    5. **Reporte** - descarga tablas, graficos y plantilla Word

    ### ¿En qué se parece a Minitab?

    - **Interpretar** el Pareto y el semi-normal para decidir que terminos son reales
    - **Seleccionar** terminos en el modelo reducido
    - **Verificar** que los resultados tengan sentido
    - **Diagnosticar** los residuales
    - **Escribir tus conclusiones**

    Al igual que en Minitab, la app te da los numeros, pero las decisiones son tuyas.
    """)
    st.info("Empieza subiendo un archivo en la barra lateral")

elif st.session_state.stage == 'design':
    df = st.session_state.df
    factores = st.session_state.factores
    via = st.session_state.via
    st.title("Paso 1 - Diseno experimental")
    
    n_grupos = df['Grupo'].nunique()
    n_total = len(df)
    n_esquinas = 2 ** len(factores)
    if via == ce.VIA_1:
        expl = f"N = {n_total} = {n_esquinas} esquinas x {n_grupos} replicas puras"
    elif via == ce.VIA_2:
        n_centros = int(df['Curvatura'].sum() / n_grupos) if n_grupos else 0
        expl = f"N = {n_total} = {n_esquinas + n_centros} puntos ({n_esquinas} esquinas + {n_centros} centro) x {n_grupos} bloques"
    else:
        expl = f"N = {n_total} = {n_esquinas} esquinas (sin replicas)"

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Factores", len(factores))
    c2.metric("N total", n_total)
    c3.metric("Bloques" if via == ce.VIA_2 else "Replicas", n_grupos if via != ce.VIA_3 else 1)
    c4.metric("Via", via)
    
    st.info(f"**De donde viene N:** {expl}")
    cond = df[df['Curvatura']==0][factores].drop_duplicates().reset_index(drop=True)
    cond.insert(0, 'Condicion', range(1, len(cond)+1))
    st.markdown("### Combinaciones del diseno")
    st.dataframe(cond, use_container_width=True)
    if st.button("Ir a Modelo completo", type='primary'):
        st.session_state.stage = 'full_model'; st.rerun()

elif st.session_state.stage == 'full_model':
    st.title("Paso 2 - Modelo completo")
    df = st.session_state.df
    factores = st.session_state.factores
    via = st.session_state.via

    if st.session_state.R_full is None:
        with st.spinner("Calculando..."):
            if via == ce.VIA_1:
                R = ce.calcular_via1_con_terminos(df, factores, terms=None)
            elif via == ce.VIA_2:
                R = ce.calcular_via2_con_terminos(df, factores, terms=None)
            else:
                R = ce.calcular_via3(df, factores)
            st.session_state.R_full = R
    R = st.session_state.R_full

    st.markdown("### 1. Tabla ANOVA")
    st.markdown("""
    Esta tabla es el **corazón del análisis**. Compara la variación que produce cada factor (la "señal") contra el ruido del experimento.
    
    **¿Cómo leer las columnas?**
    - **SS (Suma de Cuadrados):** La cantidad de variación que explica ese término. Números más grandes = mayor efecto.
    - **GL (Grados de Libertad):** Cuánta información independiente tiene ese término (para factores en 2^k, siempre es 1).
    - **MS (Cuadrado Medio):** Es la división de SS entre GL. Es la varianza promedio.
    - **F (Estadístico F):** Es la división de MS del término entre MS del Error. Si F es grande, el efecto es real y no puro ruido.
    - **p (Valor p):** La probabilidad de que ese efecto sea solo ruido. **Si p < 0.05, el término es estadísticamente significativo** (lo marcamos con ✅).
    """)
    
    st.dataframe(R['anova'].round(4), use_container_width=True)
    
    st.write(f"**R2 = {R['R2']:.4f}**" + (f" | **R2 ajustado = {R['R2_adj']:.4f}**" if R.get('R2_adj') else ""))
    st.caption("💡 **El R2** (de 0 a 1) indica qué porcentaje de la variabilidad total del proceso está explicada por el modelo. Un R2 mayor a 0.80 suele ser muy bueno.")

    if via == ce.VIA_2:
        st.markdown("### 2. Diseño y Curvatura (Exclusivo Vía 2)")
        st.markdown("""
        Como este diseño tiene un **punto central**, podemos evaluar si el comportamiento es curvo (no lineal).
        - **Esquema del cubo**: Muestra dónde están ubicadas las 8 esquinas (±1) y el punto central (0,0,0).
        - **Gráfico de Curvatura**: Compara la media del centro contra la media de las esquinas. Si el centro está muy alejado de la línea de las esquinas, hay curvatura.
        """)
        c1, c2 = st.columns(2)
        fn_cub = ce._gen_cubo(R)
        if fn_cub: c1.image(fn_cub, caption="Esquema del Diseño (Cubo 3D)", use_container_width=True)
        
        fn_curv = ce._gen_curvatura(R)
        if fn_curv: c2.image(fn_curv, caption="Gráfico de Curvatura", use_container_width=True)

    st.markdown("### 3. Diagrama de Pareto")
    st.image(ce._gen_pareto(R), use_container_width=True)

    st.markdown("### 4. Grafico semi-normal (Daniel / Lenth)")
    st.image(ce._gen_seminormal(R), use_container_width=True)
    
    # NUEVO: Leyenda de letras
    leyenda = ", ".join([f"**{chr(65+i)}** = {f}" for i, f in enumerate(factores)])
    st.caption(f"Leyenda de factores: {leyenda}")

    st.markdown("""
    ### Cómo leer estos gráficos (¡Muy importante!)
    
    **1. Diagrama de Pareto**
    - Cada barra horizontal es un término del modelo (factor o interacción).
    - La **longitud** de la barra indica cuán grande es su efecto sobre la respuesta.
    - Las barras **ROJAS** cruzan la línea vertical roja (el umbral) → son estadísticamente significativas.
    - Las barras **AZULES** no la cruzan → no son significativas (probablemente ruido).
    - El gráfico está ordenado de menor a mayor: la barra de arriba es el efecto más grande.
    
    **2. Gráfico Semi-normal (Daniel / Lenth)**
    - Cada punto es un término del modelo.
    - El **Eje X** representa los cuantiles teóricos (no te preocupes por el número exacto).
    - El **Eje Y** representa el valor absoluto del efecto (|efecto|).
    - Los puntos que siguen la línea diagonal son efectos pequeños, indistinguibles del ruido.
    - Los puntos que se **DESPEGAN** de la línea diagonal son efectos grandes, probablemente reales.
    - La línea horizontal roja es el umbral ME (Margen de Error). Los puntos por encima de ella son "activos".
    """)

    with st.expander("🤔 Preguntas de reflexion (toma notas para tu informe)"):
        st.markdown("""
        - ¿Coinciden las barras rojas del Pareto con los puntos despegados del semi-normal?
        - ¿Cual es el termino con la barra mas larga?
        - ¿Hay algun termino que parezca "marginal" (que este justo en el limite)?
        """)

    c1, c2 = st.columns(2)
    if c1.button("Volver a Diseno"): st.session_state.stage = 'design'; st.rerun()
    if c2.button("Refinar modelo", type='primary'):
        st.session_state.stage = 'refine'; st.rerun()

elif st.session_state.stage == 'refine':
    st.title("Paso 3 - Refinar modelo")
    
    df = st.session_state.df
    factores = st.session_state.factores
    via = st.session_state.via
    R_full = st.session_state.R_full
    all_terms = list(R_full['effects'].keys())

    # Explicaciones condicionales segun la via
    if via == ce.VIA_3:
        st.markdown("""
        **En esta Vía 3, el método de Lenth ya ha hecho el trabajo por ti.**
        
        Como este diseño **no tiene réplicas**, no existe un "Error puro" tradicional para calcular los p-valores. Para solucionarlo, el programa toma todos los términos que son ruido, los junta y crea un **Error Pooled** (Error agrupado). Ese es el motor que permite calcular si los términos `[ACTIVO]` son realmente significativos.
        
        Por eso verás que los términos de `[RUIDO]` dicen **p=Pooled**. 
        
        *Nota importante:* En esta vía **no hay casillas para marcar o desmarcar**. El algoritmo ya decidió qué es activo y qué es ruido basándose en el umbral ME. Solo revisa la tabla y presiona el botón para avanzar al diagnóstico.
        """)
    else:
        st.markdown("""
        **Aqui construyes el modelo final.** 
        - **Deja marcados** los terminos que identificaste como activos en el Pareto/semi-normal.
        - **Desmarca** los terminos ruido.
        - Luego presiona "Re-calcular".
        """)

    st.markdown("### Terminos disponibles")
    selected = []
    for term in all_terms:
        e = R_full['effects'][term]
        c1, c2, c3, c4 = st.columns([3, 1, 1, 1])
        display_term = term.replace('*', ' x ')
        
        if via == ce.VIA_3:
            activo = e.get('activo', False)
            c1.write(f"{'[ACTIVO]' if activo else '[RUIDO]'} {display_term}")
            if activo: selected.append(term)
        else:
            checked = c1.checkbox(display_term, value=True, key=f"chk_{term}")
            if checked: selected.append(term)
            
        c2.write(f"Efecto: {e['efecto']:+.4f}")
        p_val = e.get('p')
        c3.write(f"p={p_val:.4f}" if pd.notna(p_val) else "p=Pooled" if via == ce.VIA_3 else "p=-")

    if len(selected) == 0:
        st.warning("Selecciona al menos 1 termino.")
    else:
        if via == ce.VIA_3:
            if st.button("Continuar al Diagnostico", type='primary'):
                st.session_state.R_reduced = R_full
                st.session_state.selected_terms = selected
                st.session_state.stage = 'diagnostics'
                st.rerun()
        else:
            if st.button("Re-calcular con terminos seleccionados", type='primary'):
                with st.spinner("Ajustando modelo reducido..."):
                    try:
                        if via == ce.VIA_1:
                            R_red = ce.calcular_via1_con_terminos(df, factores, terms=selected)
                        elif via == ce.VIA_2:
                            R_red = ce.calcular_via2_con_terminos(df, factores, terms=selected)
                        st.session_state.R_reduced = R_red
                        st.session_state.selected_terms = selected
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error: {e}")

    if st.session_state.R_reduced is not None:
        R_red = st.session_state.R_reduced
        st.markdown("### 2. ANOVA del modelo reducido (Final)")
        st.markdown("""
        Este es el ANOVA final, una vez que has eliminado el ruido. 
        
        **¿Qué cambió respecto al modelo completo?**
        - Al quitar los términos no significativos, su variabilidad (SS y GL) pasa a formar parte del **Residual** (o del **Error pooled** si es Vía 3).
        - Ahora los Grados de Libertad del error son más grandes, lo que hace que el análisis sea **más sensible** para detectar los efectos verdaderos.
        - El **R2 ajustado** suele subir, porque el modelo es más simple y preciso.
        """)
        
        st.dataframe(R_red['anova'].round(4), use_container_width=True)
        
        st.write(f"**R2 = {R_red['R2']:.4f}**" + (f" | **R2 ajustado = {R_red['R2_adj']:.4f}**" if R_red.get('R2_adj') else ""))
        
        # --- INICIO EXTRA OPCIONAL: Intervalos de Confianza ---
        with st.expander("🔬 Opcional (Nivel Avanzado): Coeficientes con Intervalos de Confianza"):
            st.markdown("""
            Esta tabla muestra el error estándar y el rango donde probablemente se encuentra el verdadero coeficiente.
            **Si el intervalo NO incluye el cero**, el término es estadísticamente significativo.
            """)
            t_crit = R_red.get('t_crit', 2.0)
            mse = R_red.get('MS_resid', 0)
            N_total = R_red.get('N', 1)
            err_est = np.sqrt(mse / N_total) if mse > 0 else 0
            
            coef_data = [{
                'Término': 'Intercepto (β0)', 
                'Coef': round(R_red['y_bar'], 4), 
                'Error Est.': round(err_est, 4), 
                't': np.nan, 
                'IC 95% Inf': np.nan, 
                'IC 95% Sup': np.nan
            }]
            
            for term, e in R_red['effects'].items():
                t_val = e['coef'] / err_est if err_est > 0 else 0
                ic_inf = e['coef'] - (t_crit * err_est)
                ic_sup = e['coef'] + (t_crit * err_est)
                coef_data.append({
                    'Término': term.replace('*', ' x '),
                    'Coef': round(e['coef'], 4),
                    'Error Est.': round(err_est, 4),
                    't': round(t_val, 2),
                    'IC 95% Inf': round(ic_inf, 4),
                    'IC 95% Sup': round(ic_sup, 4)
                })
            st.dataframe(pd.DataFrame(coef_data), use_container_width=True)
        # --- FIN EXTRA OPCIONAL ---

        with st.expander("🤔 Preguntas de reflexion para tu informe"):
            st.markdown("""
            - ¿Todos los terminos que dejaste siguen siendo significativos (p < 0.05)?
            - ¿El R2 ajustado subio respecto al modelo completo?
            - Si un termino perdio significancia, ¿quitaste una interaccion importante?
            """)

    c1, c2 = st.columns(2)
    if c1.button("Volver a Modelo completo"): st.session_state.stage = 'full_model'; st.rerun()
    if c2.button("Diagnostico", type='primary'):
        st.session_state.stage = 'diagnostics'; st.rerun()

elif st.session_state.stage == 'diagnostics':
    st.title("Paso 4 - Diagnostico")
    df = st.session_state.df
    R = st.session_state.R_reduced if st.session_state.R_reduced else st.session_state.R_full

    st.markdown("### 1. Diagnostico de Residuales")
    st.markdown("""
    Los residuales son lo que el modelo **no** puede explicar. Para que el ANOVA sea confiable, deben cumplir 3 reglas:
    - **(a) Normalidad**: En el primer gráfico, los puntos deben seguir la línea diagonal recta. Si se despegan en los extremos (forma de S), la normalidad falla.
    - **(b) Homocedasticidad (Varianza constante)**: En el segundo gráfico, la nube de puntos debe ser pareja. Si se abre en forma de embudo o trompeta, la varianza no es constante.
    - **(c) Independencia**: En el tercer gráfico, no debe haber tendencias (curvas, líneas rectas). Los puntos deben estar dispersos al azar.
    """)
    st.image(ce._gen_residuales(df, R), use_container_width=True)
    with st.expander("🤔 Preguntas de reflexión (toma notas para tu informe)"):
        st.markdown("""
        - ¿Los puntos del gráfico (a) se pegan a la línea roja? 
        - ¿La nube del gráfico (b) tiene forma de embudo o es pareja?
        - ¿Ves alguna tendencia en el gráfico (c)?
        - **Conclusión:** ¿Se cumplen los supuestos? (Si sí, los p-valores del ANOVA son 100% confiables).
        """)

    st.markdown("### 2. Efectos Principales")
    st.markdown("""
    Muestra cómo cambia la media de la respuesta (Y) cuando el factor pasa de su nivel bajo (-1) a su nivel alto (+1).
    """)
    st.image(ce._gen_main_effects(R, st.session_state.spec), use_container_width=True)
    
    # --- NUEVO: Tabla de Efectos Principales Detallada ---
    st.markdown("#### Tabla Resumen de Efectos")
    filas_ef = []
    for f in st.session_state.factores:
        low_val = R['main_effects'][f]['low']
        high_val = R['main_effects'][f]['high']
        eff = R['main_effects'][f]['eff']
        filas_ef.append({
            'Factor': f,
            'Media Nivel Bajo (-1)': round(low_val, 4),
            'Media Nivel Alto (+1)': round(high_val, 4),
            'Efecto (+1 vs -1)': round(eff, 4)
        })
    st.dataframe(pd.DataFrame(filas_ef), use_container_width=True)
    st.caption("✍️ **Para tu informe:** En la columna de interpretación, escribe qué significa cada efecto en tu proceso (ej. 'El soporte grande favorece la viabilidad').")
    
    # --- NUEVO: Tabla de Ranking de Mejores Condiciones ---
    st.markdown("#### 🏆 Ranking de Condiciones Observadas (Top 5)")
    st.markdown("""
    Estas son las combinaciones que dieron los mejores resultados en el experimento. 
    *(Ordenadas de mayor a menor según la media de Y)*
    """)
    
    # 1. Calcular medias por combinación
    cond_medias = df[df['Curvatura']==0].groupby(st.session_state.factores)['Y'].mean().reset_index()
    
    # 2. Agregar el número de corrida (1 al N, en el orden original)
    cond_medias.insert(0, 'Corrida', range(1, len(cond_medias)+1))
    
    # 3. Ordenar y sacar el Top 5
    # Ordenar según el objetivo (max = descendente, min = ascendente)
    es_min = st.session_state.get('objetivo', 'max') == 'min'
    cond_medias = cond_medias.sort_values(by='Y', ascending=es_min).head(5).reset_index(drop=True)
    
    # 4. Construir la tabla
    ranking_data = []
    for idx, row in cond_medias.iterrows():
        fila = {'Posición': idx + 1, 'Corrida': int(row['Corrida'])}
        for f in st.session_state.factores:
            val = row[f]
            if val == -1: fila[f] = "Nivel Bajo (-1)"
            elif val == 1: fila[f] = "Nivel Alto (+1)"
            else: fila[f] = "Centro (0)"
        fila['Media de Y'] = round(row['Y'], 4)
        ranking_data.append(fila)
    
    st.dataframe(pd.DataFrame(ranking_data), use_container_width=True)

    with st.expander("🤔 Preguntas de reflexión (toma notas para tu informe)"):
        st.markdown("""
        - ¿Cuál es el gráfico con la pendiente más pronunciada? (Ese es el factor dominante).
        - ¿Los signos de las pendientes coinciden con los efectos que viste en el Pareto?
        - Si hay una especificación (banda verde), ¿qué niveles del factor te mantienen dentro de la spec?
        - **En la tabla de Ranking:** ¿Qué factores se repiten en las primeras posiciones? ¿Qué combinación recomendarías?
        """)

    st.markdown("### 3. Interacciones")
    st.markdown("""
    Muestra si el efecto de un factor **depende** del nivel de otro factor.
    - Líneas **paralelas**: No hay interacción (los factores actúan de forma independiente).
    - Líneas que **se cruzan**: Interacción antagónica (el efecto de uno se invierte si cambias el nivel del otro).
    - Líneas que **se separan**: Interacción sinérgica (los efectos se potencian).
    """)

    # --- NUEVO: Selector de orden de interacciones ---
    todas_interacciones = [n for n in R['effects'] if len(n.split('*'))==2]
    
    if todas_interacciones:
        orden_interacciones = st.multiselect(
            "🛠️ Selecciona y ordena las interacciones a graficar (el orden de selección será el orden de los gráficos):",
            options=todas_interacciones,
            default=todas_interacciones, 
            key="selector_interacciones"
        )
        st.session_state.orden_interacciones = orden_interacciones 
    else:
        orden_interacciones = []
    # -------------------------------------------------


    fns_int = ce._gen_interactions(R, custom_order=st.session_state.get('orden_interacciones', []))
    
    if fns_int:
        for fn in fns_int:
            if os.path.exists(fn):
                st.image(fn, use_container_width=True)
    else:
        st.info("Ninguna interaccion significativa en el modelo.")


    with st.expander("🤔 Preguntas de reflexión (toma notas para tu informe)"):
        st.markdown("""
        - ¿Las líneas son paralelas o no?
        - Si se cruzan, ¿en qué punto sucede? ¿Qué implicación práctica tiene?
        - ¿Conviene usar el mismo nivel del factor 2 sin importar el nivel del factor 1?
        """)

    if st.session_state.spec:
        st.markdown("### 4. Cumplimiento de Especificacion")
        st.markdown("""
        Muestra qué combinaciones cumplen con la meta (banda verde). Verde = cumple, Rojo = no cumple.
        """)
        st.image(ce._gen_spec(df, R, st.session_state.spec), use_container_width=True)
        with st.expander("🤔 Preguntas de reflexión (toma notas para tu informe)"):
            st.markdown("""
            - ¿Cuántas barras son verdes?
            - ¿Qué combinación de niveles comparten todas las barras rojas?
            - ¿El centro (0,0,0) cumple con la especificación?
            """)

    c1, c2 = st.columns(2)
    if c1.button("Volver a Refinar"): st.session_state.stage = 'refine'; st.rerun()
    if c2.button("Avanzar al Reporte Final", type='primary'):
        st.session_state.stage = 'report'; st.rerun()

elif st.session_state.stage == 'report':
    st.title("Paso 5 - Reporte final")
    R = st.session_state.R_reduced if st.session_state.R_reduced else st.session_state.R_full
    df = st.session_state.df
    nombre = st.session_state.nombre
    slug = nombre.replace(' ', '_').replace('/', '-')

    st.markdown("### 1. Ecuacion de Regresion")
    st.write(f"**Beta 0 (Media global / Intercepto)** = {R['y_bar']:.4f}")
    eq = ce._ecuacion_str(R)
    if eq: 
        st.code(eq, language=None)
    else: 
        st.info("Ningun termino significativo. La ecuacion es solo la media (Beta 0).")

    st.markdown("""
    #### Cómo interpretar esta ecuación (¡Muy importante para el informe!)
    Esta fórmula te permite **predecir** el valor de Y para cualquier combinación de niveles (-1 o +1).
    
    - **Beta 0 (Intercepto):** Es el valor promedio de Y cuando todos los factores están en 0 (el centro del diseño).
    - **El signo (+ / -):** Indica la dirección. Si el coeficiente de `Temperatura` es `+13.10`, significa que al subir la temperatura de -1 a +1, la respuesta **sube** 13.10 unidades.
    - **La magnitud:** Un número más grande (ignorando el signo) indica un factor con mayor impacto en el proceso.
    - **Términos de interacción (ej. A x B):** Significa que el efecto de A no es constante, sino que cambia dependiendo del nivel de B.
    
    *Ejemplo de uso:* Si sustituyes los factores por -1 o +1 y haces la suma, obtienes el Y predicho para esa condición.
    """)
    
    with st.expander("🤔 Preguntas de reflexión para tu informe"):
        st.markdown("""
        - ¿Los signos de los coeficientes de la ecuación coinciden con lo que viste en los efectos principales?
        - Si tuvieras que recomendar una configuración para maximizar (o minimizar) Y, ¿qué valores (-1 o +1) le darías a los factores con coeficientes positivos y negativos?
        - ¿Cómo usarías esta ecuación en la práctica industrial o biológica?
        """)
    st.markdown("### 2. Balance Neto de la Configuración Óptima")
    # Solo mostrar si es un diseño sin bloques ni curvatura (como el Problema 1 o 3)
    if not R.get('tiene_bloques') and not R.get('tiene_centro'):
        balance_df = ce.generar_tabla_balance_neto(R, st.session_state.get('objetivo', 'max'))
        st.dataframe(balance_df, use_container_width=True)
    else:
        st.info("El balance neto no se muestra para diseños con bloques o puntos centrales (use la ecuación de regresión).")

    st.markdown("### 2. Generar Archivos")
    if st.button("Generar descargas (Excel, Word, PNGs)", type='primary'):
        with st.spinner("Generando reporte completo..."):
            # 1. Generar Tablas
            tablas = {'ANOVA': R['anova']}
            Rf = st.session_state.R_full
            if Rf:
                ef_tab = pd.DataFrame([
                    {'Termino': n, 'Efecto': e['efecto'], 'Coef': e['coef'],
                     'p': e.get('p', np.nan),
                     'En modelo final': 'Si' if n in (st.session_state.selected_terms or list(Rf['effects'].keys())) else 'No'}
                    for n, e in Rf['effects'].items()
                ])
                tablas['Efectos_todos'] = ef_tab

            # 2. Generar Excel
            buf_xlsx = io.BytesIO()
            with pd.ExcelWriter(buf_xlsx) as w:
                for hoja, t in tablas.items():
                    t.to_excel(w, sheet_name=hoja[:31], index=False)

            # 3. GENERAR GRÁFICOS PRIMERO (para poder insertarlos en el Word)
            figs = {
                'Pareto': ce._gen_pareto(R),
                'SemiNormal': ce._gen_seminormal(R),
                'Residuales': ce._gen_residuales(df, R),
                'EfectosPrincipales': ce._gen_main_effects(R, st.session_state.spec),
            }

            # Usar el orden elegido en el paso anterior

            # Obtener el orden de forma segura
            orden_guardado = st.session_state.get('orden_interacciones', None)
            
            # Si el orden guardado no existe o está vacío, usar el orden por defecto
            if not orden_guardado:
                orden_guardado = [n for n in R['effects'] if len(n.split('*'))==2]
                
            fns_int = ce._gen_interactions(R, custom_order=orden_guardado)

            if fns_int: figs['Interacciones'] = fns_int
            if st.session_state.via == ce.VIA_2:
                fn_cub = ce._gen_cubo(R)
                if fn_cub: figs['Cubo'] = fn_cub
                fn_curv = ce._gen_curvatura(R)
                if fn_curv: figs['Curvatura'] = fn_curv
            if st.session_state.spec:
                figs['Spec'] = ce._gen_spec(df, R, st.session_state.spec)

            # 4. GENERAR WORD (Insertando todo en orden)
            doc = Document()
            doc.add_heading(f"Informe de Análisis 2^k - {nombre}", 0)
            doc.add_paragraph(f"Estudiante: {st.session_state.estudiante}")
            doc.add_paragraph(f"Curso: {st.session_state.curso}")
            doc.add_paragraph(f"Fecha: {date.today().strftime('%Y-%m-%d')}")
            
            doc.add_heading("1. Diseño experimental", level=1)
            doc.add_paragraph(f"Factores: {', '.join(st.session_state.factores)}")
            doc.add_paragraph(f"N = {len(df)}")
            doc.add_paragraph("Descripción del diseño: [Escribe aquí el contexto y los objetivos del experimento]")
            
            doc.add_heading("2. Análisis y resultados", level=1)

            # Balance Neto para el Word
            if not R.get('tiene_bloques') and not R.get('tiene_centro'):
                doc.add_heading("2.1. Balance Neto de la Configuración Óptima", level=2)
                doc.add_paragraph("La tabla muestra el desglose de ganancias y penalizaciones al configurar los factores en su nivel óptimo.")
                balance_df = ce.generar_tabla_balance_neto(R, st.session_state.get('objetivo', 'max'))
                _insertar_tabla_word(doc, balance_df)
                
                # Ajustamos la numeración del ANOVA para que sea el 2.2
                doc.add_heading("2.2. ANOVA", level=2)
            else:
                doc.add_heading("2.1. ANOVA", level=2)
            
            _insertar_tabla_word(doc, R['anova'])
            doc.add_paragraph(f"R2 = {R['R2']:.4f}" + (f" | R2 ajustado = {R['R2_adj']:.4f}" if R.get('R2_adj') else ""))
            doc.add_paragraph("Interpretación del ANOVA: [Explica qué términos son significativos]")
            
            doc.add_heading("2.2. Diagrama de Pareto", level=2)
            if 'Pareto' in figs and os.path.exists(figs['Pareto']):
                doc.add_picture(figs['Pareto'], width=Inches(6))
            doc.add_paragraph("Interpretación: [Explica qué barras cruzan el umbral y por qué]")
            
            doc.add_heading("2.3. Gráfico semi-normal (Lenth)", level=2)
            if 'SemiNormal' in figs and os.path.exists(figs['SemiNormal']):
                doc.add_picture(figs['SemiNormal'], width=Inches(6))
            doc.add_paragraph("Interpretación: [Explica qué puntos se despegan de la línea de referencia]")
            
            doc.add_heading("2.4. Efectos principales", level=2)
            # Generar tabla de efectos principales para el Word
            ef_princ_data = []
            for f in st.session_state.factores:
                ef_princ_data.append({
                    'Factor': f,
                    'Media Nivel Bajo (-1)': round(R['main_effects'][f]['low'], 4),
                    'Media Nivel Alto (+1)': round(R['main_effects'][f]['high'], 4),
                    'Efecto (+1 vs -1)': round(R['main_effects'][f]['eff'], 4)
                })
            _insertar_tabla_word(doc, pd.DataFrame(ef_princ_data))
            if 'EfectosPrincipales' in figs and os.path.exists(figs['EfectosPrincipales']):
                doc.add_picture(figs['EfectosPrincipales'], width=Inches(6))
            doc.add_paragraph("Interpretación de los efectos: [Escribe aquí qué significa cada efecto en el proceso]")

            # --- NUEVO: Ranking Top 5 ---
            doc.add_heading("2.5. Condiciones Recomendadas (Ranking)", level=2)
            
            # Calcular el ranking para el Word
            cond_medias_word = df[df['Curvatura']==0].groupby(st.session_state.factores)['Y'].mean().reset_index()
            cond_medias_word.insert(0, 'Corrida', range(1, len(cond_medias_word)+1))
            cond_medias_word = cond_medias_word.sort_values(by='Y', ascending=es_min).head(5).reset_index(drop=True)
            
            # Preparar la tabla para Word
            tabla_word = cond_medias_word.copy()
            tabla_word.insert(0, 'Posición', range(1, len(tabla_word)+1))
            
            # Reemplazar -1 y 1 por texto más amigable
            for f in st.session_state.factores:
                tabla_word[f] = tabla_word[f].apply(lambda x: "Bajo (-1)" if x==-1 else ("Alto (+1)" if x==1 else "Centro (0)"))
                
            tabla_word = tabla_word.rename(columns={'Y': 'Media de Y'})
            
            _insertar_tabla_word(doc, tabla_word)
            doc.add_paragraph("Recomendación operativa: [Escribe aquí la combinación que sugieres y por qué]")
            # -----------------------------

            if 'Interacciones' in figs and figs['Interacciones']:
                doc.add_heading("2.6. Interacciones", level=2)
                for fn_int in figs['Interacciones']:
                    if os.path.exists(fn_int):
                        doc.add_picture(fn_int, width=Inches(6)) # Ancho ideal para Word
                doc.add_paragraph("Interpretación: [Explica si las líneas son paralelas o se cruzan, y qué implican]")


            if st.session_state.via == ce.VIA_2:
                doc.add_heading("2.7. Curvatura y Diseño", level=2)
                if 'Cubo' in figs and os.path.exists(figs['Cubo']):
                    doc.add_picture(figs['Cubo'], width=Inches(4.5))
                if 'Curvatura' in figs and os.path.exists(figs['Curvatura']):
                    doc.add_picture(figs['Curvatura'], width=Inches(4.5))
                doc.add_paragraph("Interpretación: [Explica si hay comportamiento no lineal]")

            doc.add_heading("2.8. Diagnóstico de residuales", level=2)
            if 'Residuales' in figs and os.path.exists(figs['Residuales']):
                doc.add_picture(figs['Residuales'], width=Inches(6.5))
            doc.add_paragraph("Interpretación: [Analiza los 3 supuestos: normalidad, homocedasticidad e independencia]")
            
            doc.add_heading("2.9. Ecuación de regresión", level=2)
            doc.add_paragraph(f"Beta 0 = {R['y_bar']:.4f}")
            if eq: doc.add_paragraph(eq, style='Intense Quote')
            doc.add_paragraph("Interpretación de los coeficientes: [Explica qué significa cada término en la ecuación]")
       
            doc.add_heading("3. Conclusiones", level=1)
            doc.add_paragraph("1. [Conclusión 1]")
            doc.add_paragraph("2. [Conclusión 2]")
            
            doc.add_heading("4. Recomendaciones operativas", level=1)
            doc.add_paragraph("[Qué niveles fijar o evitar para optimizar el proceso]")
            
            buf_word = io.BytesIO()
            doc.save(buf_word)

            # 5. Guardar en el estado y preparar ZIP
            st.session_state.excel_bytes = buf_xlsx.getvalue()
            st.session_state.word_bytes = buf_word.getvalue()
            st.session_state.figs_dict = figs
        st.success("¡Reporte generado con tablas y gráficos insertados!")

    if st.session_state.excel_bytes:
        st.download_button("Descargar Tablas (Excel)", st.session_state.excel_bytes,
                           file_name=f'{slug}_tablas.xlsx',
                           mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        st.download_button("Descargar Plantilla de informe (Word)", st.session_state.word_bytes,
                           file_name=f'{slug}_plantilla.docx',
                           mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

        buf_zip = io.BytesIO()
        with zipfile.ZipFile(buf_zip, 'w') as z:
            z.writestr(f'{slug}_tablas.xlsx', st.session_state.excel_bytes)
            z.writestr(f'{slug}_plantilla.docx', st.session_state.word_bytes)
            for name, fns in st.session_state.figs_dict.items():
                if isinstance(fns, list):
                    for fn in fns:
                        if fn and os.path.exists(fn):
                            z.write(fn, arcname=os.path.basename(fn))
                elif fns and os.path.exists(fns):
                    z.write(fns, arcname=os.path.basename(fns))


        st.download_button("Descargar Todo en ZIP", buf_zip.getvalue(),
                           file_name=f'{slug}_analisis.zip', mime='application/zip')

    c1, c2 = st.columns(2)
    if c1.button("Volver a Diagnostico"): st.session_state.stage = 'diagnostics'; st.rerun()
    if c2.button("Empezar de nuevo"):
        for k in list(st.session_state.keys()): del st.session_state[k]
        st.rerun()