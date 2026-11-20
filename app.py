from pathlib import Path

import altair as alt
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Negocios cautivos · Cotizaciones", page_icon="📃", layout="wide")

ARCHIVO_DEFECTO = Path(__file__).parent / "BASE_DE_DATOS_PROYECTO_CAUTIVOS_migrada.xlsx"
HOJAS = ["CLIENTES", "PROVEEDORES", "RUTAS", "TIPOS_UNIDAD", "COTIZACIONES", "DETALLE_SERVICIOS",
         "PROVEEDORES_COTIZACION", "CONDICIONES_CARGOS", "PENDIENTES_REVISION", "CONTROL_COBERTURA", "DICCIONARIO_DATOS"]
TOL = 0.02
NUM_FMT = "%.2f"  # Es siempre en dolar o hay variación en cordobas / otra moneda? confirmar con ker. 


# ------------------------------------------------------------------ carga de datos  (Funcion piloto)
@st.cache_data(show_spinner="Leyendo el libro…")
def cargar(origen) -> dict:
    hojas = pd.read_excel(origen, sheet_name=HOJAS, dtype=object)
    out = {}
    for nombre, df in hojas.items():
        df = df.replace({"": np.nan})
        for c in df.columns:
            if not pd.api.types.is_numeric_dtype(df[c]):
                df[c] = df[c].apply(lambda v: v.strip() if isinstance(v, str) else v).replace({"": np.nan})
        out[nombre] = df
    num = lambda df, cols: df.assign(**{c: pd.to_numeric(df[c], errors="coerce") for c in cols})
    out["DETALLE_SERVICIOS"] = num(out["DETALLE_SERVICIOS"], ["costo", "venta", "importe_sin_clasificar", "utilidad_original", "orden"])
    out["CONDICIONES_CARGOS"] = num(out["CONDICIONES_CARGOS"], ["costo", "venta", "utilidad_original", "valor", "valor_maximo"])
    out["COTIZACIONES"] = num(out["COTIZACIONES"], ["total_costo_original", "total_venta_original",
                                                     "total_utilidad_original", "total_sin_clasificar_original"])
    return construir_modelo(out)


def construir_modelo(t: dict) -> dict:
    """Recalcula en pandas los campos derivados (no depende de valores en caché de Excel)."""
    cli = t["CLIENTES"][["id_cliente", "nombre_mostrar"]].rename(columns={"nombre_mostrar": "cliente"})
    rut = t["RUTAS"][["id_ruta", "origen_original", "destino_original", "origen_pais", "destino_pais"]].rename(
        columns={"origen_original": "origen", "destino_original": "destino"})
    uni = t["TIPOS_UNIDAD"][["id_unidad", "descripcion"]].rename(columns={"descripcion": "unidad"})
    prv = t["PROVEEDORES"][["id_proveedor", "nombre_original"]].rename(columns={"nombre_original": "proveedor"})

    base_cols = ["id_cotizacion", "id_cliente", "id_ruta", "id_unidad", "descripcion_original", "referencia_mencionada",
                 "mercancia_mencionada", "variante", "observaciones", "bloque_origen", "total_costo_original",
                 "total_venta_original", "total_utilidad_original", "total_sin_clasificar_original", "celda_total_original"]
    cot = (t["COTIZACIONES"][base_cols].merge(cli, on="id_cliente", how="left")
           .merge(rut, on="id_ruta", how="left").merge(uni, on="id_unidad", how="left"))

    det = t["DETALLE_SERVICIOS"].drop(columns=[c for c in ["cliente", "origen", "destino", "unidad", "utilidad_calculada",
                                                          "diferencia_utilidad", "control_utilidad"] if c in t["DETALLE_SERVICIOS"]])
    det["utilidad_calculada"] = np.where(det["costo"].notna() & det["venta"].notna(), det["venta"] - det["costo"], np.nan)
    det["diferencia_utilidad"] = (det["utilidad_calculada"] - det["utilidad_original"]).round(2)

    def control(r):
        if pd.notna(r.diferencia_utilidad):
            if abs(r.diferencia_utilidad) > TOL: return "Diferencia con utilidad original"
            if r.utilidad_calculada < 0 and r.venta == 0: return "Negativa con venta en cero: no es pérdida confirmada"
            return "OK"
        if pd.notna(r.utilidad_original) and pd.isna(r.venta): return "Sin venta informada: no es pérdida confirmada"
        if pd.notna(r.utilidad_calculada): return "Sin utilidad original"
        return ""
    det["control_utilidad"] = det.apply(control, axis=1)

    cond = t["CONDICIONES_CARGOS"].copy()
    cargos = cond[cond["costo"].notna() | cond["venta"].notna()]

    falt = det["estado_costo"].isin(["Vacío", "Texto en origen"]).astype(int) + det["estado_venta"].isin(["Vacío", "Texto en origen"]).astype(int)
    guion = (det["estado_costo"] == "Guion en origen").astype(int) + (det["estado_venta"] == "Guion en origen").astype(int)
    agg = det.assign(_f=falt, _g=guion, _nc=(det["estado_costo"] == "No clasificado").astype(int)).groupby("id_cotizacion").agg(
        n_servicios=("id_servicio", "count"), costo_base=("costo", "sum"), venta_base=("venta", "sum"),
        sin_clasificar_base=("importe_sin_clasificar", "sum"), importes_faltantes=("_f", "sum"),
        importes_con_guion=("_g", "sum"), no_clasificados=("_nc", "sum"))
    cag = cargos.groupby("id_cotizacion").agg(cargos_costo=("costo", "sum"), cargos_venta=("venta", "sum"))
    cot = cot.merge(agg, on="id_cotizacion", how="left").merge(cag, on="id_cotizacion", how="left")
    for c in ["n_servicios", "costo_base", "venta_base", "sin_clasificar_base", "importes_faltantes",
              "importes_con_guion", "no_clasificados", "cargos_costo", "cargos_venta"]:
        cot[c] = cot[c].fillna(0)

    def completitud(r):
        if r.n_servicios == 0: return "Sin servicios"
        if r.importes_faltantes > 0: return "Incompleta"
        if r.no_clasificados > 0: return "Importes sin clasificar"
        if r.importes_con_guion > 0: return "Completa con guiones por confirmar"
        return "Completa"
    cot["estado_completitud"] = cot.apply(completitud, axis=1)
    cot["utilidad_base"] = np.where(cot["estado_completitud"].str.startswith("Completa"), cot["venta_base"] - cot["costo_base"], np.nan)
    cot["margen_base_pct"] = np.where((cot["utilidad_base"].notna()) & (cot["venta_base"] > 0),
                                      cot["utilidad_base"] / cot["venta_base"] * 100, np.nan)

    def ok(calc, orig): return True if pd.isna(orig) else abs(calc - orig) <= TOL
    def conciliacion(r):
        if pd.isna(r.total_costo_original) and pd.isna(r.total_venta_original) and pd.isna(r.total_sin_clasificar_original):
            return "Sin total original"
        if pd.notna(r.total_sin_clasificar_original):
            return "Coincide (sin clasificar)" if abs(r.sin_clasificar_base - r.total_sin_clasificar_original) <= TOL else "Difiere"
        if ok(r.costo_base, r.total_costo_original) and ok(r.venta_base, r.total_venta_original): return "Coincide con detalle base"
        if ok(r.costo_base + r.cargos_costo, r.total_costo_original) and ok(r.venta_base + r.cargos_venta, r.total_venta_original):
            return "Coincide con base + cargos"
        return "Difiere"
    cot["estado_conciliacion"] = cot.apply(conciliacion, axis=1)

    pxc = (t["PROVEEDORES_COTIZACION"].drop(columns=[c for c in ["proveedor", "cliente", "origen", "destino", "unidad", "servicio_vinculado"]
                                                    if c in t["PROVEEDORES_COTIZACION"]])
           .merge(prv, on="id_proveedor", how="left")
           .merge(det[["id_servicio", "servicio_original"]].rename(columns={"servicio_original": "servicio_vinculado"}), on="id_servicio", how="left"))
    prov_txt = pxc.groupby("id_cotizacion")["proveedor"].apply(lambda s: " · ".join(dict.fromkeys(s.dropna())))
    cot = cot.merge(prov_txt.rename("proveedores"), on="id_cotizacion", how="left")
    cat_txt = det.groupby("id_cotizacion")["categoria_servicio"].apply(lambda s: " · ".join(dict.fromkeys(s.dropna())))
    cot = cot.merge(cat_txt.rename("servicios"), on="id_cotizacion", how="left")

    dias = cond[cond["tipo_condicion"] == "Días libres"].groupby("id_cotizacion")["valor"].apply(
        lambda s: " / ".join(f"{v:g}" for v in dict.fromkeys(s.dropna())))
    cot = cot.merge(dias.rename("dias_libres"), on="id_cotizacion", how="left")

    pend = t["PENDIENTES_REVISION"].copy()
    pend["cotizaciones_relacionadas"] = pend["cotizaciones_relacionadas"].fillna("")
    n_pend = (pend.assign(c=pend["cotizaciones_relacionadas"].str.split(r",\s*")).explode("c")
              .query("c != ''").groupby("c").size())
    cot["n_pendientes"] = cot["id_cotizacion"].map(n_pend).fillna(0).astype(int)

    t.update(cot=cot, det=det.merge(cot[["id_cotizacion", "cliente", "origen", "destino", "unidad"]], on="id_cotizacion", how="left"),
             pxc=pxc.merge(cot[["id_cotizacion", "cliente", "origen", "destino", "unidad"]], on="id_cotizacion", how="left"),
             cond=cond, pend=pend)
    return t


# ------------------------------------------------------------------ utilidades UI
def ms(label, options, **kw):
    kw.setdefault("placeholder", "Todos")
    return st.multiselect(label, options, **kw)


def opciones(serie):
    return sorted(serie.dropna().astype(str).unique(), key=str.casefold)


def fmt(v, dec=2):
    return "—" if v is None or (isinstance(v, float) and np.isnan(v)) else f"{v:,.{dec}f}"


def limpio(df):
    """Texto vacío en lugar de None en columnas de texto (los números quedan intactos)."""
    df = df.copy()
    for c in df.columns:
        if not pd.api.types.is_numeric_dtype(df[c]): df[c] = df[c].astype(object).where(df[c].notna(), "")
    return df


def tabla(df, **kw):
    """Muestra una tabla: números vacíos como — (no 'None') y decimales a 2 posiciones, sin símbolo de moneda."""
    if isinstance(df, pd.DataFrame):
        df = limpio(df)
        cfg = dict(kw.pop("column_config", None) or {})
        for c in df.columns:
            if pd.api.types.is_float_dtype(df[c]):
                if df[c].isna().any():   # Streamlit muestra "None" en nulos: se pasa a texto con "—"
                    df[c] = df[c].map(lambda v: "—" if pd.isna(v) else f"{v:,.2f}")
                    cfg.setdefault(c, st.column_config.TextColumn(c, alignment="right"))
                else:
                    cfg.setdefault(c, st.column_config.NumberColumn(c, format="%.2f"))
        kw["column_config"] = cfg
    return st.dataframe(df, **kw)


def cfg_numeros(cols):
    return {c: st.column_config.NumberColumn(c, format=NUM_FMT) for c in cols}


ESTADO_COLOR = {"Completa": "#2E7D32", "Completa con guiones por confirmar": "#9E9D24", "Incompleta": "#C62828",
                "Importes sin clasificar": "#EF6C00", "Sin servicios": "#757575",
                "Coincide con detalle base": "#2E7D32", "Coincide con base + cargos": "#558B2F",
                "Coincide (sin clasificar)": "#9E9D24", "Difiere": "#C62828", "Sin total original": "#90A4AE"}


def barras(df, campo, titulo):
    d = df[campo].value_counts().rename_axis(campo).reset_index(name="cotizaciones")
    return (alt.Chart(d, title=titulo).mark_bar(cornerRadiusEnd=3)
            .encode(y=alt.Y(f"{campo}:N", sort="-x", title=None, axis=alt.Axis(labelLimit=320)), x=alt.X("cotizaciones:Q", title="Cotizaciones"),
                    color=alt.Color(f"{campo}:N", legend=None,
                                    scale=alt.Scale(domain=list(ESTADO_COLOR), range=list(ESTADO_COLOR.values()))),
                    tooltip=[campo, "cotizaciones"]).properties(height=40 * len(d) + 50))


def filtros_cotizaciones(t, clave="f"):
    cot, det, pxc = t["cot"], t["det"], t["pxc"]
    with st.sidebar:
        st.markdown("### Filtros")
        texto = st.text_input("Buscar texto", key=f"{clave}_txt", placeholder="cliente, referencia, proveedor…")
        f_cli = ms("Cliente", opciones(cot["cliente"]), key=f"{clave}_cli")
        f_prv = ms("Proveedor", opciones(pxc["proveedor"]), key=f"{clave}_prv")
        f_ori = ms("Origen", opciones(cot["origen"]), key=f"{clave}_ori")
        f_des = ms("Destino", opciones(cot["destino"]), key=f"{clave}_des")
        f_uni = ms("Unidad", opciones(cot["unidad"]), key=f"{clave}_uni")
        f_ser = ms("Servicio (categoría)", opciones(det["categoria_servicio"]), key=f"{clave}_ser")
        f_est = ms("Estado de completitud", opciones(cot["estado_completitud"]), key=f"{clave}_est")
        f_con = ms("Conciliación", opciones(cot["estado_conciliacion"]), key=f"{clave}_con")
        solo_pend = st.toggle("Solo con pendientes", key=f"{clave}_pend")
    m = pd.Series(True, index=cot.index)
    if f_cli: m &= cot["cliente"].isin(f_cli)
    if f_ori: m &= cot["origen"].isin(f_ori)
    if f_des: m &= cot["destino"].isin(f_des)
    if f_uni: m &= cot["unidad"].isin(f_uni)
    if f_est: m &= cot["estado_completitud"].isin(f_est)
    if f_con: m &= cot["estado_conciliacion"].isin(f_con)
    if f_prv: m &= cot["id_cotizacion"].isin(pxc.loc[pxc["proveedor"].isin(f_prv), "id_cotizacion"])
    if f_ser: m &= cot["id_cotizacion"].isin(det.loc[det["categoria_servicio"].isin(f_ser), "id_cotizacion"])
    if solo_pend: m &= cot["n_pendientes"] > 0
    if texto:
        campos = ["id_cotizacion", "cliente", "referencia_mencionada", "descripcion_original", "proveedores", "origen", "destino", "variante"]
        hay = cot[campos].fillna("").astype(str).agg(" ".join, axis=1).str.contains(texto, case=False, regex=False)
        m &= hay
    return cot[m]


# ------------------------------------------------------------------ páginas
def pagina_resumen(t):
    cot, pend = t["cot"], t["pend"]
    st.title("Resumen de negocios cautivos")
    st.caption("Datos migrados desde la hoja DB RAW. Los importes se muestran sin símbolo de moneda pendiente de confirmar")
    c = st.columns(5)
    c[0].metric("Cotizaciones", len(cot))
    c[1].metric("Clientes", cot["id_cliente"].nunique())
    c[2].metric("Proveedores", len(t["PROVEEDORES"]))
    c[3].metric("Rutas", len(t["RUTAS"]))
    c[4].metric("Pendientes de revisión", int((pend["estado_revision"].fillna("Pendiente") == "Pendiente").sum()))

    a, b = st.columns(2)
    a.altair_chart(barras(cot, "estado_completitud", "¿Están completas las cotizaciones?"), width="stretch")
    b.altair_chart(barras(cot, "estado_conciliacion", "¿Cuadra el detalle con el total original?"), width="stretch")

    st.subheader("Licitaciones por cliente")
    top = cot.groupby("cliente").agg(cotizaciones=("id_cotizacion", "count")).reset_index().sort_values("cotizaciones", ascending=False)
    st.altair_chart(alt.Chart(top.head(15)).mark_bar(color="#1F4E78", cornerRadiusEnd=3)
                    .encode(y=alt.Y("cliente:N", sort="-x", title=None, axis=alt.Axis(labelLimit=420)), x=alt.X("cotizaciones:Q", title="Cotizaciones"),
                            tooltip=["cliente", "cotizaciones"]).properties(height=420), width="stretch")

    st.subheader("Margen base de las licitaciones completas")
    comp = cot[cot["margen_base_pct"].notna()]
    st.caption("Solo cotizaciones con costo y venta completos. Margen = (venta base − costo base) / venta base.")
    st.altair_chart(alt.Chart(comp).mark_circle(size=70, opacity=0.75)
                    .encode(x=alt.X("venta_base:Q", title="Venta base"), y=alt.Y("margen_base_pct:Q", title="Margen base (%)"),
                            color=alt.Color("estado_conciliacion:N", title="Conciliación",
                                            scale=alt.Scale(domain=list(ESTADO_COLOR), range=list(ESTADO_COLOR.values()))),
                            tooltip=["id_cotizacion", "cliente", "origen", "destino", "proveedores",
                                     alt.Tooltip("venta_base:Q", format=",.2f"), alt.Tooltip("margen_base_pct:Q", format=".1f")])
                    .properties(height=380).interactive(), width="stretch")


def ficha(t, cid):
    cot = t["cot"].set_index("id_cotizacion").loc[cid]
    det = t["det"][t["det"]["id_cotizacion"] == cid].sort_values("orden")
    cond = t["cond"][t["cond"]["id_cotizacion"] == cid]
    pxc = t["pxc"][t["pxc"]["id_cotizacion"] == cid]
    pend = t["pend"][t["pend"]["cotizaciones_relacionadas"].str.contains(cid, regex=False)]
    notas_cli = t["cond"][(t["cond"]["ambito"] == "Cliente") & (t["cond"]["id_cliente"] == cot["id_cliente"])]

    st.markdown(f"### {cid} · {cot['cliente']}")
    ruta = f"{cot['origen']} → {cot['destino']}" if pd.notna(cot["origen"]) else "Ruta no indicada en el origen"
    st.markdown(f"**{ruta}**  ·  Unidad: **{cot['unidad'] if pd.notna(cot['unidad']) else 'no indicada'}**"
                + (f"  ·  Variante: **{cot['variante']}**" if pd.notna(cot["variante"]) else ""))
    if pd.notna(cot["descripcion_original"]): st.caption(f"Título original: {cot['descripcion_original']}")

    m = st.columns(5)
    m[0].metric("Costo base", fmt(cot["costo_base"]))
    m[1].metric("Venta base", fmt(cot["venta_base"]))
    m[2].metric("Utilidad base", fmt(cot["utilidad_base"]),
                help="Solo se calcula si la cotización está completa.")
    m[3].metric("Cargos condicionales (venta)", fmt(cot["cargos_venta"]), help="No incluidos en la base.")
    m[4].metric("Días libres", cot["dias_libres"] if pd.notna(cot["dias_libres"]) else "—")

    estado = cot["estado_completitud"]
    (st.success if estado == "Completa" else st.warning if estado.startswith("Completa") else st.error)(f"Completitud: {estado}")
    conc = cot["estado_conciliacion"]
    if conc == "Difiere":
        st.error(f"El total original ({fmt(cot['total_costo_original'])} costo / {fmt(cot['total_venta_original'])} venta) "
                 f"no coincide con el detalle ({fmt(cot['costo_base'])} / {fmt(cot['venta_base'])}).")
    else:
        st.info(f"Conciliación: {conc}")

    tabs = st.tabs([f"Servicios ({len(det)})", f"Cargos y condiciones ({len(cond)})", f"Proveedores ({len(pxc)})",
                    f"Pendientes ({len(pend)})", "Trazabilidad"])
    with tabs[0]:
        cols = ["orden", "servicio_original", "categoria_servicio", "tramo", "costo", "venta", "importe_sin_clasificar",
                "utilidad_original", "utilidad_calculada", "control_utilidad", "estado_costo", "estado_venta", "moneda", "observaciones"]
        tabla(limpio(det[cols]), hide_index=True, width="stretch")
    with tabs[1]:
        car = cond[cond["costo"].notna() | cond["venta"].notna()]
        if len(car):
            st.markdown("**Cargos opcionales o condicionales** (fuera del precio base)")
            tabla(car[["tipo_condicion", "concepto_original", "costo", "venta", "condicion_aplicacion", "celda_origen"]],
                         hide_index=True, width="stretch")
        otras = cond[cond["costo"].isna() & cond["venta"].isna() & (cond["tipo_condicion"] != "Nota completa")]
        if len(otras):
            st.markdown("**Condiciones extraídas**")
            tabla(otras[["tipo_condicion", "concepto_original", "valor", "valor_maximo", "unidad_valor", "moneda", "condicion_aplicacion"]],
                         hide_index=True, width="stretch")
        notas = pd.concat([cond[cond["tipo_condicion"].isin(["Nota completa", "Nota libre", "Tarifa incluye", "Tarifa no incluye"])], notas_cli])
        if len(notas):
            st.markdown("**Notas completas**")
            for _, n in notas.drop_duplicates("texto_completo").iterrows():
                st.markdown(f"> {n['texto_completo']}  \n<small>{n['ambito']} · {n['celda_origen']}</small>", unsafe_allow_html=True)
    with tabs[2]:
        tabla(limpio(pxc[["proveedor", "funcion", "tramo", "etiqueta_original", "servicio_vinculado", "celda_origen"]]),
                     hide_index=True, width="stretch")
    with tabs[3]:
        if len(pend):
            tabla(pend[["id_pendiente", "motivo", "dato_encontrado", "decision_requerida", "estado_revision"]],
                         hide_index=True, width="stretch")
        else:
            st.write("Sin pendientes para esta cotización.")
    with tabs[4]:
        st.write(f"Bloque en DB RAW: `{cot['bloque_origen']}`")
        if pd.notna(cot["celda_total_original"]): st.write(f"Fila de total original: `{cot['celda_total_original']}`")
        if pd.notna(cot["referencia_mencionada"]): st.write(f"Referencias mencionadas en el bloque: {cot['referencia_mencionada']}")
        if pd.notna(cot["observaciones"]): st.write(f"Observaciones de la migración: {cot['observaciones']}")


def pagina_cotizaciones(t):
    st.title("Cotizaciones")
    vista = filtros_cotizaciones(t, "cot")
    st.caption(f"{len(vista)} de {len(t['cot'])} cotizaciones. Selecciona una fila para ver su ficha.")
    cols = ["id_cotizacion", "cliente", "origen", "destino", "unidad", "variante", "proveedores", "servicios",
            "costo_base", "venta_base", "utilidad_base", "estado_completitud", "estado_conciliacion", "n_pendientes"]
    sel = tabla(limpio(vista[cols]), hide_index=True, width="stretch", height=420,
                       on_select="rerun", selection_mode="single-row",
                       column_config={"n_pendientes": st.column_config.Column("pendientes")})
    st.download_button("Descargar esta vista (CSV)", vista[cols].to_csv(index=False).encode("utf-8-sig"),
                       "cotizaciones_filtradas.csv", "text/csv")
    filas = sel.selection.rows if sel and sel.selection else []
    if filas:
        st.divider()
        ficha(t, vista.iloc[filas[0]]["id_cotizacion"])


def pagina_comparar(t):
    st.title("Comparar alternativas")
    st.caption("Solo se comparan cotizaciones con la misma ruta y la misma unidad.")
    cot = t["cot"][t["cot"]["id_ruta"].notna()].copy()
    cot["ruta_unidad"] = cot["origen"] + " → " + cot["destino"] + "  |  " + cot["unidad"].fillna("unidad no indicada")
    entre = st.toggle("Comparar entre clientes distintos", help="Por defecto solo se comparan alternativas del mismo cliente.")
    cot["grupo"] = cot["ruta_unidad"] if entre else cot["cliente"] + "  ·  " + cot["ruta_unidad"]
    grupos = cot.groupby("grupo").filter(lambda g: len(g) > 1)
    if grupos.empty:
        st.info("No hay grupos con más de una alternativa."); return
    c1, c2 = st.columns([2, 1])
    clientes = ms("Cliente", opciones(grupos["cliente"]), key="cmp_cli")
    if clientes: grupos = grupos[grupos["cliente"].isin(clientes)]
    if grupos.empty:
        c1.info("Ese cliente no tiene alternativas comparables."); return
    lista = grupos["grupo"].value_counts()
    g = c1.selectbox("Grupo comparable", lista.index, format_func=lambda x: f"{x}  ({lista[x]} alternativas)")
    alts = grupos[grupos["grupo"] == g].sort_values("venta_base")
    if alts["variante"].nunique(dropna=False) > 1 or alts["cliente"].nunique() > 1:
        st.warning("El grupo tiene variantes o clientes distintos; revisar esas filas antes de comparar precios.")

    etiquetas = {"cliente": "Cliente", "variante": "Variante", "proveedores": "Proveedores", "dias_libres": "Días libres",
                 "costo_base": "Costo base", "venta_base": "Venta base", "utilidad_base": "Utilidad base", "margen_base_pct": "Margen base",
                 "cargos_venta": "Cargos condicionales (venta)", "estado_completitud": "Completitud", "estado_conciliacion": "Conciliación"}
    resumen = alts.set_index("id_cotizacion")[list(etiquetas)]
    vista = resumen.copy()
    for c in ["costo_base", "venta_base", "utilidad_base", "cargos_venta"]: vista[c] = vista[c].map(fmt)
    vista["margen_base_pct"] = vista["margen_base_pct"].map(lambda v: "—" if pd.isna(v) else f"{v:.1f} %")
    vista = vista.fillna("—").astype(str)
    tabla(vista.rename(columns=etiquetas).T, width="stretch")

    det = t["det"][t["det"]["id_cotizacion"].isin(alts["id_cotizacion"])]
    medida = st.radio("Comparar servicios por", ["costo", "venta", "utilidad_calculada"], horizontal=True,
                      format_func={"costo": "Costo", "venta": "Venta", "utilidad_calculada": "Utilidad calculada"}.get)
    piv = det.pivot_table(index="servicio_original", columns="id_cotizacion", values=medida, aggfunc="sum")
    piv = piv.reindex(columns=alts["id_cotizacion"])
    tabla(piv.apply(lambda col: col.map(lambda v: "—" if pd.isna(v) else f"{v:,.2f}")), width="stretch")
    st.caption("— = el servicio no existe en esa alternativa o no tiene importe informado.")
    graf = alts.melt(id_vars=["id_cotizacion", "proveedores"], value_vars=["costo_base", "venta_base"], var_name="concepto", value_name="importe")
    st.altair_chart(alt.Chart(graf).mark_bar().encode(
        x=alt.X("id_cotizacion:N", sort=list(alts["id_cotizacion"]), title=None), xOffset="concepto:N",
        y=alt.Y("importe:Q", title="Importe"), color=alt.Color("concepto:N", scale=alt.Scale(range=["#90A4AE", "#1F4E78"]), title=None),
        tooltip=["id_cotizacion", "proveedores", "concepto", alt.Tooltip("importe:Q", format=",.2f")]).properties(height=320),
        width="stretch")


def pagina_proveedores(t):
    st.title("Proveedores")
    pxc, prv = t["pxc"], t["PROVEEDORES"]
    resumen = (pxc.groupby(["id_proveedor", "proveedor"])
               .agg(cotizaciones=("id_cotizacion", "nunique"), clientes=("cliente", "nunique"),
                    funciones=("funcion", lambda s: " · ".join(sorted(set(s.dropna())))))
               .reset_index().merge(prv[["id_proveedor", "variantes_escritura_original"]], on="id_proveedor", how="left")
               .sort_values("cotizaciones", ascending=False))
    q = st.text_input("Buscar proveedor")
    if q: resumen = resumen[resumen["proveedor"].str.contains(q, case=False, regex=False)]
    st.caption("Los nombres parecidos no se unieron; las posibles equivalencias están en Pendientes.")
    sel = tabla(resumen, hide_index=True, width="stretch", height=360, on_select="rerun", selection_mode="single-row")
    filas = sel.selection.rows if sel and sel.selection else []
    if filas:
        pid = resumen.iloc[filas[0]]["id_proveedor"]
        ids = pxc.loc[pxc["id_proveedor"] == pid, "id_cotizacion"].unique()
        st.subheader(f"Cotizaciones con {resumen.iloc[filas[0]]['proveedor']}")
        cot = t["cot"][t["cot"]["id_cotizacion"].isin(ids)]
        tabla(cot[["id_cotizacion", "cliente", "origen", "destino", "unidad", "costo_base", "venta_base", "estado_completitud"]],
                     hide_index=True, width="stretch")


def tema_pendiente(m):
    m = str(m)
    reglas = [("Nombre de proveedor con", "Proveedores con '/' (una o varias empresas)"),
              ("Posible equivalencia", "Posibles equivalencias de proveedores"),
              ("Encabezados de cliente", "Posibles equivalencias de clientes"),
              ("LA CIABARA", "Posibles equivalencias de clientes"),
              ("Bloque bajo el cliente", "Título no coincide con el cliente"),
              ("origen/destino", "Ruta no indicada"), ("Origen escrito", "Ubicación ambigua"),
              ("guion", "Importes con guion '$-'"), ("costo o venta", "Importes sin clasificar"),
              ("total original", "Totales que no cuadran"), ("utilidad total", "Totales que no cuadran"),
              ("utilidad original", "Utilidad que no cuadra"), ("Costo negativo", "Utilidad que no cuadra"),
              ("alternativas", "Alternativas mezcladas"), ("FLETE MX", "Alternativas mezcladas")]
    for clave, tema in reglas:
        if clave.lower() in m.lower(): return tema
    return "Datos sueltos o sin estructura"


def pagina_pendientes(t):
    st.title("Pendientes de revisión")
    pend = t["pend"].copy()
    pend["tema"] = pend["motivo"].map(tema_pendiente)
    c1, c2 = st.columns([3, 1])
    temas = c1.multiselect("Tema", opciones(pend["tema"]), placeholder="Todos")
    estado = c2.multiselect("Estado", opciones(pend["estado_revision"]), placeholder="Todos")
    if temas: pend = pend[pend["tema"].isin(temas)]
    if estado: pend = pend[pend["estado_revision"].isin(estado)]
    cnt = t["pend"]["motivo"].map(tema_pendiente).value_counts().rename_axis("tema").reset_index(name="n")
    st.altair_chart(alt.Chart(cnt).mark_bar(color="#C62828", cornerRadiusEnd=3).encode(
        y=alt.Y("tema:N", sort="-x", title=None, axis=alt.Axis(labelLimit=520)), x=alt.X("n:Q", title="Pendientes"), tooltip=["tema", "n"]).properties(height=26 * len(cnt) + 40),
        width="stretch")
    tabla(limpio(pend[["id_pendiente", "tema", "ubicacion_original", "dato_encontrado", "motivo", "decision_requerida",
                       "cotizaciones_relacionadas", "estado_revision", "decision_tomada"]]), hide_index=True, width="stretch", height=460)
    st.caption("Las decisiones se registran en la columna decision_tomada de la hoja PENDIENTES_REVISION del Excel.")


def pagina_calidad(t):
    st.title("Calidad y cobertura de la migración")
    ctrl = t["CONTROL_COBERTURA"].copy()
    revisar = ctrl[ctrl["resultado"] == "REVISAR"]
    if revisar.empty: st.success("Todos los controles con valor esperado están OK (cobertura de celdas, IDs únicos, relaciones válidas).")
    else: st.error(f"{len(revisar)} controles requieren revisión.")
    for sec, g in ctrl.groupby("seccion", sort=False):
        with st.expander(sec, expanded=sec in ("Cobertura", "Conciliación")):
            tabla(g[["indicador", "valor", "esperado", "resultado", "nota"]], hide_index=True, width="stretch")
    st.caption("Estos indicadores vienen calculados del Excel; si editas el libro, ábrelo y guárdalo en Excel para actualizarlos.")
    with st.expander("Diccionario de datos"):
        nombre_tabla = st.selectbox("Tabla", opciones(t["DICCIONARIO_DATOS"]["tabla"]))
        tabla(t["DICCIONARIO_DATOS"][t["DICCIONARIO_DATOS"]["tabla"] == nombre_tabla].drop(columns="tabla"), hide_index=True, width="stretch")


# ------------------------------------------------------------------ arranque
def main():
    with st.sidebar:
        st.markdown("## Soporte - Negocios cautivos")
        subido = st.file_uploader("PDF cotizaciones", type="pdf", help="Si no subes nada se usa el archivo junto a app.py.")
    origen = subido if subido is not None else (ARCHIVO_DEFECTO if ARCHIVO_DEFECTO.exists() else None)
    if origen is None:
        st.warning(f"Sube el libro migrado o colócalo junto a app.py con el nombre {ARCHIVO_DEFECTO.name}.")
        st.stop()
    try:
        t = cargar(origen)
    except Exception as e:  # hoja o columna faltante
        st.error(f"No se pudo leer el libro: {e}")
        st.stop()

    paginas = {"Resumen": pagina_resumen, "Cotizaciones": pagina_cotizaciones, "Comparar alternativas": pagina_comparar,
               "Proveedores": pagina_proveedores, "Pendientes": pagina_pendientes, "Calidad de datos": pagina_calidad}
    nav = st.navigation([st.Page(lambda f=f: f(t), title=n, url_path=n.lower().replace(" ", "_")) for n, f in paginas.items()])
    nav.run()


main()
