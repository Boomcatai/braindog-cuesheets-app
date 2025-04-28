import streamlit as st
import pandas as pd
from pathlib import Path
import tempfile
import re
import plotly.express as px
import math

# Importamos funciones y constantes del extractor existente
try:
    from extractor_mejorado import (
        extract_from_excel,
        calcular_estadisticas,
        PUBLISHER_RHAPSODY,
        REGEX_EPISODIO,
        _procesar_markdown_sheet,
        formatear_tiempo,
        PDFReport,
        generar_reporte_global_pdf,
        generar_grafica_publishers,
        generar_grafica_compositores,
        generar_grafica_pistas_top_tiempo,
        generar_grafica_episodios
    )
except ImportError as imp_err:
    st.error("No se encontró extractor_mejorado.py. Asegúrate de que está en este directorio.")
    st.exception(imp_err)
    st.stop()

# Configuración de la página
st.set_page_config(
    page_title="Braindog Cuesheets Web",
    layout="wide",
    initial_sidebar_state="expanded"
)

def main():
    # Inicializar contador para key del uploader
    if "uploader_key" not in st.session_state:
        st.session_state["uploader_key"] = 0
    # Personalización del Dashboard
    dashboard_title = st.sidebar.text_input(
        "Título del Dashboard:",
        value="Braindog Cuesheets – Dashboard",
        key="dashboard_title"
    )
    # Botón para reiniciar la carga de archivos
    if st.sidebar.button("Reiniciar carga", key="reset_btn"):
        st.session_state["uploader_key"] += 1
        st.sidebar.success("Carga reiniciada. Ahora puedes subir nuevos archivos.")
        
    
    # Sidebar: carga de archivos("Configuración")
    st.title(dashboard_title)
    uploaded = st.sidebar.file_uploader(
        "Sube archivos .xlsx o .md con cuesheets",
        type=["xlsx", "md"],
        accept_multiple_files=True,
        key=f"file_uploader_{st.session_state['uploader_key']}"
    )
    if not uploaded:
        st.sidebar.info("Sube al menos un archivo para iniciar.")
        st.write("Carga tus archivos desde la barra lateral para comenzar.")
        return

    all_data = []
    total = len(uploaded)
    progress = st.sidebar.progress(0)

    # Procesar archivos
    with tempfile.TemporaryDirectory() as tmpdir:
        for i, up in enumerate(uploaded):
            tmp_path = Path(tmpdir) / up.name
            tmp_path.write_bytes(up.getbuffer())
            ext = tmp_path.suffix.lower()
            try:
                if ext == '.xlsx':
                    datos = extract_from_excel(tmp_path)
                    ep_match = re.search(REGEX_EPISODIO, tmp_path.name, re.IGNORECASE)
                    ep = ep_match.group(1).zfill(3) if ep_match else tmp_path.stem
                    for d in datos:
                        d['episode'] = ep
                else:
                    ep_match = re.search(REGEX_EPISODIO, tmp_path.name, re.IGNORECASE)
                    ep = ep_match.group(1).zfill(3) if ep_match else tmp_path.stem
                    datos = _procesar_markdown_sheet(tmp_path, ep)
                all_data.extend(datos)
            except Exception as e:
                st.sidebar.warning(f"Error procesando {up.name}: {e}")
            progress.progress(int((i+1)/total*100))

    if not all_data:
        st.error("No se extrajeron datos válidos. Revisa el formato de tus archivos.")
        return

    # Calcular estadísticas globales
    stats = calcular_estadisticas(all_data)

    # Calcular métricas comunes (necesarias tanto para dashboard como para PDF)
    total_episodes = len(stats['episodios'])
    total_tracks = stats['total_pistas']
    unique_tracks = stats['unique_tracks_count']
    total_seconds = stats.get('duracion_total_segundos', 0)
    minutes_total = math.ceil(total_seconds / 60)
    avg_secs = total_seconds // max(total_tracks, 1)
    avg_time = formatear_tiempo(avg_secs)
    unique_composers = len(stats['compositores'])
    unique_publishers = len(stats['publishers'])
    total_pub_seconds = sum(stats['publishers_tiempo'].values())
    rhapsody_secs = sum(
        d['duration_seconds'] for d in all_data
        if PUBLISHER_RHAPSODY.lower() in d.get('publisher','').lower()
    )
    rhapsody_min = math.ceil(rhapsody_secs / 60)
    rhapsody_pct = f"{(rhapsody_secs/total_pub_seconds*100 if total_pub_seconds else 0):.1f}%"

    # Preparar generación de PDF idéntico al dashboard
    from extractor_mejorado import PDFReport
    if st.sidebar.button("Generar Informe PDF", key="gen_pdf"):
        with tempfile.TemporaryDirectory() as pdf_tmp:
            # Estadísticas por episodio
            stats_por_ep = {ep: calcular_estadisticas([d for d in all_data if d['episode']==ep]) for ep in sorted(stats['episodios'], key=lambda x: int(x) if x.isdigit() else x)}
            episodios_ordenados = list(stats_por_ep.keys())
            # Generar gráficos para PDF
            chart_paths = {
                'publishers': generar_grafica_publishers(stats, Path(pdf_tmp)),
                'composers': generar_grafica_compositores(stats, Path(pdf_tmp)),
                'tracks_time': generar_grafica_pistas_top_tiempo(stats, Path(pdf_tmp)),
                'episodes_cmp': generar_grafica_episodios(stats_por_ep, Path(pdf_tmp))
            }
            # Nombre seguro
            safe_name = re.sub(r'[^A-Za-z0-9_-]+','_',dashboard_title).strip('_')
            pdf_path = Path(pdf_tmp)/f"{safe_name}_Global.pdf"
            # Crear PDF idéntico al dashboard
            pdf = PDFReport(report_name=dashboard_title)
            pdf.chart_paths = chart_paths
            pdf.add_page()
            # Título
            pdf.chapter_title(dashboard_title, level=1)
            # Resumen General: métricas
            resumen = [
                ["Total episodios", str(total_episodes)],
                ["Total pistas (usos)", str(total_tracks)],
                ["Pistas únicas", str(unique_tracks)],
                ["Tiempo total música (min)", str(minutes_total)],
                ["Duración promedio / pista", avg_time],
                ["Compositores únicos", str(unique_composers)],
                ["Editoras únicas", str(unique_publishers)],
                [f"Tiempo total {PUBLISHER_RHAPSODY} (min)", str(rhapsody_min)],
                [f"% Tiempo {PUBLISHER_RHAPSODY}", rhapsody_pct]
            ]
            pdf.add_resumen_general(resumen, title="Resumen Ejecutivo")
            # Añadir gráficos en orden
            pdf.add_chart('publishers', 'Distribución Editora')
            pdf.add_chart('composers', 'Top Compositores')
            pdf.add_chart('tracks_time', 'Top Pistas (Minutos)')
            pdf.add_chart('episodes_cmp', 'Comparativa por Episodio')
            # Guardar PDF
            pdf.output(str(pdf_path))
            # Descargar
            with open(pdf_path,'rb') as f: data=f.read()
            st.sidebar.download_button(
                "Descargar Informe PDF", data,
                file_name=pdf_path.name,
                mime="application/pdf"
            )

    # Métricas principales
    total_episodes = len(stats['episodios'])
    total_tracks = stats['total_pistas']
    unique_tracks = stats['unique_tracks_count']
    total_seconds = stats.get('duracion_total_segundos', 0)
    minutes_total = math.ceil(total_seconds / 60)
    avg_secs = total_seconds // max(total_tracks, 1)
    avg_time = formatear_tiempo(avg_secs)
    unique_composers = len(stats['compositores'])
    unique_publishers = len(stats['publishers'])
    total_pub_seconds = sum(stats['publishers_tiempo'].values())
    rhapsody_secs = sum(
        d['duration_seconds'] for d in all_data
        if PUBLISHER_RHAPSODY.lower() in d.get('publisher','').lower()
    )
    rhapsody_min = math.ceil(rhapsody_secs / 60)
    rhapsody_pct = f"{(rhapsody_secs/total_pub_seconds*100 if total_pub_seconds else 0):.1f}%"

    # Mostrar métricas
    cols = st.columns(9)
    cols[0].metric("Total episodios", total_episodes)
    cols[1].metric("Total pistas (usos)", total_tracks)
    cols[2].metric("Pistas únicas", unique_tracks)
    cols[3].metric("Tiempo total música (min)", minutes_total)
    cols[4].metric("Duración promedio / pista", avg_time)
    cols[5].metric("Compositores únicos", unique_composers)
    cols[6].metric("Editoras únicas", unique_publishers)
    cols[7].metric(f"Tiempo total {PUBLISHER_RHAPSODY} (min)", rhapsody_min)
    cols[8].metric(f"% Tiempo {PUBLISHER_RHAPSODY}", rhapsody_pct)

    # Función para paleta de colores
    def make_cmap(keys):
        non_rh = [k for k in keys if PUBLISHER_RHAPSODY.lower() not in k.lower()]
        cmap = {}
        n = len(non_rh)
        for i, p in enumerate(non_rh):
            val = int(200 - (i * (100 / max(n - 1, 1))))
            cmap[p] = f"#{val:02x}{val:02x}{val:02x}"
        cmap[PUBLISHER_RHAPSODY] = '#FF5A78'
        cmap['Otros'] = '#CCCCCC'
        return cmap

    tabs = st.tabs(["Global", "Por Episodio"])

    # Vista Global
    with tabs[0]:
        st.header("Gráficos Globales")
        # Pie Publishers
        df_pub = pd.DataFrame(stats['publishers_tiempo'].items(), columns=['Publisher', 'Seconds'])
        df_pub = df_pub[df_pub['Publisher'] != 'N/A']
        df_pub['Pct'] = df_pub['Seconds'] / df_pub['Seconds'].sum() * 100
        major = df_pub[df_pub['Pct'] >= 5].copy()
        others = df_pub[df_pub['Pct'] < 5]
        if not others.empty:
            agg = others.agg({'Seconds': 'sum', 'Pct': 'sum'})
            major = pd.concat([
                major,
                pd.DataFrame([{'Publisher': 'Otros', 'Seconds': agg['Seconds'], 'Pct': agg['Pct']}])
            ], ignore_index=True)
        major['Minutes'] = major['Seconds'].apply(lambda x: math.ceil(x / 60))
        cmap_pub = make_cmap(major['Publisher'])
        fig1 = px.pie(
            major,
            values='Minutes', names='Publisher',
            title='Distribución Editora (% min)',
            color='Publisher', color_discrete_map=cmap_pub
        )
        fig1.update_traces(
            textinfo='percent+label',
            hovertemplate='%{label}: %{value:.0f} min (%{percent:.0%})'
        )
        st.plotly_chart(fig1, use_container_width=True)

        # Top Compositores
        comp_pubs = {}
        for d in all_data:
            for comp in [c.strip() for c in d.get('composer', '').split(' / ')]:
                comp_pubs.setdefault(comp, set()).add(d.get('publisher', ''))
        df_comp = pd.DataFrame(stats['compositores_tiempo'].items(), columns=['Composer', 'Seconds'])
        df_comp = df_comp[df_comp['Composer'] != 'N/A']
        df_comp['Minutes'] = df_comp['Seconds'].apply(lambda x: math.ceil(x / 60))
        df_comp['Uses'] = df_comp['Composer'].map(stats['compositores'])
        df_comp = df_comp.sort_values('Minutes', ascending=False).head(15)
        composer_cmap = {
            comp: ('#FF5A78' if any(
                PUBLISHER_RHAPSODY.lower() in pub.lower() for pub in comp_pubs.get(comp, [])
            ) else '#888888')
            for comp in df_comp['Composer']
        }
        fig2 = px.bar(
            df_comp,
            x='Minutes', y='Composer',
            orientation='h', title='Top Compositores (Minutos)',
            hover_data=['Uses'], text='Uses',
            color='Composer', color_discrete_map=composer_cmap
        )
        fig2.update_layout(
            yaxis={'categoryorder': 'array', 'categoryarray': df_comp['Composer'].tolist()}
        )
        fig2.update_traces(textposition='auto', showlegend=False)
        st.plotly_chart(fig2, use_container_width=True)

        # Top Pistas
        df_tr = pd.DataFrame(stats['pistas_repetidas_detalle'])
        df_tr = df_tr.sort_values('tiempo_total', ascending=False)
        df_tr['Uses'] = df_tr['count']
        df_tr['TotalMin'] = df_tr['tiempo_total'].apply(lambda x: math.ceil(x / 60))
        df_tr['AvgSec'] = df_tr['tiempo_total'] / df_tr['count']
        df_tr['AvgMMSS'] = df_tr['AvgSec'].apply(lambda x: formatear_tiempo(int(math.ceil(x))))
        df_top = df_tr.head(15)
        cmap_tr = make_cmap(df_top['publisher'])
        fig3 = px.bar(
            df_top,
            x='TotalMin', y='title', orientation='h',
            title='Top Pistas (Total Minutos)',
            hover_data=['Uses', 'AvgMMSS'], text='Uses',
            color='publisher', color_discrete_map=cmap_tr
        )
        fig3.update_layout(
            yaxis={'categoryorder': 'array', 'categoryarray': df_top['title'].tolist()}
        )
        fig3.update_traces(textposition='auto', showlegend=False)
        st.plotly_chart(fig3, use_container_width=True)

    # Vista por Episodio
    with tabs[1]:
        st.header("Temas Destacados por Episodio")
        for ep in sorted(stats['episodios'], key=lambda x: int(x) if x.isdigit() else x):
            ep_data = [d for d in all_data if d['episode'] == ep]
            if not ep_data:
                continue
            st.subheader(f"Episodio {ep}")
            ep_stats = calcular_estadisticas(ep_data)
            df_ep = pd.DataFrame(ep_stats['pistas_repetidas_detalle'])
            df_ep['Minutes'] = df_ep['tiempo_total'].apply(lambda x: math.ceil(x / 60))
            df_top = df_ep.sort_values('Minutes', ascending=False).head(10)
            df_top['Color'] = df_top['publisher'].apply(
                lambda p: '#FF5A78' if PUBLISHER_RHAPSODY.lower() in p.lower() else '#D3D3D3'
            )
            fig_ep = px.bar(
                df_top,
                x='Minutes', y='title', orientation='h',
                title=f'Top 10 Temas Duración Ep {ep}',
                hover_data=['tiempo_formateado'], text='tiempo_formateado',
                color='publisher', color_discrete_map={
                    p: ('#FF5A78' if PUBLISHER_RHAPSODY.lower() in p.lower() else '#D3D3D3')
                    for p in df_top['publisher']
                }
            )
            fig_ep.update_layout(
                yaxis={'categoryorder': 'array', 'categoryarray': df_top['title'].tolist()}
            )
            fig_ep.update_traces(textposition='auto', showlegend=False)
            st.plotly_chart(fig_ep, use_container_width=True)

if __name__ == "__main__":
    main()
