import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import math
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

st.set_page_config(page_title="План визитов", layout="wide")
st.title("📊 План визитов по полигонам")

# ==============================================
# ИНИЦИАЛИЗАЦИЯ SESSION STATE
# ==============================================
if 'polygons_df' not in st.session_state:
    st.session_state.polygons_df = None
if 'points_df' not in st.session_state:
    st.session_state.points_df = None
if 'visits_df' not in st.session_state:
    st.session_state.visits_df = None
if 'result_df' not in st.session_state:
    st.session_state.result_df = None 
if 'stats_df' not in st.session_state:
    st.session_state.stats_df = None

# ==============================================
# ФУНКЦИИ
# ==============================================

def parse_wkt_polygon(wkt_string):
    """Парсит WKT строку полигона"""
    try:
        match = re.search(r'POLYGON\s*\(\s*\(([^)]+)\)\s*\)', str(wkt_string), re.IGNORECASE)
        if not match:
            return None
        coords_str = match.group(1)
        points = []
        for pair in coords_str.split(','):
            pair = pair.strip()
            if ' ' in pair:
                parts = pair.split()
                if len(parts) >= 2:
                    lon = float(parts[0])
                    lat = float(parts[1])
                    points.append((lat, lon))
        return points if len(points) >= 3 else None
    except Exception:
        return None

def point_in_polygon(lat, lon, polygon_coords):
    """Проверяет, внутри ли точка полигона"""
    if not polygon_coords or len(polygon_coords) < 3:
        return False
    inside = False
    n = len(polygon_coords)
    x, y = lat, lon
    for i in range(n):
        x1, y1 = polygon_coords[i]
        x2, y2 = polygon_coords[(i + 1) % n]
        if ((y1 > y) != (y2 > y)) and (x < (x2 - x1) * (y - y1) / (y2 - y1) + x1):
            inside = not inside
    return inside

def haversine_distance(lat1, lon1, lat2, lon2):
    """Расстояние в км между точками"""
    try:
        R = 6371
        lat1, lon1, lat2, lon2 = float(lat1), float(lon1), float(lat2), float(lon2)
        lat1_rad = lat1 * math.pi / 180
        lon1_rad = lon1 * math.pi / 180
        lat2_rad = lat2 * math.pi / 180
        lon2_rad = lon2 * math.pi / 180
        dlat = lat2_rad - lat1_rad
        dlon = lon2_rad - lon1_rad
        a = math.sin(dlat/2)**2 + math.cos(lat1_rad) * math.cos(lat2_rad) * math.sin(dlon/2)**2
        c = 2 * math.asin(min(1, math.sqrt(a)))
        return R * c
    except Exception:
        return float('inf')

def calculate_polygon_center(polygon_coords):
    """Центр полигона"""
    if not polygon_coords:
        return None, None
    lats = [p[0] for p in polygon_coords]
    lons = [p[1] for p in polygon_coords]
    return sum(lats)/len(lats), sum(lons)/len(lons)

# Словарь городов с гибкими диапазонами
CITIES = {
    "Москва": (55.75, 37.62, 0.35, 0.45),
    "Санкт-Петербург": (59.93, 30.31, 0.40, 0.70),
    "Новосибирск": (55.03, 82.92, 0.45, 0.30),
    "Екатеринбург": (56.84, 60.65, 0.30, 0.35),
    "Казань": (55.79, 49.11, 0.25, 0.30),
    "Нижний Новгород": (56.33, 44.01, 0.30, 0.40),
    "Челябинск": (55.15, 61.43, 0.25, 0.30),
    "Самара": (53.20, 50.20, 0.25, 0.40),
    "Омск": (54.98, 73.37, 0.25, 0.30),
    "Ростов-на-Дону": (47.22, 39.72, 0.25, 0.30),
    "Уфа": (54.73, 55.97, 0.25, 0.30),
    "Красноярск": (56.01, 92.87, 0.30, 0.40),
    "Пермь": (58.01, 56.25, 0.35, 0.25),
    "Воронеж": (51.67, 39.21, 0.25, 0.25),
    "Волгоград": (48.71, 44.51, 0.45, 0.45),
    "Краснодар": (45.04, 38.98, 0.20, 0.20),
}

def get_city_by_coords(lat, lon):
    """Определяет город по координатам"""
    if lat is None or lon is None:
        return "Другой"
    
    for city, (center_lat, center_lon, radius_lat, radius_lon) in CITIES.items():
        # Проверяем попадание в эллипс
        if abs(lat - center_lat) <= radius_lat and abs(lon - center_lon) <= radius_lon:
            return city
    
    return "Другой"

def load_polygons_from_csv(file):
    """Загрузка полигонов из CSV (автоопределение кодировки)"""
    try:
        # Пробуем разные кодировки
        encodings = ['utf-8', 'cp1251', 'windows-1251', 'latin1']
        
        df = None
        for encoding in encodings:
            try:
                file.seek(0)  # Возвращаемся в начало файла
                df = pd.read_csv(file, encoding=encoding)
                break
            except UnicodeDecodeError:
                continue
        
        if df is None:
            st.error("❌ Не удалось прочитать файл. Проверьте кодировку (должна быть UTF-8 или Windows-1251)")
            return None
        
        # Ищем колонку с WKT
        wkt_col = None
        for col in df.columns:
            if 'wkt' in col.lower() or 'polygon' in col.lower():
                wkt_col = col
                break
        
        if wkt_col is None:
            st.error("❌ Не найдена колонка с WKT")
            return None
        
        polygons = []
        for idx, row in df.iterrows():
            coords = parse_wkt_polygon(str(row[wkt_col]))
            if coords:
                center_lat, center_lon = calculate_polygon_center(coords)
                city = get_city_by_coords(center_lat, center_lon)
                polygons.append({
                    'id': idx,
                    'coordinates': coords,
                    'center_lat': center_lat,
                    'center_lon': center_lon,
                    'city': city,
                    'name_original': row.iloc[1] if len(row) > 1 else f"Полигон_{idx+1}"
                })
        
        if not polygons:
            st.error("❌ Не удалось распарсить полигоны")
            return None
        
        # Нумеруем полигоны по городам
        city_counts = {}
        for p in polygons:
            city_counts[p['city']] = city_counts.get(p['city'], 0) + 1
            p['number'] = city_counts[p['city']]
            p['name'] = f"{p['city']} {p['number']}"
        
        result_df = pd.DataFrame(polygons)
        st.success(f"✅ Загружено {len(polygons)} полигонов")
        
        poly_names = [p['name'] for p in polygons]
        st.info(f"**Полигоны:** {', '.join(poly_names)}")
        
        return result_df
    except Exception as e:
        st.error(f"❌ Ошибка: {str(e)}")
        return None

def assign_points_to_polygons(points_df, polygons_df):
    """Привязка точек к полигонам"""
    if points_df is None or points_df.empty:
        return points_df
    if polygons_df is None or polygons_df.empty:
        points_df['Полигон'] = 'Нет полигонов'
        return points_df
    
    results = []
    for _, point in points_df.iterrows():
        lat = point['Широта']
        lon = point['Долгота']
        
        assigned = None
        min_dist = float('inf')
        nearest = None
        
        for _, poly in polygons_df.iterrows():
            if point_in_polygon(lat, lon, poly['coordinates']):
                assigned = poly['name']
                break
            if poly['center_lat']:
                dist = haversine_distance(lat, lon, poly['center_lat'], poly['center_lon'])
                if dist < min_dist:
                    min_dist = dist
                    nearest = poly['name']
        
        if assigned:
            results.append(assigned)
        elif nearest:
            results.append(f"{nearest} (ближайший)")
        else:
            results.append('Не определен')
    
    points_df = points_df.copy()
    points_df['Полигон'] = results
    
    assigned_cnt = sum(1 for p in results if 'ближайший' not in str(p) and p != 'Не определен')
    nearest_cnt = sum(1 for p in results if 'ближайший' in str(p))
    
    st.info(f"📌 Внутри полигонов: {assigned_cnt}, по расстоянию: {nearest_cnt}")
    return points_df
    
def export_to_kml(polygons_df, points_df):
    """Экспортирует полигоны и точки в один KML файл с цветовой кодировкой"""
    try:
        POLYGON_COLORS = {
            1: {"fill": "3366cc", "line": "1a33cc", "fill_hex": "#3366cc", "name": "Синий"},
            2: {"fill": "ff9933", "line": "cc7a00", "fill_hex": "#ff9933", "name": "Оранжевый"},
            3: {"fill": "33cc33", "line": "28a028", "fill_hex": "#33cc33", "name": "Зеленый"},
            4: {"fill": "9933cc", "line": "7a28a0", "fill_hex": "#9933cc", "name": "Фиолетовый"},
            5: {"fill": "ffcc00", "line": "cca300", "fill_hex": "#ffcc00", "name": "Желтый"},
            6: {"fill": "ff3333", "line": "cc2929", "fill_hex": "#ff3333", "name": "Красный"},
            7: {"fill": "33cccc", "line": "29a3a3", "fill_hex": "#33cccc", "name": "Голубой"},
            8: {"fill": "cc9966", "line": "a37a4d", "fill_hex": "#cc9966", "name": "Коричневый"},
        }
        
        kml_header = '''<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
<Document>
<name>Полигоны и точки</name>
'''
        
        kml_styles = ""
        for num, colors in POLYGON_COLORS.items():
            kml_styles += f'''
<Style id="polygon_{num}">
    <LineStyle><color>ff{colors["line"]}</color><width>3</width></LineStyle>
    <PolyStyle><color>66{colors["fill"]}</color><fill>1</fill><outline>1</outline></PolyStyle>
</Style>
<Style id="point_fact0_{num}">
    <IconStyle>
        <color>cc{colors["fill"]}</color>
        <scale>1.0</scale>
        <Icon>
            <href>http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png</href>
        </Icon>
    </IconStyle>
</Style>
<Style id="point_fact1_{num}">
    <IconStyle>
        <color>33{colors["fill"]}</color>
        <scale>1.0</scale>
        <Icon>
            <href>http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png</href>
        </Icon>
    </IconStyle>
</Style>
'''     
        kml_body = ""
        
        # Полигоны
        for idx, row in polygons_df.iterrows():
            coords = row['coordinates']
            if not coords or len(coords) < 3:
                continue
            
            import re
            match = re.search(r'(\d+)$', row['name'])
            poly_number = int(match.group(1)) if match else 1
            color_num = ((poly_number - 1) % 8) + 1
            
            coord_string = ""
            for lat, lon in coords:
                coord_string += f"{lon},{lat},0\n"
            first_lat, first_lon = coords[0]
            coord_string += f"{first_lon},{first_lat},0"
            
            kml_body += f'''
<Placemark>
    <name>{row['name']}</name>
    <description>Город: {row['city']}</description>
    <styleUrl>#polygon_{color_num}</styleUrl>
    <Polygon><outerBoundaryIs><LinearRing><coordinates>{coord_string}</coordinates></LinearRing></outerBoundaryIs></Polygon>
</Placemark>
'''
        
        # Точки
        for _, point in points_df.iterrows():
            lat = point['Широта']
            lon = point['Долгота']
            fact = point['Факт']
            polygon = str(point['Полигон'])
            
            import re
            match = re.search(r'(\d+)', polygon)
            if match:
                color_num = ((int(match.group(1)) - 1) % 8) + 1
            else:
                color_num = 1
            
            style = f"point_fact1_{color_num}" if fact == 1 else f"point_fact0_{color_num}"
            status = "✅ Посещено" if fact == 1 else "❌ Не посещено"
            clean_polygon = polygon.replace(' (ближайший)', '')
            
            kml_body += f'''
<Placemark>
    <name>{point.get('Название_Точки', point['ID_Точки'])}</name>
    <description>ID: {point['ID_Точки']}
Адрес: {point.get('Адрес', '')}
Полигон: {clean_polygon}
Статус: {status}</description>
    <styleUrl>#{style}</styleUrl>
    <Point><coordinates>{lon},{lat},0</coordinates></Point>
</Placemark>
'''
        
        kml_footer = '</Document>\n</kml>'
        return kml_header + kml_styles + kml_body + kml_footer
    except Exception as e:
        st.error(f"❌ Ошибка создания KML: {str(e)}")
        return None

# ==============================================
# ИНТЕРФЕЙС
# ==============================================

st.markdown("---")

col1, col2, col3 = st.columns(3)

with col1:
    st.subheader("1️⃣ Полигоны (CSV)")
    polygons_file = st.file_uploader("Загрузить CSV", type=['csv'], key="polygons")
    if polygons_file:
        st.session_state.polygons_df = load_polygons_from_csv(polygons_file)

with col2:
    st.subheader("2️⃣ Точки (Excel)")
    points_file = st.file_uploader("Загрузить Excel", type=['xlsx', 'xls'], key="points")
    if points_file:
        try:
            df = pd.read_excel(points_file)
            if all(c in df.columns for c in ['ID_Точки', 'Широта', 'Долгота']):
                for col in ['Название_Точки', 'Адрес', 'Тип', 'Город']:
                    if col not in df.columns:
                        df[col] = ''
                df['Широта'] = pd.to_numeric(df['Широта'], errors='coerce')
                df['Долгота'] = pd.to_numeric(df['Долгота'], errors='coerce')
                df = df.dropna(subset=['Широта', 'Долгота'])
                st.session_state.points_df = df
                st.success(f"✅ {len(df)} точек")
            else:
                st.error("❌ Нужны колонки: ID_Точки, Широта, Долгота")
        except Exception as e:
            st.error(f"❌ {str(e)}")

with col3:
    st.subheader("3️⃣ Факт визитов (Excel)")
    visits_file = st.file_uploader("Загрузить (опционально)", type=['xlsx', 'xls'], key="visits")
    if visits_file:
        try:
            df = pd.read_excel(visits_file)
            id_col = next((c for c in df.columns if 'id' in c.lower() and ('филиал' in c.lower() or 'точк' in c.lower())), None)
            if id_col:
                df['ID_Точки'] = df[id_col].astype(str)
                st.session_state.visits_df = df
                st.success(f"✅ {len(df)} записей")
            else:
                st.warning("⚠️ Не найдена колонка с ID")
        except Exception as e:
            st.error(f"❌ {str(e)}")

# ==============================================
# РАСЧЕТ
# ==============================================

st.markdown("---")

if st.button("🚀 Рассчитать план", type="primary", use_container_width=True):
    if st.session_state.points_df is None:
        st.error("❌ Загрузите точки")
        st.stop()
    if st.session_state.polygons_df is None:
        st.error("❌ Загрузите полигоны")
        st.stop()
    
    with st.spinner("🔄 Расчет..."):
        # Привязка к полигонам
        result = assign_points_to_polygons(
            st.session_state.points_df.copy(),
            st.session_state.polygons_df
        )
        
        # Расчет факта
        fact_dict = {}
        if st.session_state.visits_df is not None:
            for _, row in st.session_state.visits_df.iterrows():
                fact_dict[str(row['ID_Точки']).strip()] = 1
        
        result['Факт'] = result['ID_Точки'].astype(str).str.strip().map(fact_dict).fillna(0).astype(int)
        
        # Итоговый DataFrame
        result_final = result[['ID_Точки', 'Название_Точки', 'Адрес', 'Широта', 'Долгота', 'Город', 'Тип', 'Полигон', 'Факт']]
        st.session_state.result_df = result_final
        
        st.success("✅ Расчет завершен!")
        
        # Сохраняем статистику (без отображения)
        clean_poly = result_final['Полигон'].str.replace(r'\s*\(ближайший\)', '', regex=True)
        stats = result_final.groupby(clean_poly).agg(
            План=('ID_Точки', 'count'),
            Факт=('Факт', 'sum')
        ).reset_index()
        stats.columns = ['Полигон', 'План', 'Факт']
        stats['Разница'] = stats['План'] - stats['Факт']
        stats['Выполнение %'] = (stats['Факт'] / stats['План'] * 100).round(1).fillna(0).astype(str) + '%'
        
        st.session_state.stats_df = stats

# ==============================================
# ВЫГРУЗКА
# ==============================================

if st.session_state.result_df is not None:
    st.markdown("---")
    st.subheader("📤 Выгрузка")
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Выгрузка Excel с точками
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            st.session_state.result_df.to_excel(writer, sheet_name='План_посещений', index=False)
        
        st.download_button(
            label="📥 Скачать план посещений (Excel)",
            data=output.getvalue(),
            file_name=f"plan_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    with col2:
        # Выгрузка KML (полигоны + точки)
        if st.session_state.polygons_df is not None and not st.session_state.polygons_df.empty:
            if st.session_state.result_df is not None and not st.session_state.result_df.empty:
                kml_data = export_to_kml(st.session_state.polygons_df, st.session_state.result_df)
                if kml_data:
                    st.download_button(
                        label="🗺️ Скачать карту (KML)",
                        data=kml_data.encode('utf-8'),
                        file_name=f"map_{datetime.now().strftime('%Y%m%d_%H%M')}.kml",
                        mime="application/vnd.google-earth.kml+xml",
                        use_container_width=True
                    )
                else:
                    st.error("❌ Ошибка создания KML")
            else:
                st.info("📌 Сначала рассчитайте план")
        else:
            st.info("📌 Загрузите полигоны для экспорта")
    
    # Статистика
    st.markdown("---")
    st.subheader("📊 Статистика по полигонам")
    
    if st.session_state.stats_df is not None:
        st.dataframe(st.session_state.stats_df, use_container_width=True, hide_index=True)
    else:
        st.info("Нажмите 'Рассчитать план' для отображения статистики")
    
    with st.expander("🔍 Предпросмотр"):
        st.dataframe(st.session_state.result_df.head(20), use_container_width=True)

# ==============================================
# ИНСТРУКЦИЯ
# ==============================================

with st.expander("📖 Инструкция"):
    st.markdown("""
    1. **Полигоны (CSV)** — колонка WKT с форматом `POLYGON ((lon lat, ...))`
    2. **Точки (Excel)** — колонки `ID_Точки`, `Широта`, `Долгота`
    3. **Факт (Excel, опционально)** — колонка с ID филиала
    4. Нажмите **Рассчитать план**
    5. Скачайте результат
    """)
