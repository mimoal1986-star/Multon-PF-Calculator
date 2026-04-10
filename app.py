import streamlit as st
import pandas as pd
import numpy as np
import io
import re
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

# ==============================================
# ФУНКЦИИ ДЛЯ РАБОТЫ С ПОЛИГОНАМИ
# ==============================================

def parse_wkt_polygon(wkt_string):
    """Парсит WKT строку полигона и возвращает список координат [(lat, lon), ...]"""
    try:
        # Извлекаем координаты из POLYGON ((x1 y1, x2 y2, ...))
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
    """Проверяет, находится ли точка внутри полигона"""
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
    """Расчет расстояния между двумя точками в километрах"""
    try:
        R = 6371
        lat1, lon1, lat2, lon2 = map(float, [lat1, lon1, lat2, lon2])
        
        lat1_rad = np.radians(lat1)
        lon1_rad = np.radians(lon1)
        lat2_rad = np.radians(lat2)
        lon2_rad = np.radians(lon2)
        
        dlat = lat2_rad - lat1_rad
        dlon = lon2_rad - lon1_rad
        
        a = np.sin(dlat/2)**2 + np.cos(lat1_rad) * np.cos(lat2_rad) * np.sin(dlon/2)**2
        c = 2 * np.arcsin(np.sqrt(a))
        
        return R * c
    except Exception:
        return float('inf')

def calculate_polygon_center(polygon_coords):
    """Вычисляет центр полигона"""
    if not polygon_coords:
        return None, None
    lats = [p[0] for p in polygon_coords]
    lons = [p[1] for p in polygon_coords]
    return sum(lats)/len(lats), sum(lons)/len(lons)

def get_city_by_coords(lat, lon):
    """Определяет город по координатам"""
    if lat is None or lon is None:
        return "Другой"
    
    # Москва
    if 55.5 <= lat <= 56.0 and 37.3 <= lon <= 38.0:
        return "Москва"
    # Санкт-Петербург
    if 59.8 <= lat <= 60.1 and 30.2 <= lon <= 30.5:
        return "Санкт-Петербург"
    # Новосибирск
    if 54.9 <= lat <= 55.2 and 82.9 <= lon <= 83.1:
        return "Новосибирск"
    # Екатеринбург
    if 56.8 <= lat <= 56.9 and 60.5 <= lon <= 60.7:
        return "Екатеринбург"
    
    return "Другой"

def load_polygons_from_csv(file):
    """Загружает полигоны из CSV файла"""
    try:
        # Пробуем разные кодировки
        for encoding in ['utf-8', 'cp1251', 'latin1']:
            try:
                file.seek(0)
                df = pd.read_csv(file, encoding=encoding)
                break
            except UnicodeDecodeError:
                continue
        else:
            st.error("❌ Не удалось прочитать файл. Проверьте кодировку (должна быть UTF-8)")
            return None
        
        # Ищем колонку WKT
        wkt_col = None
        for col in df.columns:
            if 'wkt' in col.lower() or 'polygon' in col.lower():
                wkt_col = col
                break
        
        if wkt_col is None:
            st.error("❌ В файле не найдена колонка с WKT полигонами")
            return None
        
        # Парсим полигоны
        polygons = []
        
        for idx, row in df.iterrows():
            wkt = str(row[wkt_col])
            coords = parse_wkt_polygon(wkt)
            
            if coords and len(coords) >= 3:
                center_lat, center_lon = calculate_polygon_center(coords)
                city = get_city_by_coords(center_lat, center_lon)
                
                polygons.append({
                    'id': idx,
                    'name_original': row.iloc[1] if len(row) > 1 else f"Полигон_{idx+1}",
                    'wkt': wkt,
                    'coordinates': coords,
                    'center_lat': center_lat,
                    'center_lon': center_lon,
                    'city': city
                })
        
        if not polygons:
            st.error("❌ Не удалось распарсить ни одного полигона")
            return None
        
        # Группируем по городам и присваиваем номера
        city_counts = {}
        for poly in polygons:
            city = poly['city']
            city_counts[city] = city_counts.get(city, 0) + 1
            poly['number'] = city_counts[city]
            poly['name'] = f"{city} {city_counts[city]}"
        
        result_df = pd.DataFrame(polygons)
        
        st.success(f"✅ Загружено {len(polygons)} полигонов")
        
        # Показываем список
        poly_list = [f"{p['name']}" for p in polygons]
        st.info(f"**Полигоны:** {', '.join(poly_list)}")
        
        return result_df
        
    except Exception as e:
        st.error(f"❌ Ошибка загрузки полигонов: {str(e)}")
        return None

def assign_points_to_polygons(points_df, polygons_df):
    """Привязывает точки к полигонам"""
    if points_df is None or points_df.empty:
        return points_df
    
    if polygons_df is None or polygons_df.empty:
        points_df['Полигон'] = 'Нет полигонов'
        return points_df
    
    results = []
    
    for idx, point in points_df.iterrows():
        lat = point['Широта']
        lon = point['Долгота']
        
        assigned_polygon = None
        min_distance = float('inf')
        nearest_polygon = None
        
        for _, poly in polygons_df.iterrows():
            coords = poly['coordinates']
            if point_in_polygon(lat, lon, coords):
                assigned_polygon = poly['name']
                break
            
            # Для ближайшего
            if poly['center_lat'] and poly['center_lon']:
                dist = haversine_distance(lat, lon, poly['center_lat'], poly['center_lon'])
                if dist < min_distance:
                    min_distance = dist
                    nearest_polygon = poly['name']
        
        if assigned_polygon:
            results.append(assigned_polygon)
        elif nearest_polygon:
            results.append(f"{nearest_polygon} (ближайший)")
        else:
            results.append('Не определен')
    
    points_df = points_df.copy()
    points_df['Полигон'] = results
    
    assigned = sum([p for p in results if 'ближайший' not in str(p) and p != 'Не определен'])
    nearest = sum(['ближайший' in str(p) for p in results])
    unassigned = sum([p == 'Не определен' for p in results])
    
    st.info(f"📌 Распределение: {assigned} точек внутри полигонов, {nearest} по расстоянию, {unassigned} не определены")
    
    return points_df

# ==============================================
# ЗАГРУЗКА ФАЙЛОВ
# ==============================================

st.markdown("---")

col1, col2, col3 = st.columns(3)

with col1:
    st.subheader("1️⃣ Полигоны (CSV)")
    st.caption("WKT, название, описание")
    polygons_file = st.file_uploader(
        "Загрузить CSV с полигонами",
        type=['csv'],
        key="polygons_uploader"
    )
    
    if polygons_file:
        st.session_state.polygons_df = load_polygons_from_csv(polygons_file)

with col2:
    st.subheader("2️⃣ Точки (Excel)")
    st.caption("ID_Точки, Широта, Долгота, Город, Тип")
    points_file = st.file_uploader(
        "Загрузить Excel с точками",
        type=['xlsx', 'xls'],
        key="points_uploader"
    )
    
    if points_file:
        try:
            points_df = pd.read_excel(points_file)
            
            required = ['ID_Точки', 'Широта', 'Долгота']
            missing = [col for col in required if col not in points_df.columns]
            
            if missing:
                st.error(f"❌ Отсутствуют колонки: {missing}")
            else:
                if 'Название_Точки' not in points_df.columns:
                    points_df['Название_Точки'] = points_df['ID_Точки'].astype(str)
                if 'Адрес' not in points_df.columns:
                    points_df['Адрес'] = ''
                if 'Тип' not in points_df.columns:
                    points_df['Тип'] = 'Неизвестно'
                if 'Город' not in points_df.columns:
                    points_df['Город'] = 'Неизвестно'
                
                # Конвертируем координаты в float
                points_df['Широта'] = pd.to_numeric(points_df['Широта'], errors='coerce')
                points_df['Долгота'] = pd.to_numeric(points_df['Долгота'], errors='coerce')
                
                # Удаляем точки с некорректными координатами
                before = len(points_df)
                points_df = points_df.dropna(subset=['Широта', 'Долгота'])
                after = len(points_df)
                
                if before > after:
                    st.warning(f"⚠️ Удалено {before - after} точек с некорректными координатами")
                
                st.session_state.points_df = points_df
                st.success(f"✅ Загружено {len(points_df)} точек")
        except Exception as e:
            st.error(f"❌ Ошибка: {str(e)}")

with col3:
    st.subheader("3️⃣ Факт визитов (Excel)")
    st.caption("Необязательно. ID филиала = ID_Точки")
    visits_file = st.file_uploader(
        "Загрузить Excel с фактом",
        type=['xlsx', 'xls'],
        key="visits_uploader"
    )
    
    if visits_file:
        try:
            visits_df = pd.read_excel(visits_file)
            
            # Ищем колонку с ID
            id_col = None
            for col in visits_df.columns:
                col_lower = col.lower()
                if 'id филиала' in col_lower or 'id_филиала' in col_lower or 'id точки' in col_lower:
                    id_col = col
                    break
            
            # Ищем колонку с датой
            date_col = None
            for col in visits_df.columns:
                col_lower = col.lower()
                if 'дата визита' in col_lower or 'дата_визита' in col_lower or 'дата' in col_lower:
                    date_col = col
                    break
            
            if id_col:
                visits_df['ID_Точки'] = visits_df[id_col].astype(str)
                
                if date_col:
                    visits_df['Дата_визита'] = pd.to_datetime(visits_df[date_col], errors='coerce')
                    valid_visits = visits_df['Дата_визита'].notna().sum()
                    st.success(f"✅ Загружено {len(visits_df)} записей, {valid_visits} с валидными датами")
                else:
                    st.warning("⚠️ Колонка с датой не найдена, все визиты будут считаться")
                    visits_df['Дата_визита'] = pd.Timestamp.now()
                
                st.session_state.visits_df = visits_df
            else:
                st.warning("⚠️ Колонка 'ID филиала' не найдена")
                st.session_state.visits_df = None
                
        except Exception as e:
            st.error(f"❌ Ошибка: {str(e)}")
            st.session_state.visits_df = None

# ==============================================
# РАСЧЕТ
# ==============================================

st.markdown("---")

if st.button("🚀 Рассчитать план", type="primary", use_container_width=True):
    
    if st.session_state.points_df is None:
        st.error("❌ Загрузите файл с точками")
        st.stop()
    
    if st.session_state.polygons_df is None:
        st.error("❌ Загрузите файл с полигонами")
        st.stop()
    
    with st.spinner("🔄 Расчет..."):
        # 1. Привязываем точки к полигонам
        points_with_polygons = assign_points_to_polygons(
            st.session_state.points_df.copy(),
            st.session_state.polygons_df
        )
        
        # 2. Рассчитываем факт
        fact_dict = {}
        if st.session_state.visits_df is not None and not st.session_state.visits_df.empty:
            visits = st.session_state.visits_df
            for _, row in visits.iterrows():
                point_id = str(row['ID_Точки']).strip()
                # Проверяем дату если есть
                if 'Дата_визита' in row and pd.notna(row['Дата_визита']):
                    fact_dict[point_id] = 1
                elif 'Дата_визита' not in row:
                    fact_dict[point_id] = 1
        
        points_with_polygons['Факт'] = points_with_polygons['ID_Точки'].astype(str).str.strip().map(fact_dict).fillna(0).astype(int)
        
        # 3. Формируем результат
        result_df = points_with_polygons[[
            'ID_Точки', 'Название_Точки', 'Адрес', 
            'Широта', 'Долгота', 'Город', 'Тип', 'Полигон', 'Факт'
        ]].copy()
        
        st.session_state.result_df = result_df
        
        st.success("✅ Расчет завершен!")
        
        # 4. Статистика по полигонам
        st.markdown("---")
        st.subheader("📊 Статистика по полигонам")
        
        # Очищаем названия полигонов от "(ближайший)"
        result_df['Полигон_чистый'] = result_df['Полигон'].str.replace(r'\s*\(ближайший\)', '', regex=True)
        
        poly_stats = result_df.groupby('Полигон_чистый').agg({
            'ID_Точки': 'count',
            'Факт': 'sum'
        }).reset_index()
        poly_stats.columns = ['Полигон', 'Кол-во точек план', 'Кол-во точек факт']
        poly_stats['План-Факт (визиты)'] = poly_stats['Кол-во точек план'] - poly_stats['Кол-во точек факт']
        poly_stats['План-Факт (%)'] = (poly_stats['Кол-во точек факт'] / poly_stats['Кол-во точек план'] * 100).round(1)
        poly_stats['План-Факт (%)'] = poly_stats['План-Факт (%)'].fillna(0).astype(str) + '%'
        
        st.dataframe(poly_stats, use_container_width=True, hide_index=True)

# ==============================================
# ВЫГРУЗКА
# ==============================================

if st.session_state.result_df is not None:
    st.markdown("---")
    st.subheader("📤 Выгрузка")
    
    col1, col2 = st.columns([1, 2])
    
    with col1:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            st.session_state.result_df.to_excel(writer, sheet_name='План_посещений', index=False)
        
        st.download_button(
            label="📥 Скачать план посещений (Excel)",
            data=output.getvalue(),
            file_name=f"план_посещений_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    with col2:
        st.info(f"📋 Файл содержит {len(st.session_state.result_df)} записей")
    
    with st.expander("🔍 Предпросмотр результата", expanded=False):
        st.dataframe(st.session_state.result_df.head(20), use_container_width=True, hide_index=True)

# ==============================================
# ИНСТРУКЦИЯ
# ==============================================

with st.expander("📖 Инструкция", expanded=False):
    st.markdown("""
    ### Как пользоваться:
    
    1. **Загрузите полигоны (CSV)**  
       - Формат: колонка `WKT` с полигонами
       - Пример: `POLYGON ((37.5494566 55.726112, 36.7364683 55.4825184, ...))`
    
    2. **Загрузите точки (Excel)**  
       - Обязательные колонки: `ID_Точки`, `Широта`, `Долгота`
    
    3. **Загрузите факт визитов (Excel, опционально)**  
       - Должна быть колонка с ID филиала (совпадает с `ID_Точки`)
    
    4. **Нажмите "Рассчитать план"**
    
    5. **Скачайте результат**
    """)