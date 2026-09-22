import sys
import os
import zipfile
import json
import re
import threading
import queue
import io
import colorsys
import pandas as pd
import xml.etree.ElementTree as ET

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QStackedWidget, QLabel, QPushButton, QTableWidget, QTableWidgetItem, 
    QHeaderView, QScrollArea, QComboBox, QDoubleSpinBox, QProgressBar, 
    QColorDialog, QFileDialog, QMessageBox, QListWidget, QListWidgetItem,
    QSplitter, QFrame, QSizePolicy
)
from PyQt6.QtCore import Qt, QTimer, QUrl, QPropertyAnimation, QEasingCurve
from PyQt6.QtGui import QColor, QFont, QDesktopServices

import folium
from folium.plugins import MarkerCluster
from PyQt6.QtWebEngineWidgets import QWebEngineView

# Theme Engine (Graceful fallback if not installed)
try:
    import qdarktheme
    HAS_QDARKTHEME = True
except ImportError:
    HAS_QDARKTHEME = False

AUTHOR_NAME = "Mohamed Ashraf"
TOOL_NAME = "SpatialPin Converter"
VERSION = "3.0.0" 

# --- Google Earth Standard Assets ---
GE_ICONS = {
    "Pushpin": "http://maps.google.com/mapfiles/kml/pushpin/wht-pushpin.png",
    "Paddle": "http://maps.google.com/mapfiles/kml/paddle/wht-blank.png",
    "Circle": "http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png",
    "Star": "http://maps.google.com/mapfiles/kml/shapes/star.png",
    "Triangle": "http://maps.google.com/mapfiles/kml/shapes/triangle.png",
    "Square": "http://maps.google.com/mapfiles/kml/shapes/polygon.png",
    "Target": "http://maps.google.com/mapfiles/kml/shapes/target.png",
}

def hex_to_kml_color(hex_color, opacity=1.0):
    if not hex_color or len(hex_color) != 7:
        return "ffffffff"
    hex_color = hex_color.lstrip("#")
    alpha_hex = f"{int(opacity * 255):02x}"
    return f"{alpha_hex}{hex_color[4:6]}{hex_color[2:4]}{hex_color[0:2]}".lower()

def parse_coordinate(coord_str):
    try:
        if pd.isna(coord_str):
            return None
        coord_str = str(coord_str).strip()
        try:
            return float(coord_str)
        except ValueError:
            pass

        match = re.match(r"(\d+)[°\s]+([\d\.]+)['\s]*([\d\.]+)?[\"\s]*([NSEWnsew])?", coord_str)
        if match:
            deg, min, sec, dir = match.groups()
            dd = float(deg) + float(min) / 60 + (float(sec) / 3600 if sec else 0)
            if dir and dir.upper() in ["S", "W"]:
                dd *= -1
            return dd
        return None
    except (ValueError, TypeError, AttributeError):
        return None

class KMLProcessor:
    def __init__(self):
        self.df = None
        self.current_filepath = None
        self.target_cols = {}
        self.unique_categories = []
        
    def _find_best_column(self, cols, aliases):
        cols_lower = [str(c).strip().lower() for c in cols]
        for alias in aliases:
            if alias in cols_lower: return cols[cols_lower.index(alias)]
        for alias in aliases:
            if len(alias) >= 2: 
                for i, c_lower in enumerate(cols_lower):
                    clean_c = c_lower.replace("_", " ").replace("-", " ")
                    if re.search(rf"\b{re.escape(alias)}\b", clean_c): return cols[i]
        for alias in aliases:
            if len(alias) >= 3:
                for i, c_lower in enumerate(cols_lower):
                    if alias in c_lower: return cols[i]
        return None

    def save_back_to_file(self):
        if self.df is None or not self.current_filepath: return
        export_df = self.df.drop(columns=["__parsed_lat", "__parsed_lon"], errors='ignore')
        ext = os.path.splitext(self.current_filepath)[1].lower()
        try:
            if ext in ['.xlsx', '.xls']:
                export_df.to_excel(self.current_filepath, index=False)
            elif ext in ['.csv', '.txt']:
                export_df.to_csv(self.current_filepath, index=False)
            elif ext in ['.geojson', '.json']:
                features = []
                for _, row in export_df.iterrows():
                    row_dict = row.dropna().to_dict()
                    lon = row_dict.pop('longitude', None)
                    lat = row_dict.pop('latitude', None)
                    feat = {
                        "type": "Feature",
                        "properties": row_dict,
                        "geometry": None
                    }
                    if pd.notna(lon) and pd.notna(lat):
                        try:
                            feat["geometry"] = {"type": "Point", "coordinates": [float(lon), float(lat)]}
                        except ValueError:
                            pass
                    features.append(feat)
                geojson = {"type": "FeatureCollection", "features": features}
                with open(self.current_filepath, 'w', encoding='utf-8') as f:
                    json.dump(geojson, f, indent=2)
        except Exception as e:
            print(f"Background save failed: {e}")

    def load_and_parse_file(self, filepath):
        try:
            self.current_filepath = filepath
            ext = os.path.splitext(filepath)[1].lower()
            
            if ext in ['.xlsx', '.xls']:
                self.df = pd.read_excel(filepath)
            elif ext in ['.csv', '.txt']:
                self.df = pd.read_csv(filepath)
            elif ext in ['.geojson', '.json']:
                with open(filepath, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                rows = []
                for feat in data.get('features', []):
                    prop = feat.get('properties', {})
                    geom = feat.get('geometry', {})
                    if geom and geom.get('type') == 'Point':
                        coords = geom.get('coordinates', [])
                        prop['longitude'] = coords[0] if len(coords) > 0 else None
                        prop['latitude'] = coords[1] if len(coords) > 1 else None
                    rows.append(prop)
                self.df = pd.DataFrame(rows)
            else:
                return False, "Unsupported format."

            raw_cols = list(self.df.columns)
            aliases = {
                "point name": ["point name", "site name", "name", "id", "site_id", "site id", "label", "point_id", "Nominal name", "Nominal"],
                "latitude": ["latitude", "lat", "Nominal lat", "y_coord", "y", "northing"],
                "longitude": ["longitude", "lon", "long", "Nominal long", "Nominal lon", "lng", "x_coord", "x", "easting"],
                "category": ["status", "category", "type", "class", "group"],
            }

            for key in ["point name", "latitude", "longitude"]:
                found_col = self._find_best_column(raw_cols, aliases[key])
                if found_col:
                    self.target_cols[key] = found_col
                else:
                    return False, f"Missing required column for '{key.title()}'.\nVariations checked: {', '.join(aliases[key])}"

            found_cat = self._find_best_column(raw_cols, aliases["category"])
            if found_cat:
                self.target_cols["category"] = found_cat
            else:
                if "Category" not in self.df.columns:
                    self.df["Category"] = "All Points"
                self.target_cols["category"] = "Category"

            self.df[self.target_cols["category"]] = self.df[self.target_cols["category"]].fillna("Uncategorized")
            self.df[self.target_cols["point name"]] = self.df[self.target_cols["point name"]].fillna("Unnamed Point")
            self.df["__parsed_lat"] = self.df[self.target_cols["latitude"]].apply(parse_coordinate)
            self.df["__parsed_lon"] = self.df[self.target_cols["longitude"]].apply(parse_coordinate)
            self.unique_categories = sorted(self.df[self.target_cols["category"]].astype(str).unique().tolist())
            
            invalid_count = self.df[['__parsed_lat', '__parsed_lon']].isna().any(axis=1).sum()
            msg = f"Loaded {len(self.df)} points.\nMapped:\n- Name: '{self.target_cols['point name']}'\n- Lat: '{self.target_cols['latitude']}'\n- Lon: '{self.target_cols['longitude']}'"
            if invalid_count > 0:
                msg += f"\n\n⚠ Found {invalid_count} rows with invalid coordinates.\nThey are highlighted in red in the grid. Please fix them inline."
            
            return True, msg
        except Exception as e:
            return False, f"Failed to load or parse file: {str(e)}"

    def generate_html_description(self, row_dict, visible_cols):
        html = '<table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse; width: 300px;">'
        html += '<tr style="background-color: #f2f2f2; color: black;"><th>Attribute</th><th>Value</th></tr>'
        for col in visible_cols:
            if col in row_dict and pd.notna(row_dict[col]):
                html += f"<tr><td style='color: black;'><b>{str(col)}</b></td><td style='color: black;'>{str(row_dict[col])}</td></tr>"
        html += "</table>"
        return html

    def export_to_html(self, save_path, style_config, visible_cols, export_queue):
        try:
            total_rows = len(self.df)
            update_interval = max(1, total_rows // 100)
            min_lat, max_lat = self.df["__parsed_lat"].min(), self.df["__parsed_lat"].max()
            min_lon, max_lon = self.df["__parsed_lon"].min(), self.df["__parsed_lon"].max()
            center_lat = (min_lat + max_lat) / 2
            center_lon = (min_lon + max_lon) / 2

            m = folium.Map(location=[center_lat, center_lon], zoom_start=4)
            marker_cluster = MarkerCluster().add_to(m)
            records = self.df.to_dict('records')
            
            for idx, row in enumerate(records):
                lat, lon = row["__parsed_lat"], row["__parsed_lon"]
                name = str(row[self.target_cols["point name"]])
                cat = str(row[self.target_cols["category"]])
                conf = style_config.get(cat, {})
                hex_color = conf.get("icolor", "#ff0000")
                opacity = conf.get("opacity", 1.0)
                
                html_table = self.generate_html_description(row, visible_cols)
                popup = folium.Popup(html_table, max_width=350)
                
                folium.CircleMarker(
                    location=[lat, lon], radius=6, color="white", weight=1,
                    fill=True, fill_color=hex_color, fill_opacity=opacity,
                    tooltip=name, popup=popup
                ).add_to(marker_cluster)

                if idx % update_interval == 0 or idx == total_rows - 1:
                    export_queue.put(("progress", int(((idx + 1) / total_rows) * 100)))

            export_queue.put(("status", "Writing HTML to disk..."))
            if not (min_lat == max_lat and min_lon == max_lon):
                m.fit_bounds([[min_lat, min_lon], [max_lat, max_lon]])
            m.save(save_path)
            export_queue.put(("done", save_path))
        except Exception as e:
            export_queue.put(("error", str(e)))

    def export_to_kml(self, save_path, style_config, visible_cols, export_queue):
        try:
            is_kmz = save_path.lower().endswith(".kmz")
            total_rows = len(self.df)
            update_interval = max(1, total_rows // 100)
            KML_NS = "http://www.opengis.net/kml/2.2"
            ET.register_namespace('', KML_NS)
            kml_root = ET.Element(f"{{{KML_NS}}}kml")
            doc_el = ET.SubElement(kml_root, f"{{{KML_NS}}}Document")
            name_el = ET.SubElement(doc_el, f"{{{KML_NS}}}name")
            name_el.text = os.path.basename(save_path)
            custom_images_to_copy = set()

            for cat, conf in style_config.items():
                s_id = f"style_{hash(cat)}"
                style_el = ET.SubElement(doc_el, f"{{{KML_NS}}}Style", id=s_id)
                icon_style = ET.SubElement(style_el, f"{{{KML_NS}}}IconStyle")
                ET.SubElement(icon_style, f"{{{KML_NS}}}color").text = hex_to_kml_color(conf["icolor"], conf["opacity"])
                ET.SubElement(icon_style, f"{{{KML_NS}}}scale").text = str(conf["iscale"])
                icon_el = ET.SubElement(icon_style, f"{{{KML_NS}}}Icon")
                
                if conf["shape"] == "Custom Image" and conf.get("custom_path"):
                    img_path = conf["custom_path"]
                    img_name = os.path.basename(img_path)
                    ET.SubElement(icon_el, f"{{{KML_NS}}}href").text = f"images/{img_name}"
                    custom_images_to_copy.add(img_path)
                else:
                    ET.SubElement(icon_el, f"{{{KML_NS}}}href").text = GE_ICONS.get(conf["shape"], GE_ICONS["Pushpin"])
                
                label_style = ET.SubElement(style_el, f"{{{KML_NS}}}LabelStyle")
                ET.SubElement(label_style, f"{{{KML_NS}}}color").text = hex_to_kml_color(conf["lcolor"])
                ET.SubElement(label_style, f"{{{KML_NS}}}scale").text = str(conf["lscale"])

            records = self.df.to_dict('records')
            for idx, row in enumerate(records):
                p_name = str(row[self.target_cols["point name"]])
                cat = str(row[self.target_cols["category"]])
                s_id = f"style_{hash(cat)}"
                lat, lon = row["__parsed_lat"], row["__parsed_lon"]

                pm_el = ET.SubElement(doc_el, f"{{{KML_NS}}}Placemark")
                ET.SubElement(pm_el, f"{{{KML_NS}}}name").text = p_name
                ET.SubElement(pm_el, f"{{{KML_NS}}}styleUrl").text = f"#{s_id}"
                desc_el = ET.SubElement(pm_el, f"{{{KML_NS}}}description")
                desc_el.text = self.generate_html_description(row, visible_cols)
                point_el = ET.SubElement(pm_el, f"{{{KML_NS}}}Point")
                ET.SubElement(point_el, f"{{{KML_NS}}}coordinates").text = f"{lon},{lat},0"

                if idx % update_interval == 0 or idx == total_rows - 1:
                    export_queue.put(("progress", int(((idx + 1) / total_rows) * 100)))

            export_queue.put(("status", "Writing file to disk..."))
            xml_bytes = ET.tostring(kml_root, encoding="utf-8", xml_declaration=True)

            if is_kmz:
                with zipfile.ZipFile(save_path, "w", zipfile.ZIP_DEFLATED) as kmz:
                    kmz.writestr("doc.kml", xml_bytes)
                    for img_path in custom_images_to_copy:
                        if os.path.exists(img_path):
                            kmz.write(img_path, f"images/{os.path.basename(img_path)}")
            else:
                with open(save_path, "wb") as f:
                    f.write(xml_bytes)
            export_queue.put(("done", save_path))
        except Exception as e:
            export_queue.put(("error", str(e)))

class PyQtVar:
    def __init__(self, getter, setter):
        self._getter = getter
        self._setter = setter
    def get(self): return self._getter()
    def set(self, val): self._setter(val)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{TOOL_NAME} v{VERSION}")
        self.resize(1100, 750)
        self.setMinimumSize(700, 500) 
        self.setAcceptDrops(True) 

        # Premium Typography Configuration
        font = QFont("Segoe UI", 9)
        font.setStyleHint(QFont.StyleHint.SansSerif)
        QApplication.setFont(font)

        # Core Backend State
        self.processor = KMLProcessor()
        self.style_mappings = {}
        self.export_queue = queue.Queue()

        self.init_ui()

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls(): event.accept()
        else: event.ignore()
            
    def dropEvent(self, event):
        filepath = event.mimeData().urls()[0].toLocalFile()
        if filepath.lower().endswith(('.xlsx', '.xls', '.csv', '.txt', '.geojson', '.json')):
            self.load_file(filepath)
        else:
            QMessageBox.warning(self, "Invalid File", "Please drop a supported format (.xlsx, .csv, .txt, .geojson)")

    def create_card_container(self):
        """Helper to create elevated card visual wrappers."""
        card = QFrame()
        card.setObjectName("Card")
        card.setStyleSheet("""
            QFrame#Card {
                background-color: rgba(128, 128, 128, 0.05);
                border-radius: 10px;
                border: 1px solid rgba(128, 128, 128, 0.2);
            }
        """)
        return card

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        self.sidebar = QFrame()
        self.sidebar.setObjectName("Sidebar")
        self.sidebar.setStyleSheet("""
            QFrame#Sidebar {
                background-color: rgba(10, 10, 10, 0.2);
                border-right: 1px solid rgba(128, 128, 128, 0.2);
            }
            QPushButton {
                text-align: left;
                padding: 12px 15px;
                border: none;
                border-radius: 8px;
                font-weight: bold;
                font-size: 14px;
                background-color: transparent;
            }
            QPushButton:hover {
                background-color: rgba(128, 128, 128, 0.1);
            }
            QPushButton:checked {
                background-color: #00A2FF; /* Electric Blue Accent */
                color: white;
            }
        """)
        
        self.sidebar.setMinimumWidth(220)
        self.sidebar.setMaximumWidth(220)
        sidebar_layout = QVBoxLayout(self.sidebar)
        sidebar_layout.setContentsMargins(10, 15, 10, 15)

        self.btn_hamburger = QPushButton("☰   Menu")
        self.btn_hamburger.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_hamburger.setStyleSheet("font-size: 18px; padding: 10px; text-align: left; font-weight: normal;")
        self.btn_hamburger.clicked.connect(self.toggle_sidebar)
        sidebar_layout.addWidget(self.btn_hamburger)
        sidebar_layout.addSpacing(20)

        self.stacked_widget = QStackedWidget()
        
        self.tab_data = QWidget()
        self.tab_styles = QWidget()
        self.tab_map = QWidget()
        self.tab_export = QWidget()
        self.tab_contact = QWidget()

        self.stacked_widget.addWidget(self.tab_data)
        self.stacked_widget.addWidget(self.tab_styles)
        self.stacked_widget.addWidget(self.tab_map)
        self.stacked_widget.addWidget(self.tab_export)
        self.stacked_widget.addWidget(self.tab_contact)

        self.nav_btns = []
        nav_items = [
            ("📊", "Data Upload", 0),
            ("🎨", "Style Mapping", 1),
            ("🗺️", "Live Map", 2),
            ("🚀", "Export Data", 3),
            ("✉️", "Contact Us", 4)
        ]

        for icon, text, index in nav_items:
            btn = QPushButton(f"{icon}   {text}")
            btn.setCheckable(True)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.clicked.connect(lambda checked, idx=index: self.switch_tab(idx))
            self.nav_btns.append(btn)
            sidebar_layout.addWidget(btn)

        sidebar_layout.addStretch()

        self.lbl_footer = QLabel(f"v{VERSION}")
        self.lbl_footer.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_footer.setStyleSheet("opacity: 0.5;")
        sidebar_layout.addWidget(self.lbl_footer)

        main_layout.addWidget(self.sidebar)
        
        content_wrapper = QWidget()
        content_layout = QVBoxLayout(content_wrapper)
        content_layout.setContentsMargins(15, 15, 15, 15)
        content_layout.addWidget(self.stacked_widget)
        main_layout.addWidget(content_wrapper)

        if self.nav_btns: self.nav_btns[0].setChecked(True)

        self._build_tab_data()
        self._build_tab_styles()
        self._build_tab_map()
        self._build_tab_export()
        self._build_tab_contact()

    def toggle_sidebar(self):
        width = self.sidebar.width()
        end_width = 65 if width > 100 else 220
        show_text = end_width == 220
        
        self.btn_hamburger.setText("☰   Menu" if show_text else "☰")
        nav_items = ["Data Upload", "Style Mapping", "Live Map", "Export Data", "Contact Us"]
        icons = ["📊", "🎨", "🗺️", "🚀", "✉️"]
        
        for i, btn in enumerate(self.nav_btns):
            btn.setText(f"{icons[i]}   {nav_items[i]}" if show_text else icons[i])
        self.lbl_footer.setVisible(show_text)

        self._anim1 = QPropertyAnimation(self.sidebar, b"minimumWidth")
        self._anim1.setDuration(300)
        self._anim1.setStartValue(width)
        self._anim1.setEndValue(end_width)
        self._anim1.setEasingCurve(QEasingCurve.Type.InOutQuart)
        
        self._anim2 = QPropertyAnimation(self.sidebar, b"maximumWidth")
        self._anim2.setDuration(300)
        self._anim2.setStartValue(width)
        self._anim2.setEndValue(end_width)
        self._anim2.setEasingCurve(QEasingCurve.Type.InOutQuart)

        self._anim1.start()
        self._anim2.start()

    def switch_tab(self, index):
        self.stacked_widget.setCurrentIndex(index)
        for i, btn in enumerate(self.nav_btns):
            btn.setChecked(i == index)

    def _build_tab_data(self):
        layout = QVBoxLayout(self.tab_data)
        layout.setContentsMargins(0,0,0,0)
        
        card = self.create_card_container()
        card_layout = QVBoxLayout(card)
        
        top_frame = QHBoxLayout()
        title = QLabel("Upload Data (Drag & Drop)")
        title.setFont(QFont("Segoe UI", 16, QFont.Weight.Bold))
        title.setStyleSheet("color: #00A2FF;")
        
        btn_browse = QPushButton("Browse File...")
        btn_browse.setStyleSheet("background-color: #00A2FF; color: white; padding: 8px 16px; border-radius: 4px; font-weight: bold;")
        btn_browse.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_browse.clicked.connect(self.load_file)
        
        top_frame.addWidget(title)
        top_frame.addStretch()
        top_frame.addWidget(btn_browse)
        card_layout.addLayout(top_frame)

        self.table = QTableWidget()
        self.table.itemChanged.connect(self.on_cell_changed)
        # Clean up table borders to match modern aesthetic
        self.table.setStyleSheet("QTableWidget { border: none; } QHeaderView::section { border: none; padding: 5px; }")
        card_layout.addWidget(self.table)
        
        layout.addWidget(card)

    def load_file(self, filepath=None):
        if not filepath or not isinstance(filepath, str):
            filepath, _ = QFileDialog.getOpenFileName(self, "Open File", "", "Supported Files (*.xlsx *.xls *.csv *.txt *.geojson *.json)")
        if not filepath: return

        success, msg = self.processor.load_and_parse_file(filepath)
        if success:
            df = self.processor.df
            display_cols = [c for c in df.columns if not c.startswith("__parsed")]
            
            self.table.blockSignals(True)
            self.table.clear()
            self.table.setColumnCount(len(display_cols))
            self.table.setHorizontalHeaderLabels(display_cols)
            self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)

            records = df.to_dict('records')
            self.table.setRowCount(len(records))
            
            for row_idx, row_dict in enumerate(records):
                is_invalid = pd.isna(row_dict["__parsed_lat"]) or pd.isna(row_dict["__parsed_lon"])
                for col_idx, col_name in enumerate(display_cols):
                    val = "" if pd.isna(row_dict[col_name]) else str(row_dict[col_name])
                    item = QTableWidgetItem(val)
                    item.setFlags(Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsEditable)
                    if is_invalid: item.setBackground(QColor("#8b0000")) # Deep red for errors
                    self.table.setItem(row_idx, col_idx, item)

            self.table.blockSignals(False)
            
            self.list_columns.clear()
            for col in display_cols:
                item = QListWidgetItem(col)
                item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                item.setCheckState(Qt.CheckState.Checked)
                self.list_columns.addItem(item)
            
            self.build_dynamic_styles()
            self.switch_tab(0) # Stay on Data tab to show errors
            
            if "⚠" in msg: QMessageBox.warning(self, "Validation Warning", msg)
            else: QMessageBox.information(self, "Success", msg)
        else:
            QMessageBox.critical(self, "Data Error", msg)
            
    def on_cell_changed(self, item):
        row, col = item.row(), item.column()
        col_name = self.table.horizontalHeaderItem(col).text()
        new_val = item.text()
        
        df = self.processor.df
        if df is None: return
        df.at[row, col_name] = new_val
        
        lat_col = self.processor.target_cols.get("latitude")
        lon_col = self.processor.target_cols.get("longitude")
        
        needs_recheck = False
        if col_name == lat_col:
            df.at[row, "__parsed_lat"] = parse_coordinate(new_val)
            needs_recheck = True
        elif col_name == lon_col:
            df.at[row, "__parsed_lon"] = parse_coordinate(new_val)
            needs_recheck = True
            
        if needs_recheck:
            is_invalid = pd.isna(df.at[row, "__parsed_lat"]) or pd.isna(df.at[row, "__parsed_lon"])
            color = QColor("#8b0000") if is_invalid else QColor(0,0,0,0) # Transparent for standard theme
            self.table.blockSignals(True)
            for c in range(self.table.columnCount()):
                it = self.table.item(row, c)
                if it: it.setBackground(color)
            self.table.blockSignals(False)
        self.processor.save_back_to_file()

    def _build_tab_styles(self):
        layout = QVBoxLayout(self.tab_styles)
        layout.setContentsMargins(0,0,0,0)
        card = self.create_card_container()
        card_layout = QVBoxLayout(card)
        
        top_frame = QHBoxLayout()
        title = QLabel("Map Styles & Labels")
        title.setFont(QFont("Segoe UI", 16, QFont.Weight.Bold))
        title.setStyleSheet("color: #00A2FF;")
        
        btn_load = QPushButton("Load Config")
        btn_load.clicked.connect(self.load_styles)
        btn_save = QPushButton("Save Config")
        btn_save.clicked.connect(self.save_styles)
        
        top_frame.addWidget(title)
        top_frame.addStretch()
        top_frame.addWidget(btn_load)
        top_frame.addWidget(btn_save)
        card_layout.addLayout(top_frame)

        self.splitter = QSplitter(Qt.Orientation.Horizontal)
        self.splitter.setStyleSheet("""
            QSplitter::handle { background: rgba(128, 128, 128, 0.3); }
            QSplitter::handle:horizontal { width: 4px; border-radius: 2px; margin: 4px; }
            QSplitter::handle:horizontal:hover { background: #00A2FF; }
        """)

        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0,0,0,0)

        palette_frame = QHBoxLayout()
        self.cb_palette = QComboBox()
        self.cb_palette.addItems(["Distinct/Rainbow", "Red-to-Green Heatmap", "Blue Gradient"])
        btn_apply_palette = QPushButton("Apply Palette")
        btn_apply_palette.setStyleSheet("background-color: #8A2BE2; color: white; padding: 4px 12px; border-radius: 4px;")
        btn_apply_palette.clicked.connect(self.apply_smart_palette)
        palette_frame.addWidget(QLabel("Smart Palette:"))
        palette_frame.addWidget(self.cb_palette)
        palette_frame.addWidget(btn_apply_palette)
        palette_frame.addStretch()
        left_layout.addLayout(palette_frame)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setStyleSheet("QScrollArea { border: none; }")
        self.scroll_widget = QWidget()
        self.scroll_layout = QVBoxLayout(self.scroll_widget)
        self.scroll_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        self.scroll_area.setWidget(self.scroll_widget)
        left_layout.addWidget(self.scroll_area)
        self.splitter.addWidget(left_widget)

        self.right_widget = QWidget()
        right_layout = QVBoxLayout(self.right_widget)
        right_layout.setContentsMargins(0,0,0,0)
        self.btn_toggle_cols = QPushButton("⛌ Hide Columns Panel")
        self.btn_toggle_cols.clicked.connect(self.toggle_columns)
        right_layout.addWidget(self.btn_toggle_cols)
        self.lbl_cols = QLabel("Balloon Columns (Visible)")
        self.lbl_cols.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        right_layout.addWidget(self.lbl_cols)
        self.list_columns = QListWidget()
        right_layout.addWidget(self.list_columns)
        
        self.splitter.addWidget(self.right_widget)
        self.splitter.setStretchFactor(0, 3)
        self.splitter.setStretchFactor(1, 1)
        card_layout.addWidget(self.splitter)
        layout.addWidget(card)

    def toggle_columns(self):
        is_visible = self.right_widget.isVisible()
        self.right_widget.setVisible(not is_visible)

    def apply_smart_palette(self):
        palette_type = self.cb_palette.currentText()
        cats = self.processor.unique_categories
        n = len(cats)
        if n == 0: return
        for i, cat in enumerate(cats):
            if palette_type == "Distinct/Rainbow": r, g, b = colorsys.hsv_to_rgb(i / n, 0.8, 0.9)
            elif palette_type == "Red-to-Green Heatmap": r, g, b = colorsys.hsv_to_rgb((i / max(1, n - 1)) * 0.33, 0.8, 0.9)
            elif palette_type == "Blue Gradient": r, g, b = colorsys.hsv_to_rgb(0.6, 0.8, 0.4 + (0.6 * (i / max(1, n - 1))))
            hex_color = f"#{int(r*255):02x}{int(g*255):02x}{int(b*255):02x}"
            if str(cat) in self.style_mappings: self.style_mappings[str(cat)]["icolor"].set(hex_color)

    def build_dynamic_styles(self):
        while self.scroll_layout.count():
            item = self.scroll_layout.takeAt(0)
            if item.widget(): item.widget().deleteLater()
        self.style_mappings.clear()

        hdr_layout = QHBoxLayout()
        for t, stretch in [("Category", 3), ("Icon", 2), ("Color", 1), ("Opacity", 1), ("Scale", 1), ("Lbl Color", 1), ("Lbl Scale", 1)]:
            lbl = QLabel(t)
            lbl.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
            lbl.setFont(QFont("Segoe UI", 9, QFont.Weight.Bold))
            hdr_layout.addWidget(lbl, stretch)
        
        hdr_widget = QWidget()
        hdr_widget.setLayout(hdr_layout)
        self.scroll_layout.addWidget(hdr_widget)

        for cat in self.processor.unique_categories:
            row_layout = QHBoxLayout()
            
            lbl_cat = QLabel(str(cat)[:25])
            lbl_cat.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
            row_layout.addWidget(lbl_cat, 3)

            cb_shape = QComboBox()
            cb_shape.addItems(list(GE_ICONS.keys()) + ["Custom Image"])
            cb_shape.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            cb_shape.setMinimumWidth(30) # Allow deep shrinking
            row_layout.addWidget(cb_shape, 2)

            btn_custom_icon = QPushButton("📂")
            btn_custom_icon.setSizePolicy(QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Fixed)
            btn_custom_icon.setVisible(False)
            row_layout.addWidget(btn_custom_icon, 0)
            
            def handle_shape_change(text, btn=btn_custom_icon): btn.setVisible(text == "Custom Image")
            def pick_custom_icon(btn):
                path, _ = QFileDialog.getOpenFileName(self, "Select Icon", "", "Images (*.png *.jpg *.jpeg *.gif)")
                if path:
                    btn.setProperty("custom_path", path)
                    btn.setStyleSheet("background-color: #00A2FF; color: white;") 
                    
            cb_shape.currentTextChanged.connect(lambda text, b=btn_custom_icon: handle_shape_change(text, b))
            btn_custom_icon.clicked.connect(lambda checked, b=btn_custom_icon: pick_custom_icon(b))

            btn_icolor = QPushButton()
            btn_icolor.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            btn_icolor.setMinimumWidth(20)
            self.set_btn_color(btn_icolor, "#ff0000")
            btn_icolor.clicked.connect(lambda checked, b=btn_icolor: self.pick_color(b))
            row_layout.addWidget(btn_icolor, 1)

            sb_opacity = QDoubleSpinBox()
            sb_opacity.setRange(0.1, 1.0)
            sb_opacity.setSingleStep(0.1)
            sb_opacity.setValue(1.0)
            sb_opacity.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            sb_opacity.setMinimumWidth(30)
            row_layout.addWidget(sb_opacity, 1)

            sb_iscale = QDoubleSpinBox()
            sb_iscale.setRange(0.5, 3.0)
            sb_iscale.setSingleStep(0.1)
            sb_iscale.setValue(1.2)
            sb_iscale.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            sb_iscale.setMinimumWidth(30)
            row_layout.addWidget(sb_iscale, 1)

            btn_lcolor = QPushButton()
            btn_lcolor.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            btn_lcolor.setMinimumWidth(20)
            self.set_btn_color(btn_lcolor, "#ffffff")
            btn_lcolor.clicked.connect(lambda checked, b=btn_lcolor: self.pick_color(b))
            row_layout.addWidget(btn_lcolor, 1)

            sb_lscale = QDoubleSpinBox()
            sb_lscale.setRange(0.0, 3.0)
            sb_lscale.setSingleStep(0.1)
            sb_lscale.setValue(1.0)
            sb_lscale.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            sb_lscale.setMinimumWidth(30)
            row_layout.addWidget(sb_lscale, 1)

            row_widget = QWidget()
            row_widget.setLayout(row_layout)
            self.scroll_layout.addWidget(row_widget)

            self.style_mappings[str(cat)] = {
                "shape": PyQtVar(cb_shape.currentText, cb_shape.setCurrentText),
                "custom_path": PyQtVar(lambda b=btn_custom_icon: b.property("custom_path"), lambda val, b=btn_custom_icon: b.setProperty("custom_path", val) or b.setStyleSheet("background-color: #00A2FF;" if val else "")),
                "icolor": PyQtVar(lambda b=btn_icolor: b.property("hex_val"), lambda val, b=btn_icolor: self.set_btn_color(b, val)),
                "opacity": PyQtVar(sb_opacity.value, sb_opacity.setValue),
                "iscale": PyQtVar(sb_iscale.value, sb_iscale.setValue),
                "lcolor": PyQtVar(lambda b=btn_lcolor: b.property("hex_val"), lambda val, b=btn_lcolor: self.set_btn_color(b, val)),
                "lscale": PyQtVar(sb_lscale.value, sb_lscale.setValue),
            }

    def set_btn_color(self, btn, hex_color):
        btn.setProperty("hex_val", hex_color)
        btn.setStyleSheet(f"background-color: {hex_color}; border: 1px solid rgba(128,128,128,0.5); border-radius: 4px;")

    def pick_color(self, btn):
        initial = QColor(btn.property("hex_val"))
        color = QColorDialog.getColor(initial, self, "Pick Color")
        if color.isValid(): self.set_btn_color(btn, color.name())

    def save_styles(self):
        if not self.style_mappings: return
        path, _ = QFileDialog.getSaveFileName(self, "Save Config", "", "JSON Config (*.json)")
        if not path: return
        data = {cat: {k: v.get() for k, v in config.items()} for cat, config in self.style_mappings.items()}
        with open(path, "w") as f: json.dump(data, f, indent=4)
        QMessageBox.information(self, "Saved", "Style configuration saved.")

    def load_styles(self):
        if not self.style_mappings: return
        path, _ = QFileDialog.getOpenFileName(self, "Load Config", "", "JSON Config (*.json)")
        if not path: return
        with open(path, "r") as f: data = json.load(f)
        for cat, config in data.items():
            if cat in self.style_mappings:
                for k, v in config.items():
                    if k in self.style_mappings[cat]: self.style_mappings[cat][k].set(v)
        QMessageBox.information(self, "Loaded", "Style configuration loaded.")

    def _build_tab_map(self):
        layout = QVBoxLayout(self.tab_map)
        layout.setContentsMargins(0,0,0,0)
        card = self.create_card_container()
        card_layout = QVBoxLayout(card)
        
        top_frame = QHBoxLayout()
        title = QLabel("Live Interactive Preview")
        title.setFont(QFont("Segoe UI", 16, QFont.Weight.Bold))
        title.setStyleSheet("color: #00A2FF;")
        
        self.cb_tiles = QComboBox()
        self.cb_tiles.addItems(["OpenStreetMap", "Google Normal", "Google Satellite"])
        self.cb_tiles.currentIndexChanged.connect(self.plot_map)
        
        self.btn_plot = QPushButton("Plot Points")
        self.btn_plot.setStyleSheet("background-color: #2E7D32; color: white; padding: 6px 16px; border-radius: 4px;")
        self.btn_plot.clicked.connect(self.plot_map)
        
        self.btn_clear = QPushButton("Clear Map")
        self.btn_clear.clicked.connect(self.clear_map)
        
        top_frame.addWidget(title)
        top_frame.addWidget(QLabel("Map Theme:"))
        top_frame.addWidget(self.cb_tiles)
        top_frame.addStretch()
        top_frame.addWidget(self.btn_plot)
        top_frame.addWidget(self.btn_clear)
        card_layout.addLayout(top_frame)

        self.web_view = QWebEngineView()
        card_layout.addWidget(self.web_view)
        layout.addWidget(card)
        self.clear_map()

    def clear_map(self):
        m = folium.Map(location=[0, 0], zoom_start=2)
        data = io.BytesIO()
        m.save(data, close_file=False)
        self.web_view.setHtml(data.getvalue().decode())

    def plot_map(self):
        if self.processor.df is None or self.processor.df.empty: return
        if self.processor.df[['__parsed_lat', '__parsed_lon']].isna().any().any():
            QMessageBox.warning(self, "Action Blocked", "Your dataset contains invalid coordinates.\nPlease fix them on the Data Upload tab.")
            return

        total_points = len(self.processor.df)
        if total_points > 500:
            QMessageBox.information(self, "Preview Limit", f"Preview safely limits to 500 points to ensure UI performance. All {total_points} points will be exported.")

        self.btn_plot.setText("Plotting...")
        self.btn_plot.setEnabled(False)
        QApplication.processEvents()

        plot_df = self.processor.df.head(500)
        min_lat, max_lat = plot_df["__parsed_lat"].min(), plot_df["__parsed_lat"].max()
        min_lon, max_lon = plot_df["__parsed_lon"].min(), plot_df["__parsed_lon"].max()
        
        tile_choice = self.cb_tiles.currentText()
        tiles = "OpenStreetMap"
        attr = None
        if "Google Normal" in tile_choice:
            tiles = "https://mt0.google.com/vt/lyrs=m&hl=en&x={x}&y={y}&z={z}&s=Ga"
            attr = "Google"
        elif "Google Satellite" in tile_choice:
            tiles = "https://mt0.google.com/vt/lyrs=s&hl=en&x={x}&y={y}&z={z}&s=Ga"
            attr = "Google"

        m = folium.Map(location=[(min_lat + max_lat)/2, (min_lon + max_lon)/2], tiles=tiles, attr=attr)
        for row in plot_df.to_dict('records'):
            lat, lon = row["__parsed_lat"], row["__parsed_lon"]
            name = str(row[self.processor.target_cols["point name"]])
            cat = str(row[self.processor.target_cols["category"]])
            hex_color = self.style_mappings[cat]["icolor"].get() if cat in self.style_mappings else "#00A2FF"

            folium.CircleMarker(
                location=[lat, lon], radius=6, color="white", weight=1, fill=True, fill_color=hex_color, fill_opacity=1.0, tooltip=name
            ).add_to(m)

        if not (min_lat == max_lat and min_lon == max_lon):
            m.fit_bounds([[min_lat, min_lon], [max_lat, max_lon]])

        data = io.BytesIO()
        m.save(data, close_file=False)
        self.web_view.setHtml(data.getvalue().decode())
        self.btn_plot.setText("Plot Points")
        self.btn_plot.setEnabled(True)

    def _build_tab_export(self):
        layout = QVBoxLayout(self.tab_export)
        layout.setContentsMargins(0,0,0,0)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        card = self.create_card_container()
        card_layout = QVBoxLayout(card)
        card_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        title = QLabel("Ready to Export")
        title.setFont(QFont("Segoe UI", 18, QFont.Weight.Bold))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("color: #00A2FF;")
        card_layout.addWidget(title)

        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.progress.setMinimumWidth(300)
        self.progress.setStyleSheet("""
            QProgressBar { border-radius: 4px; text-align: center; font-weight: bold; }
            QProgressBar::chunk { background-color: #00A2FF; border-radius: 4px; }
        """)
        card_layout.addWidget(self.progress)

        self.status_lbl = QLabel("Waiting...")
        self.status_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        card_layout.addWidget(self.status_lbl)
        card_layout.addSpacing(20)

        self.btn_export = QPushButton("Generate KML / KMZ")
        self.btn_export.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.btn_export.setMinimumWidth(250)
        self.btn_export.setStyleSheet("background-color: #00A2FF; color: white; padding: 12px; border-radius: 6px; font-weight: bold; font-size: 14px;")
        self.btn_export.clicked.connect(self.start_export_thread)
        card_layout.addWidget(self.btn_export)

        card_layout.addSpacing(10)

        self.btn_export_html = QPushButton("Generate HTML Web Map")
        self.btn_export_html.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.btn_export_html.setMinimumWidth(250)
        self.btn_export_html.setStyleSheet("background-color: #00897B; color: white; padding: 12px; border-radius: 6px; font-weight: bold; font-size: 14px;")
        self.btn_export_html.clicked.connect(self.start_html_export_thread)
        card_layout.addWidget(self.btn_export_html)

        h_layout = QHBoxLayout()
        h_layout.addStretch(1)
        h_layout.addWidget(card, stretch=2)
        h_layout.addStretch(1)
        layout.addLayout(h_layout)

        self.queue_timer = QTimer()
        self.queue_timer.timeout.connect(self.check_queue)

    def start_html_export_thread(self):
        if self.processor.df is None or len(self.processor.df) == 0:
            QMessageBox.warning(self, "Empty Dataset", "No valid points to export. Check your data.")
            return
        if self.processor.df[['__parsed_lat', '__parsed_lon']].isna().any().any():
            QMessageBox.warning(self, "Action Blocked", "Your dataset contains invalid coordinates.\nPlease fix the rows highlighted in red.")
            return

        style_config = {cat: {k: v.get() for k, v in config.items()} for cat, config in self.style_mappings.items()}
        save_path, _ = QFileDialog.getSaveFileName(self, "Save HTML Web Map", "", "HTML File (*.html)")
        if not save_path: return

        visible_cols = [self.list_columns.item(i).text() for i in range(self.list_columns.count()) if self.list_columns.item(i).checkState() == Qt.CheckState.Checked]

        self.btn_export.setEnabled(False)
        self.btn_export_html.setEnabled(False)
        self.progress.setValue(0)
        self.status_lbl.setText("Plotting points to map cluster...")

        threading.Thread(target=self.processor.export_to_html, args=(save_path, style_config, visible_cols, self.export_queue), daemon=True).start()
        self.queue_timer.start(100)

    def start_export_thread(self):
        if self.processor.df is None or len(self.processor.df) == 0:
            QMessageBox.warning(self, "Empty Dataset", "No valid points to export.")
            return
        if self.processor.df[['__parsed_lat', '__parsed_lon']].isna().any().any():
            QMessageBox.warning(self, "Action Blocked", "Your dataset contains invalid coordinates.")
            return

        style_config = {cat: {k: v.get() for k, v in config.items()} for cat, config in self.style_mappings.items()}
        has_custom_icons = any(c.get("shape") == "Custom Image" and c.get("custom_path") for c in style_config.values())

        save_path, _ = QFileDialog.getSaveFileName(self, "Save Map File", "", "KMZ File (*.kmz);;KML File (*.kml)")
        if not save_path: return

        if has_custom_icons and save_path.lower().endswith(".kml"):
            QMessageBox.warning(self, "Format Required", "You are using custom icon images.\nThe export format has been automatically changed to .kmz so the images can be bundled.")
            save_path = save_path[:-4] + ".kmz"

        visible_cols = [self.list_columns.item(i).text() for i in range(self.list_columns.count()) if self.list_columns.item(i).checkState() == Qt.CheckState.Checked]

        self.btn_export.setEnabled(False)
        self.btn_export_html.setEnabled(False)
        self.progress.setValue(0)
        self.status_lbl.setText("Generating XML syntax...")

        threading.Thread(target=self.processor.export_to_kml, args=(save_path, style_config, visible_cols, self.export_queue), daemon=True).start()
        self.queue_timer.start(100)

    def check_queue(self):
        try:
            while True:
                msg_type, val = self.export_queue.get_nowait()
                if msg_type == "progress": self.progress.setValue(val)
                elif msg_type == "status": self.status_lbl.setText(val)
                elif msg_type == "done":
                    self.progress.setValue(100)
                    self.status_lbl.setText("Export Complete!")
                    self.btn_export.setEnabled(True)
                    self.btn_export_html.setEnabled(True)
                    self.queue_timer.stop()
                    QMessageBox.information(self, "Success", f"File successfully written to:\n{val}")
                    return
                elif msg_type == "error":
                    self.status_lbl.setText("Exception occurred.")
                    self.btn_export.setEnabled(True)
                    self.btn_export_html.setEnabled(True)
                    self.queue_timer.stop()
                    QMessageBox.critical(self, "Export Error", val)
                    return
        except queue.Empty: pass 

    def _build_tab_contact(self):
        layout = QVBoxLayout(self.tab_contact)
        layout.setContentsMargins(0,0,0,0)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        card = self.create_card_container()
        card_layout = QVBoxLayout(card)
        card_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        title = QLabel("Contact Information")
        title.setFont(QFont("Segoe UI", 18, QFont.Weight.Bold))
        title.setStyleSheet("color: #00A2FF;")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        card_layout.addWidget(title)

        author_lbl = QLabel(f"Author: {AUTHOR_NAME}")
        author_lbl.setFont(QFont("Segoe UI", 12))
        author_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        card_layout.addWidget(author_lbl)

        email_lbl = QLabel("Email: Mohamed--Ashraf@outlook.com")
        email_lbl.setFont(QFont("Segoe UI", 12))
        email_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        card_layout.addWidget(email_lbl)

        link_lbl = QLabel('<a href="https://www.linkedin.com/in/mohamed---ashraf/" style="color: #00A2FF; text-decoration: none;">LinkedIn Profile</a>')
        link_lbl.setFont(QFont("Segoe UI", 12))
        link_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        link_lbl.setOpenExternalLinks(True)
        card_layout.addWidget(link_lbl)

        card_layout.addSpacing(30)
        footer = QLabel("Feel free to reach out for support, feedback, or feature requests!")
        footer.setFont(QFont("Segoe UI", 10, QFont.Weight.Normal, italic=True))
        footer.setAlignment(Qt.AlignmentFlag.AlignCenter)
        card_layout.addWidget(footer)
        
        h_layout = QHBoxLayout()
        h_layout.addStretch(1)
        h_layout.addWidget(card, stretch=2)
        h_layout.addStretch(1)
        layout.addLayout(h_layout)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    if HAS_QDARKTHEME:
        # Applies premium system-aware Light/Dark theme with custom Electric Blue accent
        qdarktheme.setup_theme("auto", custom_colors={"primary": "#00A2FF"})
        
    window = MainWindow()
    window.show()
    sys.exit(app.exec())