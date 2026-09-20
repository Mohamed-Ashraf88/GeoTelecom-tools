import sys
import os
import re
import json
import math
import subprocess
import webbrowser
import xml.etree.ElementTree as ET
import time
import urllib.request
import urllib.error
import urllib.parse
import random

try:
    import pandas as pd
    import pyproj
    from PyQt6.QtWidgets import (
        QApplication,
        QMainWindow,
        QWidget,
        QVBoxLayout,
        QHBoxLayout,
        QGridLayout,
        QTabWidget,
        QLabel,
        QLineEdit,
        QPushButton,
        QComboBox,
        QCheckBox,
        QTextEdit,
        QMessageBox,
        QFileDialog,
        QDialog,
        QFormLayout,
        QDialogButtonBox,
        QGroupBox,
    )
    from PyQt6.QtGui import (
        QIcon,
        QPixmap,
        QPainter,
        QColor,
        QPen,
        QAction,
        QFont,
        QPolygon,
        QBrush,
    )
    from PyQt6.QtCore import Qt, QThread, pyqtSignal, QSettings, QTimer, QPoint
    from PyQt6.QtWebEngineWidgets import QWebEngineView
    import openpyxl
except ImportError as e:
    import tkinter as tk
    from tkinter import messagebox

    root = tk.Tk()
    root.withdraw()
    messagebox.showerror(
        "Critical Dependency Missing",
        f"A required library is missing.\n\nDetails: {e}\n\nPlease install dependencies (including PyQt6-WebEngine) before running.",
    )
    sys.exit(1)

# ==========================================
# Tool Metadata
# ==========================================
__tool_name__ = "Spatial-Cordination Converter"
__version__ = "1.5.0"
__author__ = "Mohamed Ashraf"


# ==========================================
# Global Geocoding Helper
# ==========================================
def perform_reverse_geocode(lat, lon):
    url = f"https://nominatim.openstreetmap.org/reverse?format=json&lat={lat}&lon={lon}&zoom=14&addressdetails=1"
    req = urllib.request.Request(
        url, headers={"User-Agent": f"{__tool_name__}/{__version__} ({__author__})"}
    )
    try:
        with urllib.request.urlopen(req, timeout=5) as response:
            data = json.loads(response.read().decode("utf-8"))
            return data.get("display_name", "Address not found")
    except Exception as e:
        return f"Geocode Error: {str(e)}"


# ==========================================
# Column Mapping Dialog
# ==========================================
class ColumnMappingDialog(QDialog):
    def __init__(self, columns, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Map Data Columns")
        self.setModal(True)
        self.setMinimumWidth(350)
        self.layout = QFormLayout(self)

        info_label = QLabel(
            "Auto-detection failed.\nPlease map the correct columns from your file:"
        )
        info_label.setStyleSheet("color: #d35400; font-weight: bold;")
        self.layout.addRow(info_label)

        self.cb_lat = QComboBox()
        self.cb_lat.addItems(columns)
        self.layout.addRow("Latitude Column:", self.cb_lat)

        self.cb_lon = QComboBox()
        self.cb_lon.addItems(columns)
        self.layout.addRow("Longitude Column:", self.cb_lon)

        self.cb_point = QComboBox()
        self.cb_point.addItems(columns)
        self.layout.addRow("Point Name Column:", self.cb_point)

        self.buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        self.layout.addRow(self.buttons)

    def get_mapping(self):
        return (
            self.cb_lat.currentText(),
            self.cb_lon.currentText(),
            self.cb_point.currentText(),
        )


# ==========================================
# Worker Thread for Bulk Processing
# ==========================================
class ProcessWorker(QThread):
    log_msg = pyqtSignal(str)
    progress = pyqtSignal(int)
    finished = pyqtSignal(int, int, int, int, list, list)
    error = pyqtSignal(str)

    def __init__(
        self,
        input_file,
        output_file,
        precision,
        include_utm,
        export_kml,
        export_geojson,
        lat_col,
        lon_col,
        point_col,
        input_epsg,
        output_epsg,
        export_geocode,
    ):
        super().__init__()
        self.input_file = input_file
        self.output_file = output_file
        self.precision = precision
        self.include_utm = include_utm
        self.export_kml = export_kml
        self.export_geojson = export_geojson
        self.lat_col = lat_col
        self.lon_col = lon_col
        self.point_col = point_col
        self.input_epsg = input_epsg
        self.output_epsg = output_epsg
        self.export_geocode = export_geocode

    def extract_epsg(self, epsg_str):
        if not epsg_str or "none" in epsg_str.lower():
            return None
        match = re.search(r"\d+", epsg_str)
        return int(match.group()) if match else None

    def fuzzy_clean(self, coord_str):
        replacements = {"O": "0", "I": "1", "L": "1", "”": '"', "’": "'", "‘": "'"}
        for old, new in replacements.items():
            coord_str = coord_str.replace(old, new)
        return coord_str

    def parse_coordinate(self, coord_val, is_wgs84=True):
        if pd.isna(coord_val) or str(coord_val).strip() == "":
            raise ValueError("Empty or missing coordinate value")

        if not is_wgs84:
            clean_str = re.sub(r"[^\d\.\-]", "", str(coord_val))
            if not clean_str:
                raise ValueError(
                    f"No numeric value found in projected coordinate: '{coord_val}'"
                )
            return float(clean_str)

        coord_str = self.fuzzy_clean(str(coord_val).strip().upper())
        multiplier = 1

        if "S" in coord_str or "W" in coord_str:
            multiplier = -1
        elif coord_str.startswith("-"):
            multiplier = -1

        nums = re.findall(r"\d+(?:\.\d+)?", coord_str)
        if not nums:
            raise ValueError(f"No numeric values extracted from: '{coord_str}'")

        nums = [float(num) for num in nums]
        decimal_degrees = 0.0

        if len(nums) == 1:
            decimal_degrees = nums[0]
        elif len(nums) == 2:
            decimal_degrees = nums[0] + (nums[1] / 60)
        elif len(nums) >= 3:
            decimal_degrees = nums[0] + (nums[1] / 60) + (nums[2] / 3600)

        return decimal_degrees * multiplier

    def convert_to_utm(self, lat, lon):
        try:
            zone = math.floor((lon + 180) / 6) + 1
            zone = min(zone, 60)
            hemisphere = "north" if lat >= 0 else "south"

            crs_utm = pyproj.CRS.from_string(
                f"+proj=utm +zone={zone} +{hemisphere} +ellps=WGS84 +datum=WGS84 +units=m +no_defs"
            )
            crs_latlon = pyproj.CRS.from_epsg(4326)
            transformer = pyproj.Transformer.from_crs(
                crs_latlon, crs_utm, always_xy=True
            )

            easting, northing = transformer.transform(lon, lat)
            zone_letter = "N" if lat >= 0 else "S"

            return f"{zone}{zone_letter}", round(easting, 2), round(northing, 2)
        except pyproj.exceptions.CRSError:
            return "Error", 0.0, 0.0

    def run(self):
        try:
            ext = os.path.splitext(self.input_file)[1].lower()
            if ext in [".csv", ".txt"]:
                df = pd.read_csv(self.input_file)
            else:
                df = pd.read_excel(self.input_file)
        except Exception as e:
            self.error.emit(f"Failed to read file: {e}")
            return

        in_epsg_code = self.extract_epsg(self.input_epsg)
        out_epsg_code = self.extract_epsg(self.output_epsg)

        try:
            crs_wgs84 = pyproj.CRS.from_epsg(4326)
            transformer_in = (
                pyproj.Transformer.from_crs(
                    pyproj.CRS.from_epsg(in_epsg_code), crs_wgs84, always_xy=True
                )
                if in_epsg_code and in_epsg_code != 4326
                else None
            )
            transformer_out = (
                pyproj.Transformer.from_crs(
                    crs_wgs84, pyproj.CRS.from_epsg(out_epsg_code), always_xy=True
                )
                if out_epsg_code and out_epsg_code != 4326
                else None
            )
        except pyproj.exceptions.CRSError as e:
            self.error.emit(f"Invalid EPSG Code provided: {e}")
            return

        if (
            self.lat_col not in df.columns
            or self.lon_col not in df.columns
            or self.point_col not in df.columns
        ):
            self.error.emit("Selected columns are missing from the data.")
            return

        processed_data = []
        failed_rows_cache = []
        empty_rows_removed = 0
        total_rows = len(df)

        records = df.to_dict("records")
        is_wgs = not in_epsg_code or in_epsg_code == 4326

        for index, row in enumerate(records):
            if index % 250 == 0:
                self.progress.emit(index)

            pt_name = row.get(self.point_col)
            pt_name = (
                pt_name
                if pd.notna(pt_name) and str(pt_name).strip() != ""
                else f"Point_{index+1}"
            )

            lat_raw = row.get(self.lat_col)
            lon_raw = row.get(self.lon_col)

            if (
                pd.isna(lat_raw)
                or pd.isna(lon_raw)
                or str(lat_raw).strip() == ""
                or str(lon_raw).strip() == ""
            ):
                empty_rows_removed += 1
                continue

            try:
                x_raw = self.parse_coordinate(lon_raw, is_wgs84=is_wgs)
                y_raw = self.parse_coordinate(lat_raw, is_wgs84=is_wgs)

                if transformer_in:
                    lon_dd, lat_dd = transformer_in.transform(x_raw, y_raw)
                else:
                    lon_dd, lat_dd = x_raw, y_raw

                if not (-90 <= lat_dd <= 90) or not (-180 <= lon_dd <= 180):
                    raise ValueError(
                        "Coordinates out of global bounds after transformation to WGS84."
                    )

                row_data = {
                    "Point name": pt_name,
                    "Latitude": round(lat_dd, self.precision),
                    "Longitude": round(lon_dd, self.precision),
                }

                if transformer_out:
                    out_x, out_y = transformer_out.transform(lon_dd, lat_dd)
                    row_data[f"EPSG:{out_epsg_code} X"] = round(out_x, self.precision)
                    row_data[f"EPSG:{out_epsg_code} Y"] = round(out_y, self.precision)

                if self.export_geocode:
                    time.sleep(1.1)
                    row_data["Address Context"] = perform_reverse_geocode(
                        lat_dd, lon_dd
                    )

                if self.include_utm:
                    zone, east, north = self.convert_to_utm(lat_dd, lon_dd)
                    row_data["UTM Zone"] = zone
                    row_data["UTM Easting"] = east
                    row_data["UTM Northing"] = north

                processed_data.append(row_data)

            except ValueError as e:
                failed_rows_cache.append(
                    {"row_num": index + 2, "pt_name": pt_name, "error": str(e)}
                )

        self.finished.emit(
            total_rows,
            len(processed_data),
            empty_rows_removed,
            len(failed_rows_cache),
            processed_data,
            failed_rows_cache,
        )


# ==========================================
# Main GUI Application
# ==========================================
class GeoCoordApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{__tool_name__} v{__version__}")

        # Widened to support Side Panel Map & Floating Cards
        self.setMinimumSize(1250, 750)
        self.setAcceptDrops(True)

        self.settings = QSettings(__author__, __tool_name__)
        self.last_dir = self.settings.value("last_dir", os.path.expanduser("~"))

        self.failed_rows_cache = []
        self.current_lat = 0.0
        self.current_lon = 0.0

        self.apply_dynamic_icon()
        self.apply_stylesheet()
        self.setup_ui()
        self.setup_menu()

    def apply_stylesheet(self):
        """Injects a neutral, system-adaptive stylesheet for a premium UI feel."""
        qss = """
            QMainWindow { background-color: #f8f9fa; }
            QLabel { color: #2d3436; font-family: 'Segoe UI', Arial, sans-serif; }
            
            QTabWidget::pane { 
                border: 1px solid #dcdde1; 
                background: white; 
                border-radius: 8px;
            }
            QTabBar::tab {
                background: #f1f2f6;
                border: 1px solid #dcdde1;
                padding: 8px 16px;
                margin-right: 2px;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
                color: #57606f;
            }
            QTabBar::tab:selected {
                background: white;
                border-bottom-color: white;
                color: #2d3436;
                font-weight: bold;
            }
            
            QGroupBox {
                font-weight: bold;
                border: 1px solid #dfe4ea;
                border-radius: 8px;
                margin-top: 14px;
                padding-top: 15px;
                background-color: #ffffff;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                left: 15px;
                padding: 0 8px;
                color: #2f3542;
            }
            
            QLineEdit, QComboBox {
                border: 1px solid #ced6e0;
                border-radius: 5px;
                padding: 6px;
                background-color: white;
                color: #2f3542;
                selection-background-color: #3498db;
            }
            QLineEdit:focus, QComboBox:focus {
                border: 1px solid #3498db;
            }
            
            QPushButton {
                background-color: #ffffff;
                border: 1px solid #ced6e0;
                border-radius: 5px;
                padding: 6px 12px;
                color: #2f3542;
            }
            QPushButton:hover {
                background-color: #f1f2f6;
                border: 1px solid #a4b0be;
            }
            QPushButton:pressed {
                background-color: #dfe4ea;
            }
            
            /* Primary Action Button (e.g. Start Conversion) */
            QPushButton#primaryAction {
                background-color: #27ae60;
                color: white;
                border: none;
                font-weight: bold;
                padding: 10px 15px;
            }
            QPushButton#primaryAction:hover {
                background-color: #2ecc71;
            }
            QPushButton#primaryAction:pressed {
                background-color: #229954;
            }
            QPushButton#primaryAction:disabled {
                background-color: #95a5a6;
                color: #ecf0f1;
            }
            
            QTextEdit {
                border: 1px solid #ced6e0;
                border-radius: 6px;
                background-color: #fbfbfb;
                color: #57606f;
            }
        """
        self.setStyleSheet(qss)

    def apply_dynamic_icon(self):
        size = 256
        pixmap = QPixmap(size, size)
        pixmap.fill(Qt.GlobalColor.transparent)
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setBrush(QColor("#1e8449"))
        painter.setPen(Qt.PenStyle.NoPen)
        painter.drawEllipse(16, 16, 224, 224)
        pen = QPen(QColor("#2ecc71"), 6)
        painter.setPen(pen)
        painter.setBrush(Qt.BrushStyle.NoBrush)
        painter.drawEllipse(64, 16, 128, 224)
        painter.drawEllipse(104, 16, 48, 224)
        pen.setColor(QColor("#abebc6"))
        painter.setPen(pen)
        painter.drawLine(16, 128, 240, 128)
        painter.drawLine(128, 16, 128, 240)
        painter.end()
        self.setWindowIcon(QIcon(pixmap))

    def create_action_icon(self, icon_type):
        """Generates crisp programmatic icons for the embedded input actions."""
        size = 24
        pixmap = QPixmap(size, size)
        pixmap.fill(Qt.GlobalColor.transparent)
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        if icon_type == "folder":
            painter.setPen(QPen(QColor("#57606f"), 1.5))
            painter.setBrush(QBrush(QColor("#dfe4ea")))
            poly = QPolygon(
                [
                    QPoint(2, 6),
                    QPoint(9, 6),
                    QPoint(11, 9),
                    QPoint(22, 9),
                    QPoint(22, 20),
                    QPoint(2, 20),
                ]
            )
            painter.drawPolygon(poly)

        elif icon_type == "save":
            painter.setPen(QPen(QColor("#57606f"), 1.5))
            painter.setBrush(QBrush(QColor("#dfe4ea")))
            painter.drawRect(3, 3, 18, 18)
            painter.setBrush(QBrush(QColor("#ffffff")))
            painter.drawRect(7, 3, 10, 6)
            painter.drawRect(6, 13, 12, 8)

        painter.end()
        return QIcon(pixmap)

    def setup_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_h_layout = QHBoxLayout(central_widget)
        main_h_layout.setContentsMargins(15, 15, 15, 15)
        main_h_layout.setSpacing(20)

        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)

        header_label = QLabel(f"{__tool_name__}")
        header_label.setFont(QFont("Segoe UI", 24, QFont.Weight.Bold))
        header_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        left_layout.addWidget(header_label)

        sub_header = QLabel(f"v{__version__} | Developed by {__author__}")
        sub_header.setStyleSheet("color: #7f8c8d; font-size: 13px;")
        sub_header.setAlignment(Qt.AlignmentFlag.AlignCenter)
        left_layout.addWidget(sub_header)
        left_layout.addSpacing(15)

        self.tabs = QTabWidget()
        self.tab_bulk = QWidget()
        self.tab_single = QWidget()

        self.tabs.addTab(self.tab_bulk, "Bulk Processing")
        self.tabs.addTab(self.tab_single, "Single Point")
        left_layout.addWidget(self.tabs)

        self.build_bulk_tab()
        self.build_single_tab()
        self.load_settings()

        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(0, 0, 0, 0)

        self.map_view = QWebEngineView()
        self.map_view.setStyleSheet("border: 1px solid #ced6e0; border-radius: 8px;")
        right_layout.addWidget(self.map_view)
        self.init_embedded_map()

        main_h_layout.addWidget(left_panel, stretch=4)
        main_h_layout.addWidget(right_panel, stretch=5)

    def init_embedded_map(self):
        """Injects Leaflet JS, OpenStreetMap, and Floating Card UI directly into the UI."""
        html_content = """
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="utf-8">
            <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
            <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
            <style>
                body { padding: 0; margin: 0; font-family: 'Segoe UI', Arial, sans-serif;}
                html, body, #map { height: 100%; width: 100%; }
                
                #result-card {
                    display: none;
                    position: absolute;
                    top: 20px;
                    right: 20px;
                    z-index: 1000;
                    background: rgba(255, 255, 255, 0.95);
                    backdrop-filter: blur(5px);
                    padding: 15px 20px;
                    border-radius: 10px;
                    box-shadow: 0 10px 15px -3px rgba(0,0,0,0.1), 0 4px 6px -2px rgba(0,0,0,0.05);
                    border: 1px solid rgba(255,255,255,0.2);
                    width: 280px;
                    transition: all 0.3s ease;
                }
                .card-title {
                    margin: 0 0 10px 0;
                    color: #2f3640;
                    font-size: 15px;
                    font-weight: 700;
                    text-transform: uppercase;
                    letter-spacing: 0.5px;
                }
                .data-row { margin: 8px 0; font-size: 14px; }
                .rc-label { color: #7f8c8d; font-weight: 600; font-size: 12px; display: block; }
                #rc-latlon { color: #27ae60; font-weight: bold; font-size: 15px; }
                #rc-utm { color: #2f3542; font-family: monospace; }
                #rc-address { 
                    color: #2980b9; 
                    font-size: 13px; 
                    font-style: italic; 
                    margin-top: 12px;
                    padding-top: 10px;
                    border-top: 1px dashed #dcdde1;
                }
                .error-state #rc-latlon { color: #e74c3c; }
            </style>
        </head>
        <body>
            <div id="result-card">
                <h3 class="card-title">Location Data</h3>
                <div class="data-row">
                    <span class="rc-label">WGS84 Coordinates</span>
                    <span id="rc-latlon">--</span>
                </div>
                <div class="data-row">
                    <span class="rc-label">Projected UTM</span>
                    <span id="rc-utm">--</span>
                </div>
                <div class="data-row" id="address-container">
                    <span id="rc-address">--</span>
                </div>
            </div>

            <div id="map"></div>
            <script>
                var map = L.map('map', { zoomControl: false }).setView([20.0, 0.0], 2);
                L.control.zoom({ position: 'bottomright' }).addTo(map);
                
                L.tileLayer('https://{s}.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}{r}.png', {
                    maxZoom: 19,
                    attribution: '© OpenStreetMap | © CARTO'
                }).addTo(map);

                var markersGroup = L.featureGroup().addTo(map);

                function addMarker(lat, lon, title) {
                    var marker = L.marker([lat, lon]).bindPopup("<b>" + title + "</b><br>" + lat + ", " + lon);
                    markersGroup.addLayer(marker);
                }

                function clearMap() { markersGroup.clearLayers(); }

                function fitMap() {
                    if (markersGroup.getLayers().length > 0) {
                        map.fitBounds(markersGroup.getBounds(), {padding: [30, 30], maxZoom: 15});
                    }
                }
                
                function updateResultCard(latlon, utm, address, isError) {
                    var card = document.getElementById('result-card');
                    card.style.display = 'block';
                    
                    if (isError) {
                        card.className = "error-state";
                        document.getElementById('rc-latlon').innerText = latlon; 
                        document.getElementById('rc-utm').innerText = "N/A";
                        document.getElementById('rc-address').innerText = "";
                    } else {
                        card.className = "";
                        document.getElementById('rc-latlon').innerText = latlon;
                        document.getElementById('rc-utm').innerText = utm;
                        
                        var addrEl = document.getElementById('rc-address');
                        if (address) {
                            addrEl.innerText = address;
                            addrEl.style.display = 'block';
                        } else {
                            addrEl.style.display = 'none';
                        }
                    }
                }
                
                function hideResultCard() {
                    document.getElementById('result-card').style.display = 'none';
                }
            </script>
        </body>
        </html>
        """
        self.map_view.setHtml(html_content)

    def load_settings(self):
        self.combo_precision.setCurrentText(str(self.settings.value("precision", "6")))
        self.combo_in_epsg.setCurrentText(
            str(self.settings.value("in_epsg", "4326 - WGS84 (Lat/Lon)"))
        )
        self.combo_out_epsg.setCurrentText(
            str(self.settings.value("out_epsg", "None (Keep Original)"))
        )
        self.chk_utm.setChecked(self.settings.value("include_utm", True, type=bool))
        self.chk_kml.setChecked(self.settings.value("export_kml", False, type=bool))
        self.chk_geojson.setChecked(
            self.settings.value("export_geojson", False, type=bool)
        )
        self.chk_geocode.setChecked(
            self.settings.value("export_geocode", False, type=bool)
        )
        self.chk_single_geocode.setChecked(
            self.settings.value("single_geocode", True, type=bool)
        )

    def closeEvent(self, event):
        self.settings.setValue("precision", self.combo_precision.currentText())
        self.settings.setValue("in_epsg", self.combo_in_epsg.currentText())
        self.settings.setValue("out_epsg", self.combo_out_epsg.currentText())
        self.settings.setValue("include_utm", self.chk_utm.isChecked())
        self.settings.setValue("export_kml", self.chk_kml.isChecked())
        self.settings.setValue("export_geojson", self.chk_geojson.isChecked())
        self.settings.setValue("export_geocode", self.chk_geocode.isChecked())
        self.settings.setValue("single_geocode", self.chk_single_geocode.isChecked())
        self.settings.setValue("last_dir", self.last_dir)
        super().closeEvent(event)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if urls and urls[0].isLocalFile():
                ext = os.path.splitext(urls[0].toLocalFile())[1].lower()
                if ext in [".xlsx", ".xls", ".csv", ".txt"]:
                    event.acceptProposedAction()
                    return
        event.ignore()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if urls and urls[0].isLocalFile():
            file_path = urls[0].toLocalFile()
            self.entry_input_path.setText(file_path)
            out_path = file_path.rsplit(".", 1)[0] + "_converted.xlsx"
            self.entry_output_path.setText(out_path)
            self.tabs.setCurrentIndex(0)

    def build_bulk_tab(self):
        layout = QVBoxLayout(self.tab_bulk)
        layout.setSpacing(10)

        group_data = QGroupBox("1. Data Source")
        data_layout = QFormLayout(group_data)

        self.entry_input_path = QLineEdit()
        self.entry_input_path.setPlaceholderText("Select or Drag & Drop input file...")
        action_browse_in = self.entry_input_path.addAction(
            self.create_action_icon("folder"), QLineEdit.ActionPosition.TrailingPosition
        )
        action_browse_in.triggered.connect(self.select_input)

        self.entry_output_path = QLineEdit()
        self.entry_output_path.setPlaceholderText("Output destination...")
        action_browse_out = self.entry_output_path.addAction(
            self.create_action_icon("save"), QLineEdit.ActionPosition.TrailingPosition
        )
        action_browse_out.triggered.connect(self.select_output)

        data_layout.addRow("Input:", self.entry_input_path)
        data_layout.addRow("Output:", self.entry_output_path)
        layout.addWidget(group_data)

        group_crs = QGroupBox("2. Coordinate Reference System (CRS)")
        crs_layout = QGridLayout(group_crs)

        self.combo_in_epsg = QComboBox()
        self.combo_in_epsg.setEditable(True)
        self.combo_in_epsg.addItems(
            [
                "4326 - WGS84 (Lat/Lon)",
                "3857 - Web Mercator",
                "4269 - NAD83",
                "32631 - UTM 31N",
                "27700 - British National Grid",
            ]
        )

        self.combo_out_epsg = QComboBox()
        self.combo_out_epsg.setEditable(True)
        self.combo_out_epsg.addItems(
            [
                "None (Keep Original)",
                "4326 - WGS84 (Lat/Lon)",
                "3857 - Web Mercator",
                "32631 - UTM 31N",
                "27700 - British National Grid",
            ]
        )

        crs_layout.addWidget(QLabel("Input CRS (EPSG):"), 0, 0)
        crs_layout.addWidget(self.combo_in_epsg, 1, 0)
        crs_layout.addWidget(QLabel("Target CRS (EPSG):"), 0, 1)
        crs_layout.addWidget(self.combo_out_epsg, 1, 1)
        layout.addWidget(group_crs)

        group_options = QGroupBox("3. Export Options")
        opt_layout = QHBoxLayout(group_options)

        opt_layout.addWidget(QLabel("Decimals:"))
        self.combo_precision = QComboBox()
        self.combo_precision.addItems(["4", "6", "8"])
        self.combo_precision.setCurrentText("6")
        opt_layout.addWidget(self.combo_precision)

        self.chk_utm = QCheckBox("Calculate UTM")
        opt_layout.addWidget(self.chk_utm)
        self.chk_kml = QCheckBox(".KML")
        opt_layout.addWidget(self.chk_kml)
        self.chk_geojson = QCheckBox(".GeoJSON")
        opt_layout.addWidget(self.chk_geojson)
        self.chk_geocode = QCheckBox("Reverse Geocode")
        self.chk_geocode.setToolTip("Warning: Limits processing to 1 row/second.")
        opt_layout.addWidget(self.chk_geocode)
        layout.addWidget(group_options)

        layout.addSpacing(10)
        self.btn_process = QPushButton("START BATCH CONVERSION")
        self.btn_process.setObjectName("primaryAction")
        layout.addWidget(self.btn_process)
        self.btn_process.clicked.connect(self.process_files)

        self.btn_toggle_log = QPushButton("Show Activity Log ▼")
        self.btn_toggle_log.setStyleSheet(
            "background: transparent; border: none; color: #3498db; text-align: left;"
        )
        self.btn_toggle_log.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_toggle_log.clicked.connect(self.toggle_log_visibility)
        layout.addWidget(self.btn_toggle_log)

        self.log_textbox = QTextEdit()
        self.log_textbox.setReadOnly(True)
        self.log_textbox.setMaximumHeight(150)
        self.log_textbox.hide()
        layout.addWidget(self.log_textbox)

        layout.addStretch()
        self.log_message("System initialized and ready.")

    def build_single_tab(self):
        layout = QVBoxLayout(self.tab_single)
        layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        instruction = QLabel(
            "Enter coordinates in any format (DD, DDM, DMS).\nTypo correction & fuzzy matching are active."
        )
        instruction.setStyleSheet("color: #7f8c8d; font-size: 13px;")
        instruction.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(instruction)
        layout.addSpacing(15)

        group_input = QGroupBox("Target Coordinates")
        form = QFormLayout(group_input)

        self.entry_single_lat = QLineEdit()
        self.entry_single_lat.setPlaceholderText("e.g. 40° 42.768' N")
        self.lbl_feedback_lat = QLabel("")
        self.lbl_feedback_lat.setStyleSheet(
            "color: #95a5a6; font-size: 11px; font-style: italic;"
        )

        self.entry_single_lon = QLineEdit()
        self.entry_single_lon.setPlaceholderText("e.g. -74.0060")
        self.lbl_feedback_lon = QLabel("")
        self.lbl_feedback_lon.setStyleSheet(
            "color: #95a5a6; font-size: 11px; font-style: italic;"
        )

        form.addRow("Latitude:", self.entry_single_lat)
        form.addRow("", self.lbl_feedback_lat)
        form.addRow("Longitude:", self.entry_single_lon)
        form.addRow("", self.lbl_feedback_lon)
        layout.addWidget(group_input)

        self.lat_timer = QTimer()
        self.lat_timer.setSingleShot(True)
        self.lat_timer.timeout.connect(
            lambda: self.update_live_feedback(
                self.entry_single_lat, self.lbl_feedback_lat
            )
        )
        self.entry_single_lat.textChanged.connect(lambda: self.lat_timer.start(500))

        self.lon_timer = QTimer()
        self.lon_timer.setSingleShot(True)
        self.lon_timer.timeout.connect(
            lambda: self.update_live_feedback(
                self.entry_single_lon, self.lbl_feedback_lon
            )
        )
        self.entry_single_lon.textChanged.connect(lambda: self.lon_timer.start(500))

        layout.addSpacing(10)
        self.chk_single_geocode = QCheckBox("Fetch Address Context (Reverse Geocoding)")
        layout.addWidget(self.chk_single_geocode)
        layout.addSpacing(15)

        btn_convert = QPushButton("PLOT POINT ON MAP")
        btn_convert.setObjectName("primaryAction")
        btn_convert.clicked.connect(self.convert_single_point)
        layout.addWidget(btn_convert)

        layout.addStretch()

    def toggle_log_visibility(self):
        if self.log_textbox.isVisible():
            self.log_textbox.hide()
            self.btn_toggle_log.setText("Show Activity Log ▼")
        else:
            self.log_textbox.show()
            self.btn_toggle_log.setText("Hide Activity Log ▲")

    def update_live_feedback(self, entry_widget, label_widget):
        text = entry_widget.text().strip()
        if not text:
            label_widget.setText("")
            return

        text_upper = text.upper()
        fuzzies = ["O", "I", "L", "”", "’", "‘"]
        typo_count = sum(1 for char in text_upper if char in fuzzies)

        nums = re.findall(r"\d+(?:\.\d+)?", text_upper)
        if not nums:
            label_widget.setText("⚠️ Waiting for valid numbers...")
            return

        fmt = "Unknown"
        if len(nums) == 1:
            fmt = "Decimal Degrees (DD)"
        elif len(nums) == 2:
            fmt = "Degrees Minutes (DDM)"
        else:
            fmt = "Degrees/Min/Sec (DMS)"

        conf = "High" if re.search(r"[NSWE\-]", text_upper) else "Medium"
        res = f"Format: {fmt} | Confidence: {conf}"
        if typo_count > 0:
            res += f" | ⚠️ Fixing {typo_count} typo(s)"

        label_widget.setText(res)

    def setup_menu(self):
        menubar = self.menuBar()
        file_menu = menubar.addMenu("File")
        import_action = QAction("Import File...", self)
        import_action.triggered.connect(self.menu_import_file)
        file_menu.addAction(import_action)
        output_action = QAction("Choose Output Location...", self)
        output_action.triggered.connect(self.menu_choose_output)
        file_menu.addAction(output_action)
        file_menu.addSeparator()
        switch_action = QAction("Switch View", self)
        switch_action.triggered.connect(self.menu_switch_view)
        file_menu.addAction(switch_action)
        file_menu.addSeparator()
        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        help_menu = menubar.addMenu("Help")
        about_action = QAction("About the Tool", self)
        about_action.triggered.connect(self.menu_about)
        help_menu.addAction(about_action)
        contact_action = QAction("Contact Us", self)
        contact_action.triggered.connect(self.menu_contact)
        help_menu.addAction(contact_action)

    def export_kml(self, data, base_filepath):
        kml_path = base_filepath.rsplit(".", 1)[0] + ".kml"
        kml = ET.Element("kml", xmlns="http://www.opengis.net/kml/2.2")
        doc = ET.SubElement(kml, "Document")
        doc_name = ET.SubElement(doc, "name")
        doc_name.text = "Exported Coordinates"
        for row in data:
            placemark = ET.SubElement(doc, "Placemark")
            name = ET.SubElement(placemark, "name")
            name.text = str(row["Point name"])
            point = ET.SubElement(placemark, "Point")
            coords = ET.SubElement(point, "coordinates")
            coords.text = f"{row['Longitude']},{row['Latitude']},0"
        tree = ET.ElementTree(kml)
        tree.write(kml_path, encoding="utf-8", xml_declaration=True)
        self.log_message(f"GIS: Generated {kml_path}")

    def export_geojson(self, data, base_filepath):
        json_path = base_filepath.rsplit(".", 1)[0] + ".geojson"
        features = []
        for row in data:
            features.append(
                {
                    "type": "Feature",
                    "geometry": {
                        "type": "Point",
                        "coordinates": [row["Longitude"], row["Latitude"]],
                    },
                    "properties": {"name": row["Point name"]},
                }
            )
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump({"type": "FeatureCollection", "features": features}, f, indent=2)
        self.log_message(f"GIS: Generated {json_path}")

    def log_message(self, message):
        self.log_textbox.append(message)

    def select_input(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Open File",
            self.last_dir,
            "Supported Files (*.xlsx *.xls *.csv *.txt);;All files (*.*)",
        )
        if path:
            self.entry_input_path.setText(path)
            self.last_dir = os.path.dirname(path)
            if not self.entry_output_path.text().strip():
                out_path = path.rsplit(".", 1)[0] + "_converted.xlsx"
                self.entry_output_path.setText(out_path)

    def select_output(self):
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Excel File",
            self.last_dir,
            "Excel files (*.xlsx);;All files (*.*)",
        )
        if path:
            self.entry_output_path.setText(path)
            self.last_dir = os.path.dirname(path)

    def open_file_cross_platform(self, filepath):
        try:
            if sys.platform == "win32":
                os.startfile(filepath)
            elif sys.platform == "darwin":
                subprocess.call(["open", filepath])
            else:
                subprocess.call(["xdg-open", filepath])
        except Exception:
            pass

    def process_files(self):
        input_file = self.entry_input_path.text().strip('"')
        output_file = self.entry_output_path.text().strip('"')

        if not input_file or not output_file:
            QMessageBox.warning(
                self,
                "Missing Files",
                "Please provide both input and output file paths.",
            )
            return

        if self.chk_geocode.isChecked():
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Icon.Warning)
            msg.setWindowTitle("Geocoding Rate Limit Warning")
            msg.setText(
                "Reverse Geocoding is enabled.\n\nTo comply with OpenStreetMap API limits, processing is restricted to 1 row per second. A 600-row file will take roughly 10 minutes.\n\nDo you want to continue?"
            )
            msg.setStandardButtons(
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if msg.exec() == QMessageBox.StandardButton.No:
                return

        self.btn_process.setEnabled(False)
        if not self.log_textbox.isVisible():
            self.toggle_log_visibility()

        self.log_message("-" * 40)
        self.log_message("Starting coordinate conversion...")
        self.map_view.page().runJavaScript("hideResultCard();")

        try:
            ext = os.path.splitext(input_file)[1].lower()
            if ext in [".csv", ".txt"]:
                df_headers = pd.read_csv(input_file, nrows=0)
            else:
                df_headers = pd.read_excel(input_file, nrows=0)
        except Exception as e:
            self.log_message(f"CRITICAL ERROR: Failed to read file headers. {e}")
            self.btn_process.setEnabled(True)
            return

        original_cols = df_headers.columns.tolist()
        normalized_cols = [str(col).strip().lower() for col in original_cols]
        col_map = dict(zip(normalized_cols, original_cols))

        lat_aliases = ["latitude", "lat", "y", "lat.", "latitude_deg"]
        lon_aliases = ["longitude", "lon", "long", "lng", "x", "lon.", "longitude_deg"]
        name_aliases = [
            "point name",
            "site_name",
            "name",
            "id",
            "point",
            "site id",
            "station",
            "site",
        ]

        lat_key = next((col for col in normalized_cols if col in lat_aliases), None)
        lon_key = next((col for col in normalized_cols if col in lon_aliases), None)
        point_key = next((col for col in normalized_cols if col in name_aliases), None)

        if not lat_key or not lon_key or not point_key:
            dialog = ColumnMappingDialog(original_cols, self)
            if lat_key:
                dialog.cb_lat.setCurrentText(col_map[lat_key])
            if lon_key:
                dialog.cb_lon.setCurrentText(col_map[lon_key])
            if point_key:
                dialog.cb_point.setCurrentText(col_map[point_key])

            if dialog.exec() == QDialog.DialogCode.Accepted:
                lat_col, lon_col, point_col = dialog.get_mapping()
                self.log_message("Manual column mapping applied.")
            else:
                self.log_message("Conversion canceled by user during column mapping.")
                self.btn_process.setEnabled(True)
                return
        else:
            lat_col = col_map[lat_key]
            lon_col = col_map[lon_key]
            point_col = col_map[point_key]

        self.worker = ProcessWorker(
            input_file,
            output_file,
            int(self.combo_precision.currentText()),
            self.chk_utm.isChecked(),
            self.chk_kml.isChecked(),
            self.chk_geojson.isChecked(),
            lat_col,
            lon_col,
            point_col,
            self.combo_in_epsg.currentText(),
            self.combo_out_epsg.currentText(),
            self.chk_geocode.isChecked(),
        )
        self.worker.log_msg.connect(self.log_message)
        self.worker.error.connect(self.handle_worker_error)
        self.worker.finished.connect(self.handle_worker_finished)
        self.worker.start()

    def handle_worker_error(self, err_msg):
        self.log_message(f"CRITICAL ERROR: {err_msg}")
        self.btn_process.setEnabled(True)

    def handle_worker_finished(
        self, total, success, empty, failed, processed_data, failed_cache
    ):
        self.failed_rows_cache = failed_cache
        output_file = self.entry_output_path.text().strip('"')

        if processed_data:
            try:
                output_df = pd.DataFrame(processed_data)
                output_df.to_excel(output_file, index=False, engine="openpyxl")
                self.log_message(
                    f"SUCCESS: Processed {len(processed_data)} coordinates."
                )

                if self.chk_kml.isChecked():
                    self.export_kml(processed_data, output_file)
                if self.chk_geojson.isChecked():
                    self.export_geojson(processed_data, output_file)

                self.open_file_cross_platform(output_file)

                self.map_view.page().runJavaScript("clearMap();")
                sample_size = min(50, len(processed_data))
                sample = random.sample(processed_data, sample_size)

                for pt in sample:
                    lat = pt.get("Latitude", 0.0)
                    lon = pt.get("Longitude", 0.0)
                    name = str(pt.get("Point name", "")).replace("'", "\\'")
                    self.map_view.page().runJavaScript(
                        f"addMarker({lat}, {lon}, '{name}');"
                    )

                self.map_view.page().runJavaScript("fitMap();")

            except Exception as e:
                self.log_message(f"ERROR: Could not save output file. {e}")
        else:
            self.log_message("WARNING: No rows were successfully processed.")

        if self.failed_rows_cache:
            self.log_message(f"\n--- FAILED ROWS: {len(self.failed_rows_cache)} ---")

        self.btn_process.setEnabled(True)
        self.show_statistics_popup(total, success, empty, failed)

    def show_statistics_popup(self, total, success, empty, failed):
        msg = QMessageBox(self)
        msg.setWindowTitle("Process Summary")
        msg.setText("Conversion Complete!")
        details = (
            f"Total Rows Processed: {total}\n"
            f"Successfully Converted: {success}\n"
            f"Removed (Empty Cells): {empty}\n"
            f"Failed (Invalid Data): {failed}"
        )
        msg.setInformativeText(details)
        if failed > 0:
            btn_log = msg.addButton("Save Error Log", QMessageBox.ButtonRole.ActionRole)
            btn_log.clicked.connect(self.export_error_log)
        msg.addButton(QMessageBox.StandardButton.Close)
        msg.exec()

    def export_error_log(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Error Log", "", "Text File (*.txt)"
        )
        if path:
            with open(path, "w", encoding="utf-8") as f:
                f.write(f"--- {__tool_name__} Error Log ---\n\n")
                for fail in self.failed_rows_cache:
                    f.write(
                        f"Row {fail['row_num']} | Point: {fail['pt_name']}\nReason: {fail['error']}\n\n"
                    )
            self.open_file_cross_platform(path)

    def convert_single_point(self):
        lat_raw = self.entry_single_lat.text()
        lon_raw = self.entry_single_lon.text()
        success = True
        precision = 6

        util = ProcessWorker("", "", 0, False, False, False, "", "", "", "", "", False)

        try:
            self.current_lat = util.parse_coordinate(lat_raw)
            if not (-90 <= self.current_lat <= 90):
                raise ValueError("Out of bounds")
        except Exception:
            success = False

        try:
            self.current_lon = util.parse_coordinate(lon_raw)
            if not (-180 <= self.current_lon <= 180):
                raise ValueError("Out of bounds")
        except Exception:
            success = False

        if success:
            zone, east, north = util.convert_to_utm(self.current_lat, self.current_lon)
            latlon_str = (
                f"{self.current_lat:.{precision}f}, {self.current_lon:.{precision}f}"
            )
            utm_str = f"Zone {zone} | E: {east} | N: {north}"

            self.map_view.page().runJavaScript("clearMap();")
            self.map_view.page().runJavaScript(
                f"addMarker({self.current_lat}, {self.current_lon}, 'Target Location');"
            )
            self.map_view.page().runJavaScript("fitMap();")

            if self.chk_single_geocode.isChecked():
                loading_js = f"updateResultCard({json.dumps(latlon_str)}, {json.dumps(utm_str)}, 'Fetching local address data...', false);"
                self.map_view.page().runJavaScript(loading_js)
                QApplication.processEvents()
                time.sleep(1.1)
                addr = perform_reverse_geocode(self.current_lat, self.current_lon)
                final_js = f"updateResultCard({json.dumps(latlon_str)}, {json.dumps(utm_str)}, {json.dumps(addr)}, false);"
                self.map_view.page().runJavaScript(final_js)
            else:
                js = f"updateResultCard({json.dumps(latlon_str)}, {json.dumps(utm_str)}, '', false);"
                self.map_view.page().runJavaScript(js)
        else:
            err_msg = "Invalid Coordinates or Out of Bounds (-90 to 90, -180 to 180)."
            js = f"updateResultCard({json.dumps(err_msg)}, '', '', true);"
            self.map_view.page().runJavaScript("clearMap();")
            self.map_view.page().runJavaScript(js)

    def menu_import_file(self):
        self.tabs.setCurrentIndex(0)
        self.select_input()

    def menu_choose_output(self):
        self.tabs.setCurrentIndex(0)
        self.select_output()

    def menu_switch_view(self):
        idx = 1 if self.tabs.currentIndex() == 0 else 0
        self.tabs.setCurrentIndex(idx)

    # MODIFICATION 2: Help Menu: About the Tool Window
    def menu_about(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("About the Tool")
        dialog.resize(550, 450)
        layout = QVBoxLayout(dialog)

        # Title and Subtitle styling
        lbl_title = QLabel(f"{__tool_name__} v{__version__}")
        lbl_title.setFont(QFont("Segoe UI", 18, QFont.Weight.Bold))
        lbl_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(lbl_title)

        lbl_subtitle = QLabel(f"Developed by {__author__}")
        lbl_subtitle.setStyleSheet("color: #7f8c8d; font-size: 13px;")
        lbl_subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(lbl_subtitle)

        layout.addSpacing(10)

        # Scrollable text area for comprehensive tool summary
        text_edit = QTextEdit()
        text_edit.setReadOnly(True)
        text_edit.setStyleSheet("""
            QTextEdit {
                background-color: #fcfcfc; 
                border: 1px solid #ced6e0; 
                border-radius: 6px; 
                padding: 10px;
                color: #2f3542;
                font-family: 'Segoe UI', Arial, sans-serif;
            }
        """)

        about_html = """
        <h3 style="color: #2c3e50;">What is SpatialPin Converter?</h3>
        <p>SpatialPin Converter is an advanced, user-friendly desktop application designed to bridge the gap between raw, messy field data and clean, standardized geographic formats.</p>

        <h3 style="color: #2c3e50;">Core Features & Capabilities:</h3>
        <ul style="margin-top: 0px; padding-left: 20px;">
            <li style="margin-bottom: 6px;"><b>Smart Coordinate Parsing:</b> Automatically translates any format—including Decimal Degrees (DD), Degrees Minutes (DDM), and Degrees Minutes Seconds (DMS)—into standard numerical outputs.</li>
            <li style="margin-bottom: 6px;"><b>Fuzzy Typo Correction:</b> Intelligently corrects common human data-entry errors (e.g., mistaking the letter 'O' for '0', or 'I' for '1') often found in raw datasets.</li>
            <li style="margin-bottom: 6px;"><b>Dual Operating Modes:</b> Offers high-speed Bulk File Processing (.xlsx, .csv, .txt) and a fast Single Point testing interface with map previews.</li>
            <li style="margin-bottom: 6px;"><b>Advanced GIS Processing:</b> Instantly calculates projected Universal Transverse Mercator (UTM) coordinates alongside standard WGS84 outputs.</li>
            <li style="margin-bottom: 6px;"><b>Reverse Geocoding:</b> Fetches real-world physical address contexts from standard coordinates, providing immediate human-readable locations alongside numerical data.</li>
            <li style="margin-bottom: 6px;"><b>GIS Native Exports:</b> Directly exports cleaned coordinate points to standard formats like <b>.kml</b> (Google Earth) and <b>.geojson</b> for immediate spatial mapping.</li>
            <li style="margin-bottom: 6px;"><b>Robust Quality Control:</b> Features intelligent header detection, safely ignores empty rows without crashing, and generates detailed txt error logs for points it cannot resolve.</li>
        </ul>

        <p><i>Built for professionals in GIS, surveying, and spatial analysis to maximize productivity and ensure high data fidelity.</i></p>
        """
        text_edit.setHtml(about_html)
        layout.addWidget(text_edit)

        layout.addSpacing(10)

        # Action close button
        btn_close = QPushButton("Close")
        btn_close.setMinimumWidth(100)
        btn_close.clicked.connect(dialog.accept)

        # Center the close button
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        btn_layout.addWidget(btn_close)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        dialog.exec()

    # MODIFICATION 1: Modernize the Contact Us Menu
    def menu_contact(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Contact Us")
        dialog.setFixedSize(380, 220)
        layout = QVBoxLayout(dialog)

        # Header Styling
        lbl_title = QLabel("Get in Touch")
        lbl_title.setFont(QFont("Segoe UI", 16, QFont.Weight.Bold))
        lbl_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(lbl_title)

        # Information Text
        lbl_info = QLabel(
            "For support, bug reports, or feature requests:\n\nEmail: Mohamed--Ashraf@outlook.com"
        )
        lbl_info.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lbl_info.setStyleSheet("font-size: 13px; color: #2d3436;")
        layout.addWidget(lbl_info)

        layout.addSpacing(15)

        # Highlighted LinkedIn Button
        btn_linkedin = QPushButton("🔗 Open LinkedIn Profile")
        btn_linkedin.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_linkedin.setStyleSheet("""
            QPushButton {
                background-color: #0077b5;
                color: white;
                border-radius: 6px;
                padding: 10px;
                font-weight: bold;
                font-size: 13px;
                border: none;
            }
            QPushButton:hover {
                background-color: #005582;
            }
            QPushButton:pressed {
                background-color: #004471;
            }
        """)
        btn_linkedin.clicked.connect(
            lambda: webbrowser.open("https://www.linkedin.com/in/mohamed---ashraf/")
        )
        layout.addWidget(btn_linkedin)

        layout.addSpacing(5)

        # Standard Close Button
        btn_close = QPushButton("Close")
        btn_close.clicked.connect(dialog.accept)
        layout.addWidget(btn_close)

        dialog.exec()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = GeoCoordApp()
    window.show()
    sys.exit(app.exec())
