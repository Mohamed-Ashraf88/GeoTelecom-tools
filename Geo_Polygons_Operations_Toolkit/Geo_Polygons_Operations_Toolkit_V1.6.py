"""
Geo-Polygons Operations Toolkit
Author: Mohamed Ashraf
Version: 1.6 (Minimalist)
Description: A highly optimized, compact CustomTkinter tool for KML/KMZ geospatial polygon operations.
"""

import sys
import os
import ctypes
import tempfile
import pickle
import datetime
import threading
import queue
import webbrowser

# --- COMPILED EXECUTABLE ENVIRONMENT FIXES ---
# These ensure pyproj and fiona find their internal databases when bundled by Nuitka/PyInstaller
try:
    import pyproj

    # Nuitka specific data directory resolution
    proj_dir = os.path.join(os.path.dirname(pyproj.__file__), "proj_dir")
    if not os.path.exists(proj_dir):
        import pyproj.datadir

        proj_dir = pyproj.datadir.get_data_dir()

    os.environ["PROJ_LIB"] = proj_dir
    os.environ["PROJ_DATA"] = proj_dir
    pyproj.datadir.set_data_dir(proj_dir)
    pyproj.CRS("EPSG:4326")  # Pre-initialize to prevent threading crashes
except Exception:
    pass

try:
    import fiona

    gdal_dir = os.path.join(os.path.dirname(fiona.__file__), "gdal_data")
    if os.path.exists(gdal_dir):
        os.environ["GDAL_DATA"] = gdal_dir
except Exception:
    pass
# ---------------------------------------------

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageDraw, ImageTk

# Geospatial Libraries
import pandas as pd
import geopandas as gpd
from shapely.geometry import shape, mapping
import fiona
import simplekml
import warnings

warnings.filterwarnings("ignore")  # Ignore pyproj CRS deprecation warnings

# Enable KML support (Handle multiple fiona versions securely and Case-Sensitivity)
try:
    fiona.supported_drivers["KML"] = "rw"
    fiona.supported_drivers["LIBKML"] = "rw"  # Exact string expected by GDAL
    fiona.supported_drivers["libkml"] = "rw"
    fiona.supported_drivers["KMZ"] = "rw"
except AttributeError:
    pass

try:
    fiona.drvsupport.supported_drivers["KML"] = "rw"
    fiona.drvsupport.supported_drivers["LIBKML"] = "rw"  # Exact string expected by GDAL
    fiona.drvsupport.supported_drivers["libkml"] = "rw"
    fiona.drvsupport.supported_drivers["KMZ"] = "rw"
except AttributeError:
    pass

# Global Configuration
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

TOOL_NAME = "Geo-Polygons Operations Toolkit"
TOOL_VERSION = "1.6"
AUTHOR_NAME = "Mohamed Ashraf"

# Tooltip Alternative
TOOL_DESCRIPTIONS = {
    "Combine": "Fuses overlapping polygons (or multiple files) into a single, larger polygon.",
    "Intersection": "Creates a new polygon from only the overlapping area shared by inputs.",
    "Difference": "Subtracts the overlapping area of Polygon B from Polygon A.",
    "SymDifference": "Creates a new polygon keeping unique parts, deleting shared parts.",
    "SpatialJoin": "Attaches attributes of Layer B to Layer A based on spatial location.",
    "Contains": "Checks if Polygon A entirely encompasses Polygon B.",
    "Intersects": "Checks if Polygon A and B share any space.",
    "Buffer": "Expands or shrinks the polygon by a specified precise distance.",
    "Simplify": "Reduces the number of vertices while maintaining the original shape.",
    "ConvexHull": "Creates the smallest convex polygon enclosing the original shape.",
    "BoundingBox": "Calculates the smallest rectangular box containing the polygon.",
    "Dissolve": "Merges polygons that share the same value in the selected attribute column.",
    "Area": "Calculates exact area using localized UTM projection.",
    "Perimeter": "Measures the precise total length of the polygon's boundary.",
    "Centroid": "Calculates the absolute center of mass (point) of the polygon.",
}


class DropZoneLabel(ctk.CTkLabel):
    """Custom Label that acts as a drag-and-drop zone (if supported by system)."""

    def __init__(self, master, text, **kwargs):
        super().__init__(master, text=text, **kwargs)


class AboutDialog(ctk.CTkToplevel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.title(f"About {TOOL_NAME}")
        self.geometry("500x400")

        scroll = ctk.CTkScrollableFrame(self)
        scroll.pack(fill="both", expand=True, padx=10, pady=10)

        info_text = f"""🌍 {TOOL_NAME} v{TOOL_VERSION}
Author: {AUTHOR_NAME}
--------------------------------------------------

Overview:
The Geo-Polygons Operations Toolkit is a minimalist, high-performance desktop application designed for advanced geospatial analysis. Engineered with a multi-threaded architecture, it allows users to safely load and process thousands of polygons simultaneously without freezing the interface.

Key Features:
- Batch Processing: Combine 1,000+ files easily into pure, unified MultiGeometries.
- True Metric Engine: Highly accurate Area, Perimeter, and Buffer using dynamic localized UTM projections.
- 3D Google Earth Integration: Export results with Extrusion tags or directly launch to Earth.
- HTML Reports: Instantly generate professional summary reports of your spatial calculations.
"""

        info_label = ctk.CTkLabel(
            scroll, text=info_text, justify="left", wraplength=450
        )
        info_label.pack(padx=10, pady=10, anchor="w")

        close_btn = ctk.CTkButton(self, text="Close", command=self.destroy)
        close_btn.pack(pady=10)


class ContactDialog(ctk.CTkToplevel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.title("Contact Us")
        self.geometry("400x200")

        info_text = """Contact the Developer

Email: Mohamed--Ashraf@outlook.com
LinkedIn: https://www.linkedin.com/in/mohamed---ashraf/
"""
        label = ctk.CTkLabel(self, text=info_text, justify="center")
        label.pack(expand=True, padx=20, pady=20)

        close_btn = ctk.CTkButton(self, text="Close", command=self.destroy)
        close_btn.pack(pady=10)


class FileLoadWorker(threading.Thread):
    def __init__(self, paths, layer_id, queue):
        super().__init__(daemon=True)
        self.paths = paths
        self.layer_id = layer_id
        self.queue = queue

    def run(self):
        try:
            if not self.paths:
                self.queue.put(("load_finished", None, self.layer_id))
                return

            gdfs = []
            total = len(self.paths)
            errors = []

            for i, p in enumerate(self.paths):
                try:
                    gdf = gpd.read_file(p)
                    gdf["source_file"] = os.path.basename(p)
                    if gdf.crs is None:
                        gdf.set_crs(epsg=4326, inplace=True)
                    elif gdf.crs.to_epsg() != 4326:
                        gdf = gdf.to_crs(epsg=4326)
                    gdfs.append(gdf)
                except Exception as e:
                    errors.append(f"{os.path.basename(p)}: {str(e)}")

                self.queue.put(
                    (
                        "progress",
                        int(((i + 1) / total) * 100),
                        f"Loading {os.path.basename(p)}",
                    )
                )

            if errors:
                error_msg = "\n".join(errors[:3])
                if len(errors) > 3:
                    error_msg += f"\n...and {len(errors)-3} more."
                self.queue.put(("error", f"Failed to load file(s):\n{error_msg}"))

            if gdfs:
                combined = pd.concat(gdfs, ignore_index=True)
                if not isinstance(combined, gpd.GeoDataFrame):
                    combined = gpd.GeoDataFrame(combined, geometry="geometry")
                if combined.crs is None:
                    combined.set_crs(epsg=4326, inplace=True, allow_override=True)

                self.queue.put(("load_finished", combined, self.layer_id))
            else:
                self.queue.put(("load_finished", None, self.layer_id))
        except Exception as e:
            self.queue.put(("error", f"Critical load error: {str(e)}"))


class SpatialWorker(threading.Thread):
    def __init__(self, op, gdf_a, gdf_b, extra_val, extra_unit, queue):
        super().__init__(daemon=True)
        self.op = op
        self.gdf_a = gdf_a.copy() if gdf_a is not None else None
        self.gdf_b = gdf_b.copy() if gdf_b is not None else None
        self.extra_val = extra_val
        self.extra_unit = extra_unit
        self.queue = queue

    def run(self):
        try:
            self.queue.put(("progress", 10, f"Executing mathematical {self.op}..."))
            gdf_a = self.gdf_a
            gdf_b = self.gdf_b

            # --- Boolean Operations ---
            if self.op == "Combine":
                if gdf_b is not None and not gdf_b.empty:
                    self.queue.put(
                        (
                            "progress",
                            50,
                            "Fusing all geometries in Layer A and Layer B into one...",
                        )
                    )
                    combined_gdf = pd.concat([gdf_a, gdf_b], ignore_index=True)
                    geom = combined_gdf.geometry.unary_union
                else:
                    self.queue.put(
                        ("progress", 50, "Fusing all geometries in Layer A into one...")
                    )
                    geom = gdf_a.geometry.unary_union

                # Unary union returns exactly ONE geometry footprint, resolving the sub-polygon issue
                res = gpd.GeoDataFrame(geometry=[geom], crs=gdf_a.crs)
                self.queue.put(("calc_gdf_finished", res))

            elif self.op == "Intersection":
                self.queue.put(
                    ("calc_gdf_finished", gpd.overlay(gdf_a, gdf_b, how="intersection"))
                )
            elif self.op == "Difference":
                self.queue.put(
                    ("calc_gdf_finished", gpd.overlay(gdf_a, gdf_b, how="difference"))
                )
            elif self.op == "SymDifference":
                self.queue.put(
                    (
                        "calc_gdf_finished",
                        gpd.overlay(gdf_a, gdf_b, how="symmetric_difference"),
                    )
                )
            elif self.op == "SpatialJoin":
                self.queue.put(
                    (
                        "calc_gdf_finished",
                        gpd.sjoin(gdf_a, gdf_b, how="left", predicate="intersects"),
                    )
                )

            # --- Topology Operations ---
            elif self.op == "Contains":
                self.queue.put(
                    (
                        "calc_text_finished",
                        str(
                            gdf_a.geometry.unary_union.contains(
                                gdf_b.geometry.unary_union
                            )
                        ),
                    )
                )
            elif self.op == "Intersects":
                self.queue.put(
                    (
                        "calc_text_finished",
                        str(
                            gdf_a.geometry.unary_union.intersects(
                                gdf_b.geometry.unary_union
                            )
                        ),
                    )
                )

            # --- Transform Operations ---
            elif self.op == "Buffer":
                if self.extra_unit == "Degrees":
                    gdf_a.geometry = gdf_a.geometry.buffer(self.extra_val)
                else:
                    utm_crs = gdf_a.estimate_utm_crs()
                    utm_gdf = gdf_a.to_crs(utm_crs)
                    dist = (
                        self.extra_val
                        if self.extra_unit == "Meters"
                        else self.extra_val * 1000
                    )
                    utm_gdf.geometry = utm_gdf.geometry.buffer(dist)
                    gdf_a = utm_gdf.to_crs(epsg=4326)
                self.queue.put(("calc_gdf_finished", gdf_a))

            elif self.op == "Simplify":
                gdf_a.geometry = gdf_a.geometry.simplify(0.01)
                self.queue.put(("calc_gdf_finished", gdf_a))
            elif self.op == "ConvexHull":
                gdf_a.geometry = gdf_a.geometry.convex_hull
                self.queue.put(("calc_gdf_finished", gdf_a))
            elif self.op == "BoundingBox":
                gdf_a.geometry = gdf_a.geometry.envelope
                self.queue.put(("calc_gdf_finished", gdf_a))
            elif self.op == "Dissolve":
                if self.extra_val:
                    self.queue.put(
                        (
                            "calc_gdf_finished",
                            gdf_a.dissolve(by=self.extra_val).reset_index(),
                        )
                    )
                else:
                    self.queue.put(
                        ("error", "No valid attribute selected to dissolve by.")
                    )

            # --- Measure Operations ---
            elif self.op == "Area":
                utm_gdf = gdf_a.to_crs(gdf_a.estimate_utm_crs())
                area_sqm = utm_gdf.geometry.area.sum()
                self.queue.put(
                    (
                        "calc_text_finished",
                        f"Total Area:\n{area_sqm:,.2f} sq meters\n{area_sqm/1e6:,.4f} sq km\n{area_sqm*0.000247105:,.2f} acres",
                    )
                )

            elif self.op == "Perimeter":
                utm_gdf = gdf_a.to_crs(gdf_a.estimate_utm_crs())
                perim_m = utm_gdf.geometry.length.sum()
                self.queue.put(
                    (
                        "calc_text_finished",
                        f"Total Perimeter:\n{perim_m:,.2f} meters\n{perim_m/1000:,.4f} km\n{perim_m*0.000621371:,.2f} miles",
                    )
                )

            elif self.op == "Centroid":
                centroids = gdf_a.geometry.centroid
                self.queue.put(
                    (
                        "calc_gdf_finished",
                        gpd.GeoDataFrame(geometry=centroids, crs=gdf_a.crs),
                    )
                )

        except Exception as e:
            self.queue.put(("error", str(e)))


class GeoPolygonToolkit(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"{TOOL_NAME} v{TOOL_VERSION}")
        self.geometry("450x850")

        # Set dynamic in-memory icon
        self.iconphoto(False, self.generate_custom_icon())

        # State Data
        self.gdf_a = None
        self.gdf_b = None
        self.result_gdf = None
        self.undo_stack = []

        # Threading communication queue
        self.queue = queue.Queue()

        self.init_menu()
        self.init_ui()
        self.poll_queue()

    def generate_custom_icon(self):
        img = Image.new("RGBA", (64, 64), (255, 255, 255, 0))
        draw = ImageDraw.Draw(img)
        draw.polygon(
            [(10, 10), (45, 15), (40, 50), (5, 40)],
            fill=(51, 136, 255, 220),
            outline=(30, 92, 191, 255),
        )
        draw.polygon(
            [(25, 25), (60, 20), (55, 55), (20, 60)],
            fill=(0, 255, 0, 200),
            outline=(0, 179, 0, 255),
        )
        return ImageTk.PhotoImage(img)

    def init_menu(self):
        menubar = tk.Menu(self)

        filemenu = tk.Menu(menubar, tearoff=0)
        filemenu.add_command(
            label="Load Workspace (.geotool)", command=self.load_workspace
        )
        filemenu.add_command(
            label="Save Workspace (.geotool)", command=self.save_workspace
        )
        filemenu.add_separator()
        filemenu.add_command(label="Exit", command=self.quit)
        menubar.add_cascade(label="File", menu=filemenu)

        editmenu = tk.Menu(menubar, tearoff=0)
        editmenu.add_command(label="Undo Last Action", command=self.undo_action)
        menubar.add_cascade(label="Edit", menu=editmenu)

        helpmenu = tk.Menu(menubar, tearoff=0)
        helpmenu.add_command(label="About the tool", command=self.show_about)
        helpmenu.add_command(label="Contact us", command=self.show_contact)
        menubar.add_cascade(label="Help", menu=helpmenu)

        self.config(menu=menubar)

    def init_ui(self):
        # --- MAIN PANEL (Only Controls) ---
        self.main_panel = ctk.CTkScrollableFrame(self, corner_radius=0)
        self.main_panel.pack(fill="both", expand=True)

        title_lbl = ctk.CTkLabel(
            self.main_panel,
            text=f"🌍 {TOOL_NAME}",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color="#4facfe",
        )
        title_lbl.pack(pady=(10, 20))

        # 1. Inputs
        input_frame = ctk.CTkFrame(self.main_panel)
        input_frame.pack(fill="x", padx=10, pady=5)

        lbl_in = ctk.CTkLabel(
            input_frame, text="Input Layers", font=ctk.CTkFont(weight="bold")
        )
        lbl_in.pack(pady=5)

        row_a = ctk.CTkFrame(input_frame, fg_color="transparent")
        row_a.pack(fill="x", padx=5, pady=5)
        self.btn_load_a = ctk.CTkButton(
            row_a,
            text="📁 Click to Browse Layer A",
            height=40,
            command=lambda: self.browse_file("A"),
        )
        self.btn_load_a.pack(fill="x", expand=True)
        self.lbl_file_a = ctk.CTkLabel(
            input_frame, text="None loaded", text_color="gray"
        )
        self.lbl_file_a.pack()

        row_b = ctk.CTkFrame(input_frame, fg_color="transparent")
        row_b.pack(fill="x", padx=5, pady=5)
        self.btn_load_b = ctk.CTkButton(
            row_b,
            text="📁 Click to Browse Layer B",
            height=40,
            command=lambda: self.browse_file("B"),
        )
        self.btn_load_b.pack(fill="x", expand=True)
        self.lbl_file_b = ctk.CTkLabel(
            input_frame, text="None loaded", text_color="gray"
        )
        self.lbl_file_b.pack()

        self.lbl_desc = ctk.CTkLabel(
            self.main_panel,
            text="Hover over an operation to see its description.",
            text_color="#a6e3a1",
            wraplength=350,
        )
        self.lbl_desc.pack(pady=(10, 0))

        # 2. Operations Tabs
        self.tabs = ctk.CTkTabview(self.main_panel)
        self.tabs.pack(fill="x", padx=10, pady=10)

        tab_bool = self.tabs.add("Boolean (A+B)")
        for op in [
            "Combine",
            "Intersection",
            "Difference",
            "SymDifference",
            "SpatialJoin",
        ]:
            btn = ctk.CTkButton(
                tab_bool, text=op, command=lambda o=op: self.start_worker(o)
            )
            btn.pack(pady=5, fill="x")
            btn.bind(
                "<Enter>",
                lambda e, desc=TOOL_DESCRIPTIONS[op]: self.lbl_desc.configure(
                    text=desc
                ),
            )

        tab_topo = self.tabs.add("Topology (A+B)")
        for op in ["Contains", "Intersects"]:
            btn = ctk.CTkButton(
                tab_topo, text=op, command=lambda o=op: self.start_worker(o)
            )
            btn.pack(pady=5, fill="x")
            btn.bind(
                "<Enter>",
                lambda e, desc=TOOL_DESCRIPTIONS[op]: self.lbl_desc.configure(
                    text=desc
                ),
            )

        tab_trans = self.tabs.add("Transform (A)")
        buf_frame = ctk.CTkFrame(tab_trans, fg_color="transparent")
        buf_frame.pack(fill="x", pady=5)
        self.btn_buffer = ctk.CTkButton(
            buf_frame,
            text="Buffer",
            width=100,
            command=lambda: self.start_worker("Buffer"),
        )
        self.btn_buffer.pack(side="left", padx=5)
        self.btn_buffer.bind(
            "<Enter>",
            lambda e: self.lbl_desc.configure(text=TOOL_DESCRIPTIONS["Buffer"]),
        )

        self.buffer_dist = ctk.CTkEntry(buf_frame, width=60)
        self.buffer_dist.insert(0, "50")
        self.buffer_dist.pack(side="left", padx=5)
        self.buffer_units = ctk.CTkComboBox(
            buf_frame, values=["Meters", "Kilometers", "Degrees"], width=100
        )
        self.buffer_units.pack(side="left", padx=5)

        dis_frame = ctk.CTkFrame(tab_trans, fg_color="transparent")
        dis_frame.pack(fill="x", pady=5)
        self.btn_dissolve = ctk.CTkButton(
            dis_frame,
            text="Dissolve By",
            width=100,
            command=lambda: self.start_worker("Dissolve"),
        )
        self.btn_dissolve.pack(side="left", padx=5)
        self.btn_dissolve.bind(
            "<Enter>",
            lambda e: self.lbl_desc.configure(text=TOOL_DESCRIPTIONS["Dissolve"]),
        )
        self.dissolve_field = ctk.CTkComboBox(dis_frame, values=[""], width=170)
        self.dissolve_field.pack(side="left", padx=5)

        for op in ["Simplify", "ConvexHull", "BoundingBox"]:
            btn = ctk.CTkButton(
                tab_trans, text=op, command=lambda o=op: self.start_worker(o)
            )
            btn.pack(pady=5, fill="x")
            btn.bind(
                "<Enter>",
                lambda e, desc=TOOL_DESCRIPTIONS[op]: self.lbl_desc.configure(
                    text=desc
                ),
            )

        tab_meas = self.tabs.add("Measure (A)")
        for op in ["Area", "Perimeter", "Centroid"]:
            btn = ctk.CTkButton(
                tab_meas, text=op, command=lambda o=op: self.start_worker(o)
            )
            btn.pack(pady=5, fill="x")
            btn.bind(
                "<Enter>",
                lambda e, desc=TOOL_DESCRIPTIONS[op]: self.lbl_desc.configure(
                    text=desc
                ),
            )

        # 3. Console & Export
        out_frame = ctk.CTkFrame(self.main_panel)
        out_frame.pack(fill="x", padx=10, pady=10)

        lbl_out = ctk.CTkLabel(
            out_frame, text="Console & Output", font=ctk.CTkFont(weight="bold")
        )
        lbl_out.pack(pady=5)

        self.console = ctk.CTkTextbox(out_frame, height=100)
        self.console.pack(fill="x", padx=5, pady=5)

        chk_frame = ctk.CTkFrame(out_frame, fg_color="transparent")
        chk_frame.pack(fill="x", padx=5)
        self.chk_extrude = ctk.CTkCheckBox(chk_frame, text="Extrude 3D")
        self.chk_extrude.pack(side="left", padx=5)
        self.chk_tessellate = ctk.CTkCheckBox(chk_frame, text="Tessellate")
        self.chk_tessellate.pack(side="left", padx=5)
        self.chk_batch = ctk.CTkCheckBox(chk_frame, text="Batch Export")
        self.chk_batch.pack(side="left", padx=5)

        self.progress_bar = ctk.CTkProgressBar(out_frame)
        self.progress_bar.set(0)
        self.progress_bar.pack(fill="x", padx=5, pady=10)
        self.progress_bar.pack_forget()

        btn_row1 = ctk.CTkFrame(out_frame, fg_color="transparent")
        btn_row1.pack(fill="x", pady=5)
        btn_export = ctk.CTkButton(
            btn_row1,
            text="💾 Export File",
            fg_color="#2e7d32",
            hover_color="#1b5e20",
            command=self.export_file,
        )
        btn_export.pack(side="left", fill="x", expand=True, padx=5)
        btn_earth = ctk.CTkButton(
            btn_row1,
            text="🌍 Open in GE",
            fg_color="#1565c0",
            hover_color="#0d47a1",
            command=self.open_in_google_earth,
        )
        btn_earth.pack(side="right", fill="x", expand=True, padx=5)

        btn_row2 = ctk.CTkFrame(out_frame, fg_color="transparent")
        btn_row2.pack(fill="x", pady=5)
        btn_pdf = ctk.CTkButton(
            btn_row2,
            text="📄 Gen Report (HTML)",
            fg_color="#8e24aa",
            hover_color="#6a1b9a",
            command=self.generate_html_report,
        )
        btn_pdf.pack(fill="x", expand=True, padx=5)

        # Footer
        footer_lbl = ctk.CTkLabel(
            self.main_panel,
            text=f"Author: {AUTHOR_NAME} | Version: {TOOL_VERSION}",
            font=ctk.CTkFont(size=10),
            text_color="gray",
        )
        footer_lbl.pack(pady=10)

    def log(self, text):
        self.console.configure(state="normal")
        self.console.insert("end", f"> {text}\n")
        self.console.see("end")
        self.console.configure(state="disabled")

    # --- POLLING QUEUE FOR THREADS ---
    def poll_queue(self):
        try:
            while True:
                msg = self.queue.get_nowait()
                if msg[0] == "progress":
                    self.progress_bar.set(msg[1] / 100.0)
                    self.log(msg[2])
                elif msg[0] == "load_finished":
                    self.on_load_finished(msg[1], msg[2])
                elif msg[0] == "calc_gdf_finished":
                    self.on_worker_success_gdf(msg[1])
                elif msg[0] == "calc_text_finished":
                    self.on_worker_success_text(msg[1])
                elif msg[0] == "error":
                    self.on_worker_error(msg[1])
        except queue.Empty:
            pass
        finally:
            self.after(100, self.poll_queue)

    # --- DIALOGS ---
    def show_about(self):
        AboutDialog(self)

    def show_contact(self):
        ContactDialog(self)

    # --- STATE MANAGEMENT ---
    def save_state(self):
        state = {
            "gdf_a": self.gdf_a.copy() if self.gdf_a is not None else None,
            "gdf_b": self.gdf_b.copy() if self.gdf_b is not None else None,
            "result": self.result_gdf.copy() if self.result_gdf is not None else None,
        }
        self.undo_stack.append(state)
        if len(self.undo_stack) > 10:
            self.undo_stack.pop(0)

    def undo_action(self):
        if not self.undo_stack:
            self.log("Undo Stack is empty.")
            return
        state = self.undo_stack.pop()
        self.gdf_a = state["gdf_a"]
        self.gdf_b = state["gdf_b"]
        self.result_gdf = state["result"]
        self._update_dissolve_field(self.gdf_a)
        self.log("Action undone successfully.")

    def save_workspace(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".geotool", filetypes=[("GeoTool Workspace", "*.geotool")]
        )
        if not path:
            return
        try:
            ws_data = {"gdf_a": self.gdf_a, "gdf_b": self.gdf_b, "res": self.result_gdf}
            with open(path, "wb") as f:
                pickle.dump(ws_data, f)
            self.log("Workspace saved successfully.")
        except Exception as e:
            self.log(f"Failed to save workspace: {e}")

    def load_workspace(self):
        path = filedialog.askopenfilename(
            filetypes=[("GeoTool Workspace", "*.geotool")]
        )
        if not path:
            return
        try:
            with open(path, "rb") as f:
                ws_data = pickle.load(f)
            self.save_state()
            self.gdf_a = ws_data.get("gdf_a")
            self.gdf_b = ws_data.get("gdf_b")
            self.result_gdf = ws_data.get("res")
            self._update_dissolve_field(self.gdf_a)
            self.log("Workspace loaded successfully.")
        except Exception as e:
            self.log(f"Failed to load workspace: {e}")

    def _update_dissolve_field(self, gdf):
        """Updates the dissolve dropdown with columns from Layer A"""
        if gdf is not None and not gdf.empty:
            cols = [c for c in gdf.columns if c != "geometry"]
            if cols:
                self.dissolve_field.configure(values=cols)
                self.dissolve_field.set(cols[0])
            else:
                self.dissolve_field.configure(values=[""])
                self.dissolve_field.set("")
        else:
            self.dissolve_field.configure(values=[""])
            self.dissolve_field.set("")

    # --- LOADING ---
    def browse_file(self, layer_id):
        file_paths = filedialog.askopenfilenames(
            filetypes=[
                ("Spatial Files", "*.kml *.kmz *.shp *.geojson *.json *.gpkg"),
                ("All Files", "*.*"),
            ]
        )
        if file_paths:
            self.process_selected_files(file_paths, layer_id)

    def process_selected_files(self, file_paths, layer_id):
        self.save_state()
        count = len(file_paths)
        file_names = (
            f"{count} file(s) selected"
            if count > 1
            else os.path.basename(file_paths[0])
        )

        lbl = self.lbl_file_a if layer_id == "A" else self.lbl_file_b
        lbl.configure(text=f"{file_names}")
        self.log(f"Layer {layer_id}: {count} file(s) queued for loading.")

        self.progress_bar.pack(fill="x", padx=5, pady=10)
        self.progress_bar.set(0)

        # Protect against clicking start worker while a background load is in progress
        self.tabs.configure(state="disabled")

        FileLoadWorker(file_paths, layer_id, self.queue).start()

    def on_load_finished(self, gdf, layer_id):
        if layer_id == "A":
            self.gdf_a = gdf
            self._update_dissolve_field(self.gdf_a)
        else:
            self.gdf_b = gdf

        self.progress_bar.pack_forget()
        self.tabs.configure(state="normal")
        self.log(f"Layer {layer_id} loaded securely.")

    # --- WORKERS ---
    def start_worker(self, op):
        if self.gdf_a is None or self.gdf_a.empty:
            messagebox.showwarning("Missing Data", "Layer A is required.")
            return
        needs_b = [
            "Intersection",
            "Difference",
            "SymDifference",
            "SpatialJoin",
            "Contains",
            "Intersects",
        ]
        if op in needs_b and (self.gdf_b is None or self.gdf_b.empty):
            messagebox.showwarning("Missing Data", f"{op} requires Layer B.")
            return

        self.save_state()
        self.log(f"--- Starting {op} ---")
        self.progress_bar.pack(fill="x", padx=5, pady=10)
        self.progress_bar.set(0)

        self.tabs.configure(state="disabled")

        val = None
        unit = None
        if op == "Buffer":
            try:
                val = float(self.buffer_dist.get())
            except ValueError:
                val = 50.0
            unit = self.buffer_units.get()
        elif op == "Dissolve":
            val = self.dissolve_field.get()

        SpatialWorker(op, self.gdf_a, self.gdf_b, val, unit, self.queue).start()

    def on_worker_success_gdf(self, res_gdf):
        self.result_gdf = res_gdf
        self.log(f"Success! Output has {len(res_gdf)} feature(s).")
        self.finish_worker()

    def on_worker_success_text(self, text):
        self.log(f"Result:\n{text}")
        self.finish_worker()

    def on_worker_error(self, err):
        if self.undo_stack:
            self.undo_stack.pop()
        self.log(f"ERROR: {err}")
        messagebox.showerror("Operation Failed", str(err))
        self.finish_worker()

    def finish_worker(self):
        self.progress_bar.pack_forget()
        self.tabs.configure(state="normal")

    # --- EXPORT & REPORTS ---
    def _export_gdf_to_kml(self, gdf, file_path, custom_placemark_name=None):
        kml = simplekml.Kml()
        extrude = 1 if self.chk_extrude.get() else 0
        tessellate = 1 if self.chk_tessellate.get() else 0

        for _, row in gdf.iterrows():
            geom = row.geometry
            if geom is None or geom.is_empty:
                continue

            if custom_placemark_name:
                name = custom_placemark_name
            else:
                # Check if this is a true combined unified polygon
                name = "Combined Feature" if len(gdf) == 1 else "Exported Poly"
                cols = [c for c in gdf.columns if c != "geometry"]
                if cols:
                    val = row[cols[0]]
                    if pd.notnull(val):
                        name = str(val)

            if geom.geom_type == "Polygon":
                pol = kml.newpolygon(
                    name=name, outerboundaryis=list(geom.exterior.coords)
                )
                if len(geom.interiors) > 0:
                    pol.innerboundaryis = [list(ring.coords) for ring in geom.interiors]
                pol.extrude = extrude
                pol.tessellate = tessellate

            elif geom.geom_type == "MultiPolygon":
                # FIX: Add MultiPolygon support as ONE MultiGeometry so it doesn't break into pieces
                multi = kml.newmultigeometry(name=name)
                for part in geom.geoms:
                    pol = multi.newpolygon(outerboundaryis=list(part.exterior.coords))
                    if len(part.interiors) > 0:
                        pol.innerboundaryis = [
                            list(ring.coords) for ring in part.interiors
                        ]
                    pol.extrude = extrude
                    pol.tessellate = tessellate

            elif geom.geom_type == "Point":
                pnt = kml.newpoint(name=name, coords=[list(geom.coords)[0]])
                pnt.extrude = extrude

        kml.save(file_path)

    def export_file(self):
        if self.result_gdf is None:
            messagebox.showwarning(
                "No Result", "No operation result available to export."
            )
            return

        if self.chk_batch.get():
            folder_path = filedialog.askdirectory(
                title="Select Folder for Batch Export"
            )
            if not folder_path:
                return
            try:
                for i, row in self.result_gdf.iterrows():
                    single_gdf = gpd.GeoDataFrame([row], crs=self.result_gdf.crs)
                    custom_name = f"output_feature_{i}"
                    path = os.path.join(folder_path, f"{custom_name}.kml")
                    self._export_gdf_to_kml(
                        single_gdf, path, custom_placemark_name=custom_name
                    )
                self.log(f"Batch exported {len(self.result_gdf)} files.")
                messagebox.showinfo("Success", "Batch export complete!")
            except Exception as e:
                self.log(f"Batch Export Failed: {str(e)}")
        else:
            file_path = filedialog.asksaveasfilename(
                defaultextension=".shp",
                filetypes=[
                    ("Shapefile", "*.shp"),
                    ("KML", "*.kml"),
                    ("GeoJSON", "*.geojson"),
                    ("GeoPackage", "*.gpkg"),
                ],
            )
            if not file_path:
                return
            try:
                ext = os.path.splitext(file_path)[1].lower()
                if ext == ".kml":
                    custom_name = os.path.splitext(os.path.basename(file_path))[0]
                    self._export_gdf_to_kml(
                        self.result_gdf, file_path, custom_placemark_name=custom_name
                    )
                else:
                    drv = {
                        ".kmz": "KMZ",
                        ".shp": "ESRI Shapefile",
                        ".geojson": "GeoJSON",
                        ".gpkg": "GPKG",
                    }.get(ext, "ESRI Shapefile")
                    self.result_gdf.to_file(file_path, driver=drv)
                self.log(f"Saved to: {file_path}")
                messagebox.showinfo("Success", "File saved successfully!")
            except Exception as e:
                self.log(f"Export Failed: {str(e)}")

    def open_in_google_earth(self):
        if self.result_gdf is None:
            messagebox.showwarning("No Result", "No result available to view.")
            return
        if os.name != "nt":
            messagebox.showwarning(
                "OS Not Supported", "Direct launch optimized for Windows."
            )
            return
        try:
            self.log("Launching Google Earth...")
            fd, path = tempfile.mkstemp(suffix=".kml")
            os.close(fd)
            self._export_gdf_to_kml(
                self.result_gdf, path, custom_placemark_name="Preview Feature"
            )
            os.startfile(path)
        except Exception as e:
            self.log(f"Failed to launch GE: {str(e)}")

    def generate_html_report(self):
        if self.result_gdf is None:
            messagebox.showwarning(
                "No Result", "Run an operation first to generate a report."
            )
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".html",
            initialfile="GeoReport.html",
            filetypes=[("HTML Report", "*.html")],
        )
        if not file_path:
            return

        try:
            utm_gdf = self.result_gdf.to_crs(self.result_gdf.estimate_utm_crs())
            area_sqkm = utm_gdf.geometry.area.sum() / 1e6
            perim_km = utm_gdf.geometry.length.sum() / 1000
            feature_count = len(self.result_gdf)
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            df = pd.DataFrame(self.result_gdf.drop(columns="geometry")).head(15)
            table_html = df.to_html(index=False, justify="left", classes="data-table")

            html = f"""
            <html>
            <head>
                <style>
                    body {{ font-family: 'Segoe UI', Arial, sans-serif; margin: 40px; color: #333; }}
                    h1 {{ color: #2e7d32; text-align: center; border-bottom: 2px solid #2e7d32; padding-bottom: 10px; }}
                    .metrics {{ background-color: #f5f5f5; padding: 20px; border-radius: 8px; margin: 20px 0; }}
                    table.data-table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
                    .data-table th, .data-table td {{ border: 1px solid #ddd; padding: 12px; text-align: left; }}
                    .data-table th {{ background-color: #2e7d32; color: white; }}
                    .data-table tr:nth-child(even) {{ background-color: #f9f9f9; }}
                </style>
            </head>
            <body>
                <h1>Geo-Polygons Operations Report</h1>
                <p><b>Generated On:</b> {timestamp}</p>
                <p><b>Tool Version:</b> {TOOL_VERSION} by {AUTHOR_NAME}</p>
                
                <div class="metrics">
                    <h3>Result Metrics</h3>
                    <p><b>Total Features:</b> {feature_count} polygons</p>
                    <p><b>Total Area:</b> {area_sqkm:,.4f} Square Kilometers</p>
                    <p><b>Total Perimeter:</b> {perim_km:,.4f} Kilometers</p>
                    <p><b>Coordinate Reference System:</b> EPSG:4326 (WGS84)</p>
                </div>
                
                <h3>Attribute Data Sample (Top 15 Rows)</h3>
                {table_html}
            </body>
            </html>
            """

            with open(file_path, "w", encoding="utf-8") as f:
                f.write(html)

            webbrowser.open("file://" + os.path.abspath(file_path))
            self.log(f"HTML Report generated and opened.")
        except Exception as e:
            self.log(f"Report Error: {str(e)}")


if __name__ == "__main__":
    if os.name == "nt":
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
            myappid = f"mohamed.ashraf.geopolygon.{TOOL_VERSION}"
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        except Exception:
            pass

    app = GeoPolygonToolkit()
    app.mainloop()
