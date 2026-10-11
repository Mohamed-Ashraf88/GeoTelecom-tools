"""Sector Mapper: single-file Python desktop application.

Requires Python 3.10+ with Tkinter and openpyxl.
Install dependency:  py -m pip install "openpyxl>=3.1.5,<4"
Run:                 py SectorMapper.py

Optional Nuitka build with the embedded Windows EXE icon:
py -m pip install Nuitka ordered-set zstandard
py SectorMapper.py --build-exe
The helper generates its build icon temporarily and embeds it in the EXE.
No separate image file is needed to run or distribute the app.

Import either one sector per Excel row or one site with multiple azimuth columns.
Combined color groups such as 1800+2600 can be split into separate sectors.
Visual layer ordering uses heights above ground; Google Earth terrain and camera
angle can affect overlaps. Folder order is configured independently.

The combined preview uses all enabled groups at one sample north azimuth,
with proportional radii, selected beamwidths, alpha blending and layer order.
On Export map, select sector popup fields and optionally enable site-name pins
with their own popup fields. Clear all fields to hide a popup. Site pins group
exported sectors by site name and coordinates, combining distinct source values.
About, Quick Guide and Contact / Support are in the sidebar. Pin & label
appearance is under Export map > Optional site pins. Custom marker images are
inside KMZ exports; plain KML exports use the adjacent SectorMapper-assets folder.
"""
from __future__ import annotations

from dataclasses import dataclass, asdict
from pathlib import Path
import csv
import math
import os
import sys
import struct
import zlib
import subprocess
import ctypes
from functools import lru_cache
import re
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
import openpyxl
import json
import html
import hashlib
import webbrowser
from dataclasses import asdict
from collections import Counter
from concurrent.futures import ThreadPoolExecutor
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser


NS = 'http://www.opengis.net/kml/2.2'
ET.register_namespace('', NS)
PALETTE = ['#2563EB', '#EF4444', '#22C55E', '#F59E0B', '#A855F7', '#06B6D4', '#EC4899', '#64748B']

APP_USER_MODEL_ID = 'SectorMapper.Desktop'
APP_VERSION = '1.0.0'
APP_BUILD = '20261003.1'
APP_AUTHOR = 'Mohamed Ashraf'
SUPPORT_EMAIL = 'Mohamed--Ashraf@outlook.com'
SUPPORT_LINKEDIN = 'https://www.linkedin.com/in/mohamed---ashraf/'
PIN_SHAPES = ('Pin', 'Circle', 'Diamond', 'Square', 'Star')
LABEL_STYLES = ('Site name', 'Uppercase name', 'Name + groups', 'Name + sector count', 'Hidden')


def set_windows_app_identity():
    """Set the process taskbar grouping identity before creating any windows.

    https://learn.microsoft.com/windows/win32/api/shobjidl_core/nf-shobjidl_core-setcurrentprocessexplicitappusermodelid
    No registry changes or external files are involved.
    """
    if sys.platform != 'win32':
        return False
    try:
        function = ctypes.WinDLL('shell32', use_last_error=True).SetCurrentProcessExplicitAppUserModelID
        function.argtypes = [ctypes.c_wchar_p]
        function.restype = ctypes.c_long
        return function(APP_USER_MODEL_ID) >= 0
    except (AttributeError, OSError):
        # Window icons still work if the Windows shell API is unavailable.
        return False


@lru_cache(maxsize=8)
def application_icon_png(size):
    """Render the built-in sector mark to a PNG in memory using only stdlib.

    A navy rounded square, a blue bearing ring, a north-facing sector and a
    white site marker. Supersampling keeps small title-bar icons legible.
    """
    if size not in (16, 24, 32, 48, 64, 128, 256):
        raise ValueError('Unsupported icon size.')
    pixels = bytearray()
    samples = 3
    for y in range(size):
        pixels.append(0)  # PNG scanline filter: None
        for x in range(size):
            sums = [0, 0, 0]
            covered = 0
            for sy in range(samples):
                for sx in range(samples):
                    px = (x + (sx + 0.5)/samples)/size
                    py = (y + (sy + 0.5)/samples)/size
                    # Rounded rectangle with transparent outer corners.
                    qx = max(abs(px-0.5)-0.31, 0)
                    qy = max(abs(py-0.5)-0.31, 0)
                    if qx*qx + qy*qy > 0.17**2:
                        continue
                    color = (16, 28, 50)
                    dx, dy = px-0.5, py-0.53
                    radius = math.hypot(dx, dy)
                    if 0.335 <= radius <= 0.375:
                        color = (59, 130, 246)
                    if radius < 0.29 and dy < 0 and abs(dx) <= -dy*0.9:
                        color = (96, 165, 250)
                    if radius < 0.065:
                        color = (255, 255, 255)
                    for channel in range(3):
                        sums[channel] += color[channel]
                    covered += 1
            if covered:
                pixels.extend([round(v/covered) for v in sums] + [round(255*covered/(samples*samples))])
            else:
                pixels.extend((0, 0, 0, 0))
    def chunk(kind, data):
        return struct.pack('!I', len(data)) + kind + data + struct.pack('!I', zlib.crc32(kind+data) & 0xffffffff)
    return (b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('!IIBBBBB', size, size, 8, 6, 0, 0, 0))
            + chunk(b'IDAT', zlib.compress(bytes(pixels), 9)) + chunk(b'IEND', b''))


def application_icon_ico():
    """Package the same mark at seven sizes as an in-memory Windows ICO."""
    sizes = (16, 24, 32, 48, 64, 128, 256)
    frames = [application_icon_png(size) for size in sizes]
    header = struct.pack('<HHH', 0, 1, len(sizes))
    entries, offset = [], 6 + 16*len(sizes)
    for size, frame in zip(sizes, frames):
        entries.append(struct.pack('<BBBBHHII', size % 256, size % 256, 0, 0, 1, 32, len(frame), offset))
        offset += len(frame)
    return header + b''.join(entries) + b''.join(frames)


def build_windows_executable():
    """Nuitka needs a build-time icon path; generate it without a source asset.

    https://nuitka.net/user-documentation/user-manual.html
    The temporary file is removed after compilation. Explorer reads the icon
    embedded in the EXE; the running GUI uses its in-memory PNGs.
    """
    if sys.platform != 'win32':
        raise SystemExit('Build the Windows executable on Windows.')
    if '__compiled__' in globals():
        raise SystemExit('Run --build-exe from the Python source file.')
    with tempfile.TemporaryDirectory(prefix='sector-mapper-icon-') as directory:
        icon = Path(directory)/'SectorMapper.ico'
        icon.write_bytes(application_icon_ico())
        command = [sys.executable, '-m', 'nuitka', '--mode=onefile', '--enable-plugin=tk-inter',
                   '--windows-console-mode=disable', f'--windows-icon-from-ico={icon}',
                   '--product-name=Sector Mapper', '--file-description=RF sector map converter',
                   f'--product-version={APP_VERSION}', f'--file-version={APP_VERSION}',
                   '--output-dir=build', '--output-filename=SectorMapper.exe', str(Path(__file__).resolve())]
        return subprocess.run(command, check=False).returncode


@dataclass
class Mapping:
    site: str
    latitude: str
    longitude: str
    azimuths: list[str]
    group: str
    sector: str = ''
    split_groups: bool = True
    separator: str = '+'


@dataclass
class Sector:
    site: str
    name: str
    latitude: float
    longitude: float
    azimuth: float
    group: str
    row: int
    source: dict


@dataclass
class GroupStyle:
    group: str
    color: str = '#2563EB'
    radius: float = 500
    beamwidth: float = 65
    opacity: float = 85
    enabled: bool = True
    folder_order: int = 0

    def validate(self):
        if not re.fullmatch(r'#[0-9a-fA-F]{6}', self.color):
            raise ValueError(f'{self.group}: color must be #RRGGBB.')
        number(self.radius, 'Radius (meters)', 0.01, 100000)
        number(self.beamwidth, 'Beamwidth (degrees)', 0.01, 360)
        number(self.opacity, 'Opacity (%)', 0, 100)
        number(self.folder_order, 'Folder order', 0, 1000000)


@dataclass
class ExportOptions:
    title: str = 'RF sectors'
    altitude_mode: str = 'stacked'
    base_height: float = 5
    layer_gap: float = 2
    outlines: bool = True
    site_folders: bool = True
    sector_fields: list[str] | None = None
    site_pins: bool = False
    site_fields: list[str] | None = None
    pin_shape: str = 'Pin'
    pin_color: str = '#F59E0B'
    pin_size: float = 1.0
    pin_opacity: float = 100
    label_style: str = 'Site name'
    label_color: str = '#FFFFFF'
    label_size: float = 0.9
    label_opacity: float = 100

    def validate(self):
        if self.altitude_mode not in ('stacked', 'ground'):
            raise ValueError('Unknown altitude mode.')
        number(self.base_height, 'Base height', 0, 100000)
        number(self.layer_gap, 'Layer separation', 0.01, 10000)
        if not isinstance(self.site_pins, bool):
            raise ValueError('Site pins must be enabled or disabled.')
        for selected in (self.sector_fields, self.site_fields):
            if selected is not None and (not isinstance(selected, list) or any(not isinstance(f, str) for f in selected)):
                raise ValueError('Popup fields must be a list of field names.')
        if self.pin_shape not in PIN_SHAPES or self.label_style not in LABEL_STYLES:
            raise ValueError('Choose a supported pin and label style.')
        for name, value in [('Pin color', self.pin_color), ('Label color', self.label_color)]:
            if not re.fullmatch(r'#[0-9a-fA-F]{6}', value):
                raise ValueError(f'{name} must be #RRGGBB.')
        number(self.pin_size, 'Pin size', 0.1, 4)
        number(self.label_size, 'Label size', 0.1, 4)
        number(self.pin_opacity, 'Pin opacity', 0, 100)
        number(self.label_opacity, 'Label opacity', 0, 100)


@lru_cache(maxsize=5)
def map_pin_png(shape):
    """White marker silhouettes: KML IconStyle supplies tint and opacity."""
    if shape not in PIN_SHAPES:
        raise ValueError('Unknown pin shape.')
    size = 48
    rows = bytearray()
    star = [(0.5 + (0.41 if i % 2 == 0 else 0.18)*math.sin(i*math.pi/5),
             0.5 - (0.41 if i % 2 == 0 else 0.18)*math.cos(i*math.pi/5)) for i in range(10)]
    def inside(x, y):
        dx, dy = x-0.5, y-0.5
        if shape == 'Circle':
            return dx*dx+dy*dy <= 0.37**2
        if shape == 'Square':
            return max(abs(dx), abs(dy)) <= 0.33
        if shape == 'Diamond':
            return abs(dx)+abs(dy) <= 0.43
        if shape == 'Pin':
            head = dx*dx+(y-0.34)**2 <= 0.28**2
            tail = 0.34 <= y <= 0.95 and abs(dx) <= (0.95-y)*0.46
            hole = dx*dx+(y-0.34)**2 < 0.105**2
            return (head or tail) and not hole
        hit = False
        for (ax, ay), (bx, by) in zip(star, star[1:]+star[:1]):
            if (ay > y) != (by > y) and x < (bx-ax)*(y-ay)/(by-ay)+ax:
                hit = not hit
        return hit
    for y in range(size):
        rows.append(0)
        for x in range(size):
            coverage = sum(inside((x+(sx+0.5)/3)/size, (y+(sy+0.5)/3)/size) for sy in range(3) for sx in range(3))
            rows.extend((255, 255, 255, round(255*coverage/9)))
    def chunk(kind, data):
        return struct.pack('!I', len(data))+kind+data+struct.pack('!I', zlib.crc32(kind+data) & 0xffffffff)
    return b'\x89PNG\r\n\x1a\n'+chunk(b'IHDR', struct.pack('!IIBBBBB', size, size, 8, 6, 0, 0, 0))+chunk(b'IDAT', zlib.compress(bytes(rows)))+chunk(b'IEND', b'')


def site_label(members, style):
    site = members[0].site
    if style == 'Uppercase name':
        return site.upper()
    if style == 'Name + groups':
        return site + ' · ' + ', '.join(dict.fromkeys(s.group for s in members))
    if style == 'Name + sector count':
        return f'{site} · {len(members)} sectors'
    return site


SECTOR_FIELDS = ['Site', 'Sector', 'Latitude', 'Longitude', 'Color group', 'Azimuth (deg)',
                 'Radius (m)', 'Beamwidth (deg)', 'Excel row']
SITE_FIELDS = ['Site', 'Latitude', 'Longitude', 'Sector count', 'Color groups']


def sector_metadata(sector, style):
    values = dict(zip(SECTOR_FIELDS, [sector.site, sector.name, sector.latitude, sector.longitude,
                                     sector.group, sector.azimuth, style.radius, style.beamwidth, sector.row]))
    values.update({f'Excel: {key}': cell_text(value) for key, value in sector.source.items()})
    return values


def site_metadata(members):
    first = members[0]
    values = dict(zip(SITE_FIELDS, [first.site, first.latitude, first.longitude, len(members),
                                   ', '.join(dict.fromkeys(s.group for s in members))]))
    distinct = {}
    for sector in members:
        for key, value in sector.source.items():
            text = cell_text(value)
            if text:
                distinct.setdefault(key, {})[text] = None
    values.update({f'Excel: {key}': ' | '.join(items) for key, items in distinct.items()})
    return values


def popup_content(metadata, selected):
    """Escape Excel text for HTML, and emit only the selected fields."""
    names = metadata.keys() if selected is None else dict.fromkeys(selected)
    values = {key: metadata[key] for key in names if key in metadata}
    rows = ''.join('<tr><th align="left" style="padding:6px 12px 6px 0">'
                   + html.escape(key) + '</th><td style="padding:6px 0">'
                   + html.escape(cell_text(value)) + '</td></tr>' for key, value in values.items())
    return '<table style="font-family:Arial;font-size:13px">' + rows + '</table>', values


def add_balloon_style(style_node, selected):
    balloon = element(style_node, 'BalloonStyle')
    element(balloon, 'text', '$[description]')
    if selected == []:
        element(balloon, 'displayMode', 'hide')


def add_selected_data(mark, metadata):
    if metadata:
        ext = element(mark, 'ExtendedData')
        for key, value in metadata.items():
            element(element(ext, 'Data', name=key), 'value', value)


def combined_preview_rgb(styles, width=240, height=190):
    """Composite aligned sample sectors from bottom to top with real alpha.

    All groups share north azimuth and the same meters-to-pixels scale. Upper
    opaque layers can cover lower layers exactly as in a top-down stacked view.
    """
    active = [s for s in styles if s.enabled]
    radius = min(width, height)/2 - 23
    maximum = max((s.radius for s in active), default=1)
    cx, cy = width/2, height/2 + 4
    layers = [(s.radius/maximum*radius, s.beamwidth/2, s.opacity/100,
               tuple(int(s.color[i:i+2], 16) for i in (1, 3, 5))) for s in active]
    result = bytearray()
    for y in range(height):
        for x in range(width):
            dx, dy = x+0.5-cx, y+0.5-cy
            distance = math.hypot(dx, dy)
            angle = abs(math.degrees(math.atan2(dx, -dy)))
            rgb = (247, 249, 252)
            if abs(dx) < 0.6 or abs(dy) < 0.6 or any(abs(distance-radius*f) < 0.5 for f in (1/3, 2/3, 1)):
                rgb = (218, 227, 239)
            for extent, halfbeam, alpha, color in layers:
                if distance <= extent and angle <= halfbeam:
                    rgb = tuple(round(c*alpha+b*(1-alpha)) for c, b in zip(color, rgb))
            result.extend(rgb)
    return bytes(result)


def cell_text(value):
    if value is None:
        return ''
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def number(value, name, minimum, maximum):
    try:
        if isinstance(value, bool) or not cell_text(value):
            raise ValueError
        result = float(value)
    except (ValueError, TypeError):
        raise ValueError(f'{name}: expected a number, got {value!r}.') from None
    if not math.isfinite(result) or not minimum <= result <= maximum:
        raise ValueError(f'{name}: must be between {minimum:g} and {maximum:g}; got {value!r}.')
    return result


def sheet_names(path):
    book = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        return book.sheetnames
    finally:
        book.close()


def read_sheet(path, sheet, header_row=1):
    if not isinstance(header_row, int) or header_row < 1:
        raise ValueError('Header row must be a positive integer.')
    book = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = book[sheet]
        # Some generators write unreliable used-range metadata.
        ws.reset_dimensions()
        values = ws.iter_rows(min_row=header_row, values_only=True)
        raw = next(values, ())
        last = max((i for i, v in enumerate(raw) if cell_text(v)), default=-1)
        headers = [cell_text(v) for v in raw[:last + 1]]
        if not headers or any(not h for h in headers):
            raise ValueError('Header row is empty or contains an unnamed column.')
        if len(set(headers)) != len(headers):
            raise ValueError('Column headers must be unique. Rename duplicate headers in Excel.')
        rows = []
        for rownum, row in enumerate(values, header_row + 1):
            if any(cell_text(v) for v in row):
                rows.append((rownum, {h: row[i] if i < len(row) else None for i, h in enumerate(headers)}))
        return headers, rows
    finally:
        book.close()


def guess_mapping(headers):
    clean = {re.sub(r'[^a-z0-9]', '', h.lower()): h for h in headers}
    def pick(*names):
        return next((clean[n] for n in names if n in clean), '')
    return Mapping(pick('sitename', 'site', 'siteid'), pick('latitude', 'lat'),
                   pick('longitude', 'lon', 'long', 'lng'),
                   [h for h in headers if re.match(r'^(azimuth|az)[ _-]*\d*$', h, re.I)],
                   pick('band', 'technology'), pick('sectorname', 'sector', 'cellname'))


def normalize(headers, rows, mapping):
    required = [mapping.site, mapping.latitude, mapping.longitude, mapping.group] + mapping.azimuths
    if not mapping.azimuths or any(h not in headers for h in required):
        raise ValueError('Select site, latitude, longitude, color group and at least one azimuth column.')
    if mapping.sector and mapping.sector not in headers:
        raise ValueError('The sector-name column is missing.')
    if mapping.split_groups and not mapping.separator:
        raise ValueError('Enter a separator for combined groups.')
    sectors, issues = [], []
    for rownum, row in rows:
        try:
            site = cell_text(row[mapping.site])
            if not site:
                raise ValueError('Site name is empty.')
            lat = number(row[mapping.latitude], 'Latitude', -90, 90)
            lon = number(row[mapping.longitude], 'Longitude', -180, 180)
            raw_group = cell_text(row[mapping.group])
            groups = raw_group.split(mapping.separator) if mapping.split_groups else [raw_group]
            groups = list(dict.fromkeys(g.strip() for g in groups))
            if any(not g for g in groups):
                raise ValueError('Color group is empty or contains an empty separated value.')
            present = [(h, row[h]) for h in mapping.azimuths if cell_text(row[h])]
            if not present:
                raise ValueError('No azimuth is present.')
        except ValueError as error:
            issues.append({'row': rownum, 'column': 'row', 'value': '', 'error': str(error)})
            continue
        for col, value in present:
            try:
                azimuth = number(value, col, 0, 360) % 360
            except ValueError as error:
                issues.append({'row': rownum, 'column': col, 'value': cell_text(value), 'error': str(error)})
                continue
            name = cell_text(row.get(mapping.sector)) if mapping.sector else ''
            name = name or f'{site} / {col}'
            if mapping.sector and len(mapping.azimuths) > 1:
                name += f' / {col}'
            for group in groups:
                sectors.append(Sector(site, name, lat, lon, azimuth, group, rownum, dict(row)))
    return sectors, issues


def default_styles(sectors):
    groups = list(dict.fromkeys(s.group for s in sectors))
    return [GroupStyle(g, PALETTE[i % len(PALETTE)], folder_order=i) for i, g in enumerate(groups)]


def destination(lat, lon, bearing, meters):
    """Vincenty's direct solution on the WGS84 ellipsoid (distance in meters)."""
    a, f = 6378137.0, 1 / 298.257223563
    b = (1 - f) * a
    alpha = math.radians(bearing)
    phi = math.radians(lat)
    u = math.atan2((1 - f) * math.sin(phi), math.cos(phi))
    su, cu = math.sin(u), math.cos(u)
    sa, ca = math.sin(alpha), math.cos(alpha)
    sigma1 = math.atan2(su, cu * ca)
    sin_alpha = cu * sa
    cos2_alpha = 1 - sin_alpha ** 2
    u2 = cos2_alpha * (a*a - b*b) / (b*b)
    A = 1 + u2 / 16384 * (4096 + u2 * (-768 + u2 * (320 - 175*u2)))
    B = u2 / 1024 * (256 + u2 * (-128 + u2 * (74 - 47*u2)))
    sigma = meters / (b * A)
    for _ in range(100):
        cs, ss = math.cos(sigma), math.sin(sigma)
        c2 = math.cos(2 * sigma1 + sigma)
        delta = B * ss * (c2 + B / 4 * (cs * (-1 + 2*c2*c2) - B / 6 * c2 * (-3 + 4*ss*ss) * (-3 + 4*c2*c2)))
        new = meters / (b*A) + delta
        if abs(new - sigma) < 1e-12:
            sigma = new
            break
        sigma = new
    else:
        raise ValueError('Geodesic calculation did not converge.')
    cs, ss = math.cos(sigma), math.sin(sigma)
    tmp = su*ss - cu*cs*ca
    phi2 = math.atan2(su*cs + cu*ss*ca, (1-f)*math.sqrt(sin_alpha*sin_alpha + tmp*tmp))
    lam = math.atan2(ss*sa, cu*cs - su*ss*ca)
    C = f/16*cos2_alpha*(4+f*(4-3*cos2_alpha))
    c2 = math.cos(2*sigma1 + sigma)
    L = lam - (1-C)*f*sin_alpha*(sigma + C*ss*(c2+C*cs*(-1+2*c2*c2)))
    return ((lon + math.degrees(L) + 180) % 360 - 180, math.degrees(phi2))


def sector_ring(sector, style):
    # Clockwise bearings are traversed in reverse to produce CCW KML outer rings.
    n = max(8, math.ceil(style.beamwidth / 2))
    arc = [destination(sector.latitude, sector.longitude,
                       sector.azimuth + style.beamwidth/2 - style.beamwidth*i/n, style.radius)
           for i in range(n+1)]
    if style.beamwidth == 360:
        arc[-1] = arc[0]
        return arc
    center = (sector.longitude, sector.latitude)
    return [center] + arc + [center]


def kml_color(color, opacity):
    return f'{round(opacity * 255 / 100):02x}{color[5:7]}{color[3:5]}{color[1:3]}'.lower()


def element(parent, tag, text=None, **attrs):
    node = ET.SubElement(parent, f'{{{NS}}}{tag}', attrs)
    if text is not None:
        node.text = str(text)
    return node


def export_kml(path, sectors, styles, options, progress=None):
    """Stream complete group folders to a temporary file, then atomically replace output."""
    options.validate()
    path = Path(path)
    if path.suffix.lower() not in ('.kml', '.kmz'):
        raise ValueError('Choose a .kml or .kmz filename.')
    if len({s.group for s in styles}) != len(styles):
        raise ValueError('Duplicate group styles.')
    for style in styles:
        style.validate()
    known = {s.group for s in styles}
    if any(s.group not in known for s in sectors):
        raise ValueError('Some sectors have no group style. Rebuild the groups.')
    active = [s for s in styles if s.enabled]
    data = defaultdict(list)
    for sector in sectors:
        data[sector.group].append(sector)
    total = sum(len(data[s.group]) for s in active)
    if total == 0:
        raise ValueError('No sectors are enabled for export.')
    pin_image = map_pin_png(options.pin_shape) if options.site_pins else None
    pin_href = 'files/site-marker.png'
    if pin_image and path.suffix.lower() == '.kml':
        assets = path.parent/'SectorMapper-assets'
        assets.mkdir(exist_ok=True)
        asset = assets/(hashlib.sha256(pin_image).hexdigest()[:20]+'.png')
        if not asset.exists():
            asset.write_bytes(pin_image)
        pin_href = 'SectorMapper-assets/'+asset.name
    heights = {s.group: options.base_height + i * options.layer_gap for i, s in enumerate(active)}
    ids = {s.group: f'group_{i}' for i, s in enumerate(styles)}
    fd, temp_name = tempfile.mkstemp(dir=path.parent, suffix='.tmp')
    os.close(fd)
    done = 0
    try:
        with open(temp_name, 'wb') as raw:
            archive = zipfile.ZipFile(raw, 'w', zipfile.ZIP_DEFLATED) if path.suffix.lower() == '.kmz' else None
            stream = archive.open('doc.kml', 'w') if archive else raw
            try:
                stream.write(f'<?xml version="1.0" encoding="UTF-8"?><kml xmlns="{NS}"><Document>'.encode())
                title = ET.Element(f'{{{NS}}}name'); title.text = options.title
                stream.write(ET.tostring(title, encoding='utf-8'))
                for style in active:
                    node = ET.Element(f'{{{NS}}}Style', id=ids[style.group])
                    line = element(node, 'LineStyle')
                    element(line, 'color', kml_color(style.color, 100))
                    element(line, 'width', 1)
                    poly = element(node, 'PolyStyle')
                    element(poly, 'color', kml_color(style.color, style.opacity))
                    element(poly, 'fill', 1)
                    element(poly, 'outline', int(options.outlines))
                    add_balloon_style(node, options.sector_fields)
                    stream.write(ET.tostring(node, encoding='utf-8'))
                if options.site_pins:
                    pin_style = ET.Element(f'{{{NS}}}Style', id='site_pin')
                    icon_style = element(pin_style, 'IconStyle')
                    element(icon_style, 'color', kml_color(options.pin_color, options.pin_opacity))
                    element(icon_style, 'scale', options.pin_size)
                    element(element(icon_style, 'Icon'), 'href', pin_href)
                    element(icon_style, 'hotSpot', x='0.5', y='0.05' if options.pin_shape == 'Pin' else '0.5', xunits='fraction', yunits='fraction')
                    label_style = element(pin_style, 'LabelStyle')
                    element(label_style, 'color', kml_color(options.label_color, options.label_opacity))
                    element(label_style, 'scale', 0 if options.label_style == 'Hidden' else options.label_size)
                    add_balloon_style(pin_style, options.site_fields)
                    stream.write(ET.tostring(pin_style, encoding='utf-8'))
                for style in sorted(active, key=lambda s: s.folder_order):
                    # Folders are streamed too, avoiding a huge in-memory XML tree.
                    stream.write(b'<Folder>')
                    group_name = ET.Element(f'{{{NS}}}name'); group_name.text = style.group
                    stream.write(ET.tostring(group_name, encoding='utf-8'))
                    grouped = defaultdict(list)
                    for s in data[style.group]:
                        grouped[s.site if options.site_folders else ''].append(s)
                    for site, members in grouped.items():
                        if options.site_folders:
                            stream.write(b'<Folder>')
                            site_name = ET.Element(f'{{{NS}}}name'); site_name.text = site
                            stream.write(ET.tostring(site_name, encoding='utf-8'))
                        for sector in members:
                            mark = ET.Element(f'{{{NS}}}Placemark')
                            element(mark, 'name', f'{sector.name} / {sector.group}')
                            description, selected_data = popup_content(sector_metadata(sector, style), options.sector_fields)
                            element(mark, 'description', description)
                            element(mark, 'styleUrl', '#' + ids[style.group])
                            add_selected_data(mark, selected_data)
                            poly = element(mark, 'Polygon')
                            ground = options.altitude_mode == 'ground'
                            element(poly, 'tessellate', int(ground))
                            element(poly, 'altitudeMode', 'clampToGround' if ground else 'relativeToGround')
                            ring = element(element(poly, 'outerBoundaryIs'), 'LinearRing')
                            height = 0 if ground else heights[style.group]
                            element(ring, 'coordinates', ' '.join(f'{x:.9f},{y:.9f},{height:g}' for x, y in sector_ring(sector, style)))
                            stream.write(ET.tostring(mark, encoding='utf-8'))
                            done += 1
                            if progress and (done % 250 == 0 or done == total):
                                progress(done, total)
                        if options.site_folders:
                            stream.write(b'</Folder>')
                    stream.write(b'</Folder>')
                if options.site_pins:
                    sites = defaultdict(list)
                    for style in active:
                        for sector in data[style.group]:
                            sites[(sector.site, sector.latitude, sector.longitude)].append(sector)
                    stream.write(b'<Folder><name>Site pins</name>')
                    for members in sites.values():
                        first = members[0]
                        mark = ET.Element(f'{{{NS}}}Placemark')
                        element(mark, 'name', site_label(members, options.label_style))
                        description, selected_data = popup_content(site_metadata(members), options.site_fields)
                        element(mark, 'description', description)
                        element(mark, 'styleUrl', '#site_pin')
                        add_selected_data(mark, selected_data)
                        point = element(mark, 'Point')
                        element(point, 'altitudeMode', 'clampToGround')
                        element(point, 'coordinates', f'{first.longitude:.9f},{first.latitude:.9f},0')
                        stream.write(ET.tostring(mark, encoding='utf-8'))
                    stream.write(b'</Folder>')
                stream.write(b'</Document></kml>')
            finally:
                if archive:
                    stream.close()
                    if pin_image:
                        archive.writestr(pin_href, pin_image)
                    archive.close()
        os.replace(temp_name, path)
    except BaseException:
        Path(temp_name).unlink(missing_ok=True)
        raise
    return done


def write_issues(path, issues):
    with open(path, 'w', newline='', encoding='utf-8-sig') as stream:
        writer = csv.DictWriter(stream, fieldnames=['row', 'column', 'value', 'error'])
        writer.writeheader()
        # Prevent spreadsheet formula evaluation when the report is opened in Excel.
        for issue in issues:
            writer.writerow({k: "'" + v if isinstance(v, str) and v.startswith(('=', '+', '-', '@')) else v for k, v in issue.items()})


class App(tk.Tk):
    def __init__(self):
        self.windows_identity_set = set_windows_app_identity()
        super().__init__()
        self.title('Sector Mapper')
        self.configure(background='#F3F5F9')
        width, height = min(1380, self.winfo_screenwidth()-60), min(920, self.winfo_screenheight()-90)
        self.geometry(f'{width}x{height}+16+16')
        self.minsize(min(1050, width), min(680, height))
        self.executor = ThreadPoolExecutor(max_workers=1)
        self.busy = False
        self.headers, self.rows, self.sectors, self.issues, self.styles = [], [], [], [], []
        self.import_signature = None
        self.loaded_path = self.loaded_sheet = self.loaded_header = self.edit_index = None
        self.pending_profile = None
        self.swatches = {}
        self._preview_job = None
        self._filling_editor = False
        self.protocol('WM_DELETE_WINDOW', self.close)
        self.setup_theme()
        # Keep references alive; True also makes these the default for child windows.
        self.app_icons = [tk.PhotoImage(master=self, data=application_icon_png(size), format='png')
                          for size in (16, 24, 32, 48, 64, 128)]
        self.app_icon = self.app_icons[2]
        self.iconphoto(True, *self.app_icons)

        shell = ttk.Frame(self, style='Workspace.TFrame')
        shell.pack(fill='both', expand=True)
        sidebar = tk.Frame(shell, bg='#101C32', width=208)
        sidebar.pack(side='left', fill='y'); sidebar.pack_propagate(False)
        brand = tk.Canvas(sidebar, width=168, height=64, bg='#101C32', highlightthickness=0)
        brand.pack(anchor='w', padx=20, pady=(24, 3))
        brand.create_oval(1, 8, 39, 46, outline='#3B82F6', width=2)
        brand.create_arc(6, 13, 34, 41, start=30, extent=120, fill='#60A5FA', outline='')
        brand.create_oval(18, 25, 23, 30, fill='white', outline='')
        brand.create_text(51, 16, text='SECTOR', anchor='w', fill='white', font=('Segoe UI', 12, 'bold'))
        brand.create_text(51, 36, text='MAPPER', anchor='w', fill='#8FA8CF', font=('Segoe UI', 12))
        tk.Label(sidebar, text='RF PLANNING WORKSPACE', bg='#101C32', fg='#7F94B7', font=('Segoe UI', 8, 'bold')).pack(anchor='w', padx=22, pady=(0, 30))
        self.nav_buttons = []
        for i, text in enumerate(['01   Import data', '02   Colors & layers', '03   Export map']):
            btn = ttk.Button(sidebar, text=text, style='Nav.TButton', command=lambda i=i: self.navigate(i))
            btn.pack(fill='x', padx=12, pady=4)
            self.nav_buttons.append(btn)
        lower = tk.Frame(sidebar, bg='#101C32'); lower.pack(side='bottom', fill='x', padx=20, pady=24)
        for text, command in [('Quick Guide', self.show_quick_guide), ('Contact / Support', self.show_support), ('About Sector Mapper', self.show_about)]:
            ttk.Button(lower, text=text, style='Nav.TButton', command=command).pack(fill='x', pady=2)
        tk.Label(lower, text=f'v{APP_VERSION}  ·  {APP_BUILD}', bg='#101C32', fg='#7F94B7', font=('Segoe UI', 8)).pack(anchor='w', pady=(8, 12))
        tk.Frame(lower, height=1, bg='#283650').pack(fill='x', pady=(0, 16))
        tk.Label(lower, text='GOOGLE EARTH PRO', bg='#101C32', fg='#A7B8D2', font=('Segoe UI', 8, 'bold')).pack(anchor='w')
        tk.Label(lower, text='KML / KMZ export\nLocal file processing', bg='#101C32', fg='#7F94B7', font=('Segoe UI', 9), justify='left').pack(anchor='w', pady=(6, 0))

        main = ttk.Frame(shell, style='Workspace.TFrame'); main.pack(side='left', fill='both', expand=True)
        header = ttk.Frame(main, style='Workspace.TFrame', padding=(26, 22, 26, 14)); header.pack(fill='x')
        actions = ttk.Frame(header, style='Workspace.TFrame'); actions.pack(side='right', anchor='n')
        ttk.Button(actions, text='Load settings', command=self.load_profile).pack(side='left', padx=5)
        ttk.Button(actions, text='Save settings', command=self.save_profile).pack(side='left')
        self.page_title = tk.StringVar(value='Import site data')
        self.page_subtitle = tk.StringVar(value='Connect an Excel workbook and map your sector fields.')
        ttk.Label(header, textvariable=self.page_title, style='Title.TLabel').pack(anchor='w')
        ttk.Label(header, textvariable=self.page_subtitle, style='WorkspaceNote.TLabel').pack(anchor='w', pady=(5, 0))
        metrics = ttk.Frame(main, style='Workspace.TFrame', padding=(26, 0, 26, 16)); metrics.pack(fill='x')
        self.metric_vars = {}
        for i, (key, label) in enumerate([('rows', 'SOURCE ROWS'), ('sectors', 'VALID SECTORS'), ('groups', 'COLOR GROUPS'), ('issues', 'DATA ISSUES')]):
            metrics.columnconfigure(i, weight=1, uniform='metrics')
            card = ttk.Frame(metrics, style='Card.TFrame', padding=(16, 10))
            card.grid(row=0, column=i, sticky='ew', padx=(0, 10 if i<3 else 0))
            ttk.Label(card, text=label, style='Eyebrow.TLabel').pack(anchor='w')
            var = tk.StringVar(value='—'); self.metric_vars[key] = var
            ttk.Label(card, textvariable=var, style='Metric.TLabel').pack(anchor='w', pady=(3, 0))

        footer = ttk.Frame(main, style='Card.TFrame', padding=(24, 12))
        footer.pack(side='bottom', fill='x')
        self.next_button = ttk.Button(footer, text='Validate & continue  →', style='Primary.TButton', command=self.build_groups)
        self.next_button.pack(side='right', padx=(16, 0))
        self.status = tk.StringVar(value='Open an Excel workbook to begin.')
        self.status_label = ttk.Label(footer, textvariable=self.status, style='Note.TLabel', wraplength=520)
        self.status_label.pack(side='left', fill='x', expand=True)
        self.progress = ttk.Progressbar(footer, mode='indeterminate', length=80)
        self.status.trace_add('write', lambda *_: self.update_metrics())
        self.tabs = ttk.Notebook(main, style='Workspace.TNotebook')
        self.tabs.pack(fill='both', expand=True, padx=24, pady=(0, 16))
        self.input_tab, self.input_body = self.scroll_page()
        self.group_tab, self.group_body = self.scroll_page()
        self.export_tab, self.export_body = self.scroll_page()
        self.tabs.add(self.input_tab, text='Import')
        self.tabs.add(self.group_tab, text='Layers')
        self.tabs.add(self.export_tab, text='Export')
        self.make_input(); self.make_groups(); self.make_export()
        self.tabs.bind('<<NotebookTabChanged>>', self.page_changed)
        self.bind('<Control-o>', lambda _e: self.open_file() if not self.busy else None)
        self.bind('<Control-s>', lambda _e: self.save_profile() if not self.busy else None)
        self.bind('<MouseWheel>', self.page_wheel, add='+')
        self.page_changed()



    def run_job(self, task, done, label):
        if self.busy:
            messagebox.showinfo('Working', 'Please wait for the current operation to finish.', parent=self)
            return
        self.busy = True
        self.status.set(label)
        self.progress.pack(side='right', padx=(8, 0))
        self.progress.start(12)
        # Keep the UI responsive but prevent configuration changes during a job.
        self.states = []
        def disable(widget):
            for child in widget.winfo_children():
                if isinstance(child, tk.Toplevel):
                    continue
                if isinstance(child, (ttk.Button, ttk.Combobox, ttk.Entry, ttk.Checkbutton, ttk.Radiobutton, ttk.Spinbox, ttk.Treeview)):
                    self.states.append((child, child.state()))
                    child.state(['disabled'])
                elif isinstance(child, tk.Listbox):
                    self.states.append((child, child.cget('state')))
                    child.configure(state='disabled')
                disable(child)
        disable(self)
        future = self.executor.submit(task)
        def poll():
            if not future.done():
                self.after(80, poll)
                return
            self.busy = False
            self.progress.stop()
            self.progress.pack_forget()
            for widget, state in self.states:
                if isinstance(widget, tk.Listbox):
                    widget.configure(state=state)
                else:
                    widget.state(['!disabled'])
                    widget.state(state)
            try:
                done(future.result())
            except Exception as error:
                self.status.set('Operation failed. Review the error and try again.')
                messagebox.showerror('Sector Mapper', str(error), parent=self)
        self.after(80, poll)

    def guard(self, action):
        try:
            return action()
        except (ValueError, OSError, KeyError, TypeError) as error:
            messagebox.showerror('Check settings', str(error), parent=self)

    def make_input(self):
        frame = self.input_body
        source = self.card(frame, 'Excel workbook', 'Choose an .xlsx or .xlsm file. Your source workbook stays unchanged.')
        self.file = tk.StringVar()
        row = ttk.Frame(source); row.pack(fill='x')
        ttk.Entry(row, textvariable=self.file, state='readonly').pack(side='left', fill='x', expand=True)
        ttk.Button(row, text='Browse files…', style='Primary.TButton', command=self.open_file).pack(side='left', padx=(12, 0))
        settings = ttk.Frame(source); settings.pack(fill='x', pady=(14, 0))
        settings.columnconfigure(0, weight=1)
        ttk.Label(settings, text='Worksheet', style='Note.TLabel').grid(row=0, column=0, sticky='w')
        ttk.Label(settings, text='Header row', style='Note.TLabel').grid(row=0, column=1, sticky='w', padx=12)
        self.sheet = tk.StringVar(); self.header = tk.StringVar(value='1')
        self.sheet_combo = ttk.Combobox(settings, textvariable=self.sheet, state='readonly', width=24)
        self.sheet_combo.grid(row=1, column=0, sticky='ew', pady=(5, 0))
        ttk.Spinbox(settings, textvariable=self.header, from_=1, to=1048576, width=7).grid(row=1, column=1, padx=12, pady=(5, 0))
        ttk.Button(settings, text='Read worksheet', command=self.read_current).grid(row=1, column=2, pady=(5, 0))
        self.layout = tk.StringVar(value='sector')
        layout = ttk.Frame(source); layout.pack(fill='x', pady=(12, 0))
        ttk.Label(layout, text='ROW LAYOUT', style='Eyebrow.TLabel').pack(side='left', padx=(0, 18))
        ttk.Radiobutton(layout, text='One sector per row', variable=self.layout, value='sector').pack(side='left')
        ttk.Radiobutton(layout, text='One site per row', variable=self.layout, value='site').pack(side='left', padx=24)

        mapping = self.card(frame, 'Map your columns', 'Match each field to an Excel column. Select all relevant azimuth columns for site rows.')
        grid = ttk.Frame(mapping); grid.pack(fill='x')
        grid.columnconfigure(0, weight=3, uniform='map'); grid.columnconfigure(1, weight=2, uniform='map')
        fields = ttk.Frame(grid); fields.grid(row=0, column=0, sticky='nsew', padx=(0, 22))
        fields.columnconfigure(0, weight=1); fields.columnconfigure(1, weight=1)
        self.fields, self.combos = {}, {}
        items = [('site', 'Site name'), ('sector', 'Sector name · optional'), ('latitude', 'Latitude'), ('longitude', 'Longitude'), ('group', 'Color group')]
        for i, (key, label) in enumerate(items):
            r, c = divmod(i, 2)
            var = tk.StringVar(); self.fields[key] = var
            holder = ttk.Frame(fields); holder.grid(row=r, column=c, sticky='ew', padx=(0, 12 if c==0 else 0), pady=(0, 12))
            ttk.Label(holder, text=label, style='Note.TLabel').pack(anchor='w', pady=(0, 5))
            combo = ttk.Combobox(holder, textvariable=var, state='readonly', width=15)
            combo.pack(fill='x'); self.combos[key] = combo
        az = ttk.Frame(grid); az.grid(row=0, column=1, sticky='nsew')
        ttk.Label(az, text='Azimuth columns', style='Note.TLabel').pack(anchor='w', pady=(0, 5))
        listframe = ttk.Frame(az); listframe.pack(fill='x')
        self.azimuths = tk.Listbox(listframe, selectmode='multiple', exportselection=False, height=6, font=('Segoe UI', 10),
                                 bg='#FBFCFE', fg='#334155', selectbackground='#DBEAFE', selectforeground='#1D4ED8',
                                 highlightbackground='#DCE3EE', highlightcolor='#3B82F6', highlightthickness=1, bd=0, activestyle='none')
        self.azimuths.pack(side='left', fill='x', expand=True)
        scroll = ttk.Scrollbar(listframe, command=self.azimuths.yview); scroll.pack(side='right', fill='y')
        self.azimuths.configure(yscrollcommand=scroll.set)
        ttk.Label(az, text='Click to select or deselect. 0° points north.\nEmpty azimuth cells are skipped.', style='Note.TLabel', wraplength=300).pack(anchor='w', pady=(8, 0))
        ttk.Separator(mapping).pack(fill='x', pady=(2, 10))
        split = ttk.Frame(mapping); split.pack(fill='x')
        self.split = tk.BooleanVar(value=True); self.separator = tk.StringVar(value='+')
        ttk.Checkbutton(split, text='Split combined groups', variable=self.split).pack(side='left')
        ttk.Label(split, text='Separator', style='Note.TLabel').pack(side='left', padx=(24, 8))
        ttk.Entry(split, textvariable=self.separator, width=4).pack(side='left')
        ttk.Label(split, text='1800+2600 → two sectors per azimuth', style='Note.TLabel').pack(side='left', padx=18)

        preview = self.card(frame, 'Data preview', 'First 30 rows of the selected worksheet.')
        table = ttk.Frame(preview); table.pack(fill='x')
        self.preview = ttk.Treeview(table, show='headings', height=5)
        ys = ttk.Scrollbar(table, command=self.preview.yview)
        xs = ttk.Scrollbar(table, orient='horizontal', command=self.preview.xview)
        self.preview.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        self.preview.grid(row=0, column=0, sticky='nsew'); ys.grid(row=0, column=1, sticky='ns'); xs.grid(row=1, column=0, sticky='ew')
        table.columnconfigure(0, weight=1)
        self.preview.tag_configure('alternate', background='#F8FAFD')



    def open_file(self):
        path = filedialog.askopenfilename(filetypes=[('Excel workbook', '*.xlsx *.xlsm')], parent=self)
        if not path:
            return
        def done(names):
            self.file.set(path)
            self.sheet_combo.configure(values=names)
            self.sheet.set(names[0])
            self.header.set('1')
            self.read_current()
        self.run_job(lambda: sheet_names(path), done, 'Reading workbook…')

    def read_current(self):
        if not self.file.get():
            return
        def start():
            path, sheet, header = self.file.get(), self.sheet.get(), int(self.header.get())
            def done(result):
                self.headers, self.rows = result
                self.loaded_path, self.loaded_sheet, self.loaded_header = path, sheet, header
                self.import_signature = None
                self.sectors, self.issues = [], []
                for key, combo in self.combos.items():
                    combo.configure(values=([''] if key == 'sector' else []) + self.headers)
                self.azimuths.delete(0, 'end')
                for h in self.headers:
                    self.azimuths.insert('end', h)
                guess = guess_mapping(self.headers)
                self.set_mapping(guess)
                self.layout.set('site' if len(guess.azimuths) > 1 else 'sector')
                self.preview.delete(*self.preview.get_children())
                self.preview.configure(columns=self.headers)
                for h in self.headers:
                    self.preview.heading(h, text=h); self.preview.column(h, width=125, stretch=False)
                for rownum, row in self.rows[:30]:
                    self.preview.insert('', 'end', values=[row.get(h, '') if row.get(h) is not None else '' for h in self.headers], tags=('alternate',) if len(self.preview.get_children()) % 2 else ())
                self.status.set(f'{len(self.rows):,} rows read. Check the column mapping.')
                self.refresh_popup_fields()
                if self.pending_profile:
                    self.apply_profile_mapping()
            self.run_job(lambda: read_sheet(path, sheet, header), done, 'Reading worksheet…')
        self.guard(start)

    def set_mapping(self, mapping):
        for key, var in self.fields.items():
            var.set(getattr(mapping, key))
        self.azimuths.selection_clear(0, 'end')
        for i, h in enumerate(self.headers):
            if h in mapping.azimuths:
                self.azimuths.selection_set(i)
        self.split.set(mapping.split_groups); self.separator.set(mapping.separator)

    def get_mapping(self):
        azimuths = [self.headers[i] for i in self.azimuths.curselection()]
        if self.layout.get() == 'sector' and len(azimuths) != 1:
            raise ValueError('For one row per sector, select exactly one azimuth column.')
        return Mapping(**{k: v.get() for k, v in self.fields.items()}, azimuths=azimuths,
                       split_groups=self.split.get(), separator=self.separator.get())

    def signature(self):
        return (self.file.get(), self.sheet.get(), self.header.get(), self.layout.get(), asdict(self.get_mapping()))

    def build_groups(self):
        def start():
            if (self.file.get(), self.sheet.get(), int(self.header.get())) != (self.loaded_path, self.loaded_sheet, self.loaded_header):
                raise ValueError('Read the selected worksheet and header row first.')
            mapping = self.get_mapping()
            signature = self.signature()
            def done(result):
                self.sectors, self.issues = result
                existing = {s.group: s for s in self.styles}
                fresh = default_styles(self.sectors)
                # Retain custom order/settings for groups that still exist.
                current_groups = {s.group for s in fresh}
                self.styles = [s for s in self.styles if s.group in current_groups]
                self.styles += [s for s in fresh if s.group not in existing]
                self.import_signature = signature
                self.edit_index = None
                self.refresh_popup_fields()
                self.refresh_groups()
                self.show_issues()
                self.status.set(f'{len(self.sectors):,} sectors · {len(self.styles)} groups · {len(self.issues):,} issues')
                self.tabs.select(self.group_tab)
            self.run_job(lambda: normalize(self.headers, self.rows, mapping), done, 'Validating sectors…')
        self.guard(start)

    def make_groups(self):
        frame = self.group_body
        body = ttk.Frame(frame, style='Workspace.TFrame'); body.pack(fill='both', expand=True)
        body.columnconfigure(0, weight=1); body.columnconfigure(1, weight=0, minsize=312)
        left = ttk.Frame(body, style='Workspace.TFrame'); left.grid(row=0, column=0, sticky='nsew', padx=(0, 14))
        right = ttk.Frame(body, style='Workspace.TFrame', width=312); right.grid(row=0, column=1, sticky='nsew')
        card = self.card(left, 'Layer stack', 'Bottom first, top last. Select a group to edit its appearance.')
        cols = ('layer', 'group', 'color', 'radius', 'beamwidth', 'opacity', 'count', 'folder', 'enabled')
        area = ttk.Frame(card); area.pack(fill='x')
        self.group_tree = ttk.Treeview(area, columns=cols, displaycolumns=('layer', 'group', 'radius', 'beamwidth', 'count', 'enabled'),
                                      show='tree headings', selectmode='browse', height=4)
        self.group_tree.column('#0', width=32, minwidth=32, stretch=False); self.group_tree.heading('#0', text='')
        names = ['#', 'Group', 'Color', 'Meters', 'Beam °', 'Opacity', 'Sectors', 'Folder', 'Export']
        widths = [34, 90, 85, 66, 65, 65, 65, 65, 60]
        for col, width, name in zip(cols, widths, names):
            self.group_tree.heading(col, text=name)
            self.group_tree.column(col, width=width, minwidth=width, stretch=(col=='group'), anchor='w' if col=='group' else 'center')
        self.group_tree.grid(row=0, column=0, sticky='nsew')
        scroll = ttk.Scrollbar(area, command=self.group_tree.yview); scroll.grid(row=0, column=1, sticky='ns')
        horizontal = ttk.Scrollbar(area, orient='horizontal', command=self.group_tree.xview); horizontal.grid(row=1, column=0, sticky='ew')
        area.columnconfigure(0, weight=1)
        self.group_tree.configure(yscrollcommand=scroll.set, xscrollcommand=horizontal.set)
        self.group_tree.tag_configure('alternate', background='#F8FAFD')
        self.group_tree.bind('<<TreeviewSelect>>', self.select_group)
        row = ttk.Frame(card); row.pack(fill='x', pady=(12, 0))
        ttk.Button(row, text='↑  Toward bottom', command=lambda: self.guard(lambda: self.move_group(-1))).pack(side='left', padx=(0, 8))
        ttk.Button(row, text='↓  Toward top', command=lambda: self.guard(lambda: self.move_group(1))).pack(side='left')
        editor = self.card(right, 'Group inspector')
        self.inspector_title = tk.StringVar(value='Select a color group')
        ttk.Label(editor, textvariable=self.inspector_title, style='Accent.TLabel', wraplength=260).pack(anchor='w', pady=(0, 12))
        self.edit = {k: tk.StringVar() for k in ['color', 'radius', 'beamwidth', 'opacity', 'folder_order']}
        self.enabled = tk.BooleanVar(value=True)
        ttk.Label(editor, text='Fill color', style='Note.TLabel').pack(anchor='w')
        color = ttk.Frame(editor); color.pack(fill='x', pady=(5, 12))
        self.color_chip = tk.Canvas(color, width=32, height=32, highlightthickness=1, highlightbackground='#DCE3EE', bg='#E7EDF7')
        self.color_chip.pack(side='left', padx=(0, 7))
        ttk.Entry(color, textvariable=self.edit['color'], width=11).pack(side='left', fill='x', expand=True)
        ttk.Button(color, text='Pick…', command=self.choose_color).pack(side='left', padx=(6, 0))
        grid = ttk.Frame(editor); grid.pack(fill='x')
        grid.columnconfigure(0, weight=1); grid.columnconfigure(1, weight=1)
        for i, (key, text) in enumerate([('radius', 'Radius · meters'), ('beamwidth', 'Beamwidth · degrees'), ('opacity', 'Opacity · %'), ('folder_order', 'Folder rank')]):
            r, c = divmod(i, 2)
            holder = ttk.Frame(grid); holder.grid(row=r, column=c, sticky='ew', padx=(0, 10 if c==0 else 0), pady=(0, 12))
            ttk.Label(holder, text=text, style='Note.TLabel').pack(anchor='w', pady=(0, 5))
            ttk.Entry(holder, textvariable=self.edit[key], width=12).pack(fill='x')
        ttk.Checkbutton(editor, text='Include this group in export', variable=self.enabled).pack(anchor='w')
        ttk.Button(editor, text='Apply changes', style='Primary.TButton', command=lambda: self.guard(self.apply_group)).pack(fill='x', pady=(10, 0))
        folders = self.card(right, 'Folder arrangement', 'Lower ranks appear first in Google Earth’s Places panel.')
        ttk.Button(folders, text='Match ranks to layers', command=self.match_folders).pack(fill='x')
        preview = self.card(left, 'Combined sector preview')
        preview_row = ttk.Frame(preview); preview_row.pack(fill='x')
        self.canvas = tk.Canvas(preview_row, height=190, width=240, bg='#F7F9FC', highlightthickness=0)
        self.canvas.pack(side='left')
        caption = ttk.Frame(preview_row, padding=(16, 0)); caption.pack(side='left', fill='x', expand=True)
        ttk.Label(caption, text='LAYERS · BOTTOM TO TOP', style='Eyebrow.TLabel').pack(anchor='w')
        self.preview_legend = ttk.Frame(caption)
        self.preview_legend.pack(fill='x', pady=(6, 0))
        self.preview_caption = tk.StringVar(value='Choose a group to preview its sector.')
        ttk.Label(caption, textvariable=self.preview_caption, style='Note.TLabel', wraplength=220).pack(anchor='w', pady=(10, 0))
        ttk.Label(caption, text='Same north azimuth, proportional radii.\nUpper layers may hide lower colors.\nTerrain overlaps can differ.', style='Note.TLabel', wraplength=220).pack(anchor='w', pady=(10, 0))
        def arrange_preview(event):
            narrow = event.width < 480
            self.canvas.pack_configure(side='top' if narrow else 'left', anchor='w')
            caption.pack_configure(side='top' if narrow else 'left', anchor='w', pady=(10, 0) if narrow else 0)
        preview_row.bind('<Configure>', arrange_preview)
        for var in self.edit.values():
            var.trace_add('write', lambda *_: self.schedule_preview())
        self.enabled.trace_add('write', lambda *_: self.schedule_preview())
        self.canvas.bind('<Configure>', lambda _e: self.schedule_preview())



    def refresh_groups(self, selected=None):
        counts = Counter(s.group for s in self.sectors)
        self.group_tree.delete(*self.group_tree.get_children())
        self.swatches.clear()
        for i, s in enumerate(self.styles):
            image = tk.PhotoImage(width=14, height=14)
            image.put(s.color, to=(0, 0, 14, 14)); self.swatches[str(i)] = image
            self.group_tree.insert('', 'end', iid=str(i), image=image,
                                   tags=('alternate',) if i%2 else (),
                                   values=[i+1, s.group, s.color, s.radius, s.beamwidth, s.opacity, counts[s.group], s.folder_order, 'Yes' if s.enabled else 'No'])
        if self.styles:
            selected = min(selected or 0, len(self.styles)-1)
            self.edit_index = selected
            self.group_tree.selection_set(str(selected)); self.group_tree.see(str(selected))
            self.fill_editor()
        else:
            self.edit_index = None
            self.inspector_title.set('Select a color group')
            self._filling_editor = True
            for var in self.edit.values():
                var.set('')
            self._filling_editor = False
            self.draw_preview(None)
        self.update_metrics()



    def fill_editor(self):
        if self.edit_index is None:
            return
        self._filling_editor = True
        s = self.styles[self.edit_index]
        for key, var in self.edit.items():
            var.set(str(getattr(s, key)))
        self.enabled.set(s.enabled)
        self.inspector_title.set(s.group)
        self._filling_editor = False
        self.draw_preview(s)



    def select_group(self, _event=None):
        if self.busy:
            return
        selection = self.group_tree.selection()
        if not selection:
            return
        index = int(selection[0])
        if index == self.edit_index:
            return
        if self.edit_index is not None:
            try:
                self.commit_editor()
            except ValueError as error:
                self.group_tree.selection_set(str(self.edit_index))
                messagebox.showerror('Check group settings', str(error), parent=self)
                return
        self.edit_index = index
        self.fill_editor()

    def commit_editor(self):
        if self.edit_index is None or not self.styles:
            return
        old = self.styles[self.edit_index]
        candidate = GroupStyle(old.group, self.edit['color'].get().strip(),
                               float(self.edit['radius'].get()), float(self.edit['beamwidth'].get()),
                               float(self.edit['opacity'].get()), self.enabled.get(), int(self.edit['folder_order'].get()))
        candidate.validate()
        self.styles[self.edit_index] = candidate
        counts = sum(s.group == candidate.group for s in self.sectors)
        self.group_tree.item(str(self.edit_index), values=[self.edit_index+1, candidate.group, candidate.color, candidate.radius, candidate.beamwidth, candidate.opacity, counts, candidate.folder_order, 'Yes' if candidate.enabled else 'No'])
        if str(self.edit_index) in self.swatches:
            self.swatches[str(self.edit_index)].put(candidate.color, to=(0, 0, 14, 14))

    def apply_group(self):
        self.commit_editor()
        if self.edit_index is not None:
            self.draw_preview(self.styles[self.edit_index])

    def draw_preview(self, style=None):
        layers = [style if style is not None and s.group == style.group else s for s in self.styles]
        active = [s for s in layers if s.enabled]
        self.canvas.delete('all')
        width, height = 240, 190
        pixels = combined_preview_rgb(layers, width, height)
        self.combined_preview_image = tk.PhotoImage(master=self, data=f'P6 {width} {height} 255\n'.encode()+pixels, format='PPM')
        self.canvas.create_image(0, 0, image=self.combined_preview_image, anchor='nw')
        self.canvas.create_text(width/2, 10, text='N', fill='#71839C', font=('Segoe UI', 9, 'bold'))
        self.canvas.create_text(width-10, height/2+4, text='E', fill='#94A3B8', font=('Segoe UI', 8))
        self.canvas.create_text(10, height/2+4, text='W', fill='#94A3B8', font=('Segoe UI', 8))
        cx, cy = width/2, height/2+4
        self.canvas.create_oval(cx-3, cy-3, cx+3, cy+3, fill='#172D4E', outline='white')
        for child in self.preview_legend.winfo_children():
            child.destroy()
        for i, group in enumerate(active):
            line = ttk.Frame(self.preview_legend); line.pack(fill='x', pady=2)
            swatch = tk.Canvas(line, width=12, height=12, highlightthickness=0, bg=group.color)
            swatch.pack(side='left', padx=(0, 7))
            ttk.Label(line, text=f'{i+1}. {group.group} · {group.radius:g} m · {group.beamwidth:g}°',
                      style='Note.TLabel', wraplength=210).pack(side='left')
        if style is not None:
            self.color_chip.configure(bg=style.color)
        self.preview_caption.set(f'{len(active)} enabled groups · proportional scale' if active else 'No groups enabled for export.')




    def choose_color(self):
        if self.edit_index is None:
            return
        color = colorchooser.askcolor(color=self.styles[self.edit_index].color, parent=self)[1]
        if color:
            self.edit['color'].set(color)

    def move_group(self, direction):
        self.commit_editor()
        if self.edit_index is None:
            return
        i, j = self.edit_index, self.edit_index + direction
        if 0 <= j < len(self.styles):
            self.styles[i], self.styles[j] = self.styles[j], self.styles[i]
            self.edit_index = j
            self.refresh_groups(j)

    def match_folders(self):
        def apply():
            self.commit_editor()
            for i, style in enumerate(self.styles):
                style.folder_order = i
            self.refresh_groups(self.edit_index)
        self.guard(apply)

    def go_export(self):
        self.apply_group()
        self.tabs.select(self.export_tab)

    def make_export(self):
        frame = self.export_body
        columns = ttk.Frame(frame, style='Workspace.TFrame'); columns.pack(fill='both', expand=True)
        columns.columnconfigure(0, weight=3, uniform='export'); columns.columnconfigure(1, weight=2, uniform='export')
        left = ttk.Frame(columns, style='Workspace.TFrame'); left.grid(row=0, column=0, sticky='nsew', padx=(0, 14))
        right = ttk.Frame(columns, style='Workspace.TFrame'); right.grid(row=0, column=1, sticky='nsew')
        self.title_var = tk.StringVar(value='RF sectors')
        self.mode = tk.StringVar(value='stacked')
        self.base = tk.StringVar(value='5'); self.gap = tk.StringVar(value='2')
        self.outlines = tk.BooleanVar(value=True); self.site_folders = tk.BooleanVar(value=True)
        document = self.card(left, 'Output document')
        ttk.Label(document, text='Name in Google Earth', style='Note.TLabel').pack(anchor='w', pady=(0, 5))
        ttk.Entry(document, textvariable=self.title_var).pack(fill='x')
        row = ttk.Frame(document); row.pack(fill='x', pady=(10, 0))
        ttk.Checkbutton(row, text='Sector outlines', variable=self.outlines).pack(anchor='w')
        ttk.Checkbutton(row, text='Site folders within each color group', variable=self.site_folders).pack(anchor='w')
        layers = self.card(left, 'Layer presentation', 'Choose how sectors sit above the map surface.')
        ttk.Radiobutton(layers, text='Elevated layers', variable=self.mode, value='stacked', command=self.update_altitude_controls).pack(anchor='w')
        ttk.Label(layers, text='Use your bottom-to-top layer order.', style='Note.TLabel').pack(anchor='w', padx=24, pady=(0, 12))
        heights = ttk.Frame(layers); heights.pack(fill='x', padx=24)
        heights.columnconfigure(0, weight=1); heights.columnconfigure(1, weight=1)
        self.altitude_entries = []
        for i, (var, label) in enumerate([(self.base, 'Base height · m'), (self.gap, 'Layer separation · m')]):
            cell = ttk.Frame(heights); cell.grid(row=0, column=i, sticky='ew', padx=(0, 16 if i==0 else 0))
            ttk.Label(cell, text=label, style='Note.TLabel').pack(anchor='w', pady=(0, 5))
            entry = ttk.Entry(cell, textvariable=var, width=12); entry.pack(fill='x'); self.altitude_entries.append(entry)
        ttk.Separator(layers).pack(fill='x', pady=16)
        ttk.Radiobutton(layers, text='Follow terrain', variable=self.mode, value='ground', command=self.update_altitude_controls).pack(anchor='w')
        ttk.Label(layers, text='Clamp to ground. Overlap order is not guaranteed.', style='Note.TLabel', wraplength=340).pack(anchor='w', padx=24)
        note = ttk.Label(layers, text='Elevated sectors can float or shift in tilted views. Terrain and transparency affect overlaps. Folder order alone does not stack polygons.', style='Note.TLabel', wraplength=750)
        note.pack(fill='x', pady=(16, 0)); note.bind('<Configure>', lambda e: note.configure(wraplength=max(100, e.width)))
        review = self.card(right, 'Data review')
        self.issue_summary = tk.StringVar(value='Validate the Excel data before exporting.')
        ttk.Label(review, textvariable=self.issue_summary, style='Accent.TLabel').pack(anchor='w', pady=(0, 10))
        area = ttk.Frame(review); area.pack(fill='x')
        self.issue_text = tk.Text(area, height=5, wrap='word', font=('Consolas', 9), state='disabled', bg='#F7F9FC', fg='#56667F',
                                  bd=0, highlightthickness=0, padx=12, pady=10)
        self.issue_text.pack(side='left', fill='x', expand=True)
        scroll = ttk.Scrollbar(area, command=self.issue_text.yview); scroll.pack(side='right', fill='y')
        self.issue_text.configure(yscrollcommand=scroll.set)
        self.skip_invalid = tk.BooleanVar(value=False)
        ttk.Checkbutton(review, text='Skip invalid entries', variable=self.skip_invalid).pack(anchor='w', pady=(8, 0))
        ttk.Button(review, text='Save issue report…', command=self.save_issues).pack(anchor='w', pady=(8, 0))
        filetype = self.card(right, 'File format', 'KMZ compresses your map into a smaller file. Recommended for large workbooks.')
        ttk.Label(filetype, text='Use “Export KMZ” below, or save\nan uncompressed KML.', style='Note.TLabel').pack(anchor='w')
        ttk.Button(filetype, text='Export KML…', command=lambda: self.export('.kml')).pack(anchor='w', pady=(10, 0))
        self.site_pin_var = tk.BooleanVar(value=False)
        self.popup_selected = {'sector': None, 'site': list(SITE_FIELDS)}
        self.popup_lists = {}
        self.popup_names = {}
        sector_card = self.card(left, 'Sector click information', 'Choose the rows shown when a sector is clicked. Names in the Places panel remain unchanged.')
        self.make_field_selector(sector_card, 'sector')
        pin_card = self.card(right, 'Optional site pins', 'One named pin per site and coordinate pair, using only exported sectors.')
        ttk.Checkbutton(pin_card, text='Add site-name pins', variable=self.site_pin_var).pack(anchor='w', pady=(0, 8))
        defaults = ExportOptions()
        self.pin_controls = {key: tk.StringVar(value=str(getattr(defaults, key))) for key in
                             ('pin_shape', 'pin_color', 'pin_size', 'pin_opacity', 'label_style', 'label_color', 'label_size', 'label_opacity')}
        ttk.Button(pin_card, text='Pin & label appearance…', command=self.show_pin_appearance).pack(fill='x', pady=(0, 10))
        ttk.Label(pin_card, text='KMZ includes marker images. Keep the assets\nfolder beside a KML export.', style='Note.TLabel', wraplength=300).pack(anchor='w', pady=(0, 12))
        ttk.Label(pin_card, text='Choose the pin’s click information.\nDifferent source values are combined.', style='Note.TLabel').pack(anchor='w', pady=(0, 8))
        self.make_field_selector(pin_card, 'site')
        self.refresh_popup_fields()

    def make_field_selector(self, parent, kind):
        area = ttk.Frame(parent); area.pack(fill='x')
        box = tk.Listbox(area, selectmode='multiple', exportselection=False, height=7, font=('Segoe UI', 10),
                         bg='#FBFCFE', fg='#334155', selectbackground='#DBEAFE', selectforeground='#1D4ED8',
                         highlightbackground='#DCE3EE', highlightthickness=1, bd=0, activestyle='none')
        box.pack(side='left', fill='x', expand=True)
        scrollbar = ttk.Scrollbar(area, command=box.yview); scrollbar.pack(side='right', fill='y')
        box.configure(yscrollcommand=scrollbar.set)
        box.bind('<<ListboxSelect>>', lambda _e: self.remember_popup_fields(kind))
        self.popup_lists[kind] = box
        controls = ttk.Frame(parent); controls.pack(fill='x', pady=(8, 0))
        ttk.Button(controls, text='Select all', command=lambda: self.select_popup_fields(kind, True)).pack(side='left', padx=(0, 8))
        ttk.Button(controls, text='Clear', command=lambda: self.select_popup_fields(kind, False)).pack(side='left')
        ttk.Label(parent, text='Click fields to toggle. Clear all to hide the popup.', style='Note.TLabel', wraplength=300).pack(anchor='w', pady=(8, 0))

    def remember_popup_fields(self, kind):
        self.popup_selected[kind] = [self.popup_names[kind][i] for i in self.popup_lists[kind].curselection()]

    def select_popup_fields(self, kind, select_all):
        box = self.popup_lists[kind]
        box.selection_clear(0, 'end')
        if select_all:
            box.selection_set(0, 'end')
        self.remember_popup_fields(kind)

    def refresh_popup_fields(self):
        if not hasattr(self, 'popup_lists'):
            return
        for kind, builtin in [('sector', SECTOR_FIELDS), ('site', SITE_FIELDS)]:
            names = builtin + [f'Excel: {header}' for header in self.headers]
            self.popup_names[kind] = names
            box = self.popup_lists[kind]
            box.delete(0, 'end')
            for i, name in enumerate(names):
                box.insert('end', name)
                selected = self.popup_selected[kind]
                if selected is None or name in selected:
                    box.selection_set(i)



    def show_issues(self):
        self.skip_invalid.set(False)
        self.issue_summary.set(f'{len(self.sectors):,} valid sectors · {len(self.issues):,} invalid entries')
        self.issue_text.configure(state='normal'); self.issue_text.delete('1.0', 'end')
        for issue in self.issues[:200]:
            self.issue_text.insert('end', f"Row {issue['row']} [{issue['column']}]: {issue['error']}\n")
        if len(self.issues) > 200:
            self.issue_text.insert('end', 'Showing the first 200 issues. Save the report for the full list.\n')
        if not self.issues:
            self.issue_text.insert('end', 'No validation errors.\n')
        self.issue_text.configure(state='disabled')

    def save_issues(self):
        path = filedialog.asksaveasfilename(defaultextension='.csv', filetypes=[('CSV report', '*.csv')], parent=self)
        if path:
            self.guard(lambda: write_issues(path, self.issues))

    def get_options(self):
        options = ExportOptions(self.title_var.get().strip() or 'RF sectors', self.mode.get(),
                                float(self.base.get()), float(self.gap.get()), self.outlines.get(), self.site_folders.get(),
                                [self.popup_names['sector'][i] for i in self.popup_lists['sector'].curselection()],
                                self.site_pin_var.get(),
                                [self.popup_names['site'][i] for i in self.popup_lists['site'].curselection()])
        for key, variable in self.pin_controls.items():
            value = variable.get().strip()
            setattr(options, key, float(value) if key.endswith(('_size', '_opacity')) else value)
        options.validate()
        return options

    def export(self, suffix):
        def start():
            self.commit_editor()
            if not self.sectors or self.import_signature != self.signature():
                raise ValueError('Validate the data and rebuild groups after changing import settings.')
            if self.issues and not self.skip_invalid.get():
                raise ValueError('Correct the invalid Excel entries, or explicitly select “Export valid sectors” to skip them.')
            options = self.get_options()
            path = filedialog.asksaveasfilename(defaultextension=suffix, initialfile='RF sectors'+suffix,
                                              filetypes=[(suffix[1:].upper(), '*'+suffix)], parent=self)
            if not path:
                return
            # Copy settings to ensure the background export sees a stable snapshot.
            styles = [GroupStyle(**asdict(s)) for s in self.styles]
            def done(count):
                self.status.set(f'Exported {count:,} sectors to {Path(path).name}')
                messagebox.showinfo('Export complete', f'{count:,} sectors exported.\n\n{path}\n\nOpen this file in Google Earth Pro.', parent=self)
            self.run_job(lambda: export_kml(path, self.sectors, styles, options), done, 'Exporting sectors…')
        self.guard(start)

    def save_profile(self):
        def save():
            self.commit_editor()
            profile = {'version': 1, 'mapping': asdict(self.get_mapping()), 'layout': self.layout.get(),
                       'styles': [asdict(s) for s in self.styles], 'options': asdict(self.get_options())}
            path = filedialog.asksaveasfilename(defaultextension='.json', filetypes=[('Settings', '*.json')], parent=self)
            if path:
                Path(path).write_text(json.dumps(profile, indent=2, ensure_ascii=False), encoding='utf-8')
                self.status.set('Settings saved.')
        self.guard(save)

    def load_profile(self):
        path = filedialog.askopenfilename(filetypes=[('Settings', '*.json')], parent=self)
        if not path:
            return
        def load():
            profile = json.loads(Path(path).read_text(encoding='utf-8'))
            if profile.get('version') != 1:
                raise ValueError('Unsupported settings version.')
            mapping = Mapping(**profile['mapping'])
            styles = [GroupStyle(**s) for s in profile['styles']]
            if len({s.group for s in styles}) != len(styles):
                raise ValueError('Settings contain duplicate color groups.')
            for s in styles:
                s.validate()
            options = ExportOptions(**profile['options']); options.validate()
            if profile['layout'] not in ('sector', 'site') or not isinstance(mapping.azimuths, list):
                raise ValueError('Invalid import settings.')
            self.pending_profile = profile
            self.styles = styles
            self.edit_index = None
            self.import_signature = None
            self.title_var.set(options.title); self.mode.set(options.altitude_mode)
            self.base.set(str(options.base_height)); self.gap.set(str(options.layer_gap))
            self.outlines.set(options.outlines); self.site_folders.set(options.site_folders)
            self.site_pin_var.set(options.site_pins)
            for key, variable in self.pin_controls.items():
                variable.set(str(getattr(options, key)))
            self.popup_selected = {'sector': options.sector_fields, 'site': options.site_fields}
            self.refresh_popup_fields()
            self.update_altitude_controls()
            if self.headers:
                self.apply_profile_mapping()
            self.refresh_groups()
            self.status.set('Settings loaded. Check column mapping and rebuild groups.')
        self.guard(load)

    def apply_profile_mapping(self):
        profile = self.pending_profile
        mapping = Mapping(**profile['mapping'])
        required = [mapping.site, mapping.latitude, mapping.longitude, mapping.group] + mapping.azimuths
        if mapping.sector:
            required.append(mapping.sector)
        missing = [h for h in required if h not in self.headers]
        if missing:
            messagebox.showwarning('Settings columns differ', 'Map the columns for this workbook. Missing: ' + ', '.join(missing), parent=self)
        else:
            self.set_mapping(mapping)
            self.layout.set(profile['layout'])
        self.pending_profile = None

    def close(self):
        if self.busy:
            messagebox.showinfo('Working', 'Please wait for the current operation to finish before closing.', parent=self)
            return
        self.executor.shutdown(wait=False)
        self.destroy()

    def setup_theme(self):
        style = ttk.Style(self); style.theme_use('clam')
        style.configure('.', font=('Segoe UI', 10), foreground='#23324B', background='#FFFFFF')
        style.configure('TFrame', background='#FFFFFF')
        style.configure('Workspace.TFrame', background='#F3F5F9')
        style.configure('Card.TFrame', background='#FFFFFF')
        style.configure('TLabel', background='#FFFFFF')
        style.configure('Title.TLabel', background='#F3F5F9', font=('Segoe UI', 21, 'bold'), foreground='#13213A')
        style.configure('Heading.TLabel', font=('Segoe UI', 12, 'bold'), foreground='#13213A')
        style.configure('Note.TLabel', foreground='#66758C', font=('Segoe UI', 9))
        style.configure('WorkspaceNote.TLabel', background='#F3F5F9', foreground='#66758C', font=('Segoe UI', 10))
        style.configure('Eyebrow.TLabel', foreground='#718097', font=('Segoe UI', 8, 'bold'))
        style.configure('Metric.TLabel', font=('Segoe UI', 18, 'bold'), foreground='#172D4E')
        style.configure('Accent.TLabel', foreground='#2563EB', font=('Segoe UI', 9, 'bold'))
        style.configure('TButton', padding=(13, 9), borderwidth=1, relief='flat', background='#FFFFFF', bordercolor='#DCE3EE', focusthickness=2, focuscolor='#93B4F6')
        style.map('TButton', background=[('disabled', '#EEF1F6'), ('active', '#EFF5FF')], foreground=[('disabled', '#94A3B8')], bordercolor=[('focus', '#3B82F6')])
        style.configure('Primary.TButton', background='#2563EB', foreground='white', bordercolor='#2563EB', font=('Segoe UI', 10, 'bold'))
        style.map('Primary.TButton', background=[('disabled', '#B9CAED'), ('pressed', '#1E40AF'), ('active', '#1D4ED8')], foreground=[('disabled', '#F8FAFC'), ('!disabled', 'white')])
        style.configure('Nav.TButton', background='#101C32', foreground='#9CAFCD', borderwidth=0, bordercolor='#101C32', padding=(14, 15), anchor='w', font=('Segoe UI', 10))
        style.map('Nav.TButton', background=[('active', '#1D2E49')], foreground=[('active', 'white'), ('disabled', '#5D708F')])
        style.configure('NavActive.TButton', background='#234B84', foreground='#FFFFFF', borderwidth=0, bordercolor='#234B84', padding=(14, 15), anchor='w', font=('Segoe UI', 10, 'bold'))
        style.map('NavActive.TButton', background=[('active', '#2A5798')], foreground=[('disabled', '#A4B7D3')])
        for name in ['TEntry', 'TCombobox', 'TSpinbox']:
            style.configure(name, padding=(9, 7), bordercolor='#DCE3EE', fieldbackground='#FBFCFE', lightcolor='#DCE3EE', darkcolor='#DCE3EE', arrowsize=14)
            style.map(name, bordercolor=[('focus', '#3B82F6')], fieldbackground=[('readonly', '#FBFCFE'), ('disabled', '#F1F4F8')], foreground=[('disabled', '#94A3B8')])
        style.configure('TCheckbutton', padding=(0, 5), background='white')
        style.configure('TRadiobutton', padding=(0, 5), background='white')
        style.map('TCheckbutton', background=[('active', 'white')])
        style.map('TRadiobutton', background=[('active', 'white')])
        style.configure('Treeview', rowheight=38, fieldbackground='white', background='white', borderwidth=0, font=('Segoe UI', 10))
        style.configure('Treeview.Heading', font=('Segoe UI', 9, 'bold'), background='#F2F5FA', foreground='#65748A', padding=(8, 10), relief='flat', borderwidth=0)
        style.map('Treeview', background=[('selected', '#EAF1FF')], foreground=[('selected', '#1D4ED8')])
        style.map('Treeview.Heading', background=[('active', '#E7EDF6')])
        style.configure('TScrollbar', background='#D5DEEA', troughcolor='#F3F5F9', borderwidth=0, arrowsize=12)
        style.layout('Vertical.TScrollbar', [('Vertical.Scrollbar.trough', {'sticky': 'ns', 'children': [('Vertical.Scrollbar.thumb', {'expand': '1', 'sticky': 'nswe'})]})])
        style.layout('Horizontal.TScrollbar', [('Horizontal.Scrollbar.trough', {'sticky': 'ew', 'children': [('Horizontal.Scrollbar.thumb', {'expand': '1', 'sticky': 'nswe'})]})])
        style.configure('TScrollbar', bordercolor='#F3F5F9', lightcolor='#F3F5F9', darkcolor='#F3F5F9', relief='flat')
        style.configure('Treeview', bordercolor='#FFFFFF', lightcolor='#FFFFFF', darkcolor='#FFFFFF', relief='flat')
        style.configure('TProgressbar', background='#3B82F6', troughcolor='#E8EEF8', borderwidth=0, thickness=4)
        style.configure('Workspace.TNotebook', background='#F3F5F9', borderwidth=0, tabmargins=0, bordercolor='#F3F5F9', lightcolor='#F3F5F9', darkcolor='#F3F5F9')
        style.layout('Workspace.TNotebook.Tab', [])
        self.option_add('*TCombobox*Listbox.font', ('Segoe UI', 10))
        self.option_add('*TCombobox*Listbox.selectBackground', '#DBEAFE')
        self.option_add('*TCombobox*Listbox.selectForeground', '#1E40AF')



    def scroll_page(self):
        page = ttk.Frame(self.tabs, style='Workspace.TFrame')
        canvas = tk.Canvas(page, bg='#F3F5F9', highlightthickness=0)
        scroll = ttk.Scrollbar(page, orient='vertical', command=canvas.yview)
        scroll.pack(side='right', fill='y'); canvas.pack(side='left', fill='both', expand=True)
        canvas.configure(yscrollcommand=scroll.set)
        body = ttk.Frame(canvas, style='Workspace.TFrame', padding=(2, 0, 12, 2))
        window = canvas.create_window((0, 0), window=body, anchor='nw')
        body.bind('<Configure>', lambda _e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.bind('<Configure>', lambda e: canvas.itemconfigure(window, width=e.width))
        page.scroll_canvas = canvas
        return page, body



    def page_wheel(self, event):
        if isinstance(event.widget, (tk.Listbox, tk.Text, ttk.Treeview, ttk.Combobox, ttk.Spinbox)):
            return
        page = self.nametowidget(self.tabs.select())
        canvas = page.scroll_canvas
        if canvas.bbox('all') and canvas.bbox('all')[3] > canvas.winfo_height():
            canvas.yview_scroll(-int(event.delta / 120), 'units')



    def card(self, parent, title, subtitle=None):
        border = tk.Frame(parent, background='#E0E6F0', bd=0)
        border.pack(fill='x', pady=(0, 14))
        inner = ttk.Frame(border, padding=16)
        inner.pack(fill='both', expand=True, padx=1, pady=1)
        ttk.Label(inner, text=title, style='Heading.TLabel').pack(anchor='w')
        if subtitle:
            note = ttk.Label(inner, text=subtitle, style='Note.TLabel', wraplength=650)
            note.pack(fill='x', pady=(4, 12))
            note.bind('<Configure>', lambda e: note.configure(wraplength=max(100, e.width)))
        else:
            ttk.Frame(inner, height=12).pack()
        return inner



    def navigate(self, index):
        if self.busy:
            return
        self.tabs.select(index)



    def page_changed(self, _event=None):
        if not hasattr(self, 'tabs') or not self.tabs.select():
            return
        index = self.tabs.index(self.tabs.select())
        titles = ['Import site data', 'Style your sectors', 'Export your map']
        subtitles = ['Connect an Excel workbook and map your sector fields.',
                     'Define each color group and arrange the visual layers.',
                     'Review the data and prepare a Google Earth file.']
        self.page_title.set(titles[index]); self.page_subtitle.set(subtitles[index])
        for i, button in enumerate(self.nav_buttons):
            button.configure(style='NavActive.TButton' if i == index else 'Nav.TButton')
        self.next_button.configure(text=['Validate & continue  →', 'Review export  →', 'Export KMZ…'][index],
                                   command=[self.build_groups, lambda: self.guard(self.go_export), lambda: self.export('.kmz')][index])
        self.update_metrics()



    def update_metrics(self):
        if not hasattr(self, 'metric_vars'):
            return
        values = {'rows': len(self.rows), 'sectors': len(self.sectors), 'groups': len(self.styles), 'issues': len(self.issues)}
        for key, var in self.metric_vars.items():
            value = f'{values[key]:,}' if self.headers else '—'
            if key != 'rows' and self.import_signature is None:
                value = '—'
            var.set(value)



    def schedule_preview(self):
        if self._filling_editor:
            return
        if self._preview_job is not None:
            self.after_cancel(self._preview_job)
        self._preview_job = self.after(120, self.preview_edits)



    def preview_edits(self):
        self._preview_job = None
        if self.edit_index is None or self.edit_index >= len(self.styles):
            return
        try:
            candidate = GroupStyle(self.styles[self.edit_index].group, self.edit['color'].get().strip(), float(self.edit['radius'].get()),
                                   float(self.edit['beamwidth'].get()), float(self.edit['opacity'].get()), self.enabled.get())
            candidate.validate()
            self.draw_preview(candidate)
        except (ValueError, TypeError):
            self.preview_caption.set('Enter valid settings to update the preview.')



    def update_altitude_controls(self):
        for entry in self.altitude_entries:
            entry.state(['disabled'] if self.mode.get() == 'ground' else ['!disabled'])

    def show_pin_appearance(self):
        window, body = self.help_dialog('appearance', 'Pin & label appearance', 'Independent styling for site markers and map labels.', width=780, height=690)
        if window is None:
            return
        columns = ttk.Frame(body, style='Workspace.TFrame'); columns.pack(fill='x')
        for i, (prefix, title, choices) in enumerate([('pin', 'Site marker', PIN_SHAPES), ('label', 'Site label', LABEL_STYLES)]):
            columns.columnconfigure(i, weight=1, uniform='appearance')
            container = ttk.Frame(columns, style='Workspace.TFrame')
            container.grid(row=0, column=i, sticky='nsew', padx=(0, 14 if i==0 else 0))
            card = self.card(container, title)
            key = 'pin_shape' if prefix=='pin' else 'label_style'
            ttk.Label(card, text='Shape' if prefix=='pin' else 'Label text', style='Note.TLabel').pack(anchor='w', pady=(0, 5))
            ttk.Combobox(card, textvariable=self.pin_controls[key], values=choices, state='readonly').pack(fill='x')
            ttk.Label(card, text='Color', style='Note.TLabel').pack(anchor='w', pady=(12, 5))
            row = ttk.Frame(card); row.pack(fill='x')
            ttk.Entry(row, textvariable=self.pin_controls[prefix+'_color'], width=12).pack(side='left', fill='x', expand=True)
            ttk.Button(row, text='Pick…', command=lambda k=prefix+'_color': self.pick_marker_color(k)).pack(side='left', padx=(8, 0))
            for suffix, label in [('size', 'Size · scale × (0.1–4)'), ('opacity', 'Opacity · % (0–100)')]:
                ttk.Label(card, text=label, style='Note.TLabel').pack(anchor='w', pady=(12, 5))
                ttk.Entry(card, textvariable=self.pin_controls[prefix+'_'+suffix]).pack(fill='x')
        ttk.Label(body, text='Google Earth controls the label font. “Hidden” removes the map label while keeping the pin.\nThese settings apply when site-name pins are enabled and are included in saved profiles.', style='WorkspaceNote.TLabel', wraplength=700).pack(anchor='w', pady=10)
        def apply():
            try:
                self.get_options()
            except ValueError as error:
                messagebox.showerror('Check appearance', str(error), parent=window)
                return
            window.destroy()
        ttk.Button(body, text='Apply appearance', style='Primary.TButton', command=apply).pack(anchor='e')

    def pick_marker_color(self, key):
        value = self.pin_controls[key].get()
        if not re.fullmatch(r'#[0-9a-fA-F]{6}', value):
            value = '#FFFFFF'
        parent = getattr(self, 'help_windows', {}).get('appearance')
        if parent is None or not parent.winfo_exists():
            parent = self
        color = colorchooser.askcolor(value, parent=parent, title='Choose '+key.replace('_', ' '))[1]
        if color:
            self.pin_controls[key].set(color)

    def help_dialog(self, key, title, subtitle, width=680, height=580):
        if not hasattr(self, 'help_windows'):
            self.help_windows = {}
        existing = self.help_windows.get(key)
        if existing is not None and existing.winfo_exists():
            existing.lift(); existing.focus_set()
            return None, None
        window = tk.Toplevel(self)
        self.help_windows[key] = window
        window.title(title+' · Sector Mapper')
        window.transient(self)
        window.configure(bg='#F3F5F9')
        width = min(width, self.winfo_screenwidth()-80)
        height = min(height, self.winfo_screenheight()-100)
        x = max(10, min(self.winfo_rootx()+80, self.winfo_screenwidth()-width-30))
        y = max(10, min(self.winfo_rooty()+40, self.winfo_screenheight()-height-60))
        window.geometry(f'{width}x{height}+{x}+{y}')
        window.minsize(min(width, 560), min(height, 460))
        window.bind('<Escape>', lambda _e: window.destroy())
        banner = tk.Frame(window, bg='#101C32', padx=24, pady=20); banner.pack(fill='x')
        tk.Label(banner, image=self.app_icons[3], bg='#101C32').pack(side='left', padx=(0, 16))
        text = tk.Frame(banner, bg='#101C32'); text.pack(side='left', fill='x', expand=True)
        tk.Label(text, text=title, bg='#101C32', fg='white', font=('Segoe UI', 18, 'bold')).pack(anchor='w')
        tk.Label(text, text=subtitle, bg='#101C32', fg='#A7B8D2', font=('Segoe UI', 10), wraplength=width-140, justify='left').pack(anchor='w', pady=(5, 0))
        footer = ttk.Frame(window, padding=(20, 12)); footer.pack(side='bottom', fill='x')
        ttk.Label(footer, text='Sector Mapper  /  '+APP_VERSION, style='Note.TLabel').pack(side='left')
        ttk.Button(footer, text='Close', command=window.destroy).pack(side='right')
        if key == 'guide':
            body = ttk.Frame(window, style='Workspace.TFrame', padding=20)
            body.pack(fill='both', expand=True)
        else:
            wrapper = ttk.Frame(window, style='Workspace.TFrame'); wrapper.pack(fill='both', expand=True)
            canvas = tk.Canvas(wrapper, bg='#F3F5F9', highlightthickness=0)
            scroll = ttk.Scrollbar(wrapper, command=canvas.yview); scroll.pack(side='right', fill='y')
            canvas.pack(side='left', fill='both', expand=True); canvas.configure(yscrollcommand=scroll.set)
            body = ttk.Frame(canvas, style='Workspace.TFrame', padding=20)
            item = canvas.create_window((0, 0), window=body, anchor='nw')
            body.bind('<Configure>', lambda _e: canvas.configure(scrollregion=canvas.bbox('all')))
            canvas.bind('<Configure>', lambda event: canvas.itemconfigure(item, width=event.width))
            window.bind('<MouseWheel>', lambda event: canvas.yview_scroll(-int(event.delta/120), 'units') if
                        not isinstance(event.widget, (tk.Text, tk.Listbox, ttk.Combobox)) and canvas.bbox('all') and
                        canvas.bbox('all')[3] > canvas.winfo_height() else None)
        return window, body

    def copy_support_text(self, text, label):
        self.clipboard_clear(); self.clipboard_append(text)
        label.configure(text='Copied to clipboard.')

    def show_about(self):
        window, body = self.help_dialog('about', 'About Sector Mapper', 'RF data, shaped for Google Earth.', height=560)
        if window is None:
            return
        card = self.card(body, 'Sector Mapper', 'A desktop workspace for converting Excel site and sector data into styled KML and KMZ maps.')
        for name, value in [('Version', APP_VERSION), ('Build', APP_BUILD), ('Author', APP_AUTHOR)]:
            row = ttk.Frame(card); row.pack(fill='x', pady=7)
            ttk.Label(row, text=name, style='Note.TLabel', width=14).pack(side='left')
            ttk.Label(row, text=value, font=('Segoe UI', 11, 'bold')).pack(side='left')
        ttk.Label(card, text='Excel import  ·  Color groups  ·  Site pins\nCustom popups  ·  Google Earth export', style='Note.TLabel').pack(anchor='w', pady=(15, 0))
        info = ttk.Label(body, text='', style='WorkspaceNote.TLabel'); info.pack(anchor='w', pady=6)
        ttk.Button(body, text='Copy version & build', command=lambda: self.copy_support_text(f'Sector Mapper {APP_VERSION}\nBuild {APP_BUILD}\nAuthor: {APP_AUTHOR}', info)).pack(anchor='w')

    def show_support(self):
        window, body = self.help_dialog('support', 'Contact & Support', 'Questions, feedback, or help with your workflow.', height=760)
        if window is None:
            return
        card = self.card(body, 'Get in touch', 'Mohamed Ashraf')
        ttk.Label(card, text='EMAIL', style='Eyebrow.TLabel').pack(anchor='w')
        ttk.Label(card, text=SUPPORT_EMAIL, style='Accent.TLabel').pack(anchor='w', pady=(6, 10))
        row = ttk.Frame(card); row.pack(fill='x')
        notice = ttk.Label(body, text='Links open only when you choose them. No report is sent automatically.', style='WorkspaceNote.TLabel', wraplength=600)
        ttk.Button(row, text='Open email app', command=lambda: webbrowser.open('mailto:'+SUPPORT_EMAIL+'?subject=Sector%20Mapper%20Support')).pack(side='left', padx=(0, 8))
        ttk.Button(row, text='Copy email', command=lambda: self.copy_support_text(SUPPORT_EMAIL, notice)).pack(side='left')
        ttk.Separator(card).pack(fill='x', pady=18)
        ttk.Label(card, text='LINKEDIN', style='Eyebrow.TLabel').pack(anchor='w')
        ttk.Label(card, text=SUPPORT_LINKEDIN, style='Accent.TLabel', wraplength=540).pack(anchor='w', pady=(6, 10))
        row = ttk.Frame(card); row.pack(fill='x')
        ttk.Button(row, text='Open LinkedIn', command=lambda: webbrowser.open(SUPPORT_LINKEDIN)).pack(side='left', padx=(0, 8))
        ttk.Button(row, text='Copy link', command=lambda: self.copy_support_text(SUPPORT_LINKEDIN, notice)).pack(side='left')
        tip = self.card(body, 'For a useful support request', 'Include the version/build, the steps you took, the expected result, and the error message. A small anonymized sample can help reproduce a data issue.')
        ttk.Button(tip, text='Copy support template', command=lambda: self.copy_support_text(f'Sector Mapper {APP_VERSION} / build {APP_BUILD}\n\nSteps to reproduce:\n1. \n\nExpected result:\n\nActual result / error:\n', notice)).pack(anchor='w')
        notice.pack(fill='x')

    def show_quick_guide(self):
        window, body = self.help_dialog('guide', 'Quick Guide', 'From your Excel workbook to a finished sector map.', width=850, height=650)
        if window is None:
            return
        chapters = [
            ('1  Open Excel', 'Start with a workbook', 'Open Import data and choose Browse files. Select an .xlsx or .xlsm workbook. Choose the worksheet and the row containing column headers, then click Read worksheet.\n\nThe preview displays the first 30 rows. Source files remain unchanged. Legacy .xls files must first be saved as .xlsx. Formula cells use saved results; recalculate and save in Excel if needed.'),
            ('2  Map columns', 'Tell the app what each row represents', 'Choose One sector per row for one azimuth column, or One site per row for several azimuth columns.\n\nMap Site name, Latitude, Longitude and Color group. Sector name is optional. Select every applicable azimuth column for site rows.\n\nUse WGS84 decimal-degree coordinates and azimuth clockwise from north, 0–360°. A zero azimuth is valid. Empty azimuth cells are skipped.'),
            ('3  Split & validate', 'Build your color groups', 'Map Color group to Band, Technology, or another Excel column. Enable Split combined groups to expand values such as 1800+2600. The separator defaults to +.\n\nSelect Validate & continue. Review validation errors before export. Fix the Excel file and read it again, or explicitly choose to skip invalid entries on the Export map page.'),
            ('4  Style layers', 'Set size, color and order', 'Select a group in Colors & layers. Set color, radius in meters, beamwidth in degrees, opacity, and whether to include it in export. Apply changes.\n\nArrange the list bottom first and top last. The combined preview uses the same sample north azimuth and proportional radii. An opaque upper layer may hide a lower color; the legend lists every enabled group.\n\nFolder rank separately controls the Places-panel order. Use Match ranks to layers to align the two orders.'),
            ('5  Choose popups', 'Control what a click reveals', 'On Export map, scroll to Sector click information. Click a field to select or deselect it. Both calculated fields and Excel columns are available.\n\nSelect all includes every available field. Clear hides the popup. Popup selections do not change the names shown in the Places panel.'),
            ('6  Add site pins', 'Style pins and labels independently', 'Enable Add site-name pins under Optional site pins. One pin is created per site name and coordinate pair from exported sectors. Different source values are combined in site popups.\n\nOpen Pin & label appearance. Choose Pin, Circle, Diamond, Square, or Star. Set pin color, size (0.1–4×), and opacity (0–100%).\n\nChoose a label text style, including Hidden. Labels have their own color, size and opacity. Google Earth controls the font. Select the pin popup fields below the appearance controls.'),
            ('7  Export your map', 'Choose presentation and file format', 'Elevated layers use the configured base height and separation to apply the bottom-to-top order. Heights are visual offsets, not antenna heights. Terrain and tilted views may affect overlaps.\n\nFollow terrain clamps polygons to the ground, where overlap order is not guaranteed.\n\nExport KMZ for a single compressed file with marker images included. KML exports with site pins create a SectorMapper-assets folder beside the KML; keep both together. Open the exported file in Google Earth Pro and inspect the map and popups.'),
            ('8  Save & get help', 'Reuse your settings', 'Save settings stores mappings, group styles, ordering, popup selections, pin appearance and export options as JSON. Load settings for the next workbook, then review mappings and validate again.\n\nCtrl+O opens a workbook. Ctrl+S saves settings.\n\nAbout shows the version, build and author. Contact / Support includes email and LinkedIn links, plus a support-request template. Nothing is sent automatically.')]
        area = ttk.Frame(body); area.pack(fill='both', expand=True)
        steps = tk.Listbox(area, width=22, exportselection=False, bd=0, highlightthickness=0, bg='#EEF3FB', fg='#334155',
                           selectbackground='#DBEAFE', selectforeground='#1D4ED8', font=('Segoe UI', 11), activestyle='none')
        steps.pack(side='left', fill='y', padx=(0, 18))
        article = ttk.Frame(area); article.pack(side='left', fill='both', expand=True)
        heading = ttk.Label(article, style='Heading.TLabel', wraplength=460); heading.pack(anchor='w', pady=(0, 12))
        text = tk.Text(article, wrap='word', bd=0, highlightthickness=0, font=('Segoe UI', 11), padx=5, pady=5, spacing3=8,
                       fg='#334155', bg='white', state='disabled')
        scroll = ttk.Scrollbar(article, command=text.yview); scroll.pack(side='right', fill='y')
        text.configure(yscrollcommand=scroll.set); text.pack(fill='both', expand=True)
        for label, _, _ in chapters:
            steps.insert('end', label)
        progress = tk.StringVar()
        controls = ttk.Frame(body, style='Workspace.TFrame'); controls.pack(side='bottom', before=area, fill='x', pady=(14, 0))
        def show(index):
            index = max(0, min(index, len(chapters)-1))
            steps.selection_clear(0, 'end'); steps.selection_set(index); steps.see(index)
            heading.configure(text=chapters[index][1])
            text.configure(state='normal'); text.delete('1.0', 'end'); text.insert('1.0', chapters[index][2]); text.configure(state='disabled'); text.yview_moveto(0)
            progress.set(f'Step {index+1} of {len(chapters)}')
        steps.bind('<<ListboxSelect>>', lambda _e: show(steps.curselection()[0]) if steps.curselection() else None)
        ttk.Label(controls, textvariable=progress, style='WorkspaceNote.TLabel').pack(side='left')
        ttk.Button(controls, text='Next →', style='Primary.TButton', command=lambda: show(steps.curselection()[0]+1)).pack(side='right')
        ttk.Button(controls, text='← Previous', command=lambda: show(steps.curselection()[0]-1)).pack(side='right', padx=8)
        show(0)






if __name__ == '__main__':
    if '--build-exe' in sys.argv[1:]:
        raise SystemExit(build_windows_executable())
    App().mainloop()
