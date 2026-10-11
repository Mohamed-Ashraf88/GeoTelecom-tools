"""SitePin — Site Mapping & KMZ Export, version 3.4.3.

Complete single-file application, based on SpatialPin 3.4.0.
All five upgrade stages are included. Standard KMZ icons are embedded below.
Author: Mohamed Ashraf

Run with Python 3.12 (64-bit) and these dependencies installed:
    python -m pip install PyQt6 PyQt6-WebEngine pandas openpyxl xlrd folium charset-normalizer
    python SitePin_V3.4.3.py

Local logs: %LOCALAPPDATA%/SitePin/logs/SitePin.log
Existing SpatialPin .sppreset files remain compatible.
Online basemaps and HTML maps require internet access.
"""

import sys
import os
import zipfile
import json
import re
import threading
import queue
import io
import colorsys
import math
import tempfile
import html
from html import escape as escape_html
import copy
import csv
import codecs
import base64
import hashlib
import logging
from logging.handlers import RotatingFileHandler
from branca.element import MacroElement, Template
import pandas as pd
import xml.etree.ElementTree as ET

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QPushButton, QTableWidget, QTableWidgetItem, QHeaderView, QScrollArea,
    QComboBox, QDoubleSpinBox, QProgressBar, QColorDialog, QFileDialog,
    QMessageBox, QListWidget, QListWidgetItem, QSplitter, QFrame, QSizePolicy,
    QLineEdit, QCheckBox, QToolButton, QFormLayout, QAbstractItemView,
    QDialog, QDialogButtonBox, QSpinBox, QToolTip, QTableView, QMenu, QInputDialog, QTextBrowser
)
from PyQt6.QtCore import Qt, QTimer, QUrl, QObject, pyqtSlot, pyqtSignal, QPointF, QRectF, QAbstractTableModel, QModelIndex, QBuffer, QIODevice
from PyQt6.QtGui import QColor, QFont, QDesktopServices, QIcon, QPixmap, QPainter, QPen, QBrush, QImage, QImageReader, QPainterPath
from PyQt6.QtWebChannel import QWebChannel

import folium
from folium.plugins import MarkerCluster
from PyQt6.QtWebEngineWidgets import QWebEngineView
from PyQt6.QtWebEngineCore import QWebEngineSettings

AUTHOR_NAME = "Mohamed Ashraf"
APP_NAME = "SitePin"
APP_DESCRIPTION = "Site Mapping & KMZ Export"
TOOL_NAME = f"{APP_NAME} — {APP_DESCRIPTION}"
VERSION = "3.4.3"
BUILD_NUMBER = "20260928.1"
SUPPORT_EMAIL = "Mohamed--Ashraf@outlook.com"
SUPPORT_LINKEDIN = "https://www.linkedin.com/in/mohamed---ashraf/"

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

BUNDLED_ICONS = {'Pushpin': 'iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAYAAACqaXHeAAAC90lEQVR4nO2b32rUQBTGv3OSWhWLoL6LN76A6L1vIF74EEIvxHuvfAlBUBDFFkGw9VkEQaTd3XY/Oe0ZCSVNm93tmbiTHyxZspvM5Pz5JnMmAUZGRkZGRopFVnESknaeGrHMReQYuSGpKDUCSFbmBZJ3ADwDsIGrZw7AjP5dRD66AygiRCQkK98+IvmTeXhHsm6kYW9kSc8/BvDW83+GWMzz5oT3AJ4AOAiJBJLin7sk/7gnjjJFwKFvX3jfeqegLmCD2q38FMAN9/xJOmTgmmvCc5JbIjLrmwq6ROPXVzWMLom6IRbqiy7RsFl+KCyc94rCURSOonAUhaMoHEXhKApHUTiKwlEUjqJwFIWjKBzNdOyqkRwXcbjMNHSF2LR8umhfdIFjrPxlFn/jdTgrQ+Wqz0/9Gl6LyG8rifWtCUpgUbTqafDZIIuiQWXx+aDL4j0XRipPkYcA7vv3yxRRX7lXrY8c3MJI36Uxkg+8jH3U4eFUXv+AIOplTyAi80sujv5w0drsUOy5R8dXT7HNDi1YyeJojRXg4dfaUTOO/U7yNoAJgK0LTmfnSqk1s1o//ndIbvhq0raH+LRD+H6RvOfHXfm6gyIIj5LLRNxNj5QQNKidtIiy68petehA+s9n04qo5w40opHGxe103BAlQdsTkUljDXI9UoCnHq3cCGeX1lJ62IXv+76Q22uNaMQ9acp+cI4BUl+mngJtv18JinjM0+eF9sRFcP2gD2k2xPlQ1xz60rC47cNlxLNGJ+SIgC4vhwhfNgPwVAgtz7808vxEH9wwu439IWhUQw0hnLQIYZthQlDkoU0I11cALxDCSS4BzBkBbd4OF8AsBmgIYbrhMQE8yCGA4QZIM0IXwj3f3XWLvJYpcOzbfff8v0lSjifPNbrBM9PeVPTc8XlCtdY3QmcwEUzlruiXLQZTInvpleDbUSWwwUSAnIZ6EsCwEthQDDBvlMo/mTDmevWmzmyAbwBu+WPuNjweoSRI1iTtnYOREeThL9kcUH2UlkjvAAAAAElFTkSuQmCC', 'Paddle': 'iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAYAAACqaXHeAAAEz0lEQVR4nO1bTY8bRRB91Z5Nwm7gkPwZhAKJBAGUTfaICFeExA/IicsGIXGBE4gLEkfIOSbRIg6gTRDwW/hSCAkk/nqo7Op107Jnerz2uMfLk0Zjzc5UV1VXVX+9Bf7HyYY01RDJsK1izmsjuyAiRNtBUkh2SBYLfr8VOW7pkFUINaU7IjIInj1r7Z0C8I7dPbTXHYCf7NJvH0Qy9dkwawcwMpzkWQAvALgE4G0AZ8xQfT4PangHwI8ADgF8ISK/BPJFRMZpkhVIdoLf50jeJPkrZ6NPsjfjmoU/SX5D8rVA/hZyAk0hy3c1/LfAgJEZNwquMvh3BuaoEOqIy9aWC52+ziLn7Pdlkt9GvVxlbAq8M4aB3Jve+LU5gVPDC5L7piSD3l4FQkeos19ZS0pwEn7jECR5xxQaBk5YNXpBm1cadQIt7CPjn7B59AOnN+METowv7PLGz6veTWAYOGF35TWB08LTzcD42AkaERfqOsHVNF6L2x6AawD6AHIYj9UGnRjpdPtDkjv6LHUK7WpUfF2cPAPgljW43jH4vyisQ14EcENE9HexzAgQm37eMiewTvQ07IQbJNURWhcqO0mqXvCLEJIv2dx8mFnvhxiYIw5F5FLKAsqlSLXh5QPr+ZyhxqvBF0m+ah3XWdgBJDX0VeA5AM9bxOQW+jFG1lGvm/Gl+roKYb6QvGlh329yF2lBFKbjW1a7dHiU2g7g5CMd9k4D0ElGpTczwVhvi9qX7dlcvV2ppEn4n7INjcr3M4LqrXXr4rEcYNAI+BvthBbB0pR1Cfmv+3fPtST/Y93f1e23sjrgEoSdaZHhtXV3CULy24BMR6XuDiccLuGdEx8BO2gvdo5TA0Z2/w7AU3s397XALN11u87NO2t0CUK+B9DLeAVYqruIPC3T3SUI2wbwD9qHYcqmiJv3Bw0ZWwY/BPClPT467MwYGupbdv+8Sm9XIWxkuXNgaeBalv8PbFOECzsA0zrw0HIp90LodT6wvcHF9wNkmgZq9Gd2zzkNaJ30B4CvbP5fqq9LEDqwZfEnVgzH2+PIEwOz6VPjFBRVVBtXJdGiQPNIvXo94/mA3wO4D+Aj2w6rjFaXItlvLopIF0DXvls6XeUYoEXlIwDvicjjVKKVq9OIHZBczzAVetb7H4vID1q3UvlELrUFOxjRovIEwBtBBKzbCX3btbqrDjBG2qCJA9LdgLDgSQtNwx/O3gn0kiY5QbsBKSLm86wSnjZzZLwnbazc+BlOuELy90CxVdFjPEIWSte4Cs0a7xGE3fmAM7AqntAgkKkstKvWdvJReBP8wGsRRW4ZaTGMnKmOPm/t5bFZywl1RiKSpE+L4xTI0PCDiCzZaQNb9P0ZhtTN9Z8jw4+cnSU4UfCIOkPydlS1U3te68ijgPdTZNnr8+Bp8iS3ST6uOV/w4/u+ydKJTvvAiRM0IvaisE4J/XtKeGrifwaaGipTUsGH/l+LUN6yBKe0Wp8KLEmFOPRzoOEtNRWuWmGbxST3E52vLfSLVod+ydR5P+ptHxFDi5DtkI2+MaANj9a796KC5+97/h+tsIngtCBesELXC0L/dvjOxoKzU2Ec+mtb1a0pFc6SPDQn7G506JekgpIZuxtZ9KoQDnFrX9OvC9mv7HAC8C+MMNq06WoJ7AAAAABJRU5ErkJggg==', 'Circle': 'iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAYAAACqaXHeAAAFJklEQVR4nO1byY4cRRSMl91jY2MsZCO+BcRmI5vNmCsYjiD4AXPhYiMkLnACIQGWryxXywYEYjH7t4BBYjEw00ugNx2pSddkVdf0dHVX9UxIpZqprszKiHwv15fAPvaxp2HL+CjJ0u+aGVdOAE4I9yb8bDDl3b7KNfaraUGsadJmNsz8dtRvme874b8zgrApMWzeGZJ00qPCs7sAvATAydyry2s4FJJvAHgPgFvJzwB+SAWRGKN5CmHz9utYOJJ3A3gewEMA7gNwZIZsfwFwGcA3AH6MYuREXipI9pK/HyX5Gck/eCsGJDdIDkmOK674XhG/knyV5DF9x6oa04WBKgTJ4ySvZEhHwrMgCpKmv0HybE78RRMPfunvJ1WwWOjdkK4SI7UMF/v4UkSgPuiNUqHWnXjTiK7i+I3kmYWKwC2TP0byakJ83jU+DYPknrpEc+0CyZ4anzNSP5JfFka6KEs8QnKtERG45e+HSd4s1MKyEduGCyrjwbq8Qk3yrmjQQORjALcB8BGe/98G9DV4etmt08zW67YJVuelOPBwMwPg/uaDkOV0P+WIQ2uvmJNm9n2dAVOYmutESW/gnhL5QQvJIyHv1vA6ydtltTazAJz4vSt7CMCHGr+3kXzRFR4AcF4zz/5uLMDMbCzyhyRGrXajBSKcJ+lCeC9VWmlWw+8fBHC9pX5fhugK183sRFVbEKpy8T4VwGuq+S5hc9rsM1GSj6gie7UFIOmm7xn4zOseWUrbTb+IsSruMZHPlj+UJI4NxzMye/ep5U89d4a4tPZsXIrL9QihJPFY5u8TjVL1Wg4nG634pJ5t4xEqzP/OqoQdcgOvyMfjaLb4Qqgw/xelYhfNv8jlOQBHc24QpnQlXen2psHHMP/kfgiZZyPNprpu/hHuzgfK+ISM/48123s4907HEDdYDpbxCSUJvf+8idVClk+oSNDlms9hRwOhPYOAPY4w428r7wIE8B9WB6V8QvqPb2xqDuCbkO/q8bbt7Y4Rdz5/AriU4xNyqbTDuyqjQMdhAOu5H0LmmQ8coFWgti6A1kXk4tvrvnfQ24kAXwL4Xe90bUUo5eJD4Wu+V+BcisEVoZgiaQeoxdCutgPR//3+URmPUJLY43FcuU8UttLFLjFa8lduyVoYZW0BMMHXakF7HXSDyOFT7Q/UHwfYrW7wju5dcoPYi3kb9oEWQbLlDxWZDOUGbwH4VxlGVduOobi9bWYeaNUviywLZTnICtxvXMVzHeoNvNLcer8D8Ia6vlLrDVU5xQ0FM/Nd4St6vz3haduxGVCpkewrZra5BlAVVxjqZKpN0nMdcIUN1f6b2h5f2/X2uE2WyEyTiadlTnHXpU0YaenrmgugYI5hE5FhZxWOsr6EwKgyeFkc3yblbCROaE33GBo3aoEIw6Qs96t8zYTucBIhFhSJdaEQqrbMmv+C5CmVsfnJG7cixp5IaiAX39sUhkl43NWkPIubuZI8oPspkp+rMKOG4wbTcNmhrLDfqNlXIWlwPIDyYuIKo0yQ825IDwtu5oKf1renBkI1Ck7IRxM8rVD5IjYKMb7TCOdIUwHZFxPhNxvlthyYWIvngfzMgA5K+IEJPzhRxKCiPEVT/ssPS2iF6pKZ3ZjnoQnDHBHD6pJTI3foeMwJAC9ooOLnharg02/HTyJ9WROaRo7NGBqATNQ3WoeFg1JQzFHVmaH33ULMLAqBxNx9htq2EejUcUNv1ha6scjvBIYFISESagRmLPwA5T72gb2J/wHXYFQJa7PcIAAAAABJRU5ErkJggg==', 'Star': 'iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAYAAACqaXHeAAAE7klEQVR4nO1bv28cVRD+ZvfsIwqBxImRKGhQBP8BgobKRMYWUCMahERFEfeA7CK0SKlp6KgJKJB0VPBHIFqkYAewopDYd/ehuZuxx6t95z1u99Y+7pNW73S7Oz/fzLw37w5YYIEF2gJJIdlpTQAAWVuK28frAG7ad60aYqYguWTjLZJPSV6x2eCGmV/QFCW5SvIBR9iy74aGmfcQ6Nj4AYBVtQmAt0VExwHmHSRzG++Z9w9tvBHvz7vyb5rSPcsBA5Jf2r2Zh8FMDUBymeRPpnTPRlo+WC1UifkBTSmS14LXXfmDtpJhNitGADy2P7HPPQBSkGM+kyFPlr4987Z739FrIxlmLZS+FQCHwfsO9bp6f33GsrVS+tzbEfOZDFle+lJoLRnOuvSl4Pfu+buY49KXwkyTYdZApj+6QvIrK30pFJNhnqBdCySlzBiDlRmNItJL0LoC4FfL/qrYacLrM4pdAK+KyJ8Jup0ErUFqHWFrjBPoVH3Q0LerTKBnC8ZR2j1reHjpq5LUxN7TSnDT9ghOyzEQkUcVaFVidgIkVYkXAOzbtNVnMrPq6wDesM+Z3VNjdQF8DGA5kPJ7FyoqXgY12j+BluMAwFcAnoZ7LtPPAH4JMtOc9hyAByIySBqAI+XV2l8DeB/AXwVFlch5wH7BMJcBfAPgQ51J0QhSjH2d/iSfAfApgM8SXilDqqc3bcJKhWNpzknMtlsAvhCRJ67jWOF4bIgNAN+GGDwvjUuXVcf3RORuUfGxZVBGynf1RSUA4KERLE2AZwx9k/VhUF51Sc2kSsvYFZLf2wJFp9FZhcumsq7UspDisRHyYIR+hRXdLDEwmWgy5rUoX6gOw8UHye3QyHSmbcJlODTZOlHm2sDREtQNsRnW7OM2N00jyrDpije6jSbZtXEjdHfaMILz3LNqdSRb42D7ybH+ZFdjcmwazSW7KZPjTkiOTVQIp3lovJpJdlMmxw2bnnWXSS9zT0K815LsZFoCLoytKnU3+Lttmqrs/SuzMFq6yXnRdoO6JZ76/CCrRz7k1hDZMuXL2t7TQIym0t4yXvmZa4kBuITmcanxltgkCDtH7Qj9Fs78616MOM0/ALysHaHUDm8SZPXJN+z86NU0auWT1UDDewQfAbjYQPwX88BF4xV5tzsDOIrJy/9R8Umn8JBXXXlAaor/5631fW2C+I/P9StmdX9HW+bXReTvafNAhnrw2DrDVcBCmHj32e9VQdd4To2spvfXbBF02sKkb4ovWcvqXQAvAdDWW4zzcYYYGK+1OnTIpnk5eO4180rq6It25WaEHQCviMh3IrIvIptmjD0zjiT6j35o0jWeUYZ29gAkL5DUxmNZX2AQjrwVP5J8K9AYrufDIepV6+o4nYMxvyS5a7ybbYCkEIRetpPfuGMrGmOX5Dvh3aWi0HFLS3KN5P0ELeehPIcnUW0ZILdxneTjsAOMXu+ZR6+6oOP27vFHEdZn2DHjueKufN94rrfWD+Cx9beDR2JDRD3oiWoiIQuzQcPiTqB7EGbcdpSlzfjvhWbIrnkuT033CXgcHXNpCIXZ0LOrnTzA4/hX5XRT4rjj072uqXlKknwUQkbain/FD/6Tlmm8XoVnSJJaUdhKHuCx1T8neTt83+ifHkrC4rbKEGWaKXjSK/m88y2FdYTl//iHqwUWWGCBBRbA+ca/G7h3O21LZgcAAAAASUVORK5CYII=', 'Triangle': 'iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAYAAACqaXHeAAAB+UlEQVR4nO2bTU7DMBBGZ6YLNiy4CCcAUVXsWXCjHokLVAVxHi4A7UMWjuRKLc2PG8eTfpu0saJ4vveaRaWIXFMugAJWcAvXFA1wC9zPzgRgEYd+AT6aczKXEIcFNvzlMRaymNPwS2AH/ADbdG0uBWwj/e94XKbrc6BPpB8K2M/CAg7p72MBTRG+LeA4fZLPvi3gNH3cW8D/9P1bwHn6fi2gHX2/FtCevj8L6EbfnwV0p+/HAvrR92MB/enXbwHD6NdvAcPp12sBeejXawH56NdnAXnp12cB+emnJUzbAi5Dvx4LuBz9tIRpWsBl6U/fAi5PPy1hWhYwDv3pWsB49NMSpmEB49KfngWMTz8toawFlKGf3QLL0MU67EHGTRh4LyKh/KWq7vqWYH0uCjeLNw0KLuNmSvwWiQCklAHrAvSzWmAV089igVVKP5sFVjn9wRZYxfSzWGAO6A+ywCqnP9gCc0K/twXmgP4gC8wR/V4WmBP6vS0wZ/Q7W2CO6PeywBzS72SBOaPf2QJzSr+1BeaQficLzDH9VhaYU/qtLdAT9J9EJPzjuqucviQzvKvqqpnxXAGfIvIg/vKsqpu0BD0y/EpE3mJzXl5la2YJYF/D92MFqGp4cYu7eMHBuoOE59mNiHypanguXHOScLDBaz3B8tJ7uEYmlF+n7fC0eFuV7QAAAABJRU5ErkJggg==', 'Square': 'iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAYAAACqaXHeAAABHklEQVR4nO3bQWrDMBCF4X9ebLLtjdJT9Q5d9axtAqWkZIpABlNa2dBFI54+EDYID+PRLEcwDMNgLFqbmSngQOci4vrfOfTVAZmpiLhl5gl4BG5A6YYefQAvEVGeuwswl7bJzGfgif49RMRrZkZE5Hpj2vjwAlzrmulL1gN+q+8/mjaCaPXjvRagmbcwJ8wJc8KcMCfMCXPCnDAnzAlzwpwwJ8wJc8KcMCfMCXPCnDAnzAlzwpwwJ8wJc8KcMCfMCXPCnDAnzAlzwpwwJ8wJc8KcMCfMCXPTzpnbZfVkybeZ97SjQ5aB4+btkju05Dv/pQDvddz8c2e33OO0+LnVBdGMkHkEjqtgPSq5X75flBiqYEO5ZkLnxukPw8AvvgD3jzlGihkYWwAAAABJRU5ErkJggg==', 'Target': 'iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAYAAACqaXHeAAAJCklEQVR4nOVbSW9cRRCuHo+zO4uTCHGEX4CExJIVUEhQnHBnEScEJwQSXOCALQhBCEROSCAER5DgAs7CkksIQnDnkiNHSIDsxPaMp1CNvxp/Lr83896siJQ0ejPvve6u+rq6urqqJsmISFWTiEzQrespJZX/O+mS4HadUNXfVPVvXCf4+bCoKqMjE3SniGwQkbX4PXSqyGhpwSYdV7ltANAlNV/1Gbb6D20JQLAxCFo3Y6eqtfBaDffHVZu2sGGfQRvG6oCFrqaUmoKZ4PRss4hsonVv122qWkkp3Qj9GI+aUlocBJ9pULOdUmKBd4jI8wD8QXwMFAMi4bsJPi8iH5k2iMgvIvKTA4J+K4MComcyBjFb/nuHqr6qqt+q6nXtjv5Q1eOqeiiMM2rjvZJUdYy+T6rqjKr+FYSpqeqCqtZVdREfJr/n70X6JgAxLv8FUsy6gQDB/wxCu8ANCGb32pGD1CBA7LuTadSjGLMyip2jRa6KxpCqfh8Eb+CTJbAtiWtBsCs5gDgYDqTfm2bwpUtKvezjKSVj7IiIfIVtzoyX24FF+m6GbE5EPsb3D+H9XYAhtHt3i8g9IrJbRO4XkV141hwS22IF16ZBFJHTIvJMSsnc6bGhGEhdMkJNg6eqs7R+azQ7Tn9hpswgbgr9bKRZv2y/w/OtqnoQKu/EGrNAYxweil3QZcHtcypD4IUg+GRob+3WoJ8tBMAV/DY7siZj3EMwghFsBn2q1+XQkWjNzQaB3WApgJkMQrfcXLpuDgBsDhqWAEhrmarqUQK8HnYP+72H+ewrKZAFEyz8Is0GGyZzaVfZmCIAZI1N4x9Q1bOBB9eE86Sh/dsddHnwPRiMrbwzMFVka+oGAGpbpTHiEvTrLPPcD+ETZmAjEHaBfV+/QcKvLdJftwA4CADAeDpJ6s+a8HjfQFBYVjg5rHZ+nS4qfD8ACP7HBlW9SUvRbdFNeKS9HbF12RhNwrL7II74SWhG5nov0CcDsK0Mw9AA04Qj0EJ3mBZwne55a9Tl2Z+mWfe1b4Bs4BkpyjT9tv1fcXU7UwZM5+/1HP6au1FXWqDLM7WTDjYNVn08L4QwvwfHyJbULczcLewuLWepyPol++Qa2gg8zuB5eS3QZXRfCuiqo1tEZfn4ijODeXZX0Y+D6nQRTFdKgJCnpUaXaOcopwUK5FT1O9pmSq0vMnhVMMheY7tToR2sDhRZXjTGJAR2g2hjzdPJsfiOoMud7iRGG3SS287v5fUBwTfRnu2M8awzsfo23Vvf8gpqwdukBd7P++2WaiWnT3cln8PJiwOYtxC66kRVhMVeFpEptKtQcDSL7L4xWsd7X4jIuiXZChnaKzg5sgxP2anSYpNZE1bp0KEz4t+NPhER22fH8yK2OJragHtF5BUAaIIUJWPejrbrReRzO3YX4NPoU+MNIDpv1sc/eQ0rGQcQQ8lU0Rh+mN9Dp1fxfUWb0D7hVHcM0V9rX3YrMuBN8MftWIzYQ7XNeAlaZnEHJ2tvfDziwdrQLp/g4MQDj0VxtrVpxu23wwjxmaEsudF9H30WOWO8mWEHjuPZKo2vomGM0dtMT0B1toQ26+C2+nur+oRKvoBZ5MhQWapinKdV9T2LHIHfzHGRbMkCqQGe7Z062tt786YKJuivIrKVHjgQE6FDj9+3W5POoK29VcGNHuhajuBxXFu68Vwyj6Xh7yxick5UQ5a2E8Wc/rCoQfHBbsgAyTqsbWqmnUTkUgkNuAkE81TR7/vS4T67JVu71ykgmjeuQPPinj8HLYgacCN1sAEX8MwjsqaGd3GeL4dZe/8hEfkS67jbTI4zelZEnsS4ectvHGn2l0TkDfJd7P47InKc7FPLBlSxlxu6KwjGwpDzA0oDanRvSskYakvm/QGwyR60wIE/k1KyhIuF4nPtALw9yzvyRNj7l1NK17LaV9Aw7q1rILzF730m6gBgtwdA8vwAtDeBP8AY9S6FH8Pe/hm2sLzxPAxn7xwIXqVNwifsL7AfUGm+mZLyx+SD98U+uH+/T1XXu4rFtmjv6veuiPxI7m1R8nVq/D2RUrroiZic8XxmH6Zl4hpnxn2e5OJ22UTBiUNwJjwSpHBu1hVwTDiQGmN27ahRNsBJuYZp4rFOJ8u1McQuJY6xv5M3aJ3+Yz6+ByM69OMgHA5BlZjw5OO20yyO4x1D3DTO6dCXXV/Es3JBEV0afAzHST+mOqrneOAO/VTINfaESjsyY3fU2xYRHpOxn4Rnbb2j64CILF0PUsd83R+LIjr1Rf29heiPq3sDgZfXUE1SmGHi8xz1533+4BNZSviCnbsWFI0JrhBIVe9EaKyBa8sTLRFk9XDXvpxJ2sdy9ALAfnQYM8BHyw5AqauNGcnR8RLCc6L2PNkoty+Fl2kZLXAQ3JDVyiYkycBu6TYxEsJgsx2WaO8A6FJHexEP9DM6JyQ9SVEZUmZoPCRqOV1uvH3tvPckfIfQM6N9ijRlbMC5QR9nKiRqPWhzEymzQhMiJdfbRqTD2FlZiHUB7bI73QLgSVGaec5Qc6L2SF9Uv0RCkjWhVa6Cd1c5MWUB4HQavnNZju/3c7j2ng9sR7TWp6hqK5ar1JDd2U7tWkULJEweAJ76XgEe/AdP0nCN4RxpYKlEbVcU3FuuDtHg3mbWCaFtFZ4hA9ACrE19EGtbLMtxvgZfN6jLRvEA1QeyStaCa3scZbMrqsU4O0z3TDMegxadoX7Y7kRta1WODFx4J7b6wcfnFHU8/V1H/u4Y3OE5ajODmebEKc86u82uYYXKcgZGutJAHc0ok3U1LVIi265kNuYTZyk/OfqaYV2ZoY2F0ly14SVtnLSIAnMVSq1DwfTgagL7VDF+qc3sRuHytOQqagpiyXxfVD5JHyn+WQIGbxcixM8iptgpvm8xPKOfReQHi+chJDaQP00kGQBFIHDPBbeo7QOIE1rqfB0CsCcQnbaC6gWL4oY+h1MM3U+iGp7Mk2LWNhieD9apkQH/awyR1+askSBjGVmnBKt+De975Db+s6zvVJEhEYWiFyFYDJPXcL99yLrPVJHRUp0+I6HqqAaG+rv/b9d0uwEwJyIztAtwacvQ6F/4zwR63GhTpgAAAABJRU5ErkJggg=='}

class OperationCancelled(Exception):
    pass


def check_cancel(event):
    if event is not None and event.is_set():
        raise OperationCancelled('Cancelled; existing files are unchanged.')


LOG = logging.getLogger(APP_NAME)
LOG.addHandler(logging.NullHandler())
LOG_PATH = None


def configure_logging():
    global LOG_PATH
    folder = os.path.join(os.environ.get('LOCALAPPDATA', tempfile.gettempdir()), APP_NAME, 'logs')
    try:
        os.makedirs(folder, exist_ok=True)
        LOG_PATH = os.path.join(folder, APP_NAME + '.log')
        handler = RotatingFileHandler(LOG_PATH, maxBytes=1000000, backupCount=3, encoding='utf-8')
        handler.setFormatter(logging.Formatter('%(asctime)s %(levelname)s %(message)s'))
        LOG.addHandler(handler)
        LOG.setLevel(logging.INFO)
        LOG.info('Started version %s; Python %s', VERSION, sys.version.split()[0])
    except OSError:
        LOG_PATH = None


def display_text(value):
    return json.dumps(value, ensure_ascii=False) if isinstance(value, (dict, list)) else '' if pd.isna(value) else str(value)


def prepare_view(processor, cancel=None):
    columns = [c for c in processor.df.columns if not c.startswith('__parsed')]
    ready = all(key in processor.target_cols for key in ('point name', 'latitude', 'longitude'))
    issues, search = {}, {}
    duplicates = processor.duplicate_warnings()
    for row, record in enumerate(processor.df[columns].to_dict('records')):
        check_cancel(cancel)
        issues[row] = [error for axis in ('latitude', 'longitude')
                       if (error := coordinate_issue(record[processor.target_cols[axis]], axis))] if ready else ['Apply the column mapping to validate this row.']
        search[row] = ' '.join(display_text(value) for value in record.values()).casefold()
    return dict(issues=issues, search=search, duplicates=duplicates)


class DataCell:
    """Lightweight access to one source cell; no per-cell widgets are allocated."""
    def __init__(self, table, row, column):
        self.table, self.r, self.c = table, row, column
    def row(self): return self.r
    def column(self): return self.c
    def text(self): return self.table.source_text(self.r, self.c)
    def setText(self, text):
        if not self.table.signalsBlocked():
            self.table.itemChanged.emit(EditedCell(self, text))
    def toolTip(self): return self.table.message(self.r)
    def background(self): return QBrush(self.table.row_color(self.r))


class EditedCell(DataCell):
    def __init__(self, cell, text):
        super().__init__(cell.table, cell.r, cell.c)
        self.value = text
    def text(self): return self.value


class FrameModel(QAbstractTableModel):
    def __init__(self, table):
        super().__init__(table)
        self.table = table
    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.table.rows)
    def columnCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else self.table.columnCount()
    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid(): return None
        row, column = self.table.rows[index.row()], index.column()
        if role in (Qt.ItemDataRole.DisplayRole, Qt.ItemDataRole.EditRole): return self.table.source_text(row, column)
        if role == Qt.ItemDataRole.ToolTipRole: return self.table.message(row)
        if role == Qt.ItemDataRole.BackgroundRole: return QBrush(self.table.row_color(row))
    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if role != Qt.ItemDataRole.DisplayRole: return None
        if orientation == Qt.Orientation.Horizontal:
            return (self.table.columns + ['Validation / warnings'])[section]
        return str(self.table.owner.processor.source_rows[self.table.rows[section]])
    def flags(self, index):
        flags = super().flags(index)
        if index.isValid() and index.column() < len(self.table.columns): flags |= Qt.ItemFlag.ItemIsEditable
        return flags
    def setData(self, index, value, role=Qt.ItemDataRole.EditRole):
        if role != Qt.ItemDataRole.EditRole or not index.isValid() or index.column() >= len(self.table.columns): return False
        self.table.item(self.table.rows[index.row()], index.column()).setText(str(value))
        return True


class DataTable(QTableView):
    itemChanged = pyqtSignal(object)
    itemSelectionChanged = pyqtSignal()
    def __init__(self, owner):
        super().__init__()
        self.owner, self.rows, self.positions, self.columns = owner, [], {}, []
        self.frame_model = FrameModel(self)
        self.setModel(self.frame_model)
        self.selectionModel().selectionChanged.connect(lambda *_: self.itemSelectionChanged.emit())
    def rowCount(self): return len(self.owner.processor.df) if self.owner.processor.df is not None else 0
    def columnCount(self): return len(self.columns)+1 if self.columns else 0
    def currentRow(self):
        index = self.currentIndex()
        return self.rows[index.row()] if index.isValid() and index.row() < len(self.rows) else -1
    def item(self, row, column): return DataCell(self, row, column)
    def source_text(self, row, column):
        if column == len(self.columns): return self.message(row)
        return display_text(self.owner.processor.df.at[row, self.columns[column]])
    def message(self, row):
        warnings = self.owner.duplicate_notes.get(row, [])
        return ' '.join(self.owner.issue_cache.get(row, []) + (['Warning: '+', '.join(warnings)] if warnings else []))
    def row_color(self, row):
        return QColor('#fff0ec') if self.owner.issue_cache.get(row) else QColor('#fff7df') if self.owner.duplicate_notes.get(row) else QColor(0,0,0,0)
    def reset_rows(self, rows):
        previous = self.currentRow()
        if rows == self.rows: return
        self.frame_model.beginResetModel()
        self.rows = rows
        self.positions = {row:index for index,row in enumerate(rows)}
        self.frame_model.endResetModel()
        if previous in self.positions: self.selectRow(previous)
    def isRowHidden(self, row): return row not in self.positions
    def selectRow(self, row):
        if row in self.positions: super().selectRow(self.positions[row])
    def scrollToItem(self, item):
        if item.row() in self.positions: self.scrollTo(self.frame_model.index(self.positions[item.row()], item.column()))
    def refresh(self, row):
        if row in self.positions:
            index = self.positions[row]
            self.frame_model.dataChanged.emit(self.frame_model.index(index,0), self.frame_model.index(index,len(self.columns)))


class LazyWebView(QWidget):
    """Do not start Chromium until an online map is requested."""
    loadFinished = pyqtSignal(bool)
    def __init__(self, owner):
        super().__init__()
        self.owner, self.browser = owner, None
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0,0,0,0)
    def ensure_browser(self):
        if self.browser is None:
            self.browser = QWebEngineView(self)
            self.layout().addWidget(self.browser)
            self.browser.settings().setAttribute(QWebEngineSettings.WebAttribute.LocalContentCanAccessRemoteUrls, True)
            self.owner.channel = QWebChannel(self.browser.page())
            self.owner.channel.registerObject('bridge', self.owner.bridge)
            self.browser.page().setWebChannel(self.owner.channel)
            self.browser.loadFinished.connect(self.loadFinished.emit)
        return self.browser
    def load(self, url): self.ensure_browser().load(url)
    def page(self): return self.ensure_browser().page()
    def stop(self):
        if self.browser is not None: self.browser.stop()
    def url(self): return self.browser.url() if self.browser else QUrl()


def validate_style(config):
    if not isinstance(config, dict): raise ValueError('Each category needs a style object.')
    result = MainWindow.default_style()
    for key, value in config.items():
        if key not in result: continue
        if key in ('icolor','lcolor') and (not isinstance(value,str) or not re.fullmatch(r'#[0-9a-fA-F]{6}',value)): raise ValueError('Colors must use #RRGGBB.')
        if key == 'shape' and value not in list(GE_ICONS)+['Custom Image']: raise ValueError('Unknown icon shape.')
        if key in ('opacity','iscale','lscale'):
            low, high = {'opacity':(.1,1),'iscale':(.5,3),'lscale':(0,3)}[key]
            if isinstance(value,bool) or not isinstance(value,(int,float)) or not math.isfinite(value) or not low <= value <= high: raise ValueError('Invalid '+key)
        if key == 'custom_path' and value is not None and not isinstance(value,str): raise ValueError('Invalid image path.')
        result[key] = value
    return result


def png_icon(path):
    if not path or not os.path.isfile(path): raise ValueError('Custom icon is missing. Choose its image again.')
    if os.path.getsize(path) > 8*1024*1024: raise ValueError('Custom icons must be smaller than 8 MB.')
    reader = QImageReader(path)
    size = reader.size()
    if not size.isValid() or size.width()*size.height() > 16000000: raise ValueError('Invalid image or image exceeds 16 megapixels.')
    image = reader.read()
    if image.isNull(): raise ValueError('Custom icon could not be read as an image.')
    image = image.scaled(256,256,Qt.AspectRatioMode.KeepAspectRatio,Qt.TransformationMode.SmoothTransformation)
    buffer = QBuffer()
    buffer.open(QIODevice.OpenModeFlag.WriteOnly)
    image.save(buffer,'PNG')
    return bytes(buffer.data())


def save_preset(path, processor, styles, fields):
    if os.path.normcase(os.path.realpath(path)) == os.path.normcase(os.path.realpath(processor.current_filepath)): raise ValueError('The original input file is protected.')
    icons, saved_styles = {}, {}
    for category, config in styles.items():
        style = validate_style(config)
        if style['shape'] == 'Custom Image':
            data = png_icon(style['custom_path'])
            name = 'icons/'+hashlib.sha256(data).hexdigest()+'.png'
            icons[name] = data
            style['custom_path'] = name
        else: style['custom_path'] = None
        saved_styles[category] = style
    manifest = dict(format='SpatialPin preset',version=1,mapping=processor.target_cols,fields=fields,styles=saved_styles)
    fd, temporary = tempfile.mkstemp(suffix='.sppreset',dir=os.path.dirname(os.path.abspath(path)))
    os.close(fd)
    try:
        with zipfile.ZipFile(temporary,'w',zipfile.ZIP_DEFLATED) as archive:
            archive.writestr('preset.json',json.dumps(manifest,ensure_ascii=False,indent=2))
            for name,data in icons.items(): archive.writestr(name,data)
        os.replace(temporary,path)
    finally:
        if os.path.exists(temporary): os.remove(temporary)


def read_preset(path, icon_folder):
    # Validate everything before writing only hash-named PNGs to our private folder.
    with zipfile.ZipFile(path) as archive:
        infos = archive.infolist()
        if len(infos)>1000 or sum(info.file_size for info in infos)>32*1024*1024: raise ValueError('Preset is too large.')
        names = [info.filename for info in infos]
        if len(names)!=len(set(names)) or any(name!='preset.json' and not re.fullmatch(r'icons/[a-f0-9]{64}\.png',name) for name in names): raise ValueError('Invalid preset archive entries.')
        if 'preset.json' not in names: raise ValueError('Missing preset manifest.')
        manifest = json.loads(archive.read('preset.json'))
        if manifest.get('format')!='SpatialPin preset' or manifest.get('version')!=1: raise ValueError('Unsupported preset version.')
        mapping, fields, styles = manifest.get('mapping'),manifest.get('fields'),manifest.get('styles')
        if not isinstance(mapping,dict) or set(mapping)!=set(('point name','latitude','longitude','category')) or not all(isinstance(v,str) for v in mapping.values()): raise ValueError('Invalid preset mapping.')
        if len(set(mapping.values())) != len(mapping): raise ValueError('Preset roles must use distinct columns.')
        if not isinstance(fields,list) or not all(isinstance(v,str) for v in fields) or len(fields)!=len(set(fields)): raise ValueError('Invalid popup fields.')
        if not isinstance(styles,dict): raise ValueError('Invalid preset styles.')
        validated, images = {}, {}
        for category, config in styles.items():
            style = validate_style(config)
            if style['shape']=='Custom Image':
                name = style['custom_path']
                if name not in names or not re.fullmatch(r'icons/[a-f0-9]{64}\.png',name or ''): raise ValueError('Preset icon missing.')
                data = archive.read(name)
                if hashlib.sha256(data).hexdigest()!=os.path.basename(name)[:-4]: raise ValueError('Preset icon checksum mismatch.')
                image = QImage.fromData(data,'PNG')
                if image.isNull() or image.width()>256 or image.height()>256: raise ValueError('Invalid preset icon.')
                images[os.path.basename(name)] = data
                style['custom_path'] = os.path.join(icon_folder,os.path.basename(name))
            else: style['custom_path'] = None
            validated[category] = style
    os.makedirs(icon_folder,exist_ok=True)
    for name,data in images.items():
        with open(os.path.join(icon_folder,name),'wb') as handle: handle.write(data)
    return mapping,fields,validated



def hex_to_kml_color(hex_color, opacity=1.0):
    if not hex_color or len(hex_color) != 7:
        return "ffffffff"
    hex_color = hex_color.lstrip("#")
    alpha_hex = f"{int(opacity * 255):02x}"
    return f"{alpha_hex}{hex_color[4:6]}{hex_color[2:4]}{hex_color[0:2]}".lower()

def parse_coordinate(value, axis=None):
    """Parse finite WGS84 decimal degrees or strict degrees/minutes/seconds."""
    if not pd.api.types.is_scalar(value) or value is None or pd.isna(value) or isinstance(value, bool):
        return None
    text = str(value).strip().upper().replace("−", "-")
    text = text.replace("′", "'").replace("″", '"').replace("º", "°")
    if not text:
        return None
    direction = None
    if text[-1:] in ("N", "S", "E", "W"):
        direction, text = text[-1], text[:-1].strip()
        if axis == "latitude" and direction not in ("N", "S"):
            return None
        if axis == "longitude" and direction not in ("E", "W"):
            return None
    try:
        number = r"[0-9]+(?:\.[0-9]+)?"
        if re.fullmatch(r"[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:E[+-]?[0-9]+)?", text):
            result = float(text)
        elif re.fullmatch(r"[+-]?" + number + r"\s*°", text):
            result = float(text.rstrip("°").strip())
        else:
            match = re.fullmatch(
                r"([+-]?)([0-9]+)(?:\s*°\s*|\s+)(" + number +
                r")(?:(?:\s*'\s*|\s+)(" + number + r")\s*\"?|\s*'?)", text)
            if not match:
                return None
            sign, degrees, minutes, seconds = match.groups()
            minutes, seconds = float(minutes), float(seconds or 0)
            if minutes >= 60 or seconds >= 60:
                return None
            result = (float(degrees) + minutes / 60 + seconds / 3600)
            if sign == "-":
                result = -result
        if direction:
            # Signed values plus hemisphere are ambiguous; require one convention.
            if text.startswith(("-", "+")):
                return None
            result *= -1 if direction in ("S", "W") else 1
        limit = 90 if axis == "latitude" else 180
        return result if math.isfinite(result) and abs(result) <= limit else None
    except (ValueError, TypeError, OverflowError):
        return None


def coordinate_issue(value, axis):
    if parse_coordinate(value, axis) is not None:
        return ""
    limit = 90 if axis == "latitude" else 180
    if pd.api.types.is_scalar(value) and (value is None or pd.isna(value) or not str(value).strip()):
        return f"Missing {axis}."
    try:
        number = float(value)
        if not math.isfinite(number):
            return f"Invalid {axis}: enter a finite number."
        if abs(number) > limit:
            return f"Invalid {axis}: {number:g} is outside -{limit} to {limit}."
    except (TypeError, ValueError, OverflowError):
        pass
    return (f"Invalid {axis}: use WGS84 degrees from -{limit} to {limit}, "
            "or DMS with minutes/seconds below 60 and the correct hemisphere.")


def write_sample_csv(destination):
    rows = [
        ['Site ID', 'Latitude', 'Longitude', 'Category', 'Notes'],
        ['0001', '30.0444', '31.2357', 'Active', 'Example point in Cairo'],
        ['0002', '30.0131', '31.2089', 'Planned', 'Example point in Giza'],
        ['0003', '30 05 27.6 N', '31 19 19.2 E', 'Review', 'DMS example'],
    ]
    with open(destination, 'w', encoding='utf-8-sig', newline='') as handle:
        csv.writer(handle).writerows(rows)


def safe_web_text(value):
    # Folium uses JS template strings as well as HTML. Keep both contexts literal.
    return html.escape(str(value)).replace('`', '&#96;').replace('$', '&#36;')


def geographic_bounds(frame):
    """Smallest longitude arc, so a dateline-spanning dataset doesn't fit the whole world."""
    values = sorted(set(float(value) % 360 for value in frame['__parsed_lon']))
    if len(values) == 1:
        west = east = values[0]
    else:
        gaps = [((values[(i+1) % len(values)] - value) % 360, i) for i, value in enumerate(values)]
        gap, index = max(gaps)
        west = values[(index+1) % len(values)]
        east = west + 360 - gap
    if west >= 180:
        west -= 360
        east -= 360
    return [[float(frame['__parsed_lat'].min()), west], [float(frame['__parsed_lat'].max()), east]]


def display_longitude(longitude, west):
    return west + ((float(longitude) - west) % 360)


def sample_indices(frame, category_column, limit=500, selected=None):
    """Deterministic category-balanced sample. Includes selection without losing its category."""
    if frame.empty or limit < 1:
        return []
    if len(frame) <= limit:
        return list(frame.index)
    groups = {str(category): list(group.index) for category, group in frame.groupby(category_column, sort=True)}
    categories = sorted(groups)
    selected_category = str(frame.at[selected, category_column]) if selected in frame.index else None
    if len(categories) > limit:
        chosen = [categories[int(i * (len(categories)-1) / max(1, limit-1))] for i in range(limit)]
        if selected_category is not None and selected_category not in chosen:
            chosen[-1] = selected_category
        categories = chosen
    quotas = {category: 1 for category in categories}
    remaining = limit - len(categories)
    while remaining:
        changed = False
        for category in categories:
            if quotas[category] < len(groups[category]):
                quotas[category] += 1
                remaining -= 1
                changed = True
                if not remaining:
                    break
        if not changed:
            break
    result = []
    for category in categories:
        rows, count = groups[category], quotas[category]
        picks = [rows[int(i * (len(rows)-1) / max(1, count-1))] for i in range(count)]
        if category == selected_category and selected not in picks:
            picks[-1] = selected
        result.extend(picks)
    return result


def web_point(location, name, style, popup):
    """One rendering rule shared by embedded preview and HTML export."""
    return folium.CircleMarker(location=location, radius=5 * style.get('iscale', 1.2),
        color='white', weight=1.5, fill=True, fill_color=style.get('icolor', '#087f8c'),
        fill_opacity=style.get('opacity', 1.0), tooltip=safe_web_text(name),
        popup=folium.Popup(popup, max_width=360))


class CoordinateOverview(QWidget):
    """Offline coordinate plot, deliberately not presented as a geographic basemap."""
    pointSelected = pyqtSignal(int)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMouseTracking(True)
        self.setMinimumHeight(120)
        self.points = []
        self.selected = None
        self.bounds = [[-1, -1], [1, 1]]
        self.viewport = [-1, 1, -1, 1]
        self.drag_start = None
        self.drag_last = None
        self.interactive = True

    def set_points(self, points, bounds, selected=None):
        self.points = points
        self.bounds = bounds
        self.selected = selected
        self.fit_all()

    def fit_all(self):
        (south, west), (north, east) = self.bounds
        dx, dy = max(east-west, .01), max(north-south, .01)
        cx, cy = (west+east)/2, (south+north)/2
        self.viewport = [cx-dx*.58, cx+dx*.58, cy-dy*.58, cy+dy*.58]
        self.update()

    def plot_rect(self):
        return QRectF(58, 18, max(1, self.width()-76), max(1, self.height()-55))

    def point_position(self, point):
        west, east, south, north = self.viewport
        rect = self.plot_rect()
        lon = display_longitude(point['lon'], self.bounds[0][1])
        return QPointF(rect.left() + (lon-west)/(east-west)*rect.width(),
                       rect.bottom() - (point['lat']-south)/(north-south)*rect.height())

    def select_row(self, row, pan=True):
        self.selected = row
        point = next((point for point in self.points if point['row'] == row), None)
        if point and pan and not self.plot_rect().contains(self.point_position(point)):
            west, east, south, north = self.viewport
            lon = display_longitude(point['lon'], self.bounds[0][1])
            dx, dy = (east-west)/2, (north-south)/2
            self.viewport = [lon-dx, lon+dx, point['lat']-dy, point['lat']+dy]
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.fillRect(self.rect(), QColor('#edf3f6'))
        rect = self.plot_rect()
        west, east, south, north = self.viewport
        font = QFont('Segoe UI', 8)
        painter.setFont(font)
        tick_count = min(5, max(2, int(rect.height()/28)+1))
        for index in range(tick_count):
            fraction = index / (tick_count-1)
            x, y = rect.left()+rect.width()*fraction, rect.bottom()-rect.height()*fraction
            painter.setPen(QPen(QColor('#d2dfe7'), 1))
            painter.drawLine(QPointF(x, rect.top()), QPointF(x, rect.bottom()))
            painter.drawLine(QPointF(rect.left(), y), QPointF(rect.right(), y))
            painter.setPen(QColor('#61738b'))
            longitude = ((west+(east-west)*fraction+180) % 360)-180
            painter.drawText(QRectF(x-35, rect.bottom()+5, 70, 16), Qt.AlignmentFlag.AlignCenter, f'{longitude:.3f}°')
            painter.drawText(QRectF(0, y-8, 53, 16), Qt.AlignmentFlag.AlignRight, f'{south+(north-south)*fraction:.3f}°')
        painter.save()
        painter.setClipRect(rect.adjusted(-12, -12, 12, 12))
        ordered = sorted(self.points, key=lambda point: point['row'] == self.selected)
        for point in ordered:
            position = self.point_position(point)
            radius = 5 * point['style'].get('iscale', 1.2)
            fill = QColor(point['style'].get('icolor', '#087f8c'))
            fill.setAlphaF(point['style'].get('opacity', 1.0))
            painter.setBrush(fill)
            painter.setPen(QPen(QColor('#14263e') if point['row'] == self.selected else QColor('white'), 4 if point['row'] == self.selected else 1.5))
            painter.drawEllipse(position, radius, radius)
        painter.restore()
        painter.setPen(QColor('#61738b'))
        painter.drawText(QRectF(58, self.height()-17, self.width()-76, 16), Qt.AlignmentFlag.AlignLeft,
                         'Longitude →   Latitude ↑   •   Coordinate overview; no basemap')
        painter.end()

    def hit_test(self, position):
        candidates = [(math.hypot(self.point_position(point).x()-position.x(), self.point_position(point).y()-position.y()), point)
                      for point in self.points]
        if not candidates:
            return None
        distance, point = min(candidates, key=lambda entry: entry[0])
        return point if distance <= max(10, 5*point['style'].get('iscale', 1.2)+4) else None

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton and self.interactive:
            self.drag_start = self.drag_last = event.position()

    def mouseMoveEvent(self, event):
        if not self.interactive:
            return
        if self.drag_last is not None and event.buttons() & Qt.MouseButton.LeftButton:
            delta = event.position()-self.drag_last
            rect = self.plot_rect()
            west, east, south, north = self.viewport
            dx, dy = delta.x()/rect.width()*(east-west), delta.y()/rect.height()*(north-south)
            self.viewport = [west-dx, east-dx, south+dy, north+dy]
            self.drag_last = event.position()
            self.update()
        else:
            point = self.hit_test(event.position())
            if point:
                QToolTip.showText(event.globalPosition().toPoint(), '<b>' + safe_web_text(point['name']) + '</b><br>' + point['popup'], self)
            else:
                QToolTip.hideText()

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton and self.drag_start is not None:
            distance = event.position()-self.drag_start
            if abs(distance.x())+abs(distance.y()) < 5 and self.interactive:
                point = self.hit_test(event.position())
                if point:
                    self.select_row(point['row'])
                    self.pointSelected.emit(point['row'])
                    QToolTip.showText(event.globalPosition().toPoint(), '<b>'+safe_web_text(point['name'])+'</b><br>'+point['popup'], self)
            self.drag_start = self.drag_last = None

    def wheelEvent(self, event):
        if not self.interactive or not self.points:
            return
        factor = .8 if event.angleDelta().y() > 0 else 1.25
        west, east, south, north = self.viewport
        rect = self.plot_rect()
        fx, fy = (event.position().x()-rect.left())/rect.width(), (rect.bottom()-event.position().y())/rect.height()
        cx, cy = west+fx*(east-west), south+fy*(north-south)
        dx, dy = max(.0000001, min(720, (east-west)*factor)), max(.0000001, min(360, (north-south)*factor))
        self.viewport = [cx-fx*dx, cx+(1-fx)*dx, cy-fy*dy, cy+(1-fy)*dy]
        self.update()
        event.accept()


class KMLProcessor:
    def __init__(self):
        self.cancel_event = None
        self.df = None
        self.current_filepath = None
        self.target_cols = {}
        self.unique_categories = []
        self.sheet_names = []
        self.current_sheet = None
        self.import_options = {"header_row": 1, "delimiter": "auto", "encoding": "auto"}
        self.import_details = {}
        self.source_rows = []
        self.web_basemap = 'OpenStreetMap'

    @staticmethod
    def read_table(filepath, sheet_name=0, header_row=1, delimiter="auto", encoding="auto"):
        """Read a table without changing processor state; also used by import preview."""
        if not isinstance(header_row, int) or not 1 <= header_row <= 10000:
            raise ValueError("Header row must be between 1 and 10000.")
        extension = os.path.splitext(filepath)[1].lower()
        details = {"header_row": header_row}
        if extension in (".xlsx", ".xls"):
            with pd.ExcelFile(filepath) as workbook:
                sheets = workbook.sheet_names
                sheet = sheets[sheet_name] if isinstance(sheet_name, int) else sheet_name
                raw = pd.read_excel(workbook, sheet_name=sheet, header=None, dtype=object,
                                    keep_default_na=False)
            if header_row > len(raw):
                raise ValueError("The selected header row is beyond the end of the sheet.")
            headers = list(raw.iloc[header_row - 1])
            records = raw.iloc[header_row:].copy().reset_index(drop=True)
            source_rows = list(range(header_row + 1, header_row + 1 + len(records)))
            details.update(sheet=sheet, sheets=sheets)
        else:
            with open(filepath, 'rb') as handle:
                data = handle.read()
            guessed = False
            if encoding == "auto":
                if data.startswith((codecs.BOM_UTF32_LE, codecs.BOM_UTF32_BE)):
                    encoding = "utf-32"
                elif data.startswith((codecs.BOM_UTF16_LE, codecs.BOM_UTF16_BE)):
                    encoding = "utf-16"
                else:
                    try:
                        data.decode("utf-8-sig")
                        encoding = "utf-8-sig"
                    except UnicodeDecodeError:
                        # Legacy encodings cannot always be distinguished. Show the choice
                        # in import details and let the user override it in the preview.
                        try:
                            from charset_normalizer import from_bytes
                            match = from_bytes(data).best()
                            encoding = match.encoding if match else "cp1252"
                        except ImportError:
                            encoding = "cp1252"
                        guessed = True
            try:
                text = data.decode(encoding, errors="strict")
            except (UnicodeError, LookupError) as error:
                raise ValueError("Could not decode this file. Choose its encoding in Import options.") from error
            lines = text.splitlines(keepends=True)
            if header_row > len(lines):
                raise ValueError("The selected header row is beyond the end of the file.")
            selected_text = ''.join(lines[header_row-1:])
            if delimiter == "auto":
                try:
                    delimiter = csv.Sniffer().sniff(selected_text[:65536], delimiters=",;\t|").delimiter
                except csv.Error:
                    first = next(csv.reader([lines[header_row-1]]), [])
                    if len(first) > 1:
                        delimiter = ','
                    else:
                        raise ValueError("Could not detect a delimiter. Choose one in Import options.")
            if delimiter not in (",", ";", "\t", "|"):
                raise ValueError("Choose comma, semicolon, tab or pipe as the delimiter.")
            reader = csv.reader(io.StringIO(selected_text, newline=''), delimiter=delimiter, strict=True)
            headers = next(reader, [])
            rows, source_rows = [], []
            previous_line = reader.line_num
            for record in reader:
                source_row = header_row - 1 + previous_line + 1
                previous_line = reader.line_num
                if not record or not any(value.strip() for value in record):
                    continue
                if len(record) != len(headers):
                    raise ValueError(f"Row {source_row} has {len(record)} fields; expected {len(headers)}. Check the delimiter and quoting.")
                rows.append(record)
                source_rows.append(source_row)
            records = pd.DataFrame(rows)
            details.update(encoding=encoding, delimiter=delimiter, encoding_guessed=guessed)
        headers = [str(value).strip() for value in headers]
        if not headers or any(not value for value in headers):
            raise ValueError("The header row contains blank column names. Choose the correct header row.")
        normalized = [value.casefold() for value in headers]
        if len(set(normalized)) != len(normalized):
            raise ValueError("Column names must be unique (including case). Choose another header row or rename duplicate headers.")
        if records.empty:
            raise ValueError("The file contains no data rows after the selected header.")
        records.columns = headers
        # Keep literal IDs such as NA and NULL; only empty cells become missing.
        records = records.replace('', None)
        keep = ~records.isna().all(axis=1)
        source_rows = [row for row, selected in zip(source_rows, keep) if selected]
        records = records.loc[keep].reset_index(drop=True)
        if records.empty:
            raise ValueError("The file contains no nonempty data rows.")
        return records, source_rows, details
        
    def _find_best_column(self, cols, aliases):
        cols_lower = [str(c).strip().lower() for c in cols]
        for alias in aliases:
            alias = alias.lower()
            if alias in cols_lower: return cols[cols_lower.index(alias.lower())]
        for alias in aliases:
            alias = alias.lower()
            if len(alias) >= 2: 
                for i, c_lower in enumerate(cols_lower):
                    clean_c = c_lower.replace("_", " ").replace("-", " ")
                    if re.search(rf"\b{re.escape(alias)}\b", clean_c): return cols[i]
        return None

    def save_corrected_copy(self, destination):
        """Write a separate data-only copy; never rewrite the source workbook."""
        if self.df is None:
            raise ValueError("No dataset is loaded.")
        source = os.path.normcase(os.path.realpath(self.current_filepath))
        target = os.path.normcase(os.path.realpath(destination))
        if source == target or (os.path.exists(self.current_filepath) and os.path.exists(destination) and
                                os.path.samefile(self.current_filepath, destination)):
            raise ValueError("Choose a different filename. The original file is protected.")
        ext = os.path.splitext(destination)[1].lower()
        if ext not in (".xlsx", ".csv", ".geojson"):
            raise ValueError("Save a corrected copy as .xlsx, .csv, or .geojson.")
        export_df = self.df.drop(columns=["__parsed_lat", "__parsed_lon"], errors="ignore")
        fd, temporary = tempfile.mkstemp(suffix=ext, dir=os.path.dirname(target))
        os.close(fd)
        try:
            if ext == ".xlsx":
                export_df.to_excel(temporary, index=False)
            elif ext == ".csv":
                export_df.to_csv(temporary, index=False, encoding="utf-8-sig")
            else:
                if self.has_invalid_coordinates():
                    raise ValueError("Fix invalid coordinates before saving GeoJSON, or save an Excel/CSV copy.")
                properties = json.loads(export_df.to_json(orient="records", date_format="iso"))
                features = []
                for props, (_, row) in zip(properties, self.df.iterrows()):
                    features.append({"type": "Feature", "properties": props,
                        "geometry": {"type": "Point", "coordinates":
                            [float(row["__parsed_lon"]), float(row["__parsed_lat"])]}})
                with open(temporary, "w", encoding="utf-8") as handle:
                    json.dump({"type": "FeatureCollection", "features": features},
                              handle, ensure_ascii=False, allow_nan=False, indent=2)
            os.replace(temporary, target)
        finally:
            if os.path.exists(temporary):
                os.remove(temporary)

    def has_invalid_coordinates(self):
        if self.df is None or any(key not in self.target_cols for key in ("point name", "latitude", "longitude")):
            return True
        return any(coordinate_issue(row[self.target_cols[axis]], axis)
                   for _, row in self.df.iterrows()
                   for axis in ("latitude", "longitude"))

    def duplicate_warnings(self):
        warnings = {index: [] for index in self.df.index} if self.df is not None else {}
        if self.df is None or any(key not in self.target_cols for key in ("point name", "latitude", "longitude")):
            return warnings
        names = self.df[self.target_cols["point name"]].map(lambda value: str(value).strip().casefold())
        duplicates = names.duplicated(keep=False) & names.ne('') & names.ne('unnamed point')
        for index in self.df.index[duplicates]:
            warnings[index].append("Duplicate name")
        valid = ~self.df[['__parsed_lat', '__parsed_lon']].isna().any(axis=1)
        repeated = self.df[['__parsed_lat', '__parsed_lon']].duplicated(keep=False) & valid
        for index in self.df.index[repeated]:
            warnings[index].append("Duplicate coordinates")
        return warnings

    def export_selection(self, valid_only=False):
        """Return a stable export snapshot and an explicit report of omitted rows."""
        if self.df is None or any(key not in self.target_cols for key in ('point name', 'latitude', 'longitude')):
            raise ValueError("Apply the column mapping first.")
        snapshot = self.snapshot()
        errors = {}
        for index, row in snapshot.df.iterrows():
            check_cancel(self.cancel_event)
            issues = []
            for axis, internal in (("latitude", "__parsed_lat"), ("longitude", "__parsed_lon")):
                value = row[self.target_cols[axis]]
                snapshot.df.at[index, internal] = parse_coordinate(value, axis)
                issue = coordinate_issue(value, axis)
                if issue:
                    issues.append(issue)
            if issues:
                errors[index] = ' '.join(issues)
        if errors and not valid_only:
            raise ValueError("Fix invalid rows or explicitly select Valid rows only.")
        columns = [column for column in self.df.columns if not column.startswith('__parsed')]
        report = snapshot.df.loc[list(errors), columns].copy()
        # Separate metadata avoids collisions with user-provided column names.
        report.columns = ['Data: ' + column for column in columns]
        report.insert(0, 'Validation error', [errors[index] for index in errors])
        report.insert(0, 'Source row / feature', [self.source_rows[index] for index in errors])
        report.insert(0, 'Source sheet', self.current_sheet or '')
        report.insert(0, 'Source file', os.path.basename(self.current_filepath))
        snapshot.df = snapshot.df.drop(index=list(errors)).reset_index(drop=True)
        snapshot.source_rows = [row for index, row in enumerate(self.source_rows) if index not in errors]
        if snapshot.df.empty:
            raise ValueError("There are no valid rows to export.")
        snapshot.unique_categories = sorted(snapshot.df[self.target_cols['category']].astype(str).unique())
        return snapshot, report

    def export_with_report(self, destination, styles, visible_columns, messages, report, include_report=False):
        """Stage outputs first; report failures never silently count as a successful export."""
        extension = os.path.splitext(destination)[1].lower()
        target = os.path.abspath(destination)
        report_path = os.path.splitext(target)[0] + '_excluded_rows.csv'
        include_report = include_report or not report.empty
        if os.path.normcase(os.path.realpath(target)) == os.path.normcase(os.path.realpath(self.current_filepath)):
            messages.put(('error', 'The original input file cannot be overwritten.'))
            return
        if include_report and os.path.normcase(os.path.realpath(report_path)) == os.path.normcase(os.path.realpath(self.current_filepath)):
            messages.put(('error', 'The exclusion report would overwrite the original input file. Choose another output name.'))
            return
        folder = None
        preserve_recovery = False
        try:
            check_cancel(self.cancel_event)
            folder = tempfile.mkdtemp(prefix='.spatialpin_', dir=os.path.dirname(target))
            if folder:
                staged_map = os.path.join(folder, os.path.basename(target))
                staged_report = os.path.join(folder, os.path.basename(report_path))
                if include_report:
                    report.to_csv(staged_report, index=False, encoding='utf-8-sig')
                class Relay:
                    error = None
                    done = False
                    cancelled = False
                    def put(relay, event):
                        if event[0] == 'cancelled':
                            relay.cancelled = True
                        elif event[0] == 'error':
                            relay.error = event[1]
                        elif event[0] == 'done':
                            relay.done = True
                        else:
                            messages.put(event)
                relay = Relay()
                exporter = self.export_to_html if extension == '.html' else self.export_to_kml
                exporter(staged_map, styles, visible_columns, relay)
                if relay.cancelled:
                    raise OperationCancelled()
                check_cancel(self.cancel_event)
                if relay.error or not relay.done:
                    raise ValueError(relay.error or 'The export did not finish.')
                # Back up existing outputs so a failure committing either file can be rolled back.
                commits = [(staged_map, target)]
                if include_report:
                    commits.append((staged_report, report_path))
                check_cancel(self.cancel_event)
                messages.put(('status', 'Finishing export…'))
                backups, installed = [], []
                try:
                    for index, (_, final) in enumerate(commits):
                        if os.path.exists(final):
                            backup = os.path.join(folder, 'backup_' + str(index))
                            os.replace(final, backup)
                            backups.append((backup, final))
                    for staged, final in commits:
                        os.replace(staged, final)
                        installed.append(final)
                except Exception as error:
                    try:
                        for final in installed:
                            os.remove(final)
                        for backup, final in backups:
                            os.replace(backup, final)
                    except Exception as recovery_error:
                        preserve_recovery = True
                        raise OSError(f"Could not restore existing outputs. Recovery files are preserved at {folder}. {recovery_error}") from error
                    raise
            messages.put(('done', target))
        except OperationCancelled:
            messages.put(('cancelled', None))
        except Exception as error:
            LOG.exception('Export failed')
            messages.put(('error', str(error)))
        finally:
            if folder and not preserve_recovery:
                try:
                    for filename in os.listdir(folder):
                        os.remove(os.path.join(folder, filename))
                    os.rmdir(folder)
                except OSError:
                    pass

    def snapshot(self):
        result = KMLProcessor()
        result.cancel_event = self.cancel_event
        result.df = self.df.copy(deep=True)
        result.current_filepath = self.current_filepath
        result.target_cols = self.target_cols.copy()
        result.unique_categories = self.unique_categories.copy()
        result.sheet_names = self.sheet_names.copy()
        result.current_sheet = self.current_sheet
        result.import_options = self.import_options.copy()
        result.import_details = copy.deepcopy(self.import_details)
        result.source_rows = self.source_rows.copy()
        result.web_basemap = self.web_basemap
        return result

    def load_and_parse_file(self, filepath, allow_unmapped=False, sheet_name=0, header_row=1, delimiter="auto", encoding="auto"):
        candidate = KMLProcessor()
        success, message = candidate._load_and_parse_file(filepath, allow_unmapped, sheet_name, header_row, delimiter, encoding)
        if success:
            self.df = candidate.df
            self.current_filepath = candidate.current_filepath
            self.target_cols = candidate.target_cols
            self.unique_categories = candidate.unique_categories
            self.sheet_names = candidate.sheet_names
            self.current_sheet = candidate.current_sheet
            self.import_options = candidate.import_options
            self.import_details = candidate.import_details
            self.source_rows = candidate.source_rows
        return success, message

    def _load_and_parse_file(self, filepath, allow_unmapped=False, sheet_name=0, header_row=1, delimiter="auto", encoding="auto"):
        try:
            self.current_filepath = filepath
            ext = os.path.splitext(filepath)[1].lower()
            
            self.import_options = dict(header_row=header_row, delimiter=delimiter, encoding=encoding)
            if ext in ['.xlsx', '.xls', '.csv', '.txt']:
                self.df, self.source_rows, self.import_details = self.read_table(filepath, sheet_name, header_row, delimiter, encoding)
                self.sheet_names = self.import_details.get('sheets', [])
                self.current_sheet = self.import_details.get('sheet')
            elif ext in ['.geojson', '.json']:
                with open(filepath, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                rows = []
                for feat in data.get('features', []):
                    prop = dict(feat.get('properties') or {})
                    geom = feat.get('geometry', {})
                    if not geom or geom.get('type') != 'Point':
                        return False, 'Only GeoJSON Point features are supported; no data was loaded.'
                    if geom and geom.get('type') == 'Point':
                        coords = geom.get('coordinates', [])
                        prop['longitude'] = coords[0] if len(coords) > 0 else None
                        prop['latitude'] = coords[1] if len(coords) > 1 else None
                    rows.append(prop)
                self.df = pd.DataFrame(rows)
                self.source_rows = list(range(1, len(rows) + 1))
                self.import_details = {"row_kind": "GeoJSON feature"}
            else:
                return False, "Unsupported format."

            if self.df.empty:
                return False, "The file contains no data rows."
            self.df.columns = [str(c) for c in self.df.columns]
            if self.df.columns.duplicated().any():
                return False, "Column names must be unique."
            if any(c.startswith("__parsed") for c in self.df.columns):
                return False, "Column names beginning with __parsed are reserved. Please rename them."
            raw_cols = list(self.df.columns)
            aliases = {
                "point name": ["point name", "site name", "name", "id", "site_id", "site id", "label", "point_id", "Nominal name", "Nominal"],
                "latitude": ["latitude", "lat", "Nominal lat", "y_coord", "y"],
                "longitude": ["longitude", "lon", "long", "Nominal long", "Nominal lon", "lng", "x_coord", "x"],
                "category": ["status", "category", "type", "class", "group"],
            }

            for key in ["point name", "latitude", "longitude"]:
                found_col = self._find_best_column(raw_cols, aliases[key])
                if found_col:
                    self.target_cols[key] = found_col
                elif not allow_unmapped:
                    return False, f"Missing required column for '{key.title()}'. Choose a column in the mapping controls."

            found_cat = self._find_best_column(raw_cols, aliases["category"])
            if found_cat:
                self.target_cols["category"] = found_cat
            else:
                if "Category" not in self.df.columns:
                    self.df["Category"] = "All Points"
                self.target_cols["category"] = "Category"

            try:
                self.apply_mapping(self.target_cols)
                return True, "Loaded. Review the column mapping above the workspace."
            except ValueError as error:
                if not allow_unmapped:
                    return False, str(error)
                self.df["__parsed_lat"] = float("nan")
                self.df["__parsed_lon"] = float("nan")
                self.unique_categories = []
                return True, "Choose distinct Name, Latitude and Longitude columns, then Apply mapping."
        except Exception as e:
            LOG.exception("Import failed")
            return False, f"Failed to load or parse file: {str(e)}"

    def apply_mapping(self, mapping):
        required = ("point name", "latitude", "longitude")
        if any(not mapping.get(key) for key in required):
            raise ValueError("Choose Name, Latitude and Longitude columns.")
        selected = [value for value in mapping.values() if value is not None]
        if len(selected) != len(set(selected)):
            raise ValueError("Each role needs a different column.")
        if any(value not in self.df.columns or value.startswith("__parsed") for value in selected):
            raise ValueError("A selected column is unavailable in this file.")
        frame = self.df.copy(deep=True)
        mapping = mapping.copy()
        if mapping.get("category") is None:
            name = "Map category"
            while name in frame.columns:
                name += "_"
            frame[name] = "All Points"
            mapping["category"] = name
        for axis, internal in (("latitude", "__parsed_lat"), ("longitude", "__parsed_lon")):
            frame[internal] = frame[mapping[axis]].apply(lambda value: parse_coordinate(value, axis))
        frame[mapping["category"]] = frame[mapping["category"]].fillna("Uncategorized").astype(str)
        frame[mapping["point name"]] = frame[mapping["point name"]].fillna("Unnamed Point")
        self.df = frame
        self.target_cols = mapping
        self.unique_categories = sorted(frame[mapping["category"]].unique().tolist())

    def generate_html_description(self, row_dict, visible_cols):
        html = '<table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse; width: 300px;">'
        html += '<tr style="background-color: #f2f2f2; color: black;"><th>Attribute</th><th>Value</th></tr>'
        for col in visible_cols:
            if col not in row_dict:
                continue
            value = row_dict[col]
            if isinstance(value, (list, dict)):
                value = json.dumps(value, ensure_ascii=False, default=str)
            elif pd.isna(value):
                continue
            html += f"<tr><td style='color: black;'><b>{safe_web_text(col)}</b></td><td style='color: black;'>{safe_web_text(value)}</td></tr>"
        html += "</table>"
        return html

    def build_web_map(self, style_config, visible_cols, indices=None, basemap='OpenStreetMap', cluster=False, progress=None):
        valid = self.df.dropna(subset=['__parsed_lat', '__parsed_lon'])
        if valid.empty:
            raise ValueError('There are no valid coordinates to preview.')
        bounds = geographic_bounds(valid)
        (south, west), (north, east) = bounds
        result = folium.Map(location=[(south+north)/2, (west+east)/2], zoom_start=12, tiles=basemap)
        if south != north or west != east:
            result.fit_bounds(bounds, padding=(24,24), max_zoom=15)
        container = MarkerCluster().add_to(result) if cluster else result
        rows = list(valid.index) if indices is None else list(indices)
        points, markers = [], []
        interval = max(1, len(rows)//100)
        for offset, index in enumerate(rows):
            check_cancel(self.cancel_event)
            row = valid.loc[index].to_dict()
            category = str(row[self.target_cols['category']])
            style = copy.deepcopy(style_config.get(category, dict(icolor='#087f8c', iscale=1.2, opacity=1.0)))
            point = dict(row=int(index), lat=float(row['__parsed_lat']), lon=float(row['__parsed_lon']),
                         name=str(row[self.target_cols['point name']]), category=category, style=style,
                         popup=self.generate_html_description(row, visible_cols))
            marker = web_point([point['lat'], display_longitude(point['lon'], west)], point['name'], style, point['popup'])
            marker.add_to(container)
            markers.append((int(index), marker.get_name()))
            points.append(point)
            if progress and (offset % interval == 0 or offset == len(rows)-1):
                progress.put(('progress', int((offset+1)/len(rows)*100)))
        return result, points, markers, bounds

    def export_to_html(self, save_path, style_config, visible_cols, export_queue):
        try:
            if self.df is None or self.df.empty or self.has_invalid_coordinates():
                raise ValueError("No valid dataset to export. Fix coordinate errors first.")
            if os.path.normcase(os.path.realpath(save_path)) == os.path.normcase(os.path.realpath(self.current_filepath)):
                raise ValueError("The original input file cannot be overwritten.")
            m, _, _, _ = self.build_web_map(style_config, visible_cols, basemap=self.web_basemap,
                                           cluster=True, progress=export_queue)
            export_queue.put(("status", "Writing HTML to disk..."))
            m.save(save_path)
            export_queue.put(("done", save_path))
        except OperationCancelled:
            export_queue.put(("cancelled", None))
        except Exception as e:
            LOG.exception("Map generation failed")
            export_queue.put(("error", str(e)))

    def export_to_kml(self, save_path, style_config, visible_cols, export_queue):
        try:
            if self.df is None or self.df.empty or self.has_invalid_coordinates():
                raise ValueError("No valid dataset to export. Fix coordinate errors first.")
            if os.path.normcase(os.path.realpath(save_path)) == os.path.normcase(os.path.realpath(self.current_filepath)):
                raise ValueError("The original input file cannot be overwritten.")
            is_kmz = save_path.lower().endswith(".kmz")
            total_rows = len(self.df)
            update_interval = max(1, total_rows // 100)
            KML_NS = "http://www.opengis.net/kml/2.2"
            ET.register_namespace('', KML_NS)
            kml_root = ET.Element(f"{{{KML_NS}}}kml")
            doc_el = ET.SubElement(kml_root, f"{{{KML_NS}}}Document")
            name_el = ET.SubElement(doc_el, f"{{{KML_NS}}}name")
            name_el.text = os.path.basename(save_path)
            custom_images_to_copy = {}

            for cat, conf in style_config.items():
                s_id = f"style_{hash(cat)}"
                style_el = ET.SubElement(doc_el, f"{{{KML_NS}}}Style", id=s_id)
                icon_style = ET.SubElement(style_el, f"{{{KML_NS}}}IconStyle")
                ET.SubElement(icon_style, f"{{{KML_NS}}}color").text = hex_to_kml_color(conf["icolor"], conf["opacity"])
                ET.SubElement(icon_style, f"{{{KML_NS}}}scale").text = str(conf["iscale"])
                icon_el = ET.SubElement(icon_style, f"{{{KML_NS}}}Icon")
                
                check_cancel(self.cancel_event)
                if conf["shape"] == "Custom Image":
                    if not is_kmz:
                        raise ValueError('Use KMZ to include custom images.')
                    data = png_icon(conf.get('custom_path'))
                    img_name = 'images/' + hashlib.sha256(data).hexdigest() + '.png'
                    custom_images_to_copy[img_name] = data
                    href = img_name
                elif is_kmz:
                    shape = conf.get('shape', 'Pushpin')
                    href = 'images/standard_' + shape.lower() + '.png'
                    custom_images_to_copy[href] = base64.b64decode(BUNDLED_ICONS.get(shape,BUNDLED_ICONS['Pushpin']))
                else:
                    href = GE_ICONS.get(conf['shape'], GE_ICONS['Pushpin'])
                ET.SubElement(icon_el, f"{{{KML_NS}}}href").text = href
                if conf['shape'] in ('Pushpin','Paddle'):
                    ET.SubElement(icon_style, f"{{{KML_NS}}}hotSpot", x='0.5',y='0',xunits='fraction',yunits='fraction')

                label_style = ET.SubElement(style_el, f"{{{KML_NS}}}LabelStyle")
                ET.SubElement(label_style, f"{{{KML_NS}}}color").text = hex_to_kml_color(conf["lcolor"])
                ET.SubElement(label_style, f"{{{KML_NS}}}scale").text = str(conf["lscale"])

            records = self.df.to_dict('records')
            for idx, row in enumerate(records):
                check_cancel(self.cancel_event)
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
                    for name, data in custom_images_to_copy.items():
                        check_cancel(self.cancel_event)
                        kmz.writestr(name, data)
            else:
                with open(save_path, "wb") as f:
                    f.write(xml_bytes)
            export_queue.put(("done", save_path))
        except OperationCancelled:
            export_queue.put(("cancelled", None))
        except Exception as e:
            LOG.exception("Map generation failed")
            export_queue.put(("error", str(e)))

class ImportOptionsDialog(QDialog):
    def __init__(self, parent, path, options=None, sheet_name=0):
        super().__init__(parent)
        self.path = path
        self.setWindowTitle("Import options — " + os.path.basename(path))
        self.resize(790, 500)
        self.setMinimumSize(600, 400)
        self.last_preview_settings = None
        layout = QVBoxLayout(self)
        file_label = QLabel(os.path.abspath(path))
        file_label.setTextFormat(Qt.TextFormat.PlainText)
        file_label.setWordWrap(True)
        layout.addWidget(file_label)
        form = QFormLayout()
        options = options or {}
        self.header_row = QSpinBox()
        self.header_row.setRange(1, 10000)
        self.header_row.setValue(options.get('header_row', 1))
        form.addRow("Header row (1 = first row)", self.header_row)
        self.sheet = QComboBox()
        excel = os.path.splitext(path)[1].lower() in ('.xlsx', '.xls')
        self.excel = excel
        self.initial_sheet = sheet_name
        if excel:
            if isinstance(sheet_name, str): self.sheet.addItem(sheet_name)
            form.addRow('Worksheet', self.sheet)
        self.delimiter = QComboBox()
        for label, value in [('Detect automatically', 'auto'), ('Comma (,)', ','), ('Semicolon (;)', ';'), ('Tab', '\t'), ('Pipe (|)', '|')]:
            self.delimiter.addItem(label, value)
        self.delimiter.setCurrentIndex(max(0, self.delimiter.findData(options.get('delimiter', 'auto'))))
        self.encoding = QComboBox()
        for label, value in [('Detect automatically', 'auto'), ('UTF-8', 'utf-8-sig'), ('UTF-16', 'utf-16'), ('UTF-32', 'utf-32'), ('Windows Arabic (1256)', 'cp1256'), ('Windows Western (1252)', 'cp1252'), ('Latin-1', 'latin-1')]:
            self.encoding.addItem(label, value)
        self.encoding.setCurrentIndex(max(0, self.encoding.findData(options.get('encoding', 'auto'))))
        if not excel:
            form.addRow("Delimiter", self.delimiter)
            form.addRow("Text encoding", self.encoding)
        layout.addLayout(form)
        self.status = QLabel("Preview the import before applying settings.")
        self.status.setWordWrap(True)
        layout.addWidget(self.status)
        self.preview = QTableWidget()
        self.preview.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        layout.addWidget(self.preview, 1)
        controls = QHBoxLayout()
        refresh = QPushButton("Preview first 20 rows")
        refresh.clicked.connect(self.refresh_preview)
        controls.addWidget(refresh)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.button(QDialogButtonBox.StandardButton.Ok).setText("Import")
        self.import_button = buttons.button(QDialogButtonBox.StandardButton.Ok)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        controls.addWidget(buttons)
        layout.addLayout(controls)
        for control in (self.header_row, self.sheet, self.delimiter, self.encoding):
            signal = control.valueChanged if isinstance(control, QSpinBox) else control.currentIndexChanged
            signal.connect(self.settings_changed)
        self.preview_queue = queue.Queue()
        self.preview_generation = 0
        self.preview_busy = False
        self.preview_poll = QTimer(self)
        self.preview_poll.timeout.connect(self.check_preview_queue)
        self.finished.connect(self.preview_poll.stop)
        self.refresh_preview()

    def settings(self):
        return dict(header_row=self.header_row.value(), delimiter=self.delimiter.currentData(), encoding=self.encoding.currentData())

    def settings_changed(self):
        self.preview_generation += 1
        self.import_button.setEnabled(False)
        self.status.setText("Settings changed. Preview again before importing.")

    def refresh_preview(self):
        if self.preview_busy:
            self.status.setText('Reading file… wait, then preview the new settings.')
            return
        self.preview_generation += 1
        generation = self.preview_generation
        self.preview_busy = True
        self.import_button.setEnabled(False)
        self.status.setText('Reading preview…')
        options, sheet, path = self.settings(), self.sheet.currentText() or self.initial_sheet, self.path
        def work():
            try:
                frame, rows, details = KMLProcessor.read_table(path,sheet,**options)
                self.preview_queue.put((generation,(frame.head(20),rows[:20],details,len(frame)),None))
            except Exception as error:
                LOG.exception('Import preview failed')
                self.preview_queue.put((generation,None,str(error)))
        self.preview_worker = threading.Thread(target=work,daemon=True)
        self.preview_poll.start(60)
        self.preview_worker.start()

    def check_preview_queue(self):
        try: generation,result,error = self.preview_queue.get_nowait()
        except queue.Empty: return
        self.preview_poll.stop()
        self.preview_busy = False
        if generation != self.preview_generation:
            self.status.setText('Settings changed. Preview again before importing.')
            return
        self.preview.clear()
        self.preview.setRowCount(0)
        if error:
            self.status.setText(error)
            self.import_button.setEnabled(False)
            return
        frame,rows,details,total = result
        if self.excel:
            self.sheet.blockSignals(True)
            self.sheet.clear()
            self.sheet.addItems(details['sheets'])
            self.sheet.setCurrentText(details['sheet'])
            self.sheet.blockSignals(False)
        self.preview.setColumnCount(len(frame.columns))
        self.preview.setHorizontalHeaderLabels(list(frame.columns))
        self.preview.setRowCount(len(frame))
        self.preview.setVerticalHeaderLabels([str(row) for row in rows])
        for index,record in enumerate(frame.to_dict('records')):
            for column,value in enumerate(record.values()):
                self.preview.setItem(index,column,QTableWidgetItem(display_text(value)))
        self.preview.resizeColumnsToContents()
        description = f"{total:,} rows · {len(frame.columns)} columns · header row {details['header_row']}"
        if 'encoding' in details: description += f" · {details['encoding']} · delimiter {repr(details['delimiter'])}"
        if details.get('encoding_guessed'): description += ' · Encoding is a guess: check accented/Arabic text.'
        self.status.setText(description)
        self.import_button.setEnabled(True)


class MapBridge(QObject):
    def __init__(self, window):
        super().__init__(window)
        self.window = window

    @pyqtSlot(int, int)
    def selectRow(self, row, generation):
        if generation == self.window.map_generation:
            self.window.select_map_row(row)

    @pyqtSlot(int)
    def ready(self, generation):
        if generation == self.window.map_generation:
            self.window.online_ready = True
            self.window.browser_failure = ''
            self.window.update_preview_presentation()

    @pyqtSlot(int, str)
    def tiles(self, generation, state):
        if generation == self.window.map_generation:
            self.window.tile_state = state
            if state in ('available', 'partial'):
                self.window.map_timeout.stop()
                self.window.online_ready = True
                self.window.browser_failure = ''
            self.window.update_preview_presentation()

    @pyqtSlot(int, str)
    def failed(self, generation, message):
        if generation == self.window.map_generation:
            self.window.map_failed(message)


def application_icon():
    """Draw the SitePin mark locally at native icon sizes; no asset files needed."""
    icon = QIcon()
    for size in (16, 24, 32, 40, 48, 64, 128, 256):
        pixmap = QPixmap(size, size)
        pixmap.fill(Qt.GlobalColor.transparent)
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.scale(size / 64, size / 64)
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(QColor('#087f8c'))
        painter.drawRoundedRect(QRectF(1, 1, 62, 62), 14, 14)
        pin = QPainterPath()
        pin.moveTo(32, 55)
        pin.cubicTo(26, 47, 14, 34, 14, 25)
        pin.cubicTo(14, 3, 50, 3, 50, 25)
        pin.cubicTo(50, 34, 38, 47, 32, 55)
        painter.setBrush(QColor('white'))
        painter.drawPath(pin)
        painter.setBrush(QColor('#e09931'))
        painter.drawEllipse(QRectF(24, 17, 16, 16))
        painter.end()
        icon.addPixmap(pixmap)
    return icon


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.resize(1320, 880)
        self.setMinimumSize(900, 640)
        self.setAcceptDrops(True)
        self.processor = KMLProcessor()
        self.styles = {}
        self.dirty = False
        self.undo_stack, self.redo_stack = [], []
        self.mapping_ready = False
        self.mapping_pending = False
        self.pending_import_path = None
        self.duplicate_notes = {}
        self.export_report_path = None
        self.last_excluded_count = 0
        self.loading_data = False
        self.compact_mode = False
        self.restore_map_after_compact = False
        self.issue_cache, self.search_cache = {}, {}
        self.import_busy = False
        self.import_queue = queue.Queue()
        self.import_cancel = threading.Event()
        self.cancel_event = threading.Event()
        self.close_after_job = False
        self.preset_icons = tempfile.TemporaryDirectory(prefix='spatialpin_icons_')
        self.import_timer = QTimer(self)
        self.import_timer.timeout.connect(self.check_import_queue)
        self.export_queue = queue.Queue()
        self.export_busy = False
        self.output_path = None
        self.map_ready = False
        self.online_ready = False
        self.tile_state = 'pending'
        self.preview_dirty = False
        self.browser_failure = ''
        self.preview_categories = 0
        self.total_categories = 0
        self.map_generation = 0
        self.preview_rows = set()
        self.preview_summary = ""
        self.preview_documents = []
        self.browser_path = None
        self.preview_file = tempfile.TemporaryDirectory(prefix="spatialpin_preview_")
        self.preview_timer = QTimer(self)
        self.preview_timer.setSingleShot(True)
        self.preview_timer.setInterval(550)
        self.preview_timer.timeout.connect(self.plot_map)
        self.map_timeout = QTimer(self)
        self.map_timeout.setSingleShot(True)
        self.map_timeout.setInterval(18000)
        self.map_timeout.timeout.connect(self.map_failed)
        self.queue_timer = QTimer(self)
        self.queue_timer.timeout.connect(self.check_queue)
        self.setWindowIcon(application_icon())
        self.init_ui()
        self.map_panel.hide()
        self.update_data_status()
        self.clear_map()
        self.menuBar().addMenu(self.help_menu)

    @staticmethod
    def button(text, action, primary=False):
        button = QPushButton(text)
        button.setCursor(Qt.CursorShape.PointingHandCursor)
        button.clicked.connect(action)
        if primary:
            button.setObjectName("Primary")
        return button

    @staticmethod
    def card():
        frame = QFrame()
        frame.setObjectName("Card")
        return frame

    @staticmethod
    def heading(text):
        label = QLabel(text)
        label.setObjectName("Heading")
        return label

    def init_ui(self):
        self.setStyleSheet("""
            QMainWindow {background:#f3f6fa;}
            QWidget {font-family:'Segoe UI'; font-size:10pt; color:#24344a;}
            QFrame#Card {background:white; border:1px solid #dbe3ed; border-radius:10px;}
            QLabel {background:transparent;}
            QLabel#Heading {font-weight:600; font-size:11pt; color:#14263e;}
            QLabel#Muted {color:#61738b; font-size:9pt;}
            QPushButton, QToolButton {background:#fff; border:1px solid #cbd6e3; border-radius:5px; padding:6px 10px;}
            QPushButton:hover, QToolButton:hover {background:#eaf3f6; border-color:#7caeb5;}
            QPushButton:disabled {color:#8b99aa; background:#f1f4f8; border-color:#e2e8f0;}
            QPushButton#Primary {background:#087f8c; color:white; border-color:#087f8c; font-weight:600;}
            QPushButton#Primary:disabled {background:#b7cbd1; border-color:#b7cbd1;}
            QComboBox, QLineEdit, QDoubleSpinBox {background:white; border:1px solid #cbd6e3; border-radius:4px; padding:5px; min-height:18px;}
            QComboBox:disabled {background:#f1f4f8; color:#8b99aa;}
            QComboBox QAbstractItemView {background:white; selection-background-color:#dff2f3; color:#24344a;}
            QTableWidget, QListWidget {background:white; alternate-background-color:#f8fafc; border:1px solid #e2e8f0; border-radius:4px; selection-background-color:#d9eef2; selection-color:#123b49;}
            QHeaderView::section {background:#f3f6fa; padding:7px; border:none; border-bottom:1px solid #dbe3ed; font-weight:600;}
            QSplitter::handle {background:#e0e7ef; border-radius:2px;}
            QScrollArea {border:none; background:transparent;}
            QProgressBar {border:none; background:#e3ebf0; border-radius:3px; max-height:6px;}
            QProgressBar::chunk {background:#087f8c; border-radius:3px;}
            QCheckBox {spacing:6px;}
        """)
        root = QWidget()
        self.setCentralWidget(root)
        layout = QVBoxLayout(root)
        layout.setContentsMargins(16, 12, 16, 12)
        layout.setSpacing(10)
        brand = QHBoxLayout()
        monogram = QLabel()
        monogram.setPixmap(self.windowIcon().pixmap(40, 40))
        monogram.setAlignment(Qt.AlignmentFlag.AlignCenter)
        monogram.setFixedSize(40, 40)
        monogram.setStyleSheet("background:#087f8c;color:white;border-radius:9px;font-size:15pt;font-weight:700;")
        brand.addWidget(monogram)
        titles = QVBoxLayout()
        titles.setSpacing(0)
        title = QLabel(APP_NAME)
        title.setStyleSheet("font-size:17pt;font-weight:600;color:#14263e;")
        titles.addWidget(title)
        subtitle = QLabel(APP_DESCRIPTION)
        subtitle.setObjectName("Muted")
        titles.addWidget(subtitle)
        brand.addLayout(titles)
        brand.addStretch()
        self.btn_appearance = self.button("Hide appearance", self.toggle_appearance)
        self.btn_map = self.button("Show map", self.toggle_map)
        brand.addWidget(self.btn_appearance)
        brand.addWidget(self.btn_map)
        self.help_menu = QMenu("Help", self)
        self.help_menu.addAction("About SitePin", self.show_help)
        self.help_menu.addAction("Quick Guide", self.show_quick_guide)
        self.help_menu.addAction("Contact & Support", self.show_support)
        self.help_menu.addSeparator()
        self.help_menu.addAction("Open diagnostic folder", self.open_logs)
        self.help_menu.addAction("Load legacy styles (JSON)", self.load_styles)
        self.btn_help = QToolButton()
        self.btn_help.setText("Help")
        self.btn_help.setAccessibleName("Help menu")
        self.btn_help.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        self.btn_help.setMenu(self.help_menu)
        self.btn_help.setCursor(Qt.CursorShape.PointingHandCursor)
        brand.addWidget(self.btn_help)
        layout.addLayout(brand)

        source = self.card()
        self.setup_panel = source
        source_layout = QVBoxLayout(source)
        source_layout.setContentsMargins(12, 10, 12, 10)
        file_bar = QHBoxLayout()
        file_bar.addWidget(self.button("Open file…", self.load_file, True))
        self.file_label = QLabel("Drop an Excel, CSV or GeoJSON file here")
        self.file_label.setTextFormat(Qt.TextFormat.PlainText)
        self.file_label.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred)
        file_bar.addWidget(self.file_label, 1)
        self.btn_import_options = self.button("Import options…", self.show_import_options)
        file_bar.addWidget(self.btn_import_options)
        file_bar.addWidget(self.button("Sample file…", self.save_sample_file))
        file_bar.addWidget(QLabel("Sheet"))
        self.sheet_combo = QComboBox()
        self.sheet_combo.setMinimumWidth(130)
        self.sheet_combo.setMaximumWidth(210)
        self.sheet_combo.currentTextChanged.connect(self.change_sheet)
        file_bar.addWidget(self.sheet_combo)
        wgs = QLabel("WGS84 · degrees")
        wgs.setObjectName("Muted")
        file_bar.addWidget(wgs)
        source_layout.addLayout(file_bar)
        mapping = QGridLayout()
        mapping.setHorizontalSpacing(12)
        self.mapping_combos = {}
        for index, (key, label) in enumerate((("point name", "Point name"), ("latitude", "Latitude"),
                                             ("longitude", "Longitude"), ("category", "Category"))):
            mapping.addWidget(QLabel(label), 0, index)
            combo = QComboBox()
            combo.setMinimumWidth(110)
            combo.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon)
            combo.setMinimumContentsLength(12)
            combo.currentIndexChanged.connect(self.mapping_changed)
            self.mapping_combos[key] = combo
            mapping.addWidget(combo, 1, index)
            mapping.setColumnStretch(index, 1)
        self.btn_apply_mapping = self.button("Apply mapping", self.apply_mapping)
        mapping.addWidget(self.btn_apply_mapping, 1, 4)
        source_layout.addLayout(mapping)
        self.mapping_status = QLabel("Column suggestions appear when you open a file. Your original stays unchanged.")
        self.mapping_status.setObjectName("Muted")
        self.mapping_status.setWordWrap(True)
        source_layout.addWidget(self.mapping_status)
        layout.addWidget(source)

        self.workspace = QSplitter(Qt.Orientation.Horizontal)
        self.workspace.setChildrenCollapsible(False)
        self.appearance_panel = self.build_appearance()
        self.workspace.addWidget(self.appearance_panel)
        self.data_splitter = QSplitter(Qt.Orientation.Vertical)
        self.data_splitter.setChildrenCollapsible(False)
        self.map_panel = self.build_map_panel()
        self.data_panel = self.build_data_panel()
        self.data_splitter.addWidget(self.map_panel)
        self.data_splitter.addWidget(self.data_panel)
        self.data_splitter.setSizes([350, 250])
        self.workspace.addWidget(self.data_splitter)
        self.workspace.setStretchFactor(1, 1)
        self.workspace.setSizes([280, 1000])
        layout.addWidget(self.workspace, 1)

        footer = self.card()
        footer_layout = QVBoxLayout(footer)
        footer_layout.setContentsMargins(12, 8, 12, 8)
        self.data_status = QLabel()
        self.data_status.setWordWrap(True)
        status_bar = QHBoxLayout()
        status_bar.addWidget(self.data_status, 1)
        self.valid_only = QCheckBox("Valid rows only")
        self.valid_only.setToolTip("Explicitly exclude invalid coordinates and save an exclusion report next to the map.")
        self.valid_only.toggled.connect(self.update_data_status)
        status_bar.addWidget(self.valid_only)
        footer_layout.addLayout(status_bar)
        actions = QHBoxLayout()
        self.btn_save_copy = self.button("Save corrected copy…", self.save_corrected_copy)
        self.btn_save_copy.setToolTip("Save data to a separate file; formatting, formulas and other sheets are not copied.")
        actions.addWidget(self.btn_save_copy)
        self.status_lbl = QLabel("Ready when you are")
        self.status_lbl.setObjectName("Muted")
        self.status_lbl.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred)
        actions.addWidget(self.status_lbl, 1)
        self.btn_open_output = self.button("Open file", self.open_output)
        self.btn_open_output.setEnabled(False)
        actions.addWidget(self.btn_open_output)
        self.btn_open_folder = self.button('Folder', self.open_output_folder)
        self.btn_open_folder.setEnabled(False)
        actions.addWidget(self.btn_open_folder)
        self.btn_cancel = self.button('Cancel', self.cancel_job)
        self.btn_cancel.hide()
        actions.addWidget(self.btn_cancel)
        actions.addWidget(QLabel("Format"))
        self.format_combo = QComboBox()
        self.format_combo.addItems(["KMZ", "KML", "HTML"])
        self.format_combo.currentIndexChanged.connect(self.update_data_status)
        actions.addWidget(self.format_combo)
        self.btn_export = self.button("Export map…", self.start_export, True)
        actions.addWidget(self.btn_export)
        footer_layout.addLayout(actions)
        self.progress = QProgressBar()
        self.progress.setTextVisible(False)
        self.progress.setValue(0)
        footer_layout.addWidget(self.progress)
        layout.addWidget(footer)

    def collapsible(self, title, parent_layout):
        toggle = QToolButton()
        toggle.setText(title)
        toggle.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        toggle.setArrowType(Qt.ArrowType.RightArrow)
        toggle.setCheckable(True)
        toggle.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        body = QWidget()
        body.hide()
        def change(checked):
            body.setVisible(checked)
            toggle.setArrowType(Qt.ArrowType.DownArrow if checked else Qt.ArrowType.RightArrow)
        toggle.toggled.connect(change)
        parent_layout.addWidget(toggle)
        parent_layout.addWidget(body)
        return toggle, body

    def build_appearance(self):
        panel = self.card()
        panel.setMinimumWidth(260)
        panel.setMaximumWidth(400)
        outer = QVBoxLayout(panel)
        outer.addWidget(self.heading("Appearance"))
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        outer.addWidget(scroll, 1)
        contents = QWidget()
        contents.setStyleSheet("background:white;")
        content = QVBoxLayout(contents)
        content.setContentsMargins(0, 0, 3, 0)
        content.setSpacing(8)
        scroll.setWidget(contents)
        self.cb_palette = QComboBox()
        self.cb_palette.addItems(["Distinct colors", "Blue gradient", "Red to green"])
        content.addWidget(self.cb_palette)
        content.addWidget(self.button("Apply palette", self.apply_palette))
        label = QLabel("Select a category to edit its style")
        label.setObjectName("Muted")
        label.setWordWrap(True)
        content.addWidget(label)
        self.category_list = QListWidget()
        self.category_list.setMinimumHeight(100)
        self.category_list.setMaximumHeight(190)
        self.category_list.currentItemChanged.connect(self.show_category_style)
        content.addWidget(self.category_list)
        self.btn_color = self.button("Point color", lambda: self.pick_style_color("icolor"))
        content.addWidget(self.btn_color)
        self.advanced_toggle, self.advanced_body = self.collapsible("Advanced styling", content)
        form = QFormLayout(self.advanced_body)
        form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.AllNonFixedFieldsGrow)
        form.setContentsMargins(2, 4, 2, 4)
        self.shape_combo = QComboBox()
        self.shape_combo.addItems(list(GE_ICONS) + ["Custom Image"])
        self.shape_combo.currentTextChanged.connect(lambda value: self.change_style("shape", value))
        form.addRow("Icon", self.shape_combo)
        self.btn_custom = self.button("Choose image…", self.choose_custom_icon)
        form.addRow(self.btn_custom)
        self.style_spins = {}
        for key, label, low, high in (("opacity", "Opacity", 0.1, 1.0), ("iscale", "Point size", 0.5, 3.0),
                                       ("lscale", "Label size", 0.0, 3.0)):
            spin = QDoubleSpinBox()
            spin.setRange(low, high)
            spin.setSingleStep(0.1)
            spin.valueChanged.connect(lambda value, key=key: self.change_style(key, value))
            self.style_spins[key] = spin
            form.addRow(label, spin)
        self.btn_label_color = self.button("Label color", lambda: self.pick_style_color("lcolor"))
        form.addRow(self.btn_label_color)
        hint = QLabel("Google Earth only: icon shape, custom image and label styling. Preview and HTML use circles with the chosen color, size and opacity.")
        hint.setWordWrap(True)
        hint.setObjectName("Muted")
        form.addRow(hint)
        self.columns_toggle, self.columns_body = self.collapsible("Popup fields", content)
        fields = QVBoxLayout(self.columns_body)
        fields.setContentsMargins(0, 3, 0, 3)
        self.list_columns = QListWidget()
        self.list_columns.setMinimumHeight(140)
        self.list_columns.setMaximumHeight(220)
        self.list_columns.itemChanged.connect(self.schedule_preview)
        fields.addWidget(self.list_columns)
        content.addStretch()
        presets = QHBoxLayout()
        presets.addWidget(self.button("Load preset…", self.load_team_preset))
        presets.addWidget(self.button("Save preset…", self.save_team_preset))
        outer.addLayout(presets)
        return panel

    def build_map_panel(self):
        panel = self.card()
        panel.setMinimumHeight(280)
        layout = QVBoxLayout(panel)
        bar = QHBoxLayout()
        bar.addWidget(self.heading("Map preview"))
        self.preview_badge = QLabel('No data')
        self.preview_badge.setStyleSheet('color:#61738b;font-size:9pt;')

        bar.addStretch()
        self.basemap_combo = QComboBox()
        self.basemap_combo.addItems(["Coordinate overview", "OpenStreetMap"])
        self.basemap_combo.setMaximumWidth(150)
        self.basemap_combo.setToolTip("Online basemap is also used by HTML export. Coordinate overview is an offline preview; HTML then uses OpenStreetMap.")
        self.basemap_combo.currentIndexChanged.connect(self.schedule_preview)
        bar.addWidget(self.basemap_combo)
        self.auto_preview = QCheckBox("Auto refresh")
        self.auto_preview.setChecked(True)
        self.auto_preview.toggled.connect(self.auto_refresh_changed)
        bar.addWidget(self.auto_preview)
        bar.addWidget(self.button("Fit all", self.fit_map))
        bar.addWidget(self.button("Refresh", self.plot_map))
        layout.addLayout(bar)
        limitations = QLabel('Preview/HTML: color, size, opacity, popup fields. Google Earth: icon shapes and labels. Filters do not change preview or export scope.')
        limitations.setObjectName('Muted')
        limitations.setWordWrap(True)
        layout.addWidget(limitations)
        self.bridge = MapBridge(self)
        self.web_view = LazyWebView(self)
        self.web_view.loadFinished.connect(self.browser_load_finished)
        self.map_placeholder = QLabel("Your points, in view\n\nOpen a file and apply its column mapping to begin.")
        self.map_placeholder.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.map_placeholder.setWordWrap(True)
        self.map_placeholder.setStyleSheet("background:#edf3f6; color:#61738b; border-radius:6px; padding:24px; font-size:11pt;")
        layout.addWidget(self.map_placeholder, 1)
        self.web_view.hide()
        layout.addWidget(self.web_view, 1)
        self.overview = CoordinateOverview()
        self.overview.pointSelected.connect(lambda row: self.select_map_row(row) if self.map_ready else None)
        self.overview.hide()
        layout.addWidget(self.overview, 1)
        self.map_status = QLabel("Open a file to preview its points")
        self.map_status.setObjectName("Muted")
        self.map_status.setWordWrap(True)
        status_bar = QHBoxLayout()
        status_bar.addWidget(self.map_status, 1)
        status_bar.addWidget(self.preview_badge)
        layout.addLayout(status_bar)
        return panel

    def build_data_panel(self):
        panel = self.card()
        panel.setMinimumHeight(160)
        layout = QVBoxLayout(panel)
        bar = QHBoxLayout()
        bar.addWidget(self.heading("Data"))
        self.search = QLineEdit()
        self.search.setPlaceholderText("Search rows…")
        self.search.setClearButtonEnabled(True)
        self.search.textChanged.connect(self.filter_rows)
        bar.addWidget(self.search, 1)
        self.errors_only = QCheckBox("Errors only")
        self.errors_only.toggled.connect(self.filter_rows)
        bar.addWidget(self.errors_only)
        bar.addWidget(self.button("Next issue", self.next_issue))
        self.btn_undo = self.button("Undo", self.undo_edit)
        self.btn_undo.setShortcut("Ctrl+Z")
        self.btn_redo = self.button("Redo", self.redo_edit)
        self.btn_redo.setShortcut("Ctrl+Y")
        bar.addWidget(self.btn_undo)
        bar.addWidget(self.btn_redo)
        layout.addLayout(bar)
        filters = QHBoxLayout()
        filters.addWidget(QLabel("Category"))
        self.category_filter = QComboBox()
        self.category_filter.addItem("All categories", None)
        self.category_filter.setMaximumWidth(210)
        self.category_filter.currentIndexChanged.connect(self.filter_rows)
        filters.addWidget(self.category_filter)
        self.duplicates_only = QCheckBox("Duplicate warnings only")
        self.duplicates_only.toggled.connect(self.filter_rows)
        filters.addWidget(self.duplicates_only)
        filters.addStretch()
        self.btn_remove = QToolButton()
        self.btn_remove.setText("Remove…")
        self.btn_remove.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        menu = QMenu(self.btn_remove)
        self.remove_rows_action = menu.addAction("Selected rows", self.remove_selected_rows)
        self.remove_column_action = menu.addAction("Column…", self.choose_column_to_remove)
        self.btn_remove.setMenu(menu)
        self.btn_remove.setToolTip("Remove data from this working copy. Undo restores it; the source file stays unchanged.")
        filters.addWidget(self.btn_remove)
        layout.addLayout(filters)
        self.table = DataTable(self)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.table.setSortingEnabled(False)
        self.table.verticalHeader().setDefaultSectionSize(29)
        self.table.itemChanged.connect(self.on_cell_changed)
        self.table.itemSelectionChanged.connect(self.highlight_selected_row)
        self.table.itemSelectionChanged.connect(self.update_remove_actions)
        layout.addWidget(self.table, 1)
        self.filter_status = QLabel("No data loaded")
        self.filter_status.setObjectName("Muted")
        layout.addWidget(self.filter_status)
        return panel

    def toggle_appearance(self):
        hidden = not self.appearance_panel.isHidden()
        self.appearance_panel.setHidden(hidden)
        self.btn_appearance.setText("Show appearance" if hidden else "Hide appearance")

    def toggle_map(self):
        if self.compact_mode:
            show_map = self.map_panel.isHidden()
            self.map_panel.setVisible(show_map)
            self.data_panel.setVisible(not show_map)
            self.btn_map.setText('Show data' if show_map else 'Show map')
            if show_map and self.online_ready:
                self.web_view.page().runJavaScript('if(window.spatialMap) spatialMap.invalidateSize();')
            return
        hidden = not self.map_panel.isHidden()
        self.map_panel.setHidden(hidden)
        self.btn_map.setText("Show map" if hidden else "Hide map")
        if not hidden and self.online_ready:
            self.web_view.page().runJavaScript("if(window.spatialMap) spatialMap.invalidateSize();")

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if not hasattr(self, 'map_panel'):
            return
        compact = self.height() < 800
        if compact == self.compact_mode:
            return
        self.compact_mode = compact
        if compact:
            self.restore_map_after_compact = not self.map_panel.isHidden()
            self.map_panel.hide()
            self.data_panel.show()
            self.btn_map.setText('Show map')
        else:
            self.data_panel.show()
            self.map_panel.setVisible(self.restore_map_after_compact)
            self.btn_map.setText('Hide map' if self.restore_map_after_compact else 'Show map')

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls() and any(url.isLocalFile() for url in event.mimeData().urls()):
            event.acceptProposedAction()

    def dropEvent(self, event):
        paths = [url.toLocalFile() for url in event.mimeData().urls() if url.isLocalFile()]
        if paths:
            self.load_file(paths[0])

    def load_file(self, filepath=None, sheet_name=0, import_options=None, asynchronous=True):
        if self.import_busy or self.export_busy:
            return False
        if not isinstance(filepath, str):
            filepath, _ = QFileDialog.getOpenFileName(self, 'Open data', '', 'Supported files (*.xlsx *.xls *.csv *.txt *.geojson *.json)')
        if not filepath or not self.confirm_discard_edits():
            return False
        options = import_options or dict(header_row=1,delimiter='auto',encoding='auto')
        if not asynchronous:
            candidate = KMLProcessor()
            success, message = candidate.load_and_parse_file(filepath,allow_unmapped=True,sheet_name=sheet_name,**options)
            if not success:
                self.import_failed(filepath,message)
                return False
            return self.install_loaded(candidate,prepare_view(candidate),message)
        self.import_busy = True
        self.import_cancel.clear()
        self.preview_timer.stop()
        self.map_timeout.stop()
        self.set_import_controls(False)
        self.progress.setRange(0,0)
        self.btn_cancel.show()
        self.btn_cancel.setEnabled(True)
        self.status_lbl.setText('Reading and validating fileÃ¢â‚¬Â¦')
        def job():
            try:
                candidate = KMLProcessor()
                success, message = candidate.load_and_parse_file(filepath,allow_unmapped=True,sheet_name=sheet_name,**options)
                check_cancel(self.import_cancel)
                if not success: raise ValueError(message)
                prepared = prepare_view(candidate,self.import_cancel)
                check_cancel(self.import_cancel)
                self.import_queue.put(('done',(candidate,prepared,message)))
            except OperationCancelled:
                self.import_queue.put(('cancelled',None))
            except Exception as error:
                LOG.exception('Background import failed')
                self.import_queue.put(('error',(filepath,str(error))))
        self.import_worker = threading.Thread(target=job,daemon=True)
        self.import_timer.start(60)
        self.import_worker.start()
        return True

    def set_import_controls(self, enabled):
        # Leave Cancel, Help and the window itself responsive.
        self.setup_panel.setEnabled(enabled)
        self.appearance_panel.setEnabled(enabled)
        self.data_panel.setEnabled(enabled)
        self.map_panel.setEnabled(enabled)
        self.btn_export.setEnabled(enabled and self.mapping_ready)
        self.btn_save_copy.setEnabled(enabled and self.processor.df is not None)
        self.format_combo.setEnabled(enabled)
        self.valid_only.setEnabled(enabled)

    def check_import_queue(self):
        try:
            kind,value = self.import_queue.get_nowait()
        except queue.Empty:
            return
        self.import_timer.stop()
        self.import_busy = False
        self.progress.setRange(0,100)
        self.progress.setValue(0)
        self.btn_cancel.hide()
        self.set_import_controls(True)
        if self.import_cancel.is_set(): kind = 'cancelled'
        if kind == 'done': self.install_loaded(*value)
        elif kind == 'error':
            self.close_after_job = False
            self.import_failed(*value)
        else: self.status_lbl.setText('Import cancelled Ã‚· previous data retained')
        self.sheet_combo.blockSignals(True)
        if self.processor.current_sheet: self.sheet_combo.setCurrentText(self.processor.current_sheet)
        self.sheet_combo.blockSignals(False)
        self.update_data_status()
        if self.preview_dirty and self.auto_preview.isChecked(): self.schedule_preview()
        if self.close_after_job:
            self.close_after_job = False
            self.close()

    def import_failed(self, filepath, message):
        self.pending_import_path = filepath
        self.btn_import_options.setEnabled(os.path.splitext(filepath)[1].lower() in ('.xlsx','.xls','.csv','.txt'))
        self.btn_import_options.setToolTip('Configure the failed import: '+filepath)
        self.status_lbl.setText('Import failed Ã‚· previous data retained')
        QMessageBox.critical(self,'Could not open file',message+'\n\nTry Import options to choose a header row, delimiter or encoding. Diagnostic details: Help Ã¢â€ â€™ Open diagnostic folder.')

    def cancel_job(self):
        if self.import_busy:
            self.import_cancel.set()
            self.status_lbl.setText('CancellingÃ¢â‚¬Â¦ waiting for the file reader to return')
        elif self.export_busy:
            self.cancel_event.set()
            self.status_lbl.setText('CancellingÃ¢â‚¬Â¦ a final commit already started may finish')
        self.btn_cancel.setEnabled(False)

    def open_output_folder(self):
        if self.output_path:
            QDesktopServices.openUrl(QUrl.fromLocalFile(os.path.dirname(os.path.abspath(self.output_path))))

    def open_logs(self):
        if LOG_PATH:
            QDesktopServices.openUrl(QUrl.fromLocalFile(os.path.dirname(LOG_PATH)))
        else:
            QMessageBox.information(self,'Diagnostic log','Logging is unavailable in this session. Check that your local application-data folder is writable.')

    def save_team_preset(self):
        if not self.mapping_ready or self.mapping_pending or self.import_busy:
            self.status_lbl.setText('Apply the column mapping before saving a preset')
            return
        path,_ = QFileDialog.getSaveFileName(self,'Save portable preset','SitePin.sppreset','SitePin preset (*.sppreset)')
        if not path: return
        if not path.lower().endswith('.sppreset'):
            path += '.sppreset'
            if os.path.exists(path) and QMessageBox.question(self,'Replace preset?',path,QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No,QMessageBox.StandardButton.No)!=QMessageBox.StandardButton.Yes: return
        try:
            save_preset(path,self.processor,self.current_styles(),self.visible_columns())
            self.status_lbl.setText('Portable preset saved Ã‚· mappings, popup fields, styles and icons')
        except Exception as error:
            LOG.exception('Preset save failed')
            QMessageBox.critical(self,'Preset not saved',str(error))

    def load_team_preset(self):
        if self.processor.df is None or self.import_busy:
            self.status_lbl.setText('Open a dataset before applying a preset')
            return
        path,_=QFileDialog.getOpenFileName(self,'Load portable preset','','SitePin preset (*.sppreset)')
        if not path: return
        try:
            mapping,fields,styles = read_preset(path,self.preset_icons.name)
            missing = sorted(set(mapping.values()).difference(self.processor.df.columns))
            if missing: raise ValueError('Preset needs these columns: '+', '.join(missing)+'. Load a matching file or map its columns manually.')
            candidate = self.processor.snapshot()
            candidate.apply_mapping(mapping)
            # Commit only after all validation succeeds.
            self.processor = candidate
            self.mapping_ready, self.mapping_pending = True,False
            self.styles.update(styles)
            self.populate_mapping()
            self.populate_table()
            self.populate_fields()
            self.list_columns.blockSignals(True)
            for index in range(self.list_columns.count()):
                item=self.list_columns.item(index)
                item.setCheckState(Qt.CheckState.Checked if item.text() in fields else Qt.CheckState.Unchecked)
            self.list_columns.blockSignals(False)
            self.reconcile_categories()
            self.update_data_status()
            self.schedule_preview()
            missing_fields = set(fields).difference(self.display_columns)
            self.status_lbl.setText('Preset applied'+(' Ã‚· unavailable popup fields skipped: '+', '.join(sorted(missing_fields)) if missing_fields else ''))
        except Exception as error:
            LOG.exception('Preset load failed')
            QMessageBox.critical(self,'Preset not applied',str(error))

    def install_loaded(self, candidate, prepared, message):
        self.processor = candidate
        filepath = candidate.current_filepath
        self.pending_import_path = None
        self.btn_import_options.setToolTip("Preview header row, delimiter and text encoding")
        self.loading_data = True
        self.preview_timer.stop()
        self.clear_map()
        self.styles.clear()
        self.undo_stack.clear()
        self.redo_stack.clear()
        self.dirty = False
        self.mapping_ready = all(key in self.processor.target_cols for key in ("point name", "latitude", "longitude"))
        self.valid_only.setChecked(False)
        self.mapping_ready = self.mapping_ready and len(set(self.processor.target_cols.values())) == len(self.processor.target_cols)
        self.mapping_pending = not self.mapping_ready
        self.file_label.setText(os.path.basename(filepath))
        self.file_label.setToolTip(os.path.abspath(filepath))
        self.sheet_combo.blockSignals(True)
        self.sheet_combo.clear()
        self.sheet_combo.addItems(self.processor.sheet_names or ["Not an Excel file"])
        if self.processor.current_sheet:
            self.sheet_combo.setCurrentText(self.processor.current_sheet)
        self.sheet_combo.setEnabled(bool(self.processor.sheet_names))
        self.sheet_combo.blockSignals(False)
        self.populate_mapping()
        self.search.clear()
        self.errors_only.setChecked(False)
        self.duplicates_only.setChecked(False)
        self.category_filter.blockSignals(True)
        self.category_filter.setCurrentIndex(0)
        self.category_filter.blockSignals(False)
        self.duplicate_notes = prepared["duplicates"]
        self.issue_cache, self.search_cache = prepared["issues"], prepared["search"]
        self.populate_table()
        self.populate_fields()
        self.loading_data = False
        self.reconcile_categories()
        self.filter_rows()
        details = self.processor.import_details
        import_info = f"Header row {details.get('header_row', 1)}"
        if 'encoding' in details:
            import_info += f" · {details['encoding']} · delimiter {repr(details['delimiter'])}"
        if details.get('encoding_guessed'):
            import_info += " · Check text encoding in Import options."
        self.mapping_status.setText(message + ' ' + import_info)
        self.status_lbl.setText("File loaded · edits stay in memory")
        self.update_data_status()
        self.schedule_preview()
        return True

    def change_sheet(self, name):
        if not name or name == self.processor.current_sheet or not self.processor.sheet_names:
            return
        previous = self.processor.current_sheet
        if not self.load_file(self.processor.current_filepath, name, self.processor.import_options):
            self.sheet_combo.blockSignals(True)
            self.sheet_combo.setCurrentText(previous)
            self.sheet_combo.blockSignals(False)

    def show_import_options(self):
        path = self.pending_import_path or self.processor.current_filepath
        if not path:
            return
        same_file = path == self.processor.current_filepath
        dialog = ImportOptionsDialog(self, path,
            self.processor.import_options if same_file else None,
            self.processor.current_sheet if same_file and self.processor.current_sheet else 0)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.load_file(path, dialog.sheet.currentText() or 0, dialog.settings())

    def save_sample_file(self):
        path, _ = QFileDialog.getSaveFileName(self, "Save sample CSV", "SitePin_sample.csv", "CSV (*.csv)")
        if not path:
            return
        if not path.lower().endswith('.csv'):
            path += '.csv'
            if os.path.exists(path) and QMessageBox.question(self, "Replace sample?", f"Replace {path}?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No) != QMessageBox.StandardButton.Yes:
                return
        try:
            if self.processor.current_filepath and os.path.normcase(os.path.realpath(path)) == os.path.normcase(os.path.realpath(self.processor.current_filepath)):
                raise ValueError("Choose another filename. The original input file is protected.")
            write_sample_csv(path)
            self.status_lbl.setText("Sample CSV saved · use its column names as a starting point")
        except Exception as error:
            QMessageBox.critical(self, "Sample not saved", str(error))

    def populate_mapping(self):
        columns = [c for c in self.processor.df.columns if not c.startswith("__parsed")]
        for key, combo in self.mapping_combos.items():
            combo.blockSignals(True)
            combo.clear()
            combo.addItem("All points" if key == "category" else "Choose column…", None)
            for column in columns:
                combo.addItem(column, column)
            index = combo.findData(self.processor.target_cols.get(key))
            combo.setCurrentIndex(max(0, index))
            combo.blockSignals(False)

    def mapping_changed(self):
        if self.processor.df is None:
            return
        self.mapping_pending = True
        self.preview_timer.stop()
        self.clear_map()
        self.mapping_status.setText("Mapping changed. Apply it to validate and preview these columns.")
        self.update_data_status()

    def apply_mapping(self):
        if self.processor.df is None:
            return
        mapping = {key: combo.currentData() for key, combo in self.mapping_combos.items()}
        try:
            self.processor.apply_mapping(mapping)
        except ValueError as error:
            self.mapping_status.setText(str(error))
            self.update_data_status()
            return
        self.mapping_ready = True
        self.mapping_pending = False
        self.populate_mapping()
        self.duplicate_notes = self.processor.duplicate_warnings()
        self.populate_table()
        self.populate_fields(preserve=True)
        self.reconcile_categories()
        self.mapping_status.setText("Mapping applied. Coordinates are interpreted as WGS84 degrees.")
        self.update_data_status()
        self.schedule_preview()

    def populate_table(self):
        df = self.processor.df
        self.display_columns = [c for c in df.columns if not c.startswith('__parsed')]
        if not self.loading_data:
            prepared = prepare_view(self.processor)
            self.issue_cache, self.search_cache = prepared['issues'], prepared['search']
            self.duplicate_notes = prepared['duplicates']
        self.table.frame_model.beginResetModel()
        self.table.columns = self.display_columns
        self.table.rows = list(range(len(df)))
        self.table.positions = {row: row for row in self.table.rows}
        self.table.frame_model.endResetModel()
        self.table.verticalHeader().setToolTip('Original worksheet/CSV row, or GeoJSON feature number')
        for column in range(len(self.display_columns)):
            self.table.setColumnWidth(column, 130 if column else 155)
        self.table.setColumnWidth(len(self.display_columns),310)
        self.filter_rows()

    def populate_fields(self, preserve=False):
        checked = set(self.visible_columns()) if preserve else set(self.display_columns)
        self.list_columns.blockSignals(True)
        self.list_columns.clear()
        for column in self.display_columns:
            item = QListWidgetItem(column)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Checked if column in checked else Qt.CheckState.Unchecked)
            self.list_columns.addItem(item)
        self.list_columns.blockSignals(False)

    def visible_columns(self):
        return [self.list_columns.item(i).text() for i in range(self.list_columns.count())
                if self.list_columns.item(i).checkState() == Qt.CheckState.Checked]

    def row_issues(self, row):
        if not self.mapping_ready:
            return ["Apply the column mapping to validate this row."]
        return [issue for axis in ("latitude", "longitude")
                if (issue := coordinate_issue(self.processor.df.at[row, self.processor.target_cols[axis]], axis))]

    def refresh_row_validation(self, row):
        if self.mapping_ready:
            for axis, internal in (("latitude", "__parsed_lat"), ("longitude", "__parsed_lon")):
                self.processor.df.at[row, internal] = parse_coordinate(
                    self.processor.df.at[row, self.processor.target_cols[axis]], axis)
        self.issue_cache[row] = self.row_issues(row)
        self.search_cache[row] = ' '.join(display_text(self.processor.df.at[row,c]) for c in self.display_columns).casefold()
        self.table.refresh(row)

    def on_cell_changed(self, item):
        if self.processor.df is None or item.column() >= len(self.display_columns):
            return
        row, col = item.row(), item.column()
        name = self.display_columns[col]
        old, new = self.processor.df.at[row, name], item.text()
        old_text = json.dumps(old, ensure_ascii=False) if isinstance(old, (dict, list)) else "" if pd.isna(old) else str(old)
        if old_text == new:
            return
        self.undo_stack.append((row, col, name, old, new))
        self.redo_stack.clear()
        self.apply_cell_edit(row, name, new)

    def apply_cell_edit(self, row, name, value):
        self.processor.df[name] = self.processor.df[name].astype(object)
        self.processor.df.at[row, name] = value
        self.refresh_row_validation(row)
        previous = self.duplicate_notes
        self.duplicate_notes = self.processor.duplicate_warnings()
        affected = {index for index in self.duplicate_notes if self.duplicate_notes[index] != previous.get(index)} | {row}
        for index in affected:
            self.refresh_row_validation(index)
        if name == self.processor.target_cols.get("category"):
            self.reconcile_categories()
        self.dirty = True
        self.filter_rows()
        self.update_data_status()
        self.schedule_preview()

    def update_remove_actions(self):
        loaded = self.processor.df is not None and not self.import_busy
        self.btn_remove.setEnabled(loaded)
        self.remove_rows_action.setEnabled(loaded and bool(self.table.selectionModel().selectedRows()))
        self.remove_column_action.setEnabled(loaded and bool(self.table.columns))

    def structure_state(self):
        return dict(processor=self.processor.snapshot(), ready=self.mapping_ready,
                    pending=self.mapping_pending, styles=copy.deepcopy(self.styles),
                    fields=self.visible_columns(),
                    choices={key: combo.currentData() for key, combo in self.mapping_combos.items()},
                    message=self.mapping_status.text())

    def restore_structure(self, state):
        # Invalidate old map callbacks before rebasing dataframe row positions.
        self.clear_map()
        self.table.clearSelection()
        self.processor = state['processor'].snapshot()
        self.mapping_ready, self.mapping_pending = state['ready'], state['pending']
        self.styles = copy.deepcopy(state['styles'])
        self.populate_mapping()
        for key, value in state['choices'].items():
            combo = self.mapping_combos[key]
            combo.blockSignals(True)
            combo.setCurrentIndex(max(0, combo.findData(value)))
            combo.blockSignals(False)
        self.populate_table()
        self.populate_fields()
        self.list_columns.blockSignals(True)
        for index in range(self.list_columns.count()):
            item = self.list_columns.item(index)
            item.setCheckState(Qt.CheckState.Checked if item.text() in state['fields'] else Qt.CheckState.Unchecked)
        self.list_columns.blockSignals(False)
        self.reconcile_categories()
        self.mapping_status.setText(state['message'])
        self.dirty = True
        self.update_data_status()
        self.schedule_preview()

    def commit_structure(self, before, after):
        self.undo_stack.append(dict(kind='structure', before=before, after=after))
        self.redo_stack.clear()
        self.restore_structure(after)

    def remove_selected_rows(self):
        if self.processor.df is None or self.import_busy:
            return
        # Qt selection indices refer to the filtered view, not source positions.
        rows = {self.table.rows[index.row()] for index in self.table.selectionModel().selectedRows()}
        if not rows:
            return
        before = self.structure_state()
        after = self.structure_state()
        candidate = after['processor']
        keep = [row for row in range(len(candidate.df)) if row not in rows]
        candidate.df = candidate.df.iloc[keep].reset_index(drop=True)
        candidate.source_rows = [candidate.source_rows[row] for row in keep]
        self.commit_structure(before, after)

    def choose_column_to_remove(self):
        if self.processor.df is None or self.import_busy:
            return
        column, accepted = QInputDialog.getItem(
            self, "Remove column", "Column to remove from the working copy:\n"
            "Removing a mapped column requires applying the mapping again.\n"
            "You can restore the column with Undo.", self.display_columns, 0, False)
        if accepted:
            self.remove_column(column)

    def remove_column(self, column):
        if self.processor.df is None or self.import_busy or column not in self.display_columns:
            return
        before = self.structure_state()
        after = self.structure_state()
        candidate = after['processor']
        candidate.df = candidate.df.drop(columns=[column])
        after['fields'] = [name for name in after['fields'] if name != column]
        after['choices'] = {key: None if value == column else value for key, value in after['choices'].items()}
        if column in candidate.target_cols.values():
            # No partially valid mapping may reach export or preview code.
            candidate.target_cols = {}
            candidate.unique_categories = []
            candidate.df['__parsed_lat'] = float('nan')
            candidate.df['__parsed_lon'] = float('nan')
            after['ready'], after['pending'] = False, True
            after['message'] = f'Column "{column}" removed. Choose replacement columns and apply mapping, or Undo.'
        self.commit_structure(before, after)

    def replay_edit(self, source, destination, undo):
        if not source:
            return
        entry = source.pop()
        destination.append(entry)
        if isinstance(entry, dict) and entry.get('kind') == 'structure':
            self.restore_structure(entry['before'] if undo else entry['after'])
            return
        row, col, name, old, new = entry
        value = old if undo else new
        text = json.dumps(value, ensure_ascii=False) if isinstance(value, (dict, list)) else "" if pd.isna(value) else str(value)
        self.table.blockSignals(True)
        self.table.item(row, col).setText(text)
        self.table.blockSignals(False)
        self.apply_cell_edit(row, name, value)

    def undo_edit(self):
        self.replay_edit(self.undo_stack, self.redo_stack, True)

    def redo_edit(self):
        self.replay_edit(self.redo_stack, self.undo_stack, False)

    def filter_rows(self):
        if self.loading_data:
            return
        term = self.search.text().casefold()
        category = self.category_filter.currentData()
        category_values = self.processor.df[self.processor.target_cols['category']].astype(str).tolist() if category is not None and self.mapping_ready else None
        rows = [row for row in range(self.table.rowCount())
            if (not term or term in self.search_cache.get(row,'') or term in self.table.message(row).casefold())
            and (category is None or (category_values is not None and category_values[row] == category))
            and (not self.errors_only.isChecked() or self.issue_cache.get(row))
            and (not self.duplicates_only.isChecked() or self.duplicate_notes.get(row))]
        self.table.reset_rows(rows)
        self.filter_status.setText(f'{len(rows):,} of {self.table.rowCount():,} rows shown · filters affect the table only')

    def next_issue(self):
        count = self.table.rowCount()
        if not count:
            return
        start = self.table.currentRow()
        for offset in range(1, count + 1):
            row = (start + offset) % count
            if self.row_issues(row):
                self.search.clear()
                self.duplicates_only.setChecked(False)
                self.category_filter.setCurrentIndex(0)
                self.table.selectRow(row)
                self.table.scrollToItem(self.table.item(row, 0))
                return
        self.status_lbl.setText("No coordinate issues found")

    def update_data_status(self):
        loaded = self.processor.df is not None
        active = loaded and self.mapping_ready and not self.mapping_pending
        invalid = int(self.processor.df[["__parsed_lat", "__parsed_lon"]].isna().any(axis=1).sum()) if loaded else 0
        self.update_remove_actions()
        self.btn_undo.setEnabled(bool(self.undo_stack))
        self.btn_redo.setEnabled(bool(self.redo_stack))
        self.btn_save_copy.setEnabled(loaded)
        option_path = self.pending_import_path or self.processor.current_filepath or ''
        self.btn_import_options.setEnabled(os.path.splitext(option_path)[1].lower() in ('.xlsx', '.xls', '.csv', '.txt'))
        self.btn_apply_mapping.setEnabled(loaded)
        for combo in self.mapping_combos.values():
            combo.setEnabled(loaded)
        self.sheet_combo.setEnabled(bool(self.processor.sheet_names))
        reason = "Open a file to begin"
        if loaded:
            total = len(self.processor.df)
            state = "Unsaved edits" if self.dirty else "Original protected"
            reason = "Apply the column mapping" if not active else f"Will export {total-invalid:,} valid rows; exclude {invalid:,} with a report" if self.valid_only.isChecked() and invalid else f"Fix {invalid} invalid row(s) before exporting" if invalid else f"Ready to export all {total:,} points"
            duplicates = sum(bool(notes) for notes in self.duplicate_notes.values())
            self.data_status.setText(f"{total:,} rows · {invalid:,} invalid · {duplicates:,} duplicate warnings · {state}   |   {reason}")
        else:
            self.data_status.setText("Open a file to begin. Your original files stay unchanged.")
        can_export = active and len(self.processor.df) > invalid and (not invalid or self.valid_only.isChecked())
        self.btn_export.setEnabled(bool(can_export and not self.export_busy))
        self.valid_only.setEnabled(bool(active and not self.export_busy))
        self.btn_export.setToolTip(reason)
        self.format_combo.setEnabled(not self.export_busy)
        if self.import_busy:
            self.set_import_controls(False)
        self.setWindowTitle(f"{TOOL_NAME} v{VERSION}" + (" *" if self.dirty else ""))

    @staticmethod
    def default_style(index=0):
        palette = ["#087f8c", "#e09931", "#7664ae", "#c35f78", "#46936b", "#487bb2"]
        return dict(shape="Pushpin", custom_path=None, icolor=palette[index % len(palette)],
                    opacity=1.0, iscale=1.2, lcolor="#ffffff", lscale=1.0)

    def reconcile_categories(self):
        selected = self.category_list.currentItem()
        previous = selected.data(Qt.ItemDataRole.UserRole) if selected else None
        self.category_list.blockSignals(True)
        self.category_list.clear()
        if self.mapping_ready:
            column = self.processor.target_cols["category"]
            counts = self.processor.df[column].astype(str).value_counts()
            self.processor.unique_categories = sorted(counts.index.tolist())
            for index, category in enumerate(self.processor.unique_categories):
                self.styles.setdefault(category, self.default_style(index))
                item = QListWidgetItem(f"{category or '(blank)'}   ·   {counts[category]}")
                item.setData(Qt.ItemDataRole.UserRole, category)
                item.setToolTip(category)
                pixmap = QPixmap(12, 12)
                pixmap.fill(QColor(self.styles[category]["icolor"]))
                item.setIcon(QIcon(pixmap))
                self.category_list.addItem(item)
                if category == previous:
                    self.category_list.setCurrentItem(item)
        self.category_list.blockSignals(False)
        if self.category_list.currentRow() < 0 and self.category_list.count():
            self.category_list.setCurrentRow(0)
        else:
            self.show_category_style()
        previous_filter = self.category_filter.currentData()
        self.category_filter.blockSignals(True)
        self.category_filter.clear()
        self.category_filter.addItem("All categories", None)
        if self.mapping_ready:
            for category in self.processor.unique_categories:
                self.category_filter.addItem(category or '(blank)', category)
        self.category_filter.setCurrentIndex(max(0, self.category_filter.findData(previous_filter)))
        self.category_filter.blockSignals(False)
        self.filter_rows()

    def current_category(self):
        item = self.category_list.currentItem()
        return item.data(Qt.ItemDataRole.UserRole) if item else None

    def show_category_style(self):
        category = self.current_category()
        enabled = category is not None
        self.btn_color.setEnabled(enabled)
        self.advanced_body.setEnabled(enabled)
        if not enabled:
            return
        config = self.styles[category]
        self.shape_combo.blockSignals(True)
        self.shape_combo.setCurrentText(config["shape"])
        self.shape_combo.blockSignals(False)
        self.btn_custom.setVisible(config["shape"] == "Custom Image")
        self.btn_custom.setToolTip(config.get("custom_path") or "Choose an image to bundle in the KMZ")
        for key, spin in self.style_spins.items():
            spin.blockSignals(True)
            spin.setValue(config[key])
            spin.blockSignals(False)
        self.color_button(self.btn_color, "Point color", config["icolor"])
        self.color_button(self.btn_label_color, "Label color", config["lcolor"])

    @staticmethod
    def color_button(button, label, value):
        button.setText(f"{label}  {value}")
        button.setStyleSheet(f"border-left:14px solid {value};text-align:left;")

    def change_style(self, key, value):
        category = self.current_category()
        if category is None:
            return
        self.styles[category][key] = value
        self.show_category_style()
        if key == "icolor":
            self.reconcile_categories()
        self.schedule_preview()

    def pick_style_color(self, key):
        category = self.current_category()
        if category is not None:
            color = QColorDialog.getColor(QColor(self.styles[category][key]), self, "Choose color")
            if color.isValid():
                self.change_style(key, color.name())

    def choose_custom_icon(self):
        path, _ = QFileDialog.getOpenFileName(self, "Choose icon", "", "Images (*.png *.jpg *.jpeg *.gif)")
        if path:
            self.change_style("custom_path", path)

    def apply_palette(self):
        count = self.category_list.count()
        for index in range(count):
            category = self.category_list.item(index).data(Qt.ItemDataRole.UserRole)
            mode = self.cb_palette.currentIndex()
            hue = index / max(1, count) if mode == 0 else .59 if mode == 1 else index / max(1, count-1) * .33
            value = .88 if mode != 1 else .45 + .45 * index / max(1, count-1)
            r, g, b = colorsys.hsv_to_rgb(hue, .72, value)
            self.styles[category]["icolor"] = f"#{int(r*255):02x}{int(g*255):02x}{int(b*255):02x}"
        self.reconcile_categories()
        self.schedule_preview()

    def current_styles(self):
        return {category: copy.deepcopy(self.styles[category]) for category in self.processor.unique_categories if category in self.styles}

    def save_styles(self):
        if not self.current_styles():
            return
        path, _ = QFileDialog.getSaveFileName(self, "Save styles", "spatialpin_styles.json", "JSON (*.json)")
        if path:
            try:
                if self.processor.current_filepath and os.path.normcase(os.path.realpath(path)) == os.path.normcase(os.path.realpath(self.processor.current_filepath)):
                    raise ValueError("Choose a different filename. The original input file is protected.")
                with open(path, "w", encoding="utf-8") as handle:
                    json.dump(self.current_styles(), handle, ensure_ascii=False, indent=2)
                self.status_lbl.setText("Styles saved")
            except Exception as error:
                QMessageBox.critical(self, "Styles not saved", str(error))

    def load_styles(self):
        if not self.mapping_ready:
            return
        path, _ = QFileDialog.getOpenFileName(self, "Load styles", "", "JSON (*.json)")
        if not path:
            return
        try:
            with open(path, encoding="utf-8") as handle:
                data = json.load(handle)
            if not isinstance(data, dict):
                raise ValueError("The style file must contain category objects.")
            updated = copy.deepcopy(self.styles)
            for category, config in data.items():
                if category not in updated:
                    continue
                if not isinstance(config, dict):
                    raise ValueError(f"Invalid style for {category}.")
                for key, value in config.items():
                    if key in ("icolor", "lcolor") and (not isinstance(value, str) or not re.fullmatch(r"#[0-9a-fA-F]{6}", value)):
                        raise ValueError("Colors must use #RRGGBB format.")
                    if key == "shape" and value not in list(GE_ICONS) + ["Custom Image"]:
                        raise ValueError("Unknown icon shape.")
                    if key in ("opacity", "iscale", "lscale"):
                        low, high = {"opacity": (.1, 1), "iscale": (.5, 3), "lscale": (0, 3)}[key]
                        if not isinstance(value, (int, float)) or not math.isfinite(value) or not low <= value <= high:
                            raise ValueError(f"Invalid {key}.")
                    if key == "custom_path" and value is not None and not isinstance(value, str):
                        raise ValueError("Invalid image path.")
                    if key in updated[category]:
                        updated[category][key] = value
            self.styles = updated
            self.reconcile_categories()
            self.schedule_preview()
            self.status_lbl.setText("Styles loaded for matching categories")
        except Exception as error:
            QMessageBox.critical(self, "Styles not loaded", str(error))

    def auto_refresh_changed(self, checked):
        if checked and self.preview_dirty:
            self.preview_timer.start()
        elif not checked:
            self.preview_timer.stop()
        if self.preview_dirty:
            self.map_status.setText('Preview out of date · ' + ('refreshing…' if checked else 'click Refresh'))

    def schedule_preview(self, *args):
        if not self.mapping_ready or self.mapping_pending:
            return
        self.preview_dirty = True
        self.map_ready = False
        self.online_ready = False
        self.map_generation += 1
        self.map_timeout.stop()
        self.overview.interactive = False
        self.preview_badge.setText('Out of date')
        self.preview_badge.setStyleSheet('color:#9c5f08;font-weight:600;')
        self.map_status.setText('Preview out of date · ' + ('refreshing…' if self.auto_preview.isChecked() else 'click Refresh'))
        if self.auto_preview.isChecked():
            self.preview_timer.start()
        else:
            self.preview_timer.stop()

    def clear_map(self):
        self.preview_timer.stop()
        self.map_timeout.stop()
        self.map_generation += 1
        self.map_ready = self.online_ready = self.preview_dirty = False
        self.tile_state = 'pending'
        self.browser_failure = ''
        self.browser_path = None
        self.preview_rows.clear()
        self.web_view.stop()
        self.web_view.hide()
        self.overview.hide()
        self.overview.set_points([], [[-1,-1],[1,1]])
        self.map_placeholder.setText('Your points, in view\n\nOpen a file and apply its column mapping to begin.\n\nAn interactive coordinate overview is available without internet access.')
        self.map_placeholder.show()
        self.preview_badge.setText('No data')
        self.preview_badge.setStyleSheet('color:#61738b;')
        self.map_status.setText('Online basemap or coordinate overview · neither changes your source data')

    def plot_map(self):
        self.preview_timer.stop()
        self.map_timeout.stop()
        if not self.mapping_ready or self.mapping_pending:
            return
        frame = self.processor.df
        valid = frame.dropna(subset=['__parsed_lat', '__parsed_lon'])
        if valid.empty:
            self.clear_map()
            self.map_placeholder.setText('No valid points yet\n\nFix the coordinate issues in the data grid to see your points.')
            self.map_status.setText('No valid coordinates to preview. Invalid rows remain in the table.')
            return
        self.map_generation += 1
        generation = self.map_generation
        self.online_ready = False
        self.tile_state = 'pending'
        self.browser_failure = ''
        self.browser_path = None
        try:
            selected = self.table.currentRow()
            indices = sample_indices(valid, self.processor.target_cols['category'], 500, selected)
            self.preview_rows = set(indices)
            offline = self.basemap_combo.currentText() == 'Coordinate overview'
            basemap = 'OpenStreetMap' if offline else self.basemap_combo.currentText()
            map_object, points, markers, bounds = self.processor.build_web_map(
                self.current_styles(), self.visible_columns(), indices, basemap)
            self.preview_points = points
            self.map_bounds = bounds
            self.preview_categories = len({point['category'] for point in points})
            self.total_categories = valid[self.processor.target_cols['category']].astype(str).nunique()
            excluded = len(frame)-len(valid)
            self.preview_summary = (f'{len(points):,} of {len(valid):,} valid points · '
                f'{self.preview_categories:,}/{self.total_categories:,} valid categories · {excluded:,} invalid not plotted')
            self.overview.set_points(points, bounds, selected)
            self.overview.interactive = True
            self.map_ready = True
            self.preview_dirty = False
            self.map_placeholder.hide()
            self.web_view.stop()
            self.web_view.hide()
            self.overview.show()
            self.update_preview_presentation()
            entries = '\n'.join(f'spatialMarkers[{row}] = {name};' for row, name in markers)
            script = """
                window.spatialMap = MAP_NAME;
                window.spatialMarkers = {};
                MARKER_ENTRIES
                window.spatialSelected = null;
                window.selectSpatialRow = function(row) {
                    if (spatialSelected) spatialSelected.setStyle({color:'white',weight:1.5});
                    spatialSelected = null;
                    var marker = spatialMarkers[row];
                    if (marker) {
                        marker.setStyle({color:'#14263e',weight:4}); marker.bringToFront();
                        spatialMap.panTo(marker.getLatLng()); marker.openPopup(); spatialSelected=marker;
                    }
                };
                window.fitSpatialPoints = function() { spatialMap.fitBounds(BOUNDS, {padding:[24,24],maxZoom:15}); };
                var spatialBridge = null, tileLoads = 0, tileErrors = 0;
                function reportTiles() {
                    if (spatialBridge) spatialBridge.tiles(GENERATION,
                        tileLoads ? (tileErrors ? 'partial' : 'available') : (tileErrors ? 'unavailable' : 'pending'));
                }
                spatialMap.eachLayer(function(layer) {
                    if (layer instanceof L.TileLayer) {
                        layer.on('tileload', function(){tileLoads++; reportTiles();});
                        layer.on('tileerror', function(){tileErrors++; reportTiles();});
                    }
                });
                window.addEventListener('error', function() {
                    if (spatialBridge) spatialBridge.failed(GENERATION, 'Online map script error');
                });
                var channelScript = document.createElement('script');
                channelScript.src = 'qrc:///qtwebchannel/qwebchannel.js';
                channelScript.onload = function() {
                    new QWebChannel(qt.webChannelTransport, function(channel) {
                        spatialBridge = channel.objects.bridge;
                        Object.keys(spatialMarkers).forEach(function(row) {
                            spatialMarkers[row].on('click', function(){spatialBridge.selectRow(Number(row), GENERATION);});
                        });
                        spatialBridge.ready(GENERATION);
                        reportTiles();
                    });
                };
                document.head.appendChild(channelScript);
            """
            script = script.replace('MAP_NAME', map_object.get_name()).replace('MARKER_ENTRIES', entries)
            script = script.replace('BOUNDS', json.dumps(bounds)).replace('GENERATION', str(generation))
            hook = MacroElement()
            hook._template = Template('{% macro script(this, kwargs) %}' + script + '{% endmacro %}')
            map_object.add_child(hook)
            document = map_object.get_root().render()
            browser_path = os.path.join(self.preview_file.name, f'preview_{generation}.html')
            for path in (browser_path, os.path.join(self.preview_file.name, 'preview.html')):
                with open(path, 'w', encoding='utf-8') as handle:
                    handle.write(document)
            self.preview_documents.append(browser_path)
            while len(self.preview_documents) > 3:
                obsolete = self.preview_documents.pop(0)
                try:
                    os.remove(obsolete)
                except OSError:
                    pass
            if not offline:
                self.browser_path = browser_path
                self.web_view.load(QUrl.fromLocalFile(browser_path))
                self.map_timeout.start()
            self.highlight_selected_row()
        except Exception as error:
            self.map_ready = False
            self.preview_dirty = True
            self.overview.interactive = False
            self.preview_badge.setText('Not current')
            self.map_status.setText(f'Preview could not update: {error}')
            if not self.overview.points:
                self.map_placeholder.setText('Preview could not be generated.\n\nYour data is still available below.')
                self.map_placeholder.show()

    def update_preview_presentation(self):
        if self.preview_dirty or not self.map_ready:
            return
        offline = self.basemap_combo.currentText() == 'Coordinate overview'
        use_online = not offline and self.online_ready and self.tile_state in ('available', 'partial') and not self.browser_failure
        self.map_placeholder.hide()
        self.web_view.setVisible(use_online)
        self.overview.setVisible(not use_online)
        self.overview.interactive = True
        self.preview_badge.setStyleSheet('color:#087f8c;font-weight:600;')
        self.preview_badge.setText('Current')
        if use_online:
            detail = 'Online basemap' + (' · some tiles failed' if self.tile_state == 'partial' else '')
            self.web_view.page().runJavaScript('if(window.spatialMap) spatialMap.invalidateSize();')
        elif offline:
            detail = 'Coordinate overview · no basemap · drag to pan, scroll to zoom'
        elif self.browser_failure or self.tile_state == 'unavailable':
            detail = 'Online map unavailable · coordinate overview remains interactive · Refresh to retry'
        else:
            detail = 'Coordinate overview ready · online basemap loading…'
        self.map_status.setText(self.preview_summary + '\n' + detail)
        self.highlight_selected_row()

    def browser_load_finished(self, success):
        if not success and self.browser_path and self.web_view.url().toLocalFile() == self.browser_path:
            self.map_failed('Browser could not load the online map')

    def map_failed(self, message='Online map timed out'):
        if self.preview_dirty:
            return
        self.map_timeout.stop()
        self.browser_failure = message if isinstance(message, str) else 'Online map unavailable'
        self.online_ready = False
        self.update_preview_presentation()
        if not self.map_ready:
            self.map_placeholder.setText('Online map unavailable.\n\nLoad valid coordinates to use the offline overview.')
            self.map_placeholder.show()

    def fit_map(self):
        if not self.map_ready:
            self.plot_map()
        else:
            self.overview.fit_all()
            if self.online_ready and not self.web_view.isHidden():
                self.web_view.page().runJavaScript('if(window.fitSpatialPoints) fitSpatialPoints();')

    def select_map_row(self, row):
        if 0 <= row < self.table.rowCount():
            if self.compact_mode and self.data_panel.isHidden():
                self.map_panel.hide()
                self.data_panel.show()
                self.btn_map.setText('Show map')
            if self.table.isRowHidden(row):
                self.search.clear()
                self.errors_only.setChecked(False)
                self.duplicates_only.setChecked(False)
                self.category_filter.setCurrentIndex(0)
            self.table.selectRow(row)
            self.table.scrollToItem(self.table.item(row, 0))

    def highlight_selected_row(self):
        row = self.table.currentRow()
        if row < 0 or not self.map_ready:
            return
        if row not in self.preview_rows and self.mapping_ready and not self.row_issues(row):
            self.plot_map()
            return
        self.overview.select_row(row)
        if self.online_ready:
            self.web_view.page().runJavaScript(f'if(window.selectSpatialRow) selectSpatialRow({int(row)});')

    def start_export(self):
        if self.import_busy or self.export_busy or not self.mapping_ready or self.mapping_pending:
            self.update_data_status()
            return
        invalid = int(self.processor.df[['__parsed_lat','__parsed_lon']].isna().any(axis=1).sum())
        if invalid and not self.valid_only.isChecked():
            self.update_data_status()
            return
        if invalid and QMessageBox.question(self, 'Export valid rows only?',
                f'Export {len(self.processor.df)-invalid:,} valid rows and exclude {invalid:,} invalid rows?\nAn excluded-rows CSV report will be saved next to the map.',
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No) != QMessageBox.StandardButton.Yes:
            return
        snapshot = self.processor.snapshot()
        snapshot.web_basemap = 'OpenStreetMap'
        styles = self.current_styles()
        valid_categories = set(snapshot.df.loc[~snapshot.df[['__parsed_lat','__parsed_lon']].isna().any(axis=1), snapshot.target_cols['category']].astype(str))
        styles = {category: style for category,style in styles.items() if category in valid_categories}
        format_name = self.format_combo.currentText()
        custom = [style for style in styles.values() if style.get("shape") == "Custom Image"]
        if format_name in ("KML", "KMZ") and any(not style.get("custom_path") or not os.path.isfile(style['custom_path']) for style in custom):
            QMessageBox.warning(self, "Missing custom icon", "Choose an existing image for each Custom Image style before exporting.")
            return
        if format_name == "KML" and custom:
            format_name = "KMZ"
            self.format_combo.setCurrentText("KMZ")
            self.status_lbl.setText("KMZ selected to include custom images")
        extension = '.' + format_name.lower()
        default = os.path.splitext(self.processor.current_filepath)[0] + "_map" + extension
        path, _ = QFileDialog.getSaveFileName(self, f"Export {format_name}", default, f"{format_name} (*{extension})")
        if not path:
            return
        if not path.lower().endswith(extension):
            path += extension
            if os.path.exists(path) and QMessageBox.question(self, "Replace output?", f"Replace {path}?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No) != QMessageBox.StandardButton.Yes:
                return
        report_path = os.path.splitext(os.path.abspath(path))[0] + '_excluded_rows.csv'
        include_report = bool(invalid) or os.path.exists(report_path)
        if os.path.exists(report_path):
            if QMessageBox.question(self, "Replace exclusion report?", f"Replace {report_path}?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No) != QMessageBox.StandardButton.Yes:
                return
        self.export_report_path = report_path if include_report else None
        self.last_excluded_count = invalid
        self.export_busy = True
        self.cancel_event.clear()
        snapshot.cancel_event = self.cancel_event
        self.btn_cancel.show()
        self.btn_cancel.setEnabled(True)
        self.progress.setValue(0)
        self.status_lbl.setText(f"Exporting {len(snapshot.df):,} points…")
        self.update_data_status()
        self.queue_timer.start(100)
        fields, valid_only = self.visible_columns(), self.valid_only.isChecked()
        def export_job():
            try:
                selected, report = snapshot.export_selection(valid_only)
                selected.export_with_report(path, styles, fields, self.export_queue, report, include_report)
            except OperationCancelled:
                self.export_queue.put(('cancelled',None))
            except Exception as error:
                LOG.exception('Export preparation failed')
                self.export_queue.put(('error',str(error)))
        self.export_worker = threading.Thread(target=export_job, daemon=True)
        self.export_worker.start()

    def check_queue(self):
        try:
            while True:
                kind, value = self.export_queue.get_nowait()
                if kind == 'progress':
                    self.progress.setValue(value)
                elif kind == 'status':
                    self.status_lbl.setText(value)
                elif kind in ('done', 'error', 'cancelled'):
                    self.queue_timer.stop()
                    self.export_busy = False
                    self.btn_cancel.hide()
                    if kind == 'done':
                        self.output_path = value
                        self.progress.setValue(100)
                        self.status_lbl.setText(f"Export complete · {self.last_excluded_count} excluded (report saved)" if self.last_excluded_count else "Export complete")
                        self.status_lbl.setToolTip(value + ('\nExclusion report: ' + self.export_report_path if self.export_report_path else ''))
                        if self.export_report_path:
                            QMessageBox.information(self, "Map and exclusion report saved",
                                f"Map:\n{value}\n\n{self.last_excluded_count} excluded rows are listed in:\n{self.export_report_path}")
                        self.btn_open_output.setEnabled(True)
                        self.btn_open_folder.setEnabled(True)
                    elif kind == 'cancelled':
                        self.progress.setValue(0)
                        self.status_lbl.setText('Export cancelled · existing files unchanged')
                    else:
                        self.close_after_job = False
                        self.status_lbl.setText("Export failed")
                        QMessageBox.critical(self, "Export failed", value + "\n\nDiagnostic details: Help → Open diagnostic folder.")
                    self.update_data_status()
                    if self.close_after_job:
                        self.close_after_job = False
                        self.close()
                    return
        except queue.Empty:
            pass

    def open_output(self):
        if self.output_path:
            QDesktopServices.openUrl(QUrl.fromLocalFile(os.path.abspath(self.output_path)))

    def create_help_dialog(self, title, subtitle, content, support=False):
        """A shared, scrollable help surface that also fits smaller displays."""
        dialog = QDialog(self)
        dialog.setWindowTitle(f"{APP_NAME} — {title}")
        dialog.setWindowIcon(self.windowIcon())
        dialog.resize(660, 560 if support else 620)
        dialog.setMinimumSize(440, 360)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(22, 22, 22, 18)
        layout.setSpacing(14)
        header = QHBoxLayout()
        icon = QLabel()
        icon.setPixmap(self.windowIcon().pixmap(48, 48))
        header.addWidget(icon)
        titles = QVBoxLayout()
        heading = QLabel(title)
        heading.setStyleSheet('font-size:17pt; font-weight:600; color:#14263e;')
        titles.addWidget(heading)
        description = QLabel(subtitle)
        description.setWordWrap(True)
        description.setObjectName('Muted')
        titles.addWidget(description)
        header.addLayout(titles, 1)
        layout.addLayout(header)
        browser = QTextBrowser()
        browser.setObjectName('HelpContent')
        browser.setOpenLinks(False)
        browser.setOpenExternalLinks(False)
        browser.anchorClicked.connect(self.open_support_link)
        browser.setStyleSheet('QTextBrowser {background:white; border:1px solid #dbe3ed; border-radius:8px; padding:12px;}')
        browser.document().setDefaultStyleSheet(
            'body {font-family:Segoe UI; color:#24344a; font-size:10pt;} '
            'h2 {color:#087f8c; font-size:12pt; margin-top:18px;} '
            'p, li {line-height:140%;} a {color:#087f8c;}')
        browser.setHtml(content)
        layout.addWidget(browser, 1)
        if support:
            actions = QHBoxLayout()
            actions.addWidget(self.button('Email support', lambda: self.open_support_link(QUrl('mailto:' + SUPPORT_EMAIL))))
            actions.addWidget(self.button('LinkedIn', lambda: self.open_support_link(QUrl(SUPPORT_LINKEDIN))))
            actions.addWidget(self.button('Copy support details', self.copy_support_details))
            layout.addLayout(actions)
        footer = QHBoxLayout()
        release = QLabel(f'{APP_NAME} {VERSION}  ·  Build {BUILD_NUMBER}')
        release.setObjectName('Muted')
        footer.addWidget(release, 1)
        close = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        close.rejected.connect(dialog.reject)
        footer.addWidget(close)
        layout.addLayout(footer)
        return dialog

    def open_support_link(self, url):
        # Only the two published contact destinations can leave this help surface.
        if url.toString() not in ('mailto:' + SUPPORT_EMAIL, SUPPORT_LINKEDIN):
            return
        if not QDesktopServices.openUrl(url):
            QMessageBox.information(self, 'Open contact link',
                'No application could open this link. Copy the contact details and open them manually.')

    def copy_support_details(self):
        QApplication.clipboard().setText(
            f'{APP_NAME} {VERSION} | Build {BUILD_NUMBER}\nAuthor: {AUTHOR_NAME}\n'
            f'Email: {SUPPORT_EMAIL}\nLinkedIn: {SUPPORT_LINKEDIN}')

    def show_help(self):
        content = f"""
        <h2>{escape_html(APP_NAME)}</h2>
        <p>{escape_html(APP_DESCRIPTION)}</p>
        <p>Turn site coordinates into maps your team can review and share.
        Import data, check locations, customize categories, and export for Google Earth or a web browser.</p>
        <h2>Release information</h2>
        <p><b>Version:</b> {VERSION}<br><b>Build:</b> {BUILD_NUMBER}<br>
        <b>Author:</b> {escape_html(AUTHOR_NAME)}</p>
        <h2>Getting started</h2>
        <p>Open <b>Help → Quick Guide</b> for the complete workflow.
        For questions, feedback, or a problem, choose <b>Help → Contact &amp; Support</b>.</p>
        <p>Your source file stays unchanged. Save a corrected copy to keep data edits.</p>
        """
        dialog = self.create_help_dialog('About SitePin', 'Product information and release details', content)
        dialog.exec()

    def show_quick_guide(self):
        content = """
        <h2>1. Prepare your data</h2>
        <p>Use a site-name column, latitude, longitude, and an optional category such as Status.
        Coordinates must be WGS84 geographic coordinates; projected coordinates are not converted.
        Use <b>Sample file…</b> for a starting template.</p>
        <h2>2. Open and review the import</h2>
        <p>Click <b>Open file…</b> and select your Excel, CSV, or supported geographic file.
        For Excel, choose the correct sheet. Use <b>Import options…</b> to preview the header row
        and, for text files, delimiter and encoding. Check that names and coordinates appear correctly.</p>
        <h2>3. Apply the column mapping</h2>
        <p>Review <b>Point name</b>, <b>Latitude</b>, <b>Longitude</b>, and <b>Category</b>.
        Choose <b>All points</b> when no category is needed. After changing a selection, click
        <b>Apply mapping</b> to validate the data.</p>
        <h2>4. Check and edit the table</h2>
        <p>Double-click a data cell to edit it. Use <b>Errors only</b> and <b>Next issue</b> to find
        missing or invalid coordinates. Review <b>Duplicate warnings only</b>; duplicate warnings
        do not automatically remove points.</p>
        <p>Select rows with Ctrl-click or Shift-click, then choose <b>Remove… → Selected rows</b>.
        Use <b>Remove… → Column…</b> to delete a column from the working copy.
        Removing a mapped column requires applying the mapping again.
        Use <b>Undo / Redo</b> (Ctrl+Z / Ctrl+Y) to reverse or reapply edits and removals.</p>
        <p><b>Search and category/error filters affect the table only.</b> They do not limit export.
        Removed rows and columns do affect saved copies and exports.</p>
        <h2>5. Customize appearance</h2>
        <p>Select a category in <b>Appearance</b> and choose its color. Use <b>Apply palette</b>
        for category colors, <b>Advanced styling</b> for additional options, and <b>Popup fields</b>
        to choose the information shown when a point is opened.
        Icon shapes and label styling apply to Google Earth; HTML uses the supported marker appearance.</p>
        <h2>6. Preview your points</h2>
        <p>The map starts hidden. Click <b>Show map</b> to display it.
        The coordinate overview works offline: drag to pan, scroll to zoom, and select a point
        to locate its row. It is a coordinate plot, not a distance-accurate basemap.
        Choose an online basemap when internet access is available.</p>
        <p>The preview samples up to 500 valid points. Check its coverage indicator;
        <b>Fit all</b> uses all valid points. If automatic preview is paused, refresh after edits.
        On smaller windows, use <b>Show map / Show data</b> to switch views.</p>
        <h2>7. Save your work and reusable settings</h2>
        <p><b>Save corrected copy…</b> saves the edited data separately. It does not preserve
        workbook formatting, formulas, or other sheets. Your original stays unchanged.</p>
        <p><b>Save preset</b> keeps mapping, selected popup fields, and styles, including custom icons.
        <b>Load preset</b> reuses those settings with another dataset. Presets contain no data rows.</p>
        <h2>8. Export and share</h2>
        <p>Choose the output format at the bottom, then click <b>Export map…</b> and choose a destination.
        Fix invalid coordinates first, or explicitly select <b>Valid rows only</b> to export valid points
        with an excluded-rows CSV report. The preview sample does not limit export.</p>
        <p><b>KMZ</b> bundles icons for Google Earth. <b>KML</b> standard icons and <b>HTML</b> maps
        need internet access. After export, use <b>Open file</b> or <b>Folder</b> to review the result.</p>
        <h2>Need help?</h2>
        <p>Use <b>Help → Contact &amp; Support</b>. Include the version/build, the steps you took,
        and the error text. Diagnostic logs are available from <b>Help → Open diagnostic folder</b>.</p>
        """
        dialog = self.create_help_dialog('Quick Guide', 'From your first import to a finished map', content)
        dialog.exec()

    def show_support(self):
        content = f"""
        <h2>Contact {escape_html(AUTHOR_NAME)}</h2>
        <p>For questions, issue reports, feature suggestions, or feedback about SitePin:</p>
        <p><b>Email</b><br><a href="mailto:{SUPPORT_EMAIL}">{SUPPORT_EMAIL}</a></p>
        <p><b>LinkedIn</b><br><a href="{SUPPORT_LINKEDIN}">{SUPPORT_LINKEDIN}</a></p>
        <h2>Help us understand the issue</h2>
        <p>Include <b>SitePin {VERSION} / Build {BUILD_NUMBER}</b>, what you expected,
        what happened, and the steps needed to reproduce it. Add the exact error message
        and a screenshot or a small sample without confidential site data when useful.</p>
        <p>Use <b>Help → Open diagnostic folder</b> to find logs if requested.</p>
        <p><b>Email support</b> opens your email application. <b>LinkedIn</b> opens your browser.
        No message or diagnostic information is sent automatically.</p>
        """
        dialog = self.create_help_dialog('Contact & Support', 'Questions, feedback, and assistance', content, support=True)
        dialog.exec()

    def save_corrected_copy(self):
        if self.processor.df is None:
            return False
        stem = os.path.splitext(self.processor.current_filepath)[0]
        path, selected = QFileDialog.getSaveFileName(self, "Save data-only corrected copy",
            stem + "_corrected.xlsx", "Excel (*.xlsx);;CSV (*.csv);;GeoJSON (*.geojson)")
        if not path:
            return False
        if not os.path.splitext(path)[1]:
            path += {"Excel (*.xlsx)": ".xlsx", "CSV (*.csv)": ".csv",
                     "GeoJSON (*.geojson)": ".geojson"}[selected]
        try:
            self.processor.save_corrected_copy(path)
        except Exception as error:
            LOG.exception("Corrected copy save failed")
            QMessageBox.critical(self, "Copy not saved", str(error))
            return False
        self.dirty = False
        self.update_data_status()
        QMessageBox.information(self, "Corrected copy saved",
            f"Saved to:\n{path}\n\nThis is a data-only copy of the loaded sheet. "
            "Workbook formatting, formulas and other sheets are not copied. The original is unchanged.")
        return True

    def confirm_discard_edits(self):
        if not self.dirty:
            return True
        answer = QMessageBox.warning(self, "Unsaved data edits",
            "Save a corrected copy before continuing? The original file is unchanged.",
            QMessageBox.StandardButton.Save | QMessageBox.StandardButton.Discard |
            QMessageBox.StandardButton.Cancel, QMessageBox.StandardButton.Cancel)
        if answer == QMessageBox.StandardButton.Save:
            return self.save_corrected_copy()
        return answer == QMessageBox.StandardButton.Discard

    def closeEvent(self, event):
        if self.import_busy or self.export_busy:
            event.ignore()
            answer = QMessageBox.question(self, 'Operation in progress',
                'Cancel the operation and close when it stops? Existing outputs stay protected. An export already committing may finish.',
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)
            if answer == QMessageBox.StandardButton.Yes:
                self.close_after_job = True
                self.cancel_job()
        elif self.confirm_discard_edits():
            self.preview_timer.stop()
            self.map_timeout.stop()
            self.web_view.stop()
            self.preview_file.cleanup()
            self.preset_icons.cleanup()
            event.accept()
        else:
            event.ignore()


def run_self_test(destination):
    """Distribution smoke check: native UI, readers, templates and local KMZ assets."""
    result = {'version': VERSION, 'checks': [], 'ok': False}
    try:
        with tempfile.TemporaryDirectory(prefix='spatialpin_selftest_') as folder:
            source = os.path.join(folder,'sample.csv')
            write_sample_csv(source)
            processor = KMLProcessor()
            ok,message = processor.load_and_parse_file(source)
            if not ok: raise ValueError(message)
            result['checks'].append('CSV import')
            excel = os.path.join(folder,'sample.xlsx')
            processor.save_corrected_copy(excel)
            reader = KMLProcessor()
            ok,message = reader.load_and_parse_file(excel)
            if not ok: raise ValueError(message)
            result['checks'].append('Excel read/write')
            selected,report = processor.export_selection(True)
            styles = {category:MainWindow.default_style(i) for i,category in enumerate(selected.unique_categories)}
            for suffix in ('kmz','html'):
                messages = queue.Queue()
                target = os.path.join(folder,'map.'+suffix)
                selected.export_with_report(target,styles,list(processor.target_cols.values()),messages,report)
                events = list(messages.queue)
                if not events or events[-1][0] != 'done': raise ValueError(str(events))
                if suffix == 'kmz':
                    with zipfile.ZipFile(target) as archive:
                        root = ET.fromstring(archive.read('doc.kml'))
                        for node in root.iter():
                            if node.tag.endswith('}href') and node.text not in archive.namelist(): raise ValueError('KMZ icon missing')
                result['checks'].append(suffix.upper()+' export')
            window = MainWindow()
            window.auto_preview.setChecked(False)
            window.load_file(source,asynchronous=False)
            window.plot_map()
            window.show()
            QApplication.processEvents()
            for width,height in ((1320,880),(900,640)):
                window.resize(width,height)
                QApplication.processEvents()
                if width==900:
                    window.btn_map.click()
                    QApplication.processEvents()
                    if window.overview.height()<120: raise ValueError('Compact map too small')
            if window.web_view.browser is not None: raise ValueError('Offline startup initialized Chromium')
            window.close()
            result['checks'].append('Native preview and compact layout')
        result['ok'] = True
    except Exception as error:
        LOG.exception('Distribution self-test failed')
        result['error'] = str(error)
    with open(destination,'w',encoding='utf-8') as handle:
        json.dump(result,handle,ensure_ascii=False,indent=2)
    return 0 if result['ok'] else 1


if __name__ == "__main__":
    configure_logging()
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setFont(QFont("Segoe UI", 10))
    def report_exception(kind, value, trace):
        LOG.error('Unhandled application error', exc_info=(kind,value,trace))
        QMessageBox.critical(None, 'Something went wrong', 'The operation could not finish. Your original input is protected.\nTry again or reopen the file.\nDiagnostic details: Help → Open diagnostic folder.')
    sys.excepthook = report_exception
    if len(sys.argv) == 3 and sys.argv[1] == "--self-test":
        sys.exit(run_self_test(sys.argv[2]))
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
