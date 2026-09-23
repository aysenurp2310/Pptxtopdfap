#!/usr/bin/env python3
"""
PPTX -> PDF Telegram Botu
--------------------------
Kullanıcının gönderdiği .pptx dosyasını, görsel ve metin kalitesini
bozmadan PDF'e dönüştürüp geri gönderir.

Dönüştürme motoru: LibreOffice (soffice --headless).
LibreOffice, PowerPoint'i gerçek bir render motoruyla işlediği için
görselleri yeniden sıkıştırmaz / rasterize etmez; metinler vektörel
kalır, gömülü fontlar ve görseller orijinal çözünürlüğünde kalır.
Bu yüzden "python-pptx ile manuel PDF üretme" gibi kalite kaybına
yol açan yöntemler yerine LibreOffice tercih edildi.

Kurulum:
    pip install python-telegram-bot --upgrade
    sudo apt-get install libreoffice        # soffice komutu için

Çalıştırma:
    export TELEGRAM_BOT_TOKEN="123456:ABC-DEF..."
    python3 bot.py
"""

import asyncio
import gc
import io
import logging
import math
import os
import re
import shutil
import subprocess
import tempfile
import threading
import uuid
import zipfile

import gdown
import requests
from lxml import etree
from PIL import Image as PILImage
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn
from pptx.util import Pt
from pypdf import PdfReader, PdfWriter
from telegram import Update
from telegram.constants import ChatAction, ParseMode
from telegram.ext import (
    Application,
    ApplicationBuilder,
    ContextTypes,
    MessageHandler,
    CommandHandler,
    filters,
)

# --------------------------------------------------------------------------- #
# Ayarlar
# --------------------------------------------------------------------------- #

BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "")
MAX_FILE_SIZE_MB = 500          # Not: Telegram'ın standart Bot API'si, bota
                                 # DOĞRUDAN yüklenen dosyalarda indirme
                                 # sınırını 20 MB ile kısıtlar (bizim
                                 # kontrolümüz dışında bir Telegram kuralı).
                                 # Bu sınır Google Drive linki ile gönderilen
                                 # dosyalar için geçerli değildir; gdown
                                 # dosyayı doğrudan Drive'dan indirir.
SOFFICE_TIMEOUT_SEC = 600       # Büyük dosyalar için dönüştürme zaman aşımı (10 dk)

# Büyük/ağır sunumları LibreOffice'e tek seferde vermek yerine, aşağıdaki
# eşiklerden biri aşıldığında slayt gruplarına bölüp ayrı ayrı dönüştürüp
# sonra tek PDF'te birleştiriyoruz. Bu, tek bir dev dosyanın LibreOffice'in
# bellek/CPU sınırlarını (özellikle Railway gibi kısıtlı sunucularda)
# zorlamasını ve zaman aşımına takılmasını önlemeye yardımcı olur.
CHUNK_SLIDE_THRESHOLD = 100     # bu slayt sayısından fazlaysa böl (yalnızca aşırı durumlar için; asıl tetikleyici dosya boyutu)
CHUNK_FILE_SIZE_MB_THRESHOLD = 30  # bu boyuttan büyükse de böl (ağır medya) — parçalamanın asıl sebebi bu
CHUNK_SIZE = 15                 # her parçada kaç slayt olacak

# python-pptx ile bir dosyayı açmak (font tarama, autofit düzeltmesi,
# parçalama) dosyanın TAMAMINI belleğe yükler. Ama artık bu kontrol
# GÖRSEL SIKIŞTIRMADAN SONRAKİ boyuta bakıyor — ve görsel sıkıştırmanın
# kendisi (çok daha ağır bir python-pptx işlemi, görselleri gerçekten
# decode/encode ediyor) 484MB'lık dosyalarda bile başarıyla çalıştığı
# doğrulandı. Bu yüzden font/autofit gibi çok daha hafif işlemler için
# eşiği yüksek tutuyoruz; aksi hâlde büyük dosyalarda metin taşması/
# görsel kayması sorunları geri dönüyordu.
MEMORY_SAFE_PPTX_MB_THRESHOLD = 400

# Otomatik görsel sıkıştırma: dosya bu boyutu aşarsa, PowerPoint'in
# "Resimleri Sıkıştır" özelliğine benzer şekilde görselleri otomatik
# küçültüp yeniden sıkıştırıyoruz. Bu hem Telegram'ın gönderme sınırını
# aşma ihtimalini azaltır hem de LibreOffice'in bellek yükünü düşürür.
IMAGE_COMPRESS_THRESHOLD_MB = 50
IMAGE_MAX_DIMENSION_PX = 1280   # ekranda gösterim için fazlasıyla yeterli
IMAGE_JPEG_QUALITY = 60

# İlk sıkıştırma turundan sonra dosya hâlâ bu boyutun üzerindeyse
# (örn. slaytlarda az sayıda ama aşırı yüksek çözünürlüklü görsel varsa),
# daha da agresif ayarlarla ikinci bir sıkıştırma turu denenir.
IMAGE_SECOND_PASS_THRESHOLD_MB = 80
IMAGE_MAX_DIMENSION_PX_AGGRESSIVE = 900
IMAGE_JPEG_QUALITY_AGGRESSIVE = 45

# Telegram Bot API'sinin bir botun GÖNDEREBİLECEĞİ dosya boyutu sınırı
# 50 MB'dır (botun bir dosyayı İNDİREBİLECEĞİ 20 MB sınırından farklı
# ve ayrı bir kısıtlamadır). Kayıpsız kalite için görselleri koruduğumuzdan
# büyük/görsel ağırlıklı sunumlardan çıkan PDF bu sınırı aşabilir; bu
# durumda PDF'i parçalara bölüp ayrı ayrı göndeririz.
MAX_TELEGRAM_SEND_BYTES = 45 * 1024 * 1024  # 45MB (50MB'a güvenlik payı)
ALLOWED_EXTENSIONS = (".pptx", ".ppt", ".potx", ".pptm")

# Google Drive dosya/paylaşım linklerini yakalamak için desen
DRIVE_URL_PATTERN = re.compile(r"https?://(drive|docs)\.google\.com/\S+", re.IGNORECASE)

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger("pptx2pdf_bot")


# --------------------------------------------------------------------------- #
# Genel amaçlı font tespiti ve otomatik kurulum
# --------------------------------------------------------------------------- #
# Amaç: Sadece bilinen tek bir fontu değil, ileride gelecek herhangi bir
# sunumdaki eksik fontu da mümkün olduğunca otomatik çözmek. Akış:
#   1) Dosyada kullanılan tüm font isimleri (tema + çalıştırma/run seviyesi
#      + tablo hücreleri) çıkarılır.
#   2) Sistemde zaten kurulu olanlar atlanır.
#   3) Bilinen bir açık kaynak karşılığı varsa (FONT_ALIASES) o zaten
#      build sırasında kurulu olduğu için ekstra işlem gerekmez.
#   4) Kalanlar için Google Fonts'un herkese açık CSS API'si üzerinden
#      aynı isimde bir font aranır ve varsa indirilip sisteme kurulur.
#      Bu, GitHub deposundaki dosya adlandırma biçimini (bazı fontlar artık
#      "variable font" tek dosya olarak dağıtıldığı için) tahmin etmeye
#      çalışmaktan çok daha güvenilirdir.
# Bulunamayan (Google Fonts kataloğunda da olmayan, tamamen özel/lisanslı)
# fontlar için LibreOffice'in varsayılan ikamesi kullanılmaya devam eder;
# bu durumda kesin bir garanti verilemez.

FONT_ALIASES = {
    # Yaygın Microsoft fontları için, Google Fonts'ta bulunmayan ama
    # açık kaynaklı karşılığı build sırasında zaten kurulmuş olanlar.
    # (Karşılıklar Dockerfile'da fontconfig ile eşleştirilir.)
    "tw cen mt": "Poppins",
    "calibri": "Carlito",
    "cambria": "Caladea",
}

DYNAMIC_FONT_DIR = "/usr/share/fonts/truetype/dynamic"
_GOOGLE_FONTS_CSS_URL = "https://fonts.googleapis.com/css2?family={family}"
# Eski bir tarayıcı User-Agent'ı göndermek, Google'ın woff2 yerine
# doğrudan .ttf font dosyası linki döndürmesini sağlar (LibreOffice/
# fontconfig woff2'yi güvenilir şekilde desteklemez).
_OLD_BROWSER_UA = "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1)"

_dynamic_font_attempted = set()   # bu süreç ömrü boyunca denenen fontlar
_dynamic_font_lock = threading.Lock()


def _iter_text_frames(shapes):
    """
    Bir slayttaki tüm metin çerçevelerini dolaşır: normal metin kutuları,
    gruplanmış şekillerin içindekiler (iç içe olsa bile) ve tablo
    hücrelerindeki metinler dahil.
    """
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _iter_text_frames(shape.shapes)
        elif getattr(shape, "has_table", False):
            for row in shape.table.rows:
                for cell in row.cells:
                    yield cell.text_frame
        elif getattr(shape, "has_text_frame", False):
            yield shape.text_frame


def _extract_theme_fonts(pptx_path: str) -> set:
    """Pptx içindeki tema dosyalarından (theme1.xml, theme2.xml, ...)
    ana/gövde font isimlerini çıkarır."""
    fonts = set()
    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    try:
        with zipfile.ZipFile(pptx_path) as z:
            theme_files = [
                n for n in z.namelist()
                if re.match(r"ppt/theme/theme\d+\.xml$", n)
            ]
            for tf in theme_files:
                root = etree.fromstring(z.read(tf))
                for tag in ("majorFont", "minorFont"):
                    el = root.find(f".//a:fontScheme/a:{tag}/a:latin", ns)
                    if el is not None:
                        typeface = el.get("typeface")
                        if typeface and not typeface.startswith("+"):
                            fonts.add(typeface)
    except Exception:  # noqa: BLE001
        logger.exception("Tema fontları okunurken hata oluştu")
    return fonts


def extract_fonts_used(pptx_path: str) -> set:
    """Bir pptx dosyasında (tema dahil) kullanılan tüm font isimlerini döner."""
    fonts = set()
    try:
        prs = Presentation(pptx_path)
        for slide in prs.slides:
            for text_frame in _iter_text_frames(slide.shapes):
                for paragraph in text_frame.paragraphs:
                    for run in paragraph.runs:
                        name = run.font.name
                        if name and not name.startswith("+"):
                            fonts.add(name)
    except Exception:  # noqa: BLE001
        logger.exception("Font taraması sırasında hata oluştu")

    fonts |= _extract_theme_fonts(pptx_path)
    return {f.strip() for f in fonts if f and f.strip()}


def _get_installed_font_families() -> set:
    """fc-list ile sistemde kurulu olan tüm font ailesi isimlerini
    (küçük harfe çevrilmiş) döner."""
    try:
        result = subprocess.run(
            ["fc-list", ":", "family"],
            capture_output=True, text=True, timeout=15,
        )
        families = set()
        for line in result.stdout.splitlines():
            for name in line.split(","):
                cleaned = name.strip().lower()
                if cleaned:
                    families.add(cleaned)
        return families
    except Exception:  # noqa: BLE001
        logger.exception("fc-list çalıştırılamadı")
        return set()


def _download_google_font(family_name: str) -> bool:
    """
    Google Fonts'un herkese açık CSS API'sinden verilen isimde bir font
    aramayı ve indirmeyi dener. Bulunup indirilirse True döner.
    """
    try:
        family_param = requests.utils.quote(family_name.strip())
        url = _GOOGLE_FONTS_CSS_URL.format(family=family_param)
        css_resp = requests.get(
            url, headers={"User-Agent": _OLD_BROWSER_UA}, timeout=8
        )
        if css_resp.status_code != 200:
            return False

        match = re.search(
            r"url\((https://fonts\.gstatic\.com/[^)]+?\.ttf)\)", css_resp.text
        )
        if not match:
            return False

        font_resp = requests.get(match.group(1), timeout=15)
        if font_resp.status_code != 200 or len(font_resp.content) < 1000:
            return False

        os.makedirs(DYNAMIC_FONT_DIR, exist_ok=True)
        safe_name = re.sub(r"[^A-Za-z0-9]", "", family_name)
        out_path = os.path.join(DYNAMIC_FONT_DIR, f"{safe_name}-Regular.ttf")
        with open(out_path, "wb") as fh:
            fh.write(font_resp.content)
        return True

    except Exception:  # noqa: BLE001
        logger.exception("Google Fonts'tan '%s' indirilirken hata oluştu", family_name)
        return False


def ensure_fonts_available(pptx_path: str) -> None:
    """
    Dosyada kullanılan fontlardan sistemde kurulu olmayanları tespit eder
    ve mümkünse otomatik olarak temin eder (bilinen ikame veya Google
    Fonts'tan canlı indirme). Bulunamayan fontlar için LibreOffice'in
    varsayılan ikamesi kullanılmaya devam eder — bu fonksiyon en iyi
    çabayı gösterir, %100 garanti vermez.
    """
    used_fonts = extract_fonts_used(pptx_path)
    if not used_fonts:
        return

    installed = _get_installed_font_families()
    downloaded_any = False

    with _dynamic_font_lock:
        for font_name in used_fonts:
            key = font_name.lower()

            if key in installed:
                continue
            if key in FONT_ALIASES:
                # Build sırasında zaten kurulu bilinen bir ikamesi var
                # (Dockerfile'daki fontconfig eşleştirmesi devreye girer).
                continue
            if key in _dynamic_font_attempted:
                continue

            _dynamic_font_attempted.add(key)  # tekrar denemeyi engelle

            if _download_google_font(font_name):
                logger.info("Font otomatik indirildi: %s", font_name)
                downloaded_any = True

    if downloaded_any:
        try:
            subprocess.run(
                ["fc-cache", "-f"], capture_output=True, timeout=30
            )
        except Exception:  # noqa: BLE001
            logger.exception("fc-cache yenilenemedi")


# --------------------------------------------------------------------------- #
# Yardımcı fonksiyonlar
# --------------------------------------------------------------------------- #

def find_soffice() -> str:
    """Sistemde kurulu LibreOffice çalıştırılabilir dosyasını bulur."""
    for candidate in ("soffice", "libreoffice"):
        path = shutil.which(candidate)
        if path:
            return path
    raise RuntimeError(
        "LibreOffice (soffice) bulunamadı. Kurulum için: "
        "sudo apt-get install libreoffice"
    )


def convert_pptx_to_pdf(input_path: str, output_dir: str) -> str:
    """
    LibreOffice'i headless modda çalıştırarak pptx -> pdf dönüştürür.
    Yüksek kaliteli görsel/metin çıktısı için PDF export filtre
    seçenekleri ayarlanır (JPEG sıkıştırması kapalı, çözünürlük yüksek).
    """
    soffice = find_soffice()

    # PDF export filtre seçenekleri:
    # - Quality: JPEG kalitesi (0-100) -> görseller için 100
    # - ReduceImageResolution: false -> görselleri küçültme
    # - UseLosslessCompression: true -> mümkünse kayıpsız sıkıştırma
    filter_options = (
        'impress_pdf_Export:'
        'Quality=100,'
        'ReduceImageResolution=false,'
        'UseLosslessCompression=true,'
        'ExportNotes=false'
    )

    cmd = [
        soffice,
        "--headless",
        "--norestore",
        "--convert-to",
        f"pdf:{filter_options}",
        "--outdir",
        output_dir,
        input_path,
    ]

    logger.info("Dönüştürme komutu çalıştırılıyor: %s", " ".join(cmd))

    # Her dönüştürme için ayrı bir kullanıcı profili (user installation)
    # kullanmak, eşzamanlı isteklerde soffice çakışmalarını önler.
    user_profile_dir = os.path.join(output_dir, "lo_profile")
    env = os.environ.copy()

    cmd_with_profile = cmd + [
        f"-env:UserInstallation=file://{user_profile_dir}"
    ]

    result = subprocess.run(
        cmd_with_profile,
        cwd=output_dir,
        capture_output=True,
        text=True,
        timeout=SOFFICE_TIMEOUT_SEC,
        env=env,
    )

    base_name = os.path.splitext(os.path.basename(input_path))[0]
    pdf_path = os.path.join(output_dir, base_name + ".pdf")

    # ÖNEMLİ: LibreOffice bazen PDF'i başarıyla oluşturduğu hâlde
    # sıfır olmayan bir çıkış koduyla (returncode != 0) kapanabiliyor
    # — bilinen, zararsız bir tuhaflık. Bu yüzden önce çıkış koduna
    # değil, dosyanın GERÇEKTEN oluşup oluşmadığına bakıyoruz; dosya
    # varsa ve boş değilse, çıkış kodu ne olursa olsun başarılı sayarız.
    if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0:
        if result.returncode != 0:
            logger.warning(
                "soffice sıfır olmayan kodla (%s) çıktı ama PDF yine de "
                "oluşmuş, başarılı sayılıyor. stderr: %s",
                result.returncode, result.stderr,
            )
        return pdf_path

    logger.error("soffice stderr: %s", result.stderr)
    logger.error("soffice stdout: %s", result.stdout)
    raise RuntimeError(
        f"LibreOffice dönüştürme hatası (çıkış kodu {result.returncode}):\n"
        f"{result.stderr or result.stdout}"
    )


# --------------------------------------------------------------------------- #
# Büyük dosyalar için: parçalama, ayrı ayrı dönüştürme, birleştirme
# --------------------------------------------------------------------------- #

def _keep_only_slides(prs: Presentation, keep_indices: set) -> None:
    """
    Verilen Presentation nesnesinde SADECE keep_indices'teki (0-tabanlı)
    slaytları bırakır, diğerlerini tamamen kaldırır.

    ÖNEMLİ: Sadece sldIdLst'ten çıkarmak yetmez — altta yatan ilişkiyi
    (rel) de koparmak gerekir (prs.part.drop_rel). Aksi hâlde slayt
    parçası pakette "yetim" olarak kalır; bu dosya tekrar açılıp
    kaydedildiğinde (örn. sıkıştırma adımında) zip içinde çakışan
    parça adları (ör. iki adet "slide1.xml") oluşmasına ve dosyanın
    bozulmasına yol açabilir.
    """
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    for idx in reversed(range(len(slides))):
        if idx not in keep_indices:
            slide_id_elem = slides[idx]
            r_id = slide_id_elem.get(qn("r:id"))
            xml_slides.remove(slide_id_elem)
            if r_id:
                try:
                    prs.part.drop_rel(r_id)
                except Exception:  # noqa: BLE001
                    pass


def split_pptx_into_chunks(input_path: str, work_dir: str, chunk_size: int) -> list:
    """
    Bir pptx dosyasını, her biri en fazla chunk_size slayt içeren ayrı
    pptx dosyalarına böler. Bölmeye gerek yoksa (slayt sayısı zaten
    küçükse) tek elemanlı [input_path] listesi döner.
    """
    prs_full = Presentation(input_path)
    total = len(prs_full.slides)

    if total <= chunk_size:
        return [input_path]

    chunk_paths = []
    num_chunks = math.ceil(total / chunk_size)

    for c in range(num_chunks):
        start = c * chunk_size
        end = min(start + chunk_size, total)
        keep = set(range(start, end))

        # Her parça için orijinal dosyanın taze bir kopyasını aç, böylece
        # önceki parçalarda yapılan silmeler birbirini etkilemez.
        prs_chunk = Presentation(input_path)
        _keep_only_slides(prs_chunk, keep)

        chunk_path = os.path.join(work_dir, f"chunk_{c:03d}.pptx")
        prs_chunk.save(chunk_path)
        chunk_paths.append(chunk_path)

    return chunk_paths


def split_pdf_by_size(
    pdf_path: str, work_dir: str, max_bytes: int = MAX_TELEGRAM_SEND_BYTES
) -> list:
    """
    Bir PDF dosyasını, Telegram'ın gönderme sınırının altında kalacak
    şekilde sayfa bazında parçalara böler. Boyutun sayfalar arasında
    kabaca eşit dağıldığı varsayılır (görsel yoğunluğu sayfadan sayfaya
    değişebileceği için kesin bir garanti değildir, ama pratikte iyi
    sonuç verir). Bölmeye gerek yoksa tek elemanlı [pdf_path] döner.
    """
    size = os.path.getsize(pdf_path)
    if size <= max_bytes:
        return [pdf_path]

    reader = PdfReader(pdf_path)
    total_pages = len(reader.pages)
    if total_pages <= 1:
        return [pdf_path]  # tek sayfa varsa bölecek bir şey yok

    num_parts = math.ceil(size / max_bytes)
    pages_per_part = max(1, math.ceil(total_pages / num_parts))

    part_paths = []
    for start in range(0, total_pages, pages_per_part):
        writer = PdfWriter()
        for p in range(start, min(start + pages_per_part, total_pages)):
            writer.add_page(reader.pages[p])

        part_num = len(part_paths) + 1
        part_path = os.path.join(work_dir, f"part_{part_num:02d}.pdf")
        with open(part_path, "wb") as fh:
            writer.write(fh)
        part_paths.append(part_path)

    return part_paths


def merge_pdfs(pdf_paths: list, output_path: str) -> None:
    """Verilen PDF dosyalarını sırasıyla tek bir PDF'te birleştirir."""
    writer = PdfWriter()
    try:
        for pdf_path in pdf_paths:
            writer.append(pdf_path)
        with open(output_path, "wb") as fh:
            writer.write(fh)
    finally:
        writer.close()


def convert_pptx_to_pdf_chunked(
    input_path: str,
    work_dir: str,
    chunk_size: int = CHUNK_SIZE,
    progress_callback=None,
) -> tuple:
    
