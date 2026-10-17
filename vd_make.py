from __future__ import annotations

import argparse
import csv
import io
import os
import re
import sys
from collections import OrderedDict, defaultdict
from dataclasses import dataclass, field
from datetime import datetime
from decimal import Decimal, InvalidOperation, ROUND_DOWN, ROUND_HALF_UP, ROUND_UP
from pathlib import Path
from typing import Callable, Iterable, Sequence

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    HRFlowable,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)


APP_TITLE = "CSV請求書PDF作成"
PAGE_WIDTH, PAGE_HEIGHT = A4
FONT_REGULAR = "InvoiceJP"
FONT_BOLD = "InvoiceJPBold"

REQUIRED_COLUMNS = (
    "請求書番号",
    "発行日",
    "支払期限",
    "顧客名",
    "郵便番号",
    "住所",
    "品目",
    "数量",
    "単価",
    "税率",
)

OPTIONAL_COLUMNS = ("電話番号", "メール")

ROUNDING_MODES = {
    "切り捨て": ROUND_DOWN,
    "四捨五入": ROUND_HALF_UP,
    "切り上げ": ROUND_UP,
}

ACCENT = colors.HexColor("#1F4E78")
ACCENT_LIGHT = colors.HexColor("#DCE6F1")
TEXT = colors.HexColor("#222222")
MUTED = colors.HexColor("#666666")
BORDER = colors.HexColor("#AAB7C4")
ROW_ALT = colors.HexColor("#F5F8FB")


class InvoicePdfError(Exception):
    """CSVまたはPDF生成に関する利用者向けエラー。"""


@dataclass(frozen=True)
class InvoiceLine:
    item: str
    quantity: Decimal
    unit_price: Decimal
    tax_rate: Decimal


@dataclass
class Invoice:
    invoice_number: str
    issue_date: str
    due_date: str
    customer_name: str
    postal_code: str
    address: str
    phone: str = ""
    email: str = ""
    lines: list[InvoiceLine] = field(default_factory=list)


@dataclass(frozen=True)
class IssuerInfo:
    name: str
    postal_code: str = ""
    address: str = ""
    phone: str = ""
    email: str = ""
    registration_number: str = ""
    bank_info: str = ""


def read_csv_rows(csv_path: str | Path) -> list[dict[str, str]]:
    """UTF-8 BOM、UTF-8、CP932のCSVを読み込む。"""

    path = Path(csv_path)
    if not path.is_file():
        raise InvoicePdfError(f"CSVファイルが見つかりません。\n{path}")

    raw = path.read_bytes()
    if not raw:
        raise InvoicePdfError("CSVファイルが空です。")

    decoded: str | None = None
    used_encoding = ""
    for encoding in ("utf-8-sig", "cp932", "utf-8"):
        try:
            decoded = raw.decode(encoding)
            used_encoding = encoding
            break
        except UnicodeDecodeError:
            continue

    if decoded is None:
        raise InvoicePdfError(
            "CSVの文字コードを判定できません。UTF-8 BOMまたはCP932で保存してください。"
        )

    sample = decoded[:8192]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",\t;")
    except csv.Error:
        dialect = csv.excel

    reader = csv.DictReader(io.StringIO(decoded), dialect=dialect)
    if reader.fieldnames is None:
        raise InvoicePdfError("CSVにヘッダー行がありません。")

    normalized_names = [str(name or "").strip() for name in reader.fieldnames]
    missing = [name for name in REQUIRED_COLUMNS if name not in normalized_names]
    if missing:
        raise InvoicePdfError(
            "CSVに必要な列がありません。\n"
            f"不足列: {', '.join(missing)}\n"
            f"検出した文字コード: {used_encoding}"
        )

    rows: list[dict[str, str]] = []
    for row_number, raw_row in enumerate(reader, start=2):
        normalized: dict[str, str] = {}
        for key, value in raw_row.items():
            normalized[str(key or "").strip()] = str(value or "").strip()

        if not any(normalized.values()):
            continue

        normalized["__row_number__"] = str(row_number)
        rows.append(normalized)

    if not rows:
        raise InvoicePdfError("CSVにデータ行がありません。")

    return rows


def parse_decimal(value: str, field_name: str, row_number: str) -> Decimal:
    cleaned = re.sub(r"[,，￥¥\s]", "", value)
    if not cleaned:
        raise InvoicePdfError(f"{row_number}行目の「{field_name}」が空です。")

    try:
        result = Decimal(cleaned)
    except InvalidOperation as exc:
        raise InvoicePdfError(
            f"{row_number}行目の「{field_name}」を数値として読めません: {value}"
        ) from exc

    if result < 0:
        raise InvoicePdfError(
            f"{row_number}行目の「{field_name}」に負の値は指定できません。"
        )
    return result


def normalize_postal_code(value: str) -> str:
    compact = value.replace(" ", "").replace("　", "")
    if re.fullmatch(r"\d{7}", compact):
        return f"{compact[:3]}-{compact[3:]}"
    return compact


def group_invoices(rows: Sequence[dict[str, str]]) -> list[Invoice]:
    """同じ請求書番号の行を品目明細として1請求書へまとめる。"""

    grouped: OrderedDict[str, Invoice] = OrderedDict()
    header_fields = (
        "発行日",
        "支払期限",
        "顧客名",
        "郵便番号",
        "住所",
        "電話番号",
        "メール",
    )

    for row in rows:
        row_number = row["__row_number__"]
        invoice_number = row.get("請求書番号", "").strip()
        if not invoice_number:
            raise InvoicePdfError(f"{row_number}行目の「請求書番号」が空です。")

        required_text_fields = ("発行日", "支払期限", "顧客名", "住所", "品目")
        for field_name in required_text_fields:
            if not row.get(field_name, "").strip():
                raise InvoicePdfError(f"{row_number}行目の「{field_name}」が空です。")

        if invoice_number not in grouped:
            grouped[invoice_number] = Invoice(
                invoice_number=invoice_number,
                issue_date=row["発行日"],
                due_date=row["支払期限"],
                customer_name=row["顧客名"],
                postal_code=normalize_postal_code(row.get("郵便番号", "")),
                address=row["住所"],
                phone=row.get("電話番号", ""),
                email=row.get("メール", ""),
            )
        else:
            invoice = grouped[invoice_number]
            existing = {
                "発行日": invoice.issue_date,
                "支払期限": invoice.due_date,
                "顧客名": invoice.customer_name,
                "郵便番号": invoice.postal_code,
                "住所": invoice.address,
                "電話番号": invoice.phone,
                "メール": invoice.email,
            }
            current = dict(row)
            current["郵便番号"] = normalize_postal_code(current.get("郵便番号", ""))
            inconsistent = [
                name
                for name in header_fields
                if current.get(name, "").strip() != existing.get(name, "").strip()
            ]
            if inconsistent:
                raise InvoicePdfError(
                    f"請求書番号「{invoice_number}」で請求先情報が一致しません。\n"
                    f"{row_number}行目の確認項目: {', '.join(inconsistent)}"
                )

        quantity = parse_decimal(row.get("数量", ""), "数量", row_number)
        unit_price = parse_decimal(row.get("単価", ""), "単価", row_number)
        tax_rate = parse_decimal(row.get("税率", ""), "税率", row_number)
        if quantity == 0:
            raise InvoicePdfError(f"{row_number}行目の「数量」は0より大きくしてください。")
        if tax_rate > 100:
            raise InvoicePdfError(f"{row_number}行目の「税率」が100を超えています。")

        grouped[invoice_number].lines.append(
            InvoiceLine(
                item=row["品目"],
                quantity=quantity,
                unit_price=unit_price,
                tax_rate=tax_rate,
            )
        )

    return list(grouped.values())


def quantize_yen(value: Decimal, rounding_name: str) -> Decimal:
    try:
        rounding = ROUNDING_MODES[rounding_name]
    except KeyError as exc:
        raise InvoicePdfError(f"未対応の端数処理です: {rounding_name}") from exc
    return value.quantize(Decimal("1"), rounding=rounding)


def calculate_invoice(
    invoice: Invoice,
    rounding_name: str,
) -> tuple[list[Decimal], dict[Decimal, Decimal], dict[Decimal, Decimal], Decimal]:
    line_amounts = [
        quantize_yen(line.quantity * line.unit_price, rounding_name)
        for line in invoice.lines
    ]

    taxable_by_rate: dict[Decimal, Decimal] = defaultdict(lambda: Decimal("0"))
    for line, amount in zip(invoice.lines, line_amounts):
        taxable_by_rate[line.tax_rate] += amount

    tax_by_rate = {
        rate: quantize_yen(amount * rate / Decimal("100"), rounding_name)
        for rate, amount in taxable_by_rate.items()
    }
    total = sum(taxable_by_rate.values(), Decimal("0")) + sum(
        tax_by_rate.values(), Decimal("0")
    )
    return line_amounts, dict(taxable_by_rate), tax_by_rate, total


def format_number(value: Decimal) -> str:
    if value == value.to_integral_value():
        return f"{int(value):,}"
    normalized = format(value.normalize(), "f")
    integer, dot, fraction = normalized.partition(".")
    return f"{int(integer):,}{dot}{fraction}" if dot else f"{int(integer):,}"


def format_yen(value: Decimal) -> str:
    return f"¥{int(value):,}"


def format_tax_rate(value: Decimal) -> str:
    return f"{format_number(value)}%"


def format_date(value: str) -> str:
    for pattern in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d"):
        try:
            parsed = datetime.strptime(value.strip(), pattern)
            return f"{parsed.year}年{parsed.month}月{parsed.day}日"
        except ValueError:
            continue
    return value


def _font_pair_from_explicit(font_path: Path) -> tuple[Path, Path]:
    stem = font_path.stem
    candidates = (
        font_path.with_name(f"{stem}-Bold{font_path.suffix}"),
        font_path.with_name(f"{stem}Bold{font_path.suffix}"),
        font_path.with_name(font_path.name.replace("Regular", "Bold")),
    )
    bold_path = next((candidate for candidate in candidates if candidate.is_file()), font_path)
    return font_path, bold_path


def register_japanese_fonts(font_path: str | Path | None = None) -> tuple[str, str, str]:
    """
    日本語フォントを登録する。

    戻り値の3番目は利用者へ表示できるフォント情報。
    フォントファイルが見つからない場合はCIDフォントへフォールバックする。
    """

    if FONT_REGULAR in pdfmetrics.getRegisteredFontNames():
        return FONT_REGULAR, FONT_BOLD, "登録済み日本語フォント"

    pairs: list[tuple[Path, Path]] = []
    if font_path:
        explicit = Path(font_path).expanduser()
        if not explicit.is_file():
            raise InvoicePdfError(f"指定したフォントが見つかりません。\n{explicit}")
        pairs.append(_font_pair_from_explicit(explicit))

    app_dir = Path(__file__).resolve().parent
    pairs.extend(
        [
            (
                app_dir / "fonts" / "NotoSansJP-Regular.ttf",
                app_dir / "fonts" / "NotoSansJP-Bold.ttf",
            ),
            (app_dir / "fonts" / "ipaexg.ttf", app_dir / "fonts" / "ipaexg.ttf"),
        ]
    )

    windows_dir = os.environ.get("WINDIR")
    if windows_dir:
        fonts_dir = Path(windows_dir) / "Fonts"
        pairs.extend(
            [
                (fonts_dir / "YuGothM.ttc", fonts_dir / "YuGothB.ttc"),
                (fonts_dir / "meiryo.ttc", fonts_dir / "meiryob.ttc"),
                (fonts_dir / "msgothic.ttc", fonts_dir / "msgothic.ttc"),
            ]
        )

    errors: list[str] = []
    for regular_path, bold_path in pairs:
        if not regular_path.is_file():
            continue
        if not bold_path.is_file():
            bold_path = regular_path
        try:
            pdfmetrics.registerFont(
                TTFont(FONT_REGULAR, str(regular_path), subfontIndex=0)
            )
            pdfmetrics.registerFont(
                TTFont(FONT_BOLD, str(bold_path), subfontIndex=0)
            )
            pdfmetrics.registerFontFamily(
                FONT_REGULAR,
                normal=FONT_REGULAR,
                bold=FONT_BOLD,
                italic=FONT_REGULAR,
                boldItalic=FONT_BOLD,
            )
            return FONT_REGULAR, FONT_BOLD, f"埋め込み: {regular_path.name}"
        except Exception as exc:  # 次候補を試すために捕捉する
            errors.append(f"{regular_path.name}: {exc}")

    try:
        pdfmetrics.registerFont(UnicodeCIDFont("HeiseiKakuGo-W5"))
        return "HeiseiKakuGo-W5", "HeiseiKakuGo-W5", "CIDフォント（非埋め込み）"
    except Exception as exc:
        detail = "\n".join(errors[-3:])
        raise InvoicePdfError(
            "日本語フォントを登録できません。fontsフォルダーに"
            "NotoSansJP-Regular.ttfを配置してください。\n"
            f"{detail}\n{exc}"
        ) from exc


def paragraph(text: str, style: ParagraphStyle) -> Paragraph:
    escaped = (
        str(text)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace("\r\n", "\n")
        .replace("\r", "\n")
        .replace("\n", "<br/>")
    )
    return Paragraph(escaped, style)


def create_styles(font_regular: str, font_bold: str) -> dict[str, ParagraphStyle]:
    base = getSampleStyleSheet()
    return {
        "title": ParagraphStyle(
            "InvoiceTitle",
            parent=base["Title"],
            fontName=font_bold,
            fontSize=24,
            leading=30,
            textColor=ACCENT,
            alignment=TA_CENTER,
            spaceAfter=4 * mm,
        ),
        "heading": ParagraphStyle(
            "Heading",
            parent=base["Heading2"],
            fontName=font_bold,
            fontSize=10,
            leading=15,
            textColor=ACCENT,
            wordWrap="CJK",
        ),
        "body": ParagraphStyle(
            "Body",
            parent=base["BodyText"],
            fontName=font_regular,
            fontSize=9,
            leading=14,
            textColor=TEXT,
            wordWrap="CJK",
        ),
        "body_bold": ParagraphStyle(
            "BodyBold",
            parent=base["BodyText"],
            fontName=font_bold,
            fontSize=9,
            leading=14,
            textColor=TEXT,
            wordWrap="CJK",
        ),
        "small": ParagraphStyle(
            "Small",
            parent=base["BodyText"],
            fontName=font_regular,
            fontSize=7.5,
            leading=11,
            textColor=MUTED,
            wordWrap="CJK",
        ),
        "amount": ParagraphStyle(
            "Amount",
            parent=base["BodyText"],
            fontName=font_bold,
            fontSize=16,
            leading=21,
            textColor=colors.white,
            alignment=TA_RIGHT,
        ),
        "table_header": ParagraphStyle(
            "TableHeader",
            parent=base["BodyText"],
            fontName=font_bold,
            fontSize=8,
            leading=11,
            textColor=colors.white,
            alignment=TA_CENTER,
            wordWrap="CJK",
        ),
        "table_left": ParagraphStyle(
            "TableLeft",
            parent=base["BodyText"],
            fontName=font_regular,
            fontSize=8,
            leading=11,
            textColor=TEXT,
            alignment=TA_LEFT,
            wordWrap="CJK",
        ),
        "table_right": ParagraphStyle(
            "TableRight",
            parent=base["BodyText"],
            fontName=font_regular,
            fontSize=8,
            leading=11,
            textColor=TEXT,
            alignment=TA_RIGHT,
        ),
    }


def make_page_callback(font_regular: str) -> Callable:
    def add_page_footer(pdf_canvas, document) -> None:
        pdf_canvas.saveState()
        pdf_canvas.setStrokeColor(BORDER)
        pdf_canvas.setLineWidth(0.3)
        pdf_canvas.line(18 * mm, 12 * mm, PAGE_WIDTH - 18 * mm, 12 * mm)
        pdf_canvas.setFont(font_regular, 7.5)
        pdf_canvas.setFillColor(MUTED)
        pdf_canvas.drawString(18 * mm, 7.5 * mm, "CSV Invoice PDF Generator")
        pdf_canvas.drawRightString(
            PAGE_WIDTH - 18 * mm,
            7.5 * mm,
            f"Page {document.page}",
        )
        pdf_canvas.restoreState()

    return add_page_footer


def build_invoice_story(
    invoice: Invoice,
    issuer: IssuerInfo,
    rounding_name: str,
    styles: dict[str, ParagraphStyle],
) -> list:
    line_amounts, taxable_by_rate, tax_by_rate, total = calculate_invoice(
        invoice, rounding_name
    )

    story: list = [paragraph("請 求 書", styles["title"])]

    meta_table = Table(
        [
            ["請求書番号", invoice.invoice_number],
            ["発行日", format_date(invoice.issue_date)],
            ["支払期限", format_date(invoice.due_date)],
        ],
        colWidths=[28 * mm, 46 * mm],
        hAlign="RIGHT",
    )
    meta_table.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, -1), styles["body"].fontName),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("TEXTCOLOR", (0, 0), (0, -1), MUTED),
                ("TEXTCOLOR", (1, 0), (1, -1), TEXT),
                ("ALIGN", (1, 0), (1, -1), "RIGHT"),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("LINEBELOW", (0, 0), (-1, -1), 0.25, BORDER),
            ]
        )
    )
    story.extend([meta_table, Spacer(1, 5 * mm)])

    customer_lines = [
        f"〒{invoice.postal_code}" if invoice.postal_code else "",
        invoice.address,
        f"{invoice.customer_name} 御中",
    ]
    if invoice.phone:
        customer_lines.append(f"TEL: {invoice.phone}")
    if invoice.email:
        customer_lines.append(f"E-mail: {invoice.email}")

    issuer_lines = [issuer.name or "発行者名を設定してください"]
    if issuer.postal_code:
        issuer_lines.append(f"〒{normalize_postal_code(issuer.postal_code)}")
    if issuer.address:
        issuer_lines.append(issuer.address)
    if issuer.phone:
        issuer_lines.append(f"TEL: {issuer.phone}")
    if issuer.email:
        issuer_lines.append(f"E-mail: {issuer.email}")
    if issuer.registration_number:
        issuer_lines.append(f"登録番号: {issuer.registration_number}")

    parties = Table(
        [
            [
                paragraph("ご請求先", styles["heading"]),
                paragraph("発行者", styles["heading"]),
            ],
            [
                paragraph("\n".join(customer_lines), styles["body"]),
                paragraph("\n".join(issuer_lines), styles["body"]),
            ],
        ],
        colWidths=[94 * mm, 76 * mm],
    )
    parties.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), ROW_ALT),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("BOX", (0, 0), (-1, -1), 0.5, BORDER),
                ("LINEBEFORE", (1, 0), (1, -1), 0.25, BORDER),
                ("LINEBELOW", (0, 0), (-1, 0), 0.25, BORDER),
                ("LEFTPADDING", (0, 0), (-1, -1), 7),
                ("RIGHTPADDING", (0, 0), (-1, -1), 7),
                ("TOPPADDING", (0, 0), (-1, -1), 7),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
            ]
        )
    )
    story.extend([parties, Spacer(1, 6 * mm)])

    amount_table = Table(
        [
            [
                paragraph("ご請求金額（税込）", styles["body_bold"]),
                paragraph(format_yen(total), styles["amount"]),
            ]
        ],
        colWidths=[60 * mm, 110 * mm],
    )
    amount_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (0, 0), ACCENT_LIGHT),
                ("BACKGROUND", (1, 0), (1, 0), ACCENT),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (1, 0), (1, 0), "RIGHT"),
                ("BOX", (0, 0), (-1, -1), 0.7, ACCENT),
                ("LEFTPADDING", (0, 0), (-1, -1), 10),
                ("RIGHTPADDING", (0, 0), (-1, -1), 10),
                ("TOPPADDING", (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ]
        )
    )
    story.extend([amount_table, Spacer(1, 7 * mm)])

    item_rows = [
        [
            paragraph("品目", styles["table_header"]),
            paragraph("数量", styles["table_header"]),
            paragraph("単価", styles["table_header"]),
            paragraph("税率", styles["table_header"]),
            paragraph("金額", styles["table_header"]),
        ]
    ]
    for line, amount in zip(invoice.lines, line_amounts):
        item_rows.append(
            [
                paragraph(line.item, styles["table_left"]),
                paragraph(format_number(line.quantity), styles["table_right"]),
                paragraph(format_yen(quantize_yen(line.unit_price, rounding_name)), styles["table_right"]),
                paragraph(format_tax_rate(line.tax_rate), styles["table_right"]),
                paragraph(format_yen(amount), styles["table_right"]),
            ]
        )

    items_table = Table(
        item_rows,
        colWidths=[78 * mm, 18 * mm, 27 * mm, 18 * mm, 29 * mm],
        repeatRows=1,
        splitByRow=1,
    )
    item_style = [
        ("BACKGROUND", (0, 0), (-1, 0), ACCENT),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.35, BORDER),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]
    for row_index in range(2, len(item_rows), 2):
        item_style.append(("BACKGROUND", (0, row_index), (-1, row_index), ROW_ALT))
    items_table.setStyle(TableStyle(item_style))
    story.extend([items_table, Spacer(1, 5 * mm)])

    subtotal = sum(taxable_by_rate.values(), Decimal("0"))
    total_rows: list[list] = [
        [paragraph("小計", styles["body"]), paragraph(format_yen(subtotal), styles["body"])]
    ]
    for rate in sorted(taxable_by_rate):
        total_rows.append(
            [
                paragraph(f"{format_tax_rate(rate)}対象額", styles["small"]),
                paragraph(format_yen(taxable_by_rate[rate]), styles["small"]),
            ]
        )
        total_rows.append(
            [
                paragraph(f"消費税（{format_tax_rate(rate)}）", styles["small"]),
                paragraph(format_yen(tax_by_rate[rate]), styles["small"]),
            ]
        )
    total_rows.append(
        [
            paragraph("合計", styles["body_bold"]),
            paragraph(format_yen(total), styles["body_bold"]),
        ]
    )

    totals = Table(total_rows, colWidths=[42 * mm, 36 * mm], hAlign="RIGHT")
    totals.setStyle(
        TableStyle(
            [
                ("ALIGN", (1, 0), (1, -1), "RIGHT"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LINEBELOW", (0, 0), (-1, -2), 0.25, BORDER),
                ("LINEABOVE", (0, -1), (-1, -1), 0.8, ACCENT),
                ("BACKGROUND", (0, -1), (-1, -1), ACCENT_LIGHT),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    story.extend([totals, Spacer(1, 6 * mm)])

    if issuer.bank_info:
        story.extend(
            [
                HRFlowable(width="100%", thickness=0.5, color=BORDER),
                Spacer(1, 3 * mm),
                paragraph("お振込先", styles["heading"]),
                paragraph(issuer.bank_info, styles["body"]),
                Spacer(1, 3 * mm),
            ]
        )

    story.append(
        paragraph(
            f"消費税は税率ごとの対象額に対し、{rounding_name}で計算しています。",
            styles["small"],
        )
    )
    return story


def generate_invoice_pdf(
    invoices: Sequence[Invoice],
    output_path: str | Path,
    issuer: IssuerInfo,
    rounding_name: str = "切り捨て",
    font_path: str | Path | None = None,
) -> tuple[Path, str]:
    if not invoices:
        raise InvoicePdfError("出力する請求書がありません。")

    destination = Path(output_path)
    if destination.suffix.lower() != ".pdf":
        destination = destination.with_suffix(".pdf")
    destination.parent.mkdir(parents=True, exist_ok=True)

    font_regular, font_bold, font_info = register_japanese_fonts(font_path)
    styles = create_styles(font_regular, font_bold)

    story: list = []
    for index, invoice in enumerate(invoices):
        if index:
            story.append(PageBreak())
        story.extend(build_invoice_story(invoice, issuer, rounding_name, styles))

    document = SimpleDocTemplate(
        str(destination),
        pagesize=A4,
        leftMargin=20 * mm,
        rightMargin=20 * mm,
        topMargin=14 * mm,
        bottomMargin=17 * mm,
        title="請求書",
        author=issuer.name or "CSV Invoice PDF Generator",
        subject="CSVから生成した請求書",
        pageCompression=1,
    )
    footer = make_page_callback(font_regular)
    document.build(story, onFirstPage=footer, onLaterPages=footer)
    return destination, font_info


def load_invoices(csv_path: str | Path) -> list[Invoice]:
    return group_invoices(read_csv_rows(csv_path))


def invoice_summary(invoices: Sequence[Invoice], rounding_name: str) -> tuple[int, int, Decimal]:
    detail_count = sum(len(invoice.lines) for invoice in invoices)
    grand_total = sum(
        (calculate_invoice(invoice, rounding_name)[3] for invoice in invoices),
        Decimal("0"),
    )
    return len(invoices), detail_count, grand_total


def launch_gui() -> int:
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox, ttk
    except ImportError as exc:
        raise InvoicePdfError(
            "Tkinterを利用できません。Python公式版をインストールしてください。"
        ) from exc

    root = tk.Tk()
    root.title(APP_TITLE)
    root.geometry("820x690")
    root.minsize(760, 630)

    style = ttk.Style(root)
    if "vista" in style.theme_names():
        style.theme_use("vista")

    main = ttk.Frame(root, padding=16)
    main.pack(fill="both", expand=True)
    main.columnconfigure(1, weight=1)

    csv_var = tk.StringVar()
    output_var = tk.StringVar()
    issuer_name_var = tk.StringVar(value="株式会社〇〇")
    issuer_postal_var = tk.StringVar()
    issuer_address_var = tk.StringVar()
    issuer_phone_var = tk.StringVar()
    issuer_email_var = tk.StringVar()
    registration_var = tk.StringVar()
    bank_var = tk.StringVar()
    rounding_var = tk.StringVar(value="切り捨て")
    font_var = tk.StringVar()
    status_var = tk.StringVar(value="CSVファイルを選択してください。")

    ttk.Label(main, text=APP_TITLE, font=("Yu Gothic UI", 16, "bold")).grid(
        row=0, column=0, columnspan=3, sticky="w", pady=(0, 14)
    )

    def add_path_row(row: int, label: str, variable, command, button_text: str) -> None:
        ttk.Label(main, text=label).grid(row=row, column=0, sticky="w", pady=5)
        ttk.Entry(main, textvariable=variable).grid(
            row=row, column=1, sticky="ew", padx=(10, 8), pady=5
        )
        ttk.Button(main, text=button_text, command=command, width=12).grid(
            row=row, column=2, sticky="e", pady=5
        )

    def choose_csv() -> None:
        selected = filedialog.askopenfilename(
            title="請求書CSVを選択",
            filetypes=[("CSVファイル", "*.csv"), ("すべてのファイル", "*.*")],
        )
        if not selected:
            return
        csv_var.set(selected)
        source = Path(selected)
        output_var.set(str(source.with_name(f"{source.stem}_請求書.pdf")))
        inspect_csv()

    def choose_output() -> None:
        initial = output_var.get().strip() or "請求書.pdf"
        selected = filedialog.asksaveasfilename(
            title="出力PDFを保存",
            defaultextension=".pdf",
            initialfile=Path(initial).name,
            filetypes=[("PDFファイル", "*.pdf")],
        )
        if selected:
            output_var.set(selected)

    def choose_font() -> None:
        selected = filedialog.askopenfilename(
            title="日本語フォントを選択（任意）",
            filetypes=[
                ("TrueTypeフォント", "*.ttf *.ttc"),
                ("すべてのファイル", "*.*"),
            ],
        )
        if selected:
            font_var.set(selected)

    add_path_row(1, "入力CSV", csv_var, choose_csv, "参照...")
    add_path_row(2, "出力PDF", output_var, choose_output, "保存先...")

    source_group = ttk.LabelFrame(main, text="発行者情報", padding=10)
    source_group.grid(row=3, column=0, columnspan=3, sticky="nsew", pady=(12, 8))
    source_group.columnconfigure(1, weight=1)
    source_group.columnconfigure(3, weight=1)

    issuer_fields = [
        ("発行者名", issuer_name_var, "郵便番号", issuer_postal_var),
        ("住所", issuer_address_var, "電話番号", issuer_phone_var),
        ("メール", issuer_email_var, "登録番号", registration_var),
        ("振込先", bank_var, "", None),
    ]
    for row_index, (label1, var1, label2, var2) in enumerate(issuer_fields):
        ttk.Label(source_group, text=label1).grid(row=row_index, column=0, sticky="w", pady=4)
        ttk.Entry(source_group, textvariable=var1).grid(
            row=row_index, column=1, sticky="ew", padx=(8, 16), pady=4
        )
        if var2 is not None:
            ttk.Label(source_group, text=label2).grid(row=row_index, column=2, sticky="w", pady=4)
            ttk.Entry(source_group, textvariable=var2).grid(
                row=row_index, column=3, sticky="ew", padx=(8, 0), pady=4
            )

    options = ttk.LabelFrame(main, text="PDF設定", padding=10)
    options.grid(row=4, column=0, columnspan=3, sticky="ew", pady=8)
    options.columnconfigure(3, weight=1)

    ttk.Label(options, text="消費税端数").grid(row=0, column=0, sticky="w")
    ttk.Combobox(
        options,
        textvariable=rounding_var,
        values=list(ROUNDING_MODES),
        state="readonly",
        width=12,
    ).grid(row=0, column=1, sticky="w", padx=(8, 20))
    ttk.Label(options, text="日本語フォント").grid(row=0, column=2, sticky="w")
    ttk.Entry(options, textvariable=font_var).grid(
        row=0, column=3, sticky="ew", padx=(8, 8)
    )
    ttk.Button(options, text="参照...", command=choose_font, width=10).grid(
        row=0, column=4, sticky="e"
    )

    preview = tk.Text(main, height=8, wrap="word", state="disabled")
    preview.grid(row=5, column=0, columnspan=3, sticky="nsew", pady=(8, 6))
    main.rowconfigure(5, weight=1)

    def set_preview(text: str) -> None:
        preview.configure(state="normal")
        preview.delete("1.0", "end")
        preview.insert("1.0", text)
        preview.configure(state="disabled")

    def inspect_csv() -> list[Invoice] | None:
        try:
            path = csv_var.get().strip()
            if not path:
                raise InvoicePdfError("入力CSVを選択してください。")
            invoices = load_invoices(path)
            count, details, total = invoice_summary(invoices, rounding_var.get())
            lines = [
                "CSVの確認結果",
                f"請求書数: {count}件",
                f"明細行数: {details}行",
                f"合計請求額: {format_yen(total)}",
                "",
            ]
            for invoice in invoices:
                amount = calculate_invoice(invoice, rounding_var.get())[3]
                lines.append(
                    f"{invoice.invoice_number} / {invoice.customer_name} / "
                    f"{len(invoice.lines)}明細 / {format_yen(amount)}"
                )
            set_preview("\n".join(lines))
            status_var.set("CSVを正常に読み込みました。")
            return invoices
        except Exception as exc:
            set_preview("")
            status_var.set("CSVを確認できませんでした。")
            messagebox.showerror(APP_TITLE, str(exc))
            return None

    def create_pdf() -> None:
        invoices = inspect_csv()
        if invoices is None:
            return

        output = output_var.get().strip()
        if not output:
            messagebox.showerror(APP_TITLE, "出力PDFの保存先を指定してください。")
            return

        output_path = Path(output)
        if output_path.exists():
            overwrite = messagebox.askyesno(
                APP_TITLE,
                f"同名のPDFが存在します。上書きしますか？\n{output_path}",
            )
            if not overwrite:
                return

        issuer = IssuerInfo(
            name=issuer_name_var.get().strip(),
            postal_code=issuer_postal_var.get().strip(),
            address=issuer_address_var.get().strip(),
            phone=issuer_phone_var.get().strip(),
            email=issuer_email_var.get().strip(),
            registration_number=registration_var.get().strip(),
            bank_info=bank_var.get().strip(),
        )

        try:
            root.configure(cursor="watch")
            root.update_idletasks()
            result, font_info = generate_invoice_pdf(
                invoices=invoices,
                output_path=output_path,
                issuer=issuer,
                rounding_name=rounding_var.get(),
                font_path=font_var.get().strip() or None,
            )
            status_var.set(f"PDFを作成しました: {result}")
            messagebox.showinfo(
                APP_TITLE,
                f"PDFを作成しました。\n\n{result}\n\nフォント: {font_info}",
            )
        except Exception as exc:
            status_var.set("PDF生成に失敗しました。")
            messagebox.showerror(APP_TITLE, str(exc))
        finally:
            root.configure(cursor="")

    buttons = ttk.Frame(main)
    buttons.grid(row=6, column=0, columnspan=3, sticky="e", pady=(8, 6))
    ttk.Button(buttons, text="CSVを確認", command=inspect_csv, width=15).pack(
        side="left", padx=(0, 8)
    )
    ttk.Button(buttons, text="PDFを生成", command=create_pdf, width=15).pack(side="left")

    ttk.Separator(main).grid(row=7, column=0, columnspan=3, sticky="ew", pady=(8, 6))
    ttk.Label(main, textvariable=status_var, foreground="#44546A").grid(
        row=8, column=0, columnspan=3, sticky="w"
    )

    root.mainloop()
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="CSVから請求書PDFを生成します。")
    parser.add_argument("--input", help="入力CSV。省略時はGUIを起動します。")
    parser.add_argument("--output", help="出力PDF")
    parser.add_argument("--issuer-name", default="株式会社〇〇")
    parser.add_argument("--issuer-postal-code", default="")
    parser.add_argument("--issuer-address", default="")
    parser.add_argument("--issuer-phone", default="")
    parser.add_argument("--issuer-email", default="")
    parser.add_argument("--registration-number", default="")
    parser.add_argument("--bank-info", default="")
    parser.add_argument("--rounding", choices=list(ROUNDING_MODES), default="切り捨て")
    parser.add_argument("--font", help="日本語TTF/TTCフォント（任意）")
    return parser


def main(argv: Iterable[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if not args.input:
        try:
            return launch_gui()
        except InvoicePdfError as exc:
            print(f"エラー: {exc}", file=sys.stderr)
            return 1

    if not args.output:
        print("エラー: --input指定時は--outputも必要です。", file=sys.stderr)
        return 2

    try:
        invoices = load_invoices(args.input)
        issuer = IssuerInfo(
            name=args.issuer_name,
            postal_code=args.issuer_postal_code,
            address=args.issuer_address,
            phone=args.issuer_phone,
            email=args.issuer_email,
            registration_number=args.registration_number,
            bank_info=args.bank_info,
        )
        result, font_info = generate_invoice_pdf(
            invoices=invoices,
            output_path=args.output,
            issuer=issuer,
            rounding_name=args.rounding,
            font_path=args.font,
        )
        count, details, total = invoice_summary(invoices, args.rounding)
        print(f"PDF: {result.resolve()}")
        print(f"請求書: {count}件 / 明細: {details}行 / 合計: {format_yen(total)}")
        print(f"フォント: {font_info}")
        return 0
    except Exception as exc:
        print(f"エラー: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
