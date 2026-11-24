import pandas as pd
import matplotlib

matplotlib.use("Agg")
matplotlib.rcParams["font.family"] = "DejaVu Sans"
import matplotlib.pyplot as plt
import io
from datetime import datetime
from pathlib import Path

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

NAVY = RGBColor(0x1A, 0x3D, 0x7C)
GREY = RGBColor(0x66, 0x66, 0x66)
HEADER_FILL = "1A3D7C"
ROW_FILL = "EEF3FA"

DESCRIBE_RU = {
    "count": "Количество", "mean": "Среднее", "std": "Ст. отклонение",
    "min": "Мин", "25%": "25%", "50%": "Медиана", "75%": "75%", "max": "Макс",
}

TYPE_RU = {
    "int64": "Целое число",
    "float64": "Дробное число",
    "object": "Текст",
    "str": "Текст",
    "datetime64[us]": "Дата/Время",
    "datetime64[ns]": "Дата/Время",
    "bool": "Логический"
}


# ---------- вспомогательные ----------

def clean_column_names(columns):
    """Убирает случайное дублирование названий (например, ЦенаЦена -> Цена)"""
    cleaned = []
    for col in columns:
        col_str = str(col).strip()
        half = len(col_str) // 2
        # Если длина четная и первая половина равна второй - обрезаем
        if len(col_str) % 2 == 0 and col_str[:half] == col_str[half:]:
            cleaned.append(col_str[:half])
        else:
            cleaned.append(col_str)
    return cleaned


def fig_to_bytes(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=130, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf


def shade_cell(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), color_hex)
    tcPr.append(shd)


def style_table(doc, df, title=None):
    if title:
        p = doc.add_paragraph()
        run = p.add_run(title)
        run.bold = True
        run.font.color.rgb = NAVY
        run.font.size = Pt(11)

    table = doc.add_table(rows=1, cols=len(df.columns))
    table.style = "Table Grid"

    hdr = table.rows[0].cells
    for i, col in enumerate(df.columns):
        hdr[i].text = str(col)
        shade_cell(hdr[i], HEADER_FILL)
        for paragraph in hdr[i].paragraphs:
            for run in paragraph.runs:
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                run.font.bold = True
                run.font.size = Pt(10)

    for idx, row in enumerate(df.itertuples(index=False)):
        cells = table.add_row().cells
        for i, val in enumerate(row):
            # Красиво форматируем дробные числа
            if isinstance(val, float) and not pd.isna(val):
                val_str = f"{val:.2f}"
            else:
                val_str = "" if pd.isna(val) else str(val)

            cells[i].text = val_str
            for paragraph in cells[i].paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(10)
            if idx % 2 == 1:
                shade_cell(cells[i], ROW_FILL)
    return table


def add_section(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        run.font.color.rgb = NAVY
    return h


def add_caption(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.italic = True
    run.font.size = Pt(9)
    run.font.color.rgb = GREY


# ---------- отчёт ----------

def build_docx_report(df, source_name, output="report.docx"):
    # Очищаем названия колонок от багов Excel
    df.columns = clean_column_names(df.columns)

    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    # --- Обложка ---
    doc.add_paragraph("\n")
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = title.add_run("Аналитический отчёт")
    r.bold = True
    r.font.size = Pt(28)
    r.font.color.rgb = NAVY

    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = subtitle.add_run(f"по файлу «{source_name}»")
    r.italic = True
    r.font.size = Pt(14)
    r.font.color.rgb = GREY

    date_p = doc.add_paragraph()
    date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = date_p.add_run(f"Сформирован: {datetime.now().strftime('%d.%m.%Y %H:%M')}")
    r.font.size = Pt(10)
    r.font.color.rgb = GREY

    doc.add_page_break()

    # --- 1. Обзор ---
    add_section(doc, "1. Обзор", level=1)
    p = doc.add_paragraph("Общая информация о таблице:")
    p.paragraph_format.space_after = Pt(6)

    overview = pd.DataFrame({
        "Показатель": ["Строк", "Колонок", "Всего пропусков"],
        "Значение": [df.shape[0], df.shape[1], int(df.isna().sum().sum())],
    })
    style_table(doc, overview)

    # --- 2. Колонки ---
    doc.add_page_break()
    add_section(doc, "2. Описание колонок", level=1)

    # Переводим типы на русский
    types_ru = [TYPE_RU.get(str(t), str(t)) for t in df.dtypes]

    info = pd.DataFrame({
        "Колонка": [str(c) for c in df.columns],
        "Тип данных": types_ru,
        "Пропуски": df.isna().sum().values,
        "Уникальных": [df[c].nunique() for c in df.columns],
    })
    style_table(doc, info)

    # --- 3. Числовые колонки ---
    num_cols = df.select_dtypes(include="number").columns.tolist()
    if num_cols:
        doc.add_page_break()
        add_section(doc, "3. Числовые показатели", level=1)

        desc = df[num_cols].describe().round(2)
        desc.index = [DESCRIBE_RU.get(i, i) for i in desc.index]
        desc = desc.reset_index().rename(columns={"index": "Показатель"})
        style_table(doc, desc, title="Сводная статистика")

        for col in num_cols[:4]:
            doc.add_paragraph()
            fig, ax = plt.subplots(figsize=(7, 3.5))
            df[col].dropna().plot(kind="hist", bins=20, ax=ax, color="#1A3D7C", edgecolor="white")
            ax.set_title(f"Распределение: {col}", color="#1A3D7C", fontweight="bold")
            ax.set_xlabel(col)
            ax.set_ylabel("Частота")
            ax.grid(axis="y", alpha=0.3)
            ax.spines["top"].set_visible(False)
            ax.spines["right"].set_visible(False)
            doc.add_picture(fig_to_bytes(fig), width=Inches(6))
            add_caption(doc, f"Рис. Распределение значений колонки «{col}»")

    # --- 4. Категориальные колонки ---
    cat_cols = df.select_dtypes(include=["str", "object", "category", "bool"]).columns.tolist()
    if cat_cols:
        doc.add_page_break()
        add_section(doc, "4. Категориальные показатели", level=1)

        for col in cat_cols[:4]:
            vc = df[col].value_counts().head(10)
            if len(vc) <= 1:
                p = doc.add_paragraph()
                run = p.add_run(f"«{col}»: ")
                run.bold = True
                p.add_run(f"единственное значение — {vc.index[0] if len(vc) else '—'}")
                continue

            p = doc.add_paragraph()
            run = p.add_run(f"Колонка «{col}» — топ-10 значений")
            run.bold = True
            run.font.color.rgb = NAVY

            fig, ax = plt.subplots(figsize=(7, 3.5))
            vc.iloc[::-1].plot(kind="barh", ax=ax, color="#1A3D7C")
            ax.set_title(f"Топ-10 значений: {col}", color="#1A3D7C", fontweight="bold")
            ax.set_xlabel("Количество")
            ax.grid(axis="x", alpha=0.3)
            ax.spines["top"].set_visible(False)
            ax.spines["right"].set_visible(False)
            doc.add_picture(fig_to_bytes(fig), width=Inches(6))
            add_caption(doc, f"Рис. Распределение категорий колонки «{col}»")
            doc.add_paragraph()

    doc.save(output)
    print(f"✅ Отчёт сохранён: {output}")


# ---------- main ----------

def main():
    raw = input("Путь к файлу (csv/xlsx): ")
    path = raw.strip().strip('"').strip("'").strip()
    p = Path(path).resolve()
    print(f"\n📄 Файл: {p}")

    if not p.exists():
        print("❌ Файл не найден");
        return
    if not p.is_file():
        print("❌ Это папка");
        return

    ext = p.suffix.lower()
    try:
        if ext == ".csv":
            try:
                df = pd.read_csv(p, sep=None, engine="python")
            except Exception:
                df = pd.read_csv(p, sep=";")
        elif ext in (".xlsx", ".xlsm"):
            df = pd.read_excel(p, engine="openpyxl")
        elif ext == ".xls":
            df = pd.read_excel(p, engine="xlrd")
        else:
            print(f"❌ Расширение {ext!r} не поддерживается");
            return
    except ImportError as e:
        print(f"❌ Не хватает библиотеки: {e}");
        return
    except Exception as e:
        print(f"❌ Ошибка чтения: {e}");
        return

    print(f"✅ Загружено: {df.shape[0]} строк, {df.shape[1]} колонок")
    build_docx_report(df, p.name, p.with_suffix(".docx").name)


if __name__ == "__main__":
    main()