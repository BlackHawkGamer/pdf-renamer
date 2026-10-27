from pathlib import Path
from datetime import datetime

import pandas as pd
from rapidfuzz import process, fuzz

# ======================================================
# CONFIG
# ======================================================

EXCEL_FILE = Path(__file__).parent / "Zortea_map_efetivo.xlsx"

PDF_FOLDER = Path(__file__).parent

FUZZY_THRESHOLD = 25

# ======================================================
# HELPERS
# ======================================================

def normalize_text(text):

    text = str(text).upper().strip()

    replacements = {
        "Á": "A",
        "À": "A",
        "Ã": "A",
        "Â": "A",
        "É": "E",
        "Ê": "E",
        "Í": "I",
        "Ó": "O",
        "Õ": "O",
        "Ô": "O",
        "Ú": "U",
        "Ç": "C",
    }

    for old, new in replacements.items():
        text = text.replace(old, new)

    return " ".join(text.split())


def simplify_name(text):

    text = normalize_text(text)

    ignored_words = {
        "DA",
        "DE",
        "DO",
        "DAS",
        "DOS"
    }

    words = [
        word
        for word in text.split()
        if word not in ignored_words
    ]

    return " ".join(words)


def extract_name_from_filename(filename):

    name = Path(filename).stem

    name = name.replace("ZORTEA - ", "")
    name = name.replace(" - ASO", "")

    return name.strip()


def find_best_match(name, names_list, threshold=70):

    result = process.extractOne(
        name,
        names_list,
        scorer=fuzz.token_sort_ratio
    )

    if result is None:
        return None, 0

    match_name, score, _ = result

    if score >= threshold:
        return match_name, score

    return None, score


# ======================================================
# LOAD EXCEL
# ======================================================

print("Loading Excel...")
print("Excel path:", EXCEL_FILE)
print("Exists:", Path(EXCEL_FILE).exists())

xls = pd.ExcelFile(EXCEL_FILE)

sheet_name = None

for sheet in xls.sheet_names:

    if "fluig" in sheet.lower():
        sheet_name = sheet
        break

if sheet_name is None:
    raise Exception(
        f"Could not find Fluig worksheet.\n"
        f"Available sheets: {xls.sheet_names}"
    )

print("Using worksheet:", sheet_name)

df = pd.read_excel(
    EXCEL_FILE,
    sheet_name=sheet_name,
    header=2
)

df.columns = [str(col).strip() for col in df.columns]

required_columns = [
    "Name",
    "File_name",
    "Short_position"
]

for col in required_columns:

    if col not in df.columns:
        raise Exception(
            f"Column '{col}' not found.\n"
            f"Columns found:\n{list(df.columns)}"
        )

df = df[df["Name"].notna()].copy()

df["NAME_NORM"] = (
    df["Name"]
    .astype(str)
    .apply(simplify_name)
)

lookup = {}

for _, row in df.iterrows():

    employee_name = row["NAME_NORM"]

    lookup[employee_name] = {
        "file_name": str(
            row.get("File_name", "")
        ).strip(),

        "short_position": str(
            row.get("Short_position", "")
        ).strip()
    }

employee_names = list(lookup.keys())

print(f"Loaded {len(employee_names)} employees.")

# ======================================================
# PROCESS PDFS
# ======================================================

print("\nStarting...")
print("PDF folder:", PDF_FOLDER)

pdfs = list(PDF_FOLDER.glob("*.pdf"))

print("PDFs found:", len(pdfs))

for pdf in pdfs:
    print("-", pdf.name)

renamed = 0
corporate_names = 0
position_fallbacks = 0
sem_cargo_fallbacks = 0

for pdf_file in pdfs:

    try:

        employee_name = extract_name_from_filename(
            pdf_file.name
        )

        employee_norm = simplify_name(
            employee_name
        )

        best_match, score = find_best_match(
            employee_norm,
            employee_names,
            threshold=FUZZY_THRESHOLD
        )

        yy_mm = datetime.now().strftime("%y-%m")

        if best_match:

            employee_data = lookup[best_match]

            corporate_filename = (
                employee_data["file_name"]
            )

            short_position = (
                employee_data["short_position"]
            )

            if (
                corporate_filename
                and corporate_filename.lower() != "nan"
            ):

                new_name = (
                    f"{corporate_filename}"
                    f"_ASO_{yy_mm}.pdf"
                )

                corporate_names += 1

            elif (
                short_position
                and short_position.lower() != "nan"
            ):

                new_name = (
                    f"{employee_name}"
                    f"_{short_position}"
                    f"_ASO_{yy_mm}.pdf"
                )

                position_fallbacks += 1

            else:

                new_name = (
                    f"{employee_name}"
                    f"_SEM_CARGO"
                    f"_ASO_{yy_mm}.pdf"
                )

                sem_cargo_fallbacks += 1

            print("\n--------------------------------")
            print("MATCH FOUND")
            print("PDF Name   :", employee_name)
            print("Excel Name :", best_match)
            print("Score      :", round(score, 1))
            print("New Name   :", new_name)

        else:

            new_name = (
                f"{employee_name}"
                f"_SEM_CARGO"
                f"_ASO_{yy_mm}.pdf"
            )

            sem_cargo_fallbacks += 1

            print("\n--------------------------------")
            print("NO MATCH FOUND")
            print("PDF Name        :", employee_name)
            print("Best Score      :", round(score, 1))
            print("Suggested Match :", best_match)
            print("New Name        :", new_name)

        destination = pdf_file.with_name(
            new_name
        )

        if destination.exists():

            print(
                f"Skipped (already exists): "
                f"{new_name}"
            )

            continue

        pdf_file.rename(destination)

        renamed += 1

    except Exception as e:

        print(f"\nError in {pdf_file.name}")
        print(e)

# ======================================================
# SUMMARY
# ======================================================

print("\n===================================")
print("PROCESS FINISHED")
print("===================================")
print("Renamed              :", renamed)
print("Corporate names used :", corporate_names)
print("Position fallbacks   :", position_fallbacks)
print("SEM_CARGO fallbacks  :", sem_cargo_fallbacks)
print("===================================")
