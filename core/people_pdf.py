"""'Space Input Data' reports (HAP) -> number of people per space.

Every space starts with a heading line such as '#001-BD1-B-Pump Room'; below it
'1. General Details: Floor Area ... 11.2 m2' and '2.4. People: Occupancy ... 1.0 Person'.
Text comes from pdfplumber / PyMuPDF (CLI) or pdf.js (browser) as one string per page.
"""
import re
from dataclasses import dataclass

_HEAD = re.compile(r"^\s*#\s*(\S.*?)\s*$")
_AREA = re.compile(r"Floor\s+Area[\s.]*([0-9]+(?:\.[0-9]+)?)\s*m", re.I)
_OCC = re.compile(r"Occupancy[\s.]*([0-9]+(?:\.[0-9]+)?)\s*Person", re.I)


@dataclass
class PeopleSpace:
    name: str
    people: float
    area: float | None = None
    file: str = ""
    page: int = 0
    source: str = "HAP"


def parse_pages(pages, file_label=""):
    out, cur = [], None

    def close():
        if cur and cur["people"] is not None:
            out.append(PeopleSpace(cur["name"], cur["people"], cur["area"], file_label, cur["page"]))

    for pno, page in enumerate(pages, 1):
        for line in (page or "").splitlines():
            m = _HEAD.match(line)
            if m:
                if cur and cur["name"] == m.group(1):
                    continue                      # same space continues on the next page
                close()
                cur = {"name": m.group(1), "people": None, "area": None, "page": pno}
                continue
            if cur is None:
                continue
            if cur["area"] is None:
                a = _AREA.search(line)
                if a:
                    cur["area"] = float(a.group(1))
            if cur["people"] is None:
                o = _OCC.search(line)
                if o:
                    cur["people"] = float(o.group(1))
    close()
    if not out:
        raise ValueError("No spaces with an 'Occupancy ... Person' value were found in the people PDF. "
                         "Expected a HAP 'Space Input Data' report with selectable text.")
    return out


def pages_from_pdf(path):
    try:
        import fitz                                # PyMuPDF: fast for big reports
        with fitz.open(path) as doc:
            return [p.get_text() for p in doc]
    except ImportError:
        import pdfplumber
        with pdfplumber.open(path) as pdf:
            return [p.extract_text() or "" for p in pdf.pages]


def parse(path):
    import os
    return parse_pages(pages_from_pdf(path), os.path.basename(str(path)))
