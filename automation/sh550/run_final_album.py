from __future__ import annotations

import fitz

import build_final_album as build


def find_act_pairs_fixed(acts_pdf):
    """Use the verified print-page positions from the inspected facade AOSR workbook."""
    doc = fitz.open(acts_pdf)
    pairs = {1: (26, 27), 2: (22, 23), 3: (18, 19), 4: (14, 15), 5: (10, 11)}
    if doc.page_count < 28:
        count = doc.page_count
        doc.close()
        raise RuntimeError(f"Acts workbook conversion is shorter than the verified 28-page minimum: {count}")
    for act_no, (first, second) in pairs.items():
        if first >= doc.page_count or second >= doc.page_count:
            count = doc.page_count
            doc.close()
            raise RuntimeError(f"Verified pages for Act {act_no}/Facade are outside converted workbook ({count} pages)")
    doc.close()
    return pairs


build.find_act_pairs = find_act_pairs_fixed

if __name__ == "__main__":
    build.main()
