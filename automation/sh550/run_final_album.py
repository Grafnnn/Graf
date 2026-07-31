from __future__ import annotations

import re

import fitz

import build_final_album as build


def find_act_pairs_fixed(acts_pdf):
    doc = fitz.open(acts_pdf)
    fallback = {1: 26, 2: 22, 3: 18, 4: 14, 5: 10}
    result = {}
    for act_no, index in fallback.items():
        if index + 1 >= doc.page_count:
            doc.close()
            raise RuntimeError(f"Acts workbook conversion is shorter than expected: {doc.page_count}")
        text = build.normalize_for_match(doc[index].get_text("text"))
        pattern = rf"(?:^|\s){act_no}\s*/\s*фасад(?:\s|$)"
        if "фасад" not in text or not re.search(pattern, text):
            found = None
            for page_index, page in enumerate(doc):
                candidate = build.normalize_for_match(page.get_text("text"))
                if "фасад" in candidate and re.search(pattern, candidate):
                    found = page_index
                    break
            if found is None:
                doc.close()
                raise RuntimeError(f"Cannot locate Act {act_no}/Фасад in converted workbook")
            index = found
        result[act_no] = (index, index + 1)
    doc.close()
    return result


build.find_act_pairs = find_act_pairs_fixed

if __name__ == "__main__":
    build.main()
