import re
from pptx import Presentation

from survey_parser.score_rules import METRIC_PATTERNS


def clean_text(text):
    text = text.replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def extract_text_blocks(ppt_path):
    prs = Presentation(ppt_path)
    blocks = []

    for slide_idx, slide in enumerate(prs.slides, start=1):
        for shape in slide.shapes:
            # table
            if getattr(shape, "has_table", False):
                for row in shape.table.rows:
                    row_texts = []
                    for cell in row.cells:
                        t = clean_text(cell.text)
                        if t:
                            row_texts.append(t)
                            blocks.append({"slide": slide_idx, "text": t})
                    if row_texts:
                        blocks.append({"slide": slide_idx, "text": " | ".join(row_texts)})

            # text frame
            if getattr(shape, "has_text_frame", False):
                paras = []
                for p in shape.text_frame.paragraphs:
                    t = clean_text(p.text)
                    if t:
                        paras.append(t)
                        blocks.append({"slide": slide_idx, "text": t})
                if paras:
                    blocks.append({"slide": slide_idx, "text": " | ".join(paras)})

            # fallback for shape.text
            elif hasattr(shape, "text"):
                t = clean_text(shape.text)
                if t:
                    blocks.append({"slide": slide_idx, "text": t})

    return blocks


def _numbers_50_to_100(text):
    nums = [int(x) for x in re.findall(r"(?<!\d)(\d{2,3})(?!\d)", text)]
    return [n for n in nums if 50 <= n <= 100]


def _extract_engagement_score(blocks):
    for b in blocks:
        t = b["text"].lower()
        if "engagement historical trend" in t or "engagement" in t:
            nums = _numbers_50_to_100(t)
            nums = [n for n in nums if n != 100]
            if nums:
                return nums[0]
    return None


def _extract_company_score(blocks):
    for b in blocks:
        t = b["text"].lower()
        m = re.search(r"company[^0-9]{0,12}(\d{2,3})", t)
        if m:
            value = int(m.group(1))
            if 50 <= value <= 100:
                return value
    return None


def _extract_respondents(all_text):
    m = re.search(r"(\d+)\s*/\s*(\d+)\s*respondents", all_text, re.I)
    if m:
        return int(m.group(1)), int(m.group(2))
    return None, None


def _extract_response_rate(all_text):
    m = re.search(r"(\d+)\s*\(\s*(\d+)\s*%\s*\)\s*responded", all_text, re.I)
    if m:
        return int(m.group(2))
    return None


def _extract_manager(all_text):
    m = re.search(r"manager[:\s]+([^\n|]+)", all_text, re.I)
    if m:
        return clean_text(m.group(1))
    return None


def _extract_date(all_text):
    m = re.search(
        r"((?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|January|February|March|April|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4})",
        all_text,
        re.I,
    )
    if m:
        return clean_text(m.group(1))
    return None


def _extract_score_near_alias(text, alias):
    p1 = re.search(rf"{re.escape(alias)}[^0-9]{{0,25}}(\d{{2,3}})", text)
    if p1:
        v = int(p1.group(1))
        if 50 <= v <= 100:
            return v

    p2 = re.search(rf"(\d{{2,3}})[^0-9]{{0,25}}{re.escape(alias)}", text)
    if p2:
        v = int(p2.group(1))
        if 50 <= v <= 100:
            return v

    nums = _numbers_50_to_100(text)
    if nums:
        # usually real favorability score is larger than delta number
        return max(nums)

    return None


def extract_metric_scores(blocks):
    metrics = []

    for block in blocks:
        text_original = block["text"]
        text = text_original.lower()

        for canonical, aliases in METRIC_PATTERNS.items():
            if any(alias in text for alias in aliases):
                best_score = None
                for alias in aliases:
                    if alias in text:
                        s = _extract_score_near_alias(text, alias)
                        if s is not None:
                            if best_score is None or s > best_score:
                                best_score = s
                if best_score is not None:
                    metrics.append({"name": canonical, "score": best_score})

    # deduplicate, keep highest
    dedup = {}
    for m in metrics:
        if m["name"] not in dedup or m["score"] > dedup[m["name"]]["score"]:
            dedup[m["name"]] = m

    return sorted(dedup.values(), key=lambda x: x["score"], reverse=True)


def extract_overview(blocks):
    all_text = "\n".join([b["text"] for b in blocks])

    respondents, total_respondents = _extract_respondents(all_text)
    response_rate = _extract_response_rate(all_text)

    if response_rate is None and respondents and total_respondents:
        response_rate = round((respondents / total_respondents) * 100)

    overview = {
        "survey_date": _extract_date(all_text),
        "manager_or_team": _extract_manager(all_text),
        "engagement_score": _extract_engagement_score(blocks),
        "company_score": _extract_company_score(blocks),
        "respondents": respondents,
        "total_respondents": total_respondents,
        "response_rate": response_rate,
    }
    return overview


def parse_survey_ppt(ppt_path):
    blocks = extract_text_blocks(ppt_path)
    overview = extract_overview(blocks)
    metrics = extract_metric_scores(blocks)

    return {
        "overview": overview,
        "metrics": metrics,
        "text_blocks": blocks,
    }
