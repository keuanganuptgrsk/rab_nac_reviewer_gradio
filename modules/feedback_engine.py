from collections import Counter

import pandas as pd

from . import db


def learning_summary():
    feedback = db.get_feedback()
    if not feedback:
        empty = pd.DataFrame(columns=["item", "count"])
        return empty, empty, empty, empty, empty, pd.DataFrame()
    false_pos = Counter()
    false_neg = Counter()
    new_kw = Counter()
    syn = Counter()
    exc = Counter()
    for row in feedback:
        key = row.get("matched_keyword") or (row.get("original_text") or "")[:80]
        ftype = row.get("feedback_type", "")
        if ftype == "Not NAC":
            false_pos[key] += 1
            exc[(row.get("original_text") or "")[:80]] += 1
        elif ftype == "Correct NAC":
            false_neg[key] += 1
        elif ftype == "Add as New NAC Keyword":
            new_kw[(row.get("original_text") or "")[:80]] += 1
        elif ftype == "Add as Synonym":
            syn[key] += 1
    return (
        _df(false_pos),
        _df(false_neg),
        _df(new_kw),
        _df(syn),
        _df(exc),
        pd.DataFrame(feedback),
    )


def _df(counter):
    return pd.DataFrame(counter.most_common(20), columns=["item", "count"])

