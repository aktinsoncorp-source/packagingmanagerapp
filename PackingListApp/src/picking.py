"""Per-document inventory totals using the source product and variation names."""
from collections import Counter
import re


def sizes(text):
    units = {'ml': 'mL', 'l': 'L', 'liter': 'L', 'liters': 'L', 'kg': 'kg', 'g': 'g'}
    return list(dict.fromkeys(f'{amount} {units[unit.lower()]}' for amount, unit in
                             re.findall(r'(?<![\d.])(\d+(?:\.\d+)?)\s*(m[lL]|[lL]iters?|[lL]|kg|g)\b', text, re.I)))


def picking_rows(batches):
    totals = Counter()
    for batch in batches:
        for order in batch.orders:
            for item in order.items:
                # Match the order tables exactly: no aliases, abbreviation,
                # title-size inference, or merging of different listing names.
                totals[(item.product, item.variation or '—')] += item.pieces
    return [(name, variation, count) for (name, variation), count in
            sorted(totals.items(), key=lambda x: (x[0][0].casefold(), x[0][1].casefold()))]


def is_picking_table(table):
    return [c.text for c in table.rows[0].cells] == ['Done', 'Product', 'Variation / size', 'Total pieces']
