"""Non-blocking reminders for missing product identities and inventory names.

These reports describe the saved definitions and current PDF evidence only.
They never change identity matching, printed fallback fields, or source data.
"""
from dataclasses import replace
import copy
import re

from extractors import PackingError, normalize_sku
import names


REASONS = ('missing_sku', 'missing_product_name', 'missing_size')


def row_attention_reasons(row):
    """Return deficiencies without treating verified numeric aliases as blank."""
    sku = normalize_sku(row.get('sku', ''))
    # Encountered numeric PDF codes retain their platform provenance in aliases.
    # A bare hand-entered number remains unusable, as in names.validate_rows.
    numeric_source = bool(re.fullmatch(r'\d+', sku) and row.get('sku_aliases'))
    reasons = [] if names.usable_sku(sku) or numeric_source else ['missing_sku']
    if not str(row.get('alternative_product') or '').strip():
        reasons.append('missing_product_name')
    if not str(row.get('alternative_variation') or '').strip():
        reasons.append('missing_size')
    return reasons


def _matched_row(rows, item, platform, registry=None, mapping=None):
    """Use the same SKU/legacy fallback precedence as names.apply_names."""
    identity, canonical = names._identity(platform, item, registry)
    matches = names._sku_matches(rows, identity, platform, canonical)
    row = names._unique_sku_definition(matches, identity.get('sku'))
    if not row:
        mapping = names.definition_mapping(rows, platform) if mapping is None else mapping
        legacy = mapping.get(names.lookup_key(item.product, item.variation, platform), [])
        if identity:
            legacy = [candidate for candidate in legacy if not candidate.get('sku')]
        row = names.matched_definition(legacy, item.product, item.variation)
    return row


def source_attention(catalog, batches, registry=None):
    """Find missing seller SKUs even when a name resolves to a saved master SKU."""
    rows = catalog.get('rows', [])
    issues = {}
    for batch in batches:
        mapping = names.definition_mapping(rows, batch.platform)
        for order in batch.orders:
            for item in order.items:
                # Do not let a registry name match disguise an absent PDF SKU.
                identity, _ = names._identity(batch.platform, replace(item, product='', variation=''), registry)
                if identity:
                    continue
                try:
                    row = _matched_row(rows, item, batch.platform, registry, mapping)
                except PackingError:
                    # An advisory must not choose among conflicting definitions.
                    # Existing scan/generation validation handles those conflicts.
                    continue
                if not row:
                    continue
                entry = issues.setdefault(row['id'], {
                    'id': row['id'], 'sku': row.get('sku', ''),
                    'sku_scope': row.get('sku_scope', ''),
                    'original_product': row['original_product'],
                    'original_variation': row['original_variation'],
                    'reasons': ['missing_sku'], 'sources': []})
                source = {'path': str(batch.path), 'platform': batch.platform,
                          'sku': normalize_sku(item.seller_sku), 'product': item.product,
                          'variation': item.variation, 'order_id': order.order_id}
                if source not in entry['sources']:
                    entry['sources'].append(source)
    return list(issues.values())


def source_definition_ids(catalog, batches, registry=None):
    """Saved definitions used by this scan, including already-complete rows."""
    rows, identities = catalog.get('rows', []), set()
    for batch in batches:
        mapping = names.definition_mapping(rows, batch.platform)
        for order in batch.orders:
            for item in order.items:
                try:
                    row = _matched_row(rows, item, batch.platform, registry, mapping)
                except PackingError:
                    continue
                if row:
                    identities.add(row['id'])
    return [row['id'] for row in rows if row['id'] in identities]


def _current_issue_row(issue, rows, by_id):
    row = by_id.get(issue.get('id'))
    if row is not None:
        return row
    # A confirmed merge may keep another base ID. Follow a unique retained SKU,
    # then linked original pairs; never carry a deleted row's warning arbitrarily.
    sku = normalize_sku(issue.get('sku', ''))
    if sku:
        scope = issue.get('sku_scope', '')
        matches = [row for row in rows if
                   (scope in ('', 'CATALOG') and row.get('sku_scope', '') in ('', 'CATALOG') and
                    sku in names._canonical_skus(row)) or
                   any(names.sku_platform(alias['platform']) == names.sku_platform(scope) and
                       alias['sku'] == sku for alias in row.get('sku_aliases', []))]
        if len(matches) == 1:
            return matches[0]
        return None
    pair = names.lookup_key(issue.get('original_product', ''), issue.get('original_variation', ''))
    matches = names.definition_mapping(rows).get(pair, [])
    return matches[0] if len(matches) == 1 else None


def remap_attention_ids(previous_catalog, current_catalog, ids):
    """Follow confirmed merges while dropping deleted or ambiguous definitions."""
    previous = {row['id']: row for row in previous_catalog.get('rows', [])}
    rows = current_catalog.get('rows', [])
    by_id = {row['id']: row for row in rows}
    result = set()
    for identity in ids:
        if identity not in previous:
            if identity in by_id:
                result.add(identity)
            continue
        row = _current_issue_row(previous[identity], rows, by_id)
        if row is not None:
            result.add(row['id'])
    return result


def catalog_attention(catalog, source_issues=None):
    """Combine saved deficiencies with current-source reminders by definition ID."""
    rows = catalog.get('rows', [])
    by_id = {row['id']: row for row in rows}
    extras = {}
    for issue in source_issues or []:
        row = _current_issue_row(issue, rows, by_id)
        if row is not None:
            extras.setdefault(row['id'], []).append(issue)
    items = []
    for row in rows:
        reasons = set(row_attention_reasons(row))
        sources = copy.deepcopy(row.get('ecommerce_sources', []))
        for issue in extras.get(row['id'], []):
            # Source reports describe missing seller SKUs. Name/size deficiencies
            # must always be recomputed from the current edits or saved rows.
            if 'missing_sku' in issue.get('reasons', []):
                reasons.add('missing_sku')
            for source in issue.get('sources', []):
                if source not in sources:
                    sources.append(copy.deepcopy(source))
        if reasons:
            items.append({'id': row['id'], 'sku': row.get('sku', ''),
                          'original_product': row['original_product'],
                          'original_variation': row['original_variation'],
                          'reasons': [reason for reason in REASONS if reason in reasons],
                          'sources': sources})
    return {'items': items, 'count': len(items),
            'missing_sku_count': sum('missing_sku' in item['reasons'] for item in items),
            'missing_name_count': sum(any(reason in item['reasons'] for reason in REASONS[1:]) for item in items)}
