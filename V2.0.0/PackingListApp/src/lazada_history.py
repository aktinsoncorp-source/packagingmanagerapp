"""Verified tracking identities layered over immutable pre-tracking manifests."""
from collections import defaultdict


def legacy_key(key):
    return key.split(':TRACKING:', 1)[0] if key.startswith('LAZADA:') else key


def is_legacy(key):
    return key.startswith('LAZADA:') and ':TRACKING:' not in key


def output_key(output, key):
    return output.get('tracking_order_keys', {}).get(key, key)


def output_keys(output):
    return {output_key(output, key) for key in output.get('orders', [])}


def upgrade_tracking_links(history, batches):
    """Bind only one proven package from the exact previously recorded PDF bytes.

    Original keys, fingerprints, source manifests and both snapshot kinds stay
    unchanged. The additive map records what identity the old manifest covers.
    """
    changed = False
    for output in history['outputs'].values():
        if output.get('status') == 'superseded':
            continue
        for old_key, fingerprint in output.get('order_fingerprints', {}).items():
            if not is_legacy(old_key) or old_key in output.get('superseded_orders', []):
                continue
            if old_key in output.get('tracking_order_keys', {}):
                continue
            candidates = {}
            sources = [source for source in output.get('sources', []) if old_key in source.get('order_keys', [])]
            complete = bool(sources)
            for source in sources:
                batch = batches.get(source['path'])
                if batch is None or not source.get('sha256') or source['sha256'] != batch.sha256:
                    complete = False
                    break
                found = False
                for order in batch.orders:
                    if legacy_key(order.key) == old_key:
                        found = True
                        complete = complete and order.fingerprint() == fingerprint
                        candidates[order.key] = order
                complete = complete and found
            if complete and len(candidates) == 1:
                key, order = next(iter(candidates.items()))
                if key != old_key and order.fingerprint() == fingerprint:
                    output.setdefault('tracking_order_keys', {})[old_key] = key
                    changed = True
    return changed


def unresolved_keys(history):
    """Legacy coverage that cannot safely be assigned to any current package."""
    return {key for output in history['outputs'].values()
            if output.get('status') != 'superseded'
            for key in output.get('order_fingerprints', {})
            if is_legacy(key) and key not in output.get('superseded_orders', [])
            and output_key(output, key) == key}


def prior_orders(history):
    result = dict(history['orders'])
    for name, output in history['outputs'].items():
        if output.get('status') == 'superseded':
            continue
        for old_key, key in output.get('tracking_order_keys', {}).items():
            if old_key in output.get('superseded_orders', []) or key in result:
                continue
            fingerprint = output.get('order_fingerprints', {}).get(old_key)
            snapshot = output.get('order_snapshots', {}).get(old_key)
            saved = history['orders'].get(old_key, {})
            if snapshot is None and saved.get('fingerprint') == fingerprint:
                snapshot = saved.get('order')
            if snapshot is not None:
                result[key] = {'fingerprint': fingerprint, 'order': snapshot,
                               'output': name, 'source': saved.get('source', '')}
    return result


def source_keys(history, path):
    aliases = defaultdict(set)
    for output in history['outputs'].values():
        for source in output.get('sources', []):
            if source['path'] == path:
                for key in source.get('order_keys', []):
                    aliases[key].add(output_key(output, key))
    result = set()
    for key in history['sources'].get(path, {}).get('order_keys', []):
        mapped = aliases.get(key, set()) - {key}
        result.update(mapped or {key})
    return result


def key_label(key):
    value = key.split(':', 1)[-1]
    return value.replace(':TRACKING:', ' | Tracking ID: ')
