"""Private, offline marketplace-SKU registry. Shopee supplies master descriptions.

Seller aliases are platform scoped. Never guess by scent fragments or equate a
mass with a volume. Source names are evidence/display, not the identity key.
"""
from pathlib import Path
import copy
import json
import re

from extractors import PackingError


def normalized_sku(value):
    return re.sub(r'\s+', '', str(value or '')).upper()


def marketplace(value):
    value = str(value or '').upper()
    return 'SHOPEE' if value in ('SPX', 'FLASH') else value


def _pair(product, variation, platform):
    from names import lookup_key
    return lookup_key(product, variation, platform)


class SkuRegistry:
    def __init__(self, data=None):
        data = data or {'schema': 1, 'records': []}
        if data.get('schema') != 1 or not isinstance(data.get('records'), list):
            raise PackingError('The SKU catalog format is not recognized. Restore its backup.')
        self._entries = copy.deepcopy(data['records'])
        self._sku = {}
        self._pairs = {}
        self._allpairs = {}
        canonical = set()
        for entry in self._entries:
            sku = normalized_sku(entry.get('sku'))
            if not sku or sku in canonical:
                raise PackingError('The SKU catalog contains a blank or repeated master SKU.')
            canonical.add(sku)
            entry['sku'] = sku
            if not entry.get('original_product'):
                raise PackingError(f'SKU {sku} is missing its ecommerce description.')
            for link in entry.get('sku_aliases', []):
                key = marketplace(link.get('platform')), normalized_sku(link.get('sku'))
                if not all(key):
                    raise PackingError(f'SKU {sku} has an incomplete marketplace link.')
                if key in self._sku and self._sku[key]['sku'] != sku:
                    raise PackingError(f'Seller SKU {key[1]} links to more than one product in {key[0]}. Review the SKU catalog.')
                self._sku[key] = entry
            for source in entry.get('ecommerce_sources', []):
                platform = marketplace(source.get('platform'))
                pair = _pair(source.get('product', ''), source.get('variation', ''), platform)
                self._pairs.setdefault((platform, pair), {})[sku] = entry
                self._allpairs.setdefault(pair, {})[sku] = entry

    @property
    def entries(self):
        return copy.deepcopy(self._entries)

    def resolve(self, platform, seller_sku, product='', variation=''):
        platform, seller_sku = marketplace(platform), normalized_sku(seller_sku)
        if seller_sku:
            result = self._sku.get((platform, seller_sku))
            # A supplied but unknown SKU must not silently borrow a different
            # known SKU merely because its current marketplace name is equal.
            return copy.deepcopy(result) if result else {}
        if not product:
            return {}
        pair = _pair(product, variation, platform)
        matches = self._pairs.get((platform, pair), {}) if platform else self._allpairs.get(pair, {})
        if len(matches) == 1:
            return copy.deepcopy(next(iter(matches.values())))
        return {}

    def review_reason(self, platform, seller_sku):
        value = normalized_sku(seller_sku)
        if (marketplace(platform), value) in self._sku:
            return ''
        if not value or value in ('-', '—', 'N/A', 'NONE'):
            return 'No usable seller SKU was provided.'
        if re.fullmatch(r'\d+(?:\.\d+)?(?:ML|L|LITER|LITRE|G|KG|PCS?|PIECES?)', value):
            return 'This seller SKU contains only a size. Different products may reuse it.'
        return ''


def load_registry(catalog_path):
    path = Path(catalog_path).with_name('SkuCatalog.json')
    if not path.exists():
        return None
    try:
        data = json.loads(path.read_text(encoding='utf-8-sig'))
        return SkuRegistry(data)
    except PackingError:
        raise
    except Exception as exc:
        raise PackingError(f'Cannot read {path.name}: {exc}. Restore a valid SKU catalog backup.') from exc
