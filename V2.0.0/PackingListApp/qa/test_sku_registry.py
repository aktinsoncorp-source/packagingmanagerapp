"""Identity joins must survive names changing without conflating reused codes."""
from pathlib import Path
import sys
import unittest
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from sku_registry import SkuRegistry
from extractors import PackingError


class RegistryTests(unittest.TestCase):
    def record(self, sku='OIL-LEMON-100ML', name='Shopee title', variation='Lemon,100mL'):
        return {'sku': sku, 'original_product': name, 'original_variation': variation,
                'sku_aliases': [{'platform':'SHOPEE','sku':sku}, {'platform':'LAZADA','sku':'OLD-123'}],
                'ecommerce_sources':[{'platform':'SHOPEE','sku':sku,'product':name,'variation':variation},
                                     {'platform':'LAZADA','sku':'OLD-123','product':'Old listing','variation':'Scent:Lemon,Size:100 mL'}]}

    def test_renamed_titles_and_wrapped_skus_keep_identity(self):
        r=SkuRegistry({'schema':1,'records':[self.record()]})
        self.assertEqual(r.resolve('SPX','OIL-LE MON-100ML','Changed title','New label')['sku'],'OIL-LEMON-100ML')
        self.assertEqual(r.resolve('LAZADA','OLD-123','Different title','Different variation')['sku'],'OIL-LEMON-100ML')

    def test_same_seller_code_on_other_platform_does_not_match(self):
        r=SkuRegistry({'schema':1,'records':[self.record()]})
        self.assertEqual(r.resolve('TIKTOK','OLD-123','Old listing','Lemon,100mL'),{})
        self.assertTrue(r.review_reason('TIKTOK','30 ml'))

    def test_legacy_name_fallback_only_when_sku_absent(self):
        r=SkuRegistry({'schema':1,'records':[self.record()]})
        self.assertEqual(r.resolve('SHOPEE','','Shopee title','Lemon,100mL')['sku'],'OIL-LEMON-100ML')
        self.assertEqual(r.resolve('SHOPEE','OTHER-SKU','Shopee title','Lemon,100mL'),{})

    def test_ambiguous_name_never_selects_arbitrary_sku(self):
        a=self.record(); b=self.record('OIL-ROSE-100ML');b['sku_aliases']=[{'platform':'SHOPEE','sku':b['sku']}];b['ecommerce_sources']=b['ecommerce_sources'][:1]
        r=SkuRegistry({'schema':1,'records':[a,b]})
        self.assertEqual(r.resolve('SHOPEE','','Shopee title','Lemon,100mL'),{})

    def test_conflicting_aliases_are_rejected(self):
        with self.assertRaises(PackingError):SkuRegistry({'schema':1,'records':[self.record(),self.record('OTHER')]})

    def test_returned_metadata_cannot_mutate_registry(self):
        r=SkuRegistry({'schema':1,'records':[self.record()]});r.entries[0]['sku_aliases'].clear()
        result=r.resolve('SHOPEE','OIL-LEMON-100ML');result['sku_aliases'].clear()
        self.assertEqual(len(r.resolve('SHOPEE','OIL-LEMON-100ML')['sku_aliases']),2)

if __name__=='__main__':unittest.main()
