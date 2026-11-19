"""Version 1.0.0: release body footer safety net; no GitHub call."""
import importlib.util, pathlib, re, tempfile, unittest

ROOT = pathlib.Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location('release_website_footer', ROOT / 'Release/release_website_footer.py')
footer = importlib.util.module_from_spec(spec); spec.loader.exec_module(footer)
MAP = footer.load_map()


class Footer(unittest.TestCase):
    def test_appends_product_page(self):
        body = footer.with_footer('device-reboot-manager-v1.2.0', '# Notes\r\n\r\n- Fix\r\n\r\n', MAP)
        self.assertEqual(body, '# Notes\n\n- Fix\n\n---\n\nProduct page: [Smart Device Reboot Manager](https://workplacecloudhub.com/tools/smart-device-reboot-manager/) | More tools: [workplacecloudhub.com/tools](https://workplacecloudhub.com/tools/)\n')

    def test_idempotent_when_site_linked(self):
        self.assertIsNone(footer.with_footer('device-reboot-manager-v1.2.0', 'See https://workplacecloudhub.com/tools/', MAP))
        once = footer.with_footer('smartfinops-workplace-v2.0.0', 'x', MAP)
        self.assertIsNone(footer.with_footer('smartfinops-workplace-v2.0.0', once, MAP))

    def test_retired_and_unmapped(self):
        for tag in ['intune-hybrid-join-toolkit-v2.10.81', 'windows11-upgrade-toolkit-v1.0.1', 'smartworkplacecmdb-v1.4.0']:
            self.assertIsNone(footer.with_footer(tag, 'x', MAP), tag)
        self.assertTrue(footer.with_footer('brand-new-tool-v1.0.0', '', MAP).startswith('---\n\nMore tools: '))

    def test_every_product_maps_to_its_slug(self):
        for product in MAP['products']:
            line = footer.footer_line(product['prefix'] + '1.0.0', MAP)
            self.assertIn(f"/tools/{product['slug']}/", line)

    def test_map_matches_readme_table(self):
        readme = (ROOT / 'Release/README.md').read_text(encoding='utf-8')
        table = dict(re.findall(r'^\| `([\w.-]+)` \| `([\w.-]+)`', readme, re.M))
        self.assertEqual(table, {p['prefix']: p['slug'] for p in MAP['products']})

    def test_cli_writes_only_when_needed(self):
        with tempfile.TemporaryDirectory() as temp:
            src, out = pathlib.Path(temp, 'body.md'), pathlib.Path(temp, 'out.md')
            src.write_text('Done https://workplacecloudhub.com/', encoding='utf-8')
            footer.main(['--tag', 'device-registration-tool-v2.0', '--body-file', str(src), '--output', str(out)])
            self.assertFalse(out.exists())
            src.write_text('Done', encoding='utf-8')
            footer.main(['--tag', 'device-registration-tool-v2.0', '--body-file', str(src), '--output', str(out)])
            self.assertIn('/tools/smart-device-registration-tool/', out.read_text(encoding='utf-8'))


if __name__ == '__main__':
    unittest.main()
