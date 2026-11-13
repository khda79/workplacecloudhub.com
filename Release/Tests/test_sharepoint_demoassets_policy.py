"""Version1.0.1: real ZIP generation with a fake tracked-file inventory; no Git invocation."""
import importlib.util,io,json,pathlib,sys,tempfile,unittest,zipfile
from unittest.mock import patch
ROOT=pathlib.Path(__file__).resolve().parents[2]
spec=importlib.util.spec_from_file_location('release_builder',ROOT/'SmartM365/SharePointMigration/Tools/build_release.py')
builder=importlib.util.module_from_spec(spec);spec.loader.exec_module(builder)
def zipbytes(entries):
    data=io.BytesIO()
    with zipfile.ZipFile(data,'w') as z:
        for name,content in entries.items():z.writestr(name,content)
    return data.getvalue()
class Policy(unittest.TestCase):
    def test_names(self):
        for name in ['DemoAssets/picture.png','Deep/dEmOaSsEtS/image.png','Deep\\DemoAssets\\image.png','Product-DemoAssets.zip']:
            self.assertTrue(builder.is_demoassets_path(name),name)
        self.assertFalse(builder.is_demoassets_path('Release/New-DemoAssets.ps1'))
    def test_nested_archives(self):
        for name in ['DemoAssets/x.png','deep\\DemoAssets\\x.png']:
            with self.assertRaises(ValueError):builder.assert_no_demoassets_archive(zipbytes({'payload.zip':zipbytes({name:'synthetic'})}),'safe-name.zip')
        builder.assert_no_demoassets_archive(zipbytes({'app.ps1':'# runtime','New-DemoAssets.ps1':'# generator'}),'safe.zip')
    def test_corrupt_rejected(self):
        with self.assertRaises(zipfile.BadZipFile):builder.assert_no_demoassets_archive(b'corrupt','broken.zip')
    def test_renamed_nested_zip(self):
        with self.assertRaises(ValueError):builder.assert_no_demoassets_archive(zipbytes({'payload.bin':zipbytes({'DemoAssets/x.png':'synthetic'})}),'safe.zip')
    def test_real_builder_excludes_even_forced_tracked_assets(self):
        with tempfile.TemporaryDirectory() as temp:
            root=pathlib.Path(temp);prefix=builder.PREFIX;notes=prefix+f'RELEASE-NOTES-{builder.VERSION}.md'
            tracked=[notes,prefix+'app.ps1',prefix+'New-DemoAssets.ps1',prefix+'nested/DemoAssets/picture.png',prefix+'Product-DemoAssets.zip']
            for name in tracked+['LICENSE']:
                path=root/name;path.parent.mkdir(parents=True,exist_ok=True);path.write_text('# fixture',encoding='utf-8')
            def fake_git(args,**kwargs):
                if 'ls-files' in args:return '\0'.join(tracked)+'\0'
                if 'status' in args:return ''
                return 'a'*40
            output=root/'output'
            with patch.object(builder,'REPO',root),patch.object(builder,'ROOT_FILES',['LICENSE']),patch.object(builder.subprocess,'check_output',side_effect=fake_git),patch.object(sys,'argv',['build_release','--output-directory',str(output)]):builder.main()
            with zipfile.ZipFile(next(output.glob('*.zip'))) as z:
                self.assertFalse(any(builder.is_demoassets_path(n) for n in z.namelist()))
                self.assertIn(prefix+'New-DemoAssets.ps1',z.namelist())
                self.assertEqual(json.loads(z.read('release-manifest.json'))['excludedDemoAssets'],2)
if __name__=='__main__':unittest.main()
