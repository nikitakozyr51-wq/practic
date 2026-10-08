from pathlib import Path
from zipfile import ZipFile
root=Path(__file__).resolve().parent
for filename, folder in [('Nazar_Otchet_Praktika.docx','nazar'),('ОТЧЕТ.docx','friend')]:
    dest=root/'analysis'/'examples'/folder
    dest.mkdir(parents=True,exist_ok=True)
    with ZipFile(root/'originals'/filename) as archive:
        for name in archive.namelist():
            if name.startswith('word/media/') and not name.endswith('/'):
                target=dest/Path(name).name
                content=archive.read(name)
                if target.exists() and target.read_bytes()!=content:
                    raise RuntimeError('Refusing to overwrite '+str(target))
                target.write_bytes(content)
print('Images restored from original reports.')
