import re
from pathlib import Path

path = Path('./books')
for file in path.glob('*.xlsx'):
    file_name = file.name
    match = re.search(r'20(\d{4})', file_name)
    if match is None:
        continue

    month_folder = path / match[0]

    month_folder.mkdir(exist_ok=True)

    month_file = month_folder / file_name
    file.rename(month_file)
