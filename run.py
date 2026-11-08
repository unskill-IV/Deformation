from pathlib import Path
import subprocess

measurements = [mea.name for mea in Path("Deformation-data/measurements/").iterdir() if mea.is_dir()]

for mea in measurements:
    subprocess.run(
        ['python3', '-X', 'utf8', 'main.py'],
        input=(mea + '\n'),
        text=True,
        encoding='utf-8'
    )
    print(f'{mea}\nУспешно завершено')
