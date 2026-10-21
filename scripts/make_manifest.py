"""Hash the publishable files selected by Git's ignore rules; omit circular reports."""
from pathlib import Path
import subprocess,hashlib,json
ROOT=Path(__file__).resolve().parents[1]
SKIP={'FILE_MANIFEST.json','validation/publication-validation.json','validation/publication-controls.csv'}
paths=subprocess.check_output(['git','ls-files','--cached','--others','--exclude-standard','-z'],cwd=ROOT).decode().split('\0')
files=[]
for name in sorted(set(paths)-SKIP-{''}):
 p=ROOT/name
 if p.is_file():files.append({'path':name,'bytes':p.stat().st_size,'sha256':hashlib.sha256(p.read_bytes()).hexdigest()})
(ROOT/'FILE_MANIFEST.json').write_text(json.dumps({'algorithm':'SHA-256','excluded_circular_reports':sorted(SKIP),'files':files},indent=2)+'\n')
print('Manifest:',len(files),'files;',sum(x['bytes'] for x in files),'bytes')
