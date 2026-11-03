import io
import json
import re

path = r'pytools-core/probe/reports/submission_check_python_baseline.json'
BACKSLASH = chr(92)
text = io.open(path, encoding='utf-8').read()

text = re.sub(r'\(?![\/"bfnrtu])', BACKSLASH + BACKSLASH, text)
data = json.loads(text)

def fix(value):
    if isinstance(value, dict):
        return {k: fix(v) for k, v in value.items()}
    if isinstance(value, list):
        return [fix(v) for v in value]
    if isinstance(value, str) and '<ROOT>' in value:
        return value.replace(BACKSLASH, '/')
    return value

data = fix(data)
io.open(path, 'w', encoding='utf-8', newline='').write(
    json.dumps(data, ensure_ascii=False, indent=2) + '\n')
print('baseline repaired and normalized')
