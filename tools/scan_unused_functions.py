import re
from pathlib import Path

p = Path('/workspaces/KPI-Project/Code.js')
s = p.read_text()

# Find function definitions: function name(...)
defs = re.findall(r"^function\s+([A-Za-z0-9_]+)\s*\(", s, flags=re.M)
# Also capture function expressions: const name = function( and const name = (...)=>
exprs = re.findall(r"^const\s+([A-Za-z0-9_]+)\s*=\s*function\s*\(|^const\s+([A-Za-z0-9_]+)\s*=\s*\([^\)]*\)\s*=>", s, flags=re.M)
exprs = [x[0] or x[1] for x in exprs]
all_defs = sorted(set(defs + exprs))

results = []
for name in all_defs:
    # count occurrences of name( (call sites) and 'name' or "name" (menu/name refs)
    call_count = len(re.findall(r"\b"+re.escape(name)+r"\s*\(", s))
    quoted_count = len(re.findall(r"[\'\"]"+re.escape(name)+r"[\'\"]", s))
    total = call_count + quoted_count
    results.append((name, call_count, quoted_count, total))

unused = [r for r in results if r[3] <= 1]  # only definition likely

print(f'Total functions found: {len(all_defs)}')
print('\nPotentially unused (only defined, not referenced elsewhere):')
for name, calls, quotes, tot in sorted(unused):
    print(f'- {name}: calls={calls}, quoted={quotes}, total_refs={tot}')

print('\nSummary (top 30 functions by reference count):')
for name, calls, quotes, tot in sorted(results, key=lambda x: -x[3])[:30]:
    print(f'{name}: {tot} refs (calls={calls}, quoted={quotes})')
