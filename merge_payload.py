"""Merge dashboard_template2.html + dashboard_payload.json -> final HTML ready for Artifact publish.

Usage: python3 merge_payload.py <template.html> <payload.json> <output.html>
"""
import sys, json

def run(template_path, payload_path, out_path):
    with open(template_path, 'r', encoding='utf-8') as f:
        tpl = f.read()
    with open(payload_path, 'r', encoding='utf-8') as f:
        payload_raw = f.read()
    # payload_raw is already valid JSON text; just splice it in verbatim.
    json.loads(payload_raw)  # sanity check it actually parses
    if '__DATA_JSON__' not in tpl:
        raise ValueError("Template is missing the __DATA_JSON__ placeholder")
    merged = tpl.replace('__DATA_JSON__', payload_raw, 1)
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(merged)
    print(f"Wrote {out_path} ({len(merged)} bytes)")

if __name__ == '__main__':
    if len(sys.argv) != 4:
        print("Usage: python3 merge_payload.py <template.html> <payload.json> <output.html>", file=sys.stderr)
        sys.exit(1)
    run(sys.argv[1], sys.argv[2], sys.argv[3])
