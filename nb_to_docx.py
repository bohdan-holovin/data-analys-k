"""
Convert Jupyter notebook to .docx with clean formatting.

Preprocessing before pandoc:
1. Remove code cell sources (keep only outputs)
2. Fix Styler tables: use text/html instead of broken text/plain
3. Add white background to PNG images
4. Run pandoc with reference styles
"""

import json
import sys
import os
import base64
import subprocess
import tempfile
from io import BytesIO
from pathlib import Path

try:
    from PIL import Image
except ImportError:
    Image = None
    print("WARNING: Pillow not installed. PNG backgrounds won't be fixed.")
    print("  Install with: pip install Pillow")


def fix_png_background(png_b64: str) -> str:
    """Add white background to a base64-encoded PNG."""
    if Image is None:
        return png_b64
    png_bytes = base64.b64decode(png_b64)
    img = Image.open(BytesIO(png_bytes))
    if img.mode in ('RGBA', 'LA', 'PA'):
        bg = Image.new('RGBA', img.size, (255, 255, 255, 255))
        bg.paste(img, mask=img.split()[-1])
        img = bg.convert('RGB')
    elif img.mode != 'RGB':
        img = img.convert('RGB')
    buf = BytesIO()
    img.save(buf, format='PNG')
    return base64.b64encode(buf.getvalue()).decode('ascii')


def html_table_to_markdown(html: str) -> str:
    """Convert simple HTML table to markdown (best-effort)."""
    import re
    # Extract rows
    rows = re.findall(r'<tr[^>]*>(.*?)</tr>', html, re.DOTALL)
    if not rows:
        return None

    md_rows = []
    for row_html in rows:
        cells = re.findall(r'<t[hd][^>]*>(.*?)</t[hd]>', row_html, re.DOTALL)
        # Strip HTML tags from cell content
        clean = [re.sub(r'<[^>]+>', '', c).strip() for c in cells]
        md_rows.append(clean)

    if not md_rows:
        return None

    # Build markdown table
    n_cols = max(len(r) for r in md_rows)
    # Pad rows
    for r in md_rows:
        while len(r) < n_cols:
            r.append('')

    lines = []
    # Header
    lines.append('| ' + ' | '.join(md_rows[0]) + ' |')
    lines.append('| ' + ' | '.join(['---'] * n_cols) + ' |')
    # Data
    for row in md_rows[1:]:
        lines.append('| ' + ' | '.join(row) + ' |')

    return '\n'.join(lines)


def preprocess_notebook(input_path: str) -> dict:
    """Clean notebook for docx export."""
    with open(input_path, 'r', encoding='utf-8') as f:
        nb = json.load(f)

    new_cells = []
    for cell in nb['cells']:
        if cell['cell_type'] == 'markdown':
            new_cells.append(cell)
            continue

        if cell['cell_type'] != 'code':
            continue

        outputs = cell.get('outputs', [])
        if not outputs:
            continue

        # Create a markdown cell from outputs
        new_outputs = []
        for out in outputs:
            out_type = out.get('output_type', '')

            # Stream output (print statements) -> keep as-is
            if out_type == 'stream':
                new_outputs.append(out)
                continue

            # Display data
            if out_type in ('display_data', 'execute_result'):
                data = out.get('data', {})
                text_plain = ''.join(data.get('text/plain', []))

                # Fix Styler: if text/plain is just a repr, use HTML
                if 'Styler at 0x' in text_plain and 'text/html' in data:
                    html = ''.join(data['text/html'])
                    md_table = html_table_to_markdown(html)
                    if md_table:
                        # Replace with a markdown cell for the table
                        md_cell = {
                            'cell_type': 'markdown',
                            'metadata': {},
                            'source': md_table,
                        }
                        if 'id' not in cell:
                            pass
                        new_cells.append(md_cell)
                        continue
                    else:
                        # Fallback: keep HTML
                        new_outputs.append(out)
                        continue

                # Fix PNG backgrounds
                if 'image/png' in data:
                    png_data = data['image/png']
                    if isinstance(png_data, list):
                        png_data = ''.join(png_data)
                    data['image/png'] = fix_png_background(png_data)

                new_outputs.append(out)
                continue

            new_outputs.append(out)

        if new_outputs:
            # Code cell without source, only outputs
            clean_cell = {
                'cell_type': 'code',
                'execution_count': None,
                'metadata': {},
                'source': '',
                'outputs': new_outputs,
            }
            new_cells.append(clean_cell)

    nb['cells'] = new_cells
    return nb


def main():
    input_file = sys.argv[1] if len(sys.argv) > 1 else 'kanevske_analysis.ipynb'
    output_file = sys.argv[2] if len(sys.argv) > 2 else input_file.replace('.ipynb', '.docx')
    reference_doc = 'reference.docx'

    print(f"Input:  {input_file}")
    print(f"Output: {output_file}")

    # Step 1: Preprocess
    print("Preprocessing notebook...")
    nb = preprocess_notebook(input_file)

    # Save temp notebook
    with tempfile.NamedTemporaryFile(mode='w', suffix='.ipynb', delete=False, encoding='utf-8') as tmp:
        json.dump(nb, tmp, ensure_ascii=False)
        tmp_path = tmp.name

    # Step 2: Run pandoc
    try:
        cmd = ['pandoc', tmp_path, '-o', output_file, '--from=ipynb', '--to=docx']
        if os.path.exists(reference_doc):
            cmd += ['--reference-doc=' + reference_doc]
            print(f"Using style template: {reference_doc}")
        else:
            print("No reference.docx found — using default styles")

        print("Running pandoc...")
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode != 0:
            print(f"Pandoc error: {result.stderr}")
            sys.exit(1)

        size_mb = os.path.getsize(output_file) / (1024 * 1024)
        print(f"Done! {output_file} ({size_mb:.1f} MB)")
    finally:
        os.unlink(tmp_path)


if __name__ == '__main__':
    main()
