#!/usr/bin/env python3
"""
Minimal Compliance Document Generator - NO COMPLEX PARSING
"""

import os
import zipfile
import shutil
from pathlib import Path

# Avoid any complex docx operations that might fail
try:
    from docx import Document
    from docx.shared import Inches
    DOCX_AVAILABLE = True
except Exception as e:
    print(f"WARNING: python-docx not fully available: {e}")
    DOCX_AVAILABLE = False


def extract_zips():
    """Extract all zip files."""
    print("📦 Extracting zip files...")
    for zf in Path(".").glob("*.zip"):
        out = Path(f"extracted_{zf.stem}")
        if out.exists():
            shutil.rmtree(out)
        try:
            with zipfile.ZipFile(zf) as z:
                z.extractall(out)
            print(f"  ✅ {zf.name}")
        except Exception as e:
            print(f"  ❌ {zf.name}: {e}")


def get_compliances():
    """Extract compliance names from directories."""
    print("\n📋 Finding compliances...")
    names = set()
    
    for extracted in Path(".").glob("extracted_*"):
        for item in extracted.iterdir():
            if item.is_dir() and not item.name.startswith("."):
                names.add(item.name)
    
    names = sorted([n for n in names if n and len(n) > 2])
    print(f"✅ Found {len(names)} compliances")
    
    return names


def create_docs(compliances):
    """Create compliance documents."""
    if not DOCX_AVAILABLE:
        print("❌ python-docx not available")
        return []
    
    print(f"\n📝 Creating {len(compliances)} documents...")
    
    template = Path("ANVISA for Log360.docx")
    if not template.exists():
        print("❌ Template not found")
        return []
    
    out_dir = Path("generated_documents")
    out_dir.mkdir(exist_ok=True)
    created = []
    
    for i, name in enumerate(compliances, 1):
        print(f"[{i}/{len(compliances)}] {name}...", end=" ")
        try:
            doc = Document(str(template))
            
            # Simple text replacement
            for para in doc.paragraphs:
                for run in para.runs:
                    run.text = run.text.replace("Log360", name)
            
            # Safe filename
            safe = "".join(c if c.isalnum() or c in " -_" else "_" for c in name)
            out_file = out_dir / f"{safe}.docx"
            doc.save(str(out_file))
            created.append(str(out_file))
            print("✅")
        except Exception as e:
            print(f"❌ {e}")
    
    return created


def main():
    print("🚀 Compliance Generator Started\n")
    
    extract_zips()
    compliances = get_compliances()
    
    if not compliances:
        print("❌ No compliances found")
        return
    
    created = create_docs(compliances)
    
    # Report
    report = Path("COMPLIANCE_GENERATION_REPORT.md")
    with open(report, "w") as f:
        f.write(f"# Report\n\n**Created: {len(created)} documents**\n\n")
        for doc in sorted(created):
            f.write(f"- {Path(doc).name}\n")
    
    print(f"\n✨ Complete! {len(created)}/{len(compliances)} documents created")


if __name__ == "__main__":
    main()
