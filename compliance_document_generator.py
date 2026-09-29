#!/usr/bin/env python3
"""
Compliance Document Generator Agent - Simplified Version

This script automates the creation of compliance documents for all compliances.
It:
1. Extracts compliance names from zip files
2. Creates documents based on the ANVISA template
3. Adds matching logos/screenshots
"""

import os
import zipfile
import shutil
from pathlib import Path
from docx import Document
from docx.shared import Inches


def extract_zip_files():
    """Extract all zip files in the repository."""
    print("📦 Extracting zip files...")
    zip_files = list(Path(".").glob("*.zip"))
    extracted_data = []
    
    for zip_path in zip_files:
        extract_dir = Path(f"extracted_{zip_path.stem}")
        if extract_dir.exists():
            shutil.rmtree(extract_dir)
        extract_dir.mkdir(exist_ok=True)
        
        try:
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(extract_dir)
            print(f"  ✅ Extracted {zip_path.name}")
        except Exception as e:
            print(f"  ❌ Failed to extract {zip_path.name}: {e}")
            continue
        
        # Collect all files
        for item in extract_dir.rglob("*"):
            if item.is_file():
                extracted_data.append({
                    "source_zip": str(zip_path),
                    "path": str(item),
                    "name": item.name,
                    "suffix": item.suffix.lower()
                })
    
    print(f"✅ Extracted {len(list(Path('.').glob('extracted_*')))} zip files")
    return extracted_data


def extract_compliance_names(extracted_data):
    """Extract compliance names from extracted files."""
    print("\n📋 Extracting compliance names...")
    compliances = set()
    
    for item in extracted_data:
        # Skip system files
        if item["name"].startswith(".") or item["name"].startswith("_"):
            continue
            
        # Check folder names (directories have empty suffix)
        path_parts = item["path"].split(os.sep)
        if len(path_parts) > 2:
            # Look for compliance name in path structure
            for part in path_parts[1:]:
                if part and not part.startswith(".") and not part.startswith("_"):
                    if "extracted" not in part.lower() and "macosx" not in part.lower():
                        compliances.add(part)
    
    compliances = sorted([c for c in compliances if c and len(c) > 2])
    
    print(f"✅ Found {len(compliances)} compliances")
    for i, comp in enumerate(compliances[:10]):
        print(f"   {i+1}. {comp}")
    if len(compliances) > 10:
        print(f"   ... and {len(compliances)-10} more")
    
    return compliances


def find_assets_for_compliance(compliance_name, extracted_data):
    """Find logo and screenshot files for a compliance."""
    logo = None
    screenshot = None
    
    search_pattern = compliance_name.lower().replace(" ", "_").replace("-", "_")
    
    for item in extracted_data:
        item_lower = item["path"].lower()
        
        # Skip system files
        if "/__macosx/" in item_lower or item["name"].startswith("._"):
            continue
        
        # Look for images in compliance directory
        if compliance_name.lower() in item_lower and item["suffix"] in [".png", ".jpg", ".jpeg", ".gif"]:
            if not logo:
                logo = item["path"]
            elif not screenshot:
                screenshot = item["path"]
                break
    
    return logo, screenshot


def create_compliance_document(compliance_name, logo_path, screenshot_path, output_dir):
    """Create a new compliance document based on the ANVISA template."""
    template_path = Path("ANVISA for Log360.docx")
    
    if not template_path.exists():
        print(f"  ❌ Template not found: {template_path}")
        return None
    
    try:
        # Load and modify template
        doc = Document(str(template_path))
        
        # Replace Log360 with compliance name
        for para in doc.paragraphs:
            for run in para.runs:
                if "Log360" in run.text:
                    run.text = run.text.replace("Log360", compliance_name)
        
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if "Log360" in cell.text:
                        cell.text = cell.text.replace("Log360", compliance_name)
        
        # Add images
        if logo_path and Path(logo_path).exists():
            try:
                para = doc.add_paragraph()
                para.add_run().add_picture(str(logo_path), width=Inches(2.0))
            except Exception as e:
                print(f"  ⚠️ Could not add logo: {e}")
        
        if screenshot_path and Path(screenshot_path).exists():
            try:
                para = doc.add_paragraph()
                para.add_run().add_picture(str(screenshot_path), width=Inches(4.0))
            except Exception as e:
                print(f"  ⚠️ Could not add screenshot: {e}")
        
        # Save document
        output_dir.mkdir(parents=True, exist_ok=True)
        safe_name = "".join(c if c.isalnum() or c in (" ", "-", "_") else "_" for c in compliance_name)
        output_path = output_dir / f"{safe_name}.docx"
        doc.save(str(output_path))
        
        return output_path
    
    except Exception as e:
        print(f"  ❌ Error: {e}")
        return None


def generate_all_compliance_documents(compliances, extracted_data):
    """Generate compliance documents for all extracted compliances."""
    print(f"\n📝 Generating {len(compliances)} compliance documents...\n")
    
    output_dir = Path("generated_documents")
    created_documents = []
    
    for idx, compliance_name in enumerate(compliances, 1):
        print(f"[{idx}/{len(compliances)}] {compliance_name}...", end=" ")
        
        # Find associated assets
        logo_path, screenshot_path = find_assets_for_compliance(compliance_name, extracted_data)
        
        # Create document
        doc_path = create_compliance_document(
            compliance_name,
            logo_path,
            screenshot_path,
            output_dir
        )
        
        if doc_path:
            created_documents.append(str(doc_path))
            print("✅")
        else:
            print("❌")
    
    print(f"\n✅ Generated {len(created_documents)} documents\n")
    return created_documents


def create_summary_report(compliances, created_documents):
    """Create a summary report of generated documents."""
    report_path = Path("COMPLIANCE_GENERATION_REPORT.md")
    
    with open(report_path, "w") as f:
        f.write("# Compliance Documents Generation Report\n\n")
        f.write(f"**Generated:** {len(created_documents)} compliance documents\n\n")
        f.write("## Generated Documents\n\n")
        
        for doc in sorted(created_documents):
            doc_name = Path(doc).name
            f.write(f"- `{doc_name}`\n")
        
        f.write("\n## Summary\n\n")
        f.write(f"- **Total Compliances Found:** {len(compliances)}\n")
        f.write(f"- **Documents Created:** {len(created_documents)}\n")
        if len(compliances) > 0:
            success_rate = (len(created_documents) / len(compliances)) * 100
            f.write(f"- **Success Rate:** {success_rate:.1f}%\n")
    
    print(f"📊 Report saved to {report_path}")


def main():
    """Main execution flow."""
    print("🚀 Starting Compliance Document Generator Agent\n")
    
    try:
        # Step 1: Extract zip files
        extracted_data = extract_zip_files()
        
        if not extracted_data:
            print("❌ No files extracted from zip files")
            return
        
        # Step 2: Extract compliance names
        compliances = extract_compliance_names(extracted_data)
        
        if not compliances:
            print("❌ No compliances found")
            return
        
        # Step 3: Generate documents
        created_documents = generate_all_compliance_documents(compliances, extracted_data)
        
        # Step 4: Create report
        create_summary_report(compliances, created_documents)
        
        print("✨ Agent task completed successfully!")
        
    except Exception as e:
        print(f"\n❌ Error: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
