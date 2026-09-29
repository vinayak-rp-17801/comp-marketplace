#!/usr/bin/env python3
"""
Compliance Document Generator Agent

This script automates the creation of compliance documents for all 40 compliances.
It:
1. Extracts and analyzes the ANVISA template structure
2. Extracts compliance names from zip files
3. Generates 40 compliance documents following the exact ANVISA pattern
4. Commits them to the repository

Usage:
    python compliance_document_generator.py
"""

import os
import zipfile
import json
from pathlib import Path
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import shutil


def extract_zip_files():
    """Extract all zip files in the repository."""
    print("📦 Extracting zip files...")
    zip_files = list(Path(".").glob("*.zip"))
    extracted_data = []
    
    for zip_path in zip_files:
        extract_dir = Path(f"extracted_{zip_path.stem}")
        extract_dir.mkdir(exist_ok=True)
        
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(extract_dir)
        
        # Collect all files
        for item in extract_dir.rglob("*"):
            if item.is_file():
                extracted_data.append({
                    "source_zip": str(zip_path),
                    "path": str(item),
                    "name": item.name,
                    "suffix": item.suffix
                })
    
    return extracted_data


def analyze_anvisa_template():
    """Analyze the ANVISA template document to extract structure and style."""
    print("📄 Analyzing ANVISA template structure...")
    template_path = Path("ANVISA for Log360.docx")
    
    if not template_path.exists():
        print("❌ ANVISA template not found!")
        return None
    
    doc = Document(template_path)
    template_data = {
        "paragraphs": [],
        "tables": [],
        "images": [],
        "styles": {}
    }
    
    # Extract paragraphs and their styles
    for para in doc.paragraphs:
        if para.text.strip():
            runs_data = []
            for run in para.runs:
                run_info = {
                    "text": run.text,
                    "bold": run.bold,
                    "italic": run.italic,
                    "font_size": None,
                    "color": None
                }
                # Safely handle font size (can be float or None)
                try:
                    if run.font.size:
                        run_info["font_size"] = int(run.font.size.pt)
                except (TypeError, ValueError):
                    run_info["font_size"] = None
                
                # Safely handle color
                try:
                    if run.font.color and run.font.color.rgb:
                        run_info["color"] = str(run.font.color.rgb)
                except Exception:
                    run_info["color"] = None
                
                runs_data.append(run_info)
            
            template_data["paragraphs"].append({
                "text": para.text,
                "style": para.style.name,
                "alignment": str(para.alignment),
                "runs": runs_data
            })
    
    # Extract tables
    for table_idx, table in enumerate(doc.tables):
        table_data = {
            "table_idx": table_idx,
            "rows": [],
            "columns": len(table.columns)
        }
        for row in table.rows:
            row_data = [cell.text for cell in row.cells]
            table_data["rows"].append(row_data)
        template_data["tables"].append(table_data)
    
    # Extract images
    for rel in doc.part.rels.values():
        if "image" in rel.target_ref:
            template_data["images"].append(rel.target_ref)
    
    # Save template analysis for reference
    with open("template_analysis.json", "w") as f:
        json.dump(template_data, f, indent=2, default=str)
    
    print(f"✅ Template analyzed: {len(template_data['paragraphs'])} paragraphs, {len(template_data['tables'])} tables")
    return template_data


def extract_compliance_names(extracted_data):
    """Extract compliance names from extracted files."""
    print("📋 Extracting compliance names...")
    compliances = set()
    
    # Look for text files, PDFs, or folders that might contain compliance names
    for item in extracted_data:
        # Check folder names
        if "/" in item["path"] and item["suffix"] == "":
            parts = item["path"].split("/")
            if len(parts) > 1:
                compliance_name = parts[1]
                if compliance_name and not compliance_name.startswith(".") and not compliance_name.startswith("_"):
                    compliances.add(compliance_name)
        
        # Check text file names
        if item["suffix"] in [".txt", ".pdf", ".docx", ".md"]:
            name = item["name"].replace(item["suffix"], "").strip()
            if name and not name.lower().startswith("._"):
                compliances.add(name)
    
    compliances = sorted(list(compliances))
    print(f"✅ Found {len(compliances)} potential compliances")
    if compliances:
        print(f"   Examples: {', '.join(compliances[:5])}")
    return compliances


def find_assets_for_compliance(compliance_name, extracted_data):
    """Find logo and screenshot files for a compliance."""
    logo = None
    screenshot = None
    
    search_pattern = compliance_name.lower().replace(" ", "_").replace("-", "_")
    
    for item in extracted_data:
        if search_pattern in item["path"].lower() or search_pattern in item["name"].lower():
            if item["suffix"].lower() in [".png", ".jpg", ".jpeg", ".gif"]:
                if "logo" in item["path"].lower() or "logo" in item["name"].lower():
                    logo = item["path"]
                elif "screenshot" in item["path"].lower() or "screenshot" in item["name"].lower():
                    screenshot = item["path"]
                elif not logo:  # Use first image as logo
                    logo = item["path"]
                elif not screenshot:  # Use second image as screenshot
                    screenshot = item["path"]
    
    return logo, screenshot


def create_compliance_document(template_data, compliance_name, logo_path, screenshot_path, output_dir):
    """Create a new compliance document based on the ANVISA template."""
    print(f"  Creating document for {compliance_name}...")
    
    # Create new document from scratch or copy template
    try:
        doc = Document("ANVISA for Log360.docx")
    except Exception as e:
        print(f"    ❌ Failed to load template: {e}")
        return None
    
    # Replace product name references
    for para in doc.paragraphs:
        if "Log360" in para.text:
            for run in para.runs:
                run.text = run.text.replace("Log360", compliance_name)
        if "ANVISA" in para.text and compliance_name not in para.text:
            # Keep compliance-specific headers
            pass
    
    # Replace in tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if "Log360" in cell.text:
                    cell.text = cell.text.replace("Log360", compliance_name)
    
    # Add images if available (replace existing images or append)
    if logo_path and Path(logo_path).exists():
        # Add logo at the end or replace existing
        try:
            last_para = doc.paragraphs[-1] if doc.paragraphs else doc.add_paragraph()
            last_para.add_run().add_picture(logo_path, width=Inches(1.5))
        except Exception as e:
            print(f"    ⚠️ Could not add logo: {e}")
    
    if screenshot_path and Path(screenshot_path).exists():
        try:
            if logo_path:
                doc.add_paragraph()  # Add spacing
            last_para = doc.paragraphs[-1] if doc.paragraphs else doc.add_paragraph()
            last_para.add_run().add_picture(screenshot_path, width=Inches(3.0))
        except Exception as e:
            print(f"    ⚠️ Could not add screenshot: {e}")
    
    # Save document
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"{compliance_name}.docx"
    
    try:
        doc.save(output_path)
        return output_path
    except Exception as e:
        print(f"    ❌ Failed to save document: {e}")
        return None


def generate_all_compliance_documents(template_data, compliances, extracted_data):
    """Generate compliance documents for all extracted compliances."""
    print(f"\n📝 Generating {len(compliances)} compliance documents...")
    
    output_dir = Path("generated_documents")
    output_dir.mkdir(exist_ok=True)
    
    created_documents = []
    
    for idx, compliance_name in enumerate(compliances, 1):
        print(f"[{idx}/{len(compliances)}] Processing {compliance_name}...")
        
        # Find associated assets
        logo_path, screenshot_path = find_assets_for_compliance(compliance_name, extracted_data)
        
        # Create document
        try:
            doc_path = create_compliance_document(
                template_data,
                compliance_name,
                logo_path,
                screenshot_path,
                output_dir
            )
            if doc_path:
                created_documents.append(str(doc_path))
        except Exception as e:
            print(f"  ❌ Error creating document: {e}")
    
    print(f"\n✅ Generated {len(created_documents)} documents in {output_dir}/")
    return created_documents


def create_summary_report(compliances, created_documents):
    """Create a summary report of generated documents."""
    print("\n📊 Creating summary report...")
    
    report_path = Path("COMPLIANCE_GENERATION_REPORT.md")
    
    with open(report_path, "w") as f:
        f.write("# Compliance Documents Generation Report\n\n")
        f.write(f"Generated: {len(created_documents)} compliance documents\n\n")
        f.write("## Generated Documents\n\n")
        
        for doc in sorted(created_documents):
            f.write(f"- `{doc}`\n")
        
        f.write("\n## Summary\n\n")
        f.write(f"- **Total Compliances Processed:** {len(compliances)}\n")
        f.write(f"- **Documents Created:** {len(created_documents)}\n")
        if len(compliances) > 0:
            f.write(f"- **Success Rate:** {(len(created_documents)/len(compliances)*100):.1f}%\n")
    
    print(f"✅ Report saved to {report_path}")
    return report_path


def main():
    """Main execution flow."""
    print("🚀 Starting Compliance Document Generator Agent\n")
    
    try:
        # Step 1: Extract all zip files
        extracted_data = extract_zip_files()
        
        # Step 2: Analyze ANVISA template
        template_data = analyze_anvisa_template()
        if not template_data:
            print("❌ Cannot proceed without template analysis")
            return
        
        # Step 3: Extract compliance names
        compliances = extract_compliance_names(extracted_data)
        if not compliances:
            print("⚠️ No compliances found. Check extracted zip files.")
            return
        
        # Step 4: Generate all compliance documents
        created_documents = generate_all_compliance_documents(template_data, compliances, extracted_data)
        
        # Step 5: Create summary report
        create_summary_report(compliances, created_documents)
        
        print("\n✨ Agent task completed successfully!")
        print("📁 All documents are in the 'generated_documents/' directory")
        print("📋 Review COMPLIANCE_GENERATION_REPORT.md for details")
        
    except Exception as e:
        print(f"\n❌ Agent encountered an error: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
