#!/usr/bin/env python3
"""
Setup script to install dependencies and run the compliance document generator agent.
"""

import subprocess
import sys

def install_dependencies():
    """Install required Python packages."""
    print("📦 Installing dependencies...")
    packages = ["python-docx"]
    
    for package in packages:
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package, "-q"])
            print(f"  ✅ {package}")
        except subprocess.CalledProcessError as e:
            print(f"  ❌ Failed to install {package}: {e}")
            return False
    
    return True

def run_agent():
    """Run the compliance document generator agent."""
    print("\n🚀 Running Compliance Document Generator Agent...\n")
    try:
        subprocess.check_call([sys.executable, "compliance_document_generator.py"])
        return True
    except subprocess.CalledProcessError as e:
        print(f"\n❌ Agent failed: {e}")
        return False

if __name__ == "__main__":
    if install_dependencies():
        success = run_agent()
        sys.exit(0 if success else 1)
    else:
        print("\n❌ Failed to install dependencies")
        sys.exit(1)
