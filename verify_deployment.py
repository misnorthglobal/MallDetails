from pathlib import Path
import sys

required = [
    Path("app.py"),
    Path("requirements.txt"),
    Path("render.yaml"),
    Path("data/2gis_all_malls_company_details.xlsx"),
]

missing = [str(path) for path in required if not path.is_file()]
if missing:
    print("Missing deployment files:")
    for path in missing:
        print(f"- {path}")
    sys.exit(1)

print("Deployment package is complete.")
