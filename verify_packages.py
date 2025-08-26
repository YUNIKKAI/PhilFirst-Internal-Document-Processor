import importlib
import sys
from pathlib import Path

print("🔎 Verifying installed packages...\n")

req_file = Path("requirements.txt")

if not req_file.exists():
    print("⚠️ requirements.txt not found.")
    sys.exit(0)

# Mapping of pip package names → import names
PACKAGE_MAPPING = {
    "Flask": "flask",
    "python-dateutil": "dateutil",
    "python-dotenv": "dotenv",
    "Werkzeug": "werkzeug",
    "Jinja2": "jinja2",
    "MarkupSafe": "markupsafe"
}

missing = []
with req_file.open() as f:
    for line in f:
        pkg = line.strip()
        if not pkg or pkg.startswith("#"):
            continue
        pkg_name = pkg.split("==")[0].strip()
        import_name = PACKAGE_MAPPING.get(pkg_name, pkg_name.replace("-", "_"))
        try:
            importlib.import_module(import_name)
        except ImportError:
            missing.append(pkg_name)

if missing:
    print("❌ Missing packages:", ", ".join(missing))
    sys.exit(1)
else:
    print("✅ All required packages installed")
