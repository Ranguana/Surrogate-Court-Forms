"""
runner.py — Frozen launcher for Probate HQ

PyInstaller freezes THIS file (bundling the Python interpreter + all pip deps).
At runtime, it receives the app directory as an argument, adds it to sys.path,
and runs app.py from that directory.

This means:
  - Python + dependencies are locked in the binary (no pip needed on target)
  - app.py, generators.py, templates/, static/ live OUTSIDE the binary
  - Those files can be updated via GitHub releases without rebuilding the DMG
"""
import sys
import os
import importlib

def main():
    if len(sys.argv) < 2:
        print("Usage: probate-server <app_directory>", file=sys.stderr)
        sys.exit(1)

    app_dir = os.path.abspath(sys.argv[1])

    if not os.path.isfile(os.path.join(app_dir, "app.py")):
        print(f"ERROR: app.py not found in {app_dir}", file=sys.stderr)
        sys.exit(1)

    # Make the app directory the working directory and importable
    os.chdir(app_dir)
    sys.path.insert(0, app_dir)

    # Load .env if python-dotenv is available
    try:
        from dotenv import load_dotenv
        env_path = os.path.join(app_dir, ".env")
        if os.path.isfile(env_path):
            load_dotenv(env_path)
    except ImportError:
        pass

    # Import and run the Flask app
    spec = importlib.util.spec_from_file_location("app", os.path.join(app_dir, "app.py"))
    module = importlib.util.module_from_spec(spec)
    sys.modules["app"] = module
    spec.loader.exec_module(module)

if __name__ == "__main__":
    main()
