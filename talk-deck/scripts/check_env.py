"""Check (and with --install, set up) what the talk-deck skill needs on this machine.

Usage: python check_env.py [--install]

- Python packages (python-pptx, Pillow, matplotlib): installed with pip when --install is given.
- Renderer: PowerPoint (Windows) or LibreOffice + pdftoppm. Never installed by this script (system-level, large);
  the install command is printed so the agent can ask the user first.
Exit code 0 = ready to build and render, 1 = can build but cannot render, 2 = cannot build.
"""
import importlib, platform, shutil, subprocess, sys

PKGS = [('pptx', 'python-pptx', True), ('PIL', 'Pillow', True), ('matplotlib', 'matplotlib', False)]


def has(mod):
    try:
        importlib.import_module(mod); return True
    except Exception:
        return False


def powerpoint():
    if platform.system() != 'Windows': return False
    try:
        import winreg
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r'PowerPoint.Application\CLSID')); return True
    except OSError:
        return False


def soffice():
    for name in ('soffice', 'libreoffice'):
        if shutil.which(name): return shutil.which(name)
    for p in (r'C:\Program Files\LibreOffice\program\soffice.exe', '/Applications/LibreOffice.app/Contents/MacOS/soffice'):
        try:
            open(p, 'rb').close(); return p
        except OSError:
            pass
    return None


install = '--install' in sys.argv
rows, build_ok = [], True
if sys.version_info < (3, 8):
    rows.append(('Python', f'{platform.python_version()} (need 3.8+)', 'MISSING')); build_ok = False
else:
    rows.append(('Python', platform.python_version(), 'ok'))

missing = [(m, p, req) for m, p, req in PKGS if not has(m)]
if missing and install:
    subprocess.call([sys.executable, '-m', 'pip', 'install', '--quiet', *[p for _, p, _ in missing]])
    importlib.invalidate_caches()
for m, p, req in PKGS:
    ok = has(m)
    rows.append((p, 'required' if req else 'optional (redrawing graphs)', 'ok' if ok else 'MISSING'))
    if req and not ok: build_ok = False

pp, so, pdf = powerpoint(), soffice(), shutil.which('pdftoppm')
rows.append(('PowerPoint', 'render via scripts/render_pptx.ps1', 'ok' if pp else 'not found'))
rows.append(('LibreOffice', so or 'fallback renderer', 'ok' if so else 'not found'))
rows.append(('pdftoppm (poppler)', 'needed with LibreOffice', 'ok' if pdf else 'not found'))
render = 'powerpoint' if pp else ('libreoffice' if so and pdf else None)

w = max(len(r[0]) for r in rows)
for a, b, c in rows: print(f'{a:<{w}}  {c:<9}  {b}')
print()
if not build_ok:
    print('BUILD: not ready. Run:  python check_env.py --install   (or: pip install python-pptx pillow matplotlib)')
else:
    print('BUILD: ready')
if render:
    print(f'RENDER: ready ({render})')
else:
    print('RENDER: no renderer. Ask the user before installing one:')
    s = platform.system()
    if s == 'Windows':
        print('  winget install TheDocumentFoundation.LibreOffice ; winget install oschwartz10612.Poppler')
    elif s == 'Darwin':
        print('  brew install --cask libreoffice ; brew install poppler')
    else:
        print('  sudo apt install libreoffice poppler-utils')
    print('  Until then: build the deck, tell the user it was not visually checked, and ask for screenshots.')
sys.exit(0 if build_ok and render else (1 if build_ok else 2))
