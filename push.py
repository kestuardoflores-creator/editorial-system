import os, pathlib, sys, subprocess

# ── Config ────────────────────────────────────────────────────────────────────
ENV_FILE  = pathlib.Path("/home/kevin/Documents/Python/c_p/g_t_kevinflores.env")
REPO_NAME = "kestuardoflores-creator/editorial-system"
BASE      = pathlib.Path(__file__).resolve().parent

EXCLUDE_DIRS  = {".venv", "__pycache__", ".git", ".claude", "projects"}
EXCLUDE_FILES = {"repo_dump_v2.md"}
EXCLUDE_EXT   = {".pyc"}

# ── Use project venv; install PyGithub if needed ─────────────────────────────
VENV_PYTHON = BASE / ".venv" / "bin" / "python3"
if VENV_PYTHON.exists() and pathlib.Path(sys.executable).resolve() != VENV_PYTHON.resolve():
    subprocess.check_call([str(VENV_PYTHON), "-m", "pip", "install", "--quiet", "PyGithub"])
    os.execv(str(VENV_PYTHON), [str(VENV_PYTHON)] + sys.argv)

try:
    from github import Github, GithubException
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "--quiet",
                           "--break-system-packages", "PyGithub"])
    from github import Github, GithubException


def read_token():
    for line in ENV_FILE.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line and not line.startswith("#"):
            _, _, v = line.partition("=")
            return v.strip()
    raise ValueError(f"Token not found in {ENV_FILE}")


def push_file(repo, repo_path, content):
    msg = f"chore: update {repo_path}"
    try:
        existing = repo.get_contents(repo_path)
        repo.update_file(repo_path, msg, content, existing.sha)
        print(f"  Updated : {repo_path}")
    except GithubException:
        repo.create_file(repo_path, msg, content)
        print(f"  Created : {repo_path}")


def collect_files():
    for f in sorted(BASE.rglob("*")):
        if not f.is_file():
            continue
        rel   = f.relative_to(BASE)
        parts = rel.parts
        if any(p in EXCLUDE_DIRS for p in parts):
            continue
        if rel.name in EXCLUDE_FILES:
            continue
        if rel.suffix in EXCLUDE_EXT:
            continue
        yield f, str(rel).replace("\\", "/")


token = read_token()
repo  = Github(token).get_repo(REPO_NAME)

print(f"Pushing to {REPO_NAME}...\n")
ok = err = 0

for local, repo_path in collect_files():
    try:
        try:
            content = local.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            content = local.read_bytes()
        push_file(repo, repo_path, content)
        ok += 1
    except Exception as e:
        print(f"  Error   : {repo_path} — {e}")
        err += 1

print(f"\n{ok} pushed, {err} errors.")
print(f"https://github.com/{REPO_NAME}")
