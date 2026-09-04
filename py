#!/usr/bin/env python3
"""Builds MK-Rides-GitHub-Publication-Guide.docx from the drop-in files
so the document is guaranteed to match them verbatim."""

from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_LINE_SPACING

def read(p):
    with open(p, encoding="utf-8") as f:
        return f.read().rstrip("\n")

doc = Document()

# Base style tweaks
normal = doc.styles["Normal"]
normal.font.name = "Calibri"
normal.font.size = Pt(11)

def code_block(text):
    for line in text.split("\n"):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        p.paragraph_format.left_indent = Inches(0.25)
        r = p.add_run(line if line else " ")
        r.font.name = "Courier New"
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(0x1A, 0x1A, 0x1A)

def note(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    r = p.add_run(text)
    r.italic = True
    r.font.size = Pt(10)
    r.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

# ── Title ────────────────────────────────────────────────────────────────────
doc.add_heading("MK RIDES — GitHub Publication Guide", level=0)
p = doc.add_paragraph()
p.add_run("Target repository: ").bold = True
r = p.add_run("mat-tew/mk-rides")
r.font.name = "Courier New"
p.add_run("   ·   Scope: publish the existing project as-is — no Docker, no rebuilds, no code changes.")

doc.add_paragraph(
    "This guide ships the complete publication blueprint for the MK RIDES project: the exact "
    "repository file tree, the initial commit message, the .gitignore and .env.example to drop in, "
    "and the precise git / GitHub CLI commands to push it to mat-tew/mk-rides cleanly — with "
    "verification steps that guarantee no secret ever leaves your machine. The same sequence is "
    "also encoded in the executable publish-to-github.sh helper."
)

# ── Project snapshot ─────────────────────────────────────────────────────────
doc.add_heading("1 · What is being published", level=1)
doc.add_paragraph(
    "MK RIDES is Kenya's modern ride-hailing platform — “Your ride. Your route. Your choice.” "
    "The stack below was verified directly against the deployed production build "
    "(mk-rides-shsp.arcada.app), so the tree and environment files match the artifact exactly:"
)
doc.add_paragraph("Frontend: React + TypeScript + Vite + Tailwind CSS v4, lucide icons", style="List Bullet")
doc.add_paragraph("Maps: Leaflet with OpenStreetMap tiles", style="List Bullet")
doc.add_paragraph("Backend: Supabase (auth, Postgres + RLS, the `users` table confirmed in the client bundle)", style="List Bullet")
doc.add_paragraph("Payments: M-Pesa-ready via Supabase Edge Functions (Daraja keys stay server-side)", style="List Bullet")
doc.add_paragraph("Verified routes: / /home /login /signup /onboarding /driver /driver/onboarding "
                  "/driver/earnings /admin /admin/login /history /notifications /profile /safety "
                  "/support /help /legal /legal/:slug and a * not-found route", style="List Bullet")

# ── File tree ────────────────────────────────────────────────────────────────
doc.add_heading("2 · Exact repository file tree", level=1)
doc.add_paragraph(
    "Publish this tree exactly as it exists. node_modules/, dist/ and the local .env are the only "
    "things present on disk that stay out of the repository (handled by .gitignore)."
)
code_block(read("REPO_TREE.txt"))

# ── Commit message ───────────────────────────────────────────────────────────
doc.add_heading("3 · Initial commit message", level=1)
doc.add_paragraph(
    "Use this exact message (stored in COMMIT_MESSAGE.txt, which git commit -F consumes):"
)
code_block(read("COMMIT_MESSAGE.txt"))
note("Subject line alone is sufficient for a compact history; the body records the stack "
     "for anyone cloning later.")

# ── .gitignore ───────────────────────────────────────────────────────────────
doc.add_heading("4 · .gitignore (drop-in)", level=1)
doc.add_paragraph(
    "Place at the repository root. The first block is what guarantees secrets never publish: "
    ".env and all .env.* variants are ignored, with !.env.example re-included deliberately."
)
code_block(read(".gitignore"))

# ── .env.example ─────────────────────────────────────────────────────────────
doc.add_heading("5 · .env.example (drop-in)", level=1)
doc.add_paragraph(
    "This is the only environment file that ever reaches GitHub. It documents every variable "
    "with placeholder values so collaborators can cp .env.example .env and fill in their own."
)
code_block(read(".env.example"))
note("Rule of thumb: anything prefixed VITE_ is baked into the browser bundle at build time — "
     "the anon key is acceptable there only because Row Level Security is on. The service-role "
     "key and all Daraja/M-Pesa secrets never touch the repo or the Vite build; they are set "
     "with `supabase secrets set` for the Edge Functions.")

# ── Git commands ─────────────────────────────────────────────────────────────
doc.add_heading("6 · Precise git commands to publish", level=1)

doc.add_heading("6.1 Prerequisites (once per machine)", level=2)
code_block(
"""# GitHub CLI installed and authenticated as mat-tew
gh --version                                     # install from cli.github.com if missing
gh auth login                                    # GitHub.com → HTTPS → login via browser
gh auth status                                   # confirm you are 'mat-tew'

git config --global user.name  "mat-tew"
git config --global user.email "YOUR-EMAIL@users.noreply.github.com"  # keeps real email private""")

doc.add_heading("6.2 One-time setup inside the project", level=2)
code_block(
"""cd /path/to/mk-rides                    # the EXISTING project — nothing rebuilt, no Docker added

# Drop in the two safety files from this guide (once)
cp /path/to/guide/.gitignore   .gitignore
cp /path/to/guide/.env.example .env.example
cp /path/to/guide/COMMIT_MESSAGE.txt COMMIT_MESSAGE.txt

# Prove .env can never be committed
git check-ignore .env && echo ".env is ignored ✓"
grep -qxF ".env" .gitignore && echo "rule present ✓\"""")

doc.add_heading("6.3 Initialise, scan, stage, commit", level=2)
code_block(
"""git init -b main                          # fresh repo on branch 'main' (idempotent)

# Secret scan of exactly the files git would commit — must print nothing
git ls-files -c -o --exclude-standard | grep -v '.env.example' | \\
  xargs grep -nIE 'service_role|sbp_[A-Za-z0-9]{20,}|eyJhbGciOi[A-Za-z0-9_-]{10,}|PRIVATE KEY' \\
  && echo 'STOP: secret found' || echo 'clean ✓'

git add -A
git status --short                        # eyeball: .env must NOT appear
git rm --cached .env 2>/dev/null || true  # belt-and-braces: unstage it if it ever slipped in

git commit -F COMMIT_MESSAGE.txt          # one clean, well-described initial commit""")

doc.add_heading("6.4 Create mat-tew/mk-rides and push", level=2)
code_block(
"""# Option A — recommended: GitHub CLI creates the repo AND wires the remote
gh repo create mat-tew/mk-rides --private --source . --remote origin \\
  --description "MK RIDES — Kenya's modern ride-hailing platform. React + Vite + Tailwind, Supabase backend, M-Pesa ready."
git push -u origin main

# Option B — repo created in the github.com UI instead
git remote add origin git@github.com:mat-tew/mk-rides.git
git push -u origin main""")

doc.add_heading("6.5 Publish later / flip visibility when ready", level=2)
code_block(
"""gh repo edit mat-tew/mk-rides --visibility public --accept-visibility-change-consequences""")

# ── If a secret already slipped ──────────────────────────────────────────────
doc.add_heading("7 · If a secret was ever committed before", level=1)
doc.add_paragraph(
    "Deleting the file in a new commit is NOT enough — it stays in history. From the project root:"
)
code_block(
"""# Remove .env (or the leaked file) from EVERY revision, then force-push the clean history
git filter-repo --path .env --invert-paths        # or: pip install git-filter-repo
git push origin main --force-with-lease

# Then rotate the exposed key immediately:
#   Supabase Dashboard → Project Settings → API Keys → roll the anon/service keys
#   Safaricom Daraja portal → regenerate consumer key/secret + passkey""")

# ── Checklist ────────────────────────────────────────────────────────────────
doc.add_heading("8 · Final no-secrets checklist", level=1)
for item in [
    ".gitignore contains .env, .env.*, node_modules/, dist/ — with the !.env.example exception.",
    "git check-ignore .env exits 0 before the first commit.",
    "The secret scan of committable files prints “clean ✓”.",
    "git status after git add -A shows .env.example — never .env.",
    "VITE_SUPABASE_ANON_KEY is referenced via import.meta.env only, never hard-coded.",
    "Service-role key and M-Pesa (Daraja) credentials are set with supabase secrets set, "
    "never stored in the repo.",
    "Repo created as private first; flipped public only after a final history review "
    "(git log -p | grep -i key  returns nothing).",
]:
    doc.add_paragraph(item, style="List Number")

doc.add_paragraph()
note("Automated alternative: run the included helper — ./publish-to-github.sh preflight /path/to/mk-rides "
     "verifies every check above, then ./publish-to-github.sh publish /path/to/mk-rides performs steps "
     "6.3–6.4 verbatim and exits.")

doc.save("MK-Rides-GitHub-Publication-Guide.docx")
print("saved MK-Rides-GitHub-Publication-Guide.docx")
