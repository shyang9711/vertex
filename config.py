from __future__ import annotations

APP_NAME = "Vertex"
UPDATE_POLICY_ASSET_NAME = "update_policy.json"

# 🔢 bump this each time you ship a new version
APP_VERSION = "0.2.12"

# 🔗 set this to your real GitHub repo once you create it,
GITHUB_REPO = "shyang9711/vertex"
GITHUB_RELEASES_URL = f"https://github.com/{GITHUB_REPO}/releases"
GITHUB_API_LATEST   = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"

# Optional: name of the EXE asset in GitHub Releases if you want auto-download
UPDATE_ASSET_NAME = "vertex.exe"

# -------------------- Constants --------------------
ENTITY_TYPES = [
    "", "Individual / Sole Proprietor", "Partnership", "LLC",
    "S-Corporation", "Corporation (C-Corp)", "Exempt Organization",
    "Trust / Estate", "Nonprofit", "Other"
]
US_STATES = [
    "", "AL","AK","AZ","AR","CA","CO","CT","DE","DC","FL","GA","HI","ID","IL","IN",
    "IA","KS","KY","LA","ME","MD","MA","MI","MN","MS","MO","MT","NE","NV","NH","NJ",
    "NM","NY","NC","ND","OH","OK","OR","PA","RI","SC","SD","TN","TX","UT","VT","VA",
    "WA","WV","WI","WY"
]

ROLES = ["Officer", "Employee", "Spouse", "Parent", "Business Owner", "Business"]