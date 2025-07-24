"""
Metadata:
    File Name: test_github.py
    Owner: Andrew (SilicaStormSiam)
    Purpose: Tests GitHub API authentication and repository access for SilicaStormSiam/SilicaStormSiam.
    Version: 1.0.0
    Version Control:
        Version 1.0.0 (2025-07-24): Initial version to test PAT and repository access.
        Version 0.3.0 (2025-07-23): N/A.
        Version 0.2.0 (2025-07-22): N/A.
        Version 0.1.0 (2025-07-21): N/A.
    Change Log:
        2025-07-24 (v1.0.0): Created script to verify GitHub PAT and repository.
        2025-07-23 (v0.3.0): N/A.
        2025-07-22 (v0.2.0): N/A.
        2025-07-21 (v0.1.0): N/A.
"""
from github import Github
from dotenv import load_dotenv
import os
load_dotenv(r"C:\Users\Andrew\silicastormsiam\scripts\.env", encoding="utf-8")
token = os.getenv("GITHUB_TOKEN")
if not token:
    print("Error: GITHUB_TOKEN not found in .env")
else:
    g = Github(token)
    try:
        repo = g.get_repo("SilicaStormSiam/SilicaStormSiam")
        print(f"Successfully accessed repository: {repo.full_name}")
    except Exception as e:
        print(f"Error accessing repository: {e}")