Branch Protection Template

Purpose
- Ensure all code changes are validated by the repository's agent checks before merging into protected branches (main/master).

Policy (recommended settings)
1. Protect branches: main, master
2. Require pull requests for all changes (disable direct pushes by developers)
3. Require status checks to pass before merging:
   - Agent Guard (workflow file: .github/workflows/agent-guard.yml)
   - Any other CI tests you add
4. Require at least 1 approving review from a reviewer (more if desired)
5. Require linear history (no merge commits) — optional but recommended
6. Enforce branch protection for administrators: Optional (choose based on org policy)

How to enable (via GitHub UI)
1. Go to the repository on GitHub -> Settings -> Branches -> Branch protection rules
2. Click "Add rule"
3. Enter branch name pattern: main (repeat for master if used)
4. Check "Require a pull request before merging"
5. Under "Require status checks to pass before merging", select the check: "Agent Guard" (or its workflow run name). Ensure the workflow has run at least once so the check appears.
6. Optionally require reviews and enable "Require linear history"
7. Save changes

Notes
- Administrators can be exempted; consider enforcing the checks for everyone to fully prevent accidental merges.
- If your CI check names differ, pick the correct status check in step 5. The agent-guard.yml workflow created earlier must be run at least once on main to appear as a status check to require.
