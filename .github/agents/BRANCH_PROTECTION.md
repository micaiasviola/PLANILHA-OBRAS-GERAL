Branch protection instructions for the Agent Guard

This file explains how to activate branch protection that requires the Agent Guard workflow to pass before merges.

Quick steps:
- Ensure .github/workflows/agent-guard.yml exists (it does in this repo).
- Push this branch and create a pull request.
- Run the Agent Guard workflow on the branch (it should run automatically on PR creation).
- In the GitHub repository Settings -> Branches -> Add branch protection rule for "main" (and/or "master").
- Under "Require status checks to pass before merging", select the workflow run "Agent Guard" (it will appear after the workflow has run).
- Save the rule. From now on, merges into protected branches will be blocked until the Agent Guard checks pass.

If you want automation to enforce this policy across multiple repos, consider using organization-level branch protection templates or GitHub REST API scripts.
