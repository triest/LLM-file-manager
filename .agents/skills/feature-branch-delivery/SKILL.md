---
name: feature-branch-delivery
description: >-
  Use for every feature or fix in this repository. Start on a dedicated Git
  branch, commit and push the change, verify it, then merge it into main only
  after the verification succeeds.
---

# Feature branch delivery

## Workflow

1. Start from an up-to-date `main` and create a branch named `feature/<short-name>` or `fix/<short-name>`.
2. Implement one reviewable feature and update the relevant project documentation.
3. Commit the change on the feature branch with a concise, specific message.
4. Run the checks requested by the user and any applicable project-required checks. Report failures and fix them before merging.
5. Push the feature branch to `origin` so the work is backed up and reviewable.
6. Only after the checks succeed, merge the feature branch into `main` and push `main`.
7. Confirm the final branch and commit IDs. Keep the feature branch available on the remote.

Do not merge a failing feature. Do not amend or overwrite another person's commits. If authentication or a merge conflict blocks the push/merge, preserve the branch and report the exact blocker.
