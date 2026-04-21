# Git Rollback and Cleanup Guide (Solo Repo + VS Code)

Practical, low-risk ways to clean up Git history while protecting your local files.

Use this guide when you have already committed (and maybe pushed) and want to undo, rewrite, or tidy history safely.

## Before You Touch History

Run these first:

```bash
git status
git branch backup/pre-cleanup-$(Get-Date -Format yyyyMMdd-HHmm)
```

Why this matters:
- `git status` confirms your current state.
- A backup branch gives you an instant recovery point.

## Quick Decision Guide

Pick the line that matches your situation:

1. Option A (Primary): I pushed several commits and want one clean commit while keeping local files.
   **Command:** `git reset --soft <good-commit>`
   **Rewrite history:** Yes
   **Force push:** Usually `git push --force-with-lease`

2. Option B: I pushed bad code and want to undo safely without rewriting history.
   **Command:** `git revert <commit>`
   **Rewrite history:** No
   **Force push:** No

3. Option C: I have multiple local commits not pushed yet and want to squash.
   **Command:** `git reset --soft HEAD~N`
   **Rewrite history:** Local only
   **Force push:** No (if never pushed)

4. Option D: I only want to undo the most recent commit but keep all edits staged.
   **Command:** `git reset --soft HEAD~1`
   **Rewrite history:** Local only
   **Force push:** No (if never pushed)

## Option A (Primary): Soft Reset, Re-Commit, Force With Lease

Best for solo repositories when you want one clean commit on remote.

### What it does

- Moves branch history back to a known good commit.
- Keeps your files exactly as-is.
- Lets you create a cleaner commit history.

### Steps

1. Find the commit you want to reset to:

```bash
git log --oneline --decorate -n 20
```

2. Soft reset:

```bash
git reset --soft <good-commit>
```

3. Confirm staged changes:

```bash
git status
```

4. Re-commit cleanly:

```bash
git commit -m "Clean final implementation"
```

5. Update remote safely:

```bash
git push --force-with-lease
```

Notes:
- Prefer `--force-with-lease` over `--force`.
- If someone else has pushed in the meantime, lease protection stops you from overwriting unexpectedly.

## Option B: Revert (No History Rewrite)

Best when you want maximum safety and no force push.

```bash
git revert <commit-hash>
git push
```

Use this when:
- The bad commit is already on shared or protected history.
- You want explicit traceability of what was undone.

## Option C: Squash Local Commits Before First Push

If commits are local only:

```bash
git reset --soft HEAD~3
git commit -m "Single clean commit"
git push
```

## Option D: Undo Last Commit, Keep Work

```bash
git reset --soft HEAD~1
```

Result:
- Last commit removed.
- All edits still staged.

## Safety Reference

- `git reset --soft`
   File impact: Keeps files exactly as-is and keeps changes staged.
   History impact: Moves branch pointer backward.

- `git reset --mixed`
   File impact: Keeps files exactly as-is and unstages changes.
   History impact: Moves branch pointer backward.

- `git reset --hard`
   File impact: Discards tracked working-tree changes.
   History impact: Moves branch pointer backward.

- `git revert`
   File impact: Applies inverse changes in a new commit.
   History impact: Adds commit, does not rewrite existing history.

## VS Code Equivalent (Soft Reset)

1. Open Source Control.
2. Open commit history (via Source Control or command palette).
3. Right-click target commit.
4. Choose reset branch to this commit.
5. Select Soft.

## Recovery if You Reset Wrong

Use reflog to recover previous branch tips:

```bash
git reflog -n 20
git reset --soft <reflog-entry-or-commit>
```

As long as objects are still in reflog history, recovery is usually possible.

## Copy/Paste Playbooks

### Rewrite to one clean commit after push

```bash
git status
git branch backup/pre-cleanup-$(Get-Date -Format yyyyMMdd-HHmm)
git log --oneline --decorate -n 20
git reset --soft <good-commit>
git commit -m "Clean final implementation"
git push --force-with-lease
```

### Safe undo with no rewrite

```bash
git status
git revert <bad-commit>
git push
```

## Final Guidance

- For solo repos, rewriting with `reset --soft` plus `--force-with-lease` is usually the cleanest workflow.
- For shared history, prefer `git revert`.
- Avoid `reset --hard` unless you explicitly want to discard local tracked changes.
