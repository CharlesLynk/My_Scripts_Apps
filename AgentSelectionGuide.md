# GitHub Copilot Multi-Agent Model Selection Guide

**Reference date:** November 24, 2026  
**Basis:** `taskmasterAgentBootstrap.md` and the supplied GitHub Copilot model-picker screenshot.

> Model availability, context limits, and credit rates can change. Recheck the Copilot model picker before changing the long-term role map.

## TL;DR

### Recommended starting configuration

| Role | Primary model | Fallback or escalation |
|---|---|---|
| Master/orchestrator | **GPT-6.2 Sol** | Claude Fable 5.2 |
| Independent integration reviewer | **Claude Sonnet 5.5** | GPT-6.2 Sol |
| Premium architectural escalation | **Claude Fable 5.2** | GPT-6.2 Astra |
| Difficult terminal/integration diagnosis | **GPT-6.2 Astra** | Claude Fable 5.2 |
| Substantial implementation worker | **MAI Code 2** | GPT-6.1 Terra |
| Windows/C# worker | **GPT-6.1 Terra** | MAI Code 2 |
| Simple coding worker | **MAI Code 2 Flash** | GPT-6 Luna |
| Read-only research/documentation worker | **Claude Haiku 5.5** | GPT-6 Luna |
| Utility/indexing worker | **GPT-6 Luna** | Claude Haiku 5.5 |

### Short selection rules

- Use **GPT-6.2 Sol** for normal planning, delegation, review, and integration.
- Use **Claude Sonnet 5.5** as an independent reviewer so the implementing and approving model families differ.
- Escalate ambiguous architecture, repeated failures, migrations, public-interface changes, and destructive operations to **Claude Fable 5.2**.
- Use **GPT-6.2 Astra** only when a difficult terminal-heavy investigation justifies premium pricing.
- Route substantial but bounded implementation to **MAI Code 2** or **GPT-6.1 Terra**.
- Route deterministic, easy-to-check tasks to **MAI Code 2 Flash**, **Claude Haiku 5.5**, or **GPT-6 Luna**.
- Judge models by **cost per accepted integration**, including retries and review rework, not cost per request.

### Master candidates to test

1. **GPT-6.2 Sol**: expected best default balance.
2. **Claude Sonnet 5.5**: conservative alternative and independent reviewer.
3. **Claude Fable 5.2**: premium capability benchmark.
4. **GPT-6.1 Terra**: economy control.
5. **GPT-6.2 Astra**: optional terminal-heavy premium candidate.

### Top simple workers

1. **MAI Code 2 Flash**: bounded source changes and routine fixes.
2. **Claude Haiku 5.5**: research, documentation, and lightweight independent review.
3. **GPT-6 Luna**: indexing, extraction, classification, and mechanical work.

---

## Role requirements

The master agent should handle:

- Repository boundaries and working-tree state.
- Task decomposition, interfaces, invariants, acceptance criteria, permissions, and stop conditions.
- Risk classification and worker selection.
- Review of actual diffs and validation evidence.
- Integration decisions and escalation after repeated failure.

The master should not perform routine implementation when a bounded worker can do it safely. Prioritize instruction adherence, review quality, integration judgment, uncertainty handling, and cost per accepted result.

---

## Master-agent options

| Model | Context | Input / output credits per 1M | Best role | Main trade-off |
|---|---:|---:|---|---|
| **GPT-6.2 Sol** | 1.1M | 200 / 1,000 | Default master and integration reviewer | More expensive than economy models, but materially cheaper than premium models |
| **Claude Sonnet 5.5** | 1M | 200 / 1,000 | Alternate master and cross-provider reviewer | Similar price to Sol; select based on internal instruction-adherence testing |
| **Claude Fable 5.2** | 1.1M | 800 / 4,000 | Architecture and high-risk escalation | Four times Sol/Sonnet pricing |
| **GPT-6.2 Astra** | 1.1M | 800 / 4,000 | Terminal-heavy failure diagnosis | Premium price is unjustified for normal orchestration |
| **GPT-6.1 Terra** | 1.1M | 100 / 500 | Economy master or Windows/C# specialist | Weaker default choice where deep review is the central requirement |
| **MAI Code 2** | 1M | 100 / 500 | Senior implementation worker | Better aligned with implementation than governance-heavy orchestration |

### GPT-6.2 Sol

Use for routine orchestration, task contracts, result review, milestone records, and integration decisions. Escalate when requirements conflict, several workers fail the same criterion, or the task changes a public interface, file format, security boundary, migration path, or destructive operation.

### Claude Sonnet 5.5

Use as the independent reviewer when GPT-6.2 Sol is master. Test Sonnet as the primary master if strict adherence to scope, stop conditions, and validation distinctions is more important than aggressive execution.

### Claude Fable 5.2

Reserve for architectural ambiguity, hidden coupling, conflicting evidence, high-risk changes, and repeated failures. Require explicit escalation approval unless the task is preclassified as high risk.

### GPT-6.2 Astra

Reserve for command-heavy investigations across builds, tests, repositories, or toolchains. Compare Astra and Fable on real failure cases, then designate one as the normal premium escalation model.

### GPT-6.1 Terra

Use as an economy orchestration baseline and as a specialist for C#/.NET, Windows PowerShell, and terminal-based build/test loops.

### MAI Code 2

Use for bounded multi-file implementation, deep debugging, refactoring, and implementation-plus-test work. Do not make it the sole default master until it demonstrates equivalent governance and review performance.

---

## Worker routing

| Task type | Primary worker | Escalate when |
|---|---|---|
| Targeted source edit or routine fix | **MAI Code 2 Flash** | Scope expands, existing pattern is unclear, or a public contract must change |
| Repository inventory or evidence gathering | **Claude Haiku 5.5** | Findings conflict or architectural interpretation is required |
| Documentation synchronization | **Claude Haiku 5.5** | Behavior is not approved or cannot be verified from source/tests |
| Indexing, extraction, classification, formatting | **GPT-6 Luna** | Judgment or nonmechanical code changes are required |
| Multi-file implementation or nontrivial refactor | **MAI Code 2** | Invariants are uncertain, tests are weak, or side effects cross system boundaries |
| C#/.NET or Windows PowerShell | **GPT-6.1 Terra** | Autodesk/Office/vendor runtime behavior is uncertain |
| Public API, migration, file format, security, or state-recovery change | **GPT-6.2 Sol** or **Claude Sonnet 5.5** | Architecture remains ambiguous or validation cannot adequately detect failure |

### Minimum worker contract

Every worker assignment should include:

- Exact in-scope files and symbols.
- Explicit out-of-scope areas.
- Acceptance criteria.
- Validation commands and expected evidence.
- Permission boundaries.
- Stop and escalation conditions.
- Required structured result format.

Workers must stop rather than improvise when the expected pattern is missing, a public contract must change, an out-of-scope file is required, or validation cannot establish the requested result.

---

## Candidate evaluation

Run the same repository-bootstrap and review exercises through the master candidates. Score:

- Correct repository-boundary detection.
- Detection of modified and untracked files before editing.
- Separation of workspace-wide and project-specific instructions.
- Quality and boundedness of task contracts.
- Correct validation selection.
- Distinction between build success and runtime validation in AutoCAD, Plant 3D, Excel, VSD Craft, or hardware.
- Detection of unauthorized interface, dependency, binary, or file-format changes.
- Compliance with stop and escalation conditions.
- Review of diffs and evidence instead of reliance on worker summaries.
- Credits consumed per accepted, integrated result.

Retain one normal master, one reviewer from a different provider, and one premium escalation model.

---

## Models to deprioritize

- **Claude Fable 5 / 5.1:** Fable 5.2 shows more context and lower displayed pricing.
- **Claude Haiku 4.5:** 160K context at 100/500 versus Haiku 5.5 at 1M and 10/50.
- **GPT-5.6 Luna:** 20/120 versus GPT-6 Luna at 10/50.
- **GPT-5.3-Codex:** 400K context and 175/1,400 make it a poor default value compared with enabled alternatives.
- **Crossed-out or policy-disabled models:** do not include unavailable models in the durable role map.

---

## Cost and governance rules

1. Use premium models only for explicit escalation classes.
2. Send workers only the task contract, relevant instructions, selected files/symbols, and validation commands.
3. Return decisions, diffs, evidence, and unresolved risks to the master, not full worker transcripts.
4. Do not delegate work when delegation overhead exceeds the edit.
5. Limit concurrency until repository boundaries, worktrees, and validation are stable.
6. Track first-pass acceptance, retries, failed validation, review findings, merge conflicts, and credits by model.
7. Place inexpensive models on deterministic tasks where failure is easy to detect.
8. Escalate ambiguity early instead of allowing low-cost workers to loop.
9. Keep implementation and approval separate for material changes.

---

## Execution safeguards

- Use **Local** execution for AutoCAD, Plant 3D, Office/VBA, AutoHotkey, VSD Craft, Windows registry, and hardware-dependent work.
- Use a **new Git worktree** for each independently directed editing task when practical.
- Use **Plan** for unfamiliar or consequential work, then **Interactive** for implementation.
- Use **Assisted permissions** only with terminal sandboxing and clear repository boundaries.
- Use **Manual permissions** for AppData, registry, installed applications, network shares, package installation, deployment, hardware, destructive operations, and Git publication.
- Enable **terminal sandboxing by default**.
- Never use local **Autopilot + All permissions + no sandbox** on this repository.
- Restrict Autopilot to isolated, reversible, well-tested work in disposable worktrees.
- Report build validation separately from runtime validation inside AutoCAD, Plant 3D, Excel, VSD Craft, or real hardware.

---

## Provenance

- `taskmasterAgentBootstrap.md` supplied the workflow, authority, risk, task-contract, and escalation requirements.
- The supplied GitHub Copilot model-picker screenshot supplied the displayed model names, availability, context sizes, and credit rates.
- Treat benchmark claims as directional because results vary by benchmark version, reasoning effort, and agent harness.
