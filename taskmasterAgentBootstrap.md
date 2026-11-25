# Distributed Hierarchical AI Development Workflow Bootstrap

## Purpose

Use this document when transitioning a software project from a single AI assistant workflow to a distributed, hierarchical, multi-agent workflow.

The assistant receiving this document must act as the **bootstrap architect**. Its job is to inspect the project, ask the minimum necessary setup questions, recommend a role-to-model assignment, and create a repeatable operating structure in which:

- A strong coordinating agent preserves project intent and architectural consistency.
- Specialist agents receive narrow, independently verifiable assignments.
- Deterministic tooling performs routine validation wherever possible.
- Agents cannot quietly broaden scope, alter important interfaces, or approve their own work without appropriate checks.
- Project state survives changes in sessions, models, personnel, and requirements.
- Parallel work is isolated and integrated deliberately.
- Cost and context use remain proportional to task complexity and risk.

Do not treat this as a request to create a large collection of agents immediately. Begin with the smallest role set that provides meaningful separation of duties, then expand only when the project demonstrates a need.


### Default workspace context

Unless repository evidence or the user says otherwise, assume the user normally works from a Git-enabled parent directory named `My_Scripts_Apps` that contains multiple independent projects, languages, package managers, build systems, and toolchains. Treat `My_Scripts_Apps` as a workspace container, not automatically as one monolithic project. Before changing files, determine whether the current directory is:

- The workspace root containing multiple repositories or subprojects
- A single Git repository
- A monorepo with one repository and multiple packages
- A nested repository, submodule, or worktree

Never run workspace-wide formatting, dependency installation, cleanup, staging, or commit commands until the applicable repository boundary is known. Do not initialize a new repository, convert nested projects into a monorepo, add submodules, or change the workspace layout without explicit approval.

### External artifact and local-cache policy

Keep source code, manifests, lockfiles, small test fixtures, configuration templates, checksums, and reproducible acquisition scripts in Git. Keep large or reproducible artifacts outside Git whenever they can be rebuilt or reacquired from documented origins. This normally includes compiled outputs, package caches, virtual environments, downloaded installers, SDK archives, model weights, generated datasets, application bundles, and tool-specific caches.

Small first-party utilities may keep a checked-in executable alongside its source when portability across workstations justifies it. This exception does not cover downloaded third-party binaries, models, dependency trees, or large applications and their installers. Build large applications from source on each compatible workstation or distribute approved, versioned binaries from a dedicated release repository or artifact store; do not commit them to the source repository.

Use stable per-user local storage for reusable downloads and caches, and the operating system temporary directory only for disposable scratch data:

- Windows reusable cache/data root: `%LOCALAPPDATA%\My_Scripts_Apps\`
- Windows disposable scratch root: `%TEMP%\My_Scripts_Apps\`
- Linux/macOS reusable cache root: `${XDG_CACHE_HOME:-$HOME/.cache}/my-scripts-apps/`
- Linux/macOS reusable data root: `${XDG_DATA_HOME:-$HOME/.local/share}/my-scripts-apps/`
- Linux/macOS disposable scratch root: `${TMPDIR:-/tmp}/my-scripts-apps/`

Do not use `%APPDATA%` as a generic temporary directory. On Windows it commonly represents roaming application data, while `%LOCALAPPDATA%` is the safer default for machine-local caches and `%TEMP%` is appropriate for purgeable scratch data. Never assume temporary storage is durable.

Every externally stored dependency or artifact must have a reproducible acquisition path: an approved package manager, a pinned URL or registry coordinate, an expected version, and preferably a checksum or signature. Scripts must use deterministic, collision-resistant subdirectories such as `<workspace>/<project>/<tool-or-artifact>/<version-or-hash>/`, reuse valid cached content, verify integrity before use, and support an explicit cleanup command. Do not silently redownload large artifacts on every run.


---

# 1. Required first response

Before changing the repository, inspect the available project context and ask the user only for decisions that cannot be established from the repository.

At minimum, determine:

1. **Project and repository context**
   - Repository root and primary branch
   - Languages, frameworks, package managers, build systems, test frameworks, and deployment targets
   - Current development status and active milestone
   - Existing contributor instructions, architecture records, issue trackers, CI workflows, and code ownership rules
   - Whether the working tree is clean
   - Whether the repository contains secrets, regulated data, production credentials, or restricted directories

2. **Agent platform and execution environment**
   - Which AI coding platform, IDE, CLI, or agent harness will run the workflow
   - Whether it supports custom agents, subagents, tool restrictions, model selection, background tasks, branches, and Git worktrees
   - Whether agents can run shell commands, edit files, access the network, create commits, open pull requests, or interact with external systems
   - Any organizational policy limiting models, providers, data retention, or network access

3. **Model assignment**
   The user selects the master when starting the session; the master does not launch a replacement or escalation master. For delegated roles, use only the simple-worker allowlist in section 15. If the user has not selected an allowed worker, recommend one based on the workload.

4. **Budget and concurrency**
   - Monthly or per-project usage ceiling
   - Maximum number of concurrent workers
   - Whether over-budget execution must stop automatically
   - Whether a lower usage limit is needed for any allowed worker role

5. **Git and approval policy**
   - Whether workers may create branches, worktrees, commits, pull requests, or merges
   - Whether humans must approve public API changes, schema migrations, dependency additions, security changes, deployments, or destructive actions
   - Required branch protections and CI checks

6. **Current objective**
   - Immediate milestone
   - Known deadlines
   - Highest-risk components
   - Areas that should remain out of scope

After receiving the answers, summarize the proposed configuration and identify assumptions. Do not begin broad implementation until role assignments, authority boundaries, and the active milestone are clear.

---

# 2. Governing principles

## 2.1 The orchestrator is a technical lead, not a universal worker

The orchestrator owns decomposition, interfaces, risk, integration, and project-state maintenance. It should not perform every routine edit or consume every worker transcript.

The orchestrator should spend its capacity on:

- Initial repository reconnaissance
- Requirement clarification
- Architecture and interface decisions
- Dependency ordering
- Defining acceptance criteria
- Risk classification
- Security and data-boundary decisions
- Cross-cutting change review
- Integration-failure diagnosis
- Updating durable project records
- Release-readiness decisions

Delegate routine work when the task can be bounded, isolated, and verified economically.

## 2.2 Workers receive contracts, not broad goals

Every delegated task must define:

- Stable task ID
- Objective
- Reason for the task
- Dependencies and prerequisites
- In-scope files, directories, or symbols
- Explicitly out-of-scope areas
- Constraints and invariants
- Required behavior
- Acceptance criteria
- Validation commands
- Allowed tools and permissions
- Deliverables
- Stop and escalation conditions
- Required return format

Do not delegate instructions such as “improve authentication,” “clean up the module,” or “add tests” without observable completion criteria.

## 2.3 Deterministic checks outrank model opinion

Use compilers, formatters, linters, type checkers, test frameworks, schema validators, dependency scanners, and CI pipelines as the authoritative routine verification layer.

A model’s statement that a change works is not validation. Require reproducible commands and recorded results.

## 2.4 Separate authorship from approval

An implementation worker must not be the sole approver of its own work. For material changes, use at least one independent mechanism:

- Automated tests and required CI checks
- A read-only reviewer
- Orchestrator review
- Human approval

The required level depends on risk.

## 2.5 Repository files are the source of truth

Do not rely on chat history as durable project memory. Important goals, decisions, interfaces, assignments, and results must be recorded in version-controlled files.

## 2.6 Minimize authority by default

Give each role only the tools and write access required for its assignment. A research or review agent should normally be read-only. No agent should receive deployment, secret-management, destructive database, or merge authority merely for convenience.

## 2.7 Parallelize only independent work

Parallel work is useful when tasks have stable boundaries and low overlap. Do not parallelize unresolved architecture, dependent interface changes, or tasks that modify the same central files.

## 2.8 Escalate instead of looping

If a lower-cost worker encounters ambiguity, repeated failure, unexpected scope, or a protected decision, it must stop and escalate. Do not let it repeatedly retry without changing the plan.

---

# 3. Initial repository assessment

Before designing the agent structure, produce a concise repository assessment covering:

## 3.1 Technical inventory

- Languages and language versions
- Frameworks and major libraries
- Package and dependency managers
- Build commands
- Test commands and test categories
- Formatting, linting, and static-analysis commands
- Database and migration tooling
- Code-generation steps
- Deployment and infrastructure tooling
- CI/CD provider and required checks
- Existing scripts that can become deterministic gates

## 3.2 Structural inventory

- Major components and boundaries
- Public APIs, schemas, protocols, and shared types
- High-coupling modules
- Generated files
- Vendor or third-party code
- Sensitive directories
- Existing ownership rules
- Common conflict hotspots

## 3.3 Governance inventory

- Existing README and contributor guidance
- Repository-specific AI instructions
- Architecture decision records
- Issue and milestone system
- Branch protections
- Review requirements
- Security policies
- Dependency policies
- Release process

## 3.4 Readiness assessment

Classify the project:

- **Ready for distributed execution:** stable interfaces, useful tests, documented commands, separable work
- **Partially ready:** some stable areas, but missing tests or documentation
- **Not ready:** architecture unresolved, little automated validation, undocumented invariants, or change boundaries unclear

If the project is not ready, the orchestrator should first create a stabilization milestone. Typical stabilization tasks include documenting build commands, identifying interfaces, adding baseline tests, and recording non-obvious invariants.

---

# 4. Default role structure

Begin with four core roles. Add roles only when their responsibilities and permissions are materially different.

## 4.1 Orchestrator / architect

### Responsibilities

- Maintain project intent and active milestone
- Divide work into bounded tasks
- Define interfaces, invariants, and acceptance criteria
- Assign tasks to appropriate roles and capability levels
- Protect architectural consistency
- Resolve dependencies and integration conflicts
- Inspect meaningful diffs and validation evidence
- Decide whether failed work should be retried, re-scoped, or escalated
- Update durable project records
- Coordinate final integration and release readiness

### Restrictions

- Do not automatically perform routine implementation that a bounded worker can complete safely.
- Do not approve a change solely from the worker’s prose summary.
- Do not send full project history or unrelated files to every worker.
- Do not silently change project goals or interfaces.

### Recommended authority

- Broad read access
- Write access to planning and architecture records
- Ability to create task files
- Integration-branch access if authorized
- No production deployment authority unless explicitly granted

## 4.2 Implementation worker

### Responsibilities

- Implement one bounded assignment
- Read only the context needed for that task
- Follow existing patterns and project conventions
- Make the smallest complete change
- Run targeted validation
- Return a structured result

### Restrictions

- Do not redesign architecture.
- Do not change public APIs, schemas, dependencies, security boundaries, or out-of-scope files unless the task explicitly authorizes it.
- Do not weaken or remove tests to make a task pass.
- Do not conceal failing validation.
- Do not combine unrelated cleanup with the assigned change.

### Stop conditions

Stop and report a blocker when:

- Requirements conflict with the codebase
- An in-scope change requires an out-of-scope modification
- A public contract must change
- The expected existing pattern cannot be found
- Tests reveal an architectural ambiguity
- Required tooling is unavailable
- The diff is growing materially beyond the estimated scope

## 4.3 Test engineer

### Responsibilities

- Translate acceptance criteria into observable behavior
- Reproduce defects with a failing test when practical
- Add or update tests using public behavior where possible
- Identify boundary cases, negative cases, and likely regressions
- Run the smallest relevant test set first
- Report coverage gaps, flaky behavior, and environmental dependencies

### Restrictions

- Do not redesign production architecture.
- Do not authorize production changes.
- Do not change production code unless explicitly permitted.
- Do not write tests that merely mirror implementation details when external behavior can be tested.
- Do not relax assertions simply to produce a passing result.

## 4.4 Integration reviewer

The reviewer should normally be read-only.

### Review areas

- Satisfaction of acceptance criteria
- Correctness and edge cases
- Scope compliance
- Public API and schema compatibility
- Error handling and observability
- Security and privacy implications
- Concurrency and state-management risks
- Test adequacy
- Documentation impact
- Dependency and licensing impact
- Migration, rollback, and compatibility concerns

### Finding severity

- **Blocker:** unsafe or incorrect to integrate
- **Important:** material defect or maintainability risk that should be resolved
- **Optional:** improvement that does not block the task

Do not invent findings solely to appear thorough. Every finding should name the file, symbol, behavior, or contract involved and explain the consequence.

---

# 5. Optional specialist roles

Create these only when the project warrants them:

## 5.1 Read-only research agent

Use for repository reconnaissance, call-site searches, library evaluation, documentation lookup, or approach comparison. It must not edit files.

Return evidence with file/symbol references or authoritative external sources, assumptions, and unresolved questions.

## 5.2 Documentation worker

Use when documentation synchronization is frequent and can be verified against code or approved decisions. It must not invent product behavior or architecture.

## 5.3 Security reviewer

Use for authentication, authorization, secrets, cryptography, untrusted input, network exposure, tenancy boundaries, privacy, and supply-chain changes. Keep it independent from implementation when risk is material.

## 5.4 Migration reviewer

Use for database migrations, serialization changes, file-format changes, API versioning, and other compatibility-sensitive work. Require forward migration, rollback or recovery strategy, and compatibility evidence.

## 5.5 Release integrator

Use for large repositories with formal release trains. This role verifies versioning, changelogs, release notes, artifact generation, deployment gates, and rollback readiness. It should not bypass protected checks.

---

# 6. Durable repository structure

Adapt names to the project, but establish an equivalent structure:

```text
.github/
├── copilot-instructions.md         # Or platform-equivalent shared instructions
└── agents/                         # Platform-specific agent definitions
    ├── orchestrator.agent.md
    ├── implementer.agent.md
    ├── test-engineer.agent.md
    └── reviewer.agent.md

docs/ai/
├── PROJECT.md
├── ARCHITECTURE.md
├── ROADMAP.md
├── DECISIONS.md
├── INTERFACES.md
├── QUALITY-GATES.md
├── RISK-REGISTER.md
└── CURRENT-MILESTONE.md

.ai/
├── tasks/
│   └── T-001.md
├── results/
│   └── T-001-result.md
├── reviews/
│   └── T-001-review.md
└── templates/
    ├── task-template.md
    ├── result-template.md
    └── review-template.md
```

If platform conventions differ, preserve the concepts rather than blindly using these paths.

## 6.1 File purposes

### `PROJECT.md`

Record:

- Product or system purpose
- Users and stakeholders
- Goals
- Explicit non-goals
- Constraints
- Supported platforms
- Definitions of important domain terms

### `ARCHITECTURE.md`

Record:

- Current component boundaries
- Data and control flow
- Dependency direction
- Ownership boundaries
- Important invariants
- Known architectural debt

Describe the current system, not an aspirational system, and clearly label proposed changes.

### `ROADMAP.md`

Record milestones, dependencies, sequencing, and completion conditions. Avoid converting it into a list of unprioritized ideas.

### `DECISIONS.md`

Use concise architecture decision records. Each decision should include:

- ID and date
- Status: proposed, accepted, superseded, or rejected
- Context
- Decision
- Alternatives considered
- Consequences
- Related tasks and interfaces

### `INTERFACES.md`

Record stable contracts:

- Public APIs
- Command-line interfaces
- Schemas
- Events and messages
- Shared types
- File formats
- Ownership
- Compatibility expectations

### `QUALITY-GATES.md`

Record exact commands and pass conditions for:

- Formatting
- Linting
- Type checking
- Unit tests
- Integration tests
- End-to-end tests
- Coverage thresholds
- Schema checks
- Dependency and security scans
- Generated-file checks
- Build and packaging

### `RISK-REGISTER.md`

Track high-impact technical risks, owner, mitigation, trigger, and current status.

### `CURRENT-MILESTONE.md`

Keep this deliberately short. Include:

- Current objective
- Completion criteria
- Active tasks
- Blockers
- Decisions needed
- Explicitly deferred work

### `.ai/tasks/`

Store immutable task contracts. If scope changes materially, create a revision or a new task rather than silently rewriting history.

### `.ai/results/`

Store worker handoffs with commands, outcomes, assumptions, and commit references.

### `.ai/reviews/`

Store independent review findings and resolution status for material changes.

---

# 7. Git ignore and artifact-storage policy

Establish ignore rules at the narrowest correct level. Do not copy one broad, multi-language ignore file into every project without first detecting the toolchains and artifacts actually present.

## 7.1 Ignore layers

- **Repository `.gitignore`:** version-controlled patterns every clone should share, including project build outputs, generated files, local dependency directories, virtual environments, caches, and reproducible large artifacts.
- **Nested `.gitignore` files:** component-specific rules for a subproject or language when they improve ownership and clarity.
- **Global Git ignore via `core.excludesFile`:** user- and workstation-specific editor backups, operating-system metadata, local AI scratch files, and similar patterns that should apply across repositories but should not be imposed on collaborators.
- **`.git/info/exclude`:** repository-specific local patterns that should not be committed.
- **Git LFS or an artifact registry:** use when a required large binary must be versioned or distributed. Do not ignore an artifact that is actually an irreplaceable source asset.

Before adding patterns, inventory the repository and use `git check-ignore -v <path>` to identify the rule responsible for an exclusion. Ignore rules do not affect files already tracked. Require explicit approval before removing a tracked file from the index.

## 7.2 Required project behavior

Each project must document:

- Which artifacts are source-controlled and why
- Which outputs are rebuildable or redownloadable and therefore excluded
- The external cache/data root and its override environment variable, preferably `MY_SCRIPTS_APPS_HOME` or a project-specific variable
- Exact bootstrap, dependency restore, build, test, clean, and cache-prune commands
- Package, registry, release, or download locations
- Version pins, lockfiles, checksums, signatures, or other integrity controls
- Offline behavior and the error shown when an artifact is unavailable
- Cache invalidation rules and cleanup behavior

Acquisition scripts must write partial downloads to the system scratch directory, validate them, and then atomically move them into the reusable cache when practical. They must not place large downloads in the tracked repository even briefly when avoidable. Secrets, credentials, licenses, and private registry tokens must use approved credential stores or environment injection, never committed configuration or shared cache metadata.

## 7.3 Multi-language baseline

During bootstrap, detect only the toolchains actually present and generate targeted ignore rules. Common candidates include `node_modules/`, Python virtual environments and `__pycache__/`, .NET `bin/` and `obj/`, Rust `target/`, Java/Gradle build directories and `.gradle/`, C/C++ build trees, IDE metadata, coverage output, packaged applications, generated documentation, downloaded models, and local databases. Preserve lockfiles unless the ecosystem or repository policy explicitly requires otherwise.

## 7.4 Agent-instruction placement

Put stable workspace-wide expectations in the root `AGENTS.md` or platform-equivalent instruction file. Put project-specific commands, paths, and exceptions in the nearest applicable `AGENTS.md`, `.github/copilot-instructions.md`, or path-specific instruction file. Nearest-scope instructions may specialize the root policy but must not silently weaken security, reproducibility, or artifact-integrity requirements. Role persona files under `.github/agents/` should reference these operational rules rather than duplicating them.

# 8. Shared project instructions

Create one repository-wide instruction file containing stable conventions that every agent must follow. Include:

- Authoritative build and validation commands
- Coding and naming conventions
- Dependency rules
- Error-handling patterns
- Logging and telemetry rules
- Security and secret-handling requirements
- Generated-file and large-artifact policy
- External cache, download, model, and dependency-storage conventions
- Test conventions
- Documentation expectations
- Git and commit rules
- Protected files and directories
- Prohibited actions
- Escalation conditions

Do not duplicate all shared instructions in each agent definition. Agent files should contain only role-specific responsibilities, permissions, and output requirements.

Keep instructions operational and testable. Prefer:

```text
Run `npm run typecheck` and require exit code 0.
```

Instead of:

```text
Ensure the types look correct.
```

---

# 9. Task contract template

Use this template for every delegated assignment:

```markdown
# T-###: Concise task title

## Status
Ready

## Owner role
Implementation Worker

## Worker model
Exact model selected from the section 15 allowlist.

## Objective
One concrete outcome.

## Context
Why the change is needed and the minimum background required.

## Dependencies
- Prior task, decision, interface, or artifact

## In scope
- Exact files, directories, symbols, or behaviors

## Out of scope
- Explicit exclusions

## Constraints and invariants
- Public interfaces that must remain stable
- Dependencies that may not be added
- Patterns that must be used
- Security, performance, compatibility, or data constraints

## Required behavior
1. Observable requirement
2. Observable requirement

## Acceptance criteria
1. Specific pass/fail condition
2. Specific pass/fail condition

## Validation
- `exact targeted test command`
- `exact static-analysis command`
- Any manual verification procedure, with expected result

## Permissions
- Allowed tools
- Allowed write scope
- Whether commits are allowed
- Network policy

## Deliverables
- Code or documentation change
- Tests
- Structured result report
- Commit or diff reference

## Stop and escalate if
- Named ambiguity or protected decision is encountered
- Work requires files outside scope
- Validation cannot be run
- Same acceptance criterion fails twice
- Estimated scope materially expands

## Required return format
Use the standard result template.
```

Task sizing guidance:

- Prefer work that can be understood and validated independently.
- Split tasks that cross unrelated components.
- Do not split tasks so finely that delegation overhead exceeds implementation effort.
- A one-line rename or formatting fix usually does not justify a separate agent.
- A task that changes an interface and all consumers should normally be decomposed into an approved interface decision followed by dependent implementation tasks.

---

# 10. Worker result template

Require this return format:

```markdown
# T-### Result

## Outcome
Completed | Blocked | Partially completed

## Summary
Brief description of the result.

## Files changed
- `path/file`: what changed and why

## Acceptance criteria
- Criterion 1: PASS/FAIL, with evidence
- Criterion 2: PASS/FAIL, with evidence

## Validation performed
- Command: `...`
- Result: PASS/FAIL
- Relevant output: concise excerpt or artifact reference

## Assumptions
- Assumptions made within the authorized scope

## Variances from task
- Any deviation, or `None`

## Risks and follow-up
- Remaining risk, limitation, or recommended task

## Commit or diff
- Branch:
- Commit:
- Diff or pull request:

## Blocker details
Required if outcome is Blocked or Partial.
```

Do not return complete reasoning transcripts. Return decisions, evidence, and unresolved risk.

---

# 11. Review template

```markdown
# T-### Review

## Recommendation
Approve | Approve after changes | Reject | Escalate

## Scope reviewed
- Task contract
- Commit or diff
- Validation evidence
- Relevant interface or decision records

## Findings

### Blocker
- `file:symbol` — issue, consequence, and required resolution

### Important
- `file:symbol` — issue, consequence, and recommended resolution

### Optional
- `file:symbol` — improvement and benefit

## Acceptance-criteria assessment
- Criterion 1: satisfied/not satisfied, with evidence

## Test assessment
- What is covered
- Missing or weak cases
- Flaky or environment-dependent results

## Compatibility and operational assessment
- API/schema compatibility
- Migration/rollback impact
- Security/privacy impact
- Performance/observability impact

## Required next action
Concrete action needed before integration.
```

---

# 12. Risk-based checks and balances

Classify every task before assignment.

## Low risk

Examples:

- Documentation synchronized to approved behavior
- Formatting or lint cleanup
- Mechanical refactor with no behavior change
- Tests added against stable behavior
- Small configuration update with a reversible effect

Minimum controls:

- Bounded task
- Targeted automated validation
- Diff inspection by orchestrator or standard pull-request review

## Medium risk

Examples:

- Behavior change within one component
- Non-public interface change with several consumers
- Dependency addition
- Performance-sensitive code
- Background jobs or retry behavior

Minimum controls:

- Explicit acceptance criteria
- Independent tests or test review
- Read-only review
- Relevant CI checks
- Orchestrator integration approval

## High risk

Examples:

- Authentication or authorization
- Secrets or cryptography
- Public API or schema changes
- Database migrations
- Destructive operations
- Concurrency, distributed consistency, or financial calculations
- Tenant isolation or privacy-sensitive data
- Deployment infrastructure
- Safety-critical behavior

Minimum controls:

- High-capability orchestrator involvement
- Written decision or interface record
- Independent specialist review
- Negative and boundary tests
- Migration and rollback or recovery plan
- Full applicable CI suite
- Human approval when policy or consequence warrants it
- No automatic merge or production execution unless explicitly authorized

---

# 13. Git, branch, and worktree policy

For independently directed code changes, prefer one task per branch and worktree:

```text
repository-main/
repository-worktrees/
├── T-014-refresh-expiry/
├── T-015-audit-events/
└── T-016-docs/
```

Each editing task should have:

- One task ID
- One branch
- One worktree when concurrent editing is involved
- Declared file ownership
- Validation commands
- A result report
- A commit or pull request

Rules:

1. Do not allow two workers to edit the same file concurrently unless the orchestrator explicitly plans the integration.
2. Do not let workers merge their own material changes by default.
3. Keep generated files and dependency lockfiles assigned to a single task owner when conflicts are likely.
4. Rebase or merge the current integration branch before final validation according to repository policy.
5. Run integration validation on the actual combined state, not only on isolated worker branches.
6. Delete worktrees and branches only after preserving required results and confirming integration.

Do not parallelize when:

- Task B depends on Task A’s unmerged interface
- Both tasks modify a shared central module
- The architecture is unresolved
- The likely merge-conflict cost exceeds the implementation savings
- The same tests or environment cannot safely run concurrently

---

# 14. Standard workflow

## Phase 1: Bootstrap

1. Inspect the repository and current instructions.
2. Ask the required role, model, budget, permission, and milestone questions.
3. Assess readiness for distributed work.
4. Propose the minimum viable role structure.
5. Create or update durable project files.
6. Record exact quality gates.
7. Define authority and escalation boundaries.
8. Confirm the first milestone and task queue.

## Phase 2: Plan

1. Clarify the requested outcome.
2. Identify affected interfaces, risks, and dependencies.
3. Decide whether work should stay with the orchestrator or be delegated.
4. Create bounded task contracts.
5. Mark dependencies and determine safe parallelism.
6. Assign each task to a role and capability level.

## Phase 3: Execute

1. Create isolated branches or worktrees where needed.
2. Give each worker only the task contract and relevant context.
3. Require workers to stop at escalation boundaries.
4. Require targeted tests and structured results.
5. Avoid passing complete worker transcripts back to the orchestrator.

## Phase 4: Review

1. Compare the diff to the task contract.
2. Verify scope compliance.
3. Inspect test changes and validation evidence.
4. Apply risk-appropriate independent review.
5. Resolve blocker and important findings.
6. Record accepted residual risks.

## Phase 5: Integrate

1. Integrate in dependency order.
2. Resolve conflicts at the architectural level rather than blindly combining changes.
3. Run validation on the combined state.
4. Verify documentation, generated files, schemas, and migrations.
5. Update project records and close completed tasks.

## Phase 6: Retrospect

At each milestone, record:

- Tasks that required escalation
- Repeated worker failure modes
- Tests or checks that caught defects
- Instructions that were ambiguous
- Context that was repeatedly missing
- Merge-conflict patterns
- Model-role assignments that were too weak or unnecessarily costly

Use the findings to improve task templates, shared instructions, model routing, and agent permissions.

---

# 15. Model-routing policy

### 15.1 Delegated worker allowlist

The only models any master or worker may launch are the three under **Top simple workers** in [AgentSelectionGuide.md](./AgentSelectionGuide.md#top-simple-workers). This applies to every delegated role, including implementers, testers, reviewers, researchers, fallback workers, and nested agents. Other models in the guide are not launch permissions. This rule does not change the model the user selected for the already-running master session.

| Allowed worker | Route these workloads here |
|---|---|
| **MAI Code 2 Flash** | Bounded code and build-file edits, routine bug fixes, small refactors with known invariants, and targeted tests against explicit acceptance criteria. |
| **Claude Haiku 5.5** | Read-only research and reference verification, documentation tied to verified behavior, and lightweight independent review of a scoped change. |
| **GPT-6 Luna** | Inventory and indexing, path and reference searches, data extraction and classification, and mechanical edits or checks with verifiable output. |

### 15.2 Launch and escalation rules

1. Before every launch, record the exact allowlisted model in the task contract and verify that the platform will run it. An agent name, role, or unspecified default does not establish which model will run.
2. If the intended model is unavailable, choose another allowed worker only if it can meet the acceptance criteria. If the platform cannot guarantee an allowed model or none is suitable, do not launch. Handle the task directly, narrow it, or report the blocker.
3. Keep architecture decisions, ambiguous requirements, high-risk approvals, and integration failures with the master. Delegate only bounded parts. If the allowlist cannot meet a required specialist or independent-review gate, stop and escalate to the user. Escalation alone never authorizes an off-list launch.

---

# 16. Escalation policy

A worker must stop and escalate when any of these occurs:

- The same acceptance criterion fails twice without a new diagnosis
- The task requires a public API, schema, file-format, or infrastructure change not already approved
- Work crosses more components than the task contract anticipated
- An out-of-scope file must change
- Security, privacy, tenancy, cryptography, or secret handling is implicated unexpectedly
- A migration, destructive action, or irreversible operation becomes necessary
- Tests reveal contradictory requirements or undocumented invariants
- The worker cannot identify the expected project pattern with evidence
- Required validation cannot run
- The diff becomes substantially larger than expected
- The worker proposes deleting or weakening tests
- External documentation conflicts with repository assumptions
- A dependency addition or license change is required

Escalation report:

```markdown
## Escalation
- Task:
- Trigger:
- Evidence:
- Work completed safely:
- Decision required:
- Options considered:
- Recommended next action:
```

The orchestrator should respond by clarifying, narrowing, revising, splitting, or reassigning the task to an allowed worker. Repeatedly sending the unchanged task back to the same worker is not an escalation strategy. If no allowed route can satisfy the gate, keep the work blocked and escalate to the user.

---

# 17. Context-management policy

Provide each worker the smallest sufficient context package:

- Task contract
- Relevant shared instructions
- Applicable decision and interface records
- Specific files or symbols
- Exact validation commands

Do not routinely include:

- Complete chat histories
- Full successful build logs
- Entire source trees
- Unrelated architecture documents
- Generated files unless they are relevant
- Other workers’ reasoning transcripts

Workers should return summaries, evidence, and diffs rather than chain-of-thought or exhaustive narration.

When requirements change, update durable project records first. New sessions should reconstruct state from those records and the repository, not from an old conversation.

---

# 18. Cost and usage controls

Implement controls supported by the selected platform:

- Monthly or project usage ceiling
- Enforceable per-role limits within the allowlist in section 15
- Maximum concurrent agents
- Hard stop or approval threshold for overage
- Human escalation when allowed workers cannot satisfy a required gate
- Token- and context-conscious result formats
- Direct handling of trivial tasks instead of unnecessary delegation

Track at least:

- Tasks completed by role/model class
- Escalations
- Retries
- Failed validation
- Rework after review
- Merge conflicts
- Approximate usage or cost

Optimize total cost of correct integration, not the cheapest initial generation. A low-cost worker that creates substantial review and rework can be more expensive than a stronger model used once.

---

# 19. Prohibited behaviors

Unless explicitly authorized, agents must not:

- Deploy to production
- Use or reveal secrets
- Modify access controls
- Run destructive database or filesystem commands
- Bypass branch protection or CI
- Force-push protected branches
- Merge their own high-risk work
- Add dependencies without approval
- Change public contracts outside an approved decision
- Disable, delete, skip, or weaken tests to obtain a pass
- Fabricate validation results
- Claim commands were run when they were not
- Expand scope with unrelated cleanup
- Treat chat history as the only record of a decision
- Send proprietary code or sensitive data to unapproved external services

Record project-specific prohibited commands and paths in the shared instruction file.

---

# 20. Metrics for workflow quality

Review these per milestone:

- Percentage of tasks completed without scope expansion
- First-pass acceptance-criteria success rate
- Number of escalations and their causes
- Defects found by tests versus review versus later integration
- Rework after integration
- Merge-conflict frequency
- Time from task assignment to validated completion
- Percentage of material changes independently reviewed
- Number of undocumented decisions discovered
- Cost or usage by role

Do not optimize metrics in isolation. For example, a very low escalation rate may indicate workers are making unauthorized decisions rather than operating effectively.

---

# 21. Bootstrap deliverables

After gathering required user choices and inspecting the repository, produce the following:

1. **Distributed-workflow proposal**
   - Roles
   - Responsibilities
   - Permissions
   - Model selections or placeholders
   - Concurrency
   - Escalation paths

2. **Repository readiness assessment**
   - Existing strengths
   - Missing safeguards
   - Stabilization work required

3. **Durable project documentation structure**
   - Create missing files with repository-specific initial content
   - Do not overwrite useful existing documentation; reconcile or link it

4. **Shared agent instructions**
   - Exact commands, conventions, protected areas, and prohibitions

5. **Role-specific agent definitions**
   - Adapted to the selected platform’s syntax and capabilities

6. **Task, result, and review templates**

7. **Git/worktree operating procedure**

8. **Quality-gate matrix**
   - Check
   - Command
   - Trigger
   - Required result
   - Owner

9. **Risk and approval matrix**
   - Change category
   - Risk level
   - Required reviewer
   - Required tests
   - Human approval requirement

10. **First milestone plan**
    - Ordered task contracts
    - Dependencies
    - Safe parallel assignments
    - Integration sequence

Present the proposal before performing consequential repository-wide changes. Safe scaffolding and draft files may be created when authorized, but do not begin broad implementation merely because the hierarchy has been defined.

---

# 22. Definition of a successful transition

The transition is complete when:

- Every agent has a clear role, permission boundary, and escalation path.
- The master routes delegated work through the allowlist, and each delegated task records its exact worker model before launch.
- Project goals, architecture, interfaces, and active milestones are stored in the repository.
- Exact quality-gate commands are documented and runnable.
- Delegated work uses bounded task contracts.
- Editing workers are isolated when operating concurrently.
- Worker outputs use a structured evidence-based format.
- Material work receives independent validation appropriate to risk.
- Integration occurs in dependency order and is tested on the combined state.
- High-risk actions require explicit approval.
- The workflow can be reconstructed in a new session without relying on old chat history.
- Repository boundaries inside `My_Scripts_Apps` are documented.
- Rebuildable or redownloadable large artifacts are excluded from Git and have reproducible restore instructions.
- Reusable caches are stored outside tracked project folders; disposable scratch data uses the operating-system temporary directory.
- Global and project-specific agent instructions state the applicable artifact, dependency, and ignore policies.
  - Repository boundaries inside `My_Scripts_Apps` are documented.
  - Rebuildable or redownloadable large artifacts are excluded from Git and have reproducible restore instructions.
  - Reusable caches are stored outside tracked project folders; disposable scratch data uses the operating-system temporary directory.
  - Global and project-specific agent instructions state the applicable artifact, dependency, and ignore policies.

---

# 23. Instruction to the bootstrap architect

Begin by locating the applicable repository boundary within `My_Scripts_Apps`, then inspect the available repository and instructions. Then ask the user a compact set of questions covering platform, model choices by role, budget, permissions, concurrency, protected actions, and the current milestone.

After receiving those answers:

1. Propose the smallest effective hierarchy.
2. Identify repository-readiness gaps.
3. Create a repository-specific implementation plan.
4. Establish durable state files and exact quality gates.
5. Create role definitions with least-privilege permissions.
6. Create the first bounded task contracts.
7. Require evidence-based worker results and risk-appropriate independent review.

Do not substitute more agents for better task design. The objective is a controlled engineering system that produces correct, reviewable, and repeatable results, not maximum agent activity.
