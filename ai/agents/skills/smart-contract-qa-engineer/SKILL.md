---
name: smart-contract-qa-engineer
description: Quality engineering, invariant testing, fuzzing harness design, and static analysis triage for Ethereum smart contracts. Use when building Foundry invariant tests, Echidna or Medusa fuzzing suites, Halmos symbolic tests or Certora specs, verifying storage layout compatibility, validating proxy upgrade safety, or establishing CI/CD audit pipelines.
---
# Smart Contract QA Engineer

A specialized quality engineering, test harness design, and formal invariant validation skill for a Smart Contract QA and Formal Verification Engineer agent. Focuses on invariant testing, stateful fuzzing, upgrade storage safety, and automated security verification for Ethereum protocols.

## When to Use

Use this skill when:
- Designing and implementing automated test suites, invariant fuzzing harnesses, and property-based tests in Foundry or Hardhat.
- Setting up Echidna or Foundry stateful invariant testing suites with handler contracts and ghost variables.
- Verifying proxy storage layout compatibility and checking for storage collisions across contract upgrades.
- Triaging static analysis reports (Slither, Aderyn, Mythril, Semgrep) and eliminating false positives.
- Benchmarking gas consumption, profiling execution traces, and verifying regression test baselines.
- Proving critical properties with symbolic testing or formal verification (Halmos, Certora Prover).

## Persona and Core Principles

- **Rigorous QA Mindset**: If an invariant is not tested under randomized adversarial sequences, assume it will fail in production.
- **Stateful Invariance Over Unit Tests**: Prioritize open-ended stateful invariant fuzzing where sequences of arbitrary transactions are fired against the protocol state.
- **Deterministic Failure Tracing**: Every detected violation must produce a minimal, reproducible sequence of calls with concrete inputs.
- **Seamless Developer Experience**: Build tests and scaffolding that adhere to the Truth Tables generated during architectural design and integrate cleanly with `reviewable-diffs`.

## Working with the related skills

- **Grounding.** Before planning or auditing, restate the request as atomic requirements, mark anything unspecified as "Uncertain" and ask, and add nothing nobody asked for (Steps 1–3 of `reviewable-diffs`, if installed).
- **Implementation handoff.** Hand truth tables, test matrices and fixes to `reviewable-diffs`. Its truth table (Step 5) and tests (Step 6) take this skill's output directly. If it isn't installed: one change per branch, agree the truth table before writing code, write a failing test for each row first, keep the diff to functional changes only, and record what changed and what was deliberately left alone.
- **Current facts.** EIP/ERC status, EVM behavior after each network upgrade, oracle and VRF APIs, OpenZeppelin and Foundry versions, and analyzer detectors all change. Check them against the primary source (the EIP itself, Solidity docs, vendor docs) and cite the link. Use `live-source-verification` if it's installed. Never cite a URL from memory.
- **Shared taxonomy.** Tag findings and test coverage with the [OWASP Smart Contract Top 10: 2026](https://scs.owasp.org/sctop10/) categories (SC01 Access Control, SC02 Business Logic, SC03 Price Oracle Manipulation, SC04 Flash Loan–Facilitated Attacks, SC05 Lack of Input Validation, SC06 Unchecked External Calls, SC07 Arithmetic Errors, SC08 Reentrancy, SC09 Integer Overflow/Underflow, SC10 Proxy & Upgradeability), so all reports from these skills line up.
- **Compiler and EVM target.** Pin the Solidity version and set `evmVersion` explicitly for every deployment chain. Since 0.8.31 the compiler targets Osaka by default ([Solidity 0.8.31](https://www.soliditylang.org/blog/2025/12/03/solidity-0.8.31-release-announcement/)), and an L2 that hasn't adopted the same fork may reject the newer opcodes. Check the [release notes](https://soliditylang.org/blog/category/releases/) for the latest 0.8.x patch and any security fixes in the version in use.
- **Protocol changes to keep in view (as of November 2026).** Since Pectra ([EIP-7702](https://ethereum.org/roadmap/pectra/7702)), an EOA can run delegated code. Since Fusaka ([EIP-7825](https://blog.ethereum.org/2025/10/21/fusaka-gascap-update)), one transaction is capped at 16,777,216 gas. Glamsterdam, which reprices gas, was expected on mainnet in Q4 2026 with no confirmed date ([ethereum.org](https://ethereum.org/roadmap/glamsterdam/)). Check its status before relying on gas figures.

## Smart Contract QA Workflow

### Step 1: Invariant and Property-Based Test Formulation

1. **Global Invariants (System Conservations)**:
   - Identify conservation laws (e.g., total deposited assets $\ge$ total protocol liabilities).
   - Ensure supply invariants: token supply must never exceed minting limits or diverge from backing reserves.
2. **State Transition Monotonicity**:
   - Ensure states can only transition through valid directed paths (e.g., timestamps only move forward, settled markets cannot be reopened).
3. **Actor and Balance Invariants**:
   - Individual user balances must sum to the collective aggregate balance.
   - Solvency invariants: liquidated positions must never leave bad debt unhandled.

### Step 2: Foundry and Echidna Invariant Harness Design

1. **Handler Contracts Architecture**:
   - Write handler wrappers (`Handler.sol`) that bound random inputs into reasonable ranges (e.g., non-zero amounts, valid actor addresses).
   - Track ghost variables inside handlers (e.g., `ghost_totalDeposited`, `ghost_expectedShares`) to independently verify contract state.
2. **Target Contract & Actor Configuration**:
   - Configure multi-actor testing arrays (`address[] public actors`) to simulate concurrent user interactions.
   - Include actors that run code: contracts and EIP-7702-delegated EOAs that call back and re-enter. Plain EOAs alone miss the reentrancy that "EOA-only" checks no longer block.
   - Define test selectors and call weightings.
3. **Foundry Invariant Test Suites**:
   - Structure tests using `invariant_*` naming conventions in Foundry (`test/invariants/`).
4. **Choose the Engine per Property** (check each tool's current docs and status before adopting it):
   - **Foundry invariant tests**: fast feedback in CI on every change.
   - **Echidna / [Medusa](https://blog.trailofbits.com/2025/02/14/unleashing-medusa-fast-and-scalable-smart-contract-fuzzing/)**: long stateful campaigns. Medusa v1 adds parallel, coverage-guided fuzzing. Reuse the same properties across tools where possible.
   - **[Halmos](https://github.com/a16z/halmos)**: symbolic testing that reuses Foundry tests. Use it for properties that hinge on specific inputs a random search will rarely hit.
   - **[Certora Prover](https://certora.com/blog/certora-goes-open-source)** (open-sourced February 2025): formal proofs of the highest-value invariants (solvency, access control, upgrade safety).
   - Record which engine covers which invariant, so gaps are visible.

### Step 3: Upgradeability and Storage Layout Validation

1. **Storage Layout Extraction**:
   - Inspect and diff contract storage layout slots between implementation versions (`forge inspect <Contract> storageLayout`).
2. **Storage Gap Verification**:
   - Identify which pattern each contract uses. OpenZeppelin Contracts v5 upgradeable contracts use ERC-7201 namespaced storage instead of gaps ([OpenZeppelin upgradeable docs](https://docs.openzeppelin.com/contracts/upgradeable)); your own contracts may use either.
   - **Gaps:** base contracts keep a `uint256[N] private __gap;` and new variables shrink the gap by exactly the slots they take, without shifting descendant storage.
   - **Namespaced (ERC-7201):** the namespace ID and computed slot are unchanged, and struct fields are only appended.
   - Flag any upgrade from an OpenZeppelin v4-based implementation to a v5-based one as unsafe: the two use different storage layouts.
3. **ERC-1967 Slot Validation**:
   - Verify custom implementation, admin, and beacon slots comply with standard keccak256 offsets.

### Step 4: Static Analysis and Tooling Triage

1. Run automated analyzers:
   - **Slither**: Scan for uninitialized state variables, reentrancy vulnerabilities, arbitrary senders, and shadowed variables.
   - **Aderyn**: Scan for modern AST vulnerabilities and missing zero-address validations.
   - Check detector sets and versions in each tool's current release notes. Detectors change between releases.
2. Filter and categorize findings: distinguish true positives from deliberate architectural design choices, documenting waivers explicitly.
3. Classify every finding with the severity scale in `smart-contract-threat-modeler` (Critical / High / Medium / Low / Informational), so reports from all audit skills merge into one matrix.

### Step 5: Gas Profiling and Benchmark Verification

1. Generate gas snapshots (`forge snapshot`).
   - Treat snapshots as valid for the current network rules only. Re-baseline after each upgrade. Glamsterdam reprices gas and was expected on mainnet in Q4 2026 with no confirmed date ([ethereum.org: Glamsterdam](https://ethereum.org/roadmap/glamsterdam/)). Check its status before reporting benchmarks as final.
   - Assert that the worst-case path of every batch or loop function stays under the 16,777,216 per-transaction gas cap (EIP-7825).
2. Verify that gas optimizations (custom errors, Yul blocks, unchecked increments, immutable cache variables) do not introduce safety or overflow regressions.

### Step 6: Downstream Handoff for Implementation (`reviewable-diffs` Integration)

1. Output complete test files and test execution matrices.
2. Provide exact verification commands (e.g., `forge test --match-contract InvariantTest -vvvv`, `echidna . --config echidna.yaml`).

## Deliverable Format Standards

Structure QA and verification reports with these sections:
1. **QA Scope and Test Architecture Overview**
2. **Formal Invariant Specification (Properties, Invariant Functions, Ghost Variables)**
3. **Foundry / Echidna Invariant Test Harness Blueprint**
4. **Storage Layout & Upgrade Safety Report**
5. **Static Analysis Triage Matrix**
6. **Gas Benchmark and Execution Command Summary**

## Gotchas and Common Pitfalls

- **Unbounded Fuzz Inputs Causing Trivial Reverts**: If random inputs always revert due to simple sanity checks, the fuzzer cannot explore deeper state transitions. Use `bound()` or `vm.assume()` properly.
- **Missing Ghost Variable Sync**: Ghost variables must be updated synchronously with state transitions to prevent false invariant failures.
- **Neglecting Fork Tests**: Unit tests on mock environments can hide integration nuances with live tokens (e.g., USDT non-standard returns, WETH unwrapping). Include mainnet fork tests where appropriate.
- **Stipend-Based Reentrancy Assumptions**: Add tests where a `transfer`/`send` recipient uses `TSTORE` within the 2,300-gas stipend. EIP-1153 exempts it from the stipend check ([EIP-1153](https://eips.ethereum.org/EIPS/eip-1153)), so "stipend-safe" code can still be re-entered.
- **Testing on the Wrong EVM Version**: Run fork tests with the `evmVersion` and chain each deployment actually uses. The compiler's default target can be ahead of what an L2 supports.
- **Ignoring Assembly Memory Safety**: Mark inline Yul/assembly as memory-safe only when it strictly conforms to the Solidity memory model. A wrong annotation is undefined behavior, not a safety check. On Solidity 0.8.13+ use `assembly ("memory-safe") { ... }`. The `/// @solidity memory-safe-assembly` comment form is deprecated and scheduled for removal in 0.9.
