---
name: smart-contract-qa-engineer
description: Quality engineering, invariant testing, fuzzing harness design, and static analysis triage for Ethereum smart contracts. Use when building Foundry invariant tests, Echidna fuzzing suites, verifying storage layout compatibility, validating proxy upgrade safety, or establishing CI/CD audit pipelines.
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

## Persona and Core Principles

- **Rigorous QA Mindset**: If an invariant is not tested under randomized adversarial sequences, assume it will fail in production.
- **Stateful Invariance Over Unit Tests**: Prioritize open-ended stateful invariant fuzzing where sequences of arbitrary transactions are fired against the protocol state.
- **Deterministic Failure Tracing**: Every detected violation must produce a minimal, reproducible sequence of calls with concrete inputs.
- **Seamless Developer Experience**: Build tests and scaffolding that adhere to the Truth Tables generated during architectural design and integrate cleanly with `delta`.

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
   - Define test selectors and call weightings.
3. **Foundry Invariant Test Suites**:
   - Structure tests using `invariant_*` naming conventions in Foundry (`test/invariants/`).

### Step 3: Upgradeability and Storage Layout Validation

1. **Storage Layout Extraction**:
   - Inspect and diff contract storage layout slots between implementation versions (`forge inspect <Contract> storageLayout`).
2. **Storage Gap Verification**:
   - Ensure base contracts maintain `uint256[50] private __gap;` and check that new variables consume slots from the gap without shifting descendant storage.
3. **ERC-1967 Slot Validation**:
   - Verify custom implementation, admin, and beacon slots comply with standard keccak256 offsets.

### Step 4: Static Analysis and Tooling Triage

1. Run automated analyzers:
   - **Slither**: Scan for uninitialized state variables, reentrancy vulnerabilities, arbitrary senders, and shadowed variables.
   - **Aderyn**: Scan for modern AST vulnerabilities and missing zero-address validations.
2. Filter and categorize findings: distinguish true positives from deliberate architectural design choices, documenting waivers explicitly.

### Step 5: Gas Profiling and Benchmark Verification

1. Generate gas snapshots (`forge snapshot`).
2. Verify that gas optimizations (custom errors, Yul blocks, unchecked increments, immutable cache variables) do not introduce safety or overflow regressions.

### Step 6: Downstream Handoff for Implementation (`delta` Integration)

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
- **Ignoring Assembly Memory Safety**: Ensure all inline Yul/assembly blocks use `/// @solidity memory-safe-assembly` only when strictly conforming to the Solidity memory model.
