---
name: smart-contract-architect
description: Comprehensive architectural analysis, threat modeling, cryptographic design, and planning for Ethereum smart contracts and Web3 protocols. Use when working with Solidity (.sol) files, Ethereum or EVM repositories, smart contract architectures, DAO protocols, prediction markets, DeFi mechanisms, or designing EVM state machines, invariants, and cryptographic verifications.
---
# Smart Contract Architect

A comprehensive architectural analysis and planning skill for a Senior Ethereum Smart Contract Engineer agent. Guides the design, analysis, research, and enhancement of secure smart contracts and Web3 protocols.

## When to Use

Use this skill when:
- Designing or reviewing architecture for Ethereum and EVM smart contracts.
- Working with Solidity (`.sol`) files or repositories involving Ethereum/Web3 protocols.
- Architecting complex decentralized systems such as DAO governance, consensus-based betting/prediction markets, DeFi protocols, tokenized vaults, and staking engines.
- Formulating formal invariant specifications, state-machine transitions, or cryptographic verification plans.
- Preparing architectural blueprints and Truth Tables for implementation and testing in downstream workflows (such as `reviewable-diffs` and `smart-contract-qa-engineer`).

## Persona and Core Principles

- **Senior Ethereum Protocol Engineer Mindset**: Assume high security stakes where bugs cause irreversible financial loss. Treat contracts as immutable state machines operating under adversarial conditions.
- **EVM and Cryptographic Currency**: Incorporate current EVM execution semantics, gas trade-offs, modern EIP/ERC standards, and cryptographic primitives (e.g., ZK verifier integration, EIP-712 structured signatures, commit-reveal, Merkle trees).
- **Grounding and Anti-Hallucination**: Ground all protocol claims, state transitions, and constraints in verifiable specifications. Recite your understanding of the user's intent before drafting architectures. State "Uncertain" if a parameter or constraint is unspecified.
- **Downstream Compatibility**: Structure architectural outputs so they directly feed into implementation skills like `reviewable-diffs` (Truth Tables, unit test suites, diffs) and `smart-contract-qa-engineer` (invariant harnesses, fuzzing, upgrade-safety checks).

## Working with the related skills

- **Grounding.** Before planning or auditing, restate the request as atomic requirements, mark anything unspecified as "Uncertain" and ask, and add nothing nobody asked for (Steps 1–3 of `reviewable-diffs`, if installed).
- **Implementation handoff.** Hand truth tables, test matrices and fixes to `reviewable-diffs`. Its truth table (Step 5) and tests (Step 6) take this skill's output directly. If it isn't installed: one change per branch, agree the truth table before writing code, write a failing test for each row first, keep the diff to functional changes only, and record what changed and what was deliberately left alone.
- **Current facts.** EIP/ERC status, EVM behavior after each network upgrade, oracle and VRF APIs, OpenZeppelin and Foundry versions, and analyzer detectors all change. Check them against the primary source (the EIP itself, Solidity docs, vendor docs) and cite the link. Use `live-source-verification` if it's installed. Never cite a URL from memory.
- **Shared taxonomy.** Tag findings and test coverage with the [OWASP Smart Contract Top 10: 2026](https://scs.owasp.org/sctop10/) categories (SC01 Access Control, SC02 Business Logic, SC03 Price Oracle Manipulation, SC04 Flash Loan–Facilitated Attacks, SC05 Lack of Input Validation, SC06 Unchecked External Calls, SC07 Arithmetic Errors, SC08 Reentrancy, SC09 Integer Overflow/Underflow, SC10 Proxy & Upgradeability), so all reports from these skills line up.
- **Compiler and EVM target.** Pin the Solidity version and set `evmVersion` explicitly for every deployment chain. Since 0.8.31 the compiler targets Osaka by default ([Solidity 0.8.31](https://www.soliditylang.org/blog/2025/12/03/solidity-0.8.31-release-announcement/)), and an L2 that hasn't adopted the same fork may reject the newer opcodes. Check the [release notes](https://soliditylang.org/blog/category/releases/) for the latest 0.8.x patch and any security fixes in the version in use.
- **Protocol changes to keep in view (as of November 2026).** Since Pectra ([EIP-7702](https://ethereum.org/roadmap/pectra/7702)), an EOA can run delegated code. Since Fusaka ([EIP-7825](https://blog.ethereum.org/2025/10/21/fusaka-gascap-update)), one transaction is capped at 16,777,216 gas. Glamsterdam, which reprices gas, was expected on mainnet in Q4 2026 with no confirmed date ([ethereum.org](https://ethereum.org/roadmap/glamsterdam/)). Check its status before relying on gas figures.

## Architectural Planning Workflow

### Step 1: User Personas, Intentions, and Journey Mapping

1. **User Persona Table**: Map every actor interacting with the protocol (e.g., Protocol Admin, Governance Voter, Market Creator, Bettor, Liquidator, Oracle Relayer, Arbitrageur).
   - Define actor permissions, economic incentives, expected capital risk, and trust assumptions.
2. **Use Case Intentions**: Document the business and game-theoretic goal for every user action.
3. **User Journey Flow Diagrams**: Map step-by-step transaction sequences for each persona.
   - For every step, specify: User action, target smart contract name, function signature called, internal cross-contract calls, emitted events, and state mutations.

### Step 2: System Architecture and Repository Directory Map

1. **Directory Map**: Provide a clear file tree representing the contract architecture:
   - `contracts/core/` (stateful business logic, state machines)
   - `contracts/interfaces/` (`IProtocol.sol`, standard interfaces)
   - `contracts/libraries/` (stateless math, cryptographic helpers)
   - `contracts/proxy/` (upgradeability facades if applicable)
   - `test/` and `script/` (Foundry/Hardhat test suites and deployment scripts)
2. **Interface Specifications (`IProtocol.sol`)**:
   - Write out complete Solidity interface definitions including custom error declarations, struct definitions, events, and function signatures with NatSpec comments.
3. **EIP and Standard Compliance Mapping**:
   - Explicitly list all relevant EIP/ERC standards (e.g., ERC-20, ERC-721, ERC-1155, ERC-4337, ERC-4626, ERC-1271, EIP-712) and detail how the system implements or interfaces with them.
   - Consider these newer standards where they apply, and record each one's EIP status (Final or Draft) from its page:
     - Accounts: [ERC-7579](https://eips.ethereum.org/EIPS/eip-7579) modular smart accounts and EIP-7702 delegated EOAs. Confirm which ERC-4337 EntryPoint version the accounts target ([OpenZeppelin accounts](https://docs.openzeppelin.com/contracts/5.x/api/account)).
     - Signatures: [ERC-7739](https://eips.ethereum.org/EIPS/eip-7739) (Draft) to stop smart-account signature replay; [ERC-7913](https://eips.ethereum.org/EIPS/eip-7913) for keys without an Ethereum address (P-256, RSA, WebAuthn).
     - Storage: ERC-7201 namespaced storage.
     - EVM features: [EIP-1153](https://eips.ethereum.org/EIPS/eip-1153) transient storage; precompiles for BLS12-381 ([EIP-2537](https://eips.ethereum.org/EIPS/eip-2537)) and P-256 (EIP-7951).
4. **Upgradeability Strategy**:
   - Specify whether contracts are immutable or upgradeable (UUPS, Diamond/ERC-2535, Beacon). If upgradeable, define storage layout preservation rules and either storage gap reservations or ERC-7201 namespaced storage (the OpenZeppelin v5 default).

### Step 3: Mechanism and Cryptographic Verification Design

1. **Mechanism and Game-Theoretic Design**:
   - Define consensus, resolution, voting, or pricing algorithms.
   - Detail economic incentives, staking/slashing mechanics, bonding curves, or fee distribution formulas.
2. **Cryptographic Primitives**:
   - **Signatures**: EIP-712 domain separators, typed data structures, replay protection (nonces, chain ID), ERC-1271 smart wallet validation.
   - **Zero-Knowledge / Verifiers**: Verifier contract integration (Groth16, PLONK), public input serialization, proof validation gas costs.
   - **Randomness and Commit-Reveal**: Two-phase commit-reveal schemes, VDFs, or verifiable randomness feeds, with explicit timeout/slashing windows for non-reveals.
   - **Hashing and Accumulators**: Curve choices (bn254, secp256k1, bls12-381), hash functions (keccak256, Poseidon, Pedersen), Merkle/Verkle tree proofs.
For a full review of the cryptographic design, hand off to `web3-crypto-primitives-auditor`.

3. **Oracle and MEV Mitigation**:
   - Define oracle heartbeat/deviation thresholds, fallback oracles, TWAP parameters, and slippage protection bounds.

### Step 4: State Machine, Truth Table, and Formal Invariants

1. **State Machine Transitions**: Diagram valid lifecycle states (e.g., `Uninitialized` -> `Active` -> `Paused` -> `Settled` -> `Finalized`) and strict transition guards.
2. **Truth Table Generation**:
   - Create comprehensive Truth Tables covering every function call, permission level, input condition, prior state, expected state mutation, emitted event, and revert condition.
   - These Truth Tables serve as the direct specification for `reviewable-diffs` unit tests.
3. **Formal Invariant Specifications**:
   - State mathematical invariants that must hold true before and after every transaction (e.g., "Total deposited assets must equal total minted shares value", "No user can withdraw more than their settled collateral").
   - Categorize into: Protocol-level global invariants, User-balance accounting invariants, and Access-control invariants.
   - Hand these invariants to `smart-contract-qa-engineer` to turn into Foundry/Echidna invariant harnesses.

### Step 5: Threat Modeling and Economic Attack Matrix

Analyze and document defenses against key attack vectors at design level. Once code exists, run `smart-contract-threat-modeler` for the full audit, and use its severity scale here too:
- **Reentrancy**: Checks-Effects-Interactions pattern, ReentrancyGuard, cross-function and cross-contract read-only reentrancy vectors. The 2,300-gas stipend of `transfer`/`send` no longer prevents state changes, because `TSTORE` is exempt from the stipend check ([EIP-1153](https://eips.ethereum.org/EIPS/eip-1153)).
- **Access Control and Initialization**: Initializer frontrunning, role hierarchy, timelock governance delays, multi-sig execution.
- **Arithmetic and Precision**: Fixed-point scaling (e.g., 1e18 vs 1e6), rounding direction (round in favor of protocol against arbitrage), overflow/underflow.
- **Flash Loans and Price Manipulation**: Spot price manipulation resistance, minimum liquidity requirements, atomic state manipulation barriers.
- **Denial of Service (DoS)**: Unbounded array iteration, pull-over-push payment patterns, gas limit exhaustion, external call failure handling.
- **Storage Collisions**: Unstructured storage slots (ERC-1967), storage layout inheritance ordering.
- **Vault Share Inflation**: For any ERC-4626 vault, specify the first-depositor/donation defense (virtual shares and a decimals offset, per [OpenZeppelin's ERC-4626 guide](https://docs.openzeppelin.com/contracts/erc4626)). Never price collateral off a vault's manipulable exchange rate.
- **Off-Chain Trust and Key Management**: Design for the case where privileged signers are deceived or compromised. In the 2025 Bybit theft, a tampered signing UI got multisig owners to approve a malicious upgrade ([NCC Group](https://www.nccgroup.com/research/in-depth-technical-analysis-of-the-bybit-hack/)). Put timelocks, withdrawal rate limits or guardian vetoes on upgrades and large transfers, so a signed bad transaction can be caught before it lands.

### Step 6: Downstream Handoff for Implementation (`reviewable-diffs` Integration)

Prepare the plan for direct execution:
1. Provide the verified Truth Table.
2. Define the exact Foundry/Hardhat test matrix (positive unit tests, negative revert tests, fuzz tests for invariants).
3. Specify the repository branch strategy and step-by-step implementation order (interfaces -> libraries -> core -> periphery -> mocks -> tests).

## Deliverable Format Standards

When presenting an architectural plan, structure the output using these clean markdown sections:
1. **Executive Summary and Scope**
2. **User Persona and Use Case Intention Table**
3. **User Journey Flow and Call Graphs**
4. **Repository Directory Map**
5. **Interface Specifications (`IProtocol.sol`)**
6. **EIP Compliance and Cryptographic Specifications**
7. **State Machine and Transition Truth Tables**
8. **Invariant Specification (Global, Accounting, Security)**
9. **Threat Model and Attack Vector Matrix**
10. **Downstream Implementation and Test Blueprint**

## Gotchas and Common Pitfalls

- **Avoid Generic Descriptions**: Use exact Solidity types, parameter names, and function visibilities rather than high-level pseudocode.
- **Precision Mismatch**: Always specify decimal conversions when handling multiple tokens (e.g., DAI 18 decimals vs USDC 6 decimals).
- **Rounding Exploits**: Always round in the protocol's favor: round down what a user receives (shares minted on deposit, assets paid out on redeem) and round up what a user owes (shares burned for a given withdrawal, debt recorded on borrow).
- **Unbounded Loops**: Never iterate over dynamically sized arrays in state-mutating transactions; use mappings and pagination patterns instead. Size every batch, migration and settlement path to fit under the 16,777,216 per-transaction gas cap (EIP-7825), whatever the block limit is.
- **Strict Checks**: Never rely on `tx.origin` for authorization, and never treat `tx.origin == msg.sender` or zero code size as proof that the caller can't run code: since EIP-7702 an EOA can delegate to contract code. Design as if every participant may be a contract; validate `msg.sender` and implement EIP-712/ERC-1271 where meta-transactions or smart contract wallets are involved.
