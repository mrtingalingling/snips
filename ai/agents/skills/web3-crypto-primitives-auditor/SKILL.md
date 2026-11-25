---
name: web3-crypto-primitives-auditor
description: Security audit, verification, and cryptographic review for Web3 systems. Use when validating zero-knowledge verifier integrations, EIP-712/ERC-1271 signature schemes, commit-reveal mechanisms, hash functions, elliptic curves, randomness beacons, and off-chain to on-chain cryptographic boundaries.
---
# Web3 Crypto Primitives Auditor

A specialized cryptographic verification and security audit skill for an Applied Cryptography and Web3 QA Engineer agent. Focuses on auditing cryptographic boundaries, signature schemes, zero-knowledge verifier contracts, randomness protocols, and mathematical assumptions on Ethereum.

## When to Use

Use this skill when:
- Auditing smart contracts that integrate zero-knowledge (ZK) SNARK/STARK proof verifiers (e.g., Groth16, PLONK, UltraHonk).
- Reviewing digital signature validation (ECDSA, EIP-712 typed structured data, ERC-1271 smart contract wallet signatures, ERC-6492, EIP-2098 compact signatures).
- Evaluating commit-reveal protocols, Verifiable Random Functions (VRF), or Verifiable Delay Functions (VDF).
- Validating hashing schemes, Merkle tree implementations, and curve parameter choices (e.g., bn254, secp256k1, bls12-381).
- Analyzing off-chain to on-chain data passing, serialization boundaries, and cryptographic proof verification gas costs.

## Persona and Core Principles

- **Applied Cryptographer Mindset**: Cryptographic failures are subtle, non-reverting, and critical. Small parameter misconfigurations or missing field bounds cause complete protocol insolvency.
- **Mathematical Rigor**: Verify that all curve operations, finite field constraints, and hashing operations strictly satisfy their formal security definitions.
- **Malleability and Replay Defense**: Treat every signature and proof as malleable until explicitly bounded by domain separators, chain IDs, nonces, and scalar field ranges.
- **Downstream QA Integration**: Generate deterministic mathematical edge-case tests and negative test vectors ready for execution in `reviewable-diffs`.

## Working with the related skills

- **Grounding.** Before planning or auditing, restate the request as atomic requirements, mark anything unspecified as "Uncertain" and ask, and add nothing nobody asked for (Steps 1–3 of `reviewable-diffs`, if installed).
- **Implementation handoff.** Hand truth tables, test matrices and fixes to `reviewable-diffs`. Its truth table (Step 5) and tests (Step 6) take this skill's output directly. If it isn't installed: one change per branch, agree the truth table before writing code, write a failing test for each row first, keep the diff to functional changes only, and record what changed and what was deliberately left alone.
- **Current facts.** EIP/ERC status, EVM behavior after each network upgrade, oracle and VRF APIs, OpenZeppelin and Foundry versions, and analyzer detectors all change. Check them against the primary source (the EIP itself, Solidity docs, vendor docs) and cite the link. Use `live-source-verification` if it's installed. Never cite a URL from memory.
- **Shared taxonomy.** Tag findings and test coverage with the [OWASP Smart Contract Top 10: 2026](https://scs.owasp.org/sctop10/) categories (SC01 Access Control, SC02 Business Logic, SC03 Price Oracle Manipulation, SC04 Flash Loan–Facilitated Attacks, SC05 Lack of Input Validation, SC06 Unchecked External Calls, SC07 Arithmetic Errors, SC08 Reentrancy, SC09 Integer Overflow/Underflow, SC10 Proxy & Upgradeability), so all reports from these skills line up.
- **Compiler and EVM target.** Pin the Solidity version and set `evmVersion` explicitly for every deployment chain. Since 0.8.31 the compiler targets Osaka by default ([Solidity 0.8.31](https://www.soliditylang.org/blog/2025/12/03/solidity-0.8.31-release-announcement/)), and an L2 that hasn't adopted the same fork may reject the newer opcodes. Check the [release notes](https://soliditylang.org/blog/category/releases/) for the latest 0.8.x patch and any security fixes in the version in use.
- **Protocol changes to keep in view (as of November 2026).** Since Pectra ([EIP-7702](https://ethereum.org/roadmap/pectra/7702)), an EOA can run delegated code. Since Fusaka ([EIP-7825](https://blog.ethereum.org/2025/10/21/fusaka-gascap-update)), one transaction is capped at 16,777,216 gas. Glamsterdam, which reprices gas, was expected on mainnet in Q4 2026 with no confirmed date ([ethereum.org](https://ethereum.org/roadmap/glamsterdam/)). Check its status before relying on gas figures.

## Cryptographic Audit Workflow

### Step 1: Signature Schemes and Authentication Auditing

1. **ECDSA Security**:
   - Verify `ecrecover` return value: ensure address(0) is rejected on invalid signatures.
   - Enforce low-s value constraints (EIP-2) to prevent signature malleability.
   - Never use a signature as a unique identifier (for example, as a "used" key). Use nonces or invalidate the signed hash. OpenZeppelin marks its malleability restriction as deprecated, to be removed in v6.0 ([OpenZeppelin cryptography](https://docs.openzeppelin.com/contracts/api/utils/cryptography)).
   - Ensure explicit replay protection: `nonce` tracking, `chainId` inclusion in hash digest, and contract address binding.
4. **EIP-7702 Authorizations** (wallets, delegation targets, relayers):
   - An authorization signed with `chain_id` 0 is valid on every chain ([ethereum.org: EIP-7702](https://ethereum.org/roadmap/pectra/7702)). Flag any flow that produces or accepts one.
   - The EOA's own key still controls the account after delegation, so a delegated multisig or policy contract can be bypassed by that key. Don't model it as a true smart account.
   - Delegation-target contracts must protect their initialization and storage against reuse by other delegating accounts.
2. **EIP-712 Verification**:
   - Validate Domain Separator computation and ensure `verifyingContract` is bound to the deployed instance.
   - Verify type hashes and encoded struct packing match EIP-712 specification exactly.
3. **P-256 / Passkey (WebAuthn) Signatures**:
   - Fusaka added a secp256r1 verification precompile on mainnet (EIP-7951) ([ethereum.org: Fusaka](https://ethereum.org/roadmap/fusaka/)). Other chains may expose it under RIP-7212, at another address, or not at all. Confirm per deployment chain.
   - Require the precompile's exact 32-byte success value. A failed verification returns empty data, and so does a call to an address where no precompile exists, so treating empty data as success accepts every signature.
   - Check WebAuthn wrapping: challenge binding, `authenticatorData` flags and `clientDataJSON` parsing.
5. **ERC-1271 / Smart Wallet Signatures**:
   - Verify proper calling of `isValidSignature(bytes32,bytes)`.
   - Prevent cross-contract signature replays across different smart contract wallet instances. When one EOA owns several smart accounts, a plain `isValidSignature` can be replayed across them. Check for a defensive rehashing scheme such as [ERC-7739](https://eips.ethereum.org/EIPS/eip-7739) (Draft).
   - For keys without an Ethereum address (P-256, RSA, WebAuthn), check [ERC-7913](https://eips.ethereum.org/EIPS/eip-7913) verifier usage: the signer encoding (verifier address + key), that verifier contracts are trusted, and how a key shorter than 20 bytes is rejected.
   - For counterfactual (not-yet-deployed) accounts, check ERC-6492 handling and the safety of its factory call.
6. **BLS Signatures (BLS12-381)**:
   - Pectra added BLS12-381 precompiles ([EIP-2537](https://eips.ethereum.org/EIPS/eip-2537)). Take the precompile addresses, input encodings and gas schedule from the EIP itself.
   - Check which operations perform subgroup checks under the EIP, and validate points wherever they don't. Also check hash-to-curve domain separation tags, rogue-key defenses (proof of possession) for aggregated signatures, and that precompile failures are handled rather than read as empty output.

### Step 2: Zero-Knowledge and Verifier Contract Assurance

1. **Public Input Validation**:
   - Verify that all public inputs passed to the verifier are strictly reduced modulo the scalar field size ($r$).
   - Check input serialization order between off-chain proof generation (Circom, Noir, Halo2) and the on-chain Solidity verifier.
2. **Under-Constrained Verification**:
   - Verify that the on-chain contract enforces all necessary domain checks that are not guaranteed by the circuit itself.
   - Check for over-constrained circuits too. They reject valid witnesses and can deny service to honest provers ([zkFuzz](https://arxiv.org/abs/2504.11961v2)).
   - Check for proof replay vulnerabilities (e.g., re-submitting a valid proof for a different user address or different state root).
3. **Verification Gas**:
   - Confirm that a verification, and any batch of verifications in one call, fits under EIP-7825's 16,777,216 per-transaction gas cap with margin.
4. **Circuit Tooling** (when circuit source is in scope):
   - Run a static pass ([Circomspect, ZKAP, halo2-analyzer](https://blog.zksecurity.xyz/posts/zksecurity-tools/index.html)), then a solver-backed under-constraint check ([Picus](https://veridise.com/tools/zk-tool/)) or circuit fuzzing ([zkFuzz](https://arxiv.org/abs/2504.11961v2)).
   - Treat static hits as leads, and confirm each with a concrete witness that the circuit accepts but shouldn't.

### Step 3: Commit-Reveal and Randomness Protocols

1. **Commit-Reveal Boundaries**:
   - Verify commitment binding and hiding properties ($H(value \parallel secret \parallel msg.sender)$).
   - Ensure the commitment digest binds the specific committer address to prevent frontrunning/copying commitments.
   - Verify timeout and slashing parameters: enforce deterministic default outcomes if a party fails to reveal.
2. **Randomness Feeds**:
   - Check VRF subscription funding and fulfill callback gas limits.
   - Ensure random seeds cannot be influenced by miner/validator timestamps or blockhash manipulations.

### Step 4: Hash Functions and Merkle Accumulators

1. **Pre-image and Collision Resistance**:
   - Check for Merkle second-preimage attacks: if a leaf preimage is 64 bytes, it can be passed off as an internal node (the concatenation of two child hashes). `abi.encode` alone does not fix this (two `bytes32` values still encode to 64 bytes). Double-hash leaves or domain-separate leaf and node hashing.
   - Check `abi.encodePacked` with more than one dynamic-type argument (see Gotchas).
   - Verify zero-value leaf initialization in sparse Merkle trees.
2. **Poseidon and Pedersen Hashing**:
   - Verify correct round constants and MDS matrix configurations when evaluating ZK-friendly hash implementations.

### Step 5: Severity, Downstream Handoff and Edge-Case Test Plan

Classify every finding with the severity scale in `smart-contract-threat-modeler` (Critical / High / Medium / Low / Informational), so reports from all audit skills merge into one matrix.


1. Formulate negative test vectors:
   - Zero-address signatures and signature malleability variants.
   - Scalar field overflow public inputs.
   - Expired timestamps, duplicate nonces, and mutated proofs.
2. Deliver the test specification ready for implementation under `reviewable-diffs`. Property-style vectors (e.g. "no signature validates twice") go to `smart-contract-qa-engineer` as fuzz/invariant targets.

## Deliverable Format Standards

Structure cryptographic audit reports with these sections:
1. **Cryptographic Scope and Boundary Overview**
2. **Signature and Authentication Security Review**
3. **Zero-Knowledge Circuit and Verifier Analysis**
4. **Commit-Reveal, Randomness, and Timing Assessment**
5. **Hash and Accumulator Collision Analysis**
6. **Findings by Severity**
7. **Cryptographic Edge-Case Test Matrix**

## Gotchas and Common Pitfalls

- **Do Not Use `abi.encodePacked` on Dynamic Types**: Packing multiple variable-length arrays or strings causes hash collisions.
- **Missing Modulo Reductions**: If public inputs in ZK verifiers are not checked to be $< r$, attackers can forge valid proofs with aliased field elements.
- **Unchecked ERC-1271 Magic Value**: Always compare the returned 4-byte selector directly against `0x1626ba7e`.
- **Commitment Frontrunning**: Always include `msg.sender` in the committed hash digest; otherwise an attacker can frontrun and submit the same commitment.
