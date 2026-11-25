---
name: smart-contract-threat-modeler
description: Comprehensive threat modeling, vulnerability surface analysis, and economic exploit simulation for Ethereum and EVM smart contracts. Use when auditing Solidity contracts, evaluating protocol attack surfaces, modeling flash loan or oracle manipulation vectors, analyzing MEV risks, or preparing audit-readiness matrices.
---
# Smart Contract Threat Modeler

A specialized threat modeling and security analysis skill for an Audit and Quality Engineer agent. Focuses on identifying protocol vulnerabilities, economic exploit vectors, access control flaws, and preparing smart contracts for rigorous security audits.

## When to Use

Use this skill when:
- Auditing Solidity smart contracts and EVM protocol repositories.
- Evaluating threat models, attack surfaces, and permission hierarchies.
- Simulating economic exploit scenarios such as flash-loan attacks, oracle manipulation, and sandwich/MEV vectors.
- Analyzing reentrancy, state machine race conditions, and cross-contract interdependencies.
- Generating structured audit reports, severity matrices, and actionable remediation blueprints.

## Persona and Core Principles

- **Security Auditor Mindset**: Assume all external inputs are adversarial and public mempools are monitored by MEV searchers.
- **Economic Realism**: Treat economic incentives as security boundaries. If an action is profitable after gas, flash loan fees, and slippage, assume it will occur.
- **Evidence-Based Grounding**: Ground all identified vulnerabilities in specific code execution paths, state mutations, or mathematical models. Avoid vague warnings.
- **Actionable Remediation**: Provide concrete, non-breaking remediation code and test reproduction guidelines that feed directly into implementation skills like `reviewable-diffs`.

## Working with the related skills

- **Grounding.** Before planning or auditing, restate the request as atomic requirements, mark anything unspecified as "Uncertain" and ask, and add nothing nobody asked for (Steps 1–3 of `reviewable-diffs`, if installed).
- **Implementation handoff.** Hand truth tables, test matrices and fixes to `reviewable-diffs`. Its truth table (Step 5) and tests (Step 6) take this skill's output directly. If it isn't installed: one change per branch, agree the truth table before writing code, write a failing test for each row first, keep the diff to functional changes only, and record what changed and what was deliberately left alone.
- **Current facts.** EIP/ERC status, EVM behavior after each network upgrade, oracle and VRF APIs, OpenZeppelin and Foundry versions, and analyzer detectors all change. Check them against the primary source (the EIP itself, Solidity docs, vendor docs) and cite the link. Use `live-source-verification` if it's installed. Never cite a URL from memory.
- **Shared taxonomy.** Tag findings and test coverage with the [OWASP Smart Contract Top 10: 2026](https://scs.owasp.org/sctop10/) categories (SC01 Access Control, SC02 Business Logic, SC03 Price Oracle Manipulation, SC04 Flash Loan–Facilitated Attacks, SC05 Lack of Input Validation, SC06 Unchecked External Calls, SC07 Arithmetic Errors, SC08 Reentrancy, SC09 Integer Overflow/Underflow, SC10 Proxy & Upgradeability), so all reports from these skills line up.
- **Compiler and EVM target.** Pin the Solidity version and set `evmVersion` explicitly for every deployment chain. Since 0.8.31 the compiler targets Osaka by default ([Solidity 0.8.31](https://www.soliditylang.org/blog/2025/12/03/solidity-0.8.31-release-announcement/)), and an L2 that hasn't adopted the same fork may reject the newer opcodes. Check the [release notes](https://soliditylang.org/blog/category/releases/) for the latest 0.8.x patch and any security fixes in the version in use.
- **Protocol changes to keep in view (as of November 2026).** Since Pectra ([EIP-7702](https://ethereum.org/roadmap/pectra/7702)), an EOA can run delegated code. Since Fusaka ([EIP-7825](https://blog.ethereum.org/2025/10/21/fusaka-gascap-update)), one transaction is capped at 16,777,216 gas. Glamsterdam, which reprices gas, was expected on mainnet in Q4 2026 with no confirmed date ([ethereum.org](https://ethereum.org/roadmap/glamsterdam/)). Check its status before relying on gas figures.

## Threat Modeling and Audit Workflow

### Step 1: Attack Surface and Entry Point Mapping

1. Map all external and public function entry points.
2. Identify untrusted external calls (`call`, `delegatecall`, `staticcall`, ERC-777/ERC-1155 token hooks).
3. Review privileged roles (Owner, Admin, Pauser, Operator) and privilege escalation vectors.
4. Check contract initialization, constructor parameters, and factory deployment patterns.
5. Find every "caller is an EOA" assumption (`tx.origin == msg.sender`, `extcodesize == 0`, `code.length == 0`). Since EIP-7702, an EOA can delegate to code, receive callbacks and re-enter, so treat these as broken guards against reentrancy, flash loans or bots ([ethereum.org: EIP-7702](https://ethereum.org/roadmap/pectra/7702)).
6. Map the off-chain and operational surface: privileged keys held by off-chain services, the multisig signing flow, deployment and upgrade pipelines, and front-ends that build transactions. In the 2025 Bybit theft, a tampered signing UI led signers to approve a malicious implementation upgrade ([NCC Group](https://www.nccgroup.com/research/in-depth-technical-analysis-of-the-bybit-hack/)). For each privileged path, find out whether a compromised or deceived quorum can upgrade or drain instantly, or whether a timelock or rate limit bounds the damage.

### Step 2: Economic and Oracle Exploit Simulation

1. **Oracle Manipulation**:
   - Evaluate dependencies on spot Automated Market Maker (AMM) prices.
   - Assess Time-Weighted Average Price (TWAP) window lengths and manipulation costs across blocks.
   - Verify Chainlink staleness checks on `updatedAt` against each feed's heartbeat, `answer > 0`, and min/max circuit-breaker bounds. On L2s, also check the sequencer uptime feed. Don't rely on `answeredInRound`: Chainlink marks it deprecated ([Data Feeds API Reference](https://docs.chain.link/data-feeds/api-reference)). Re-check the current docs during each audit.
2. **Flash Loan Resilience**:
   - Model protocol solvency if an attacker accesses infinite flash-loan liquidity within a single transaction.
   - Check whether deposits, borrowings, or liquidations can be manipulated atomically.
   - Treat any ERC-4626 vault exchange rate used as a price as manipulable by donation. Check first-depositor/inflation defenses (virtual shares and decimals offset, per [OpenZeppelin](https://docs.openzeppelin.com/contracts/erc4626)) on every vault the protocol issues or accepts.
3. **MEV and Frontrunning**:
   - Check for transaction ordering vulnerabilities in liquidations, auctions, arbitrage, and governance execution.

### Step 3: Reentrancy and State Mutation Verification

1. Check for violation of the Checks-Effects-Interactions (CEI) pattern across all state-mutating functions.
2. Identify cross-function reentrancy vectors where sharing state variables creates inconsistent intermediate states.
3. Identify cross-contract read-only reentrancy where external contracts query a manipulated balance before settlement.
4. Check transient storage (EIP-1153) usage. `TSTORE` is not subject to the 2,300-gas stipend check, so `transfer`/`send` callbacks can change state and re-enter ([EIP-1153](https://eips.ethereum.org/EIPS/eip-1153); [ChainSecurity](https://chainsecurity.com/tstore-low-gas-reentrancy/)). Also check that transient locks and flash-accounting deltas are cleared or settled correctly within the transaction, including when several calls are batched together.

### Step 4: Access Control and Governance Threat Modeling

1. **Governance Exploits**:
   - Check if voting power can be flash-loaned or flash-minted within the same block or proposal window.
   - Validate timelock delays to ensure users can exit before malicious parameter updates take effect.
2. **Signature & Authentication Flaws**:
   - Verify proper nonce tracking for signature replays.
   - Check permit flows for front-running and nonce griefing: a third party can submit a user's `permit` first and make the user's bundled call revert. Wrap `permit` calls in `try`/`catch` and fall back to checking the existing allowance.
   - Ensure invalid signatures revert rather than returning zero-addresses.
   - For a deeper review of signatures, ZK verifiers, commit-reveal or randomness, hand off to `web3-crypto-primitives-auditor`. Its findings use this skill's severity scale.

### Step 5: Severity Classification and Vulnerability Matrix

Classify every finding using standard severity definitions:
- **Critical**: Direct loss of funds, permanent contract freeze, or unauthorized minting/draining without prerequisites.
- **High**: Loss of funds under specific market conditions, temporary freeze of user assets, or serious griefing.
- **Medium**: Broken protocol functionality without direct fund theft, oracle deviation vulnerability, or governance blockage.
- **Low**: Minor logic bugs, incorrect event emissions, or missing validation with low impact.
- **Informational / Gas**: Code style improvements, gas optimizations, or dead code.

Also tag every finding with its [OWASP Smart Contract Top 10: 2026](https://scs.owasp.org/sctop10/) category (SC01–SC10). OWASP's honourable mentions (permit front-running and nonce DoS, MEV, cross-chain MEV, gas/loop/state-bloat DoS, governance attacks, and cryptographic signature-scheme flaws) are in scope here too. Many real exploits chain several categories, so report the chain as well as each link.

### Step 6: Remediation and Downstream Integration

1. Write out the precise recommended code fix.
2. Specify the exact test scenario required in Foundry/Hardhat to prove vulnerability reproduction and subsequent fix validation for `reviewable-diffs`.
3. Where the root cause is a broken invariant (accounting, solvency, access control), also hand it to `smart-contract-qa-engineer` as an invariant test, so the whole class of bug stays caught, not just this instance.

## Deliverable Format Standards

Structure threat modeling reports with these sections:
1. **Audit Scope and Contract Summary**
2. **Attack Surface Mapping and Trust Assumptions**
3. **Vulnerability Matrix (Summary of Findings by Severity)**
4. **Detailed Findings (Root Cause, Exploit Scenario, Code Snippet, Remediation)**
5. **Economic Attack Vector Simulation (Flash Loans, Oracles, MEV)**
6. **Remediation Plan and Test Reproduction Checklist**

## Gotchas and Common Pitfalls

- **Do Not Guess Exploits**: Always trace the complete execution call trace from `msg.sender` to the final state mutation.
- **Beware of Assumptions**: Never assume ERC-20 tokens follow standard behavior; account for fee-on-transfer, rebasing tokens, and non-standard return values (use `SafeERC20`).
- **Pay Attention to Selfdestruct and Forced Ether**: Contracts must never rely strictly on `address(this).balance` for internal accounting. Since [EIP-6780](https://eips.ethereum.org/EIPS/eip-6780), `SELFDESTRUCT` only deletes a contract within its creation transaction; otherwise it just sends the balance. That means ether can still be forced in, and code at a CREATE2 address can still change through same-transaction create-and-destroy. Don't assume code at an address is permanent unless it was deployed with no path to `SELFDESTRUCT`.
- **Check Proxy Storage Clashes**: Ensure upgradeable implementations maintain exact storage variable alignment.
- **Per-Transaction Gas Cap**: Since Fusaka, EIP-7825 caps any single transaction at 16,777,216 gas, even when the block limit is higher ([EF blog](https://blog.ethereum.org/2025/10/21/fusaka-gascap-update)). A liquidation sweep, settlement loop, migration or emergency withdrawal that needs more can never execute. Rate that as permanent DoS, not a gas optimization.
