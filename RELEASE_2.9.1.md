# LDP 2.9.1

Version 22 execution sample: Overview 7.261 s (later 3.561 s), participant database 4.433 s, login 3.048 s. These are individual production calls, not controlled benchmarks.

- Overview reads a smaller projection of submissions and mentoring evidence, and computes counts without building full participant view objects. Eligibility and role/segment scope are preserved.
- Database phase filtering uses the existing narrow enrollment model in light mode.
- Participant login checks roster membership using NIK and segment columns. Admin authentication is unchanged.
- Full detail uses the existing targeted Phase 1/2 read for large sources. Phase assessment caches depend on source and Settings revisions instead of unrelated audit and queue writes.
- B2C contains SLDP and TSHDP. SLDP uses ASHDP grading rules with separate source sheets. TSHDP retains its weights with separate B2C sources. An untagged Phase 3 score is treated as B2B, even when an identical NIK/program/batch exists in B2C.

Validation: 31 backend segment regression cases and a focused eligibility/Phase 3 regression passed in a local Apps Script mock. The same small synthetic Overview case reads 268 vs 278 cells with five reads on both versions. These counts do not measure Google Apps Script latency.

## Release to Apps Script

1. Keep a rollback copy of the current deployment. Copy Code.gs and Index.html from this PR to the correct project and save.
2. Run setupLdpSystem from the editor and wait for setup to finish. Check that B2C source tabs are separate. Import the actual B2C roster with Segmen=B2C and upload the B2C score files through the B2C workspace.
3. Verify Admin and Participant in both segments on a test deployment. For a NIK shared across segments, check that the enrollments and scores stay separate.
4. Compare Overview, Database and Detail timings on the same data before moving the live deployment. Check initial and repeat loads; compare server duration with time until the page can be used. A live 200-user load test has not been performed.
