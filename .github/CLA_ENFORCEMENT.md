# Contributor License Agreement Enforcement

## Status

**Not enabled.** The repository's
[Contributor License Agreement draft](../CONTRIBUTOR_LICENSE_AGREEMENT-DRAFT.md)
is not a binding agreement. Do not configure a required CLA check until legal
has approved both the final text and the electronic acceptance process.

## Activation checklist

1. Legal approves a versioned final CLA, including the governing law,
   contributor representations, rights grant, and acceptance method.
2. Privacy reviews the selected CLA provider, its data-processing terms, data
   locations, retention, and the public/private location of signature records.
3. Install the approved CLA GitHub App or Action and configure it to verify
   every pull-request author and commit author against that version.
4. Protect `master` with a ruleset that requires pull requests, an approving
   review, and the CLA check from that specific GitHub App.
5. Verify the gate using a test pull request from an account without a signed
   CLA and one with a signed CLA. Document approved bot exceptions.

The check is a merge gate only. Maintainers remain responsible for confirming
that the contributor and any employer have authority to grant the rights.
