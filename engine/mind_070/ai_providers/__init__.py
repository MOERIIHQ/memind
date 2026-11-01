"""AIProvider transport providers (Phase 2 concrete adapters).

This package holds provider-agnostic *adapters* that implement the AIProvider
contract over a concrete transport (HTTP/HTTPS).  The contract (engine/
mind_070/ai_provider.py) is unchanged; adapters are thin, optional overlays
and never execute operations, write workbooks, publish artefacts, or run
PlanValidator/executor logic.

Default suite is network-free: every transport test mocks urlopen.  Real
network integration is gated behind explicit user configuration and is not
run by the default suite.
"""
