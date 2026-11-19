"""What several replicas must agree on: the send allowance of each person, and the file links."""

from m365_mcp.state.store import StateStore, StoreUnavailable

__all__ = ["StateStore", "StoreUnavailable"]
