from typing import Any, Callable, Optional, TypeVar
from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations

ContextT = TypeVar("ContextT", default=Any)


class MCPServer(FastMCP[ContextT]):
    """FastMCP subclass that only registers tools present in *enabled_tools*.

    If *enabled_tools* is ``None`` (default) every tool is registered,
    matching the standard FastMCP behaviour.
    """

    def __init__(self, *args, enabled_tools: Optional[list[str]] = None, **kwargs):
        super().__init__(*args, **kwargs)
        self.enabled_tools = enabled_tools

    def tool(
        self,
        func: Callable | None = None,
        *,
        annotations: ToolAnnotations | None = None,
        **kwargs,
    ) -> Callable:
        def decorator(f: Callable) -> Callable:
            tool_name = f.__name__
            if self.enabled_tools is None or tool_name in self.enabled_tools:
                parent_decorator = super(MCPServer, self).tool(
                    annotations=annotations, **kwargs
                )
                return parent_decorator(f)
            return f  # skip registration, return function unmodified

        if func is not None:
            # Called as @mcp.tool (no parentheses)
            return decorator(func)
        # Called as @mcp.tool(...) — return the decorator
        return decorator
