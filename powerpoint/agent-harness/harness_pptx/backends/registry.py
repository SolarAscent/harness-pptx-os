"""Backend registry — auto-detect and instantiate backends."""

from __future__ import annotations

import platform

from harness_pptx.backends.interface import RendererInterface


class BackendRegistry:
    """Registry of available backend implementations with auto-detection."""

    def __init__(self):
        self._backends: dict[str, type[RendererInterface]] = {}

    def register(self, name: str, backend_cls: type[RendererInterface]) -> None:
        self._backends[name] = backend_cls

    def get(self, name: str) -> RendererInterface:
        cls = self._backends.get(name)
        if cls is None:
            available = sorted(self._backends.keys())
            raise KeyError(f"Backend '{name}' not found. Available: {available}")
        return cls()

    def list(self) -> list[str]:
        return sorted(self._backends.keys())

    def auto_detect(self) -> RendererInterface:
        """Select the best backend for the current platform.

        macOS → AppleScript (native PowerPoint automation)
        Windows / Linux → pptx-xml (python-pptx cross-platform)
        """
        system = platform.system()
        if system == "Darwin":
            # macOS — try AppleScript backend first
            if "applescript" in self._backends:
                return self.get("applescript")
        # Fall back to pptx-xml for all platforms
        if "pptx-xml" in self._backends:
            return self.get("pptx-xml")
        raise RuntimeError(
            f"No suitable backend found for platform '{system}'. "
            f"Available backends: {self.list()}"
        )


# Singleton
_backend_registry: BackendRegistry | None = None


def get_backend_registry() -> BackendRegistry:
    global _backend_registry
    if _backend_registry is None:
        _backend_registry = BackendRegistry()
        from harness_pptx.backends.applescript_backend import AppleScriptBackend
        from harness_pptx.backends.pptx_xml_backend import PPTXXmlBackend
        _backend_registry.register("applescript", AppleScriptBackend)
        _backend_registry.register("pptx-xml", PPTXXmlBackend)
    return _backend_registry


def get_backend(name: str | None = None) -> RendererInterface:
    """Get a backend by name, or auto-detect."""
    reg = get_backend_registry()
    if name:
        return reg.get(name)
    return reg.auto_detect()
