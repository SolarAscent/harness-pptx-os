"""Design system: theme tokens for colors, fonts, spacing, radii, shadows."""

from __future__ import annotations

from typing import Annotated

from pydantic import BaseModel, Field, StringConstraints

HexColor = Annotated[str, StringConstraints(pattern=r"^#[0-9a-fA-F]{6}$")]


# ---- Color tokens -----------------------------------------------------------

class ColorTokens(BaseModel):
    """Semantic color palette. All values are #RRGGBB hex strings."""

    primary: HexColor = "#003366"
    accent: HexColor = "#0078D7"
    background: HexColor = "#FFFFFF"
    text: HexColor = "#1A1A1A"
    muted: HexColor = "#767676"
    surface: HexColor = "#F5F5F5"
    border: HexColor = "#D0D0D0"
    success: HexColor = "#107C10"
    warning: HexColor = "#FF8C00"
    error: HexColor = "#D13438"


class FontTokens(BaseModel):
    """Font stack with CJK fallbacks."""

    title: str = "Calibri"
    title_zh: str = "Microsoft YaHei"
    body: str = "Calibri"
    body_zh: str = "Microsoft YaHei"
    mono: str = "Consolas"


class SpacingTokens(BaseModel):
    """Semantic spacing scale in points."""

    xs: float = 4.0
    sm: float = 8.0
    md: float = 16.0
    lg: float = 24.0
    xl: float = 36.0
    xxl: float = 48.0


class RadiiTokens(BaseModel):
    """Corner radii in points."""

    none: float = 0.0
    sm: float = 3.0
    md: float = 6.0
    lg: float = 12.0


class ShadowTokens(BaseModel):
    """Shadow presets as CSS-like box-shadow strings (not rendered directly)."""

    none: str = "none"
    sm: str = "0 1pt 2pt rgba(0,0,0,0.10)"
    md: str = "0 2pt 6pt rgba(0,0,0,0.15)"
    lg: str = "0 4pt 12pt rgba(0,0,0,0.20)"


# ---- Theme ------------------------------------------------------------------

class Theme(BaseModel):
    """Complete design theme for a deck.

    Agents reference tokens (e.g. ``color: token.primary``) rather than
    hard-coding raw color values.
    """

    name: str = "corporate"
    colors: ColorTokens = Field(default_factory=ColorTokens)
    fonts: FontTokens = Field(default_factory=FontTokens)
    spacing: SpacingTokens = Field(default_factory=SpacingTokens)
    radii: RadiiTokens = Field(default_factory=RadiiTokens)
    shadows: ShadowTokens = Field(default_factory=ShadowTokens)
    chart_palette: list[HexColor] = Field(
        default_factory=lambda: [
            "#003366",
            "#0078D7",
            "#107C10",
            "#FF8C00",
            "#D13438",
            "#767676",
            "#8764B8",
            "#00B7C3",
        ]
    )

    def color(self, token: str) -> str:
        """Resolve a color token name to its hex value."""
        return getattr(self.colors, token)

    def space(self, token: str) -> float:
        """Resolve a spacing token name to its point value."""
        return getattr(self.spacing, token)

    def font(self, token: str) -> str:
        """Resolve a font token name to its family string."""
        return getattr(self.fonts, token)
