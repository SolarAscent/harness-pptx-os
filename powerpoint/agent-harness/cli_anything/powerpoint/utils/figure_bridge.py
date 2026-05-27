"""
Nature-Figure Bridge: Generate publication-quality academic charts for PPT insertion.

Generates high-resolution matplotlib figures following Nature-style aesthetics
and returns the image path for seamless embedding into PowerPoint slides.

Chart types: bar, hbar, line, scatter, pie, heatmap, radar, donut
Color themes: nature, pku, dark, light
"""

from __future__ import annotations

import json
import tempfile
from pathlib import Path
from typing import Any


class FigureBridge:
    """Generates publication-quality matplotlib figures and saves as high-res PNG."""

    # ── Color themes (RGB hex) ───────────────────────────────────────

    COLOR_THEMES = {
        "nature": {
            "bg": "#FFFFFF",
            "fg": "#222222",
            "grid": "#E5E5E5",
            "palette": ["#3A7CA5", "#D64933", "#5B8C5A", "#F4A261", "#264653",
                        "#E9C46A", "#A8DADC", "#457B9D", "#6D6875", "#B5838D"],
            "accent": "#D64933",
        },
        "pku": {
            "bg": "#FFFFFF",
            "fg": "#222222",
            "grid": "#E8E0D5",
            "palette": ["#8B0000", "#C8A030", "#2C3E50", "#8B4513", "#D4A574",
                        "#A0522D", "#CD853F", "#800000", "#B8860B", "#556B2F"],
            "accent": "#C8A030",
        },
        "dark": {
            "bg": "#1E1E2E",
            "fg": "#E8E8E8",
            "grid": "#3A3A50",
            "palette": ["#89B4FA", "#F38BA8", "#A6E3A1", "#F9E2AF", "#89DCEB",
                        "#CBA6F7", "#FAB387", "#94E2D5", "#B4BEFE", "#F5C2E7"],
            "accent": "#F38BA8",
        },
        "light": {
            "bg": "#FAFAFA",
            "fg": "#333333",
            "grid": "#E0E0E0",
            "palette": ["#4A90D9", "#E05555", "#4CAF50", "#FF9800", "#9C27B0",
                        "#00BCD4", "#8BC34A", "#FF5722", "#607D8B", "#E91E63"],
            "accent": "#E05555",
        },
    }

    def generate(
        self,
        chart_type: str,
        title: str = "",
        xlabel: str = "",
        ylabel: str = "",
        width: int = 600,
        height: int = 400,
        data: str | None = None,
        color_theme: str = "nature",
    ) -> dict[str, Any]:
        """Generate a publication-quality chart and return the file path."""
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        import matplotlib.ticker as mticker
        import numpy as np

        theme = self.COLOR_THEMES.get(color_theme, self.COLOR_THEMES["nature"])

        # Publication-quality rcParams
        plt.rcParams.update({
            "font.family": "sans-serif",
            "font.sans-serif": ["Arial", "Helvetica", "DejaVu Sans", "sans-serif"],
            "font.size": 10,
            "axes.titlesize": 13,
            "axes.labelsize": 10,
            "axes.spines.right": False,
            "axes.spines.top": False,
            "axes.linewidth": 0.6,
            "axes.edgecolor": theme["fg"],
            "xtick.color": theme["fg"],
            "ytick.color": theme["fg"],
            "text.color": theme["fg"],
            "legend.frameon": False,
            "legend.fontsize": 8,
            "grid.alpha": 0.3,
            "figure.dpi": 150,
        })

        fig, ax = plt.subplots(figsize=(width / 100, height / 100))
        fig.patch.set_facecolor(theme["bg"])
        ax.set_facecolor(theme["bg"])

        # Parse data
        parsed = self._parse_data(data)

        # Draw chart
        dispatcher = {
            "bar": self._draw_bar,
            "hbar": self._draw_hbar,
            "line": self._draw_line,
            "scatter": self._draw_scatter,
            "pie": self._draw_pie,
            "heatmap": self._draw_heatmap,
            "radar": self._draw_radar,
            "donut": self._draw_donut,
        }
        draw_fn = dispatcher.get(chart_type)
        if draw_fn is None:
            return {"status": "error", "error": f"Unknown chart type: {chart_type}. "
                    f"Available: {list(dispatcher.keys())}"}

        draw_fn(ax, parsed, theme, title, xlabel, ylabel)
        fig.tight_layout()

        # Save to temp file
        tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False, prefix="figure_")
        fig.savefig(tmp.name, dpi=300, bbox_inches="tight", facecolor=theme["bg"], edgecolor="none")
        plt.close(fig)

        return {"status": "ok", "path": tmp.name, "chart_type": chart_type, "theme": color_theme}

    # ── Data parsing ──────────────────────────────────────────────────

    def _parse_data(self, data: str | None) -> list[dict]:
        """Parse input data into a list of dicts.

        Formats:
          JSON: '[{"label": "A", "value": 10}, ...]'
          Simple: 'A:10,B:20,C:30' or 'series1:10,20;series2:15,25'
        """
        if data is None:
            return self._demo_data()

        data = data.strip()
        # JSON format
        if data.startswith("["):
            return json.loads(data)

        # Simple format: 'label1:v1,label2:v2'
        if ":" in data and ";" not in data:
            result = []
            for part in data.split(","):
                if ":" in part:
                    label, val = part.split(":", 1)
                    result.append({"label": label.strip(), "value": float(val.strip())})
            if result:
                return result

        # Multi-series format: 'series1:v1,v2,v3;series2:v1,v2,v3'
        if ";" in data and ":" in data:
            result = []
            for series in data.split(";"):
                if ":" in series:
                    name, vals = series.split(":", 1)
                    result.append({
                        "series": name.strip(),
                        "values": [float(v.strip()) for v in vals.split(",")],
                    })
            if result:
                return result

        return self._demo_data()

    def _demo_data(self) -> list[dict]:
        return [
            {"label": "Category A", "value": 45},
            {"label": "Category B", "value": 62},
            {"label": "Category C", "value": 38},
            {"label": "Category D", "value": 71},
            {"label": "Category E", "value": 55},
        ]

    def _demo_multiseries(self) -> list[dict]:
        return [
            {"series": "Group 1", "values": [45, 62, 38, 71, 55]},
            {"series": "Group 2", "values": [52, 48, 65, 58, 43]},
        ]

    # ── Chart drawing functions ──────────────────────────────────────

    def _draw_bar(self, ax, data, theme, title, xlabel, ylabel):
        import numpy as np

        if data and "series" in data[0]:
            # Multi-series grouped bar
            series_list = data
            labels = [f"Cat {i+1}" for i in range(len(series_list[0]["values"]))]
            x = np.arange(len(labels))
            n_series = len(series_list)
            bar_w = 0.8 / n_series
            for i, s in enumerate(series_list):
                offset = (i - n_series / 2 + 0.5) * bar_w
                ax.bar(x + offset, s["values"], bar_w * 0.9,
                       label=s["series"], color=theme["palette"][i % len(theme["palette"])],
                       edgecolor="none")
            ax.set_xticks(x)
            ax.set_xticklabels(labels)
            ax.legend()
        else:
            labels = [d.get("label", f"Item {i+1}") for i, d in enumerate(data)]
            values = [d.get("value", 0) for d in data]
            bars = ax.bar(range(len(labels)), values,
                          color=[theme["palette"][i % len(theme["palette"])] for i in range(len(labels))],
                          edgecolor="none", width=0.65)
            ax.set_xticks(range(len(labels)))
            ax.set_xticklabels(labels, rotation=25, ha="right", fontsize=8)

            # Value labels on bars
            for bar, val in zip(bars, values):
                ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + max(values) * 0.01,
                        str(val), ha="center", va="bottom", fontsize=7, color=theme["fg"])

        ax.grid(axis="y", color=theme["grid"], linewidth=0.4)
        if title:
            ax.set_title(title, fontweight="bold", pad=12)
        if xlabel:
            ax.set_xlabel(xlabel)
        if ylabel:
            ax.set_ylabel(ylabel)

    def _draw_hbar(self, ax, data, theme, title, xlabel, ylabel):
        labels = [d.get("label", f"Item {i+1}") for i, d in enumerate(data)]
        values = [d.get("value", 0) for d in data]
        colors = [theme["palette"][i % len(theme["palette"])] for i in range(len(labels))]
        bars = ax.barh(range(len(labels)), values, color=colors, edgecolor="none", height=0.55)
        ax.set_yticks(range(len(labels)))
        ax.set_yticklabels(labels, fontsize=8)
        ax.invert_yaxis()
        for bar, val in zip(bars, values):
            ax.text(bar.get_width() + max(values) * 0.01, bar.get_y() + bar.get_height() / 2,
                    str(val), ha="left", va="center", fontsize=7, color=theme["fg"])
        ax.grid(axis="x", color=theme["grid"], linewidth=0.4)
        if title:
            ax.set_title(title, fontweight="bold", pad=12)
        if xlabel:
            ax.set_xlabel(xlabel)
        if ylabel:
            ax.set_ylabel(ylabel)

    def _draw_line(self, ax, data, theme, title, xlabel, ylabel):
        if not data or "series" not in data[0]:
            data = self._demo_multiseries()

        for i, s in enumerate(data):
            vals = s["values"]
            color = theme["palette"][i % len(theme["palette"])]
            ax.plot(range(len(vals)), vals, color=color, linewidth=2.0,
                    marker="o", markersize=5, label=s["series"], markeredgecolor="none")
        ax.legend(fontsize=7)
        ax.grid(color=theme["grid"], linewidth=0.4)
        if title:
            ax.set_title(title, fontweight="bold", pad=12)
        if xlabel:
            ax.set_xlabel(xlabel)
        if ylabel:
            ax.set_ylabel(ylabel)

    def _draw_scatter(self, ax, data, theme, title, xlabel, ylabel):
        if not data or "series" not in data[0]:
            data = self._demo_multiseries()
        for i, s in enumerate(data):
            vals = s["values"]
            color = theme["palette"][i % len(theme["palette"])]
            x = list(range(len(vals)))
            ax.scatter(x, vals, c=color, s=50, alpha=0.8,
                       label=s["series"], edgecolors="none")
        ax.legend(fontsize=7)
        ax.grid(color=theme["grid"], linewidth=0.4)
        if title:
            ax.set_title(title, fontweight="bold", pad=12)
        if xlabel:
            ax.set_xlabel(xlabel)
        if ylabel:
            ax.set_ylabel(ylabel)

    def _draw_pie(self, ax, data, theme, title, xlabel, ylabel):
        labels = [d.get("label", f"Item {i+1}") for i, d in enumerate(data)]
        values = [d.get("value", 0) for d in data]
        colors = theme["palette"][:len(labels)]
        wedges, texts, autotexts = ax.pie(
            values, labels=None, autopct="%1.1f%%",
            colors=colors, startangle=90, pctdistance=0.75,
            wedgeprops={"edgecolor": "white", "linewidth": 1.5},
        )
        for t in autotexts:
            t.set_fontsize(8)
        ax.legend(wedges, labels, title="", loc="center left",
                  bbox_to_anchor=(1, 0, 0.5, 1), fontsize=7)
        if title:
            ax.set_title(title, fontweight="bold", pad=12)

    def _draw_donut(self, ax, data, theme, title, xlabel, ylabel):
        labels = [d.get("label", f"Item {i+1}") for i, d in enumerate(data)]
        values = [d.get("value", 0) for d in data]
        colors = theme["palette"][:len(labels)]
        wedges, texts, autotexts = ax.pie(
            values, labels=None, autopct="%1.1f%%",
            colors=colors, startangle=90, pctdistance=0.78,
            wedgeprops={"edgecolor": "white", "linewidth": 1.5, "width": 0.4},
        )
        for t in autotexts:
            t.set_fontsize(8)
        ax.legend(wedges, labels, title="", loc="center left",
                  bbox_to_anchor=(1, 0, 0.5, 1), fontsize=7)
        if title:
            ax.set_title(title, fontweight="bold", pad=12)

    def _draw_heatmap(self, ax, data, theme, title, xlabel, ylabel):
        import numpy as np

        # Expect data as a 2D array or list of value objects
        if data and isinstance(data[0], dict) and "values" in data[0]:
            matrix = np.array([s["values"] for s in data])
        else:
            np.random.seed(42)
            matrix = np.random.rand(6, 8) * 100

        im = ax.imshow(matrix, cmap="YlOrRd", aspect="auto")
        cbar = plt.colorbar(im, ax=ax, shrink=0.85)
        cbar.ax.tick_params(labelsize=7)
        cbar.outline.set_visible(False)

        # Annotate cells
        for i in range(matrix.shape[0]):
            for j in range(matrix.shape[1]):
                ax.text(j, i, f"{matrix[i, j]:.0f}", ha="center", va="center",
                        fontsize=7, color="white" if matrix[i, j] > matrix.max() * 0.5 else "black")

        if title:
            ax.set_title(title, fontweight="bold", pad=12)
        if xlabel:
            ax.set_xlabel(xlabel)
        if ylabel:
            ax.set_ylabel(ylabel)

    def _draw_radar(self, ax, data, theme, title, xlabel, ylabel):
        import numpy as np

        categories = [d.get("label", f"Dim {i+1}") for i, d in enumerate(data)]
        values = [d.get("value", 0) for d in data]
        N = len(categories)
        angles = np.linspace(0, 2 * np.pi, N, endpoint=False).tolist()
        values += values[:1]
        angles += angles[:1]

        ax = plt.gca()
        ax.set_facecolor(theme["bg"])
        ax = plt.subplot(111, polar=True)
        ax.set_facecolor(theme["bg"])

        ax.fill(angles, values, alpha=0.25, color=theme["palette"][0])
        ax.plot(angles, values, "o-", linewidth=2, color=theme["palette"][0], markersize=5)
        ax.fill(angles, values, alpha=0.08, color=theme["palette"][0])
        ax.set_xticks(angles[:-1])
        ax.set_xticklabels(categories, fontsize=8, color=theme["fg"])
        ax.set_yticklabels([])
        ax.spines["polar"].set_color(theme["grid"])
        if title:
            plt.title(title, fontweight="bold", pad=20, color=theme["fg"])
