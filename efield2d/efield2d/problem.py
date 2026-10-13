"""Problem definition: regions, electrodes and mesh controls.

Overlap rule (painter's algorithm): where entities overlap, the one that comes
later wins; conductors always win over regions.  So the usual workflow is
"draw a big air region first, then place insulators and electrodes in it".
Anything not covered by a region or a conductor is simply not meshed.

Boundaries that are not an electrode / Dirichlet line are natural (Neumann,
zero normal D or J) - this is also the symmetry condition, and for an
axisymmetric model the r = 0 axis.
"""
from __future__ import annotations

from dataclasses import dataclass, field

from .geometry import Polygon
from .materials import Material


@dataclass
class Region:
    polygon: Polygon
    material: Material
    name: str = ""


@dataclass
class Conductor:
    """Equipotential body. Its interior is not meshed; its outline is Dirichlet."""
    polygon: Polygon
    voltage: complex = 0.0
    name: str = ""


@dataclass
class DirichletLine:
    """A fixed-potential line lying on the outline of a region."""
    polygon: Polygon  # open polyline (closed=False)
    voltage: complex = 0.0
    name: str = ""


@dataclass
class Refinement:
    """Local mesh refinement: size `size` within `radius` of (x, y), growing
    back to the global size over `transition`."""
    x: float
    y: float
    size: float
    radius: float = 0.0
    transition: float | None = None


@dataclass
class Problem:
    regions: list[Region] = field(default_factory=list)
    conductors: list[Conductor] = field(default_factory=list)
    lines: list[DirichletLine] = field(default_factory=list)
    symmetry: str = "planar"  # "planar" (depth 1 m) | "axisym" (x = r, y = z)
    unit: float = 1e-3        # length of one geometry unit in metres (mm)
    mesh_size: float = 5.0    # global target element size, geometry units
    # element size on electrode / Dirichlet outlines, growing to mesh_size over
    # `edge_transition`. None = no edge refinement.
    edge_size: float | None = None
    edge_transition: float | None = None
    refinements: list[Refinement] = field(default_factory=list)

    def __post_init__(self):
        if self.symmetry not in ("planar", "axisym"):
            raise ValueError("symmetry must be 'planar' or 'axisym'")
