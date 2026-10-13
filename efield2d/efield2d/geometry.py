"""2D geometry primitives.

A polygon is a list of vertices (x, y, bulge) - the same representation as a
DXF LWPOLYLINE.  `bulge` belongs to the segment that starts at that vertex:
0 = straight line, tan(theta/4) otherwise (theta = included angle, positive
counter-clockwise).  Coordinates are in the problem's length unit (mm by
default).
"""
from __future__ import annotations

import math
from dataclasses import dataclass, field


@dataclass
class Segment:
    p0: tuple[float, float]
    p1: tuple[float, float]
    bulge: float = 0.0

    # --- arc description (only meaningful when bulge != 0) -------------
    def arc(self):
        """Return (center, radius, a_start, sweep) of the equivalent CCW arc
        (a_start is the angle of the CCW start point, which is p1 for a
        clockwise segment)."""
        (x0, y0), (x1, y1) = self.p0, self.p1
        dx, dy = x1 - x0, y1 - y0
        chord = math.hypot(dx, dy)
        theta = 4.0 * math.atan(self.bulge)
        half = abs(theta) / 2.0
        radius = chord / (2.0 * math.sin(half))
        sgn = 1.0 if self.bulge > 0 else -1.0
        px, py = -dy / chord, dx / chord
        d = radius * math.cos(half)
        cx = (x0 + x1) / 2.0 + sgn * px * d
        cy = (y0 + y1) / 2.0 + sgn * py * d
        sx, sy = (x0, y0) if self.bulge > 0 else (x1, y1)
        a_start = math.atan2(sy - cy, sx - cx)
        return (cx, cy), radius, a_start, abs(theta)

    def distance(self, x: float, y: float) -> float:
        """Distance from (x, y) to this segment."""
        if self.bulge == 0.0:
            (x0, y0), (x1, y1) = self.p0, self.p1
            dx, dy = x1 - x0, y1 - y0
            L2 = dx * dx + dy * dy
            t = 0.0 if L2 == 0 else max(0.0, min(1.0, ((x - x0) * dx + (y - y0) * dy) / L2))
            return math.hypot(x - (x0 + t * dx), y - (y0 + t * dy))
        (cx, cy), R, a0, sweep = self.arc()
        ang = (math.atan2(y - cy, x - cx) - a0) % (2.0 * math.pi)
        if ang <= sweep:
            return abs(math.hypot(x - cx, y - cy) - R)
        return min(math.hypot(x - self.p0[0], y - self.p0[1]),
                   math.hypot(x - self.p1[0], y - self.p1[1]))


@dataclass
class Polygon:
    """Closed (or open, for lines) chain of line / arc segments."""
    vertices: list[tuple[float, float, float]]
    closed: bool = True

    @classmethod
    def from_points(cls, pts, closed: bool = True) -> "Polygon":
        return cls([(float(p[0]), float(p[1]), 0.0) for p in pts], closed)

    def segments(self) -> list[Segment]:
        v = self.vertices
        n = len(v)
        count = n if self.closed else n - 1
        out = []
        for i in range(count):
            x0, y0, b = v[i]
            x1, y1, _ = v[(i + 1) % n]
            out.append(Segment((x0, y0), (x1, y1), b))
        return out

    def distance(self, x: float, y: float) -> float:
        return min(s.distance(x, y) for s in self.segments())

    def signed_area(self) -> float:
        """Enclosed area, positive for counter-clockwise outlines (arcs included)."""
        v = self.vertices
        n = len(v)
        a = 0.0
        for i in range(n):
            x0, y0, b = v[i]
            x1, y1, _ = v[(i + 1) % n]
            a += 0.5 * (x0 * y1 - x1 * y0)
            if b != 0.0:
                th = 4.0 * math.atan(b)
                chord = math.hypot(x1 - x0, y1 - y0)
                r = chord / (2.0 * math.sin(abs(th) / 2.0))
                a += 0.5 * r * r * (th - math.sin(th))
        return a

    def reversed(self) -> "Polygon":
        """Same outline traversed the other way round."""
        v = self.vertices
        n = len(v)
        out = []
        for k in range(n):
            j = n - 1 - k
            b = -v[(j - 1) % n][2]
            if not self.closed and k == n - 1:
                b = 0.0
            out.append((v[j][0], v[j][1], b))
        return Polygon(out, self.closed)

    def bbox(self):
        xs, ys = [], []
        for s in self.segments():
            xs += [s.p0[0], s.p1[0]]
            ys += [s.p0[1], s.p1[1]]
            if s.bulge != 0.0:  # sample the arc so the bbox covers the bulge
                (cx, cy), R, a0, sw = s.arc()
                for k in range(0, 33):
                    a = a0 + sw * k / 32
                    xs.append(cx + R * math.cos(a))
                    ys.append(cy + R * math.sin(a))
        return min(xs), min(ys), max(xs), max(ys)


def rect(x0: float, y0: float, x1: float, y1: float) -> Polygon:
    return Polygon.from_points([(x0, y0), (x1, y0), (x1, y1), (x0, y1)])
