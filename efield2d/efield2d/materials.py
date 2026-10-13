"""Materials and the analysis modes that decide which coefficient is used.

All three analyses solve  div(kappa grad V) = 0  with a different kappa:

  mode "es"  electrostatic     kappa = eps0 * eps_r            (instantaneous / impulse)
  mode "dc"  DC steady current kappa = sigma                   (DC withstand: sigma decides)
  mode "ac"  time-harmonic     kappa = sigma + j*omega*eps0*eps_r   (AC withstand)

Note that for a multi-dielectric insulation the DC field distribution follows
the conductivity, not the permittivity.
"""
from __future__ import annotations

import math
from dataclasses import dataclass

EPS0 = 8.8541878128e-12  # F/m

MODES = ("es", "dc", "ac")


@dataclass(frozen=True)
class Material:
    name: str
    eps_r: float = 1.0
    sigma: float = 0.0  # S/m
    # allowable (breakdown) field, V/m. Optional: used by the withstand check.
    # Either a constant or a curve [(x, E_allow), ...] - see post.allowable().
    e_allow: float | None = None

    def kappa(self, mode: str, freq: float = 0.0) -> complex | float:
        if mode == "es":
            return EPS0 * self.eps_r
        if mode == "dc":
            return self.sigma
        if mode == "ac":
            return complex(self.sigma, 2.0 * math.pi * freq * EPS0 * self.eps_r)
        raise ValueError(f"unknown mode {mode!r}, expected one of {MODES}")
