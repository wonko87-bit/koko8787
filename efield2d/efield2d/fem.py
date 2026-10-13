"""Linear-triangle FEM for  div(kappa grad V) = 0  (planar / axisymmetric)."""
from __future__ import annotations

from dataclasses import dataclass

import numpy as np
import scipy.sparse as sp
import scipy.sparse.linalg as spla

from .materials import MODES
from .mesh import Mesh
from .problem import Problem


@dataclass
class Solution:
    problem: Problem
    mesh: Mesh
    mode: str
    freq: float
    phi: np.ndarray            # (N,) potential, complex for mode "ac"
    kappa_region: np.ndarray   # (R,) coefficient per region


def element_geometry(mesh: Mesh):
    """Per-triangle b, c (shape functions' gradient numerators), area."""
    p = mesh.nodes[mesh.tris]
    x, y = p[:, :, 0], p[:, :, 1]
    b = np.stack([y[:, 1] - y[:, 2], y[:, 2] - y[:, 0], y[:, 0] - y[:, 1]], axis=1)
    c = np.stack([x[:, 2] - x[:, 1], x[:, 0] - x[:, 2], x[:, 1] - x[:, 0]], axis=1)
    area = 0.5 * ((x[:, 1] - x[:, 0]) * (y[:, 2] - y[:, 0]) - (x[:, 2] - x[:, 0]) * (y[:, 1] - y[:, 0]))
    return b, c, area, x.mean(axis=1)


def volume_weight(problem: Problem, mesh: Mesh) -> np.ndarray:
    """Volume of each element: area * depth (planar, 1 m) or 2*pi*r_centroid*area."""
    _, _, area, xc = element_geometry(mesh)
    return area * (2.0 * np.pi * xc if problem.symmetry == "axisym" else 1.0)


def assemble(problem: Problem, mesh: Mesh, kappa_region: np.ndarray) -> sp.csr_matrix:
    b, c, area, _ = element_geometry(mesh)
    w = volume_weight(problem, mesh) / (area * area)  # volume / area^2
    kap = kappa_region[mesh.region] * w / 4.0
    ke = kap[:, None, None] * (b[:, :, None] * b[:, None, :] + c[:, :, None] * c[:, None, :])
    rows = np.repeat(mesh.tris, 3, axis=1).ravel()
    cols = np.tile(mesh.tris, (1, 3)).ravel()
    n = mesh.n_nodes
    return sp.coo_matrix((ke.ravel(), (rows, cols)), shape=(n, n)).tocsr()


def solve(problem: Problem, mesh: Mesh, mode: str = "es", freq: float = 0.0) -> Solution:
    if mode not in MODES:
        raise ValueError(f"mode must be one of {MODES}")
    if mode == "ac" and freq <= 0:
        raise ValueError("mode 'ac' needs freq > 0")
    kappa = np.array([r.material.kappa(mode, freq) for r in problem.regions])
    if mode == "dc" and np.all(kappa == 0):
        raise ValueError("DC analysis needs at least one region with sigma > 0")
    # a region with zero kappa makes the matrix singular; give it a tiny value
    # relative to the largest one (it carries no field anyway)
    kmax = np.max(np.abs(kappa))
    kappa = np.where(np.abs(kappa) < 1e-12 * kmax, 1e-12 * kmax, kappa)

    if len(mesh.bc_nodes) == 0:
        raise ValueError("no Dirichlet (electrode) nodes - the problem is singular")
    complex_case = mode == "ac" or np.iscomplexobj(mesh.bc_values) and np.any(mesh.bc_values.imag != 0)
    dtype = complex if complex_case else float
    K = assemble(problem, mesh, kappa.astype(dtype)).astype(dtype)

    n = mesh.n_nodes
    fixed = np.zeros(n, dtype=bool)
    fixed[mesh.bc_nodes] = True
    free = ~fixed
    vd = np.zeros(n, dtype=dtype)
    vd[mesh.bc_nodes] = mesh.bc_values.astype(dtype) if complex_case else mesh.bc_values.real

    Kcsc = K.tocsc()
    rhs = -(Kcsc[:, np.flatnonzero(fixed)] @ vd[fixed])[free]
    phi = vd.copy()
    phi[free] = spla.spsolve(Kcsc[free][:, free].tocsc(), rhs)
    return Solution(problem, mesh, mode, freq, phi, kappa)
