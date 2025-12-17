"""
Simcenter Magnet 3D Box Modeling Script
========================================

This script creates a 3D rectangular box (cuboid) in Simcenter Magnet.
The box is created with one corner at the origin (0,0,0) and extends
in the positive direction.

Author: Claude
Date: 2025-12-17
"""

import sys
try:
    from MagNet import *
except ImportError:
    print("Warning: MagNet module not found. This script requires Simcenter Magnet to run.")
    print("The script structure is ready for use within Simcenter Magnet environment.")


def create_box(width, depth, height, material_name="Air"):
    """
    Create a 3D rectangular box in Simcenter Magnet.

    Parameters:
    -----------
    width : float
        Width of the box in the X direction (mm)
    depth : float
        Depth of the box in the Y direction (mm)
    height : float
        Height of the box in the Z direction (mm)
    material_name : str, optional
        Material to assign to the box (default: "Air")

    Returns:
    --------
    component : MagNet Component object
        The created box component
    """

    # Validate input parameters
    if width <= 0 or depth <= 0 or height <= 0:
        raise ValueError("All dimensions must be positive values")

    print(f"Creating box with dimensions:")
    print(f"  Width (X):  {width} mm")
    print(f"  Depth (Y):  {depth} mm")
    print(f"  Height (Z): {height} mm")
    print(f"  Material:   {material_name}")

    try:
        # Get the current MagNet document
        magnet = GetDocument()

        # Define the 8 vertices of the box
        # Starting from origin (0,0,0) and extending in positive direction
        vertices = [
            [0, 0, 0],           # Vertex 0: origin
            [width, 0, 0],       # Vertex 1: +X
            [width, depth, 0],   # Vertex 2: +X, +Y
            [0, depth, 0],       # Vertex 3: +Y
            [0, 0, height],      # Vertex 4: +Z
            [width, 0, height],  # Vertex 5: +X, +Z
            [width, depth, height],  # Vertex 6: +X, +Y, +Z
            [0, depth, height]   # Vertex 7: +Y, +Z
        ]

        # Define the 6 faces of the box using vertex indices
        # Each face is defined by 4 vertices in counter-clockwise order
        faces = [
            [0, 1, 2, 3],  # Bottom face (Z=0)
            [4, 7, 6, 5],  # Top face (Z=height)
            [0, 4, 5, 1],  # Front face (Y=0)
            [2, 6, 7, 3],  # Back face (Y=depth)
            [0, 3, 7, 4],  # Left face (X=0)
            [1, 5, 6, 2]   # Right face (X=width)
        ]

        # Create the component using MagNet API
        # Method 1: Using MakeComponentInABox (simpler method)
        component = magnet.getView().newComponent("Box")

        # Create a box using the built-in box creation method
        component.makeComponentInABox(
            0, 0, 0,              # Starting point (X, Y, Z)
            width, depth, height  # Dimensions
        )

        # Assign material to the component
        component.setMaterial(material_name)

        print(f"✓ Box component created successfully: {component.getName()}")

        # Update the view
        magnet.getView().viewAll()

        return component

    except Exception as e:
        print(f"Error creating box: {str(e)}")
        raise


def create_box_advanced(width, depth, height, material_name="Air", name="Box"):
    """
    Create a 3D rectangular box using vertex and face definition (advanced method).

    This method provides more control over the geometry definition.

    Parameters:
    -----------
    width : float
        Width of the box in the X direction (mm)
    depth : float
        Depth of the box in the Y direction (mm)
    height : float
        Height of the box in the Z direction (mm)
    material_name : str, optional
        Material to assign to the box (default: "Air")
    name : str, optional
        Name of the component (default: "Box")

    Returns:
    --------
    component : MagNet Component object
        The created box component
    """

    # Validate input parameters
    if width <= 0 or depth <= 0 or height <= 0:
        raise ValueError("All dimensions must be positive values")

    print(f"Creating box (advanced method) with dimensions:")
    print(f"  Width (X):  {width} mm")
    print(f"  Depth (Y):  {depth} mm")
    print(f"  Height (Z): {height} mm")
    print(f"  Material:   {material_name}")

    try:
        # Get the current MagNet document
        magnet = GetDocument()
        view = magnet.getView()

        # Create arrays for edges and faces
        # Bottom face
        edge1 = view.newLine(0, 0, 0, width, 0, 0)
        edge2 = view.newLine(width, 0, 0, width, depth, 0)
        edge3 = view.newLine(width, depth, 0, 0, depth, 0)
        edge4 = view.newLine(0, depth, 0, 0, 0, 0)

        # Create bottom face from edges
        bottom_edges = [edge1, edge2, edge3, edge4]

        # Extrude the bottom face to create the box
        component = view.newComponent(name)
        component.sweepAlongVector(bottom_edges, 0, 0, height)

        # Assign material
        component.setMaterial(material_name)

        print(f"✓ Box component created successfully: {component.getName()}")

        # Update the view
        view.viewAll()

        return component

    except Exception as e:
        print(f"Error creating box: {str(e)}")
        raise


def main():
    """
    Main function to demonstrate box creation.
    """
    print("=" * 60)
    print("Simcenter Magnet - 3D Box Creation Script")
    print("=" * 60)

    # Default parameters
    default_width = 100.0   # mm
    default_depth = 50.0    # mm
    default_height = 30.0   # mm

    # Get parameters from command line arguments if provided
    if len(sys.argv) >= 4:
        try:
            width = float(sys.argv[1])
            depth = float(sys.argv[2])
            height = float(sys.argv[3])
            material = sys.argv[4] if len(sys.argv) >= 5 else "Air"
        except ValueError:
            print("Error: Invalid parameters. Using default values.")
            width, depth, height, material = default_width, default_depth, default_height, "Air"
    else:
        print("Usage: python create_box.py <width> <depth> <height> [material]")
        print(f"\nUsing default parameters:")
        width, depth, height, material = default_width, default_depth, default_height, "Air"

    print()

    # Create the box
    try:
        box = create_box(width, depth, height, material)
        print("\n" + "=" * 60)
        print("Box created successfully!")
        print("=" * 60)
    except Exception as e:
        print(f"\nFailed to create box: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()
