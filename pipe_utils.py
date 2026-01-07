from scipy.optimize import brentq
import math

def calculate_insulation_thickness(
        UValue: float,
        lambdaInsulation: float,
        lambdaWall: float,
        di: float,
        da: float,
        dOuterLayer: float,
        max_thickness: float = 1.0
    ) -> float:
        """
        Calculates the required insulation thickness to achieve a target U-Value.

        This function uses a numerical root-finding algorithm (SciPy's brentq) to
        solve the transcendental heat transfer equation for `dInsulation`.

        Args:
            UValue (float): The target overall heat transfer coefficient [W/(m*K)].
                            Must be a positive value.
            lambdaInsulation (float): Thermal conductivity of the insulation material [W/(m*K)].
                                      Must be a positive value.
            lambdaWall (float): Thermal conductivity of the main wall/pipe material [W/(m*K)].
                                Must be a positive value.
            di (float): Inner diameter of the pipe/wall [m]. Must be positive.
            da (float): Outer diameter of the pipe/wall (inner diameter of insulation) [m].
                        Must be greater than di.
            dOuterLayer (float): Thickness of the outer protective layer [m]. Must be non-negative.
            max_thickness (float, optional): The upper search bound for the insulation thickness [m].
                                             Defaults to 1.0 meter.

        Returns:
            float: The calculated required thickness of the insulation (`dInsulation`) [m].
                   Returns 0.0 if no insulation is needed.

        Raises:
            ValueError: If input parameters are not physically valid or if the target U-Value
                        cannot be achieved even with the maximum specified insulation thickness.
        """
        # --- Input Validation ---
        if UValue <= 0:
            raise ValueError("UValue must be positive.")
        if lambdaInsulation <= 0 or lambdaWall <= 0:
            raise ValueError("Thermal conductivities (lambdaInsulation, lambdaWall) must be positive.")
        if not (di > 0 and da > 0 and dOuterLayer >= 0):
            raise ValueError("All diameters and thicknesses must be positive (dOuterLayer can be zero).")
        if da <= di:
            raise ValueError("Outer diameter (da) must be greater than inner diameter (di).")

        # --- Calculation Setup ---
        
        # 1. Required total thermal resistance for the target U-Value
        r_total_required = math.pi / UValue

        # 2. Resistance of the inner wall layer (constant)
        r_wall_inner = (1 / (2 * lambdaWall)) * math.log(da / di)
        
        # Check if the wall alone already meets the requirement
        if r_wall_inner >= r_total_required:
            print("Warning: No insulation needed. The base wall already meets or exceeds the target U-Value.")
            return 0.0

        # 3. Define the function f(d_ins) = 0 for the root-finder
        # f(d_ins) = (calculated_total_r) - (required_r)
        def residual_function(d_ins: float) -> float:
            # Avoid math errors for d_ins <= 0, although solver stays in positive bounds
            if d_ins <= 0:
                # Resistance with no insulation
                return r_wall_inner - r_total_required
            
            # Resistance of the insulation layer
            r_insulation = (1 / (2 * lambdaInsulation)) * math.log((da + 2 * d_ins) / da)
            
            # Resistance of the outer protective layer
            r_wall_outer = (1 / (2 * lambdaWall)) * math.log(
                (da + 2 * d_ins + 2 * dOuterLayer) / (da + 2 * d_ins)
            )
            
            # Total calculated resistance for a given d_ins
            r_calculated = r_wall_inner + r_insulation + r_wall_outer
            
            # The residual: we want this to be zero
            return r_calculated - r_total_required

        # --- Numerical Solver ---
        
        # The lower bound for insulation thickness is essentially zero
        lower_bound = 1e-9  # A small positive number to avoid log(1) issues
        
        # Check the signs at the bounds to ensure a root exists in the interval
        f_lower = residual_function(lower_bound)
        f_upper = residual_function(max_thickness)

        if f_lower > 0:
             # This case is already handled by the check r_wall_inner >= r_total_required,
             # but serves as a robust safeguard.
             print("Warning: No insulation needed. The base wall already meets the target U-Value.")
             return 0.0
        
        if f_upper < 0:
            raise ValueError(
                f"Target U-Value of {UValue} is too low to be achieved even with "
                f"{max_thickness*1000:.1f} mm of insulation. "
                f"Try increasing max_thickness or using a better insulation material (lower lambda)."
            )
        
        if f_lower * f_upper >= 0:
            # This should not happen if the above checks pass, but is a solver prerequisite
            raise RuntimeError(
                "Cannot solve: The residual function does not cross zero in the search interval. "
                "There might be an issue with the provided parameters."
            )

        # Use brentq to find the root (the value of d_ins where residual is zero)
        try:
            insulation_thickness = brentq(residual_function, a=lower_bound, b=max_thickness)
            return insulation_thickness
        except ValueError as e:
            # This catches errors from brentq if it fails to converge
            raise RuntimeError(f"Numerical solver failed to find a solution: {e}")
