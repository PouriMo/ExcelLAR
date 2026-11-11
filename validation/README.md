# Validation

These Python scripts check ExcelLAR without using any of its code. They need NumPy and SciPy, and they run from this folder.

`lbla.py` reads a data sheet, assembles the equilibrium and admissibility conditions and solves the full linear program with HiGHS. `python lbla.py ../examples` prints the multiplier of every example.

`refine.py` generates the radial refinement series of case study b in `../examples/convergence/`. `ref2.py` computes the reference of the continuous arch, which ExcelLAR cannot model since it works on discrete blocks: the four-hinge kinematic multiplier of the continuous ring, with closed-form integrals of the self-weight and hinge positions free along the arch. The value is 10.84.

`minthick.py` finds the minimum thickness of a semicircular arch under self-weight, for radial and for vertical joints. With 400 blocks it gives t/R = 0.1075 and 0.1095. It also writes the 100-block models of `../examples/stereotomy/`, which ExcelLAR solves with t/R = 0.1075 and 0.1093.

`conditioning.py` computes the condition number of the matrix G1 for different choices of the interface kept as independent.

`models.py` writes data sheets in the ExcelLAR format.
