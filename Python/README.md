Using Python to write LARSA 4D Macros
=====================================

The [Python](https://www.python.org/) programming language can be used to write macros for LARSA 4D.

Since Python does not have a code editor like the VBA code editor in Microsoft Excel which provides hints about object class, property, and method names, it is hard to know ahead of time which methods and properties you can use. It is recommended to prototype a macro in VBA in Excel first, or contact the LARSA support team for help.

Preparing Your Workstation
--------------------------

* Install the latest Windows installer for Python from https://www.python.org/downloads/windows/. Use 32-bit Python for Approach 1, which accesses the LarsaElements library directly. Approach 2 uses the running LARSA 4D Application object and can use 64-bit Python.

In a Command Prompt (Windows start menu -> Command Prompt), run these commands:

```
py -m pip install pywin32

py -m win32com.client.makepy "LARSA 2000 Library Elements"
```

The first command installs the open source pywin32 package which enables connecting to COM-enabled applications in Python. The second command runs a tool it provides for it to learn about the LARSA 4D COM interface. (Note that if you have multiple versions of Python installed, you may need to select a version. Try running `py --help` and `py --list` for additional instructions.)

You can then start Python's IDLE tool from the Windows start menu.

After running the third command above, it will output that it's creating a file and will give the path to that file. Looking inside that file can help find the names of constants and methods for the LARSA 4D interface, although it’s quite complex.

Macro Approaches
----------------

There are two approaches for writing LARSA 4D macros. Examples will be given for each. (The examples, and the two approaches for using the LARSA 4D COM interface, correspond exactly to how macros are written in VBA in Excel.)

### 1) Using the LarsaElements library directly

The first approach uses the "LarsaElements library" directly. This approach has the following benefits:

* It is faster than the second approach when creating or accessing a very large amount of project or results information.
* LARSA 4D .lar files can be read and written without LARSA 4D being open.

It has the following limitations:

* A 32-bit version of Python is required.
* Live-updating a project that is open in LARSA 4D is not possible.

Example: [LARSA4D_Macro_Approach1.py](LARSA4D_Macro_Approach1.py)

### 2) Using the Application object

The second approach manipulates a running instance of LARSA 4D using the LARSA 4D Application object. It has the following benefits:

* It is possible to live-update a project open in a running instance of LARSA 4D.
* A 64-bit version of Python can be used.

It has the following limitations:

* It is slower than the first approach when creating or accessing a very large amount of project or results information.
* Only the project open in the running instance of LARSA 4D can be modified.

Example: [LARSA4D_Macro_Approach2.py](LARSA4D_Macro_Approach2.py)

### Preliminary concrete anchor assessment

`Concrete_Anchor_Capacity.ipynb` is an interactive screening calculator for
anchor tension, shear, and combined utilisation. Open it in Jupyter or VS Code,
edit the assessment-inputs cell, and run the notebook cells from top to bottom.

The notebook exports a CSV report and includes an optional PDF export cell
(requires the optional `reportlab` package). This is not a complete EN 1992-4
design:
pull-out and concrete edge-breakout require characteristic group resistances
from the anchor's applicable ETA. The concrete-cone and pry-out checks use
simplified assumptions; review the notebook report and verify all inputs,
partial factors, geometry, and failure modes against the governing National
Annex and product ETA before relying on any result.
