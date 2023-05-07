using Pkg

dependencies = [
    "IJulia",
    "JuMP",
    "HiGHS",
    "DataFrames",
    "XLSX",
    "Dates",
    "Plots"
]

Pkg.add(dependencies)