# ExcelReaders.jl v0.11.0 Release Notes
* Update to PyCall.jl 1.90.0

# ExcelReaders.jl v0.10.3 Release Notes
* Fix julia 1.0 compat bug

# ExcelReaders.jl v0.10.2 Release Notes
* Remove a faulty promotion rule

# ExcelReaders.jl v0.10.1 Release Notes
* Fix remaining julia 0.7/1.0 compat issues

# ExcelReaders.jl v0.10.0 Release Notes
* Drop julia 0.6 support, add julia 0.7 support

# ExcelReaders.jl v0.9.0 Release Notes
* Drop support for DataFrames.
* Use Dates.Time.
* Use DataValue for missing values.
* Fix deprecated syntax.

# ExcelReaders.jl v0.8.2 Release Notes
* Fix bug in readxlsheet

# ExcelReaders.jl v0.8.1 Release Notes
* Drop unnecessary dependencies
* Fix import of numerical column headers

# ExcelReaders.jl v0.8.0 Release Notes
* Drop julia 0.5 support.
* Fix julia 0.6 deprecation warnings.

# ExcelReaders.jl v0.7.0 Release Notes
* julia 0.6 compatability.
* Drop julia 0.4 support.
* Autogenerate names for columns without a name in readxlsheet
* Add readxlnames and readxlrange functions

# ExcelReaders.jl v0.6.0 Release Notes
* Throw more meaningful errors for ill-specified Excel ranges
* julia 0.5 compatability

# ExcelReaders.jl v0.5.0 Release Notes
* Add readxlsheet(DataFrame,...) support
* Add inline documentation
* Detect empty cells in header row

# ExcelReaders.jl v0.4.1 Release Notes
* Use pyimport_conda from PyCall to interact with conda

# ExcelReaders.jl v0.4.0 Release Notes
* Drop julia 0.3 support.
* Use conda to install xlrd dependency.

# ExcelReaders.jl v0.3.0 Release Notes
* Compatible with julia 0.4.

# ExcelReaders.jl v0.2.0 Release Notes
* Add ``readxlsheet`` function.
