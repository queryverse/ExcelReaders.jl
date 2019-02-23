using Documenter, ExcelReaders

makedocs(
	modules = [ExcelReaders],
	sitename = "ExcelReaders.jl",
	analytics="UA-132838790-1",
	pages = [
        "Introduction" => "index.md"
    ]
)

deploydocs(
    repo = "github.com/queryverse/ExcelReaders.jl.git"
)
