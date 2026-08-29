using Documenter, ExcelReaders

makedocs(modules = [ExcelReaders],
	sitename = "ExcelReaders.jl",
	format = Documenter.HTML(analytics = "UA-132838790-1"),
	warnonly = [:missing_docs],
	pages = [
        "Introduction" => "index.md"
    ])

deploydocs(repo = "github.com/queryverse/ExcelReaders.jl.git")
