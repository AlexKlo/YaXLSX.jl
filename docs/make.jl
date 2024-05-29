# docs/make.jl
using Documenter
using YaXLSX

makedocs(
    sitename = "YaXLSX Documentation",
    format = Documenter.HTML(),
    pages = [
        "Home" => "index.md"
    ],
)
