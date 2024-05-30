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

deploydocs(;
    repo = "github.com/AlexKlo/YaXLSX.jl",
    devbranch = "main",
    push_preview = true,
)