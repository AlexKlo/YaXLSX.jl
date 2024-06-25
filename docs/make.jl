using YaXLSX
using Documenter

DocMeta.setdocmeta!(YaXLSX, :DocTestSetup, :(using YaXLSX); recursive = true)

makedocs(;
    modules = [YaXLSX],
    sitename = "YaXLSX.jl",
    format = Documenter.HTML(;
        repolink = "https://github.com/AlexKlo/YaXLSX.jl",
        canonical = "https://AlexKlo.github.io/YaXLSX.jl",
        edit_link = "master",
        assets = ["assets/favicon.ico"],
        sidebar_sitename = true,  # Set to 'false' if the package logo already contain its name
    ),
    pages = [
        "Home"    => "index.md",
        "API Reference" => "pages/content.md",
    ],
    warnonly = [:doctest, :missing_docs],
)

deploydocs(;
    repo = "github.com/AlexKlo/YaXLSX.jl",
    devbranch = "master",
    push_preview = true,
)
