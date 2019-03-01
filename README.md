# JSON2XLSX - Convert JSON to Excel format

## Basic invocation

    $ json2xlx input.json
    $ open out.xlsx


## Lines mode

Sometimes a file is not inherently a well formed json file, instead of getting

    [
        {"name:" "Bruce Wayne"},
        {"name:" "Vicky Vale"}
    ]

You may get just

    {"name:" "Bruce Wayne"}
    {"name:" "Vicky Vale"}

In this case each line in the file in well formed but the file as a whole is not. For
this case use the --lines parameter.

