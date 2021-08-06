# CSV to XLSX Example

```sh
cargo build
cargo run <input_csv> <output_xlsx>
```

You could also specify the worksheet name if wanted (otherwise "Worksheet" will be used)

```sh
cargo build
cargo run <input_csv> <output_xlsx> <worksheet_name>
```

Inspired to https://gist.github.com/lethean/fd23eb2b30fc287cee24cafd36b7db09
