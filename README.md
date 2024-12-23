# SN parser

Simple command line tool for transforming excel files with given structure into the expected output structure

## Usage

### Flags

These are the default flags:

```sh
Usage of ./sn_excel:
  -input string
        input file (default "input.xlsx")
  -inputSheet string
        input sheet name (default "Sheet1")
  -output string
        output file (default "output.xlsx")
```

Changeable for example:

```sh
sn_parser -input input.xlsx -inputSheet Sheet1 -output output.xlsx
```

### Help

```sh
sn_parser -h | --help
```

Displays the options of the program
