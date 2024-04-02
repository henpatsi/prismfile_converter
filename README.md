# prismfile_converter
Convert and organize GraphPad Prism pzfx files into excel files to view them without the program.
Also calculates the average and SEM.

### Running

For a single file, input must be a .pzfx file and output must be a .xlsx file.

`python3 pzfx_to_excel.py <name of input file> <name of output file>`

For a directory, input must be a directory containing at least one .pzfx file, output can be any directory.

`python3 pzfx_to_excel.py <input dir> <output dir>`

(the command for running python might differ)

### Dependencies

- pandas
- openpyxl
