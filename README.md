# Tube labels

A script to generate a sheet of circle labels for 12" x 16" sticker paper -- 1/2" diameter or 3/8" diameter.

## Getting Started
Make sure you have Python 3 installed. Download tube_labels.py and run in the command line.

## Examples
Generate 1 sheet (192 labels) starting at A1
`$ python tube_labels.py A1`

Generate all stickers A1-A1000
`$ python tube_labels.py A1 --total 1000`
or
`$ python tube_labels.py A1 --end A1000`

Make all the stickers small (3/8" diameter)
`$ python tube_labels.py A1 --size small`

Generate stickers that are all text
`$ python tube_labels.py Taq --text`

Save output in a specific place
`$ python tube_labels.py A1 --outfile stickers.xlsx`

For more help:
`$ python tube_labels.py --help`
