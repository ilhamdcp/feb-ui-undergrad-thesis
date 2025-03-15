# feb-ui-undergrad-thesis
This repository contains python code as data wrangling tools based on report generated from S&P Capital IQ.
My thesis topic discusses regarding institutional ownership effect to [corporate financialization](https://www.scirp.org/reference/referencespapers?referenceid=2474781).

## Input
The inputs are excel file, already formated with `IDX:` prefix which are reports generated from S&P Capital IQ.
Each excel report containing balance sheet and ownership information to be parsed later

## Output
The output will be generated into `output.xlsx` which contains excel file with multiple sheets.
Those sheets include *Total  Asset*, *Financial Asset* and *Institutional Ownership History*.

## How to run
1. execute `python -m venv env`
2. execute `source env/bin/activate`
3. execute `pip3 install -r requirements.txt`
4. execute `python3 crawler.py`

## What is financial asset
TODO