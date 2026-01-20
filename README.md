# Topsis-Anjani-102303480

**TOPSIS (Technique for Order of Preference by Similarity to Ideal Solution)** implementation in Python for multi-criteria decision analysis.

## Description

TOPSIS is a multi-criteria decision analysis method that ranks alternatives based on their similarity to the ideal solution. This package provides a simple command-line tool to perform TOPSIS analysis on CSV data files.

## Installation
Package listed on PyPI:- https://pypi.org/project/Topsis-Anjani-102303480
Install the package using pip:
```bash
pip install Topsis-Anjani-102303480
```

## Usage

After installation, you can use the `topsis` command from anywhere in your terminal:

```bash
topsis <InputDataFile> <Weights> <Impacts> <OutputResultFileName>
```

### Example

```bash
topsis data.csv "1,1,1,2" "+,+,-,+" result.csv
```

## Input File Format

The input CSV file must follow this structure:

- **First column**: Names of alternatives/options
- **Remaining columns**: Numeric values for each criterion
- **Minimum**: 3 columns (1 name column + at least 2 criteria)

### Example Input (`data.csv`)

```csv
Model,Price,Storage,Camera,Battery
P1,250,64,12,4000
P2,200,32,8,3500
P3,300,128,16,4500
P4,275,64,12,4200
P5,225,32,16,3800
```

## Output Format

The output CSV includes all original columns plus:

- **Topsis Score**: Score between 0 and 1 (higher is better)
- **Rank**: Ranking based on TOPSIS score (1 is best)

### Example Output (`result.csv`)

```csv
Model,Price,Storage,Camera,Battery,Topsis Score,Rank
P3,300,128,16,4500,0.691,1
P4,275,64,12,4200,0.535,2
P1,250,64,12,4000,0.534,3
P5,225,32,16,3800,0.401,4
P2,200,32,8,3500,0.308,5
```
