import sys
import os
import numpy as np
import pandas as pd

def normalize_matrix(matrix, weights):
    norm_matrix = matrix / np.sqrt((matrix**2).sum(axis=0))
    weighted_matrix = norm_matrix * weights
    return weighted_matrix

def calculate_ideal_solutions(weighted_matrix, impacts):
    ideal_best = []
    ideal_worst = []
    for i, impact in enumerate(impacts):
        if impact == '+':
            ideal_best.append(weighted_matrix[:, i].max())
            ideal_worst.append(weighted_matrix[:, i].min())
        elif impact == '-':
            ideal_best.append(weighted_matrix[:, i].min())
            ideal_worst.append(weighted_matrix[:, i].max())
    return np.array(ideal_best), np.array(ideal_worst)

def calculate_topsis(matrix, weights, impacts):
    weighted_matrix = normalize_matrix(matrix, weights)
    ideal_best, ideal_worst = calculate_ideal_solutions(weighted_matrix, impacts)
    distances_best = np.sqrt(((weighted_matrix - ideal_best)**2).sum(axis=1))
    distances_worst = np.sqrt(((weighted_matrix - ideal_worst)**2).sum(axis=1))
    performance_scores = distances_worst / (distances_best + distances_worst)
    return performance_scores

def main():
    if len(sys.argv) != 5:
        print("Usage: python topsis.py <input_file> <weights> <impacts> <output_file>")
        sys.exit(1)

    input_file = sys.argv[1]
    weights = list(map(float, sys.argv[2].split(',')))
    impacts = sys.argv[3].split(',')
    output_file = sys.argv[4]

    try:
        data = pd.read_excel(input_file)
        if data.shape[1] < 3:
            raise ValueError("Input file must have at least three columns (ID, Criteria1, Criteria2, ...).")

        criteria_matrix = data.iloc[:, 1:].values
        if not np.issubdtype(criteria_matrix.dtype, np.number):
            raise ValueError("Criteria columns must contain only numeric values.")

        if len(weights) != criteria_matrix.shape[1] or len(impacts) != criteria_matrix.shape[1]:
            raise ValueError("The number of weights and impacts must match the number of criteria.")

        if not all(impact in ['+', '-'] for impact in impacts):
            raise ValueError("Impacts must be either '+' or '-'.")

        scores = calculate_topsis(criteria_matrix, weights, impacts)
        data['TOPSIS Score'] = scores
        data['Rank'] = data['TOPSIS Score'].rank(ascending=False).astype(int)

        data.to_excel(output_file, index=False)
        print(f"Results saved to {output_file}")

    except FileNotFoundError:
        print(f"Error: The file '{input_file}' was not found.")
    except ValueError as ve:
        print(f"Error: {ve}")
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()
