import json
from app.parser.sov_parser_ai import SOVParser


def main():
    input_path = "SOV_docs/LaBarre/LaBarre2_SOV.xlsx"  # <-- change this
    output_path = "tests/results/LaBarre2.json"  # <-- output file

    parser = SOVParser()
    result = parser.parse_excel(input_path)

    with open(output_path, "w") as f:
        json.dump(result, f, indent=2)

    print(f"JSON written to: {output_path}")


if __name__ == "__main__":
    main()
