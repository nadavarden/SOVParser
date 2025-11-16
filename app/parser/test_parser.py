import json
from app.parser.sov_parser import SOVParser



def print_nested(d, indent=0):
    """Pretty print nested dicts / lists."""
    space = " " * indent
    if isinstance(d, dict):
        for k, v in d.items():
            print(f"{space}{k}:")
            print_nested(v, indent + 4)
    elif isinstance(d, list):
        for idx, item in enumerate(d):
            print(f"{space}- Item {idx+1}:")
            print_nested(item, indent + 4)
    else:
        print(f"{space}{d}")


def test_excel(path: str):
    parser = SOVParser()

    print(f"\n==============================")
    print(f"   TESTING SOV PARSER")
    print(f"   File: {path}")
    print(f"==============================\n")

    result = parser.parse_excel(path)

    # Pretty print all extracted fields
    print("\n📌 Extracted Fields and Values:")
    print("------------------------------")
    print_nested(result)


    print("\nDone.\n")


if __name__ == "__main__":
    # Put your test file name here
    test_excel("SOV_docs/LaBarre/LaBarre2_SOV.xlsx")