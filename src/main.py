import dnp as dnp
import argparse

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=".NET files parser")
    parser.add_argument('input_path', type=str, help="Path to .NET-files.")
    parser.add_argument('output_path', type=str, help="Path where results will be saved.")
    parser.add_argument('obj_name', type=str, help="Name of searching object.")
    args = parser.parse_args()

    input_path = args.input_path
    output_path = args.output_path
    obj_name = args.obj_name
    # input_path = "data\\"
    # output_path = "output\\"
    # obj_name = "DD1"

    dnp = dnp.DotNetParser(input_path, obj_name, output_path)
    dnp.run()



