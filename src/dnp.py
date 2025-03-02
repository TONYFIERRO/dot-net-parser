import glob
import os
import re
import pandas as pd


class DotNetParser:
    def __init__(self, input_path, obj_name, output_path, extension="NET"):
        self.__input_path = input_path
        self.__obj_name = obj_name
        self.__output_path = output_path
        self.__extension = extension

        raw_string = r"-\d+$"
        self.__pattern = f'^{self.__obj_name}{raw_string}'

    def run(self):
        files = self.__find_all_files(self.__input_path)
        for file in files:
            with open(os.path.join(self.__input_path, file), "r") as file_opened:
                lines = file_opened.readlines()
                lines_split = [[sub_row for sub_row in row.split(" ") if sub_row] for row in lines]
                lines_split = self.__check_on_stars(lines_split)
                df = pd.DataFrame({"Name": [], "Number": []})
                for line in lines_split:
                    for i in range(1, len(line)):
                        if re.match(self.__obj_name, line[i]):
                            obj_split = line[i].split("-")
                            if obj_split[0] == self.__obj_name:
                                obj_split[1] = obj_split[1].replace(r"\n", "")
                                df = pd.concat(
                                    [df, pd.DataFrame({"Name": [line[0]], "Number": [obj_split[1]]})],
                                    ignore_index=True)
                df = df.map(lambda x: x.replace("\n", "") if isinstance(x, str) else x)

                try:
                    os.makedirs(os.path.dirname(self.__output_path), exist_ok=True)
                    output_path = os.path.join(self.__output_path, f"{file}_output.xlsx")
                    df.to_excel(output_path, index=False)
                    print(f'File {file} is parsed. The result is saved in {file}_output.xlsx.')
                except PermissionError:
                    print(
                        "Permission Denied. Possible solution: close all active Excel documents.")
                except OSError:
                    print("OS Error. Check the output directory.")

    def __find_all_files(self, path):
        files = []
        os.chdir(path)
        for file in glob.glob(f"*.{self.__extension}"):
            files.append(file)
        os.chdir(os.path.abspath(os.path.join(os.path.dirname(__file__), '')))
        return files

    @staticmethod
    def __check_on_stars(elements):
        result = []
        for row in elements:
            if row[0] == "*":
                for elem in row[1:]:
                    result[-1].append(elem)
            else:
                result.append(row)
        return result
