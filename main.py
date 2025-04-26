import pandas as pd
from pathlib import Path
import sys
import keyboard
import random
import shutil
import os

CONFIG = {
    "input_dir": r"route_input",
    "file_prefix": "words_day",
    "file_suffix": ".xlsx",
    "required_columns":
        ["words", "definition", "times", "importance"],  # mandatory column validation
    "enable_backup": True,  # back up files or not
    "backup_dir": r"route_backup"
}


class MultiExcelLoader:
    def __init__(self, config, numbers):
        self.config = config
        self.numbers = numbers  # the number of input excel
        self.files = {}
        self._load()

    def validate_columns(self, df):
        missing = set(self.config["required_columns"]) - set(df.columns)
        if missing:
            raise ValueError(f"Missing required columns: {missing}")

    def create_backup(self, path):
        if self.config["enable_backup"]:
            backup_dir = Path(self.config["backup_dir"])
            backup_dir.mkdir(parents=True, exist_ok=True)
            backup_path = backup_dir / f"backup_{path.name}"
            shutil.copy2(path, backup_path)
            return backup_path
        return None

    def load_files(self, num):
        file_name = f"{self.config['file_prefix']}{num}{self.config['file_suffix']}"
        path = Path(self.config["input_dir"]) / file_name

        if not path.exists():
            print(f"File not found: {path}")
            return

        try:
            df = pd.read_excel(path, engine='openpyxl')
            df["origin"] = path
            df['times'] = df['times'].fillna(0)

            if self.config["required_columns"]:
                self.validate_columns(df)

            self.create_backup(path)

            self.files[file_name] = {
                "df": df,
                "original_path": path,
                "backup_path": Path(self.config["backup_dir"]) / f"backup_{file_name}",
                "modified": False
            }
            print(f"Loaded: {file_name}")

        except Exception as e:
            print(f"Failed to load {file_name}: {str(e)}")

    def _load(self):
        for num in self.numbers:
            self.load_files(num)

    def combine_dataframes(self):
        if not self.files:
            print("No data to merge")
            return None

        dfs = [info["df"] for info in self.files.values()]
        self.combined_df = pd.concat(dfs, ignore_index=True)

        if "output_combined" in self.config:
            self.combined_df.to_excel(self.config["output_combined"], index=False)
            print(f"Merged data saved to: {self.config['output_combined']}")

        return self.combined_df


def get_user_input():
    try:
        m = int(input("Enter number of files to load (m): "))
        if m <= 0:
            raise ValueError("m must be a positive integer")
    except ValueError as e:
        print(f"Input error: {e}")
        sys.exit(1)

    numbers = []
    for i in range(m):
        while True:
            try:
                num = int(input(f"Enter {i + 1}/{m} number: "))
                numbers.append(num)
                break
            except ValueError:
                print("Please enter a valid integer")

    return numbers, m


def mistake_count(j, combined_df):
    mistake_num = int(input("Correct? (1=Yes, 0=No): "))
    if mistake_num == 0:
        combined_df.at[j, "times"] += 1
    return mistake_num, combined_df


def data_back(combined_df, path_back, path_delete):
    if 'origin' not in combined_df.columns:
        raise ValueError("DataFrame missing 'origin' column")

    os.makedirs(path_back, exist_ok=True)
    grouped = combined_df.groupby('origin')

    for origin_value, group_df in grouped:
        group_finish = group_df.drop("origin", axis=1)
        filename = f"{str(origin_value)}"
        filepath = os.path.join(path_back, filename)

        group_finish.to_excel(filepath, index=False)
        print(f"Saved: {filepath}")


def present_value(combined_df, j, fault_sum):
    for i in range(len(CONFIG["required_columns"]) - 2):
        print(combined_df.at[j, CONFIG['required_columns'][i]])
        keyboard.wait('space')

    mistake, combined_df = mistake_count(j, combined_df)
    if mistake == 0:
        fault_sum += 1
    return fault_sum, combined_df


def num_judge():
    num_all = int(input("Review all words randomly? (1=All, 0=Partial): "))
    if num_all == 0:
        num_choose = int(input("How many words to review?: "))
        return num_choose


def weigh_judge(combined_df):
    return combined_df.sort_values("times", ascending=False)


def order_judge(combined_df):
    whether_weigh = int(input("Review by mistake frequency? (1=Yes, 0=No): "))
    if whether_weigh == 1:
        combined_df = weigh_judge(combined_df)

    order = int(input("Choose review order (1=Sequential/2=Random) "))
    print("Begin your review! Press spacebar to continue after each item.")
    fault_sum = 0
    words_num = 0

    if order == 1:
        for j in range(len(combined_df)):
            words_num += 1
            fault_sum, combined_df = present_value(combined_df, j, fault_sum)

    elif order == 2:
        shuffled_indices = random.sample(range(len(combined_df)), len(combined_df))
        words_end = num_judge()
        for j in shuffled_indices:
            words_num += 1
            fault_sum, combined_df = present_value(combined_df, j, fault_sum)
            if words_num == words_end:
                break

    else:
        print("Invalid input")

    print(f"Reviewed {words_num} words today with {fault_sum} errors. Keep going!")
    data_back(combined_df, path_back=CONFIG["input_dir"], path_delete=CONFIG["backup_dir"])


if __name__ == "__main__":
    numbers, m = get_user_input()
    loader = MultiExcelLoader(CONFIG, numbers)
    combined_df = loader.combine_dataframes()

    if combined_df is not None:
        order_judge(combined_df)
