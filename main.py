import pandas as pd
from pathlib import Path
import sys
import keyboard
import random
import shutil
import os


CONFIG = {
    "input_dir": r"route.",
    "file_prefix": "words_day",
    "file_suffix": ".xlsx",
    "required_columns": ["words", "remember","definition","complement","times","importance"],  # mandatory column validation
    "enable_backup": True,  #back up files or not
    "backup_dir": r"route."
}


class MultiExcelLoader:
    def __init__(self, config, numbers):
        self.config = config
        self.numbers = numbers  #the number of input excel
        self.files = {}
        self._load()

    def validate_columns(self, df):
        missing = set(self.config["required_columns"]) - set(df.columns)
        if missing:
            raise ValueError(f"缺失必要列: {missing}")


    def create_backup(self, path):
        if self.config["enable_backup"]:
            backup_dir = Path(self.config["backup_dir"])
            backup_dir.mkdir(parents=True, exist_ok=True)
            backup_path = backup_dir / f"backup_{path.name}"
            #shutil.copy2(path, backup_path)
            path.rename(backup_path)
            return backup_path
        return None

    def load_files(self, num):
        file_name = f"{self.config['file_prefix']}{num}{self.config['file_suffix']}"
        path = Path(self.config["input_dir"]) / file_name

        if not path.exists():
            print(f"File does not exist.: {path}")
            return

        try:
            df = pd.read_excel(path, engine='openpyxl')
            df["origin"]=path
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
            print(f"已加载: {file_name}")

        except Exception as e:
            print(f"文件 {file_name} 加载失败: {str(e)}")

    def _load(self):
        for num in self.numbers:
            self.load_files(num)

    def combine_dataframes(self):
        if not self.files:
            print("没有可合并的数据")
            return None

        dfs = [info["df"] for info in self.files.values()]

        self.combined_df = pd.concat(dfs, ignore_index=True)  # seperate

        if "output_combined" in self.config:
            self.combined_df.to_excel(self.config["output_combined"], index=False)
            print(f"合并后的数据已保存至: {self.config['output_combined']}")

        return self.combined_df

def get_user_input():

    try:
        m = int(input("请输入要加载的文件数量 (m): "))
        if m <= 0:
            raise ValueError("m 必须是正整数")
    except ValueError as e:
        print(f"输入错误: {e}")
        sys.exit(1)

    numbers = []
    for i in range(m):
        while True:
            try:
                num = int(input(f"请输入第 {i + 1}/{m} 个数字: "))
                numbers.append(num)
                break
            except ValueError:
                print("请输入有效的整数")

    return numbers,m

def mistake_count(j,combined_df):
    mistake_num=int(input("这个单词是否正确（1表示正确，0表示错误）"))
    if(mistake_num==0):
        combined_df.at[j,"times"]+=1
    return mistake_num,combined_df


def data_back(combined_df,path_back):
    if 'origin' not in combined_df.columns:
        raise ValueError("DataFrame 中缺少 'origin' 列")

    os.makedirs(path_back, exist_ok=True)

    grouped = combined_df.groupby('origin')

    for origin_value, group_df in grouped:

        group_finish=group_df.drop("origin",axis=1)
        filename = f"{str(origin_value)}"
        filepath = os.path.join(path_back, filename)

        group_finish.to_excel(filepath, index=False)
        print(f"已保存：{filepath}")

def present_value(combined_df,j,fault_sum):
    print(combined_df.at[j, "words"])
    keyboard.wait('space')
    print(combined_df.at[j, "remember"])
    keyboard.wait('space')
    print(combined_df.at[j, "definition"])
    keyboard.wait('space')
    print(combined_df.at[j, "complement"])
    keyboard.wait('space')
    mistake,combined_df = mistake_count(j, combined_df)
    if mistake==0:
        fault_sum += 1
    return fault_sum,combined_df

def num_judge():
    num_all=int(input("是要乱序背所有单词吗：（1表示所有，0表示部分）"))
    if (num_all==0):
        num_choose=int(input("要背多少个单词呢"))
        return num_choose

def weigh_judge(combined_df):
    combined_df=combined_df.sort_values("times",ascending= False)
    return combined_df


def order_judge(combined_df):
    whether_weigh=int(input("是否按照错误次数背单词？（1表示按照，0表示不按照）"))
    if whether_weigh==1:
        combined_df=weigh_judge(combined_df)

    order = int(input("请输入复习单词时的顺序：（1表示正序，2表示乱序）"))
    print("欢迎开始您的背单词之旅！若中途暂停只需按下 a 键，即可停止背单词并生成反馈")
    fault_sum = 0 #错误个数
    words_num=0 #背单词的数量

    if order == 1:
        for j in range(len(combined_df)):
            words_num +=1
            fault_sum,combined_df=present_value(combined_df, j, fault_sum)
            '''
            if keyboard.is_pressed():
                print("检测到 a 键,停止背单词")
                break
            keyboard.wait('space')
            '''

    elif order == 2:
        shuffled_indices = random.sample(range(len(combined_df)), len(combined_df))
        words_end = num_judge()
        for j in shuffled_indices:
            words_num += 1
            fault_sum,combined_df=present_value(combined_df, j, fault_sum)
            '''
            if keyboard.is_pressed('a'):
                print("检测到 a 键,停止背单词")
                break
            '''
            if words_num==words_end:
                break

    else:
        print("无效的输入")
    print(f"今天你背了{words_num}个单词，错了{fault_sum}个单词，再接再厉！")
    data_back(combined_df,path_back=r"C:\Users\NUC\Desktop\gre词汇\原始文件")


if __name__ == "__main__":
    numbers, m = get_user_input()
    loader = MultiExcelLoader(CONFIG, numbers)
    combined_df = loader.combine_dataframes()

    if combined_df is not None:
        order_judge(combined_df)
