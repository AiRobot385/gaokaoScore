import pandas as pd
import openpyxl
import os 
import re
import math
import time

TOTALPEOPLES = 600000
SEGMENTPEOPLES = 1000
SEGMENTGROUPS = TOTALPEOPLES // SEGMENTPEOPLES

class School:
    def __init__(self, name):
        self.name = name
        self.scores = []
        self.min_score = 750
        self.min_rank = 0
        self.avg_score = 0
        
    def add_score(self, score):
        self.scores.append(score)
        
    def set_min_score(self, min_score):
        self.min_score = min_score
        
    def set_min_rank(self, min_rank):
        self.min_rank = min_rank
    
    def get_scores(self):
        return self.scores
    
    def get_min_score(self):
        return self.min_score
    
    def get_min_rank(self):
        return self.min_rank

    def calculate_average_score(self):
        self.scores = sorted(self.scores)
        if len(self.scores) == 0:
            self.avg_score = 0
        elif len(self.scores) == 1:
            self.avg_score = int(self.scores[0]) 
        elif len(self.scores) == 2:
            self.avg_score = math.ceil(sum(self.scores) / 2) 
        else:
            trimmed_scores = self.scores[1:-1]
            avg_score = sum(trimmed_scores) / len(trimmed_scores)
            self.avg_score = math.ceil(avg_score) 


class SchoolRank:
    def __init__(self, school_name, rank):
        self.school_name = school_name
        self.rank = rank
        
class ListNode:
    def __init__(self, val=None, next=None):
        self.val = val
        self.next = next
    
class LinkedList:
    def __init__(self):
        self.head = None
        self.tail = None
    
    def insert(self, val):
        new_node = ListNode(val)
        if not self.head:
            self.head = new_node
            self.tail = new_node
        else:
            self.tail.next = new_node
            self.tail = new_node
    
    def search(self):
        result = []
        curr = self.head
        while curr:
            result.append(curr.val)
            curr = curr.next
        return result
    
class ArrayOfLinkedLists:
    def __init__(self, size):
        self.arr = [LinkedList() for i in range(size)]
    
    def insert(self, idx, val):
        self.arr[idx].insert(val)
    
    def search(self, start_idx, end_idx):
        result = []
        if start_idx >= 0 and end_idx <= len(self.arr):
            for i in range(start_idx, end_idx + 1):
                result += self.arr[i].search()
        return result



def parse_school_line_excel(filename, sheet_name, start_row):
    
    data = {}
    # 读取 Excel 表格
    df = pd.read_excel(filename, sheet_name=sheet_name)

    # 提取选定行和列的内容
    selected_cols = [1, 3, 11]
    selected_df = df.iloc[start_row-1:, selected_cols]

    # 获取列名称列表
    col_names = selected_df.columns.tolist()

    # 将第一列重命名为 'school'，第二列重命名为 'score'
    new_col_names = [0] * 3
    new_col_names[0] = 'school'
    new_col_names[1] = 'score'
    new_col_names[2] = 'comment'

    # 使用字典将列重命名
    col_dict = {col_names[i]: new_col_names[i] for i in range(len(col_names))}
    selected_df = selected_df.rename(columns=col_dict)

    # 迭代每行数据
    for index, row in selected_df.iterrows():
        # print(f"行号：{index+start_row}")
        # 跳过备注列
        if not pd.isnull(row['comment']):
            # print(row['comment'])
            continue

        pattern = re.compile(r'第\d+组')
        school_nm = re.sub(pattern, '', row['school'])
        school_score = row['score']
        # print(school_nm, school_score)
        
        if school_nm not in data:
            new_school_obj = School(school_nm)
            data[school_nm] = new_school_obj
            new_school_obj.scores.append(school_score)
        else:
            school_obj = data[school_nm]
            school_obj.scores.append(school_score)

    return data

    # for key, obj in data.items():
    #     print(key, obj.scores, obj.min_score)



def parse_score_rank_excel(filename, start_row, first_col=True, third_col=True):
    """
    Parse an Excel file and return a dictionary with the specified columns.

    Args:
        filename (str): The name of the Excel file to parse.
        start_row (int): The row number to start reading from.
        first_col (bool): Whether or not to include values from the first column.
        third_col (bool): Whether or not to include values from the third column.

    Returns:
        dict: A dictionary that maps items from the first column to items from the third column.
    """

    wb = openpyxl.load_workbook(filename)
    ws = wb.active
    data = {}

    for row in ws.iter_rows(min_row=start_row):
        key = row[0].value if first_col else None
        value = row[2].value if third_col else None
        if key is not None and value is not None:
            data[int(key)] = int(value)

    return data


def generate_school_ranked(score_rankings_map, school_scoreline_map, score_type=True):

    school_ranked = {}
    # set school obj averge score 
    for key, obj in school_scoreline_map.items():
        obj.calculate_average_score()

    for school_nm, school_obj in school_scoreline_map.items():
        score = school_obj.avg_score
        if score not in score_rankings_map:
            continue
        rank = score_rankings_map[score]
        school_ranked[school_nm] = rank
        school_obj.min_rank = rank
        # print(school_obj.name, school_obj.scores, school_obj.avg_score, school_obj.min_rank)
    
    return school_ranked


def generate_rank2school_arrlist(school_ranked_map):
    arr = ArrayOfLinkedLists(SEGMENTGROUPS)

    school_rank_list = list(school_ranked_map.items())
    tuple_lists = sorted(school_rank_list, key=lambda x: x[1])

    for tuple_item in tuple_lists:
        school_nm = tuple_item[0]
        rank = tuple_item[1]
        index = rank // SEGMENTPEOPLES
        arr.insert(index, SchoolRank(school_nm, rank))    
    return arr

def rank_range_school(rank1, rank2, schoolist):
    rank1 = rank1 // SEGMENTPEOPLES
    rank2 = rank2 // SEGMENTPEOPLES

    res = schoolist.search(rank1, rank2)
    for item in res:
        print(item.school_name, item.rank)

def find_matching_entry(dictionary, needle):
    result = []
    for key, value in dictionary.items():
        if needle in str(key):
            result.append((key, value))
    result = sorted(result, key=lambda x: x[1])
    for item in result:
        print(item[0], item[1])


def search_school(map):
    str = input()
    find_matching_entry(map, str)

def search_rank_range_school(arr):
    r1 = int(input())
    r2 = int(input())
    rank_range_school(r1, r2, arr)



def main():
    school_line_excel = os.path.abspath("data/school_line/2022_school_score_line.xlsx")
    score_rank_excel = os.path.abspath("data/score_rank/2022_score_rank.xlsx")

    score_rank_map = parse_score_rank_excel(score_rank_excel, 2)
    school_scoreline_map = parse_school_line_excel(school_line_excel, "Sheet1", 2)

    school_ranked_map = generate_school_ranked(score_rank_map, school_scoreline_map)

    school_arr = generate_rank2school_arrlist(school_ranked_map)

    # search_school(school_ranked_map)
    search_rank_range_school(school_arr)


if __name__ == "__main__":
    main()


