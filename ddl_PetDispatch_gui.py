import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import pandas as pd
import itertools
from tqdm import tqdm
from typing import List, Dict, Tuple, Optional
import time
from concurrent.futures import ProcessPoolExecutor, as_completed
import multiprocessing

# 技能分值映射
skill_score_map = {
    'C':7,
    'B':12,
    'A':17,
    'S':22
}

def check_dependencies():
    try:
        import pandas
        import openpyxl
    except ImportError:
        return False
    return True

def read_pet_list() -> List[Dict]:
    """读取宠物列表.xlsx文件"""
    try:
        df = pd.read_excel('./data/宠物列表.xlsx', sheet_name='Sheet1')
    except FileNotFoundError:
        raise FileNotFoundError("未找到./data/宠物列表.xlsx文件，请确保文件存在于data文件夹中。")
    pets = []
    # 稀有度基础分映射
    rarity_base_map = {
        '普通宠物':2,
        '高级宠物':2,
        '稀有宠物':3,
        '传说宠物':5
    }
    for idx, row in df.iterrows():
        pet_name = row.iloc[0]
        if pd.isna(pet_name):
            continue
        rarity = row.iloc[1]
        skill1 = row.iloc[2]
        skill1_level = row.iloc[3]
        skill2 = row.iloc[4]
        skill2_level = row.iloc[5]
        base_score = rarity_base_map.get(rarity, 2)
        skills = {}
        if pd.notna(skill1) and pd.notna(skill1_level):
            skills[skill1] = skill_score_map.get(skill1_level, 0)
        if pd.notna(skill2) and pd.notna(skill2_level):
            skills[skill2] = skill_score_map.get(skill2_level, 0)
        pets.append({
            'id': idx+1,
            'name': pet_name,
            'rarity': rarity,
            'base_score': base_score,
            'skills': skills,
            'is_borrowed': False  # 默认不是借用的
        })
    return pets

def read_regions() -> Dict[str, List[Dict]]:
    """读取跑腿地区.xlsx文件"""
    try:
        df = pd.read_excel('./data/跑腿地区.xlsx', sheet_name='Sheet1')
    except FileNotFoundError:
        raise FileNotFoundError("未找到./data/跑腿地区.xlsx文件，请确保文件存在于data文件夹中。")
    regions = {}
    current_region = None
    for idx, row in df.iterrows():
        region_name = row.iloc[0]
        if pd.notna(region_name):
            current_region = region_name
            if current_region not in regions:
                regions[current_region] = []
        if current_region is None:
            continue
        area = row.iloc[1]
        task = row.iloc[2]
        bonus1 = row.iloc[3]
        bonus2 = row.iloc[4]
        bonus_skills = []
        if pd.notna(bonus1):
            bonus_skills.append(bonus1)
        if pd.notna(bonus2):
            bonus_skills.append(bonus2)
        if pd.notna(area) and pd.notna(task):
            regions[current_region].append({
                'area': area,
                'task': task,
                'bonus_skills': bonus_skills,
                'id': len(regions[current_region])  # 任务ID
            })
    # 确保每个区域有5个任务，不足的用空任务填充
    for region in regions:
        while len(regions[region]) <5:
            regions[region].append({
                'area': '',
                'task': '',
                'bonus_skills': [],
                'id': len(regions[region])
            })
    return regions

def precompute_pet_task_scores(pets: List[Dict], tasks: List[Dict]) -> Dict[int, Dict[int, int]]:
    """预计算每个宠物对每个任务的得分"""
    pet_task_scores = {}
    for pet in pets:
        pet_scores = {}
        for task in tasks:
            total = 0
            for skill, score in pet['skills'].items():
                if skill in task['bonus_skills']:
                    total += score
            pet_scores[task['id']] = total if total !=0 else pet['base_score']
        pet_task_scores[pet['id']] = pet_scores
    return pet_task_scores

def calculate_team_score(combo: List[Dict], task: Dict, pet_task_scores: Dict[int, Dict[int, int]]) -> int:
    """计算宠物组合的总得分"""
    return sum([pet_task_scores[pet['id']][task['id']] for pet in combo])

def generate_task_combinations(tasks: List[Dict], task_count: int) -> List[Tuple[Dict]]:
    """生成指定数量的任务组合，优先选择加成技能多的任务"""
    # 过滤掉空任务
    valid_tasks = [task for task in tasks if task['task']]
    # 按加成技能数量排序，优先计算加成技能多的任务组合
    valid_tasks.sort(key=lambda x: len(x['bonus_skills']), reverse=True)
    # 生成指定数量的任务组合
    if len(valid_tasks) >= task_count:
        return list(itertools.combinations(valid_tasks, task_count))
    else:
        # 如果有效任务数量不足，返回空列表
        return []

def assign_no_borrow(task_list: List[Dict], used_pet_mask: int, current_assignments: List[Dict],
                     available_pets: List[Dict], pet_task_scores: Dict[int, Dict[int, int]],
                     task_max_scores_no_borrow: Dict[int, int],
                     best_score: Dict, best_assignments: List[List[Dict]], all_special_found: List[bool]):
    """尝试不使用借用宠物的全特级方案"""
    # 如果已经找到所有任务都达到特级的方案，直接返回
    if all_special_found[0]:
        return
    if not task_list:
        # 计算总得分
        total = sum([assign['score'] for assign in current_assignments])
        # 计算借用的宠物数量
        borrowed = sum([1 for assign in current_assignments for pet in assign['team'] if pet.get('is_borrowed', False)])
        total_pets = sum([len(assign['team']) for assign in current_assignments])
        # 检查是否所有任务都达到特级
        all_special = all([assign['score'] > 37 for assign in current_assignments])
        # 如果找到所有任务都达到特级的方案，标记并保存
        if all_special:
            all_special_found[0] = True
            best_score['total'] = total
            best_score['borrowed'] = borrowed
            best_score['total_pets'] = total_pets
            best_assignments.clear()
            best_assignments.append([a.copy() for a in current_assignments])
        return
    # 处理当前任务
    current_task = task_list[0]
    # 剪枝：如果当前任务不使用借用宠物的最大可能得分都无法达到特级，直接返回
    if task_max_scores_no_borrow[current_task['id']] <= 37:
        return
    # 生成可用宠物列表（只使用自有宠物）
    available = []
    for pet in available_pets:
        if not pet.get('is_borrowed', False):
            # 只使用自己的宠物，且未被使用（使用位运算检查）
            if not (used_pet_mask & (1 << (pet['id'] - 1))):
                available.append(pet)
    # 计算当前任务的最大可能得分（使用可用宠物）
    current_task_max = 0
    pet_scores = []
    for pet in available:
        pet_scores.append(pet_task_scores[pet['id']][current_task['id']])
    pet_scores.sort(reverse=True)
    if pet_scores:
        current_task_max = sum(pet_scores[:min(3, len(pet_scores))])
    # 剪枝：如果当前任务的最大可能得分都无法达到特级，直接返回
    if current_task_max <= 37:
        return
    # 先计算1-3只宠物的组合，优先尝试能达到特级的组合
    # 先生成所有可能的组合，并按得分从高到低排序
    valid_combos = []
    # 优先尝试1只宠物的组合
    for pet in available:
        score = pet_task_scores[pet['id']][current_task['id']]
        if score > 37:
            valid_combos.append(([pet], score, 1, 0))  # 0表示没有借用宠物
    # 尝试2只宠物的组合
    if len(valid_combos) == 0 or current_task_max > max([c[1] for c in valid_combos], default=0):
        if len(available) >=2:
            for combo in itertools.combinations(available, 2):
                score = calculate_team_score(combo, current_task, pet_task_scores)
                if score > 37:
                    valid_combos.append((list(combo), score, 2, 0))
    # 尝试3只宠物的组合
    if len(valid_combos) == 0 or current_task_max > max([c[1] for c in valid_combos], default=0):
        if len(available) >=3:
            for combo in itertools.combinations(available, 3):
                score = calculate_team_score(combo, current_task, pet_task_scores)
                if score > 37:
                    valid_combos.append((list(combo), score, 3, 0))
    # 如果没有能达到特级的组合，返回
    if not valid_combos:
        return
    # 按得分从高到低排序，优先尝试得分高的组合
    valid_combos.sort(key=lambda x: (-x[1], x[2]))
    for combo, score, combo_size, combo_borrowed in valid_combos:
        # 检查任务中没有相同的宠物
        pet_names = [pet['name'] for pet in combo]
        if len(set(pet_names)) != len(pet_names):
            continue
        # 检查自有宠物是否已经被使用（使用位运算）
        combo_owned_ids = [pet['id'] for pet in combo if not pet.get('is_borrowed', False)]
        conflict = False
        new_used_mask = used_pet_mask
        for pet_id in combo_owned_ids:
            if used_pet_mask & (1 << (pet_id - 1)):
                conflict = True
                break
            new_used_mask |= (1 << (pet_id - 1))
        if conflict:
            continue
        # 记录分配
        new_assignments = current_assignments + [
            {
                'task': current_task,
                'team': combo,
                'score': score
            }
        ]
        # 递归处理下一个任务
        assign_no_borrow(task_list[1:], new_used_mask, new_assignments,
                         available_pets, pet_task_scores, task_max_scores_no_borrow,
                         best_score, best_assignments, all_special_found)
        # 如果已经找到全特级方案，直接返回
        if all_special_found[0]:
            return

def assign_with_borrow(task_list: List[Dict], used_pet_mask: int, current_assignments: List[Dict],
                       available_pets: List[Dict], pet_task_scores: Dict[int, Dict[int, int]],
                       task_max_scores: Dict[int, int],
                       best_score: Dict, best_assignments: List[List[Dict]], all_special_found: List[bool]):
    """尝试使用借用宠物的全特级方案"""
    # 如果已经找到所有任务都达到特级的方案，直接返回
    if all_special_found[0]:
        return
    if not task_list:
        # 计算总得分
        total = sum([assign['score'] for assign in current_assignments])
        # 计算借用的宠物数量
        borrowed = sum([1 for assign in current_assignments for pet in assign['team'] if pet.get('is_borrowed', False)])
        total_pets = sum([len(assign['team']) for assign in current_assignments])
        # 检查是否所有任务都达到特级
        all_special = all([assign['score'] > 37 for assign in current_assignments])
        # 如果找到所有任务都达到特级的方案，标记并保存
        if all_special:
            all_special_found[0] = True
            best_score['total'] = total
            best_score['borrowed'] = borrowed
            best_score['total_pets'] = total_pets
            best_assignments.clear()
            best_assignments.append([a.copy() for a in current_assignments])
        return
    # 处理当前任务
    current_task = task_list[0]
    # 剪枝：如果当前任务的最大可能得分都无法达到特级，直接返回
    if task_max_scores[current_task['id']] <= 37:
        return
    # 生成可用宠物列表
    available = []
    for pet in available_pets:
        if pet.get('is_borrowed', False):
            # 农场的宠物都可用
            available.append(pet)
        else:
            # 自己的宠物未被使用才可用（使用位运算检查）
            if not (used_pet_mask & (1 << (pet['id'] - 1))):
                available.append(pet)
    # 计算当前任务的最大可能得分（使用可用宠物）
    current_task_max = 0
    pet_scores = []
    for pet in available:
        pet_scores.append(pet_task_scores[pet['id']][current_task['id']])
    pet_scores.sort(reverse=True)
    if pet_scores:
        current_task_max = sum(pet_scores[:min(3, len(pet_scores))])
    # 剪枝：如果当前任务的最大可能得分都无法达到特级，直接返回
    if current_task_max <= 37:
        return
    # 先计算1-3只宠物的组合，优先尝试能达到特级的组合
    # 先生成所有可能的组合，并按得分从高到低排序
    valid_combos = []
    # 优先尝试1只宠物的组合（优先自有宠物）
    for pet in available:
        score = pet_task_scores[pet['id']][current_task['id']]
        if score > 37:
            borrowed = 1 if pet.get('is_borrowed', False) else 0
            valid_combos.append(([pet], score, 1, borrowed))
    # 尝试2只宠物的组合
    if len(valid_combos) == 0 or current_task_max > max([c[1] for c in valid_combos], default=0):
        if len(available) >=2:
            for combo in itertools.combinations(available, 2):
                score = calculate_team_score(combo, current_task, pet_task_scores)
                if score > 37:
                    borrowed = sum(1 for pet in combo if pet.get('is_borrowed', False))
                    valid_combos.append((list(combo), score, 2, borrowed))
    # 尝试3只宠物的组合
    if len(valid_combos) == 0 or current_task_max > max([c[1] for c in valid_combos], default=0):
        if len(available) >=3:
            for combo in itertools.combinations(available, 3):
                score = calculate_team_score(combo, current_task, pet_task_scores)
                if score > 37:
                    borrowed = sum(1 for pet in combo if pet.get('is_borrowed', False))
                    valid_combos.append((list(combo), score, 3, borrowed))
    # 如果没有能达到特级的组合，返回
    if not valid_combos:
        return
    # 按得分从高到低排序，优先尝试得分高的组合，其次是借用宠物少的组合
    valid_combos.sort(key=lambda x: (-x[1], x[3], x[2]))
    # 检查每个任务最多借1只农场宠物
    # 检查总共借用的农场宠物不超过3只
    current_borrowed_total = sum(1 for assign in current_assignments for pet in assign['team'] if pet.get('is_borrowed', False))
    for combo, score, combo_size, combo_borrowed in valid_combos:
        # 检查每个任务最多借1只农场宠物
        if combo_borrowed > 1:
            continue
        # 检查总共借用的农场宠物不超过3只
        if current_borrowed_total + combo_borrowed > 3:
            continue
        # 检查任务中没有相同的宠物
        pet_names = [pet['name'] for pet in combo]
        if len(set(pet_names)) != len(pet_names):
            continue
        # 检查自有宠物是否已经被使用（使用位运算）
        combo_owned_ids = [pet['id'] for pet in combo if not pet.get('is_borrowed', False)]
        conflict = False
        new_used_mask = used_pet_mask
        for pet_id in combo_owned_ids:
            if used_pet_mask & (1 << (pet_id - 1)):
                conflict = True
                break
            new_used_mask |= (1 << (pet_id - 1))
        if conflict:
            continue
        # 记录分配
        new_assignments = current_assignments + [
            {
                'task': current_task,
                'team': combo,
                'score': score
            }
        ]
        # 递归处理下一个任务
        assign_with_borrow(task_list[1:], new_used_mask, new_assignments,
                           available_pets, pet_task_scores, task_max_scores,
                           best_score, best_assignments, all_special_found)
        # 如果已经找到全特级方案，直接返回
        if all_special_found[0]:
            return

def assign_normal(task_list: List[Dict], used_pet_mask: int, current_assignments: List[Dict],
                  available_pets: List[Dict], pet_task_scores: Dict[int, Dict[int, int]],
                  best_score: Dict, best_assignments: List[List[Dict]]):
    """普通最优方案寻找"""
    if not task_list:
        # 所有任务分配完毕，计算总得分
        total = sum([assign['score'] for assign in current_assignments])
        borrowed = sum([1 for assign in current_assignments for pet in assign['team'] if pet.get('is_borrowed', False)])
        total_pets = sum([len(assign['team']) for assign in current_assignments])
        # 比较是否更优
        if total > best_score['total']:
            best_score['total'] = total
            best_score['borrowed'] = borrowed
            best_score['total_pets'] = total_pets
            best_assignments.clear()
            best_assignments.append([a.copy() for a in current_assignments])
        elif total == best_score['total']:
            if borrowed < best_score['borrowed']:
                best_score['borrowed'] = borrowed
                best_score['total_pets'] = total_pets
                best_assignments.clear()
                best_assignments.append([a.copy() for a in current_assignments])
            elif borrowed == best_score['borrowed']:
                if total_pets < best_score['total_pets']:
                    best_score['total_pets'] = total_pets
                    best_assignments.clear()
                    best_assignments.append([a.copy() for a in current_assignments])
                elif total_pets == best_score['total_pets']:
                    best_assignments.append([a.copy() for a in current_assignments])
        # 处理完成后返回
        return
    # 剪枝：计算当前总得分加上剩余任务的最大可能得分，如果小于已有的最佳得分，提前终止
    current_total = sum([assign['score'] for assign in current_assignments])
    # 生成可用宠物列表
    available = []
    for pet in available_pets:
        if pet.get('is_borrowed', False):
            # 农场的宠物都可用
            available.append(pet)
        else:
            # 自己的宠物未被使用才可用（使用位运算检查）
            if not (used_pet_mask & (1 << (pet['id'] - 1))):
                available.append(pet)
    # 计算剩余任务的最大可能得分
    remaining_max = 0
    for task in task_list:
        # 计算单个任务的最大可能得分：3只最高得分宠物的得分之和
        pet_scores = []
        for pet in available:
            pet_scores.append(pet_task_scores[pet['id']][task['id']])
        pet_scores.sort(reverse=True)
        remaining_max += sum(pet_scores[:3]) if pet_scores else 0
    # 如果当前总得分加上剩余最大得分小于等于已有最佳得分，提前终止
    if best_score['total'] != -1 and current_total + remaining_max <= best_score['total']:
        return
    # 处理当前任务
    current_task = task_list[0]
    # 先计算2只宠物的最大得分，判断是否需要尝试3只宠物
    max_two_score = 0
    # 先检查是否有至少2只可用宠物
    if len(available) >=2:
        for combo in itertools.combinations(available, 2):
            score = calculate_team_score(combo, current_task, pet_task_scores)
            if score > max_two_score:
                max_two_score = score
    # 生成1-3只宠物的组合，当2只宠物的最大得分已经达到特级，就不需要尝试3只
    for i in range(1,4):
        # 如果是3只宠物，且2只的最大得分已经达到特级，就跳过
        if i == 3 and max_two_score > 37:
            continue
        # 如果可用宠物数量不足i只，跳过
        if len(available) < i:
            continue
        for combo in itertools.combinations(available, i):
            # 检查每个任务最多借1只农场宠物
            combo_borrowed = sum(1 for pet in combo if pet.get('is_borrowed', False))
            if combo_borrowed > 1:
                continue
            # 检查任务中没有相同的宠物
            pet_names = [pet['name'] for pet in combo]
            if len(set(pet_names)) != len(pet_names):
                continue
            # 检查总共借用的农场宠物不超过3只
            current_borrowed = sum(1 for assign in current_assignments for pet in assign['team'] if pet.get('is_borrowed', False))
            if current_borrowed + combo_borrowed > 3:
                continue
            # 检查自有宠物是否已经被使用（使用位运算）
            combo_owned_ids = [pet['id'] for pet in combo if not pet.get('is_borrowed', False)]
            conflict = False
            new_used_mask = used_pet_mask
            for pet_id in combo_owned_ids:
                if used_pet_mask & (1 << (pet_id - 1)):
                    conflict = True
                    break
                new_used_mask |= (1 << (pet_id - 1))
            if conflict:
                continue
            # 计算该组合的得分
            score = calculate_team_score(combo, current_task, pet_task_scores)
            # 记录分配
            new_assignments = current_assignments + [
                {
                    'task': current_task,
                    'team': list(combo),
                    'score': score
                }
            ]
            # 递归处理下一个任务
            assign_normal(task_list[1:], new_used_mask, new_assignments,
                          available_pets, pet_task_scores, best_score, best_assignments)

def calculate_best_assignment(task_combination: Tuple[Dict], available_pets: List[Dict], pet_task_scores: Dict[int, Dict[int, int]]) -> Tuple[Dict, List[List[Dict]], bool]:
    """计算给定任务组合的最佳宠物分配"""
    best_score = {'total': -1, 'borrowed': float('inf'), 'total_pets': float('inf')}
    best_assignments = []
    all_special_found = [False]  # 用列表包装，方便修改
    # 获取自有宠物的最大ID，用于位运算
    owned_pets = [pet for pet in available_pets if not pet.get('is_borrowed', False)]
    max_owned_pet_id = max([pet['id'] for pet in owned_pets], default=0)
    # 预计算每个任务的最大可能得分（使用3只最高得分宠物）
    def calculate_task_max_score(task, use_borrowed=True):
        pet_scores = []
        for pet in available_pets:
            if use_borrowed or not pet.get('is_borrowed', False):
                pet_scores.append(pet_task_scores[pet['id']][task['id']])
        pet_scores.sort(reverse=True)
        return sum(pet_scores[:3]) if pet_scores else 0
    # 预计算每个任务的最大可能得分
    task_max_scores = {}
    task_max_scores_no_borrow = {}
    for task in task_combination:
        task_max_scores[task['id']] = calculate_task_max_score(task)
        task_max_scores_no_borrow[task['id']] = calculate_task_max_score(task, use_borrowed=False)
    # 开始分阶段计算
    # 第一阶段：尝试不使用借用宠物的全特级方案
    assign_no_borrow(list(task_combination), 0, [],
                     available_pets, pet_task_scores, task_max_scores_no_borrow,
                     best_score, best_assignments, all_special_found)
    # 第二阶段：如果第一阶段没有找到，尝试使用借用宠物的全特级方案
    if not all_special_found[0]:
        assign_with_borrow(list(task_combination), 0, [],
                           available_pets, pet_task_scores, task_max_scores,
                           best_score, best_assignments, all_special_found)
    # 第三阶段：如果前两个阶段都没有找到，执行原来的逻辑寻找最优方案
    if not all_special_found[0]:
        assign_normal(list(task_combination), 0, [],
                      available_pets, pet_task_scores, best_score, best_assignments)
    return best_score, best_assignments, all_special_found[0]

def get_reward_level(score: int) -> str:
    """根据总得分获取奖励等级"""
    if score > 37:
        return "特阶"
    elif 25 < score <= 37:
        return "一阶"
    elif 13 < score <= 25:
        return "二阶"
    elif 5 < score <= 13:
        return "三阶"
    elif 1 < score <= 5:
        return "四阶"
    else:
        return "无奖励"

class DispatchCalculatorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("宠物派遣计算器")
        self.root.geometry("800x700")
        
        # 初始化变量
        self.pets = []
        self.regions = {}
        self.selected_region = None
        self.owned_pets = []
        self.farm_pets = []
        self.task_count = 1
        
        # 创建界面元素
        self.create_widgets()
        
        # 加载数据
        self.load_data()
    
    def create_widgets(self):
        # 标题标签
        title_label = ttk.Label(self.root, text="宠物派遣计算器", font=("Arial", 16))
        title_label.pack(pady=10)
        
        # 区域选择框架
        region_frame = ttk.LabelFrame(self.root, text="选择派遣区域")
        region_frame.pack(padx=10, pady=5, fill=tk.X)
        
        self.region_var = tk.StringVar()
        self.region_combobox = ttk.Combobox(region_frame, textvariable=self.region_var, state="readonly")
        self.region_combobox.pack(padx=10, pady=5, fill=tk.X)
        self.region_combobox.bind("<<ComboboxSelected>>", self.on_region_selected)
        
        # 宠物选择框架
        pet_frame = ttk.LabelFrame(self.root, text="选择拥有的宠物（可多选，按住Ctrl或Shift多选）")
        pet_frame.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)
        
        self.owned_pet_listbox = tk.Listbox(pet_frame, selectmode=tk.MULTIPLE, exportselection=0)
        self.owned_pet_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 滚动条
        owned_scroll = ttk.Scrollbar(pet_frame, orient=tk.VERTICAL, command=self.owned_pet_listbox.yview)
        owned_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.owned_pet_listbox.configure(yscrollcommand=owned_scroll.set)
        
        # 农场宠物选择框架
        farm_frame = ttk.LabelFrame(self.root, text="选择农场宠物（可多选，按住Ctrl或Shift多选）")
        farm_frame.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)
        
        self.farm_pet_listbox = tk.Listbox(farm_frame, selectmode=tk.MULTIPLE, exportselection=0)
        self.farm_pet_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 滚动条
        farm_scroll = ttk.Scrollbar(farm_frame, orient=tk.VERTICAL, command=self.farm_pet_listbox.yview)
        farm_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.farm_pet_listbox.configure(yscrollcommand=farm_scroll.set)
        
        # 任务数量选择
        task_frame = ttk.LabelFrame(self.root, text="选择任务数量")
        task_frame.pack(padx=10, pady=5, fill=tk.X)
        
        self.task_count_var = tk.StringVar()
        self.task_count_combobox = ttk.Combobox(task_frame, textvariable=self.task_count_var, values=["1", "2", "3", "4", "5"], state="readonly")
        self.task_count_combobox.current(0)
        self.task_count_combobox.pack(padx=10, pady=5, fill=tk.X)
        
        # 计算按钮
        calc_button = ttk.Button(self.root, text="计算最优派遣方案", command=self.calculate)
        calc_button.pack(padx=10, pady=10)
        
        # 结果显示区域
        result_frame = ttk.LabelFrame(self.root, text="计算结果")
        result_frame.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)
        
        self.result_text = scrolledtext.ScrolledText(result_frame, wrap=tk.WORD)
        self.result_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.result_text.config(state=tk.DISABLED)
    
    def load_data(self):
        # 检查依赖
        if not check_dependencies():
            messagebox.showerror("错误", "需要安装pandas和openpyxl库，请运行以下命令：\npip install pandas openpyxl")
            self.root.quit()
            return
        
        # 读取宠物数据
        try:
            self.pets = read_pet_list()
        except FileNotFoundError as e:
            messagebox.showerror("错误", str(e))
            self.root.quit()
            return
        
        # 读取地区数据
        try:
            self.regions = read_regions()
        except FileNotFoundError as e:
            messagebox.showerror("错误", str(e))
            self.root.quit()
            return
        
        # 填充区域下拉框
        self.region_combobox["values"] = list(self.regions.keys())
        if self.regions:
            self.region_combobox.current(0)
            self.on_region_selected(None)
        
        # 填充宠物列表
        score_to_level = {v:k for k,v in skill_score_map.items()}
        for pet in self.pets:
            skill_str = ', '.join([f"{k}({score_to_level[v]})" for k, v in pet['skills'].items()])
            self.owned_pet_listbox.insert(tk.END, f"{pet['id']}. {pet['name']} - {pet['rarity']} - 特性：{skill_str}")
            self.farm_pet_listbox.insert(tk.END, f"{pet['id']}. {pet['name']} - {pet['rarity']} - 特性：{skill_str}")
    
    def on_region_selected(self, event):
        self.selected_region = self.region_var.get()
    
    def calculate(self):
        # 获取选择的区域
        if not self.selected_region:
            messagebox.showwarning("警告", "请先选择派遣区域")
            return
        
        # 获取选择的拥有的宠物
        selected_owned_indices = self.owned_pet_listbox.curselection()
        if not selected_owned_indices:
            messagebox.showwarning("警告", "请选择至少一只拥有的宠物")
            return
        self.owned_pets = [self.pets[i] for i in selected_owned_indices]
        
        # 获取选择的农场宠物
        selected_farm_indices = self.farm_pet_listbox.curselection()
        self.farm_pets = []
        for i in selected_farm_indices:
            pet = self.pets[i].copy()
            pet['is_borrowed'] = True
            self.farm_pets.append(pet)
        
        # 获取任务数量
        try:
            self.task_count = int(self.task_count_var.get())
        except ValueError:
            messagebox.showwarning("警告", "请选择有效的任务数量")
            return
        
        # 开始计算
        self.result_text.config(state=tk.NORMAL)
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END, "正在计算，请稍候...\n")
        self.result_text.update()
        
        # 预计算宠物-任务得分矩阵
        available_pets = self.owned_pets + self.farm_pets
        tasks = self.regions[self.selected_region]
        pet_task_scores = precompute_pet_task_scores(available_pets, tasks)
        
        # 生成任务组合
        task_combinations = generate_task_combinations(tasks, self.task_count)
        if not task_combinations:
            self.result_text.insert(tk.END, f"无法生成{self.task_count}个任务组合，该区域只有{len([t for t in tasks if t['task']])}个有效任务，请选择较小的任务数量。\n")
            self.result_text.config(state=tk.DISABLED)
            return
        
        # 初始化全局最佳
        overall_best = {
            'total': -1,
            'borrowed': float('inf'),
            'total_pets': float('inf'),
            'assignments': []
        }
        
        # 记录开始时间
        start_time = time.time()
        
        # 使用并行计算
        self.result_text.insert(tk.END, f"正在使用{multiprocessing.cpu_count()}个CPU核心并行计算最优派遣方案...\n")
        self.result_text.update()
        
        # 创建进程池
        try:
            with ProcessPoolExecutor(max_workers=multiprocessing.cpu_count()) as executor:
                # 提交所有任务组合的计算
                futures = []
                future_to_task = {}  # 映射future到task_combo
                for task_combo in task_combinations:
                    valid_tasks = [task for task in task_combo if task['task']]
                    if valid_tasks:
                        future = executor.submit(
                            calculate_best_assignment,
                            valid_tasks,
                            available_pets,
                            pet_task_scores
                        )
                        futures.append(future)
                        future_to_task[future] = task_combo
                
                # 处理完成的任务
                all_special_found = False
                processed_count = 0
                self.result_text.insert(tk.END, "正在计算任务组合...\n")
                self.result_text.update()
                
                for future in as_completed(futures):
                    task_combo = future_to_task[future]
                    processed_count += 1
                    try:
                        best_score, best_assignments, combo_all_special = future.result()
                    except Exception as e:
                        self.result_text.insert(tk.END, f"\n计算任务组合时出错：{e}\n")
                        continue
                    
                    # 如果找到全特级方案，立即停止所有计算并输出结果
                    if combo_all_special:
                        # 取消所有未完成的任务
                        for f in futures:
                            if not f.done():
                                f.cancel()
                        # 记录为全局最佳方案
                        overall_best['total'] = best_score['total']
                        overall_best['borrowed'] = best_score['borrowed']
                        overall_best['total_pets'] = best_score['total_pets']
                        overall_best['assignments'] = best_assignments
                        # 跳出循环，准备输出结果
                        all_special_found = True
                        break
                    
                    # 更新全局最佳
                    if best_score['total'] > overall_best['total']:
                        overall_best['total'] = best_score['total']
                        overall_best['borrowed'] = best_score['borrowed']
                        overall_best['total_pets'] = best_score['total_pets']
                        overall_best['assignments'] = best_assignments
                    elif best_score['total'] == overall_best['total']:
                        if best_score['borrowed'] < overall_best['borrowed']:
                            overall_best['total'] = best_score['total']
                            overall_best['borrowed'] = best_score['borrowed']
                            overall_best['total_pets'] = best_score['total_pets']
                            overall_best['assignments'] = best_assignments
                        elif best_score['borrowed'] == overall_best['borrowed']:
                            if best_score['total_pets'] < overall_best['total_pets']:
                                overall_best['total'] = best_score['total']
                                overall_best['borrowed'] = best_score['borrowed']
                                overall_best['total_pets'] = best_score['total_pets']
                                overall_best['assignments'] = best_assignments
                            elif best_score['total_pets'] == overall_best['total_pets']:
                                overall_best['assignments'].extend(best_assignments)
            
            # 计算总耗时
            end_time = time.time()
            total_calc_time = end_time - start_time
            
            # 输出结果
            self.result_text.insert(tk.END, "\n===== 最优派遣方案结果 =====\n")
            self.result_text.insert(tk.END, f"✅ 计算完成！方案计算总耗时：{total_calc_time:.2f} 秒\n")
            self.result_text.insert(tk.END, f"派遣区域：{self.selected_region}\n")
            
            if not overall_best['assignments']:
                self.result_text.insert(tk.END, "没有找到有效的派遣方案。\n")
            else:
                # 取第一个最佳方案
                best_assignment = overall_best['assignments'][0]
                self.result_text.insert(tk.END, f"执行任务数量：{len(best_assignment)}\n")
                self.result_text.insert(tk.END, f"总得分：{overall_best['total']}\n")
                self.result_text.insert(tk.END, f"借用宠物数量：{overall_best['borrowed']}\n")
                self.result_text.insert(tk.END, f"总使用宠物数量：{overall_best['total_pets']}\n")
                
                # 输出每个任务
                for i, assign in enumerate(best_assignment, 1):
                    task = assign['task']
                    team = assign['team']
                    score = assign['score']
                    reward_level = get_reward_level(score)
                    
                    self.result_text.insert(tk.END, f"\n--- 任务{i} ---\n")
                    self.result_text.insert(tk.END, f"任务名称：{task['task']}\n")
                    self.result_text.insert(tk.END, f"任务区域：{task['area']}\n")
                    self.result_text.insert(tk.END, f"加成特性：{', '.join(task['bonus_skills']) if task['bonus_skills'] else '无'}\n")
                    
                    # 处理宠物名称，农场宠物加（借）
                    pet_names = []
                    for pet in team:
                        if pet.get('is_borrowed', False):
                            pet_names.append(f"{pet['name']}（借）")
                        else:
                            pet_names.append(pet['name'])
                    self.result_text.insert(tk.END, f"推荐派遣宠物：{', '.join(pet_names)}\n")
                    self.result_text.insert(tk.END, f"任务得分：{score}\n")
                    self.result_text.insert(tk.END, f"预计奖励等级：{reward_level}\n")
                
                # 如果有多个同优先方案，提示用户
                if len(overall_best['assignments']) > 1:
                    self.result_text.insert(tk.END, f"\n注：共有{len(overall_best['assignments'])}种同优先的最优方案，以上为其中一种。\n")
        
        except Exception as e:
            self.result_text.insert(tk.END, f"\n计算过程中出错：{e}\n")
        
        self.result_text.config(state=tk.DISABLED)

if __name__ == "__main__":
    # 处理打包后的多进程支持
    multiprocessing.freeze_support()
    # 设置启动方式为spawn，兼容Windows打包
    try:
        multiprocessing.set_start_method('spawn', force=True)
    except RuntimeError:
        # 如果已经设置过启动方式，忽略错误
        pass
    
    root = tk.Tk()
    app = DispatchCalculatorGUI(root)
    root.mainloop()