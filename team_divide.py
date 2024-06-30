import random
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

df = pd.read_csv("data.csv")
# df["years"] = df["years"].apply(lambda x: 6 if (x >= 6) & (x < 10)  else x)
df["years"] = df["years"].apply(lambda x: 10 if (x >= 10)  else x)
df["name"] = df["name"].apply(lambda x: x.replace("\u3000", ""))
teams_num = 4

# 必須ペアリングおよび分離の指定
# must_pair = ['あああ', 'いいい']
must_pair = []
# must_separate = ['AAA', 'BBB']
must_separate = []

# 初期チーム分けの関数
def initialize_teams(members, must_pair, must_separate):
    teams = {i: [] for i in range(teams_num)}
    remaining_members = members.copy()

    # 必須ペアを同じチームに配置
    pair_team = random.choice(list(teams.keys()))
    for member in must_pair:
        for m in remaining_members:
            if m[0] == member:
                teams[pair_team].append(m)
                remaining_members.remove(m)
                break

    # 必須分離を異なるチームに配置
    separate_teams = random.sample(list(teams.keys()), 2)
    for i, member in enumerate(must_separate):
        for m in remaining_members:
            if m[0] == member:
                teams[separate_teams[i]].append(m)
                remaining_members.remove(m)
                break

    # 残りのメンバーを均等に配置
    team_keys = list(teams.keys())
    team_sizes = [int(len(df)/teams_num)] * teams_num  # 25人を均等に分けるためのサイズ
    for i in range(len(df)%teams_num):
        team_sizes[i] = team_sizes[i] + 1
    print(team_sizes)
    # 各チームのサイズを均等に保つようにメンバーを割り当てる
    for i, member in enumerate(remaining_members):
        assigned = False
        for j in range(teams_num):
            if len(teams[team_keys[j]]) < team_sizes[j]:
                teams[team_keys[j]].append(member)
                assigned = True
                break
        if not assigned:
            raise ValueError("Failed to assign members evenly to teams.")

    return teams

# メンバーをリストに変換
members = list(df.itertuples(index=False, name=None))
random.shuffle(members)

teams = initialize_teams(members, must_pair, must_separate)

# 評価関数の定義
def evaluate_teams(teams):
    gender_balance = []
    level_balances = []
    years_sum = []
    first_year_distribution = []
    first_years_balance = []

    for team in teams.values():
        genders = [member[1] for member in team]
        levels = [member[2] for member in team]
        years = [member[3] for member in team]
        first_years = years.count(1)
        first_year_distribution.append(first_years)

        gender_balance.append(abs(genders.count('M') - genders.count('F')))
        level_balances.append(np.mean(levels))
        years_sum.append(sum(years))
        first_years_balance.append(np.std(first_year_distribution))
    level_balance = np.std(level_balances)
    year_balance = np.std(years_sum)

    return sum(gender_balance) + level_balance + sum(first_years_balance) + year_balance
    # 男女比(差し引き)、レベルバランス(分散)、1年目の人数(合計)、年次のバランス(分散)を評価対象に

# 制約を満たすか確認
def check_constraints(teams, must_pair, must_separate):
    pair_satisfied = any(all(any(member[0] == p for member in team) for p in must_pair) for team in teams.values())
    separate_satisfied = not any(all(any(member[0] == s for member in team) for s in must_separate) for team in teams.values())
    
    return pair_satisfied and separate_satisfied

# シミュレーテッド・アニーリング法による最適化
def simulated_annealing(teams, max_iter=10000, initial_temp=30, cooling_rate=0.85):
    current_teams = teams.copy()
    current_score = evaluate_teams(current_teams)
    temp = initial_temp

    for _ in range(max_iter):
        new_teams = {i: team.copy() for i, team in current_teams.items()}
        
        # ランダムに2人のメンバーを交換
        team_a, team_b = random.sample(new_teams.keys(), 2)
        if not new_teams[team_a] or not new_teams[team_b]:
            continue
        member_a = random.choice(new_teams[team_a])
        member_b = random.choice(new_teams[team_b])
        
        new_teams[team_a].remove(member_a)
        new_teams[team_b].remove(member_b)
        new_teams[team_a].append(member_b)
        new_teams[team_b].append(member_a)
        
        # 制約を確認
        if not check_constraints(new_teams, must_pair, must_separate):
            continue
        
        new_score = evaluate_teams(new_teams)
        
        if new_score < current_score or random.uniform(0, 1) < np.exp((current_score - new_score) / temp):
            current_teams = new_teams
            current_score = new_score
        
        temp *= cooling_rate

    return current_teams, current_score

# 最適化の実行
optimized_teams, final_score = simulated_annealing(teams)
# 結果の表示

wb = Workbook()
ws = wb.active
ws.title = "Teams"

dataframes = []
print(f"score: {final_score}")
for team_id, members in optimized_teams.items():
    sum1 = 0
    sum2 = 0
    F, M = 0, 0
    output_data = []
    print(f"Team {team_id + 1}:")
    for member in members:
        print(f"  {member}")
        sum1 += member[2]
        sum2 += member[3]
        if member[1] == 'F':
            F += 1
        else:
            M += 1
        output_data.append([f"Team {team_id + 1}", member[0], member[1], member[2], member[3]])
    print(f"level ave: {sum1/len(members)}  year sum: {sum2}  F: {F}  M: {M}")
    df = pd.DataFrame(output_data, columns=["Team", "Name", "Gender", "Level", "Year"]).sort_values("Year")
    dataframes.append(df)

start_col = 1
for df in dataframes:
    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
        for c_idx, value in enumerate(row, start_col):
            ws.cell(row=r_idx, column=c_idx, value=value)
    start_col += len(df.columns) + 1  # データフレームの列数+空白列1列分を追加
wb.save("teams_output.xlsx")