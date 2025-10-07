#!/usr/bin/env python3

import openpyxl
import pandas as pd
import random
import os
from openpyxl.styles import Font

# ✅ 파일 로드 (엑셀)
file_path = 'Auto_Table.xlsx'

# ✅ Participants 시트에서 참가자 명단 읽기
df_participants = pd.read_excel(file_path, sheet_name='Participants')
male_players = df_participants['남자'].dropna().tolist()
female_players = df_participants['여자'].dropna().tolist()

num_male = len(male_players)
num_female = len(female_players)

print(f"참가자: 남 {num_male}, 여 {num_female}")

# ✅ 남여 비율별 시퀀스 표

# 혼복 6
sequence_table = {
    (6,14): [2,2,2,13,13], (7,13): [7,7,7,7,10], (8,12): [2,3,3,13,13],
    (9,11): [3,3,3,13,13], (10,10): [7,7,8,8,11], (11,9): [3,3,8,8,15],
    (12,8): [3,8,8,8,14], (13,7): [4,8,8,8,14], (14,6): [4,4,8,12,14],
    (15,5): [4,4,12,12,12], (16,4): [4,9,9,12,12],
}

# 혼복 8
sequence_table_v2 = {
    (6,14): [2,7,10,10,13], (7,13): [7,7,7,10,13], (8,12): [3,10,10,11,11],
    (9,11): [3,3,7,13,15], (10,10): [3,7,8,11,15], (11,9): [8,8,11,11,11],
    (12,8): [8,8,8,8,15], (13,7): [8,8,8,12,14], (14,6): [4,8,12,12,14],
    (15,5): [4,12,12,12,12], (16,4): [5,12,12,12,12],
}

use_v2 = True  # ✅ True면 혼복8, False면 혼복6

if use_v2:
    round_combinations = sequence_table_v2.get((num_male, num_female), None)
else:
    round_combinations = sequence_table.get((num_male, num_female), None)

if round_combinations is None:
    raise ValueError(f"⚠️ 현재 참가자 수 (남 {num_male}, 여 {num_female}) 에 대한 시퀀스가 정의되어 있지 않습니다.")
print(f"사용 시퀀스: {round_combinations}")

wb = openpyxl.load_workbook(file_path)
ws = wb['Match_schedule']

# 엑셀 로드: 3행부터 데이터 읽기
df_life = pd.read_excel(file_path, sheet_name='LIFE_members', header=None, skiprows=2)
life_members_male = ['김종현', '문광식', '박동언', '박종성', '오성목', '임채경', '정기완', '조창현', '홍상현']
life_members_female = ['김예인', '문지정', '서가연', '서자랑', '장은비', '정예원', '최은진']
life_members = life_members_male + life_members_female

# 라이프 회원 그룹 정보 로드
group_dict = {}
for name in df_life.iloc[:,1].dropna():
    group_dict[name] = 'A'
for name in df_life.iloc[:,2].dropna():
    group_dict[name] = 'B'
for name in df_life.iloc[:,3].dropna():
    group_dict[name] = 'A'
for name in df_life.iloc[:,4].dropna():
    group_dict[name] = 'B'

# ✅ 게스트 선수들을 group_dict 에 'guest' 로 등록
for p in male_players + female_players:
    if p not in group_dict:
        group_dict[p] = 'guest'

combi_table = {
    1: (0,0,4), 2: (0,1,3), 3: (0,2,2), 4: (0,3,1), 5: (0,4,0),
    6: (1,0,3), 7: (1,1,2), 8: (1,2,1), 9: (1,3,0), 10: (2,0,2),
    11: (2,1,1), 12: (2,2,0), 13: (3,0,1), 14: (3,1,0), 15: (4,0,0)
}

round_rows = [5,7,9,11,13]
all_players = male_players + female_players

# ✅ swap_if_needed 함수 (cross-pair swap)
def swap_if_needed(previous_round, current_round, max_attempts=20):
    attempt = 0
    swap_warning = False
    while attempt < max_attempts:
        need_retry = False
        for prev_team in previous_round:
            for curr_team in current_round:
                prev_players = prev_team[1]
                curr_players = curr_team[1]
                common_players = set(map(tuple, prev_players)) & set(map(tuple, curr_players))
                if len(common_players) >= 3 and len(curr_players) >= 4:
                    p1, p2, p3, p4 = curr_players[:4]
                    curr_players[0] = p1
                    curr_players[1] = p3
                    curr_players[2] = p2
                    curr_players[3] = p4
                    need_retry = True
        if not need_retry:
            break
        attempt += 1
    if attempt >= max_attempts:
        swap_warning = True
    return current_round, swap_warning


MAX_TRIALS = 100
trial = 0
while trial < MAX_TRIALS:
    trial += 1
    # ✅ 루프 시작 시 변수 초기화
    rest_count = {p:0 for p in all_players}
    game_count = {p:0 for p in all_players}
    mixed_played_men = {p:0 for p in male_players}
    mixed_played_women = {p:0 for p in female_players}
    previous_round = []
    swap_warning = False

    # ✅ main loop
    for rnd, comb_num in enumerate(round_combinations):
        mixed, men_d, women_d = combi_table[comb_num]
        men_need = mixed*2 + men_d*4
        women_need = mixed*2 + women_d*4

        rest_num = len(all_players) - (men_need + women_need)
        rest_this_round = []
        active_men = male_players.copy()
        active_women = female_players.copy()

        # ✅ 휴식자 선정
        while len(rest_this_round) < rest_num:
            candidates = [p for p in all_players if p not in rest_this_round and rest_count[p]==0]
            if not candidates:
                candidates = [p for p in all_players if p not in rest_this_round]
            p = random.choice(candidates)
            if p in active_men and len(active_men)-1 >= men_need:
                active_men.remove(p)
                rest_this_round.append(p)
                rest_count[p] += 1
            elif p in active_women and len(active_women)-1 >= women_need:
                active_women.remove(p)
                rest_this_round.append(p)
                rest_count[p] += 1

        match_list = []

        # ✅ 혼복 (모든 선수 최소 1회 참여 우선)
        for _ in range(mixed):
            unplayed_men = [p for p in active_men if mixed_played_men[p]==0]
            unplayed_women = [p for p in active_women if mixed_played_women[p]==0]

            men_pool = unplayed_men if len(unplayed_men) >= 2 else active_men
            women_pool = unplayed_women if len(unplayed_women) >= 2 else active_women

            men_pool.sort(key=lambda x: game_count[x])
            women_pool.sort(key=lambda x: game_count[x])
            random.shuffle(men_pool)
            random.shuffle(women_pool)

            men_pair = men_pool[:2]
            women_pair = women_pool[:2]

            for p in men_pair:
                mixed_played_men[p] += 1
            for p in women_pair:
                mixed_played_women[p] += 1

            team = [men_pair[0], women_pair[0], men_pair[1], women_pair[1]]
            match_list.append(('혼복', team))

            for p in men_pair + women_pair:
                game_count[p] += 1
            active_men = [m for m in active_men if m not in men_pair]
            active_women = [w for w in active_women if w not in women_pair]

        # ✅ 남복 (A/B 그룹 우선)
        for _ in range(men_d):
            active_men.sort(key=lambda x: game_count[x])
            random.shuffle(active_men)
            a_men = [p for p in active_men if group_dict.get(p) == 'A']
            guest_men = [p for p in active_men if group_dict.get(p) == 'guest']
            if len(a_men) + len(guest_men) >= 4 and len(a_men) >= 1:
                combined = a_men + guest_men
                team = combined[:4]
            else:
                b_men = [p for p in active_men if group_dict.get(p) == 'B']
                if len(b_men) + len(guest_men) >= 4 and len(b_men) >= 1:
                    combined = b_men + guest_men
                    team = combined[:4]
                else:
                    team = active_men[:4]

            match_list.append(('남복', team))
            for p in team:
                game_count[p] += 1
            active_men = [m for m in active_men if m not in team]

        # ✅ 여복 (A/B 그룹 우선)
        for _ in range(women_d):
            active_women.sort(key=lambda x: game_count[x])
            random.shuffle(active_women)
            a_women = [p for p in active_women if group_dict.get(p) == 'A']
            guest_women = [p for p in active_women if group_dict.get(p) == 'guest']
            if len(a_women) + len(guest_women) >= 4 and len(a_women) >= 1:
                combined = a_women + guest_women
                team = combined[:4]
            else:
                b_women = [p for p in active_women if group_dict.get(p) == 'B']
                if len(b_women) + len(guest_women) >= 4 and len(b_women) >= 1:
                    combined = b_women + guest_women
                    team = combined[:4]
                else:
                    team = active_women[:4]

            match_list.append(('여복', team))
            for p in team:
                game_count[p] += 1
            active_women = [w for w in active_women if w not in team]

        # ✅ swap 로직 적용
        match_list, swap_flag = swap_if_needed(previous_round, match_list)
        if swap_flag:
            swap_warning = True
        previous_round = match_list.copy()

        # ✅ 코트 배정
        match_list_sorted = sorted(match_list, key=lambda x: 0 if x[0]=='여복' else 1)
        match_players = []
        for m in match_list_sorted:
            match_players.extend(m[1])

        final_players = match_players + rest_this_round
        final_players = final_players[:20] + [None]*(20-len(final_players))

        row = round_rows[rnd]
        # 엑셀 파일에 저장할 때 라이프 멤버는 밑줄 추가
        for idx, name in enumerate(final_players):
            if name:
                gender = '(m)' if name in male_players else '(f)'
                cell = ws.cell(row=row, column=idx+3)
                cell.value = f"{name}{gender}"
                name_without_gender = name.split('(')[0]  # Extract name without gender
                if name_without_gender in life_members:
                    # Remove print statements related to adding asterisks
                    # print(f"Adding asterisk to {name_without_gender}")
                    # print(f"Cell value set to: {cell.value}")
                    cell.value = f"*{name}{gender}"
                # Set background color based on team composition
                if team.count('(f)') == 4:
                    cell.fill = openpyxl.styles.PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
                    cell.fill.opacity = 0.2
                elif team.count('(m)') == 4:
                    cell.fill = openpyxl.styles.PatternFill(start_color='0000FF', end_color='0000FF', fill_type='solid')
                    cell.fill.opacity = 0.2
                elif team.count('(f)') == 2 and team.count('(m)') == 2:
                    cell.fill = openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00', fill_type='solid')
                    cell.fill.opacity = 0.2
            else:
                ws.cell(row=row, column=idx+3).value = None

        # 매칭이 완료된 후에 라이프 멤버를 팀의 맨 앞으로 배치하고, 성별 식별자를 추가
        for match in match_list:
            for i, team in enumerate(match[1]):
                life_members_in_team = [p for p in team if p.split('(')[0] in life_members]
                non_life_members = [p for p in team if p.split('(')[0] not in life_members]
                match[1][i] = life_members_in_team + non_life_members
                # Remove debugging print statements
                # print(f"Updated team: {match[1][i]}")

        # 각 선수의 이름에 성별 식별자 추가
        for i, team in enumerate(match_list):
            players_with_gender = [f"{p} (m)" if p in male_players else f"{p} (f)" for p in team[1]]
            match_list[i] = (team[0], players_with_gender)

    # ✅ 혼복 최소 1회 미참여 선수 확인
    unplayed_men_final = [p for p,v in mixed_played_men.items() if v==0]
    unplayed_women_final = [p for p,v in mixed_played_women.items() if v==0]

    print(f"Trial {trial}: 미혼복 남={len(unplayed_men_final)}, 여={len(unplayed_women_final)}, swap_warning={swap_warning}")

    if not swap_warning and not unplayed_men_final and not unplayed_women_final:
        print("✅ 성공적으로 매칭 완료")
        break

# ✅ 파일명 자동 증가 저장
base_filename = 'LIFE_Auto_Table'
file_ext = '.xlsx'
file_path_save = base_filename + file_ext
counter = 2
while os.path.exists(file_path_save):
    file_path_save = f"{base_filename}_{counter}{file_ext}"
    counter += 1

wb.save(file_path_save)
print(f"✅ 저장 완료: {file_path_save}")
