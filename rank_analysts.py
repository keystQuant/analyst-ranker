'''
Keyst.

 ʕ•ﻌ•ʔ     ʕ•ﻌ•ʔ
( >  <) ♡ (>  < )
 u   u     u   u

feat. peepee
with love...

--COMMENT--
파일 이름은 'tenminscan.xlsx' 그리고 'hankyung.xlsx'로 저장한 후에 폴더 안에 넣는다.
제대론 된 형식의 파일들이 이 프로그램과 같은 폴더 안에 있다면, 프로그램을 실행시킨다.
'''
import os
import numpy as np
import pandas as pd

tenminscan_excel_file = [report for report in os.listdir() if 'tenminscan' in report][0]
hankyung_excel_file = [report for report in os.listdir() if 'hankyung' in report][0]

tenminscan = pd.read_excel(tenminscan_excel_file)
hankyung = pd.read_excel(hankyung_excel_file)

## STEP 1:
######################
#### 한경 데이터 정리 ####
######################

# 루프를 돌려 산업별 분류가 시작하는 곳을 찾는다
# 값이 1이면, 산업분류 차트 첫 번째라는 뜻이다. 인덱스값을 가져오고 루프를 종료한다
hankyung_report_start_index = 0

for index in range(hankyung['Unnamed: 2'].shape[0]):
    value = hankyung['Unnamed: 2'].iloc[index]
    if value == 1:
        hankyung_report_start_index = index
        break

print('산업 데이터 시작 인덱스: {}'.format(index))

# 리포트가 시작하는 위치 이전의 값들을 제거한 한경리포트 DataFrame을 만든다
clean_hankyung = hankyung.loc[hankyung_report_start_index:, :'Unnamed: 1']
clean_hankyung.head()

# 산업별 애널들의 이름과 랭킹을 묶어서 리스트에 넣어준다
dataset = []

new_industry_start = True # 새로운 산업 랭킹의 시작으로 보기
for index in range(clean_hankyung.shape[0]):
    row = clean_hankyung.iloc[index]
    row_list = list(row)

    if new_industry_start:
        industry = row['Unnamed: 1'] # 산업의 이름을 가져온다
        new_industry_start = False # 새로운 산업을 가져왔으니, 잠깐 산업은 고정시킨다
        continue

    # index값 행의 밸류가 모두 nan일 경우 새로운 산업의 시작으로 간주
    if pd.isnull(row_list[0]) and pd.isnull(row_list[1]):
        new_industry_start = True
        continue
    else:
        try:
            rank = row_list[0]
            affiliation = row_list[1].split('] ')[0][1:]
            analyst = row_list[1].split('] ')[1]
            full_data = '{}|{}|{}|{}'.format(industry, rank, affiliation, analyst)
            dataset.append(full_data)
        except:
            continue

# 데이터를 flat하게 바꾸어 리스트에 넣은 것을 찾기 쉽도록 딕셔너리로 바꿔준다
datadict = {}

for flatdata in dataset:
    listdata = flatdata.split('|')
    key = '{}_{}'.format(listdata[2], listdata[3])
    value = '{} {}'.format(listdata[0], listdata[1])
    if key not in datadict:
        datadict[key] = [value]
    else:
        if datadict[key] == None:
            datadict[key] = [value]
        else:
            new_value = datadict[key] + [value] # append로 리스트에 값을 추가해주면 이상하게 None이 리턴된다
            datadict[key] = new_value


## STEP 2:
#############################
#### 랭킹 데이터 리스트 만들기 ####
#############################

col_len = tenminscan['Unnamed: 2'].shape[0] # DataFrame의 길이

index = 0 # DataFrame의 첫 번째 값부터 가져와서 루프를 돌린다
rank_list = [] # 랭킹을 찾아서 리스트 안에 넣는다. (주의: DF의 인덱스 넘버와 length가 일치해야한다)
sort_list = [] # 소팅 리스트를 만들어서, 랭킹이 높은 애널리스트들의 항목을 df에서 위로 올린다

while index < col_len:
    analyst = tenminscan['Unnamed: 2'].iloc[index] # 애널리스트의 이름을 가져온다
    if pd.isnull(analyst) or analyst == '작성자':
        # 애널리스트의 이름이 아닌 nan 혹은 증권사의 이름일 경우 루프를 넘긴다
        rank_list.append(np.nan) # 루프를 그냥 넘기기 전에 rank_list에 nan값을 채워준다
        sort_list.append(999) # 999를 넣어서 랭킹이 낮은 항목을 위로 올린다
        index += 1
        continue
    # 진짜 데이터 프로세싱 시작:
    else:
        affiliation = tenminscan['Unnamed: 2'].iloc[index + 1][:2] # 앞 두 글자만 가져오기
        analyst = analyst.split(' ')[0] if ' ' in analyst else analyst # '이름 외 2명' 같은 형식의 데이터 처리
        analyst = analyst.split('.') if '.' in analyst else [analyst] # 두 명 이상인 경우 리스트로 스플릿하고, 아니면 하나의 엘레먼트만 있는 리스트이다

        print('**DATA: {}, {}'.format(analyst, affiliation))

        ranks = [] # 소팅 리스트에 추가하기 위해서 그 행 데이터 총 랭킹의 최저 랭킹을 뽑아낸다
        final_data = ''
        for a in analyst:
            # datadict를 들고와서 키값을 찾는다
            for key, value in datadict.items():
                if (affiliation in key) and (a in key):
                    # 소속과 애널리스트가 모두 일치한다면...
                    print('{}: {}'.format(key, value)) # 값 프린트 (디버깅용)
                    print('데이터 있음, 랭킹 데이터 가져오기')
                    # 밸류가 None이 아니면 랭킹 데이터를 넣어준다
                    all_values = ''
                    if len(value) != 1:
                        for v in value:
                            all_values = '{0} {1} ({2})'.format(all_values, v, a)
                            ranks.append(v.split(' ')[-1])
                    else:
                        all_values = '{0} ({1})'.format(value[0], a)
                        ranks.append(value[0].split(' ')[-1])
                    break
                else: # if (affiliation in key) and (a in key):
                    # 딕셔너리 키를 모두 돌려 확인하기 때문에 여러번 돈다
                    all_values = ''
            ##### for key, value in datadict.items(): #####
            final_data = final_data + all_values + ' '
        ##### for a in analyst: 여기서 끝 #####

        print('리턴 데이터: {}'.format(final_data))

        rank_list.append(final_data)
        rank_list.append(np.nan)
        if len(ranks) != 0:
            sort_list.append(min(ranks))
            sort_list.append(999)
        else:
            sort_list.append(999)
            sort_list.append(999)
        index += 2
        print('\n')

# 새로운 열을 만들어서 랭킹 데이터 추가하기
tenminscan['랭킹'] = rank_list
tenminscan['sort'] = [int(rank) for rank in sort_list]

# 엑셀로 저장하기
tenminscan.sort_values(by='sort', ascending=True, inplace=True)
tenminscan = tenminscan.drop(['sort'], axis=1)
writer = pd.ExcelWriter('output.xlsx')
tenminscan.to_excel(writer, 'Sheet1')
writer.save()
