import openpyxl
import pandas as pd 

# 데이터 로드
path_taga = r'C:\Users\kdp\OneDrive\바탕 화면\KDT_7\02.PANDAS\miniproject\taga_data.xlsx'
taga_DF = pd.read_excel(path_taga)

# 팀명을 기준 인덱스로 모든 열 추출
# 해당 연도를 행으로 삽입
taga_DF = taga_DF.set_index('연도')

# 승리리 데이터 로드
path_victory = r'C:\Users\kdp\OneDrive\바탕 화면\KDT_7\02.PANDAS\miniproject\victory.xlsx'

# 해당 연도를 행으로 삽입
vic_DF = pd.read_excel(path_victory, header=None, names=['순위', '팀명', '연도', '가산점'], index_col='연도')

AVG_name = [ ]
R_name = [ ]

# 2001년부터 2024년까지의 데이터 분석
for year in range(2001, 2025):
    
    # 타자 AVG 데이터 추출
    taga_DF_year = taga_DF.loc[year]                 # 해당 연도에 해당 되는 데이터 추출
    taga_DF_year.reset_index(inplace=True)            # 해당 연도에 해당 되는 데이터가 모두 추출 되었으므로 연도를 중심으로 한 인덱스 리셋
    taga_DF_year.set_index('팀명', inplace=True)       # 팀을 기준으로 인덱스 재설정 => 나중에 팀명으로 리그 순위를 반영하여 상관관계를 계산하기 위함
    taga_DF_year = taga_DF_year.drop('연도', axis=1)  # 연도는 이제 불필요한 데이터 이므로 삭제
    taga_DF_year = taga_DF_year.drop('G', axis=1)     # 게임 횟수는 상관관계 분석 시 nan이 나오므로 삭제

    # 승리 점수수 데이터 추출
    vic_DF_year = vic_DF.loc[year]                      # 해당 연도에 해당 되는 데이터 추출
    vic_DF_year.reset_index(inplace=True)               # 해당 연도에 해당 되는 데이터가 모두 추출 되었으므로 연도를 중심으로 한 인덱스 리셋
    vic_DF_year.set_index('팀명', inplace=True)         # 팀을 기준으로 인덱스 재설정 => 나중에 팀명으로 리그 순위를 반영하여 상관관계를 계산하기 위함
    vic_DF_year = vic_DF_year.drop('연도', axis=1)      # 연도는 이제 불필요한 데이터 이므로 삭제
    vic_DF_year = vic_DF_year.drop('순위', axis=1)      # 순위위는 중요하지 않다고 판단하여 삭제

    # 데이터 확인
    merged_df = pd.concat([taga_DF_year, vic_DF_year], axis=1) # 이제 taga_DF의 데이터 프레임에 승리 점수 가산점의 시리즈를 합합니다.

    # 상관관계 분석
    correlation = merged_df.corr() # 상관분석 시작!!
    AVG_name.append(correlation.loc['AVG','가산점'])
    R_name.append(correlation.loc['R','가산점'])
    correlation_value = (correlation[correlation['가산점']>0.5])['가산점'].drop(index=['가산점']) # 유의미한 상관관계 지표를 알기 위한 조건문 
    
    print(f"{year}년의 상관관계 분석 결과:")
    print(correlation)
    print()
    print(f"{year}년의 강한 상관관계를 나타내는 지표")
    print(correlation_value)   
    print()

    # 결과 해석!!!!
    # 칼럼명이 가산점인 부분의 행이나 열을 보면 각 요인별 우승에 대한 상관계수를 확인 할 수 있음!!
     
    # 상관관계 분석 결과 저장
    output_path = r'C:\Users\kdp\OneDrive\바탕 화면\KDT_7\02.PANDAS\miniproject\correlation_results.csv'
    correlation.to_csv(output_path, mode='a',index_label=f'{year}년')
    # 유의한 상관관계 분석 결과 저장
    output_path1 = r'C:\Users\kdp\OneDrive\바탕 화면\KDT_7\02.PANDAS\miniproject\correlation_value.csv'
    correlation_value.to_csv(output_path1, mode='a', index_label=f'{year}년')

print(f'AVG와 가산점 상관관계 24년치 평균값 : {sum(AVG_name)/len(AVG_name)}')
print(f'R와 가산점 상관관계 24년치 평균값 : {sum(R_name)/len(R_name)}')

result = pd.read_csv(output_path1, index_col=0)
print(result.describe)
# print(result.index.value_counts().nlargest(10)) # 해당 상관 계수 간에 가장 많은 결과가 나온 것