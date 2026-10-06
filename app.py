import streamlit as st
import pandas as pd
import datetime
import io
import gspread #구글시트용

st.set_page_config(
    page_title='성남등린이',
    layout='wide',  # 'centered'(기본값)에서 'wide'로 변경
)




# ========================================================================================= 구글시트 연결
# 1. 서비스 계정 JSON 열쇠로 구글 시트 안전하게 연결
@st.cache_resource
def get_connection():
    spreadsheet_url = st.secrets["connections"]["gsheets"]["spreadsheet"]
    
    # 만약 Streamlit Secrets에 [gspread] 정보가 있다면 (클라우드 환경)
    if "gspread" in st.secrets:
        import json
        from google.oauth2.service_account import Credentials
        
        scope = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        # st.secrets에 있는 정보를 dict 형태로 변환해서 인증 객체 생성
        creds_dict = dict(st.secrets["gspread"])
        creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
        gc = gspread.authorize(creds)
    else:
        # 로컬 환경 (컴퓨터에 service_account.json 파일이 있을 때)
        gc = gspread.service_account(filename="service_account.json")
        
    doc = gc.open_by_url(spreadsheet_url)
    return doc

if 'gs' not in st.session_state:
    st.session_state.gs = {}
    doc = get_connection()
    # --- [1] 멤버 관리 탭 읽기 ---
    member_ws = doc.worksheet("member")
    member = pd.DataFrame(member_ws.get_all_records()).reset_index(drop=True)
    member = member.sort_values(by='Index').reset_index(drop=True)
    st.session_state.gs['Member'] = member

    # --- [2] 산행기록 탭 읽기 및 저장(쓰기) ---
    hiking_ws = doc.worksheet("hiking")
    st.session_state.gs['Hiking'] = pd.DataFrame(hiking_ws.get_all_records()).reset_index(drop=True)

Member = st.session_state.gs['Member']
Hiking = st.session_state.gs['Hiking']
# st.dataframe(Member, width="stretch")
# st.dataframe(Hiking, width="stretch")




# ========================================================================================= 기타 준비물
if 'nowtime' not in st.session_state: st.session_state.nowtime = datetime.datetime.now().strftime('%Y%m%d')
else:pass
Year = int(st.session_state.nowtime[:4])
Month = int(st.session_state.nowtime[4:-2])
Day = int(st.session_state.nowtime[-2:])

if 'num' not in st.session_state: st.session_state.num = int(st.session_state.nowtime)
else:pass
NUM = st.session_state.num

def to_excel(dataframe):
    output = io.BytesIO()
    # openpyxl 엔진을 사용하여 엑셀 파일로 쓰기
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        dataframe.to_excel(writer, index=False, sheet_name='등산이력')
    processed_data = output.getvalue()
    return processed_data




# ========================================================================================= 본격적 시작
ncol1, ncol2, ncol3 = st.columns([1,4,0.4])
with ncol1:
    st.markdown("""<div style="font-size: 24px; font-weight: bold; color: black; background-color: lightgreen; padding: 1px; border-radius: 8px; text-align: center;margin-bottom: 10px;">
        【 성남등린이 】👦👧
        </div>""",unsafe_allow_html=True)
with ncol2: MODE = st.selectbox('mode', options=['ＭＯＤＥⅠ．　인원/등산이력 장부 📒','ＭＯＤＥⅡ．　산행 일지 ⛰️'], label_visibility='collapsed')
with ncol3: CODE = st.text_input('code', label_visibility='collapsed', type='password')
if CODE == 'tmd': trigger_code = 'GO'
else: trigger_code = 'STOP'


# ========================================================================================= 데이터 로딩, 재료 정리
list_hiking = ['날짜','산','참석자','리딩자','당취자','지각자']
list_member = ['ID','이름','성별','생년','가입일','Index']

# Hiking = pd.read_csv('./savefiles/산행기록.csv',encoding='cp949') # 윈도우파일
# Member = pd.read_csv('./savefiles/멤버관리.csv',encoding='cp949') # 윈도우파일

Hiking_edit = pd.DataFrame(index=Hiking.index, columns=Hiking.columns)
Hiking_edit[['날짜','산']] = Hiking[['날짜','산']]
for i in Hiking.index:
    for col in ['참석자','리딩자','당취자','지각자']:
        val_ = Hiking.loc[i,col]
        try: val_ = val_.split('|')
        except: pass
        Hiking_edit.at[i, col] = val_
# st.write(Hiking, Hiking_edit)

list_Index = []
for i in Member.index: list_Index.append(Member.loc[i,'Index'])
# st.write(list_Index)

Member_show = Member.copy()




# ========================================================================================= 모드별 페이지
if 'Ⅱ．' in MODE:  
    mcol21, mcol22 = st.columns([1,4], gap='large')
    with mcol21:
        st.markdown("""<div style="font-size: 20px; font-weight: bold; color: black; background-color: white; padding: 1px; border-radius: 8px; text-align: left;margin-bottom: 10px;">
                ⅰ）산행 추가 📌
                </div>""",unsafe_allow_html=True)

        list_scol = [1,3.6]
        scol1, scol2 = st.columns(list_scol)
        with scol1: st.write('날짜')
        with scol2: Date = st.text_input('날짜', placeholder='예시 : 20260110', label_visibility='collapsed', key='date'+str(NUM))
        scol1, scol2 = st.columns(list_scol)
        with scol1: st.write('산')
        with scol2: Mountain = st.text_input('산', placeholder='예시 : (야등)청계산_매봉', label_visibility='collapsed', key='mountain'+str(NUM))
        scol1, scol2 = st.columns(list_scol)
        with scol1: st.write('참석자')
        with scol2: Join = st.multiselect('참석자', options=list_Index, label_visibility='collapsed', key='join'+str(NUM))
        scol1, scol2 = st.columns(list_scol)
        with scol1: st.write('리딩자')
        with scol2: Lead = st.multiselect('리딩자', options=Join, label_visibility='collapsed', key='lead'+str(NUM))
        scol1, scol2 = st.columns(list_scol)
        with scol1: st.write('당취자')
        with scol2: Cancel = st.multiselect('당취자', options=Join, label_visibility='collapsed', key='cancel'+str(NUM))
        scol1, scol2 = st.columns(list_scol)
        with scol1: st.write('지각자')
        with scol2: Late = st.multiselect('지각자', options=Join, label_visibility='collapsed', key='late'+str(NUM))

        Hiking_add = pd.DataFrame(index=range(1), columns=list_hiking)
        if len(Date) > 0 and len(Mountain) > 0 and len(Join) > 0 and len(Lead) > 0 and trigger_code == 'GO':
            trigger_hiking = False
            Hiking_add.at[0,'날짜'] = int(Date)
            Hiking_add.at[0,'산'] = Mountain
            Join = '|'.join(Join)
            Lead = '|'.join(Lead)
            Cancel = '|'.join(Cancel)
            Late = '|'.join(Late)
            # st.write(Join, Lead, Cancel, Late)
            Hiking_add.at[0,'참석자'] = Join
            Hiking_add.at[0,'당취자'] = Cancel
            Hiking_add.at[0,'지각자'] = Late
            Hiking_add.at[0,'리딩자'] = Lead
            # st.write(Hiking_add.loc[0].to_list())
        else: trigger_hiking = True


        # 구글시트 저장
        if st.button('**산행 추가하기 ✍️**', disabled=trigger_hiking, width='stretch', type='tertiary'):
            doc = get_connection()
            hiking_ws = doc.worksheet("hiking")
            hiking_ws.append_row(Hiking_add.loc[0].to_list())
            
            doc = get_connection()
            hiking_ws = doc.worksheet("hiking")
            st.session_state.gs['Hiking'] = pd.DataFrame(hiking_ws.get_all_records()).reset_index(drop=True)

            st.session_state.num = NUM+1
            st.rerun()

        # 컴퓨터에 파일에 저장
        # if st.button('**산행 추가하기 ✍️**', disabled=trigger_hiking, width='stretch', type='tertiary'):
        #     Hiking_add = pd.concat([Hiking, Hiking_add], axis=0).reset_index(drop=True)
        #     Hiking_add.to_csv('./savefiles/산행기록.csv', encoding='cp949', index=False)
        #     st.session_state.num = NUM+1
        #     st.rerun()
        #     # st.dataframe(Hiking_add)
        # else:pass



    with mcol22:
        st.markdown("""<div style="font-size: 20px; font-weight: bold; color: black; background-color: white; padding: 1px; border-radius: 8px; text-align: left;margin-bottom: 10px;">
                        ⅱ）일지 확인 📋　［최신순 나열］
                        </div>""",unsafe_allow_html=True)

        st.dataframe(Hiking_edit.sort_index(ascending=False), hide_index=True, width='stretch', height=650)







elif 'Ⅰ．' in MODE:
    mcol11, mcol12 = st.columns([1,3], gap='large')
    with mcol11:
        st.markdown("""<div style="font-size: 20px; font-weight: bold; color: black; background-color: white; padding: 1px; border-radius: 8px; text-align: left;margin-bottom: 10px;">
                ⅰ）인원 관리 👤　［추가/삭제］
                </div>""",unsafe_allow_html=True)
        scol1, scol2 = st.columns([1,4])
        with scol2: Mode_Mem = st.radio('추가삭제', options=['**추가모드**','**삭제모드**','**검색하기**'], label_visibility='collapsed', horizontal=True)


        if '검색' in Mode_Mem:
            scol1, scol2 = st.columns([1,4])
            with scol1: st.write('이름 검색')
            with scol2: st.multiselect('search_name', options=list_Index, label_visibility='collapsed', key='cut'+str(NUM), max_selections=10)
                        

        elif '추가' in Mode_Mem:
            scol1, scol2 = st.columns([1,4])
            with scol1: st.write('이름')
            with scol2: Name = st.text_input('name', placeholder='예시 : 홍길동', label_visibility='collapsed', key='name'+str(NUM))
            scol1, scol2 = st.columns([1,4])
            with scol1: st.write('성별')
            with scol2: Sex = st.selectbox('sex', options=['선택하세요','남','여'], label_visibility='collapsed', key='sex'+str(NUM))
            scol1, scol2 = st.columns([1,4])
            with scol1: st.write('생년')
            with scol2: Birth = st.text_input('birth', placeholder='예시 : YY (2자리)', label_visibility='collapsed', key='birth'+str(NUM))
            scol1, scol2 = st.columns([1,4])
            with scol1: st.write('가입일')
            with scol2: Join = st.text_input('join', placeholder='예시 : YYMMDD (6자리)', label_visibility='collapsed', key='join'+str(NUM))
            scol1, scol2 = st.columns([1,4])
            with scol1: st.write('아이디')
            with scol2: ID = st.text_input('id', placeholder='예시 : 호형호제 (없으면 ; 공란)', label_visibility='collapsed', key='id'+str(NUM))

            if len(Name) > 0 and len(Birth) == 2 and len(Join) == 6 and len(Sex) == 1 and trigger_code == 'GO':
                Membering = pd.DataFrame(index=range(1), columns=list_member)

                Index = '{}({}/{})_{}'.format(Name, Sex, Birth, Join)
                
                Membering.at[0,'이름'] = Name
                Membering.at[0,'성별'] = Sex
                Membering.at[0,'생년'] = int(Birth)
                Membering.at[0,'가입일'] = int(Join)
                Membering.at[0,'ID'] = ID
                Membering.at[0,'Index'] = Index
                
                # st.dataframe(Membering[['이름','성별','생년','가입일','ID','Index']], width='stretch', hide_index=True)
                # st.write(Membering.loc[0].to_list())
                trigger_member = False
            else: trigger_member = True    

            # 구글시트 저장
            if st.button('**멤버 추가하기 ✍️**', disabled=trigger_member, width='stretch', type='tertiary'):
                doc = get_connection()
                member_ws = doc.worksheet("member")
                member_ws.append_row(Membering.loc[0].to_list())

                #새로 읽고 다시 저장
                doc = get_connection()
                member_ws = doc.worksheet("member")
                member = pd.DataFrame(member_ws.get_all_records()).reset_index(drop=True)
                member = member.sort_values(by='Index').reset_index(drop=True)
                st.session_state.gs['Member'] = member

                st.session_state.num = NUM+1
                st.rerun()

            # 컴퓨터에 파일에 저장
            # if st.button('**멤버 추가하기 ✍️**', disabled=trigger_member, width='stretch', type='tertiary'):
            #     Member_add = Member.copy()
            #     Member_add = pd.concat([Member_add, Membering],axis=0).reset_index(drop=True)
            #     Member_add = Member_add.sort_values(by='가입일', ascending=True).reset_index(drop=True)
            #     # st.write(Member_add)

            #     # Member_add.to_csv('./savefiles/멤버관리.csv', encoding='cp949', index=False) # 향후 구글시트
            #     # st.session_state.num = NUM+1
            #     # st.rerun()
            # else:pass
        
        
        
        else: # 삭제모드
            scol1, scol2 = st.columns([1,4])
            with scol1: st.write('이름 검색')
            # with scol2: Cut = st.selectbox('cut', options=['삭제할 1명 선택'] + list_Index, label_visibility='collapsed', key='cut'+str(NUM))
            with scol2: Cut = st.multiselect('cut', options=list_Index, label_visibility='collapsed', key='cut'+str(NUM), max_selections=1)
            
            if len(Cut) != 0 and trigger_code == 'GO': trigger_member_cut = False
            else: trigger_member_cut = True

            # 구글시트 저장
            if st.button('**멤버 삭제하기 ✍️**', disabled=trigger_member_cut, width='stretch', type='tertiary'):
                doc = get_connection()
                member_ws = doc.worksheet("member")
                cell = member_ws.find(Cut[0])
                # st.write(cell)
                row_index = cell.row
                member_ws.delete_rows(row_index)

                #새로 읽고 다시 저장
                doc = get_connection()
                member_ws = doc.worksheet("member")
                member = pd.DataFrame(member_ws.get_all_records()).reset_index(drop=True)
                member = member.sort_values(by='Index').reset_index(drop=True)
                st.session_state.gs['Member'] = member

                st.session_state.num = NUM+1
                st.rerun()
            else:pass

            # 컴퓨터에 파일에 저장
            # if st.button('**멤버 삭제하기 ✍️**', disabled=trigger_member_cut, width='stretch', type='tertiary'):
            #     cut_idx = []
            #     for c_ in Cut: cut_idx.append(Member[Member['Index'] == c_].index[0])
            #     # st.write(cut_idx)
            #     Member_cut = Member.drop(cut_idx, axis=0)
            #     Member_cut = pd.concat([Member_cut, pd.DataFrame(index=range(len(cut_idx)), columns=list_member)], axis=0).reset_index(drop=True)
            #     Member_cut = Member_cut.sort_values(by='가입일', ascending=True).reset_index(drop=True)
            #     # st.write(Member_cut)
                
            #     Member_cut.to_csv('./savefiles/멤버관리.csv', encoding='cp949', index=False) # 향후 구글시트
            #     st.session_state.num = NUM+1
            #     st.rerun()
            # else:pass



    with mcol12:
        st.markdown("""<div style="font-size: 20px; font-weight: bold; color: black; background-color: white; padding: 1px; border-radius: 8px; text-align: left;margin-bottom: 10px;">
                        ⅱ）멤버확인 📋　［현재 멤버수：{} 명］
                        </div>""".format(len(Member_show.index)),unsafe_allow_html=True)
        st.dataframe(Member_show[['이름','성별','생년','가입일','ID']], hide_index=True, width='stretch')
    



    st.write('#####')
    scol31, scol32 = st.columns([8, 1])
    with scol31:
        st.markdown("""<div style="font-size: 20px; font-weight: bold; color: black; background-color: white; padding: 1px; border-radius: 8px; text-align: left;margin-bottom: 10px;">
                                ⅲ）등산 이력 📊　［{}년도］
                                </div>""".format(str(int(Year))),unsafe_allow_html=True)

    if 'Hiking' not in st.session_state:
        list_hiking = ['Index','당취날짜','당취횟수','지각날짜','지각횟수','Q1','Q1_산','Q2','Q2_산','상반기','Q3','Q3_산','Q4','Q4_산','하반기']
        Hiking = pd.DataFrame(index=Member_show.index, columns=list_hiking)
        # 인원별 등산이력 ; Hiking
        Hiking['Index'] = Member['Index']
        
        # 당취자 정리 ; Index로 매칭
        df_cancel = Hiking_edit.dropna(axis=0, subset='당취자').copy()
        # st.write(df_cancel)
        for i in Hiking.index:
            mem_ = Hiking.loc[i,'Index']
            if pd.isna(mem_):pass
            else:
                date_ = []
                for j in df_cancel.index:
                    cancel_ = df_cancel.loc[j,'당취자'] # cancel_ ; 리스트
                    if mem_ in cancel_: date_.append(str(round(df_cancel.loc[j,'날짜'])))
                    else:pass
                date_ = '|'.join(date_)
                Hiking.at[i,'당취날짜'] = date_
        
        for i in Hiking.index:
            mem_ = Hiking.loc[i,'Index']
            if pd.isna(mem_):pass
            else:
                date_ = Hiking.loc[i,'당취날짜']
                if '|' in date_: 
                    date_ = date_.split('|') # 2개이상시 ; 리스트
                    Hiking.at[i,'당취횟수'] = len(date_)
                else:
                    if len(date_) == 0: Hiking.at[i,'당취횟수'] = 0
                    else: Hiking.at[i,'당취횟수'] = 1
                    

        # 지각자 정리
        df_late = Hiking_edit.dropna(axis=0, subset='지각자').copy()
        # st.write(df_late)
        for i in Hiking.index:
            mem_ = Hiking.loc[i,'Index']
            if pd.isna(mem_):pass
            else:
                date_ = []
                for j in df_late.index:
                    cancel_ = df_late.loc[j,'지각자'] # cancel_ ; 리스트
                    if mem_ in cancel_: date_.append(str(round(df_late.loc[j,'날짜'])))
                    else:pass
                date_ = '|'.join(date_)
                Hiking.at[i,'지각날짜'] = date_

        for i in Hiking.index:
            mem_ = Hiking.loc[i,'Index']
            if pd.isna(mem_):pass
            else:
                date_ = Hiking.loc[i,'지각날짜']
                if '|' in date_: 
                    date_ = date_.split('|') # 2개이상시 ; 리스트
                    Hiking.at[i,'지각횟수'] = len(date_)
                else:
                    if len(date_) == 0: Hiking.at[i,'지각횟수'] = 0
                    else: Hiking.at[i,'지각횟수'] = 1



        # 등산이력 정리
        Hiking_edit['날짜'] = Hiking_edit['날짜'].astype(int)
        btw_q1_l = int('{}0100'.format(Year))
        btw_q1_u = int('{}0399'.format(Year))
        btw_q2_l = int('{}0400'.format(Year))
        btw_q2_u = int('{}0699'.format(Year))
        btw_q3_l = int('{}0700'.format(Year))
        btw_q3_u = int('{}0999'.format(Year))
        btw_q4_l = int('{}1000'.format(Year))
        btw_q4_u = int('{}1299'.format(Year))
        # st.write(btw_q1_l, btw_q1_u, btw_q2_l, btw_q2_u, btw_q3_l, btw_q3_u, btw_q4_l, btw_q4_u)
        
        ### Q1
        list_Q = ['Q1','Q2','Q3','Q4']
        for Q in list_Q:
            if Q == 'Q1': Hiking_Q = Hiking_edit[Hiking_edit['날짜'].between(btw_q1_l, btw_q1_u)]
            elif Q == 'Q2': Hiking_Q = Hiking_edit[Hiking_edit['날짜'].between(btw_q2_l, btw_q2_u)]
            elif Q == 'Q3': Hiking_Q = Hiking_edit[Hiking_edit['날짜'].between(btw_q3_l, btw_q3_u)]
            elif Q == 'Q4': Hiking_Q = Hiking_edit[Hiking_edit['날짜'].between(btw_q4_l, btw_q4_u)]
            else:pass
            # st.write(Hiking_Q)
            for i in Hiking.index:
                mem_ = Hiking.loc[i,'Index']
                mountain_ = []
                for j in Hiking_Q.index:
                    join_ = Hiking_Q.loc[j,'참석자'] #리스트
                    cancel_ = Hiking_Q.loc[j,'당취자'] #float or 리스트
                    if type(cancel_) == float: cancel_ = []
                    else:pass
                    # st.write(join_, cancel_)
                    if mem_ in join_: # 참석명단에 있으면?
                        if mem_ in cancel_: pass # 당취면 pass
                        else: mountain_.append(Hiking_Q.loc[j,'산']) # 당취 아니면 산 기록!
                    else:pass # 참석명단에 없으면 pass
                Hiking.at[i,'{}'.format(Q)] = len(mountain_)
                Hiking.at[i,'{}_산'.format(Q)] = mountain_

        for i in Hiking.index:
            Hiking.at[i,'상반기'] = Hiking.loc[i,'Q1'] + Hiking.loc[i,'Q2']
            Hiking.at[i,'하반기'] = Hiking.loc[i,'Q3'] + Hiking.loc[i,'Q4']


        # 컷 대상자
        if Month >= 1 and Month <= 3: Season = 'Q1'
        elif Month >= 4 and Month <= 6: Season = 'Q2'
        elif Month >= 7 and Month <= 9: Season = 'Q3'
        elif Month >= 10 and Month <= 12: Season = 'Q4'
        else:pass
        Colname = '{}_Cut대상자'.format(Season)
        Hiking[[Colname]] = None

        for i in Hiking.index:
            mem_ = Hiking.loc[i,'Index']
            if pd.isna(mem_):pass
            else:
                passfail = Hiking.loc[i,Season]
                if passfail > 0: Hiking.at[i, Colname] = 'Pass'
                else: Hiking.at[i, Colname] = '해당！'
        st.session_state.Hiking = Hiking
    else:pass    
    Hiking = st.session_state.Hiking
    Hiking = Hiking.sort_values(by='Index',ascending=True)
    st.dataframe(Hiking, width='stretch', hide_index=True, height=660)

    with scol32:
        st.download_button(
            label='엑셀파일 다운로드 📥',
            data=to_excel(st.session_state.Hiking),
            file_name='등산이력_{}년도.xlsx'.format(Year),
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            width='stretch',
            type='primary',

        )






else:pass



