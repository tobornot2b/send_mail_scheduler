import smtplib
from email.mime.multipart import MIMEMultipart # 메일의 Data 영역의 메시지를 만드는 모듈 (MIMEText, MIMEApplication, MIMEImage, MIMEAudio가 attach되면 바운더리 형식으로 변환)
from email.mime.text import MIMEText # MIMEText('메일 내용', '메일 내용의 MIME 타입')
# from email.mime.application import MIMEApplication # MIMEApplication('첨부할 파일', '첨부할 파일의 MIME 타입')
# from email.mime.audio import MIMEAudio # MIMEAudio('음악 파일', '음악 파일의 MIME 타입')
from email.mime.image import MIMEImage # MIMEImage('이미지 파일', '이미지 파일의 MIME 타입')
from email.utils import formatdate # 메일의 헤더에 날짜를 넣어주는 모듈
from email.header import Header

import re
import pandas as pd
import numpy as np
from datetime import datetime, date
from apscheduler.schedulers.background import BackgroundScheduler # 백그라운드 스케줄러
import time
import sys


# 키 가져오기
sys.path.append('c:/settings')
import config
mail_acc = config.send_mail_info['mail_account']
mail_pass = config.send_mail_info['mail_pw']


# 이메일 형식 검사 함수
def check_email(email: str) -> bool:
    regex = '^[a-zA-Z0-9+-_.]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$' # 이메일 정규식
    if re.match(regex, email): # re.match('정규식', '문자열')
        return True
    else:
        return False


# 목록을 읽어서 전처리하는 함수
def read_list_and_data_preprocessing() -> pd.DataFrame:
    df = pd.read_csv('./birth_list.csv', encoding='utf-8', dtype=str) # dtype을 str로 일괄지정해주는게 편하다.

    df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x) # 데이터프레임 전체 공백제거
    df['생일'] =  df['생일'].str.replace('/', '-') # 생일 데이터에서 '/'를 '-'로 변경
    df['입사일'] =  df['입사일'].str.replace('/', '-') # 입사일 데이터에서 '/'를 '-'로 변경

     # 이메일 형식 검사
    df['이메일검사결과'] = df['이메일'].apply(check_email)
    if df['이메일검사결과'].sum() != len(df): # 이메일 형식이 잘못된 데이터가 있으면
        print('이메일 형식이 잘못된 데이터가 있습니다.')
        print(df[df['이메일검사결과'] == False])
        return None
    
    today = datetime.now().date() # 오늘 날짜
    df['발송일'] = str(today.year) + df['생일'].str[4:] # '년도'를 현재 년도로 변경
    df['생일'] = pd.to_datetime(df['생일'], format='%Y-%m-%d') # datetime 형식으로 변경
    df['입사일'] = pd.to_datetime(df['입사일'], format='%Y-%m-%d')
    df['발송일'] = pd.to_datetime(df['발송일'], format='%Y-%m-%d')
    
    df = df[df['발송일'] >= str(today)] # 발송일이 오늘보다 큰 데이터만 추출 today가 str임에 주의
    df = df.sort_values('발송일').reset_index(drop=True) # 발송일 기준으로 오름차순 정렬
    df['그림'] = pd.Series(np.random.randint(1, 10, size=len(df))).astype(str) # 그림선택 컬럼에 1~3 사이의 랜덤 숫자를 넣음
    
    df['근속일'] = (df['발송일'] - df['입사일']).dt.days.astype(str) # 근속일 = 발송일 - 입사일
    df['나이'] = ((df['발송일'] - df['생일']).dt.days // 365).astype(str) # 나이 = 발송일 - 생일

    title_list = []
    content_list = []
    for i in range(len(df)):
        title = f"[아이비클럽] {df['부서명'][i]} {df['사원명'][i]} {df['직급'][i]}님의 {df['나이'][i]}번째 생일을 축하합니다!"
        content = f'''
        <html>
        <head></head>
        <body>
        안녕하세요, {df['사원명'][i]} {df['직급'][i]}님.<br>
        <br>
        <br>
        {df['사원명'][i]}님의 생일을 진심으로 축하드립니다.<br>
        <br>
        뜻깊은 오늘, 그 어느 때보다도 아름다운 날이 되기를 바랍니다.<br>
        <br>
        {df['사원명'][i]} {df['직급'][i]}님께서 그간 우리 아이비클럽을 위해 {df['근속일'][i]}일간 근무하신 노고에 깊은 감사를 전하며<br>
        <br>
        앞으로도 더욱더 발전하는 {df['사원명'][i]} {df['직급'][i]}님이 되기를 기원합니다.<br>
        <br>
        오늘 하루 생애 최고로 행복하시길 바랍니다.<br>
        <br>
        감사합니다.<br>
        <br>
        '''
        title_list.append(title)
        content_list.append(content)
    df['제목'] = title_list
    df['본문'] = content_list

    print(df)
    return df


# 메일 전송 함수
def send_mail(from_addr: str, to_addr: str, to_name: str, subject: str, content: str, image_number: str):
    msg = MIMEMultipart() # MIMEMultipart()를 사용하면 여러개의 MIME 타입을 사용할 수 있음
    
    msg['From'] = '아이비클럽' # 보내는 사람 명칭
    msg['To'] = to_name # 받는 사람 명칭
    msg['Date'] = formatdate(localtime=True)
    # msg['Cc'] = cc_addr # 참조 메일 주소
    # msg['Bcc'] = bcc_addr # 숨은 참조 메일 주소
    msg['Subject'] = Header(s=subject, charset='utf-8') # 메일 제목
    
    with open(f'./image/{image_number}.jpg', 'rb') as f: # 이미지 파일 추가
        img = MIMEImage(f.read(), Name = 'happybirthday.jpg') # Name은 메일 수신자에서 설정되는 파일 이름
        img.add_header('Content-ID', '<capture>') # 해더에 Content-ID 추가(본문 내용에서 cid로 링크를 걸 수 있다.)
        msg.attach(img) # Data 영역의 메시지에 바운더리 추가
    
    body = MIMEText(f"{content}</br><img src='cid:capture'>", 'html')
    msg.attach(body)
    print(msg)

    smtp = smtplib.SMTP('smtp.gmail.com', 587) # smtplib.SMTP('사용할 SMTP 서버의 URL', PORT)
    smtp.ehlo() # SMTP 서버 연결
    smtp.starttls() # TLS 암호화 (TLS 사용할 때에만 해당코드 입력)
    smtp.login(from_addr, mail_pass) # smtp.login('메일 주소', '비밀번호')

    smtp.sendmail(from_addr, to_addr, msg.as_string()) # smtp.sendmail('보내는 사람 메일 주소', '받는 사람 메일 주소', '메시지')
    time.sleep(1)
    smtp.quit() # SMTP 서버 연결 종료
    print('메일 전송이 완료되었습니다.')


# 메일 전송 스케줄러
def background_scheduler():
    scheduler = BackgroundScheduler(daemon=True, timezone='Asia/Seoul')
    scheduler.start()
    df = read_list_and_data_preprocessing()
    for i in df.index:
        scheduler.add_job(
            send_mail,
            'date',
            run_date=datetime(df.loc[i, '발송일'].year, df.loc[i, '발송일'].month, df.loc[i, '발송일'].day, 16, 52, i),
            args=[mail_acc, df.loc[i, '이메일'], df.loc[i, '사원명'], df.loc[i, '제목'], df.loc[i, '본문'], df.loc[i, '그림']],
            id=df.loc[i, '사원코드'],
            )
    scheduler.print_jobs()


if __name__ == '__main__':
    background_scheduler()
    while True:
        time.sleep(1)