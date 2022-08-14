from msilib.schema import MIME
from telnetlib import DO
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime
from docx.shared import Cm, Inches
from docx.text.run import Font
from docx.oxml.ns import qn 

import smtplib
import os
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication  #메일의 첨부 파일을 base64 형식으로 변환
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email import encoders

document=Document()
# 스타일 적용
style=document.styles['Normal']
font=style.font
font.name='Arial'

banks=['하나','우리','신한','국민','농협']
internetBanks=['카카오뱅크','케이뱅크','토스']

document.add_heading('\"고객과 함께, 신뢰와 책임, 열정과 혁신, 소통과 팀웍\"', level = 1) 
document.add_paragraph('')
dateToday=datetime.today()
document.add_paragraph(datetime.today().strftime("%Y. %m. %d"))  #해당 날짜

send=document.add_paragraph('')
send.add_run('수 신             ').bold=True
send.add_run('모바일 앱 개발 이해 관련 부서').bold=True

title=document.add_paragraph('')
title.add_run('제 목           『'+datetime.today().strftime("%Y년 %m월")+'IBK 모바일 앱 사용자 반응 비교』').bold=True

document.add_paragraph('')
document.add_paragraph('')

objective=document.add_paragraph('')
objective.add_run('□ 발간 목적').bold=True
#objective.add_run('당행의 모바일 앱에 대한 사용자들의 반응을 이해관계자에 효과적으로 전달하고, 타행과의 비교를 통해 개선점을 찾고자 함').bold=True
document.add_paragraph('당행의 모바일 앱에 대한 사용자들의 반응을 이해관계자에 효과적으로 전달하고, 타행과의 비교를 통해 개선점을 찾고자 함')

document.add_paragraph('')

index=document.add_paragraph('')
index.add_run('□ 주요 내용 목차').bold=True

document.add_paragraph(' 1. 당행 개인고객용 모바일 앱 (i-one bank) 사용자 반응 분석')
document.add_paragraph(' 2. 타행 개인고객용 모바일 앱 사용자 반응 비교 분석')
document.add_paragraph(' 3. 당행 기업고객용 모바일 앱 (i-one bank) 사용자 반응 분석')
document.add_paragraph(' 4. 타행 기업고객용 모바일 앱 사용자 반응 비교 분석')
document.add_paragraph(' 5. 인터넷 전문 은행 모바일 앱 사용자 반응 비교 분석')
document.add_paragraph('')

for i in range(7):
    document.add_paragraph('')  # 다음 장으로 이동

main=document.add_paragraph('')
main.add_run('□ 주요 내용').bold=True

# 1번 당행 개인 앱 리뷰 현황
document.add_paragraph(' 1. 당행 개인고객용 모바일 앱 (i-one bank) 사용자 반응 분석')

#표 생성
grid_t_style = document.styles["Table Grid"]
IBKTable = document.add_table(3,2,grid_t_style)

IBKTableCells1=IBKTable.rows[0].cells
IBKTableCells1[0].paragraphs[0].add_run('긍정적 반응').bold=True
IBKTableCells1[0].paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER

IBKTableCells1=IBKTable.rows[0].cells
IBKTableCells1[1].paragraphs[0].add_run('부정적 반응').bold=True
IBKTableCells1[1].paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER

# 표에 워드클라우드 삽입
IBKCell10=IBKTable.cell(1,0)
IBKPara10=IBKCell10.add_paragraph()
IBKRun10=IBKPara10.add_run()
IBKRun10.add_picture("WordCloudEx.PNG",width=Cm(7),height=Cm(5))

IBKCell11=IBKTable.cell(1,1)
IBKPara11=IBKCell11.add_paragraph()
IBKRun11=IBKPara11.add_run()
IBKRun11.add_picture("WordCloudEx.PNG",width=Cm(7),height=Cm(5))

# 긍정 빈출 단어 Top3
IBKTableCells3=IBKTable.rows[2].cells
IBKTableCells3[0].paragraphs[0].add_run('빈출 단어 Top3\n')
IBKTableCells3[0].paragraphs[0].add_run('1\n')
IBKTableCells3[0].paragraphs[0].add_run('2\n')
IBKTableCells3[0].paragraphs[0].add_run('3\n')

# 부정 빈출 단어 Top3
IBKTableCells3[1].paragraphs[0].add_run('빈출 단어 Top3\n')
IBKTableCells3[1].paragraphs[0].add_run('1\n')
IBKTableCells3[1].paragraphs[0].add_run('2\n')
IBKTableCells3[1].paragraphs[0].add_run('3\n')

document.add_paragraph('')

#  하나은행 개인 앱 리뷰 현황
document.add_paragraph(' 2. 타행 개인고객용 모바일 앱 사용자 반응 분석')
document.add_paragraph('    ㅇ 하나은행')

# 하나은행 표 생성
HANA1Table = document.add_table(3,2,grid_t_style)

HANA1Cells1=HANA1Table.rows[0].cells
HANA1Cells1[0].paragraphs[0].add_run('긍정적 반응').bold=True
HANA1Cells1[0].paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER

HANA1TableCells1=HANA1Table.rows[0].cells
HANA1Cells1[1].paragraphs[0].add_run('부정적 반응').bold=True
HANA1TableCells1[1].paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER

# 표에 워드클라우드 삽입
HANACell10=HANA1Table.cell(1,0)
HANAPara10=HANACell10.add_paragraph()
HANARun10=HANAPara10.add_run()
HANARun10.add_picture("WordCloudEx.PNG",width=Cm(7),height=Cm(5))

HANACell11=HANA1Table.cell(1,1)
HANAPara11=HANACell11.add_paragraph()
HANARun11=HANAPara11.add_run()
HANARun11.add_picture("WordCloudEx.PNG",width=Cm(7),height=Cm(5))

# 긍정 빈출 단어 Top3
HANATableCells3=HANA1Table.rows[2].cells
HANATableCells3[0].paragraphs[0].add_run('빈출 단어 Top3\n')
HANATableCells3[0].paragraphs[0].add_run('1\n')
HANATableCells3[0].paragraphs[0].add_run('2\n')
HANATableCells3[0].paragraphs[0].add_run('3\n')

# 부정 빈출 단어 Top3
HANATableCells3[1].paragraphs[0].add_run('빈출 단어 Top3\n')
HANATableCells3[1].paragraphs[0].add_run('1\n')
HANATableCells3[1].paragraphs[0].add_run('2\n')
HANATableCells3[1].paragraphs[0].add_run('3\n')
    
document.add_paragraph('')
document.add_paragraph('')

#국민은행 앱 리뷰 현황
document.add_paragraph('    ㅇ 국민은행')

# 국민은행 표 생성
KB1Table = document.add_table(3,2,grid_t_style)

KBCells1=KB1Table.rows[0].cells
KBCells1[0].paragraphs[0].add_run('긍정적 반응').bold=True
KBCells1[0].paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER

KB1TableCells1=KB1Table.rows[0].cells
KBCells1[1].paragraphs[0].add_run('부정적 반응').bold=True
KB1TableCells1[1].paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER

# 표에 워드클라우드 삽입
KBCell10=KB1Table.cell(1,0)
KBPara10=KBCell10.add_paragraph()
KBRun10=KBPara10.add_run()
KBRun10.add_picture("WordCloudEx.PNG",width=Cm(7),height=Cm(5))

KBCell11=KB1Table.cell(1,1)
KBPara11=KBCell11.add_paragraph()
KBRun11=KBPara11.add_run()
KBRun11.add_picture("WordCloudEx.PNG",width=Cm(7),height=Cm(5))

# 긍정 빈출 단어 Top3
KBTableCells3=KB1Table.rows[2].cells
KBTableCells3[0].paragraphs[0].add_run('빈출 단어 Top3\n')
KBTableCells3[0].paragraphs[0].add_run('1\n')
KBTableCells3[0].paragraphs[0].add_run('2\n')
KBTableCells3[0].paragraphs[0].add_run('3\n')

# 부정 빈출 단어 Top3
KBTableCells3[1].paragraphs[0].add_run('빈출 단어 Top3\n')
KBTableCells3[1].paragraphs[0].add_run('1\n')
KBTableCells3[1].paragraphs[0].add_run('2\n')
KBTableCells3[1].paragraphs[0].add_run('3\n')

    
# 임시 변수 !!!!!!!!!!!!!!!!!!!!!! 나중에 지우기
bestWord='UI 개선'    
resultPersonal=document.add_paragraph('')    
resultPersonal.add_run('결 론           ').bold=True
resultPersonal.add_run(bestWord+'에 힘쓰는 것이 좋겠다고 판단됨.')
    
# 당행 기업 앱 리뷰 현황
document.add_paragraph(' 3. 당행 기업고객용 모바일 앱 (i-one bank) 사용자 반응 분석')
document.add_paragraph('   ㅇ 워드클라우드로 나타낸 당행 모바일 앱 사용자 반응')
document.add_paragraph('     ㄱ. 긍정적 반응')
document.add_picture('WordCloudEx.PNG', width=Cm(16), height=Cm(8))  #추후에 실제 워드클라우드 이미지로 변경할 것.
document.add_paragraph('     ㄴ. 부정적 반응')
document.add_picture('WordCloudEx.PNG', width=Cm(16), height=Cm(8))  #추후에 실제 워드클라우드 이미지로 변경할 것.
document.add_paragraph('')

# 타행 기업 앱 리뷰 현황
document.add_paragraph(' 4. 타행 기업고객용 모바일 앱 사용자 반응 분석')
for i in banks:
    document.add_paragraph(i+"은행")
    document.add_paragraph('   ㅇ 워드클라우드로 나타낸 '+i+'은행 모바일 앱 사용자 반응')
    document.add_paragraph('     ㄱ. 긍정적 반응')
    document.add_picture(i+'EWordCloudP.PNG', width=Cm(16), height=Cm(8))  #추후에 실제 워드클라우드 이미지로 변경할 것.
    document.add_paragraph('     ㄴ. 부정적 반응')
    document.add_picture(i+'EWordCloudN.PNG',width=Cm(16), height=Cm(8))  #추후 변경
    document.add_paragraph('   ㅇ 빈출 단어 Top3') 
    
# 임시 변수 !!!!!!!!!!!!!!!!!!!!!! 나중에 지우기
bestWord2='기업 이미지 개선'    
resultEnterprise=document.add_paragraph('')    
resultEnterprise.add_run('결 론           ').bold=True
resultEnterprise.add_run(bestWord2+'에 힘쓰는 것이 좋겠다고 판단됨.')
    
# 인터넷 전문 은행 앱 리뷰 현황
document.add_paragraph(' 5. 인터넷 전문 은행 모바일 앱 사용자 반응 분석')
for i in internetBanks:
    document.add_paragraph(i)
    document.add_paragraph('   ㅇ 워드클라우드로 나타낸 '+i+' 모바일 앱 사용자 반응')
    document.add_paragraph('     ㄱ. 긍정적 반응')
    document.add_picture(i+'WordCloudP.PNG', width=Cm(16), height=Cm(8))  #추후에 실제 워드클라우드 이미지로 변경할 것.
    document.add_paragraph('     ㄴ. 부정적 반응')
    document.add_picture(i+'WordCloudN.PNG',width=Cm(16), height=Cm(8))  #추후 변경
    document.add_paragraph('   ㅇ 빈출 단어 Top3') 

# 임시 변수!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
bestWordInternet='부가 서비스'
resultInternet=document.add_paragraph('')
resultInternet.add_run('결 론           ').bold=True
resultInternet.add_run(bestWordInternet+'에 힘쓰는 것이 좋겠다고 판단됨.')

#마지막 꼬릿말
document.add_paragraph('\"새로운 60년, 고객을 향한 혁신\"')

# 문단별 정렬
paragraph1=document.paragraphs[0]   #첫번째 문단
paragraph1.alignment=WD_ALIGN_PARAGRAPH.CENTER

paragraph2=document.paragraphs[2]
paragraph2.alignment=WD_ALIGN_PARAGRAPH.RIGHT


paragraphLast=document.paragraphs[-1]
#paragraphLast.font.size=Document.shared.Pt(20)
paragraphLast.alignment=WD_ALIGN_PARAGRAPH.CENTER   #마지막 문단

#파일 저장 // 마지막 단계
document.save("report.docx")


'''

# 보고서를 이메일로 발송
s=smtplib.SMTP('smtp.gmail.com',587)    #gmail 포트번호 587
s.starttls()    # TLS(Transport Layer Security) 보안

s.login('IBK.ITgroup.2@gmail.com','czhoerpcnkfzqsdh')  # 메일을 보내는 계정

#메일 정보
msg=MIMEMultipart()
msg['From']='IBK.ITgroup.2@gmail.com'
msg['To']='bethh05108@gmail.com'
msg['Subject']=datetime.today().strftime("%Y. %m. %d")+ " I-one bank 사용자 반응 보고서입니다."

#메일 내용
content=datetime.today().strftime("%Y. %m. %d")+ "I-one bank 사용자 반응 보고서입니다."
part2=MIMEText(content,'plain')
msg.attach(part2)

# 보고서 첨부
part = MIMEBase('application','octet-stream')
with open("report.docx",'rb') as file:
    part.set_payload(file.read())
encoders.encode_base64(part)
part.add_header('Content-Disposition','attachment',filename='report.docx')
msg.attach(part)

#메일 전송
s.sendmail("IBK.ITgroup.2@gmail.com","bethh05108@gmail.com",msg.as_string())

#세션 종료
s.quit()
'''
