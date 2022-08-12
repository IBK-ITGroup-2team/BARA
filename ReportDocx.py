from telnetlib import DO
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime
from docx.shared import Cm, Inches
from docx.text.run import Font
from docx.oxml.ns import qn

document=Document()
# 스타일 적용
style=document.styles['Normal']
font=style.font
font.name='Arial'

banks=['하나','우리','신한','국민','농협']
internetBanks=['카카오뱅크','케이뱅크','토스']

document.add_heading('\"고객과 함께, 신뢰와 책임, 열정과 혁신, 소통과 팀웍\"', level = 1) 
dateToday=datetime.today()
document.add_paragraph(datetime.today().strftime("%Y. %m. %d"))  #해당 날짜

send=document.add_paragraph('')
send.add_run('수 신     ').bold=True
send.add_run('모바일 앱 개발 이해 관련 부서').bold=True
document.add_heading('제 목  『IBK 모바일 앱 사용자 반응 비교』', level=2)
document.add_paragraph('')

objective=document.add_paragraph('□ 발간 목적')
objective.bold=True
#objective.add_run('당행의 모바일 앱에 대한 사용자들의 반응을 이해관계자에 효과적으로 전달하고, 타행과의 비교를 통해 개선점을 찾고자 함').bold=True
document.add_paragraph('당행의 모바일 앱에 대한 사용자들의 반응을 이해관계자에 효과적으로 전달하고, 타행과의 비교를 통해 개선점을 찾고자 함')

document.add_paragraph('')

document.add_paragraph('□ 주요 내용 목차')
document.add_paragraph(' 1. 당행 개인고객용 모바일 앱 (i-one bank) 사용자 반응 분석')
document.add_paragraph(' 2. 타행 개인고객용 모바일 앱 사용자 반응 비교 분석')
document.add_paragraph(' 3. 당행 기업고객용 모바일 앱 (i-one bank) 사용자 반응 분석')
document.add_paragraph(' 4. 타행 기업고객용 모바일 앱 사용자 반응 비교 분석')
document.add_paragraph(' 5. 인터넷 전문 은행 모바일 앱 사용자 반응 비교 분석')
document.add_paragraph('')

for i in range(9):
    document.add_paragraph('')  # 다음 장으로 이동

document.add_paragraph('□ 주요 내용')

# 1번 당행 개인 앱 리뷰 현황
document.add_paragraph(' 1. 당행 개인고객용 모바일 앱 (i-one bank) 사용자 반응 분석')
document.add_paragraph('   ㅇ 워드클라우드로 나타낸 당행 모바일 앱 사용자 반응')
document.add_paragraph('     ㄱ. 긍정적 반응')
document.add_picture('WordCloudEx.PNG', width=Cm(16), height=Cm(8))  #추후에 실제 워드클라우드 이미지로 변경할 것.
document.add_paragraph('     ㄴ. 부정적 반응')
document.add_picture('WordCloudEx.PNG', width=Cm(16), height=Cm(8))  #추후에 실제 워드클라우드 이미지로 변경할 것.
document.add_paragraph('')

#빈출 단어 도출
document.add_paragraph('   ㅇ 빈출 단어 Top3')

# 추후 워드 클라우드 코드와 연결
'''
for i in range(3):
    document.add_paragraph(i'. 'array[i])
'''
document.add_paragraph('')

# 2번 타행 개인 앱 리뷰 현황
document.add_paragraph(' 2. 타행 개인고객용 모바일 앱 사용자 반응 분석')
for i in banks:
    document.add_paragraph(i+"은행")
    document.add_paragraph('   ㅇ 워드클라우드로 나타낸 '+i+'은행 모바일 앱 사용자 반응')
    document.add_paragraph('     ㄱ. 긍정적 반응')
    document.add_picture(i+'WordCloudP.PNG', width=Cm(16), height=Cm(8))  #추후에 실제 워드클라우드 이미지로 변경할 것.
    document.add_paragraph('     ㄴ. 부정적 반응')
    document.add_picture(i+'WordCloudN.PNG',width=Cm(16), height=Cm(8))  #추후 변경
    document.add_paragraph('   ㅇ 빈출 단어 Top3')    
    # 추후 워드 클라우드 코드와 연결
    '''
    for i in range(3):
        document.add_paragraph(i'. 'array[i])
    '''
    
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
resultPersonal=document.add_paragraph('')    
resultPersonal.add_run('결 론           ').bold=True
resultPersonal.add_run(bestWord2+'에 힘쓰는 것이 좋겠다고 판단됨.')
    
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


#마지막 꼬릿말
document.add_paragraph('\"새로운 60년, 고객을 향한 혁신\"')

# 문단별 정렬
paragraph1=document.paragraphs[0]   #첫번째 문단
paragraph1.alignment=WD_ALIGN_PARAGRAPH.CENTER

paragraph2=document.paragraphs[1]
paragraph2.alignment=WD_ALIGN_PARAGRAPH.RIGHT


paragraphLast=document.paragraphs[-1]
#paragraphLast.font.size=Document.shared.Pt(20)
paragraphLast.alignment=WD_ALIGN_PARAGRAPH.CENTER   #마지막 문단



#파일 저장 // 마지막 단계
document.save("report.docx")

