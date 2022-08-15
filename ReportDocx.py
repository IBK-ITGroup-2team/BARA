from msilib.schema import MIME
from telnetlib import DO
from docx import Document
from docx.shared import Pt
# from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from datetime import datetime
from docx.shared import Cm, Inches
from docx.text.run import Font
from docx.oxml.ns import qn

import smtplib
import os
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication  # 메일의 첨부 파일을 base64 형식으로 변환
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email import encoders


def main():
    document = Document()
    # 스타일 적용
    style = document.styles['Normal']
    font = style.font
    font.name = 'Arial'

    banks = ['하나', '우리', '신한', '국민', '농협']
    internetBanks = ['카카오뱅크', '케이뱅크', '토스']

    document.add_heading('\"고객과 함께, 신뢰와 책임, 열정과 혁신, 소통과 팀웍\"', level=1)
    document.add_paragraph('')
    dateToday = datetime.today()
    document.add_paragraph(datetime.today().strftime("%Y. %m. %d"))  # 해당 날짜

    send = document.add_paragraph('')
    send.add_run('수 신             ').bold = True
    send.add_run('모바일 앱 개발 이해 관련 부서').bold = True

    title = document.add_paragraph('')
    title.add_run('제 목           『' + datetime.today().strftime("%Y년 %m월") + 'IBK 모바일 앱 사용자 반응 비교』').bold = True

    document.add_paragraph('')
    document.add_paragraph('')

    objective = document.add_paragraph('')
    objective.add_run('□ 발간 목적').bold = True
    # objective.add_run('당행의 모바일 앱에 대한 사용자들의 반응을 이해관계자에 효과적으로 전달하고, 타행과의 비교를 통해 개선점을 찾고자 함').bold=True
    document.add_paragraph('당행의 모바일 앱에 대한 사용자들의 반응을 이해관계자에 효과적으로 전달하고, 타행과의 비교를 통해 개선점을 찾고자 함')

    document.add_paragraph('')

    index = document.add_paragraph('')
    index.add_run('□ 주요 내용 목차').bold = True

    document.add_paragraph(' 1. 당행 개인고객용 모바일 앱 (i-one bank) 사용자 반응 분석')
    document.add_paragraph(' 2. 타행 개인고객용 모바일 앱 사용자 반응 비교 분석')
    document.add_paragraph(' 3. 당행 기업고객용 모바일 앱 (i-one bank) 사용자 반응 분석')
    document.add_paragraph(' 4. 타행 기업고객용 모바일 앱 사용자 반응 비교 분석')
    document.add_paragraph(' 5. 인터넷 전문 은행 모바일 앱 사용자 반응 비교 분석')
    document.add_paragraph('')

    for i in range(7):
        document.add_paragraph('')  # 다음 장으로 이동

    main = document.add_paragraph('')
    main.add_run('□ 주요 내용').bold = True

    # 1번 당행 개인 앱 리뷰 현황
    document.add_paragraph(' 1. 당행 개인고객용 모바일 앱 (i-one bank) 사용자 반응 분석')

    # 표 생성
    grid_t_style = document.styles["Table Grid"]
    IBKTable = document.add_table(3, 2, grid_t_style)

    IBKTableCells1 = IBKTable.rows[0].cells
    IBKTableCells1[0].paragraphs[0].add_run('긍정적 반응').bold = True
    IBKTableCells1[0].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    # IBKTableCells1[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    IBKTableCells1 = IBKTable.rows[0].cells
    IBKTableCells1[1].paragraphs[0].add_run('부정적 반응').bold = True
    IBKTableCells1[1].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    # IBKTableCells1[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 표에 워드클라우드 삽입
    IBKCell10 = IBKTable.cell(1, 0)
    IBKPara10 = IBKCell10.add_paragraph()
    IBKRun10 = IBKPara10.add_run()
    IBKRun10.add_picture("./wordcloud/개인고객/IBK_WordCloud_P.png", width=Cm(7), height=Cm(5))

    IBKCell11 = IBKTable.cell(1, 1)
    IBKPara11 = IBKCell11.add_paragraph()
    IBKRun11 = IBKPara11.add_run()
    IBKRun11.add_picture("./wordcloud/개인고객/IBK_WordCloud_N.png", width=Cm(7), height=Cm(5))

    # 긍정 빈출 단어 Top3
    IBKTableCells3 = IBKTable.rows[2].cells
    IBKTableCells3[0].paragraphs[0].add_run('빈출 단어 Top3\n')
    # IBKTableCells3[0].paragraphs[0].add_run('1. ' + IBK_positive_top3[0] + '\n')
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
    HANA1Table = document.add_table(3, 2, grid_t_style)

    HANA1Cells1 = HANA1Table.rows[0].cells
    HANA1Cells1[0].paragraphs[0].add_run('긍정적 반응').bold = True
    HANA1Cells1[0].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    # HANA1Cells1[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    HANA1TableCells1 = HANA1Table.rows[0].cells
    HANA1Cells1[1].paragraphs[0].add_run('부정적 반응').bold = True
    HANA1TableCells1[1].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    # HANA1TableCells1[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 표에 워드클라우드 삽입
    HANACell10 = HANA1Table.cell(1, 0)
    HANAPara10 = HANACell10.add_paragraph()
    HANARun10 = HANAPara10.add_run()
    HANARun10.add_picture("./wordcloud/개인고객/HANA_WordCloud_P.png", width=Cm(7), height=Cm(5))

    HANACell11 = HANA1Table.cell(1, 1)
    HANAPara11 = HANACell11.add_paragraph()
    HANARun11 = HANAPara11.add_run()
    HANARun11.add_picture("./wordcloud/개인고객/HANA_WordCloud_N.png", width=Cm(7), height=Cm(5))

    # 긍정 빈출 단어 Top3
    HANATableCells3 = HANA1Table.rows[2].cells
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

    # 국민은행 앱 리뷰 현황
    document.add_paragraph('    ㅇ 국민은행')

    # 국민은행 표 생성
    KB1Table = document.add_table(3, 2, grid_t_style)

    KBCells1 = KB1Table.rows[0].cells
    KBCells1[0].paragraphs[0].add_run('긍정적 반응').bold = True
    KBCells1[0].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    # KBCells1[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    KB1TableCells1 = KB1Table.rows[0].cells
    KBCells1[1].paragraphs[0].add_run('부정적 반응').bold = True
    KB1TableCells1[1].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    # KB1TableCells1[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 표에 워드클라우드 삽입
    KBCell10 = KB1Table.cell(1, 0)
    KBPara10 = KBCell10.add_paragraph()
    KBRun10 = KBPara10.add_run()
    KBRun10.add_picture("./wordcloud/개인고객/KB_WordCloud_P.png", width=Cm(7), height=Cm(5))

    KBCell11 = KB1Table.cell(1, 1)
    KBPara11 = KBCell11.add_paragraph()
    KBRun11 = KBPara11.add_run()
    KBRun11.add_picture("./wordcloud/개인고객/KB_WordCloud_N.png", width=Cm(7), height=Cm(5))

    # 긍정 빈출 단어 Top3
    KBTableCells3 = KB1Table.rows[2].cells
    KBTableCells3[0].paragraphs[0].add_run('빈출 단어 Top3\n')
    KBTableCells3[0].paragraphs[0].add_run('1\n')
    KBTableCells3[0].paragraphs[0].add_run('2\n')
    KBTableCells3[0].paragraphs[0].add_run('3\n')

    # 부정 빈출 단어 Top3
    KBTableCells3[1].paragraphs[0].add_run('빈출 단어 Top3\n')
    KBTableCells3[1].paragraphs[0].add_run('1\n')
    KBTableCells3[1].paragraphs[0].add_run('2\n')
    KBTableCells3[1].paragraphs[0].add_run('3\n')

    document.add_paragraph('')
    document.add_paragraph('')

    # 신한은행 앱 리뷰 현황
    document.add_paragraph('    ㅇ 신한은행')

    # 신한은행 표 생성
    SH1Table = document.add_table(3, 2, grid_t_style)

    SHCells1 = SH1Table.rows[0].cells
    SHCells1[0].paragraphs[0].add_run('긍정적 반응').bold = True
    SHCells1[0].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    # SHCells1[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    SH1TableCells1 = SH1Table.rows[0].cells
    SHCells1[1].paragraphs[0].add_run('부정적 반응').bold = True
    SH1TableCells1[1].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    # SH1TableCells1[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 표에 워드클라우드 삽입-신한
    SHCell10 = SH1Table.cell(1, 0)
    SHPara10 = SHCell10.add_paragraph()
    SHRun10 = SHPara10.add_run()
    SHRun10.add_picture("./wordcloud/개인고객/SHINHAN_WordCloud_P.png", width=Cm(7), height=Cm(5))

    SHCell11 = SH1Table.cell(1, 1)
    SHPara11 = SHCell11.add_paragraph()
    SHRun11 = SHPara11.add_run()
    SHRun11.add_picture("./wordcloud/개인고객/SHINHAN_WordCloud_N.png", width=Cm(7), height=Cm(5))

    # 긍정 빈출 단어 Top3
    SHTableCells3 = SH1Table.rows[2].cells
    SHTableCells3[0].paragraphs[0].add_run('빈출 단어 Top3\n')
    SHTableCells3[0].paragraphs[0].add_run('1\n')
    SHTableCells3[0].paragraphs[0].add_run('2\n')
    SHTableCells3[0].paragraphs[0].add_run('3\n')

    # 부정 빈출 단어 Top3
    SHTableCells3[1].paragraphs[0].add_run('빈출 단어 Top3\n')
    SHTableCells3[1].paragraphs[0].add_run('1\n')
    SHTableCells3[1].paragraphs[0].add_run('2\n')
    SHTableCells3[1].paragraphs[0].add_run('3\n')

    for i in range(3):
        document.add_paragraph('')

    # 농협 은행 개인 앱 리뷰 현황
    document.add_paragraph('    ㅇ 농협은행')

    # 농협은행 표 생성
    NH1Table = document.add_table(3, 2, grid_t_style)

    NHCells1 = NH1Table.rows[0].cells
    NHCells1[0].paragraphs[0].add_run('긍정적 반응').bold = True
    NHCells1[0].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    # NHCells1[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    NH1TableCells1 = NH1Table.rows[0].cells
    NHCells1[1].paragraphs[0].add_run('부정적 반응').bold = True
    NH1TableCells1[1].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    # NH1TableCells1[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 표에 워드클라우드 삽입-농협
    NHCell10 = NH1Table.cell(1, 0)
    NHPara10 = NHCell10.add_paragraph()
    NHRun10 = NHPara10.add_run()
    NHRun10.add_picture("./wordcloud/개인고객/NH_WordCloud_P.png", width=Cm(7), height=Cm(5))

    NHCell11 = NH1Table.cell(1, 1)
    NHPara11 = NHCell11.add_paragraph()
    NHRun11 = NHPara11.add_run()
    NHRun11.add_picture("./wordcloud/개인고객/NH_WordCloud_N.png", width=Cm(7), height=Cm(5))

    # 긍정 빈출 단어 Top3
    NHTableCells3 = NH1Table.rows[2].cells
    NHTableCells3[0].paragraphs[0].add_run('빈출 단어 Top3\n')
    NHTableCells3[0].paragraphs[0].add_run('1\n')
    NHTableCells3[0].paragraphs[0].add_run('2\n')
    NHTableCells3[0].paragraphs[0].add_run('3\n')

    # 부정 빈출 단어 Top3
    NHTableCells3[1].paragraphs[0].add_run('빈출 단어 Top3\n')
    NHTableCells3[1].paragraphs[0].add_run('1\n')
    NHTableCells3[1].paragraphs[0].add_run('2\n')
    NHTableCells3[1].paragraphs[0].add_run('3\n')

    document.add_paragraph('')
    document.add_paragraph('')

    # 우리은행 개인 앱 리뷰 현황
    document.add_paragraph('    ㅇ 우리은행')

    # 우리은행 표 생성
    WOORI1Table = document.add_table(3, 2, grid_t_style)

    WOORICells1 = WOORI1Table.rows[0].cells
    WOORICells1[0].paragraphs[0].add_run('긍정적 반응').bold = True
    WOORICells1[0].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    # WOORICells1[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    WOORI1TableCells1 = WOORI1Table.rows[0].cells
    WOORICells1[1].paragraphs[0].add_run('부정적 반응').bold = True
    WOORI1TableCells1[1].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    # WOORI1TableCells1[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 표에 워드클라우드 삽입-우리
    WOORICell10 = WOORI1Table.cell(1, 0)
    WOORIPara10 = WOORICell10.add_paragraph()
    WOORIRun10 = WOORIPara10.add_run()
    WOORIRun10.add_picture("./wordcloud/개인고객/WOORI_WordCloud_P.png", width=Cm(7), height=Cm(5))

    WOORICell11 = WOORI1Table.cell(1, 1)
    WOORIPara11 = WOORICell11.add_paragraph()
    WOORIRun11 = WOORIPara11.add_run()
    WOORIRun11.add_picture("./wordcloud/개인고객/WOORI_WordCloud_N.png", width=Cm(7), height=Cm(5))

    # 긍정 빈출 단어 Top3
    WOORITableCells3 = WOORI1Table.rows[2].cells
    WOORITableCells3[0].paragraphs[0].add_run('빈출 단어 Top3\n')
    WOORITableCells3[0].paragraphs[0].add_run('1\n')
    WOORITableCells3[0].paragraphs[0].add_run('2\n')
    WOORITableCells3[0].paragraphs[0].add_run('3\n')

    # 부정 빈출 단어 Top3
    WOORITableCells3[1].paragraphs[0].add_run('빈출 단어 Top3\n')
    WOORITableCells3[1].paragraphs[0].add_run('1\n')
    WOORITableCells3[1].paragraphs[0].add_run('2\n')
    WOORITableCells3[1].paragraphs[0].add_run('3\n')

    document.add_paragraph('')

    # 임시 변수 !!!!!!!!!!!!!!!!!!!!!! 나중에 지우기
    bestWord = 'UI 개선'
    resultPersonal = document.add_paragraph('best 은행: ' + '\n')
    resultPersonal.add_run('결 론           ').bold = True
    resultPersonal.add_run(bestWord + '에 힘쓰는 것이 좋겠다고 판단됨.')

    # 당행 기업 앱 리뷰 현황
    document.add_paragraph(' 3. 당행 기업고객용 모바일 앱 (i-one bank) 사용자 반응 분석')

    IBKTableE = document.add_table(3, 2, grid_t_style)

    IBKTableECells1 = IBKTable.rows[0].cells
    IBKTableECells1[0].paragraphs[0].add_run('긍정적 반응').bold = True
    IBKTableECells1[0].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    # IBKTableECells1[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    IBKTableECells1 = IBKTable.rows[0].cells
    IBKTableECells1[1].paragraphs[0].add_run('부정적 반응').bold = True
    IBKTableECells1[1].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    # IBKTableECells1[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 표에 워드클라우드 삽입
    IBKECell10 = IBKTableE.cell(1, 0)
    IBKEPara10 = IBKECell10.add_paragraph()
    IBKERun10 = IBKEPara10.add_run()
    IBKERun10.add_picture("./wordcloud/기업고객/IBK_E_WordCloud_P.png", width=Cm(7), height=Cm(5))

    IBKECell11 = IBKTableE.cell(1, 1)
    IBKEPara11 = IBKECell11.add_paragraph()
    IBKERun11 = IBKEPara11.add_run()
    IBKERun11.add_picture("./wordcloud/기업고객/IBK_E_WordCloud_N.png", width=Cm(7), height=Cm(5))

    # 긍정 빈출 단어 Top3
    IBKTableECells3 = IBKTableE.rows[2].cells
    IBKTableECells3[0].paragraphs[0].add_run('빈출 단어 Top3\n')
    IBKTableECells3[0].paragraphs[0].add_run('1\n')
    IBKTableECells3[0].paragraphs[0].add_run('2\n')
    IBKTableECells3[0].paragraphs[0].add_run('3\n')

    # 부정 빈출 단어 Top3
    IBKTableECells3[1].paragraphs[0].add_run('빈출 단어 Top3\n')
    IBKTableECells3[1].paragraphs[0].add_run('1\n')
    IBKTableECells3[1].paragraphs[0].add_run('2\n')
    IBKTableECells3[1].paragraphs[0].add_run('3\n')

    document.add_paragraph('')

    # 타행 기업 앱 리뷰 현황
    document.add_paragraph(' 4. 타행 기업고객용 모바일 앱 사용자 반응 분석')

    document.add_paragraph('    ㅇ 하나은행')

    # 하나은행 표 생성
    HANA1TableE = document.add_table(3, 2, grid_t_style)

    HANA1Cells1E = HANA1TableE.rows[0].cells
    HANA1Cells1E[0].paragraphs[0].add_run('긍정적 반응').bold = True
    HANA1Cells1E[0].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    # HANA1Cells1E[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    HANA1TableCells1E = HANA1TableE.rows[0].cells
    HANA1Cells1E[1].paragraphs[0].add_run('부정적 반응').bold = True
    HANA1TableCells1E[1].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    # HANA1TableCells1E[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 표에 워드클라우드 삽입
    HANACell10E = HANA1TableE.cell(1, 0)
    HANAPara10E = HANACell10E.add_paragraph()
    HANARun10E = HANAPara10E.add_run()
    HANARun10E.add_picture("./wordcloud/기업고객/HANA_E_WordCloud_P.png", width=Cm(7), height=Cm(5))

    HANACell11E = HANA1TableE.cell(1, 1)
    HANAPara11E = HANACell11E.add_paragraph()
    HANARun11E = HANAPara11E.add_run()
    HANARun11E.add_picture("./wordcloud/기업고객/HANA_E_WordCloud_N.png", width=Cm(7), height=Cm(5))

    # 긍정 빈출 단어 Top3
    HANATableCells3E = HANA1TableE.rows[2].cells
    HANATableCells3E[0].paragraphs[0].add_run('빈출 단어 Top3\n')
    HANATableCells3E[0].paragraphs[0].add_run('1\n')
    HANATableCells3E[0].paragraphs[0].add_run('2\n')
    HANATableCells3E[0].paragraphs[0].add_run('3\n')

    # 부정 빈출 단어 Top3
    HANATableCells3E[1].paragraphs[0].add_run('빈출 단어 Top3\n')
    HANATableCells3E[1].paragraphs[0].add_run('1\n')
    HANATableCells3E[1].paragraphs[0].add_run('2\n')
    HANATableCells3E[1].paragraphs[0].add_run('3\n')

    document.add_paragraph('')
    document.add_paragraph('')
    document.add_paragraph('')

    document.add_paragraph('    ㅇ 국민은행')

    # 국민은행 표 생성
    KB1TableE = document.add_table(3, 2, grid_t_style)

    KB1Cells1E = KB1TableE.rows[0].cells
    KB1Cells1E[0].paragraphs[0].add_run('긍정적 반응').bold = True
    KB1Cells1E[0].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    # KB1Cells1E[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    KB1TableCells1E = KB1TableE.rows[0].cells
    KB1Cells1E[1].paragraphs[0].add_run('부정적 반응').bold = True
    KB1TableCells1E[1].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    # KB1TableCells1E[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 표에 워드클라우드 삽입
    KBCell10E = KB1TableE.cell(1, 0)
    KBPara10E = KBCell10E.add_paragraph()
    KBRun10E = KBPara10E.add_run()
    KBRun10E.add_picture("./wordcloud/기업고객/KB_E_WordCloud_P.png", width=Cm(7), height=Cm(5))

    KBCell11E = KB1TableE.cell(1, 1)
    KBPara11E = KBCell11E.add_paragraph()
    KBRun11E = KBPara11E.add_run()
    KBRun11E.add_picture("./wordcloud/기업고객/KB_E_WordCloud_N.png", width=Cm(7), height=Cm(5))

    # 긍정 빈출 단어 Top3
    KBTableCells3E = KB1TableE.rows[2].cells
    KBTableCells3E[0].paragraphs[0].add_run('빈출 단어 Top3\n')
    KBTableCells3E[0].paragraphs[0].add_run('1\n')
    KBTableCells3E[0].paragraphs[0].add_run('2\n')
    KBTableCells3E[0].paragraphs[0].add_run('3\n')

    # 부정 빈출 단어 Top3
    KBTableCells3E[1].paragraphs[0].add_run('빈출 단어 Top3\n')
    KBTableCells3E[1].paragraphs[0].add_run('1\n')
    KBTableCells3E[1].paragraphs[0].add_run('2\n')
    KBTableCells3E[1].paragraphs[0].add_run('3\n')

    document.add_paragraph('')
    document.add_paragraph('')

    # 신한 은행 기업 앱 리뷰 현황
    document.add_paragraph('    ㅇ 신한은행')

    # 신한은행 표 생성
    SH1TableE = document.add_table(3, 2, grid_t_style)

    SH1Cells1E = SH1TableE.rows[0].cells
    SH1Cells1E[0].paragraphs[0].add_run('긍정적 반응').bold = True
    SH1Cells1E[0].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    # SH1Cells1E[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    SH1TableCells1E = SH1TableE.rows[0].cells
    SH1Cells1E[1].paragraphs[0].add_run('부정적 반응').bold = True
    SH1TableCells1E[1].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    # SH1TableCells1E[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 표에 워드클라우드 삽입
    SHCell10E = SH1TableE.cell(1, 0)
    SHPara10E = SHCell10E.add_paragraph()
    SHRun10E = SHPara10E.add_run()
    SHRun10E.add_picture("./wordcloud/기업고객/SHINHAN_E_WordCloud_P.png", width=Cm(7), height=Cm(5))

    SHCell11E = SH1TableE.cell(1, 1)
    SHPara11E = SHCell11E.add_paragraph()
    SHRun11E = SHPara11E.add_run()
    SHRun11E.add_picture("./wordcloud/기업고객/SHINHAN_E_WordCloud_N.png", width=Cm(7), height=Cm(5))

    # 긍정 빈출 단어 Top3
    SHTableCells3E = SH1TableE.rows[2].cells
    SHTableCells3E[0].paragraphs[0].add_run('빈출 단어 Top3\n')
    SHTableCells3E[0].paragraphs[0].add_run('1\n')
    SHTableCells3E[0].paragraphs[0].add_run('2\n')
    SHTableCells3E[0].paragraphs[0].add_run('3\n')

    # 부정 빈출 단어 Top3
    SHTableCells3E[1].paragraphs[0].add_run('빈출 단어 Top3\n')
    SHTableCells3E[1].paragraphs[0].add_run('1\n')
    SHTableCells3E[1].paragraphs[0].add_run('2\n')
    SHTableCells3E[1].paragraphs[0].add_run('3\n')

    document.add_paragraph('')
    document.add_paragraph('')
    document.add_paragraph('')

    # 농협 은행 기업 고객 용 앱 리뷰 현황
    document.add_paragraph('    ㅇ 농협은행')

    # 농협은행 표 생성
    NH1TableE = document.add_table(3, 2, grid_t_style)

    NH1Cells1E = NH1TableE.rows[0].cells
    NH1Cells1E[0].paragraphs[0].add_run('긍정적 반응').bold = True
    NH1Cells1E[0].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    # NH1Cells1E[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    NH1TableCells1E = SH1TableE.rows[0].cells
    NH1Cells1E[1].paragraphs[0].add_run('부정적 반응').bold = True
    NH1TableCells1E[1].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    # NH1TableCells1E[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 표에 워드클라우드 삽입
    NHCell10E = NH1TableE.cell(1, 0)
    NHPara10E = NHCell10E.add_paragraph()
    NHRun10E = NHPara10E.add_run()
    NHRun10E.add_picture("./wordcloud/기업고객/NH_E_WordCloud_P.png", width=Cm(7), height=Cm(5))

    NHCell11E = NH1TableE.cell(1, 1)
    NHPara11E = NHCell11E.add_paragraph()
    NHRun11E = NHPara11E.add_run()
    NHRun11E.add_picture("./wordcloud/기업고객/NH_E_WordCloud_N.png", width=Cm(7), height=Cm(5))

    # 긍정 빈출 단어 Top3
    NHTableCells3E = NH1TableE.rows[2].cells
    NHTableCells3E[0].paragraphs[0].add_run('빈출 단어 Top3\n')
    NHTableCells3E[0].paragraphs[0].add_run('1\n')
    NHTableCells3E[0].paragraphs[0].add_run('2\n')
    NHTableCells3E[0].paragraphs[0].add_run('3\n')

    # 부정 빈출 단어 Top3
    NHTableCells3E[1].paragraphs[0].add_run('빈출 단어 Top3\n')
    NHTableCells3E[1].paragraphs[0].add_run('1\n')
    NHTableCells3E[1].paragraphs[0].add_run('2\n')
    NHTableCells3E[1].paragraphs[0].add_run('3\n')

    document.add_paragraph('')
    document.add_paragraph('')

    # 우리 은행 기업 고객용 앱 현황
    document.add_paragraph('    ㅇ 우리은행')

    # 우리은행 표 생성
    WOORI1TableE = document.add_table(3, 2, grid_t_style)

    WOORI1Cells1E = WOORI1TableE.rows[0].cells
    WOORI1Cells1E[0].paragraphs[0].add_run('긍정적 반응').bold = True
    WOORI1Cells1E[0].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    # WOORI1Cells1E[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    WOORI1TableCells1E = WOORI1TableE.rows[0].cells
    WOORI1Cells1E[1].paragraphs[0].add_run('부정적 반응').bold = True
    WOORI1TableCells1E[1].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    # WOORI1TableCells1E[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 표에 워드클라우드 삽입
    WOORICell10E = WOORI1TableE.cell(1, 0)
    WOORIPara10E = WOORICell10E.add_paragraph()
    WOORIRun10E = WOORIPara10E.add_run()
    WOORIRun10E.add_picture("./wordcloud/기업고객/WOORI_E_WordCloud_P.png", width=Cm(7), height=Cm(5))

    WOORICell11E = WOORI1TableE.cell(1, 1)
    WOORIPara11E = WOORICell11E.add_paragraph()
    WOORIRun11E = WOORIPara11E.add_run()
    WOORIRun11E.add_picture("./wordcloud/기업고객/WOORI_E_WordCloud_N.png", width=Cm(7), height=Cm(5))

    # 긍정 빈출 단어 Top3
    WOORITableCells3E = WOORI1TableE.rows[2].cells
    WOORITableCells3E[0].paragraphs[0].add_run('빈출 단어 Top3\n')
    WOORITableCells3E[0].paragraphs[0].add_run('1\n')
    WOORITableCells3E[0].paragraphs[0].add_run('2\n')
    WOORITableCells3E[0].paragraphs[0].add_run('3\n')

    # 부정 빈출 단어 Top3
    WOORITableCells3E[1].paragraphs[0].add_run('빈출 단어 Top3\n')
    WOORITableCells3E[1].paragraphs[0].add_run('1\n')
    WOORITableCells3E[1].paragraphs[0].add_run('2\n')
    WOORITableCells3E[1].paragraphs[0].add_run('3\n')

    document.add_paragraph('')
    document.add_paragraph('')

    # 임시 변수 !!!!!!!!!!!!!!!!!!!!!! 나중에 지우기
    bestWord2 = '기업 이미지 개선'
    resultEnterprise = document.add_paragraph('')
    resultEnterprise.add_run('결 론           ').bold = True
    resultEnterprise.add_run(bestWord2 + '에 힘쓰는 것이 좋겠다고 판단됨.')

    # 인터넷 전문 은행 앱 리뷰 현황

    # 임시 변수!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    bestWordInternet = '부가 서비스'
    resultInternet = document.add_paragraph('')
    resultInternet.add_run('결 론           ').bold = True
    resultInternet.add_run(bestWordInternet + '에 힘쓰는 것이 좋겠다고 판단됨.')

    # 마지막 꼬릿말
    document.add_paragraph('\"새로운 60년, 고객을 향한 혁신\"')

    # 문단별 정렬
    paragraph1 = document.paragraphs[0]  # 첫번째 문단
    paragraph1.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    # paragraph1.alignment = WD_ALIGN_PARAGRAPH.CENTER

    paragraph2 = document.paragraphs[2]
    paragraph2.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    # paragraph2.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    paragraphLast = document.paragraphs[-1]
    # paragraphLast.font.size=Document.shared.Pt(20)
    paragraphLast.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER  # 마지막 문단
    # paragraphLast.alignment = WD_ALIGN_PARAGRAPH.CENTER  # 마지막 문단

    # 파일 저장 // 마지막 단계
    document.save("report.docx")

    # 보고서를 이메일로 발송
    s = smtplib.SMTP('smtp.gmail.com', 587)  # gmail 포트번호 587
    s.starttls()  # TLS(Transport Layer Security) 보안
    s.login('IBK.ITgroup.2@gmail.com', 'czhoerpcnkfzqsdh')  # 메일을 보내는 계정
    # 메일 정보
    msg = MIMEMultipart()
    msg['From'] = 'IBK.ITgroup.2@gmail.com'
    msg['To'] = 'bethh05108@gmail.com'
    msg['Subject'] = datetime.today().strftime("%Y. %m. %d") + " I-one bank 사용자 반응 보고서입니다."
    # 메일 내용
    content = datetime.today().strftime("%Y. %m. %d") + "I-one bank 사용자 반응 보고서입니다."
    part2 = MIMEText(content, 'plain')
    msg.attach(part2)
    # 보고서 첨부
    part = MIMEBase('application', 'octet-stream')
    with open("report.docx", 'rb') as file:
        part.set_payload(file.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment', filename='report.docx')
    msg.attach(part)
    # 메일 전송
    s.sendmail("IBK.ITgroup.2@gmail.com", "ghlwls111@gmail.com", msg.as_string())
    # 세션 종료
    s.quit()
