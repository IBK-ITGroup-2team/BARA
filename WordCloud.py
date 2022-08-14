from wordcloud import WordCloud
import matplotlib.pyplot as plt
from collections import Counter
import pandas as pd
import re
from konlpy.tag import Okt

okt = Okt()
global SW
global c


def text_cleaning(doc):
    # 한국어와 띄어쓰기를 제외한 글자를 제거
    doc = re.sub("[^ㄱ-ㅎㅏ-ㅣ가-힇 ]", "", doc)

    # 이모티콘 제거
    EMOJI = re.compile('[\U00010000-\U0010ffff]', flags=re.UNICODE)
    doc = EMOJI.sub(r'', doc)

    # py-hanspell 맞춤법 검사 추가

    return doc


def define_stopwords(path):
    global SW
    SW = set()

    with open(path, encoding='utf-8') as f:
        for word in f:
            SW.add(word)
        return SW


def text_tokenizing(doc):  # 형태소 분석
    global c
    # tokenized_doc = []
    # for word in okt.nouns(doc):
    #     if word is not in SW and len(word) >1 :
    #         tokenized_doc.append(word)
    # return tokenized_doc

    words = [word for word in okt.nouns(doc) if word not in SW and len(word) > 1]
    c = Counter(words)  # 단어별 빈도수 형태의 딕셔너리 데이터
    return c


def create_WordCloud(path, rate1, rate2, colormap, savepath):
    global SW
    global top3
    text = ""

    # csv = pd.read_csv('./dataset/개인고객/HANAreview_dataset.csv',
    csv = pd.read_csv(path, encoding='UTF8')
    find_row = csv.loc[(csv['dateYear'] >= 2020) & ((csv['rating'] == rate1) | (csv['rating'] == rate2))]
    cat = find_row['content']
    print(cat)
    for t in cat:
        text += (okt.normalize(t) + ". ")
    print(text)
    # print(okt.pos(text, norm=True, stem=True))

    SW = define_stopwords("C:\\Users\\Cyber\\PycharmProjects\\Bank\\stopwords-ko.txt")
    cleaned_text = text_cleaning(text)
    print("전처리: ", cleaned_text)

    tokenized_text = text_tokenizing(cleaned_text)
    print("/n형태소 분석(명사 추출): ", tokenized_text)

    # 가장 많이 나온 단어부터 3개 저장
    top3 = []
    counts = Counter(c)
    tags = counts.most_common(3)
    for x in tags:
        top3.append(x[0])
    print(top3)

    wc = WordCloud(font_path='C:\\Windows\\Fonts\\malgun.ttf', colormap=colormap,
                   width=400, height=400, scale=2.0, max_words=130, max_font_size=250)
    gen = wc.generate_from_frequencies(c)
    plt.figure()
    plt.imshow(gen)
    wc.to_file(savepath)


# 은행 별 부정,긍정 워드클라우드 생성 (개인, 기업 순)
if __name__ == '__main__':
    # 하나은행 (개인)
    create_WordCloud('./reviews/개인고객/HANAreview_dataset.csv', 1, 2, 'Oranges_r', './wordcloud/개인고객/HANA_WordCloud_N.png')
    HANA_negative_top3 = top3

    create_WordCloud('./reviews/개인고객/HANAreview_dataset.csv', 4, 5, 'RdBu_r', './wordcloud/개인고객/HANA_WordCloud_P.png')
    HANA_positive_top3 = top3

    # 하나은행 (기업)
    create_WordCloud('./reviews/기업고객/HANA_enterprise_review_dataset.csv', 1, 2, 'Oranges_r',
              './wordcloud/기업고객/HANA_E_WordCloud_N.png')
    HANA_E_negative_top3 = top3

    create_WordCloud('./reviews/기업고객/HANA_enterprise_review_dataset.csv', 4, 5, 'RdBu_r',
              './wordcloud/기업고객/HANA_E_WordCloud_P.png')
    HANA_E_positive_top3 = top3

    # IBK기업은행 (개인)
    create_WordCloud('./reviews/개인고객/ibkbank_individual_review_dataset.csv', 1, 2, 'Oranges_r',
              './wordcloud/개인고객/IBK_WordCloud_N.png')
    IBK_negative_top3 = top3

    create_WordCloud('./reviews/개인고객/ibkbank_individual_review_dataset.csv', 4, 5, 'RdBu_r',
              './wordcloud/개인고객/IBK_WordCloud_P.png')
    IBK_positive_top3 = top3

    # IBK기업은행 (기업)
    create_WordCloud('./reviews/기업고객/IBKreview_dataset(iONEBank기업).csv', 1, 2, 'Oranges_r',
              './wordcloud/기업고객/IBK_E_WordCloud_N.png')
    IBK_E_negative_top3 = top3

    create_WordCloud('./reviews/기업고객/IBKreview_dataset(iONEBank기업).csv', 4, 5, 'RdBu_r',
              './wordcloud/기업고객/IBK_E_WordCloud_P.png')
    IBK_E_positive_top3 = top3

    # 국민은행 (개인)
    create_WordCloud('./reviews/개인고객/KBreview_dataset.csv', 1, 2, 'Oranges_r',
              './wordcloud/개인고객/KB_WordCloud_N.png')
    KB_negative_top3 = top3

    create_WordCloud('./reviews/개인고객/KBreview_dataset.csv', 4, 5, 'RdBu_r',
              './wordcloud/개인고객/KB_WordCloud_P.png')
    KB_positive_top3 = top3

    # 국민은행 (기업)
    create_WordCloud('./reviews/기업고객/KBreview_dataset(KB스타기업뱅킹).csv', 1, 2, 'Oranges_r',
              './wordcloud/기업고객/KB_E_WordCloud_N.png')
    KB_E_negative_top3 = top3

    create_WordCloud('./reviews/기업고객/KBreview_dataset(KB스타기업뱅킹).csv', 4, 5, 'RdBu_r',
              './wordcloud/기업고객/KB_E_WordCloud_P.png')
    KB_E_positive_top3 = top3

    # 농협은행 (개인)
    create_WordCloud('./reviews/개인고객/NHreview_dataset.csv', 1, 2, 'Oranges_r',
              './wordcloud/개인고객/NH_WordCloud_N.png')
    NH_negative_top3 = top3

    create_WordCloud('./reviews/개인고객/NHreview_dataset.csv', 4, 5, 'RdBu_r',
              './wordcloud/개인고객/NH_WordCloud_P.png')
    NH_positive_top3 = top3

    # 농협은행 (기업)
    create_WordCloud('./reviews/기업고객/NHreview_dataset(NH기업뱅킹).csv', 1, 2, 'Oranges_r',
              './wordcloud/기업고객/NH_E_WordCloud_N.png')
    NH_E_negative_top3 = top3

    create_WordCloud('./reviews/기업고객/NHreview_dataset(NH기업뱅킹).csv', 4, 5, 'RdBu_r',
              './wordcloud/기업고객/NH_E_WordCloud_P.png')
    NH_E_positive_top3 = top3

    # 우리은행 (개인)
    create_WordCloud('./reviews/개인고객/WONreview_dataset.csv', 1, 2, 'Oranges_r',
              './wordcloud/개인고객/WOORI_WordCloud_N.png')
    WOORI_negative_top3 = top3

    create_WordCloud('./reviews/개인고객/WONreview_dataset.csv', 4, 5, 'RdBu_r',
              './wordcloud/개인고객/WOORI_WordCloud_P.png')
    WOORI_positive_top3 = top3

    # 우리은행 (기업)
    create_WordCloud('./reviews/기업고객/WOORIbank_enterprise_review_dataset.csv', 1, 2, 'Oranges_r',
              './wordcloud/기업고객/WOORI_E_WordCloud_N.png')
    WOORI_E_negative_top3 = top3

    create_WordCloud('./reviews/기업고객/WOORIbank_enterprise_review_dataset.csv', 4, 5, 'RdBu_r',
              './wordcloud/기업고객/WOORI_E_WordCloud_P.png')
    WOORI_E_positive_top3 = top3

    # 신한은행 (개인)
    create_WordCloud('./reviews/개인고객/신한review_dataset.csv', 1, 2, 'Oranges_r',
              './wordcloud/개인고객/SHINHAN_WordCloud_N.png')
    SHINHAN_negative_top3 = top3

    create_WordCloud('./reviews/개인고객/신한review_dataset.csv', 4, 5, 'RdBu_r',
              './wordcloud/개인고객/SHINHAN_WordCloud_P.png')
    SHINHAN_positive_top3 = top3

    # 신한은행 (기업)
    create_WordCloud('./reviews/기업고객/SHINHANbank_enterprise_review_dataset.csv', 1, 2, 'Oranges_r',
              './wordcloud/기업고객/SHINHAN_E_WordCloud_N.png')
    SHINHAN_E_negative_top3 = top3

    create_WordCloud('./reviews/기업고객/SHINHANbank_enterprise_review_dataset.csv', 4, 5, 'RdBu_r',
              './wordcloud/기업고객/SHINHAN_E_WordCloud_P.png')
    SHINHAN_E_positive_top3 = top3