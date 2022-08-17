from multiprocessing import freeze_support
import matplotlib.pyplot as plt
from collections import Counter
import pandas as pd
import re
from hanspell import spell_checker
import warnings  # 경고 알람 제거를 위한 라이브러리
from gensim.corpora import Dictionary
import gensim
from gensim.models import CoherenceModel
import pyLDAvis
# import pyLDAvis.gensim # don't skip this
import pyLDAvis.gensim_models  # don't skip this
from konlpy.tag import Okt

warnings.filterwarnings("ignore", category=DeprecationWarning)

global SW
global c
global tokenized_text
global ldamodel
# global news_list
okt = Okt()

rv_csv = ['./reviews/개인고객/HANAreview_dataset.csv', './reviews/기업고객/HANA_enterprise_review_dataset.csv',
          './reviews/개인고객/ibkbank_individual_review_dataset.csv', './reviews/기업고객/IBKreview_dataset(iONEBank기업).csv',
          './reviews/개인고객/KBreview_dataset.csv', './reviews/기업고객/KBreview_dataset(KB스타기업뱅킹).csv',
          './reviews/개인고객/NHreview_dataset.csv', './reviews/기업고객/NHreview_dataset(NH기업뱅킹).csv',
          './reviews/개인고객/WONreview_dataset.csv', './reviews/기업고객/WOORIbank_enterprise_review_dataset.csv',
          './reviews/개인고객/신한review_dataset.csv', './reviews/기업고객/SHINHANbank_enterprise_review_dataset.csv',
          './reviews/인터넷뱅크/KAKAO_review_dataset.csv', './reviews/인터넷뱅크/KBank_review_dataset.csv',
          './reviews/인터넷뱅크/TOSS_review_dataset.csv']

tm_html = ['./topicmodeling/개인고객/HANA_N.html', './topicmodeling/개인고객/HANA_P.html',
           './topicmodeling/기업고객/HANA_E_N.html', './topicmodeling/기업고객/HANA_E_P.html',
           './topicmodeling/개인고객/IBK_N.html', './topicmodeling/개인고객/IBK_P.html',
           './topicmodeling/기업고객/IBK_E_N.html', './topicmodeling/기업고객/IBK_E_P.html',
           './topicmodeling/개인고객/KB_N.html', './topicmodeling/개인고객/KB_P.html',
           './topicmodeling/기업고객/KB_E_N.html', './topicmodeling/기업고객/KB_E_P.html',
           './topicmodeling/개인고객/NH_N.html', './topicmodeling/개인고객/NH_P.html',
           './topicmodeling/기업고객/NH_E_N.html', './topicmodeling/기업고객/NH_E_P.html',
           './topicmodeling/개인고객/WOORI_N.html', './topicmodeling/개인고객/WOORI_P.html',
           './topicmodeling/기업고객/WOORI_E_N.html', './topicmodeling/기업고객/WOORI_E_P.html',
           './topicmodeling/개인고객/SHINHAN_N.html', './topicmodeling/개인고객/SHINHAN_P.html',
           './topicmodeling/기업고객/SHINHAN_E_N.html', './topicmodeling/기업고객/SHINHAN_E_P.html',
           './topicmodeling/인터넷뱅크/KAKAO_N.html', './topicmodeling/인터넷뱅크/KAKAO_P.html',
           './topicmodeling/인터넷뱅크/KBank_N.html', './topicmodeling/인터넷뱅크/KBank_P.html',
           './topicmodeling/인터넷뱅크/TOSS_N.html', './topicmodeling/인터넷뱅크/TOSS_P.html']

tm_csv = ['./topicmodeling/토픽순위/HANA_N.csv', './topicmodeling/토픽순위/HANA_P.csv',
          './topicmodeling/토픽순위/HANA_E_N.csv', './topicmodeling/토픽순위/HANA_E_P.csv',
          './topicmodeling/토픽순위/IBK_N.csv', './topicmodeling/토픽순위/IBK_P.csv',
          './topicmodeling/토픽순위/IBK_E_N.csv', './topicmodeling/토픽순위/IBK_E_P.csv',
          './topicmodeling/토픽순위/KB_N.csv', './topicmodeling/토픽순위/KB_P.csv',
          './topicmodeling/토픽순위/KB_E_N.csv', './topicmodeling/토픽순위/KB_E_P.csv',
          './topicmodeling/토픽순위/NH_N.csv', './topicmodeling/토픽순위/NH_P.csv',
          './topicmodeling/토픽순위/NH_E_N.csv', './topicmodeling/토픽순위/NH_E_P.csv',
          './topicmodeling/토픽순위/WOORI_N.csv', './topicmodeling/토픽순위/WOORI_P.csv',
          './topicmodeling/토픽순위/WOORI_E_N.csv', './topicmodeling/토픽순위/WOORI_E_P.csv',
          './topicmodeling/토픽순위/SHINHAN_N.csv', './topicmodeling/토픽순위/SHINHAN_P.csv',
          './topicmodeling/토픽순위/SHINHAN_E_N.csv', './topicmodeling/토픽순위/SHINHAN_E_P.csv',
          './topicmodeling/토픽순위/KAKAO_N.csv', './topicmodeling/토픽순위/KAKAO_P.csv',
          './topicmodeling/토픽순위/KBank_N.csv', './topicmodeling/토픽순위/KBank_P.csv',
          './topicmodeling/토픽순위/TOSS_N.csv', './topicmodeling/토픽순위/TOSS_P.csv']


def text_cleaning(doc):
    # 한국어와 띄어쓰기를 제외한 글자를 제거
    doc = re.sub("[^가-힇 ]", "", doc)
    # doc = re.sub("[^ㄱ-ㅎㅏ-ㅣ가-힇 ]", "", doc)

    # 이모티콘 제거
    EMOJI = re.compile('[\U00010000-\U0010ffff]', flags=re.UNICODE)
    doc = EMOJI.sub(r'', doc)

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
    global SW
    # tokenized_doc = []
    # for word in okt.nouns(doc):
    #     if word is not in SW and len(word) >1 :
    #         tokenized_doc.append(word)
    # return tokenized_doc

    # SW = [line.rstrip() for line in SW]
    # words = []
    # for word in okt.pos(doc, stem=True):  # 어간 추출
    #     if word[1] in ['Noun']:  # 명사, 형용사
    #         if word not in SW and len(word) > 1:
    #             words.append(word[0])
    # print(words)  # ['이런', '형식']

    #
    SW = [line.rstrip() for line in SW]
    words = [word for word in okt.nouns(doc) if word not in SW and len(word) > 1]
    print(words)

    minimum_count = 3
    remove_words = []
    for i in range(len(words)):
        tmp = words[i]
        if words.count(tmp) >= minimum_count:
            remove_words.append(tmp)
    print(remove_words)

    c = Counter(remove_words)  # 단어별 빈도수 형태의 딕셔너리 데이터
    print("----------------------------------")
    print(c)
    return c


def morpheme_Analysis(rv_csv, rate1, rate2):
    global SW
    global news_list
    text = ""
    news_list = []

    csv = pd.read_csv(rv_csv, encoding='UTF8')
    review = csv.loc[(csv['dateYear'] >= 2020) & ((csv['rating'] == rate1) | (csv['rating'] == rate2))]
    cat = review['content']
    SW = define_stopwords("./stopwords/stopwords-ko.txt")

    for t in cat:
        result = text_cleaning(t)
        result = spell_checker.check([result])
        t = result[0].checked
        print(t)
        text += (okt.normalize(t) + ". ")
    print(text)

    cleaned_text = text_cleaning(text)
    print("전처리: ", cleaned_text)

    tokenized_text = text_tokenizing(cleaned_text)
    print("/n형태소 분석(명사 추출): ", tokenized_text)

    counts = Counter(c)
    tags = counts.items()
    for x in tags:
        news_list.append(x[0])
    news_list = [d.split() for d in news_list]

    print(news_list)


def compute_coherence_values(dictionary, corpus, texts, limit, start=2, step=3):
    # 토픽 개수에 따라 Coherence score, Perplexity score 값을 반복 도출하는 사용자 정의 함수
    coherence_values = []
    perplexity_values = []
    model_list = []
    for num_topics in range(start, limit, step):
        model = gensim.models.ldamodel.LdaModel(corpus, num_topics=num_topics, id2word=dictionary, iterations=50,
                                                passes=50)
        model_list.append(model)
        coherencemodel = CoherenceModel(model=model, texts=texts, dictionary=dictionary, coherence='c_v')
        coherence_values.append(coherencemodel.get_coherence())
        perplexity_values.append(ldamodel.log_perplexity(corpus))
        print("coherence, perplexity 값 출력")
        print(coherence_values)
        print(perplexity_values)

    return model_list, coherence_values, perplexity_values


def TM(rv_csv, rate1, rate2, tm_html, tm_csv):
    global ldamodel

    # 형태소 분석
    morpheme_Analysis(rv_csv, rate1, rate2)
    dct = Dictionary(news_list)  # 형태소 분석을 통해 만든 명사 리스트를 사전으로 생성

    # 출현빈도가 적거나 자주 등장하는 단어 제거
    dct.filter_extremes(no_below=1)
    print(dct)

    corpus = [dct.doc2bow(text) for text in news_list]  # 코퍼스 생성

    NUM_TOPICS = 3  # 토픽의 개수를 지정하여 진행하는 경우
    ldamodel = gensim.models.ldamodel.LdaModel(corpus, num_topics=NUM_TOPICS, id2word=dct, iterations=50, passes=40)

    # Coherence score : 값이 클수록 정확한 데이터
    # Perplexity score : 값이 작을수록 정확한 데이터

    # Ldamodel = gensim.models.ldamodel.LdaModel(corpus, num_topics=NUM_TOPICS, id2word=dct, iterations=50)
    # model_list, coherence_values, perplexity_values = compute_coherence_values(dictionary=dct, corpus=corpus,
    #                                                                            texts=news_list, start=3, limit=6,
    #                                                                            step=1)

    limit = 40
    start = 2
    step = 6
    model_list, coherence_values, perplexity_values = compute_coherence_values(dictionary=dct, corpus=corpus,
                                                                               texts=news_list, start=start,
                                                                               limit=limit,
                                                                               step=step)

    # 토픽의 개수를 3~5개로 제한하여 Coherence score, Perplexity score 값을 연속적으로 도출 (6, 3, 1)
    x = range(start, limit, step)
    plt.plot(x, coherence_values)
    plt.xlabel("Num Topics")
    plt.ylabel("Coherence score")
    plt.legend(("coherence_values"), loc='best')
    plt.show()

    x = range(start, limit, step)
    plt.plot(x, perplexity_values)
    plt.xlabel("Num Topics")
    plt.ylabel("Perplexity score")
    plt.legend(("perplexity_values"), loc='best')
    plt.show()

    # pyLDAvis.enable_notebook()

    # ldamodel = model_list[2]  # Coherence score 값이 가장 크고, Perplexity score 값이 가장 작은 점

    lda_display = pyLDAvis.gensim_models.prepare(ldamodel, corpus, dct, sort_topics=False)
    # pyLDAvis.display(lda_display)
    pyLDAvis.save_html(lda_display, tm_html)

    # 단어의 해당 토픽에 대한 기여도 보기
    df = pd.DataFrame()
    topics = ldamodel.print_topics(num_words=5)

    # csv로 저장
    for i in range(len(topics)):
        print(topics[i])
        topic = str(topics[i])
        topic = re.sub("[^가-힇 ]", "", topic)
        topic_list = topic.split(' ')
        topic_list = [v for v in topic_list if v]
        df[i] = topic_list
    print(df)
    df.to_csv(tm_csv, encoding='cp949', index=False)


def main():
    # 하나은행 (개인 부정긍정, 기업 부정긍정 순)
    TM(rv_csv[0], 1, 2, tm_html[0], tm_csv[0])
    TM(rv_csv[0], 4, 5, tm_html[1], tm_csv[1])
    TM(rv_csv[1], 1, 2, tm_html[2], tm_csv[2])
    TM(rv_csv[1], 4, 5, tm_html[3], tm_csv[3])

    # 기업은행 (개인 부정긍정, 기업 부정긍정 순)
    TM(rv_csv[2], 1, 2, tm_html[4], tm_csv[4])
    TM(rv_csv[2], 4, 5, tm_html[5], tm_csv[5])
    TM(rv_csv[3], 1, 2, tm_html[6], tm_csv[6])
    TM(rv_csv[3], 4, 5, tm_html[7], tm_csv[7])

    # 국민은행 (개인 부정긍정, 기업 부정긍정 순)
    TM(rv_csv[4], 1, 2, tm_html[8], tm_csv[8])
    TM(rv_csv[4], 4, 5, tm_html[9], tm_csv[9])
    TM(rv_csv[5], 1, 2, tm_html[10], tm_csv[10])
    TM(rv_csv[5], 4, 5, tm_html[11], tm_csv[11])

    # 농협은행 (개인 부정긍정, 기업 부정긍정 순)
    TM(rv_csv[6], 1, 2, tm_html[12], tm_csv[12])
    TM(rv_csv[6], 4, 5, tm_html[13], tm_csv[13])
    TM(rv_csv[7], 1, 2, tm_html[14], tm_csv[14])
    TM(rv_csv[7], 4, 5, tm_html[15], tm_csv[15])

    # 우리은행 (개인 부정긍정, 기업 부정긍정 순)
    TM(rv_csv[8], 1, 2, tm_html[16], tm_csv[16])
    TM(rv_csv[8], 4, 5, tm_html[17], tm_csv[17])
    TM(rv_csv[9], 1, 2, tm_html[18], tm_csv[18])
    TM(rv_csv[9], 4, 5, tm_html[19], tm_csv[19])

    # 신한은행 (개인 부정긍정, 기업 부정긍정 순)
    TM(rv_csv[10], 1, 2, tm_html[20], tm_csv[20])
    TM(rv_csv[10], 4, 5, tm_html[21], tm_csv[21])
    TM(rv_csv[11], 1, 2, tm_html[22], tm_csv[22])
    TM(rv_csv[11], 4, 5, tm_html[23], tm_csv[23])

    # 카카오뱅크
    TM(rv_csv[12], 1, 2, tm_html[24], tm_csv[24])
    TM(rv_csv[12], 4, 5, tm_html[25], tm_csv[25])

    # 케이뱅크
    TM(rv_csv[13], 1, 2, tm_html[26], tm_csv[26])
    TM(rv_csv[13], 4, 5, tm_html[27], tm_csv[27])

    # 토스
    TM(rv_csv[14], 1, 2, tm_html[28], tm_csv[28])
    TM(rv_csv[14], 4, 5, tm_html[29], tm_csv[29])


if __name__ == '__main__':
    freeze_support()
    main()
