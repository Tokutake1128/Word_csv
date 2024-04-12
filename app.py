import streamlit as st
from docx import Document
import pandas as pd
from io import BytesIO

def load_word_file(docx_file):
    doc = Document(docx_file)
    text = [p.text for p in doc.paragraphs if p.text]
    return text

def split_sentences(text):
    sentences = []
    for paragraph in text:
        # パラグラフを句点で分割
        parts = paragraph.split('。')
        for i, part in enumerate(parts):
            if i < len(parts) - 1:  # 最後の要素以外は必ず「。」を追加
                sentences.append(part + '。')
            else:
                if part:  # 最後の部分が空でなければ、リストに追加
                    sentences.append(part)
    return sentences

def save_sentences_to_csv(sentences, filename='output.csv'):
    df = pd.DataFrame(sentences, columns=['Sentence'])
    df.to_csv(filename, index=False)
    return filename

st.title('Word→CSV変換')

uploaded_file = st.file_uploader("Word ファイルをアップロードしてください", type=['docx'])
if uploaded_file is not None:
    with st.spinner('ファイルを処理中...'):
        text = load_word_file(uploaded_file)
        sentences = split_sentences(text)
        output_file = save_sentences_to_csv(sentences)
        st.success('処理が完了しました！')

        st.download_button(
            label="CSV ファイルをダウンロード",
            data=open(output_file, "rb"),
            file_name="sentences.csv",
            mime='text/csv',
        )