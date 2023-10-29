import streamlit as st
from docx import Document
import io
import zipfile

def read_docx(file):
    doc = Document(io.BytesIO(file.read()))
    fullText = []
    for para in doc.paragraphs:
        fullText.append(para.text)
    return '\n'.join(fullText)

def correct_text(text, corrections):
    for wrong, correct in corrections.items():
        text = text.replace(wrong, f'<span style="color:red;">{correct}</span>')
    return text

st.title('表記揺れチェックツール')

uploaded_files = st.file_uploader("WordまたはMarkdownファイルをアップロードしてください", type=["docx", "md"], accept_multiple_files=True)

# ユーザーが選択できる表記揺れのリスト
options = {
    "下さ": "くださ",
    "頂": "いただ",
    "虫歯": "むし歯",
    "出来": "でき",
    "致し": "いたし"
}

# ユーザーが選択した表記揺れ
selected_options = st.multiselect("修正したい表記揺れを選択してください", list(options.keys()))

# ユーザーが自由に入力できるキーワードのリスト
user_keywords = st.text_area("追加でチェックしたいキーワードを「キーワード:変換後の文字」の形式で1行ずつ入力してください").split('\n')

# 開始ボタン
if st.button('開始'):
    # ユーザーが選択した表記揺れと自由に入力したキーワードを修正するための辞書
    corrections = {key: options[key] for key in selected_options}
    corrections.update({key.split(':')[0]: key.split(':')[1] for key in user_keywords if ':' in key})

    if len(uploaded_files) == 1:
        # Only one file, so download it directly
        uploaded_file = uploaded_files[0]
        if uploaded_file.type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
            # This is a Word file (.docx)
            text = read_docx(uploaded_file)
        elif uploaded_file.type in ["text/markdown", "application/octet-stream"]:
            # This is a Markdown file (.md)
            text = uploaded_file.read().decode('utf-8')
        else:
            st.error("Unsupported file type: " + uploaded_file.type)

        st.write('アップロードされたファイルの内容:')
        st.write(text)

        corrected_text = correct_text(text, corrections)
        st.write('修正後のテキスト:')
        st.markdown(corrected_text, unsafe_allow_html=True)

        # Create a download button for the corrected text
        st.download_button(
            label="修正後のファイルをダウンロード",
            data=corrected_text.encode('utf-8'),
            file_name='corrected.md',
            mime='text/markdown',
        )
    else:
        # Multiple files, so zip them
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'a', zipfile.ZIP_DEFLATED, False) as zip_file:
            for i, uploaded_file in enumerate(uploaded_files):
                if uploaded_file is not None:
                    if uploaded_file.type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
                        # This is a Word file (.docx)
                        text = read_docx(uploaded_file)
                    elif uploaded_file.type in ["text/markdown", "application/octet-stream"]:
                        # This is a Markdown file (.md)
                        text = uploaded_file.read().decode('utf-8')
                    else:
                        st.error("Unsupported file type: " + uploaded_file.type)
                        continue

                    st.write('アップロードされたファイルの内容:')
                    st.write(text)

                    corrected_text = correct_text(text, corrections)
                    st.write('修正後のテキスト:')
                    st.markdown(corrected_text, unsafe_allow_html=True)

                    # Add corrected text to the zip file
                    zip_file.writestr(f'corrected_{i}.md', corrected_text)

        # Finish the zip file
        zip_file.close()
        zip_buffer.seek(0)

        # Create a download button for the zip file
        st.download_button(
            label="修正後のファイルをダウンロード",
            data=zip_buffer.getvalue(),
            file_name='corrected_files.zip',
            mime='application/zip',
        )