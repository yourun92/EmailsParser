import mailbox
from email.header import decode_header, make_header
import pandas as pd
import re


mbox = mailbox.mbox('Inbox')
# nlp = stanza.Pipeline('ru', processors='tokenize,pos,lemma', use_gpu=False)
data = []
ind = 0
for msg in mbox:

    try:
        # Отправитель
        from_header = str(make_header(decode_header(msg["From"])))
    except:
        from_header = ''

    try:
        # Тема
        subject = str(make_header(decode_header(msg["Subject"])))
    except:
        subject = ''

    # Тело письма
    body = ""

    try:
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = part.get("Content-Disposition", "")

                if content_type == "text/plain" and "attachment" not in content_disposition:
                    payload = part.get_payload(decode=True)
                    if payload:
                        body = payload.decode(part.get_content_charset() or "utf-8", errors="ignore")
                        break
        else:
            payload = msg.get_payload(decode=True)
            if payload:
                body = payload.decode(msg.get_content_charset() or "utf-8", errors="ignore")
    except:
        body = ''

    row = {
        'sender': from_header,
        'subject': subject,
        'body': body
    }

    data.append(row)
    ind += 1

    if ind % 500 == 0:
        print(f'Обработано {ind} писем')


    # # Лемматизируем весь текст
    # doc = nlp(body.lower())
    # lemmas_in_text = 'паролями' in " ".join([word.lemma for sent in doc.sentences for word in sent.words])

    # if lemmas_in_text:
    #     print(body)
    #     ind += 1

def clean_illegal_chars(x):
    if isinstance(x, str):
        # Удаляет все символы с кодами < 32, кроме табуляции, перевода строки и возврата каретки
        return re.sub(r'[\x00-\x08\x0B-\x0C\x0E-\x1F]', '', x)
    return x


df = pd.DataFrame(data)
df_clean = df.apply(lambda col: col.map(clean_illegal_chars))
df_clean.to_excel("output.xlsx", index=False)
