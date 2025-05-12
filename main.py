import mailbox
from email.header import decode_header, make_header
import stanza

mbox = mailbox.mbox('Inbox_my')
nlp = stanza.Pipeline('ru', processors='tokenize,pos,lemma', use_gpu=False)
ind = 0
for msg in mbox:

    # Отправитель
    # from_header = str(make_header(decode_header(msg["From"])))


    # # Тема
    # subject = str(make_header(decode_header(msg["Subject"])))


    # Тело письма
    body = ""

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



    # Лемматизируем весь текст
    doc = nlp(body.lower())
    lemmas_in_text = 'паролями' in " ".join([word.lemma for sent in doc.sentences for word in sent.words])

    if lemmas_in_text:
        print(body)
        ind += 1


print(ind)
