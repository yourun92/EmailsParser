import mailbox
from email.header import decode_header, make_header
import pandas as pd
import re
import tqdm
from bs4 import BeautifulSoup
import os
from email.utils import parsedate_to_datetime

def clean_illegal_chars(x):
    if isinstance(x, str):
        # Удаляет все символы с кодами < 32, кроме табуляции, перевода строки и возврата каретки
        return re.sub(r'[\x00-\x08\x0B-\x0C\x0E-\x1F]', '', x)
    return x

def delete_unnecessary_messages(body, sender, subject):
    sender = sender.lower()
    subject = subject.lower()
    body = body.lower()

    # # Удалить сообщения с повторной отправкой
    # if subject.startswith('re'):
    #     return True

    # Проверка по отправителю
    if any(domain in sender for domain in ['batissur.ru', 'specstali.ru', 'enix.ru', 'enyx.ru']):
        return True

    # Ключевые фразы для фильтрации
    blacklist_phrases = [
        'оплата по накладной',
        'отписаться от рассылки',
        'снижение цен',
        'рассылка',
        'подтвердите подписку',
        'unsubscribe'
    ]

    if any(phrase in body for phrase in blacklist_phrases):
        return True

    return False

def extract_text_from_msg(msg):
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
                        return body.strip()

                # fallback to HTML if no plain text
                if content_type == "text/html" and "attachment" not in content_disposition and not body:
                    payload = part.get_payload(decode=True)
                    if payload:
                        html = payload.decode(part.get_content_charset() or "utf-8", errors="ignore")
                        soup = BeautifulSoup(html, "html.parser")
                        body = soup.get_text(separator="\n")  # преобразуем в читаемый текст
        else:
            payload = msg.get_payload(decode=True)
            if payload:
                content_type = msg.get_content_type()
                charset = msg.get_content_charset() or "utf-8"
                decoded = payload.decode(charset, errors="ignore")
                if content_type == "text/plain":
                    body = decoded
                elif content_type == "text/html":
                    soup = BeautifulSoup(decoded, "html.parser")
                    body = soup.get_text(separator="\n")
    except:
        pass

    return body.strip()


def get_data(mbox):
    data = []
    for msg in tqdm.tqdm(mbox, desc='Обработка писем', unit='шт.'):
        try:
            from_header = str(make_header(decode_header(msg.get("From", ""))))
        except:
            from_header = ''
        try:
            subject = str(make_header(decode_header(msg.get("Subject", ""))))
        except:
            subject = ''

        body = extract_text_from_msg(msg)

        # Получаем дату письма
        try:
            date_str = msg.get("Date", "")
            # Преобразуем строку даты в datetime
            date = parsedate_to_datetime(date_str) if date_str else None
        except:
            date = None

        row = {
            'sender': from_header,
            'subject': subject,
            'body': body,
            'date': date
        }
        data.append(row)
    return data

folder = r'D:\data\enix_info\extracted\info@enix.ru'

# Собираем только .mbox-файлы, которые начинаются с Inbox и не заканчиваются на .msf
mbox_files = [
    os.path.join(folder, f)
    for f in os.listdir(folder)
    if f.startswith('Inbox') and not f.endswith('.msf') and os.path.isfile(os.path.join(folder, f))
]

all_data = []

print("Найдено файлов:", len(mbox_files))

for filename in mbox_files:
    if os.path.exists(filename):
        print(f'Чтение данных из файла: {filename}')
        mbox = mailbox.mbox(filename)
        data = get_data(mbox)
        all_data.extend(data)
    else:
        print(f'Файл не найден: {filename}')

# Объединение всех данных
df = pd.DataFrame(all_data)

print('Удаление ненужных сообщений')
df_clean = df[~df.apply(lambda row: delete_unnecessary_messages(row['body'], row['sender'], row['subject']), axis=1)]
df_clean = df_clean.apply(lambda x: x.apply(clean_illegal_chars))

# Сохраняем в parquet
df_clean.to_parquet("output_last_v2.parquet", index=False)
print(f'Данные записаны в output.parquet — всего {len(df_clean)} писем')