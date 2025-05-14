import pandas as pd
from groups import groups, new_keys_in_groups, delete_mails, position
import tqdm
import re


# Функция для удаления строк с нежелательными словами
def delete_mail_with_na_words(body):
    if not isinstance(body, str):
        return False
    body_lower = body.lower()
    for word in delete_mails:  # предполагается, что delete_mails - это список слов
        if re.search(rf'\b{re.escape(word.lower())}\b', body_lower):
            return True  # Если слово найдено, удаляем строку
    return False  # Если слов из delete_mails нет, не удаляем

def delete_mail_with_re(subject):
    if not isinstance(subject, str):
        return False
    subject_lower = subject.lower().strip()
    return subject_lower.startswith('re')

def delete_mail_with_enix(sender):
    if not isinstance(sender, str):
        return False
    sender_lower = sender.lower().strip()
    return 'enix' in sender_lower


def find_position(position_list, text):
    if not isinstance(text, str):
        return None

    text_lower = text.lower().split('уважением')[-1].split('\n')
    for position in position_list:
        for t in text_lower:
            if position.lower() in t:
                return t.replace(',', '').strip()

    return None


def find_phones(text):
    if not isinstance(text, str):
        return None

    phone_pattern = re.compile(
        r'''
        (?:
            (?:\+7|8)                # код страны
            [\s\u00A0\-]*            # возможные пробелы/тире
            \(?\d{3,5}\)?            # код города
            [\s\u00A0\-]*\d{2,3}    # первая часть номера
            [\s\u00A0\-]*\d{2}       # вторая часть
            [\s\u00A0\-]*\d{2,4}     # третья часть
            (?:\s*(?:доб\.?|ext\.?)\s*\d{1,5})? # добавочный, опционально
        )
        ''',
        re.VERBOSE
    )

    # фильтруем по количеству цифр (не менее 10, чтобы отсечь тех. номера)
    phones = []
    for match in phone_pattern.finditer(text):
        number = match.group().strip()
        digits_only = re.sub(r'\D', '', number)
        if 10 <= len(digits_only) <= 15:
            phones.append(number)

    return ' | '.join(set(phones))


def detect_group(sender, body):
    if not isinstance(body, str) or not isinstance(sender, str):
        return 'группа не определена', None, None, None  # Возвращаем 4 значения

    sender_lower = sender.lower()
    body_lower = body.lower()

    for group, key_in_sender, key_in_body in zip(groups['Группы'], groups['Ключи в email'], new_keys_in_groups):

        # ищем в адресе отправителя
        for keyword in key_in_sender.split(', '):
            keyword_lower = keyword.lower()
            if keyword_lower in ('tsk', 'tk', 'td'):
                parts = re.split(r'[@\-._]', sender_lower)
                if any(part.startswith(keyword_lower) for part in parts):
                    pos = find_position(position, body)
                    phones = find_phones(body)

                    return group, keyword, pos, phones
            else:
                if keyword_lower in sender_lower:
                    pos = find_position(position, body)
                    phones = find_phones(body)

                    return group, keyword, pos, phones

        # ищем в теле письма
        for keyword in key_in_body:
            # if keyword.lower() in body_lower:
            if re.search(rf'\b{re.escape(keyword.lower())}\b', body_lower):
                pos = find_position(position, body)
                phones = find_phones(body)

                return group, keyword, pos, phones

    # Если ничего не найдено
    return 'группа не определена', None, None, None



df = pd.read_excel('output.xlsx')

# Применяем функцию для удаления строк с нежелательными словами, переодправленные сообщения и с почтами от enix
df = df[~df['body'].apply(delete_mail_with_na_words)]
df = df[~df['subject'].apply(delete_mail_with_re)]
df = df[~df['sender'].apply(delete_mail_with_enix)]

df['Группа'], df['Определено по ключу'], df['Должность'], df['Телефон'] = zip(*[
    detect_group(sender, body)
    for sender, body in tqdm.tqdm(zip(df['sender'], df['body']), total=len(df), desc="Определение групп")
])

df = df.rename({'sender': 'Отправитель', 'subject': 'Тема', 'body': 'Тело письма'}, axis='columns')
df.to_excel('mail_parser.xlsx', index=False)
