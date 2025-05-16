import pandas as pd
from groups_to_filter import groups, position_list
import re
from tqdm import tqdm
tqdm.pandas()

def extract_name_email(text):
    text = text.strip()

    # 1) Если есть формат: Имя <почта>
    match = re.match(r'^(.*?)\s*<([^>]+)>$', text)
    if match:
        name = match.group(1).strip()
        email = match.group(2).strip()
        return name, email

    # 2) Если просто почта (содержит @)
    match = re.search(r'[\w\.-]+@[\w\.-]+', text)
    if match:
        email = match.group(0)
        # Попытка взять всё, что до почты как имя
        name_part = text[:match.start()].strip()
        name = name_part if name_part else None
        return name, email

    # 3) Если ничего не найдено, возвращаем None
    return '', ''


def find_phones(text):
    if not isinstance(text, str):
        return None

    phone_pattern = re.compile(
        r'''
        (?:
            (?:\+7|8)                # код страны
            [\s\u00A0\-]+            # ← теперь хотя бы один пробел/тире
            \(?\d{3,5}\)?            # код города
            [\s\u00A0\-]*\d{2,3}
            [\s\u00A0\-]*\d{2}
            [\s\u00A0\-]*\d{2,4}
            (?:\s*(?:доб\.?|ext\.?)\s*\d{1,5})?
        )
        ''',
        re.VERBOSE
    )

    phones = []
    for match in phone_pattern.finditer(text):
        number = match.group().strip()
        digits_only = re.sub(r'\D', '', number)
        # допфильтрация: длина и наличие хотя бы одного разделителя
        if (
            10 <= len(digits_only) <= 15
            and re.search(r'[\s\-()]', number)  # ← есть хотя бы один разделитель
        ):
            phones.append(number)

    return ' | '.join(set(phones))

def find_position(position_list, text):
    if not isinstance(text, str):
        return 'не определено'

    text = text.lower()
    for position, values in position_list.items():
        for v in values:
            if v.lower() in text:
                return position.lower()

    return 'не определено'


def detect_group(sender, body):
    if not isinstance(body, str) or not isinstance(sender, str):
        return 'группа не определена', ''

    sender = sender.lower()
    body = body.lower()

    sender_full = sender
    sender_domain = sender.split('@')[-1]

    for k, v in groups.items():

        # По доменной части отправителя
        for keyword in v[1]:
            if keyword:
                keyword = keyword.lower()
                if keyword in ('td', 'tk'):
                    pattern = rf"(^|[\W_-]){re.escape(keyword)}($|[\W_-])"
                    if re.search(pattern, sender_domain):
                        return k, keyword
                elif keyword == 'tsk':
                    sender_domain = sender_domain.split('.')[0]
                    if sender_domain.startswith('tsk') or sender_domain.endswith('tsk'):
                        return k, keyword
                elif keyword in sender_domain:
                    return k, keyword

        # По полному адресу отправителя
        for keyword in v[0]:
            if keyword:
                keyword = keyword.lower()
                if keyword in sender_full:
                    return k, keyword

        # По телу письма
        for keyword in v[2]:
            keyword = keyword.lower()
            if keyword in ('мост', 'краска', 'краски', 'замок'): # ищем не по вхождению а по началу с
                pattern = rf"\b{re.escape(keyword)}"
                pattern1 = rf"{re.escape(keyword)}\b"
                if re.search(pattern, body) or re.search(pattern1, body):
                    return k, keyword
            elif keyword in body:
                return k, keyword

    return 'группа не определена', ''

def first_not_undefined(series):
    for val in series:
        if val and str(val).lower() != 'не определено':
            return val
    return series.iloc[0]



df = pd.read_excel('output.xlsx')

# Применяем функцию
print('Определение группы')
result = df.progress_apply(lambda row: detect_group(row['sender'], row['body']), axis=1)
df[['Группа', 'Ключ']] = pd.DataFrame(result.tolist(), index=df.index)

print('Ищем номера телефонов')
df['Телефон'] = df['body'].progress_apply(find_phones)

print('Ищем должность')
df['Должность'] = df['body'].progress_apply(lambda x: find_position(position_list, x))

print('Емаил и имя')
df[['Имя', 'Почта']] = df['sender'].progress_apply(lambda x: pd.Series(extract_name_email(x)))

df_unique = df.groupby('Почта', dropna=False).agg({
    'Имя': 'first',
    'Группа': first_not_undefined,
    'Должность': first_not_undefined,
    'Телефон': 'first',
    'Ключ': 'first',
    'body': 'first'
}).reset_index()

df_unique.to_excel('mail_parser.xlsx', index=False)