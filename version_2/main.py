from natasha import NewsEmbedding, NewsMorphTagger, Segmenter, Doc, NamesExtractor, MorphVocab, NewsNERTagger
import pandas as pd
from groups_to_filter import groups, position_list
import re
from tqdm import tqdm
tqdm.pandas()

embedding = NewsEmbedding()
morph_tagger = NewsMorphTagger(embedding)
ner_tagger = NewsNERTagger(embedding)
segmenter = Segmenter()

morph = MorphVocab()
names_extractor = NamesExtractor(morph)


def extract_name_email(sender, body):
    sender = sender.strip()

    # 1) Если есть формат: Имя <почта>
    match = re.match(r'^(.*?)\s*<([^>]+)>$', sender)
    if match:
        names = extract_name_with_natasha(body) + [match.group(1).strip()]
        name = max(names, key=len)
        email = match.group(2).strip()
        return name, email

    # 2) Если просто почта (содержит @)
    match = re.search(r'[\w\.-]+@[\w\.-]+', sender)
    if match:
        email = match.group(0)
        names = extract_name_with_natasha(
            body) + [sender[:match.start()].strip()]
        name = max(names, key=len) if names else None
        return name, email

    return '', ''


def extract_name_with_natasha(text):
    if not text:
        return []
    doc = Doc(text)
    doc.segment(segmenter)
    doc.tag_morph(morph_tagger)
    doc.tag_ner(ner_tagger)

    persons = []
    for span in doc.spans:
        if span.type == 'PER':
            span.normalize(morph)
            persons.append(span.normal)
    return persons



def find_phones(text):
    if not isinstance(text, str):
        return None

    phone_pattern = re.compile(
        r'''
        (?:(?:тел\.?|моб\.?)\s*[:\-]?\s*)?  # необязательное "тел." или "моб."
        (?:
            (?:\+7|8)                       # код страны
            [\s\-()]*\d{3,5}               # код города (возможно в скобках)
            [\s\-]*\d{2,3}                 # часть номера
            [\s\-]*\d{2}                   # часть номера
            [\s\-]*\d{2,4}                 # часть номера
            (?:\s*(?:доб\.?|ext\.?)\s*\d{1,5})?  # добавочный номер
        )
        ''',
        re.VERBOSE | re.IGNORECASE
    )

    phones = []
    for match in phone_pattern.finditer(text):
        number = match.group().strip()
        digits_only = re.sub(r'\D', '', number)
        if 10 <= len(digits_only) <= 15:
            phones.append(number)

    return ' | '.join(set(phones))


def find_position(position_list, text):
    if not isinstance(text, str):
        return 'не определено'

    text = text.lower()
    for position, values in position_list.items():
        for v in values:
            v = v.lower()
            if v == 'проект':
                if any(i in text for i in ['уважением', 'пожеланиями']):
                    regards = text.split('уважением' if 'уважением' in text else 'пожеланиями')[-1]
                    pattern = rf"\b{re.escape(v)}\b"
                    if re.search(pattern, regards):
                        return position
            elif v in text:
                return position

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
            if keyword == 'мост':
                pattern = r'\bмост\w*|\w*мост\b'
                if re.search(pattern, body):
                    return k, keyword
            elif keyword in ('краска', 'краски', 'замок'):
                pattern = rf"\b{re.escape(keyword)}\b"
                if re.search(pattern, body):
                    return k, keyword
            elif keyword == 'проект':
                pattern = rf"\b{re.escape(keyword)}\b"
                if re.search(pattern, body):
                    return k, keyword
            elif keyword == 'строй':
                pattern = rf'\b{re.escape(keyword)}'
                if re.search(pattern, body):
                    return k, keyword
            elif keyword in body:
                return k, keyword

    return 'группа не определена', ''


def first_not_undefined(series):
    for val in series:
        if val and str(val).lower() != 'не определено':
            return val
    return series.iloc[0]


df = pd.read_parquet('output_last_v2.parquet')

# Применяем функцию
print('Определение группы')
result = df.progress_apply(lambda row: detect_group(
    row['sender'], row['body']), axis=1)
df[['Группа', 'Ключ']] = pd.DataFrame(result.tolist(), index=df.index)

print('Ищем номера телефонов')
df['Телефон'] = df['body'].progress_apply(find_phones)

print('Ищем должность')
df['Должность'] = df['body'].progress_apply(lambda x: find_position(position_list, x))

print('Емаил и имя')
df[['Имя', 'Почта']] = df.progress_apply(
    lambda row: pd.Series(extract_name_email(row['sender'], row['body'])),
    axis=1
)

print('Почти готово')
df_unique = df.groupby('Почта', dropna=False).agg({
    'Имя': 'first',
    'Группа': first_not_undefined,
    'Должность': first_not_undefined,
    'Телефон': 'first',
    'Ключ': 'first',
    'body': 'first'
}).reset_index()

df_unique.to_excel('mail_parser_2.xlsx', index=False)