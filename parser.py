import pandas as pd
from groups import groups, new_keys_in_groups
import tqdm
import re

def detect_group(sender, body):
    if not isinstance(body, str) or not isinstance(sender, str):
        return 'группа не определена', None

    sender_lower = sender.lower()
    body_lower = body.lower()

    for group, key_in_sender, key_in_body in zip(groups['Группы'], groups['Ключи в email'], new_keys_in_groups):

        # ищем в адресе отправителя
        for keyword in key_in_sender.split(', '):
            keyword_lower = keyword.lower()
            if keyword_lower in ('tsk', 'tk', 'td'):
                parts = re.split(r'[@\-._]', sender_lower)
                if any(part.startswith(keyword_lower) for part in parts):
                    return group, keyword
            else:
                if keyword_lower in sender_lower:
                    return group, keyword

        # ищем в теле письма
        for keyword in key_in_body:
            # if keyword.lower() in body_lower:
            if re.search(rf'\b{re.escape(keyword.lower())}\b', body_lower):
                return group, keyword

    return 'группа не определена', None


df = pd.read_excel('output.xlsx')

df['Группа'], df['Определено по ключу'] = zip(*[
    detect_group(sender, body)
    for sender, body in tqdm.tqdm(zip(df['sender'], df['body']), total=len(df), desc="Определение групп")
])

df = df.rename({'sender': 'Отправитель', 'subject': 'Тема', 'body': 'Тело письма'}, axis='columns')
df.to_excel('mail_parser.xlsx', index=False)
