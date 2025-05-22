import pandas as pd
from email_validator import validate_email, EmailNotValidError


ind = 0


def check_for_valid(email):
    global ind

    ind += 1

    if ind % 10 == 0:
        print(f'{ind} - email отработано')
    try:
        emailinfo = validate_email(email, check_deliverability=True)
        email = emailinfo.normalized
        return 'valid and normalized'
    except EmailNotValidError as e:
        # Возвращаем сообщение об ошибке для отладки
        return f'is invalid: {str(e)}'
    except Exception as e:
        return f'Error: {str(e)}'  # Ловим другие ошибки


df = pd.read_excel(
    'mail_parser_1.xlsx', sheet_name='Sheet1')

# Создаём новую колонку со значением по умолчанию
df['check_for_valied_email'] = ''

# Фильтруем по условию: только где "Группа" != "группа не определена"
mask = df['Группа'].str.lower().str.strip() != 'группа не определена'

# Применяем проверку только к этим строкам
df.loc[mask, 'check_for_valied_email'] = df.loc[mask, 'Почта'].apply(check_for_valid)

df.to_excel('Парсинг почты.xlsx', index=False)
