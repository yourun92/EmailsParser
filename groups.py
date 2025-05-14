import pandas as pd
import pymorphy3

groups = {
    "Группы": [
        "Торговые компании",
        "Строители монолитчики",
        "Строители кровельщики",
        "Строители фасадчики",
        "Проектировщики",
        "Строители дорожники",
        "Строители отделочники"
    ],
    "Ключи в email": [
        "torg, tsk, tk, td",
        "monolit, beton",
        "krovlia, krovla, krov, roof",
        "fasad",
        "project, proekt, arch",
        "dorog, road",
        "fghfgjfgjfgnv"
    ],
    "Ключи в теле": [
        "Для торговой компании, торговая компания, перепродажа, менеджер по продажам, менеджер отдела продаж",
        "Гидрошпонка, фиксатор арматуры, бетон, арматура, эмульсол, опалубка, опалубочные, БДК, балка, замок, ригель, бентонитовый, гернитовый, бентонит, гернит",
        "Техноэласт, рулонная кровля, кровля, крыша, Roof, кровельная, рулонка, грибки для мембраны, сварка кровли, горелка для кровли, для кровли",
        "Фасад, фасадные, фасадная, вентфасад, СФТК, штукатурная, штукатурного",
        "Проект, архитектурное, проектирование, проектирую, разработки проекта, разработка",
        "Дорожное полотно, георешетка, геосетка, дороги, асфальт, дорожное",
        "Отделка, для отделки, гипсовые, из гипса, гипсокартон, краска, краски, лакокрасочный, окрашивание, ковролин, линолеум, потолок, для стен, внутри помещения"
    ]
}

position = [
    'старший',
    'менеджер',
    'по снабжению',
    'проектировщик',
    'руководитель',
    'руководитель отдела',
    'директор',
    'генеральный директор',
    'омтс',
    'специалист по',
    'специалист отдела',
    'заместитель директора',
    'генеральный директор',
    'ведущий специалист',
    'инженер по',
    'инженер-проектировщик',
    'менеджер отдела',
    'cпециалист отдела снабжения',
    'менеджер по закупкам',
    'Ведущий специалист отдела продаж',
    'Отдел Материально-технического снабжения',
    'отдел закупок',
    'начальник отдела',
    'отдел'
]

delete_mails = ['Тендер', 'процедура', '7713416124', '7713783501', '7713416124', 'отписаться от рассылки', 'сметчик', 'смета', 'карго', 'снижение цен', 'оплата по накладной', 'metalloinvest', 'ozon']

morph = pymorphy3.MorphAnalyzer()
keys_in_groups = [group.split(', ') for group in groups['Ключи в теле']]

cases = ['nomn', 'gent', 'datv', 'accs', 'ablt', 'loct']
numbers = ['sing', 'plur']

new_keys_in_groups = []

for group in keys_in_groups:
    data = []
    for phrase in group:
        phrs = phrase.split()
        parses = [morph.parse(p)[0] for p in phrs]
        pos_tags = [p.tag.POS for p in parses]

        data.append(phrase.lower())
        if len(phrs) != 1 and 'NOUN' in pos_tags:
            noun = parses[pos_tags.index('NOUN')]

            for number in numbers:
                for case in cases:
                    new_words = []
                    for p in parses:
                        # Для существительного — склоняем с учетом рода
                        if p.tag.POS == 'NOUN':
                            inflected = p.inflect({case, number})
                        # Для прилагательного — с согласованием по роду и числу
                        elif p.tag.POS == 'ADJF':
                            if number == 'sing':
                                inflected = p.inflect({case, number, noun.tag.gender})
                            else:
                                inflected = p.inflect({case, number})
                        else:
                            inflected = p  # оставляем без изменений, если не умеем склонять

                        new_words.append(inflected.word if inflected else p.word)

                    data.append(' '.join(new_words))

        else:
            data.extend([f.word for f in morph.parse(phrase)[0].lexeme])

    data = list(dict.fromkeys(data))
    new_keys_in_groups.append(data)
