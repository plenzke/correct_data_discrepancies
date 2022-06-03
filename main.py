import re
import pandas as pd

df = pd.read_excel('client_acc.xls', header=0)

def convert(name: str): # Функция для перевода ФИО в единый формат /Фамилия И.О./
    mod_name = name.split(' ')
    temp = '.'.join([mod_name[i][0] for i in range(1, len(mod_name))])

    if temp == '':
        return mod_name[-1]
    else:
        return mod_name[0] + ' ' + temp + '.'

def get_initials(fullname): # Функция для вытаскивания инициалов из ФИО
    name_list = fullname.split()

    initials = ''

    for n in name_list:
        initials += n[0].upper()

    return initials

def create_pattern(initials): # Функция для создания паттерна поиска ФИО на основе инициала имени (инициал имени присутствует всегда)
    """Первое слово - Загл. буква, далее не пробел, не точка, не запятая, не косая черта, далее могут идти строчные буквы,
     Второе слово - Заглавная буква инициала имени + строчные или одна заглавная как инициал,
     после - пробел, пустая строка или точка для отделения инициала,
     Третье слово - аналогично второму без инициала имени."""
    pattern = fr'(([А-Я][^A-Я\s\.\,\/][a-я]*)(\s+)([{initials[1]}][a-я]*)' + \
              fr'(\.*\s*|\s+|\B)([А-Я]?\.*[a-я]*)?)'

    return pattern

def to_single_format(substring, sample, orig): # Функция для перевода столбца 'Наименование счета' в единый формат
    orig = orig.replace('Ё', 'Е')
    orig = orig.replace('ё', 'е')
    if len(sample) != 0:
        correct_str = orig.replace(substring, sample)
        index = correct_str.find(sample)
        string = " ".join([correct_str[:index], correct_str[index:]])
        string = re.sub(' +', ' ', string)
        return string
    else:
        return orig


def text_is_upper(text): # Функция для проверки текста на процент заглавных букв
    text = re.sub('[^А-Яа-я]', '', text)
    try:
        percent_upper = sum(map(str.isupper, text)) / len(text)
        if percent_upper > 0.8:
            return True
        else:
            return False
    except Exception as e:
        return False

def find_name(text, initials): # Функция для поиска ФИО в тексте на основе созданного паттерна
    text = str(text)
    text = text.replace('Ё', 'Е')
    text = text.replace('ё', 'е')
    text = re.sub(r'\s+', ' ', text)

    pattern = create_pattern(initials)

    if not text_is_upper(text):
        name = re.findall(pattern, text)
        if name:
            return name[0][0]
        else:
            return 'пусто, ' + 'пусто, ' + 'пусто'
    else:
        try:
            start_number_index = re.search(
                r'(\bРВПС\w{0,} )|(\bРВПС\. )',
                # При случае когда текст написан заглавными буквами, ищем по ключевым словам по типу (РВПС, П/отчет и т.д.) после которого идет вхождение имени
                # В данном случае написано только для РВПС
                text, flags=re.I).end()
            text_cut = text[start_number_index:]
            FIO = text_cut.replace('.', ' ')
            FIO = re.sub(r'\s+', ' ', FIO).split(' ')
            if len(FIO) >= 3:
                return FIO[0] + ' ' + FIO[1] + ' ' + FIO[2]
            elif len(FIO) == 2:
                return FIO[0] + ' ' + FIO[1] + ' ' + 'пусто'
            elif len(FIO) == 1:
                return FIO[0] + ' ' + 'пусто' + 'пусто'
        except Exception:
            return 'пусто, ' + 'пусто, ' + 'пусто'

df['INITIALS'] = df['ФИО'].apply(get_initials) # Создаем столбец инициалов
df['SAMPLE'] = df['ФИО'].apply(convert) # Создаем столбец образца
df['NAME_SUBSTRING'] = df.apply(lambda x: find_name(x['Наименование счета'], x['INITIALS']), axis=1) # Ищем подстроку с ФИО
df['Correct'] = df.apply(lambda x: to_single_format(x['NAME_SUBSTRING'], x['SAMPLE'], x['Наименование счета']), axis=1) # Приводим ФИО к единому формату
df.loc[(df['Наименование счета'] != df['Correct']), 'Наименование счета'] = df['Correct'] # Заменяем значения в столбце 'Наименование счета'

# Выгружаем данные в Excel
writer = pd.ExcelWriter('output.xlsx')
df[['№ клиента', 'ФИО', 'Наименование счета']].to_excel(writer, 'marks')
writer.save()








