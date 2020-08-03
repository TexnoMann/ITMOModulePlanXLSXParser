# Описание

Програма для создания JSON фала из XLSX файла с расписанием модулей.

## Установка

Используйте пакетный менеджер [pip3](https://pip.pypa.io/en/stable/) для установки всех необходимых пакетов.
```bash
git clone https://github.com/TexnoMann/ITMOModulePlanXLSXParser.git
```

```bash
pip3 install -r requirements.txt
```

## Запуск
Для запуска скрипта парсинга **модульного плана** из консоли(Ubuntu):
```bash
python3 parse_json.py [-in|--input Путь до xlsx файла] [-out|--output Путь для сохранения сгенерированного JSON]
```
Для запуска скрипта парсинга **информации о занятости аудиторий** из консоли(Ubuntu):
```bash
python3 parse_rooms_json.py [-in|--input Путь до docx файла] [-out|--output Путь до выходного json файла с занятостью] [-rinfo|--rooms_info Путь до json файла с информацией об аудиториях] [-tc|--time_config Путь до конфига с таблицей времени]
```

Для запуска скрипта парсинга **информации о расписании занятий** из консоли(Ubuntu):
```bash
python3 parse_lessons_plan_json.py [-in|--input Путь до xlsx файла] [-out|--output Путь до выходного json файла с рассписанием] [-tc|--time_config Путь до конфига с таблицей времени]
```

## License
[MIT](https://choosealicense.com/licenses/mit/)
