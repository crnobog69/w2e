# <p align="center">Word 2 Excel</p>

### <p align="center">(w2e)</p>

---

> [!WARNING]
> Овај алат је специјализован само за одређене кориснике. Употреба на документима који нису прилагођени за рад са овим алатом може да створити грешке и нестабилност у резултату.

---

## Опис
Специјализовани `Python` алат за конверзију `Word` докумената у `Excel` табеле, са посебним фокусом на очување форматирања и обраду нумеричких података. Идеалан за обраду финансијских и табеларних докумената.

## Могућности
- Аутоматска обрада свих `.docx` датотека у тренутном директоријуму
- Паметна детекција и конверзија бројчаних формата
- Очување форматирања текста (подебљано, поравнање)
- Интелигентно прилагођавање ширине колона
- Посебна обрада табеларних структура
- Подршка за више пасуса у ћелији
- Аутоматско креирање излазног директоријума (`ex/`)

## Инсталација
```bash
git clone https://github.com/crnobog69/w2e.git
cd w2e
pip install -r requirements.txt
```

## Употреба
Једноставно поставите ваше .docx датотеке у исти директоријум са скриптом и покрените:
```bash
python w2e.py
```
Конвертоване датотеке ће бити сачуване у `ex/` директоријуму са истим именом али .xlsx екстензијом.

## Технички детаљи
- Конвертује европски формат бројева (200.000,00) у стандардни формат
- Користи `Calibri` фонт (11pt за табеле, 12pt за обичан текст)
- Имплементира паметно преламање текста
- Чува подебљани текст из изворних докумената
- Аутоматски прилагођава висину редова за преломљени садржај

## Захтеви
- Python 3.7+
- pandas
- python-docx
- openpyxl
- setuptools

> [!NOTE]
> Активирајте виртуелно окружење ако га користите:
> <br>
> ```bash
> source ~/orbit-server/venv/bin/activate

## Лиценца
MIT Лиценца

## Доприноси
Доприноси су добродошли! Слободно пошаљите [Pull Request](https://github.com/crnobog69/w2e/pulls).

## Подршка
Ако наиђете на проблеме, молимо вас да их пријавите у GitHub систему за праћење проблема.
