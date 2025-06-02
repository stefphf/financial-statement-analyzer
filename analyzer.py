import pandas as pd
import warnings
import os
import re

def suppress_warnings():
    """Подавление предупреждений openpyxl"""
    warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl.styles.stylesheet')

def get_file_path():
    """Запрос пути к файлу у пользователя"""
    while True:
        file_path = input("Введите полный путь к файлу Excel: ").strip()
        if os.path.isfile(file_path):
            return file_path
        print("Файл не найден. Пожалуйста, введите правильный путь.")

def parse_excel_file(file_path):
    """Функция для чтения и обработки данных из Excel файла"""
    # Чтение листа с информацией об организации
    company_sheet = pd.read_excel(file_path, sheet_name="Сведения об организации", header=None)
    company_name = None
    for index, row in company_sheet.iterrows():
        if isinstance(row[0], str) and "Полное наименование юридического лица" in row[0]:
            company_name = row[7]
            break

    # Чтение листа с бухгалтерским балансом
    balance_sheet = pd.read_excel(file_path, sheet_name="Бухгалтерский баланс", header=None)

    # Чтение листа с отчетом о финансовых результатах
    income_statement = pd.read_excel(file_path, sheet_name="Отчет о финансовых результатах", header=None)

    # Определение года отчета
    report_year = None
    for index, row in balance_sheet.iterrows():
        if isinstance(row[0], str) and "На" in row[0] and "г." in row[0]:
            # Ищем год в строке типа "На 31 декабря 2024 г."
            year_match = re.search(r'\b(20\d{2})\b', row[0])
            if year_match:
                report_year = year_match.group(1)
                break

    if not report_year:
        for index, row in income_statement.iterrows():
            if isinstance(row[0], str) and "За" in row[0] and "г." in row[0]:
                # Ищем год в строке типа "За 2024 г."
                year_match = re.search(r'\b(20\d{2})\b', row[0])
                if year_match:
                    report_year = year_match.group(1)
                    break

    if not report_year:
        # Если год не найден, попробуем извлечь из имени файла
        file_name = os.path.basename(file_path)
        year_match = re.search(r'_(\d{4})_', file_name)
        if year_match:
            report_year = year_match.group(1)
        else:
            report_year = "отчетный"

    # Извлечение данных из бухгалтерского баланса
    balance_data = {}
    for index, row in balance_sheet.iterrows():
        if isinstance(row[3], str):
            if "БАЛАНС" in row[3] and isinstance(row[8], str) and "1600" in row[8]:
                balance_data["total_assets"] = float(str(row[10]).replace(" ", "").replace("-", "0"))
            elif "Нераспределенная прибыль" in row[3] and isinstance(row[8], str) and "1370" in row[8]:
                balance_data["retained_earnings"] = float(str(row[10]).replace(" ", "").replace("-", "0"))
            elif "Итого по разделу III" in row[3] and isinstance(row[8], str) and "1300" in row[8]:
                balance_data["total_equity"] = float(str(row[10]).replace(" ", "").replace("-", "0"))
            elif "Кредиторская задолженность" in row[3] and isinstance(row[8], str) and "1520" in row[8]:
                balance_data["accounts_payable"] = float(str(row[10]).replace(" ", "").replace("-", "0"))
            elif "Итого по разделу V" in row[3] and isinstance(row[8], str) and "1500" in row[8]:
                balance_data["total_liabilities"] = float(str(row[10]).replace(" ", "").replace("-", "0"))

    # Если сумма обязательств не найдена, вычисляем ее
    if "total_liabilities" not in balance_data and "total_assets" in balance_data and "total_equity" in balance_data:
        balance_data["total_liabilities"] = balance_data["total_assets"] - balance_data["total_equity"]

    # Извлечение данных из отчета о финансовых результатах
    income_data = {}
    for index, row in income_statement.iterrows():
        if isinstance(row[4], str):
            if "Выручка" in row[4] and isinstance(row[9], str) and "2110" in row[9]:
                income_data["revenue"] = float(str(row[12]).replace(" ", "").replace("-", "0"))
            elif "Чистая прибыль" in row[4] and isinstance(row[9], str) and "2400" in row[9]:
                income_data["net_profit"] = float(str(row[12]).replace(" ", "").replace("-", "0"))
            elif "Прибыль (убыток) до налогообложения" in row[4] and isinstance(row[9], str) and "2300" in row[9]:
                income_data["ebt"] = float(str(row[12]).replace(" ", "").replace("-", "0"))

    return company_name, balance_data, income_data, report_year

def calculate_ratios(balance_data, income_data):
    """Функция для расчета финансовых показателей"""
    ratios = {}

    # Проверка наличия необходимых данных
    required_income = ["revenue", "net_profit"]
    required_balance = ["total_assets", "total_equity", "total_liabilities"]

    missing = [k for k in required_income if k not in income_data] + [k for k in required_balance if
                                                                      k not in balance_data]
    if missing:
        raise ValueError(f"Не удалось извлечь необходимые данные: {', '.join(missing)}")

    # Показатели рентабельности
    ratios["net_profit_margin"] = (income_data["net_profit"] / income_data["revenue"]) * 100 if income_data[
                                                                                                    "revenue"] != 0 else 0
    ratios["return_on_assets"] = (income_data["net_profit"] / balance_data["total_assets"]) * 100 if balance_data[
                                                                                                         "total_assets"] != 0 else 0
    ratios["return_on_equity"] = (income_data["net_profit"] / balance_data["total_equity"]) * 100 if balance_data[
                                                                                                         "total_equity"] != 0 else 0

    # Показатели автономности (финансовой устойчивости)
    ratios["equity_ratio"] = (balance_data["total_equity"] / balance_data["total_assets"]) * 100 if balance_data[
                                                                                                        "total_assets"] != 0 else 0
    ratios["debt_to_equity"] = (balance_data["total_liabilities"] / balance_data["total_equity"]) * 100 if balance_data[
                                                                                                               "total_equity"] != 0 else 0

    # Добавляем основные показатели для вывода
    ratios["net_profit_value"] = income_data["net_profit"]
    ratios["revenue_value"] = income_data["revenue"]
    ratios["equity_value"] = balance_data["total_equity"]
    ratios["assets_value"] = balance_data["total_assets"]
    ratios["liabilities_value"] = balance_data["total_liabilities"]

    return ratios


def print_results(company_name, ratios, report_year):
    """Функция для вывода результатов"""
    print(f"\nРезультаты анализа финансовых показателей {company_name} за {report_year} год:")
    print("=" * 60)

    # Форматируем вывод с выравниванием по правому краю
    print(f"Рентабельность продаж (ROS), %        : {ratios['net_profit_margin']:>10.2f} %")
    print(f"Рентабельность активов (ROA), %       : {ratios['return_on_assets']:>10.2f} %")
    print(f"Рентабельность капитала (ROE), %      : {ratios['return_on_equity']:>10.2f} %")
    print(f"Коэффициент автономности, %           : {ratios['equity_ratio']:>10.2f} %")
    print(f"Коэффициент финансового левериджа, %  : {ratios['debt_to_equity']:>10.2f} %")
    print(f"Чистая прибыль, тыс. руб.             : {ratios['net_profit_value']:>10.2f} тыс. руб.")
    print(f"Выручка, тыс. руб.                    : {ratios['revenue_value']:>10.2f} тыс. руб.")
    print(f"Собственный капитал, тыс. руб.        : {ratios['equity_value']:>10.2f} тыс. руб.")
    print(f"Всего активов, тыс. руб.              : {ratios['assets_value']:>10.2f} тыс. руб.")
    print(f"Обязательства, тыс. руб.              : {ratios['liabilities_value']:>10.2f} тыс. руб.")
    print("=" * 60)


def get_indicator_analysis(indicator, value):
    """Возвращает анализ и оценку для выбранного показателя"""
    analysis = ""
    recommendation = ""

    if indicator == "net_profit_margin":
        analysis = "Рентабельность продаж (ROS) показывает, сколько копеек чистой прибыли получает компания с каждого рубля выручки.\n"
        if value > 20:
            assessment = "Отличная рентабельность (выше среднерыночной)"
        elif value > 10:
            assessment = "Хорошая рентабельность (на уровне среднерыночной)"
        elif value > 5:
            assessment = "Удовлетворительная рентабельность (ниже среднерыночной)"
        else:
            assessment = "Низкая рентабельность (требуется анализ издержек и ценовой политики)"
        recommendation = "Для повышения ROS можно оптимизировать издержки, пересмотреть ценообразование или ассортимент."

    elif indicator == "return_on_assets":
        analysis = "Рентабельность активов (ROA) показывает эффективность использования активов компании.\n"
        if value > 15:
            assessment = "Отличная эффективность использования активов"
        elif value > 8:
            assessment = "Хорошая эффективность использования активов"
        elif value > 3:
            assessment = "Удовлетворительная эффективность использования активов"
        else:
            assessment = "Низкая эффективность использования активов"
        recommendation = "Для повышения ROA можно оптимизировать активы или увеличить оборачиваемость."

    elif indicator == "return_on_equity":
        analysis = "Рентабельность капитала (ROE) показывает доходность для акционеров.\n"
        if value > 25:
            assessment = "Отличная доходность для акционеров"
        elif value > 15:
            assessment = "Хорошая доходность для акционеров"
        elif value > 8:
            assessment = "Удовлетворительная доходность"
        else:
            assessment = "Низкая доходность для акционеров"
        recommendation = "Для повышения ROE можно увеличить прибыль или оптимизировать структуру капитала."

    elif indicator == "equity_ratio":
        analysis = "Коэффициент автономности показывает долю активов, финансируемых за счет собственных средств.\n"
        if value > 60:
            assessment = "Высокая финансовая устойчивость"
        elif value > 40:
            assessment = "Удовлетворительная финансовая устойчивость"
        else:
            assessment = "Низкая финансовая устойчивость (зависимость от заемных средств)"
        recommendation = "Оптимальное значение 50-70%. Слишком высокое значение может указывать на неэффективное использование заемного капитала."

    elif indicator == "debt_to_equity":
        analysis = "Коэффициент финансового левериджа показывает соотношение заемных и собственных средств.\n"
        if value < 50:
            assessment = "Консервативная структура капитала"
        elif value < 100:
            assessment = "Умеренная долговая нагрузка"
        elif value < 150:
            assessment = "Высокая долговая нагрузка"
        else:
            assessment = "Очень высокая долговая нагрузка (рискованная структура капитала)"
        recommendation = "Оптимальное значение зависит от отрасли, обычно 50-100%."

    return f"\n{analysis}Текущее значение: {value:.2f}%\nОценка: {assessment}\nРекомендация: {recommendation}"


def ask_for_analysis(ratios):
    """Запрашивает у пользователя показатель для анализа"""
    indicators = {
        "1": ("Рентабельность продаж (ROS)", "net_profit_margin"),
        "2": ("Рентабельность активов (ROA)", "return_on_assets"),
        "3": ("Рентабельность капитала (ROE)", "return_on_equity"),
        "4": ("Коэффициент автономности", "equity_ratio"),
        "5": ("Коэффициент финансового левериджа", "debt_to_equity")
    }

    print("\n" + "=" * 60)
    print("ДЕТАЛЬНЫЙ АНАЛИЗ ПОКАЗАТЕЛЕЙ".center(60))
    print("=" * 60)
    print("Выберите показатель для подробного анализа:")
    for key, (name, _) in indicators.items():
        print(f"{key}. {name}")
    print("0. Завершить анализ")

    while True:
        choice = input("\nВведите номер показателя (1-5) или 0 для выхода: ")
        if choice == "0":
            break
        if choice in indicators:
            indicator_key = indicators[choice][1]
            value = ratios[indicator_key]
            print(get_indicator_analysis(indicator_key, value))
            print("-" * 60)
        else:
            print("Неверный ввод. Пожалуйста, выберите номер от 1 до 5 или 0 для выхода.")


def main():
    suppress_warnings()  # Подавляем предупреждения openpyxl

    print("АНАЛИЗ ФИНАНСОВОЙ ОТЧЕТНОСТИ")
    print("=" * 50)

    file_path = get_file_path()

    try:
        # Чтение и обработка данных из файла
        print("\nОбработка файла...")
        company_name, balance_data, income_data, report_year = parse_excel_file(file_path)

        # Расчет финансовых показателей
        ratios = calculate_ratios(balance_data, income_data)

        # Вывод результатов
        print_results(company_name, ratios, report_year)

        # Детальный анализ показателей
        ask_for_analysis(ratios)

    except Exception as e:
        print(f"\nОШИБКА: {str(e)}")
        print("Проверьте структуру файла и наличие всех необходимых данных.")


if __name__ == "__main__":
    main()
