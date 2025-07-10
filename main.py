import pptx.exc
from pptx.enum.shapes import *
from pptx import *


def fractional_in_text_checker(some_text: str) -> str:
    """
    Функция для извлечения дробных чисел из входной строки.
    :param some_text: Входная строка.
    :return: Строку с дробными числами, записанными через \n, если таковые имеются.
    """
    some_text_copy = some_text.split()
    normalized_text = some_text.replace(",", ".").replace("/", ".").split()
    j = 0
    fractional_numbers = ''
    for word in normalized_text:
        try:
            if len(word) == 1:
                if ord(word) in [188, 189, 190]:
                    fractional_numbers += word + "\n"
            else:
                float(word)
                if "." in word:
                    fractional_numbers += (some_text_copy[j]) + "\n"
        except ValueError:
            pass
        j += 1
    return fractional_numbers


def chart_handler(chart_shape):
    """
    Функция для извлечения информации из диаграмм.
    :param chart_shape: Объект Shape, пренадлежащий MSO_SHAPE_TYPE.CHART и представляющий собой диаграмму.
    :return: Строку, содержащую все значения, присутствующие на диаграмме.
    """
    text = ''
    for series in chart_shape.chart.series:
        for value in series.values:
            text += f" {value}"
    return text


def table_handler(table_shape):
    """
    Функция для извлечения информации из таблиц.
    :param table_shape: Объект Shape, пренадлежащий MSO_SHAPE_TYPE.TABLE и представляющий собой таблицу.
    :return: Строку, содержащую все значения, присутствующие в таблице.
    """
    text = ''
    for row in table_shape.table.rows:
        for cell in row.cells:
            text += f" {cell.text}"
    return text


def slide_processing(slide, slide_number):
    """
    Функция для обработки слайдов .pptx презентации и извлечения инфорамции из них.
    :param slide: Объект Slide, представляющий собой слайд
    :param slide_number: Целое число, отражающее номер обрабатываемого слайда
    :return: Выводит в консоль номера всех слайдов, где встречаются дробные числа,
    а также сами числа, по одному в строке.
    """
    fractional_numbers = ''
    for shape in slide.shapes:
        try:
            if shape.has_text_frame:
                numbers = fractional_in_text_checker(shape.text)
                if numbers:
                    fractional_numbers += numbers
            else:
                match shape.shape_type:
                    case MSO_SHAPE_TYPE.TABLE:
                        numbers = fractional_in_text_checker(table_handler(shape))
                        if numbers:
                            fractional_numbers += numbers
                    case MSO_SHAPE_TYPE.CHART:
                        numbers = fractional_in_text_checker(chart_handler(shape))
                        if numbers:
                            fractional_numbers += numbers
                    case _:
                        pass
        except Exception as err:
            print(f"Ошибка обработчика слайдов: {err}")

    if len(fractional_numbers) > 0:
        print(f"Слайд номер {slide_number}\n" + fractional_numbers)
        return True


if __name__ == '__main__':
    try:
        path = input("Введите путь до .pptx презентации: ")
        if not (path.endswith(".pptx")):
            path += ".pptx"
        try:
            prs = Presentation(path)
            slide_number = 1
            are_there_fractionals = ''
            for slide in prs.slides:
                are_there_fractionals += str(slide_processing(slide, slide_number))
                slide_number += 1
            if not ("True" in are_there_fractionals):
                print("Дробных чисел в презентации не найдено")

        except pptx.exc.PackageNotFoundError:
            print(
                "По указанному пути презентации .pptx не найдено.\nПроверьте правильность указания ссылки и тип файла")

    except Exception as e:
        print(f"Произошла непредвиденная ошибка: {e}")
