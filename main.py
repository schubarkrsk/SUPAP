import win32api
import win32print
import glob
import pathlib


def print_pdf(input_pdf_file, mode=2):
    printer_name = win32print.GetDefaultPrinter()
    # тут нужные права на использование принтеров
    print_defaults = {"DesiredAccess": win32print.PRINTER_ALL_ACCESS}
    # начинаем работу с принтером ("открываем" его)
    handle = win32print.OpenPrinter(printer_name, print_defaults)
    # Если изменить level на другое число, то не сработает
    level = 2
    # Получаем значения принтера
    attributes = win32print.GetPrinter(handle, level)
    # Настройка двухсторонней печати
    attributes['pDevMode'].Duplex = mode  # flip over  3 - это короткий 2 - это длинный край

    ## Передаем нужные значения в принтер
    win32print.SetPrinter(handle, level, attributes, 0)
    win32print.GetPrinter(handle, level)['pDevMode'].Duplex
    ## Предупреждаем принтер о старте печати
    win32print.StartDocPrinter(handle, 1, [input_pdf_file, None, "raw"])
    ## 2 в начале для открытия pdf и его сворачивания, для открытия без сворачивания поменяйте на 1
    win32api.ShellExecute(2, 'print', input_pdf_file, '.', '/manualstoprint', 0)
    ## "Закрываем" принтер
    win32print.ClosePrinter(handle)

if __name__ == "__main__":
    CURRENT_DIRECTORY = pathlib.Path.cwd()
    pdf_files = glob.glob("*.pdf")
    print(f"Печать {len(pdf_files)} файлов из каталога {CURRENT_DIRECTORY}")
    for file in pdf_files:
        path = pathlib.Path(CURRENT_DIRECTORY, file)
        print(path)
        print_pdf(str(path), 2)

    print("Печать завершена!")
