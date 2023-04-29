import dearpygui.dearpygui as dpg
from excel import Excel_regex


#Параметры окна
title = 'Excel regex'
width, height = 500, 500
always_on_top = True
window_color = (210, 210, 210)
font_color = (0,0,0)
resizable = False
no_title_bar = True
no_move = True
no_background = True
small_icon = 'icon2.ico'

excel_regex = Excel_regex()

#Параметры для полей ввода
x, y = 50, 32
button_width = 25
row_widgets, column_widgets = [], []

def find_matches() -> None:
    '''Собирает все введенные данные, делает проверку и после прохождения проверок запускает обработку.
    Если обнаружится исключение, то выводит сообщение на экран'''

    delete_unused_fields()
    sequence = set_sequence()

    if dpg.get_value('prices') == '' or dpg.get_value('nom') == '' or sequence == [] or (len(row_widgets) == 1 and len(column_widgets) == 1):
        warning_message('Заполните все поля')
        return

    tmp_row_ranges = []
    tmp_row_sep = []

    tmp_column_ranges = []
    tmp_column_sep = []

    tmp_len = 0

    for i in range(len(row_widgets)):
        tmp_row_ranges.append(dpg.get_value(row_widgets[i][1]))
        tmp_row_sep.append(dpg.get_value(row_widgets[i][2]))

    for i in range(len(column_widgets)):
        tmp_column_ranges.append(dpg.get_value(column_widgets[i][1]))
        tmp_column_sep.append(dpg.get_value(column_widgets[i][2]))

    tmp_row_ranges.pop(0)
    tmp_row_sep.pop(0)
    tmp_column_ranges.pop(0)
    tmp_column_sep.pop(0)

    tmp_len = excel_regex.create_filters(tmp_row_ranges, tmp_row_sep, tmp_column_ranges, tmp_column_sep)
    
    if len(sequence) < tmp_len:
        warning_message('Переберите все строки/столбцы')
        return

    for i in sequence:
        if int(i) > tmp_len:
            warning_message()
            return

    # print(excel_regex.__dict__)
    excel_regex.execute()
    return

def set_sequence() -> list[str]:
    '''Устанавливает длину последовательности строки (regex) исходя из полученных данных в книге Excel'''

    tmp = dpg.get_value('sequence')
    tmp = excel_regex.set_sequence(tmp)
    return tmp

def check_sequence() -> None:
    '''Удаление символов кроме цифр из последовательности (человеческий фактор)'''

    widget = 'sequence'
    tmp = list(dpg.get_value(widget))

    for i in range(len(tmp)):
        if tmp[i] not in '0123456789 ':
            tmp[i] = ' '
    
    tmp = ''.join(tmp)
    dpg.set_value(widget, tmp)

def update_source(sender) -> None:
    '''Передача в объект класса диапазона ячеек, название листа и путь к файлу, откуда брать данные'''

    tmp = True
    tmp_names = ['prices', 'prices_sheet', 'prices_file']
    
    if sender == 'nom' or sender == 'nom_button':
        tmp = False
        tmp_names = ['nom', 'nom_sheet', 'nom_file']

    tmp_data = excel_regex.source_file_and_cells(tmp)

    dpg.set_value(tmp_names[0], tmp_data[0])
    dpg.set_value(tmp_names[1], tmp_data[1])
    dpg.set_value(tmp_names[2], tmp_data[2])

def add_row_widget() -> None:
    '''Добавление поля ввода строки

    p.s.
    Я хотел сократить код и уместить все в одну функцию, но у меня возникла проблема с получением исходных таблиц, поэтому снял проект с поддержки'''

    # print('before:',row_widgets)
    
    next_x, next_y = dpg.get_item_pos(row_widgets[-1][-3])[0], dpg.get_item_pos(row_widgets[-1][-3])[1]

    tmp = []

    tmp.append(dpg.add_button(label='-', parent=group_row, width=button_width, pos=[next_x, next_y+30], callback=lambda id: delete_widget(id, row_widgets)))
    tmp.append(dpg.add_input_text(parent=group_row, width=x*2+15, uppercase=True, hint=hint, callback=lambda id: set_value(id, row_widgets), on_enter=True, pos=[next_x*4, next_y+30]))
    tmp.append(dpg.add_input_text(parent=group_row, width=button_width, uppercase=True, pos=[next_x*16, next_y+30]))

    row_widgets.append(tmp)

    tmp = ''

def add_column_widget() -> None:
    '''Добавление поля ввода столбцов'''

    next_x, next_y = dpg.get_item_pos(column_widgets[-1][-3])[0], dpg.get_item_pos(column_widgets[-1][-3])[1]

    tmp = []

    tmp.append(dpg.add_button(label='-', parent=group_column, width=button_width, pos=[next_x, next_y+30], callback=lambda id: delete_widget(id, column_widgets)))
    tmp.append(dpg.add_input_text(parent=group_column, width=x*2+15, uppercase=True, hint=hint, callback=lambda id: set_value(id, column_widgets), on_enter=True, pos=[next_x+button_width+5, next_y+30]))
    tmp.append(dpg.add_input_text(parent=group_column, width=button_width, uppercase=True, pos=[next_x+button_width*6, next_y+30]))

    column_widgets.append(tmp)

    tmp = ''

def delete_widget(sender, widget_list) -> None:
    '''Удаление полей ввода'''

    x = get_index(sender, widget_list)
    # print(x)
    for i in range(3):
        dpg.delete_item(widget_list[x[0]][i])
    widget_list.pop(x[0])

    update_pos()

def update_pos() -> None:
    '''Обновление стека полей на экране'''

    i = 1
    while i < len(row_widgets):

        next_x, next_y = dpg.get_item_pos(row_widgets[i-1][-3])[0], dpg.get_item_pos(row_widgets[i-1][-3])[1]

        dpg.set_item_pos(row_widgets[i][0], [next_x, next_y+30])
        dpg.set_item_pos(row_widgets[i][1], [next_x*4, next_y+30])
        dpg.set_item_pos(row_widgets[i][2], [next_x*16, next_y+30])

        # print(row_widgets)
        i+=1
    
    i = 1
    while i < len(column_widgets):

        next_x, next_y = dpg.get_item_pos(column_widgets[i-1][-3])[0], dpg.get_item_pos(column_widgets[i-1][-3])[1]

        dpg.set_item_pos(column_widgets[i][0], [next_x, next_y+30])
        dpg.set_item_pos(column_widgets[i][1], [next_x+button_width+5, next_y+30])
        dpg.set_item_pos(column_widgets[i][2], [next_x+button_width*6, next_y+30])

        # print(column_widgets)
        i+=1

def get_index(sender, widget_list) -> list:
    '''Получение id виджета на экране'''

    x = [x for x in widget_list if sender in x][0]
    # print(x)
    x = [widget_list.index(x), x.index(sender)]
    return x

def set_value(sender, widget_list) -> None:
    '''Получение значений по введенным диапазонам ячеек'''

    x = get_index(sender, widget_list)
    tmp = excel_regex.get_address()
    dpg.set_value(widget_list[x[0]][1], tmp)

def delete_unused_fields() -> str:
    '''Удаляет все неиспользуемые поля ввода строк/столбцов, а также разделители'''

    tmp_widget_list = row_widgets
    i = 0
    switched = False

    while True:
        widget = tmp_widget_list[i][1]
        
        if widget != '' and dpg.get_value(widget) == '':
            delete_widget(widget, tmp_widget_list)
            i=0
        i+=1

        if i >= len(row_widgets) and not switched:
            i = 0
            tmp_widget_list = column_widgets
            switched = True
        elif i >= len(column_widgets) and switched:
            break
        
    return 'All unused fields deleted!'

def warning_message(message: str='Превышение последовательности') -> None:
    '''Сообщение об ошибке пользователю'''

    with dpg.window(label='Warning!', width=250, height=100, no_move=no_move, no_resize=not resizable, modal=True, no_close=True, tag='warning_window', pos=[100, 100]):
        dpg.add_text(message, color=(255, 255,255))
        dpg.add_button(label='OK', width=250, callback=lambda:dpg.delete_item('warning_window'))

def theme_change() -> None:
    '''Изменение темы приложения на темную или светлую'''

    global window_color, font_color
    window_color, font_color = font_color, window_color
    items = dpg.get_all_items()
    for i in items:
        try:
            dpg.configure_item(i, color=font_color)
        except:
            pass
    dpg.configure_item('theme_changer', default_value=font_color)
    dpg.configure_viewport('Excel regex',clear_color = window_color)
    
    print(tmp)

def always_top() -> None:
    '''Переключение режима окна: поверх других окон или нет'''

    if dpg.is_viewport_always_top():
        always_on_top = False
    elif not dpg.is_viewport_always_top():
        always_on_top = True
    dpg.set_viewport_always_top(always_on_top)


if __name__ == '__main__':

    dpg.create_context()

    with dpg.font_registry():
        '''Для отображения кириллицы, это скорее костыль'''

        with dpg.font(f'C:\\Windows\\Fonts\\Calibri.ttf', 15, default_font=True, id="Default font"):
            dpg.add_font_range_hint(dpg.mvFontRangeHint_Cyrillic)

    dpg.bind_font("Default font")

    with dpg.window(label='Excel', tag='main_windows', width=width, height=height, no_title_bar=no_title_bar, no_move=no_move, no_background=no_background, no_resize=not resizable):
        '''Создание окна с группами виджетов'''

        hint = 'Выделите ячейки в Excel файле'
        horizontal = True
        spacing = 5


        with dpg.group(horizontal=horizontal, horizontal_spacing=spacing) as group_param:

            dpg.add_text('На переднем плане', color=font_color)
            dpg.add_checkbox(callback=always_top, default_value=True)
            dpg.add_text('Смена цвета', color=font_color)
            dpg.add_color_button(tag='theme_changer', callback=theme_change,  no_border = True)


        with dpg.group(horizontal=horizontal, horizontal_spacing=spacing) as group_prices:

            dpg.add_text('Цены', color=font_color)
            dpg.add_input_text(tag='prices', width=x*2+15, uppercase=True, hint=hint, callback=update_source, on_enter=True)
            dpg.add_button(label='-', tag='prices_button', width=button_width, callback=update_source)
            dpg.add_text('Лист цен;', tag='prices_sheet', color=font_color)
            dpg.add_text('Файл цен', tag='prices_file', color=font_color)


        with dpg.group(horizontal=horizontal, horizontal_spacing=spacing) as group_nom:

            dpg.add_text('Ном. ', color=font_color)
            dpg.add_input_text(tag='nom', width=x*2+15, uppercase=True, hint=hint, callback=update_source, on_enter=True)
            dpg.add_button(label='-', tag='nom_button', width=button_width, callback=update_source)
            dpg.add_text('Лист ном;', tag='nom_sheet', color=font_color)
            dpg.add_text('Файл номенклатуры', tag='nom_file', color=font_color)


        with dpg.group(horizontal=horizontal, horizontal_spacing=spacing) as group_sequence:

            dpg.add_text('Последовательность', color=font_color)
            dpg.add_input_text(tag='sequence', width=x*4, callback=check_sequence)
            dpg.add_button(label='+', tag='sequence_button', width=button_width, callback=find_matches)


        with dpg.group(horizontal=horizontal, horizontal_spacing=x*5)     as group_row_column:

            dpg.add_text('Строки', color=font_color)
            dpg.add_text('Столбцы', color=font_color)


        with dpg.group(pos=[x/5, y*4])                                    as group_row:

            tmp = []
            tmp.append(dpg.add_button(label='+', width=button_width*4.5, callback=add_row_widget))
            tmp.append('')
            tmp.append('')
            row_widgets.append(tmp)
            tmp=''


        with dpg.group(pos=[x*6, y*4])                                    as group_column:

            tmp = []
            tmp.append(dpg.add_button(label='+', width=button_width*4.5, callback=add_column_widget))
            tmp.append('')
            tmp.append('')
            column_widgets.append(tmp)
            tmp=''

    dpg.create_viewport(title='Excel regex',
                        width=width,
                        height=height,
                        always_on_top=always_on_top,
                        clear_color=window_color,
                        resizable=resizable,
                        small_icon=small_icon)

    dpg.setup_dearpygui()
    dpg.show_viewport()
    while dpg.is_dearpygui_running():
        # put old render callback code here!
        dpg.render_dearpygui_frame()  
    dpg.destroy_context()