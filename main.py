import os
from openpyxl import Workbook

file_name = 'example.txt'  # название файла

crane_type = 1
# 1 - мостовой
# 2 - башенный
# 3 - козловой

BRIDGE_S = 0
BRIDGE_L = 0

# пусть пользователь введет crane_type
crane_type = int(input('Введите тип крана (1 - мостовой, 2 - башенный, 3 - козловой): '))

BRIDGE_S = float(input('Введите S: '))
BRIDGE_L = float(input('Введите L: '))
BRIDGE_H = float(input(
    'Введите H: '))  # высота рельсины, первая и последняя точка поднимаются на H, от этого создается еще одна эталонная прямая, и от нее будут считаться отклонения по z для остальных точек

crane_type_names = {
    1: 'мостовой',
    2: 'башенный',
    3: 'козловой',
}

workbook = Workbook()
sheet = workbook.active
sheet["A1"] = "тип крана"
sheet["A2"] = "s"
sheet["A3"] = "l"
sheet["B1"] = crane_type_names[crane_type]
sheet["B2"] = BRIDGE_S
sheet["B3"] = BRIDGE_L

# ОТКЛОНЕНИЕ 1
# мостовые
# 0,002 S, но не более 40
BRIDGE_CRANE_1_MULTIPLIER = 0.002
BRIDGE_CRANE_1_NOT_BIGGER_THAN = 40

# башенные (45-60)
TOWER_CRANE_1 = 45  # TODO 45-60

# козловые (40)
GANTRY_CRANE_1 = 40

# ОТКЛОНЕНИЕ 2
# мостовой
# 0,0015 L, но не более 10 мм.
BRIDGE_CRANE_2_MULTIPLIER = 0.0015
BRIDGE_CRANE_2_NOT_BIGGER_THAN = 0.001  # ЭТО СКАЗАЛА АЙГУЛЬ

# для башенного и козлового прочерк

# ОТКЛОНЕНИЕ 3

# мостовой
# 0,002S, но не более 15
BRIDGE_CRANE_3_MULTIPLIER = 0.002
BRIDGE_CRANE_3_NOT_BIGGER_THAN = 15

TOWER_CRANE_3 = 10

GANTRY_CRANE_3 = 15

dots = []  # массив для хранения точек (x, y, z)

with open(file_name, 'r') as f:
    a = f.readlines()
    # remove all lines from a that start with 'ST' or with 'PR'
    a = [x for x in a if (not x.startswith('ST')) and (not x.startswith('PR'))]
    a.sort(key=lambda x: int(x.split(',')[0]))
    for i in a:
        if not i:  # если строка пустая, то пропускаем её
            continue
        dots.append([float(j) for j in i.split(',')[1:4]])

print(dots)

dots_on_rail = int(len(dots) / 2)  # кол во точке на рельсе (вместе с начальной и конечной)

# первые 4 точки это:
# 1 - начальная первой рельсы
# 2 - конечная первой рельсы
# 3 - начальная второй рельсы
# 4 - конечная второй рельсы

d1 = dots[0]
d2 = dots[1]
d3 = dots[2]
d4 = dots[3]
dots_first = dots[4:2 + dots_on_rail]  # точки первой рельсы
dots_second = dots[2 + dots_on_rail:]  # точки второй рельсы

lendots = len(dots_first)  # кол-во точек в одном массиве


def getKBbyTwoDots(x1, y1, x2, y2):
    if x2 != x1:
        k = (y2 - y1) / (x2 - x1)  # k = (y1 - y0) / (x1 - x0)
    else:
        k = 0
    b = (y2 - k * x2)  # b = y0 - a * x0
    return k, b


def getABCbyKB(k, b):
    return -k, 1, -b


# y = kx + b
# -kx + y - b = 0 (в форме Ax + By + C = 0) => A = -k; B = 1; C = -b

first_line_K, first_line_B = getKBbyTwoDots(d1[0], d2[0], d1[1], d2[1])  # прямая в системе (x, y)

line1_A, line1_B, line1_C = getABCbyKB(first_line_K, first_line_B)

second_line_K, second_line_B = getKBbyTwoDots(d3[0], d3[1], d4[0], d4[1])

line2_A, line2_B, line2_C = getABCbyKB(second_line_K, second_line_B)

flXZ_K, flXZ_B = getKBbyTwoDots(d1[0], d2[2], d1[0], d2[2])  # первая прямая в системе (x, z)
slXZ_K, slXZ_B = getKBbyTwoDots(d3[0], d4[2], d3[0], d4[2])  # вторая прямая в системе (x, z)

heightened_flXZ_B = flXZ_B + BRIDGE_H  # поднимаем прямую на H
heightened_slXZ_B = slXZ_B + BRIDGE_H  # поднимаем прямую на H

print(line1_A, line1_B, line1_C)
print(line2_A, line2_B, line2_C)
# 22.386 + 0.145
# 22.522 - 22.377


def first_H_line_func(x):  # функция первой прямой, поднятной на H
    return flXZ_K * x + heightened_flXZ_B


def second_H_line_func(x):  # функция второй прямой, поднятной на H
    return slXZ_K * x + heightened_slXZ_B


def point_to_line_distance(A, B, C, x0, y0):  # расстояние от точки до прямой (прямая в форме Ax + By + C = 0)
    return abs(A * x0 + B * y0 + C) / ((A ** 2 + B ** 2) ** 0.5)


def two_point_distance(x0, y0, x1, y1):  # расстояние между двумя точками
    return ((x1 - x0) ** 2 + (y1 - y0) ** 2) ** 0.5


'''
Отклонение 1 - Разность отметок головок рельсов в одном поперечном сечении
(отклонение по Z у противоположных рельс)
z1-z0 не более чем допустимое отклонение

Отклонение 2 - Разность отметок рельсов на соседних колоннах
(отклонение по Z у соседних рельс)
z1-z0 не более чем допустимое отклонение

Отклонение 3 - Сужение или расширение колеи рельсового пути
|S - (y1-y0)| не более чем допустимое отклонение

Отклонение 4 - Разность высотных отметок головок рельсов
взять эталонную прямую в системе (x, z), сместить ее на h вверх (y(x)=kx+b+h)
для каждой нижней точки подставлять x0 в функцию, c=y(x0), и находить разность между y0 и c
'''

next_empty_line = 5  # used globally


def xwrite(row, line, text, incrementNextLine=False):  # to write text to cell
    if incrementNextLine:
        global next_empty_line  # uses global variable
        next_empty_line += 1
    sheet[f"{row}{line}"] = text


for i in range(lendots):
    line1_index = 4 + i  # индекс первой точки в массиве dots
    line2_index = 4 + i + lendots  # индекс второй точки в массиве dots
    next_empty_line += 2

    if line1_index + 1 == 14:
        print('HERE')

    xwrite('A', next_empty_line, f'{line1_index + 1}-{line2_index + 1}', True)

    # ПРОВЕРКА ОТКЛОНЕНИЯ 1
    if crane_type == 1 and (abs(dots_first[i][2] - dots_second[i][2]) > BRIDGE_CRANE_1_MULTIPLIER * BRIDGE_S or abs(
            dots_first[i][2] - dots_second[i][
                2]) > BRIDGE_CRANE_1_NOT_BIGGER_THAN):  # если разность по z больше чем 0.002S и не больше 40 (ОТКЛОНЕНИЕ 1)
        print(
            f'ОТКЛОНЕНИЕ 1 НА {line1_index + 1}-{line2_index + 1}: {abs(dots_first[i][2] - dots_second[i][2])}; допустимое отклонение {BRIDGE_CRANE_1_MULTIPLIER * BRIDGE_S} и не больше чем {BRIDGE_CRANE_1_NOT_BIGGER_THAN}')
        xwrite('A', next_empty_line, '1. Разность отметок головок рельсов в одном поперечном сечении (отклонение по Z)',
               True)
        xwrite('A', next_empty_line, 'Отклонение')
        xwrite('B', next_empty_line, 'Допустимое отклонение', True)
        xwrite('A', next_empty_line, f'{abs(dots_first[i][2] - dots_second[i][2])}')  # TODO выше или ниже
        xwrite('B', next_empty_line, f'{BRIDGE_CRANE_1_MULTIPLIER * BRIDGE_S}', True)
        next_empty_line += 1
    elif crane_type == 2 and abs(dots_first[i][2] - dots_second[i][
        2]) > TOWER_CRANE_1:  # если разность отметок головок рельсов в одном поперечном сечении больше, чем 45, то это ошибка
        print(
            f'ОТКЛОНЕНИЕ 1 НА {line1_index + 1}-{line2_index + 1}: {abs(dots_first[i][2] - dots_second[i][2])}; допустимое отклонение {TOWER_CRANE_1}')  # если разность отметок головок рельсов в одном поперечном сечении больше, чем 45, то это ошибка
        xwrite('A', next_empty_line, '1. Разность отметок головок рельсов в одном поперечном сечении (отклонение по Z)',
               True)
        xwrite('A', next_empty_line, 'Отклонение')
        xwrite('B', next_empty_line, 'Допустимое отклонение', True)
        xwrite('A', next_empty_line, f'{abs(dots_first[i][2] - dots_second[i][2])}')  # TODO выше или ниже
        xwrite('B', next_empty_line, f'{TOWER_CRANE_1}', True)
    elif crane_type == 3 and abs(dots_first[i][2] - dots_second[i][
        2]) > GANTRY_CRANE_1:  # если разность отметок головок рельсов в одном поперечном сечении больше, чем 40, то это ошибка
        print(
            f'ОТКЛОНЕНИЕ 1 НА {line1_index + 1}-{line2_index + 1}: {abs(dots_first[i][2] - dots_second[i][2])}; допустимое отклонение {GANTRY_CRANE_1}')  # если разность отметок головок рельсов в одном поперечном сечении больше, чем 40, то это ошибка
        xwrite('A', next_empty_line, '1. Разность отметок головок рельсов в одном поперечном сечении (отклонение по Z)',
               True)
        xwrite('A', next_empty_line, 'Отклонение')
        xwrite('B', next_empty_line, 'Допустимое отклонение', True)
        xwrite('A', next_empty_line, f'{abs(dots_first[i][2] - dots_second[i][2])}')  # TODO выше или ниже
        xwrite('B', next_empty_line, f'{GANTRY_CRANE_1}', True)
    next_empty_line += 1

    # ПРОВЕРКА ОТКЛОНЕНИЯ 3
    dist1 = point_to_line_distance(line2_A, line2_B, line2_C, dots_first[i][0], dots_first[i][1])
    dist2 = point_to_line_distance(line2_A, line2_B, line2_C, dots_second[i][0], dots_second[i][1])
    dist = 0

    # d3, d4, dots_second[i]
    D = (dots_second[i][0] - d3[0]) * (d4[1] - d3[1]) - (dots_second[i][1] - d3[1]) * (d4[0] - d3[0])
    if D < 0:
        dist = dist1 + dist2
    elif D > 0:
        dist = dist1 - dist2
    else:
        dist = dist1

    xwrite('A', next_empty_line, 'СУЖЕНИЕ И РАСШИРЕНИЕ', True)
    xwrite('A', next_empty_line, f'{line1_index + 1}-{line2_index + 1}')
    xwrite('B', next_empty_line, f'{dist}', True)

    next_empty_line += 1
    print(
        f'РАССТОЯНИЕ МЕЖДУ {line1_index + 1}-{line2_index + 1}\': {two_point_distance(dots_first[i][0], dots_first[i][1], dots_first[i][0], dots_second[i][1])}')  # измеряем расстояние между точкой на верхней рельсе, и сдвинутой точкой на нижней релье, поэтому координата x1 = x0 в этом случае, т.к. они параллельны
    print(
        f'РАССТОЯНИЕ ОТ ТОЧКИ {line1_index + 1} ДО ЭТАЛОННОЙ ПРЯМОЙ 1-2: {point_to_line_distance(line1_A, line1_B, line1_C, dots_first[i][0], dots_first[i][1])}')
    print(
        f'РАССТОЯНИЕ ОТ ТОЧКИ {line2_index + 1} ДО ЭТАЛОННОЙ ПРЯМОЙ 3-4: {point_to_line_distance(line2_A, line2_B, line2_C, dots_second[i][0], dots_second[i][1])}')
    xwrite('A', next_empty_line, f'РАССТОЯНИЕ МЕЖДУ {line1_index + 1}-{line2_index + 1}\'')
    xwrite('B', next_empty_line,
           f'{two_point_distance(dots_first[i][0], dots_first[i][1], dots_first[i][0], dots_second[i][1])}', True)
    xwrite('A', next_empty_line, f'РАССТОЯНИЕ ОТ ТОЧКИ {line1_index + 1} ДО ЭТАЛОННОЙ ПРЯМОЙ 1-2')
    xwrite('B', next_empty_line,
           f'{point_to_line_distance(line1_A, line1_B, line1_C, dots_first[i][0], dots_first[i][1])}', True)
    xwrite('A', next_empty_line, f'РАССТОЯНИЕ ОТ ТОЧКИ {line2_index + 1} ДО ЭТАЛОННОЙ ПРЯМОЙ 3-4')
    xwrite('B', next_empty_line,
           f'{point_to_line_distance(line2_A, line2_B, line2_C, dots_second[i][0], dots_second[i][1])}', True)

next_empty_line += 1
xwrite('A', next_empty_line, 'Отклонения по Z между точками и эталонной прямой, поднятой на H', True)

for i in range(4, 4 + lendots):
    xwrite('A', next_empty_line, f'{i + 1}')
    # calculate using H line function
    xwrite('B', next_empty_line,
           f'{first_H_line_func(dots[i][2]) - abs(dots[i][2])}', True)
for i in range(4 + lendots, 4 + lendots * 2):
    xwrite('A', next_empty_line, f'{i + 1}')
    # calculate using H line function
    xwrite('B', next_empty_line,
           f'{second_H_line_func(dots[i][2]) - abs(dots[i][2])}', True)

next_empty_line += 1
xwrite('A', next_empty_line, 'Разность отметок рельсов на соседних колоннах (Расстояния по z между соседними точками)', True)

for i in range(4, 4 + lendots - 1):
    xwrite('A', next_empty_line, f'{i + 1}-{i}')
    xwrite('B', next_empty_line, f'{abs(dots[i][2] - dots[i + 1][2])}')
    xwrite('C', next_empty_line, f'{i+1} {"выше" if dots[i][2] - dots[i + 1][2] < 0 else "ниже"} чем {i}')
    xwrite('D', next_empty_line, f'Допустимое отклонение: {BRIDGE_CRANE_2_MULTIPLIER * BRIDGE_L}', True)
for i in range(4 + lendots, 4 + lendots * 2 - 1):
    xwrite('A', next_empty_line, f'{i + 1}-{i}')
    xwrite('B', next_empty_line, f'{abs(dots[i][2] - dots[i + 1][2])}')
    xwrite('C', next_empty_line, f'{i + 1} {"выше" if dots[i][2] - dots[i + 1][2] < 0 else "ниже"} чем {i}')
    xwrite('D', next_empty_line, f'Допустимое отклонение: {BRIDGE_CRANE_2_MULTIPLIER * BRIDGE_L}', True)

# if os.path.isfile(file_name + '.xlsx'):
#     os.remove(file_name + '.xlsx')
workbook.save(filename=file_name + '.xlsx')

# https://www.geogebra.org/3d/wk5z5rep

# поднять 1 и 2 точку на h, и оттуда считать отклонение по z
# соседние точки по z найти разницу


'''
точка слева или справа от прямой
Определяется так. Предположим, у нас есть 3 точки: А(х1,у1), Б(х2,у2), С(х3,у3). Через точки А и Б проведена прямая. И нам надо определить, как расположена точка С относительно прямой АБ. Для этого вычисляем значение:
D = (х3 - х1) * (у2 - у1) - (у3 - у1) * (х2 - х1)
- Если D = 0 - значит, точка С лежит на прямой АБ.
- Если D < 0 - значит, точка С лежит слева от прямой.
- Если D > 0 - значит, точка С лежит справа от прямой.
'''