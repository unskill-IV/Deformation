from openpyxl import Workbook
from modules.deformation_module import (
    load_data,
    get_height_and_deformation,
    get_values,
    mean_value,
    mean_abs_value,
    error_finder,
    error_correction,
    create_print_skeleton,
    velocity_f,
    poisson_f,
    anisotropy_f
)

file = input('Enter measurement name: ')
data_books = load_data(file)
work_book = Workbook()
height, deformation = get_height_and_deformation(file, len(data_books[0].sheetnames))

print_names = ['time', 'velocity', 'Poisson', 'delta', 'shift', 'UV|TV']
print_values = []

for i in range(3):
    shts = data_books[i].sheetnames
    ti_print_values = []

    for sht in shts:
        tn, t2 = get_values(data_books[i], sht)
        total_amount_of_values = len(tn)

        if total_amount_of_values >= 3:
            ta = list(map(lambda x: x * 2, t2))
            shift_constant = error_finder(tn, t2)
            delta_t = []

            new_tn = [0.0] * total_amount_of_values
            new_t2 = [0.0] * total_amount_of_values
            new_ta = [0.0] * total_amount_of_values

            for j in range(total_amount_of_values):
                new_tn[j], new_t2[j] = error_correction(tn[j], t2[j], shift_constant)
                new_ta[j] = new_t2[j] * 2
                delta_t.append(round(new_tn[j] - new_ta[j], 4))

            while max(list(map(lambda x: abs(x), delta_t))) > 0.1:
                for j in range(len(delta_t)):
                    if abs(delta_t[j]) > 0.1:
                        new_tn.pop(j)
                        new_t2.pop(j)
                        new_ta.pop(j)
                        delta_t.pop(j)
                        break

            mt = list(map(lambda x: round(x / 2, 4), new_tn))
            mean_mt = mean_value(mt)
            velocity = velocity_f(mean_mt, height[shts.index(sht)])
            mean_delta_t = mean_abs_value(delta_t)
            used_amount_of_values = len(new_tn)
        else:
            mean_mt = 0.0
            velocity = 0.0
            mean_delta_t = 0.0
            shift_constant = 0.0
            used_amount_of_values = 0

        sheet_print_values = [
            mean_mt,
            velocity,
            '-',
            mean_delta_t,
            shift_constant,
            f'{used_amount_of_values}|{total_amount_of_values}'
        ]

        ti_print_values.append(sheet_print_values)

    print_values.append(ti_print_values)

create_print_skeleton(work_book, print_names, data_books[0].sheetnames)

rows_for_print = len(print_values[0][0])
columns_for_print = len(print_values[0])

for t in range(2):
    for sh in range(columns_for_print):
        if print_values[t][sh][print_names.index('UV|TV')][0] != '0':
            ti = print_values[t][sh][print_names.index('time')]
            t3 = print_values[2][sh][print_names.index('time')]

            poisson = poisson_f(ti, t3)
            print_values[t][sh][print_names.index('Poisson')] = poisson
        else:
            print_values[t][sh][print_names.index('Poisson')] = 0.0

for sh in range(columns_for_print):
    if print_values[0][sh][print_names.index('UV|TV')][0] != '0':
        if print_values[1][sh][print_names.index('UV|TV')][0] != '0':
            t1 = print_values[0][sh][print_names.index('time')]
            t2 = print_values[1][sh][print_names.index('time')]

            anisotropy = anisotropy_f(t1, t2)
        else:
            anisotropy = 0.0
    else:
        anisotropy = 0.0

    work_book['Sheet'].cell(row=((3 * rows_for_print) + 2), column=(sh + 3), value=anisotropy)

for t in range(3):
    for i in range(columns_for_print):
        for j in range(rows_for_print):
            work_book['Sheet'].cell(row=(t * rows_for_print + j + 2), column=(i + 3), value=print_values[t][i][j])

work_book.save(f'results/{file}/{file}.xlsx')
