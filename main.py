import pandas as pd
import matplotlib.pyplot as plt
from math import pi
import xlsxwriter


def calculate_kp(Den: float, CALI: float) -> float:
    CALI /= 10
    v = 10 * (CALI ** 2) * pi / 4
    v_porod = v * Den / 2.9
    v_por = v - v_porod
    Kp = v_por / v
    return Kp


def calculate_kp_neutron_log(GK: float, W: float) -> float:
    W /= 100
    Kp = (1.9 - GK * 0.15) * W
    return Kp


def first_customization(worksheet) -> None:
    worksheet.write(0, 1, "Пористость")
    worksheet.write(1, 1, "KP")
    worksheet.write(2, 1, "%")

    worksheet.write(0, 2, "Коллекторы")
    worksheet.write(0, 1, "KL")
    worksheet.write(2, 2, "boolean")

    worksheet.write(0, 0, "Высота")
    worksheet.write(1, 0, "MD")
    worksheet.write(2, 0, "m")


def main() -> None:
    data = pd.read_excel('data.xlsx', 0, header=5)
    with xlsxwriter.Workbook('result.xlsx') as workbook:
        Kp_ggk = []
        Kp_net = []
        worksheet = workbook.add_worksheet("result")
        first_customization(worksheet)

        for row, MD, CALI, Den, GK, W, PE in zip(range(3, 665), data['MD'][6:],
                                                 data['CALI'][6:],
                                                 data['Den'][6:], data['GK'][6:],
                                                 data['W'][6:], data['PE'][6:]):
            Kp = (calculate_kp_neutron_log(GK, W) + calculate_kp(Den, CALI)) / 2
            Kp_net.append(calculate_kp_neutron_log(GK, W))
            Kp_ggk.append(calculate_kp(Den, CALI))
            worksheet.write(row, 0, MD)
            worksheet.write(row, 1, Kp)
            if Kp > 0.072:
                worksheet.write(row, 2, 1)
            else:
                worksheet.write(row, 2, 0)
    plt.plot(Kp_ggk, data["MD"][6:])
    plt.show()
    plt.plot(Kp_net, data["MD"][6:])
    plt.show()
