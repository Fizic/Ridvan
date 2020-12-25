import pandas as pd
import matplotlib.pyplot as plt


def ggk_lp_to_Kp(Den, PE) -> float:
    pass


data = pd.read_excel('result.xlsx', 0, header=1)
print(data)
plt.title("Геофизический планшет")
plt.plot(data["KP"][3:], data["MD"][3:])
plt.show()
