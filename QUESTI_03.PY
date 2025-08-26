#import numpy as np
import matplotlib.pyplot as plt

fig, ax1 = plt.subplots(1,1)

x = [1, 2, 3, 4, 5]
y = [30, 15, 20, 10, 45]
rotulos = ['sc', 'pr', 'rs', 'sp', 'rj']

ax1.barh(x, y, tick_label=rotulos)
ax1.set_xlabel('valor')      # eixo horizontal
ax1.set_ylabel('estado')     # eixo vertical
plt.show()
