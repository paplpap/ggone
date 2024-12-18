from pyhwpx import Hwp
from src.KDS_3p_12 import K312
from src.K_4p_15_1 import K415_1
from src.K_4p_15_2 import K415_2, K415_3
from src.KDS_4p_16_1 import K416_1

import time

def KDScode(file_path):
    hwp = Hwp()
    hwp.open(file_path)

    K312(hwp)
    #time.sleep(5)
    K415_1(hwp)
    #time.sleep(5)
    K415_2(hwp)
    #time.sleep(5)

    K415_3(hwp)

    #K416_1(hwp)




if __name__ == '__main__':
    KDScode(r'../data/KDS 11 50 25 기초 내진 설계기준(21.05).hwp')


