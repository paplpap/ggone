from pyhwpx import Hwp

from src.KDS_3p_12 import KDS_312
from src.K_4p_15_1 import K415_1
from src.K_4p_15_2 import K415_2, K415_3
from src.KDS_4p_16_1 import K416_1
from src.K_4p_17_1 import K417_1
from src.KDS_4p_17_2 import K417_2
from src.KDS_4p_18_1 import K418
from src.K_4p_19 import K419
from src.K_4p_20 import K420
from src.K_4p_23 import K423




def KDScode(file_path):
    hwp = Hwp()
    hwp.open(file_path)

    KDS_312(hwp)
    K415_1(hwp)
    K415_2(hwp)
    K415_3(hwp)
    K416_1(hwp)
    K417_1(hwp)
    K417_2(hwp)
    K418(hwp)
    K419(hwp)
    K420(hwp)
    K423(hwp)

if __name__ == '__main__':
    KDScode(r'../data/KDS 11 50 25 기초 내진 설계기준(21.05).hwp')
