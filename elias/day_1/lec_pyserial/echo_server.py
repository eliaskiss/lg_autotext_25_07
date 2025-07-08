import sys
from icecream import ic

# ic(sys.path)
# sys.path.append('../lec_pyserial')
# ic(sys.path)

# or
# copy lec_pyserial_class --> .venv/lib/python3.10/site-packages

from lec_pyserial_class import Serial

ic.configureOutput(includeContext=True)

# Echo Server
# 입력받다가 Enter키가 입력되는 순간 그동안의 데이터를 출력하는 에코서버
# Exit, exit, EXIT, eXit, EXit lower() 들어오면 프로그램 종료`
# b'\x0d'

def main():
    ser = Serial()
    ser.openSerial('com1')

    ic('Echo Server is running...')

    RETURN_CODE = b'\x0d'  # Carriage Return CR (13:0x0D)