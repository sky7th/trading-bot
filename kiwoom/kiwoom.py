from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from config.errorCode import *


class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()

        self.__login_event_loop = None

        self.__account_num = None

        self.__get_ocx_instance()
        self.__event_slots()

        self.signal_login_commConnect()
        self.get_acount_info()

    def __get_ocx_instance(self):
        self.setControl('KHOPENAPI.KHOpenAPICtrl.1')

    def __event_slots(self):
        self.OnEventConnect.connect(self.__login_slot)

    def __login_slot(self, errCode):
        print(errors(errCode))
        self.__login_event_loop.exit()

    def signal_login_commConnect(self):
        self.dynamicCall('CommConnect()')
        self.__login_event_loop = QEventLoop()
        self.__login_event_loop.exec_()

    def get_acount_info(self):
        account_list = self.dynamicCall("GetLoginInfo(String)", "ACCNO")
        self.__account_num = account_list.split(":")[0]
        print("나의 보유 계좌번호 %s" % self.__account_num)
