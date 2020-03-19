from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from config.errorCode import *


class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()

        self.__login_event_loop = None
        self.__detail_account_info_event_loop = None

        self.__account_num = None

        self.__get_ocx_instance()
        self.__event_slots()

        self.signal_login_commConnect()
        self.get_acount_info()
        self.detail_account_info()

    def __get_ocx_instance(self):
        self.setControl('KHOPENAPI.KHOpenAPICtrl.1')

    def __event_slots(self):
        self.OnEventConnect.connect(self.__login_slot)
        self.OnReceiveTrData.connect(self.__trdata_slot)

    def __login_slot(self, errCode):
        print(errors(errCode))
        self.__login_event_loop.exit()

    def __trdata_slot(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext):
        '''
        tr요청을 받는 구역, 슬롯
        :param sScrNo: 스크린번호
        :param sRQName: 내가 요청했을 때 지은 이름
        :param sTrCode: 요청id, tr코드
        :param sRecordName: 사용 안함
        :param sPrevNext: 다음 페이지가 있는지
        :return:
        '''
        if sRQName == "예수금상세현황요청":
            deposit = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "예수금")
            ok_deposit = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "출금가능금액")
            print("예수금: %s" % int(deposit))
            print("출금가능금액: %s" % int(ok_deposit))

        self.__detail_account_info_event_loop.exit()

    def signal_login_commConnect(self):
        self.dynamicCall('CommConnect()')
        self.__login_event_loop = QEventLoop()
        self.__login_event_loop.exec_()

    def get_acount_info(self):
        account_list = self.dynamicCall("GetLoginInfo(String)", "ACCNO")
        self.__account_num = account_list.split(":")[0][:-1]
        print("내 계좌번호: %s" % self.__account_num)

    def detail_account_info(self):
        self.dynamicCall("SetInputValue(String, String)", "계좌번호", self.__account_num)
        self.dynamicCall("SetInputValue(String, String)", "비밀번호", "0000")
        self.dynamicCall("SetInputValue(String, String)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(String, String)", "조회구분", "2")
        self.dynamicCall("CommRqData(String, String, int, String)", "예수금상세현황요청", "opw00001", "0", "2000")
        self.__detail_account_info_event_loop = QEventLoop()
        self.__detail_account_info_event_loop.exec_()
