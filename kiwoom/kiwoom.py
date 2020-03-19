from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from config.errorCode import *


class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()

        ###### eventloop 모음
        self.__login_event_loop = None
        self.__detail_account_info_event_loop = QEventLoop()

        ###### 스크린번호 모음
        self.screen_my_info = "2000"

        ###### 변수 모음
        self.__account_num = None
        self.__account_stock_dict = {}

        ###### 계좌 관련 변수 모음
        self.__use_money = 0
        self.__use_money_percent = 0.5

        self.__get_ocx_instance()
        self.__event_slots()

        self.signal_login_commConnect()
        self.get_acount_info()
        self.detail_account_info()  # 예수금 요청
        self.detail_account_mystock()  # 계좌평가 잔고 내역 요청

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
            self.__use_money = int(deposit) * self.__use_money_percent
            self.__use_money = self.__use_money / 4
            self.__detail_account_info_event_loop.exit()

        if sRQName == "계좌평가잔고내역요청":
            total_buy_money = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "총매입금액")
            total_buy_money_result = int(total_buy_money)
            total_profit_loss_rate = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "총수익률(%)")
            total_profit_loss_rate_result = float(total_profit_loss_rate)

            row_size = self.dynamicCall("GetRepeatCnt(QString, QString", sTrCode, sRQName)

            for i in range(row_size):
                code = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목번호")
                code_nm = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목명")
                stock_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "보유수량")
                buy_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "매입가")
                learn_rate = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "수익률(%)")
                current_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "현재가")
                total_buy_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "매입금액")
                possible_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "매매가능수량")

                code = code.strip()[1:]
                if code not in self.__account_stock_dict:
                    self.__account_stock_dict.update({code: {}})

                self.__account_stock_dict[code].update({"종목명": code_nm.strip()})
                self.__account_stock_dict[code].update({"보유수량": int(stock_quantity.strip())})
                self.__account_stock_dict[code].update({"매입가": int(buy_price.strip())})
                self.__account_stock_dict[code].update({"수익률(%)": float(learn_rate.strip())})
                self.__account_stock_dict[code].update({"현재가": int(current_price.strip())})
                self.__account_stock_dict[code].update({"매입금액": int(total_buy_price.strip())})
                self.__account_stock_dict[code].update({"매매가능수량": int(possible_quantity.strip())})

            if sPrevNext == "2":
                self.detail_account_mystock(sPrevNext="2")
            else:
                print("총매입금액: %s" % total_buy_money_result)
                print("총수익률(%%): %s" % total_profit_loss_rate_result)
                print("계좌 내 종목 개수: %s" % len(self.__account_stock_dict))
                print("계좌 내 종목 내역: %s" % self.__account_stock_dict)
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
        self.dynamicCall("CommRqData(String, String, int, String)", "예수금상세현황요청", "opw00001", "0", self.screen_my_info)
        self.__detail_account_info_event_loop.exec_()

    def detail_account_mystock(self, sPrevNext="0"):
        self.dynamicCall("SetInputValue(String, String)", "계좌번호", self.__account_num)
        self.dynamicCall("SetInputValue(String, String)", "비밀번호", "0000")
        self.dynamicCall("SetInputValue(String, String)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(String, String)", "조회구분", "2")
        self.dynamicCall("CommRqData(String, String, int, String)", "계좌평가잔고내역요청", "opw00018", sPrevNext, self.screen_my_info)
        self.__detail_account_info_event_loop.exec_()
