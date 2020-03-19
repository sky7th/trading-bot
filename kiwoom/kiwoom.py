from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from config.errorCode import *


class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()

        # 스크린번호
        self.SCREEN_MY_INFO = "2000"

        self.deposit = None
        self.deposit_ok = None
        self.total_buy_money = None
        self.total_profit_loss_rate = None
        self.mystock_dict = {}
        self.mystock_not_concluded_dict = {}

        # 계좌 관련 변수
        self.account_num = None
        self.use_money = 0
        self.USE_MONEY_PERCENT = 0.5

        # eventloop
        self.login_event_loop = QEventLoop()
        self.account_detail_event_loop = QEventLoop()

        self.__get_ocx_instance()
        self.__event_slots()

        self.signal_login_commConnect()

        self.set_acount_num()
        self.signal_account_detail_info()  # 예수금상세현황요청
        self.signal_account_detail_mystock()  # 계좌평가잔고내역요청
        self.signal_account_detail_mystock_not_concluded()  # 실시간미체결요청

    def __get_ocx_instance(self):
        self.setControl('KHOPENAPI.KHOpenAPICtrl.1')

    def __event_slots(self):
        self.OnEventConnect.connect(self.login_slot)
        self.OnReceiveTrData.connect(self.trdata_slot)

    def login_slot(self, errCode):
        print(errors(errCode))
        self.login_event_loop.exit()

    def trdata_slot(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext):
        """
        tr요청을 받는 구역, 슬롯
        :param sScrNo: 스크린번호
        :param sRQName: 내가 요청했을 때 지은 이름
        :param sTrCode: 요청id, tr코드
        :param sRecordName: 사용 안함
        :param sPrevNext: 다음 페이지가 있는지
        :return:
        """
        if sRQName == "예수금상세현황요청":
            self.set_deposit_details(sRQName, sTrCode)
            print("예수금: %s" % int(self.deposit))
            print("출금가능금액: %s" % int(self.deposit_ok))

        elif sRQName == "계좌평가잔고내역요청":
            self.set_mystock_details(sRQName, sTrCode, sPrevNext)
            print("총매입금액: %s" % int(self.total_buy_money))
            print("총수익률(%%): %s" % float(self.total_profit_loss_rate))
            print("계좌 내 종목 개수: %s" % len(self.mystock_dict))
            print("계좌 내 종목 내역: %s" % self.mystock_dict)

        elif sRQName == "실시간미체결요청":
            self.set_mystock_not_concluded_details(sRQName, sTrCode)
            print("계좌 내 미체결 종목 개수: %s" % len(self.mystock_not_concluded_dict))
            print("계좌 내 미체결 종목 내역: %s" % self.mystock_not_concluded_dict)

        self.account_detail_event_loop.exit()

    def signal_login_commConnect(self):
        self.dynamicCall('CommConnect()')
        self.login_event_loop.exec_()

    def signal_account_detail_info(self):
        self.dynamicCall("SetInputValue(String, String)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(String, String)", "비밀번호", "0000")
        self.dynamicCall("SetInputValue(String, String)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(String, String)", "조회구분", "2")
        self.dynamicCall("CommRqData(String, String, int, String)", "예수금상세현황요청", "opw00001", "0", self.SCREEN_MY_INFO)
        self.account_detail_event_loop.exec_()

    def signal_account_detail_mystock(self, sPrevNext="0"):
        self.dynamicCall("SetInputValue(String, String)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(String, String)", "비밀번호", "0000")
        self.dynamicCall("SetInputValue(String, String)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(String, String)", "조회구분", "2")
        self.dynamicCall("CommRqData(String, String, int, String)", "계좌평가잔고내역요청", "opw00018", sPrevNext,
                         self.SCREEN_MY_INFO)
        self.account_detail_event_loop.exec_()

    def signal_account_detail_mystock_not_concluded(self, sPrevNext="0"):
        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(QString, QString)", "체결구분", "1")
        self.dynamicCall("SetInputValue(QString, QString)", "매매구분", "0")
        self.dynamicCall("CommRqData(String, String, int, String)", "실시간미체결요청", "opt10075", sPrevNext,
                         self.SCREEN_MY_INFO)
        self.account_detail_event_loop.exec_()

    def set_acount_num(self):
        account_list = self.dynamicCall("GetLoginInfo(String)", "ACCNO")
        self.account_num = account_list.split(":")[0][:-1]
        print("내 계좌번호: %s" % self.account_num)

    def set_deposit_details(self, sRQName, sTrCode):
        self.deposit = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "예수금")
        self.deposit_ok = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "출금가능금액")
        self.use_money = int(self.deposit) * self.USE_MONEY_PERCENT
        self.use_money = self.use_money / 4

    def set_mystock_not_concluded_details(self, sRQName, sTrCode):  # 미체결 요청
        rows = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)

        for i in range(rows):
            code = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목코드")
            code_nm = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목명")
            order_no = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문번호")
            order_status = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                            "주문상태")  # 접수, 확인, 체결
            order_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                              "주문수량")
            order_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문가격")
            order_classification = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                                    "주문구분")  # -매도, +매수, 정정주문, 취소주문
            not_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "미체결수량")
            ok_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "체결량")

            self.mystock_not_concluded_dict = {}
            order_no = order_no.strip()

            if order_no not in self.mystock_not_concluded_dict:
                self.mystock_not_concluded_dict[order_no] = {}

            mystock_not_concluded_dict_item = self.mystock_not_concluded_dict[order_no]
            mystock_not_concluded_dict_item.update({"종목코드": code.strip()})
            mystock_not_concluded_dict_item.update({"종목명": code_nm.strip()})
            mystock_not_concluded_dict_item.update({"주문번호": order_no})
            mystock_not_concluded_dict_item.update({"주문상태": order_status.strip()})
            mystock_not_concluded_dict_item.update({"주문수량": int(order_quantity.strip())})
            mystock_not_concluded_dict_item.update({"주문가격": int(order_price.strip())})
            mystock_not_concluded_dict_item.update({"주문구분": order_classification.strip().lstrip("+").lstrip("-")})
            mystock_not_concluded_dict_item.update({"미체결수량": int(not_quantity.strip())})
            mystock_not_concluded_dict_item.update({"체결량": int(ok_quantity.strip())})

    def set_mystock_details(self, sRQName, sTrCode, sPrevNext):
        self.total_buy_money = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "총매입금액")
        self.total_profit_loss_rate = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0,
                                                  "총수익률(%)")
        row_size = self.dynamicCall("GetRepeatCnt(QString, QString", sTrCode, sRQName)

        for i in range(row_size):
            code = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목번호")
            code_nm = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목명")
            stock_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                              "보유수량")
            buy_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "매입가")
            learn_rate = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "수익률(%)")
            current_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "현재가")
            total_buy_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                               "매입금액")
            possible_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                                 "매매가능수량")

            code = code.strip()[1:]
            if code not in self.mystock_dict:
                self.mystock_dict.update({code: {}})

            mystock_dict_item = self.mystock_dict[code]
            mystock_dict_item.update({"종목명": code_nm.strip()})
            mystock_dict_item.update({"보유수량": int(stock_quantity.strip())})
            mystock_dict_item.update({"매입가": int(buy_price.strip())})
            mystock_dict_item.update({"수익률(%)": float(learn_rate.strip())})
            mystock_dict_item.update({"현재가": int(current_price.strip())})
            mystock_dict_item.update({"매입금액": int(total_buy_price.strip())})
            mystock_dict_item.update({"매매가능수량": int(possible_quantity.strip())})

        if sPrevNext == "2":
            self.signal_account_detail_mystock(sPrevNext="2")
