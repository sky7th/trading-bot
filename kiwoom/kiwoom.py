from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from PyQt5.QtTest import *
from config.errorCode import *
from config import utils
from config.typeCode import *
from buyingLaw import granvileLaw


class Kiwoom(QAxWidget):

    USE_MONEY_PERCENT = 0.5
    KOSDAQ_NUM = "10"

    SCREEN_MY_INFO = "2000"
    SCREEN_CALCULATION_STOCK = "4000"
    SCREEN_REAL_STOCK = "5000"  # 종목별로 할당할 스크린 번호
    SCREEN_REAL_SELLING_STOCK = "6000"  # 종목별로 할당할 주문용 스크린 번호
    SCREEN_REAL_START_STOP = "1000"

    def __init__(self):
        super().__init__()

        self.typeCode = TypeCode()

        self.deposit = None
        self.deposit_ok = None
        self.total_buy_money = None
        self.total_profit_loss_rate = None
        self.mystock_dict = {}
        self.mystock_not_concluded_dict = {}
        self.portfolio_stock_dict = {}
        self.jango_dict = {}

        self.analysis_data = []

        # 계좌 관련 변수
        self.account_num = None
        self.use_money = 0

        # eventloop
        self.login_event_loop = QEventLoop()
        self.account_detail_event_loop = QEventLoop()
        self.analysis_event_loop = QEventLoop()

        self.__get_ocx_instance()
        self.__event_slots()
        self.__real_event_slots()

        self.signal_login_commConnect()

        self.account_num = self.get_acount_num()
        self.signal_account_detail_info()  # 예수금상세현황요청
        self.signal_account_detail_mystock()  # 계좌평가잔고내역요청
        self.signal_account_detail_mystock_not_concluded()  # 실시간미체결요청

        # self.analyze_chart()  # 종목 분석용, 임시용으로 실행

        self.set_portfolio_stock_dict()  # 스크린 번호를 할당

        self.real_signal_market_start_time()  # 장시작시간
        self.real_signal_stock_conclusion()  # 주식체결

        print("내 계좌번호: %s" % self.account_num)

    def __get_ocx_instance(self):
        self.setControl('KHOPENAPI.KHOpenAPICtrl.1')

    def __event_slots(self):
        self.OnEventConnect.connect(self.login_slot)
        self.OnReceiveTrData.connect(self.trdata_slot)

    def __real_event_slots(self):
        self.OnReceiveRealData.connect(self.realdata_slot)
        self.OnReceiveChejanData.connect(self.chejan_slot)

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
            self.account_detail_event_loop.exit()

        elif sRQName == "계좌평가잔고내역요청":
            self.set_mystock_details(sRQName, sTrCode, sPrevNext)

            if sPrevNext == "2":
                self.signal_account_detail_mystock(sPrevNext="2")
            else:
                print("총매입금액: %s" % int(self.total_buy_money))
                print("총수익률(%%): %s" % float(self.total_profit_loss_rate))
                print("계좌 내 종목 개수: %s" % len(self.mystock_dict))
                print("계좌 내 종목 내역: %s" % self.mystock_dict)
                self.account_detail_event_loop.exit()

        elif sRQName == "실시간미체결요청":
            self.set_mystock_not_concluded_details(sRQName, sTrCode)
            print("계좌 내 미체결 종목 개수: %s" % len(self.mystock_not_concluded_dict))
            print("계좌 내 미체결 종목 내역: %s" % self.mystock_not_concluded_dict)
            self.account_detail_event_loop.exit()

        elif sRQName == "주식일봉차트조회":
            self.set_analysis_data(sRQName, sTrCode)
            code = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "종목코드")
            code = code.strip()

            if sPrevNext == "2":
                self.signal_bars_of_day(code, sPrevNext="2")
            else:
                if granvileLaw.is_possible_4th_law(self.analysis_data):
                    code_nm = self.dynamicCall("GetMasterCodeName(QString)", code)
                    utils.save_stock_info(code, code_nm, str(self.analysis_data[0][1]))

                self.analysis_data.clear()
                self.analysis_event_loop.exit()

    def realdata_slot(self, sCode, sRealType, sRealData):
        if sRealType == "장시작시간":
            self.print_market_status(sCode, sRealType)

        elif sRealType == "주식체결":
            self.update_portfolio_by_real_stock_conclusion(sCode, sRealType)
            self.trade(sCode)
            self.cancel_mystock_not_concluded(sCode)

    def print_market_status(self, sCode, sRealType):
        status_code = self.dynamicCall("GetCommRealData(QString, int)", sCode, self.typeCode.REALTYPE[sRealType]["장운영구분"])

        if status_code == "0":
            print("장 시작 전")
        elif status_code == "3":
            print("장 시작")
        elif status_code == "2":
            print("장 종료, 동시호가로 넘어감")
        elif status_code == "4":
            print("3시 30분 장 종료")

    def update_portfolio_by_real_stock_conclusion(self, sCode, sRealType):
        a = self.dynamicCall("GetCommRealData(QString, int)", sCode,
                             self.typeCode.REALTYPE[sRealType]["체결시간"])  # HHMMSS
        b = self.dynamicCall("GetCommRealData(QString, int)", sCode, self.typeCode.REALTYPE[sRealType]["현재가"])
        c = self.dynamicCall("GetCommRealData(QString, int)", sCode, self.typeCode.REALTYPE[sRealType]["전일대비"])
        d = self.dynamicCall("GetCommRealData(QString, int)", sCode, self.typeCode.REALTYPE[sRealType]["등락율"])
        e = self.dynamicCall("GetCommRealData(QString, int)", sCode, self.typeCode.REALTYPE[sRealType]["(최우선)매도호가"])
        f = self.dynamicCall("GetCommRealData(QString, int)", sCode, self.typeCode.REALTYPE[sRealType]["(최우선)매수호가"])
        g = self.dynamicCall("GetCommRealData(QString, int)", sCode, self.typeCode.REALTYPE[sRealType]["거래량"])
        h = self.dynamicCall("GetCommRealData(QString, int)", sCode, self.typeCode.REALTYPE[sRealType]["누적거래량"])
        i = self.dynamicCall("GetCommRealData(QString, int)", sCode, self.typeCode.REALTYPE[sRealType]["고가"])
        j = self.dynamicCall("GetCommRealData(QString, int)", sCode, self.typeCode.REALTYPE[sRealType]["시가"])
        k = self.dynamicCall("GetCommRealData(QString, int)", sCode, self.typeCode.REALTYPE[sRealType]["저가"])

        if sCode not in self.portfolio_stock_dict:
            self.portfolio_stock_dict.update({sCode: {}})

        self.portfolio_stock_dict[sCode].update({"체결시간": a})
        self.portfolio_stock_dict[sCode].update({"현재가": abs(int(b))})
        self.portfolio_stock_dict[sCode].update({"전일대비": abs(int(c))})
        self.portfolio_stock_dict[sCode].update({"등락율": float(d)})
        self.portfolio_stock_dict[sCode].update({"(최우선)매도호가": abs(int(e))})
        self.portfolio_stock_dict[sCode].update({"(최우선)매수호가": abs(int(f))})
        self.portfolio_stock_dict[sCode].update({"거래량": abs(int(g))})
        self.portfolio_stock_dict[sCode].update({"누적거래량": abs(int(h))})
        self.portfolio_stock_dict[sCode].update({"고가": abs(int(i))})
        self.portfolio_stock_dict[sCode].update({"시가": abs(int(j))})
        self.portfolio_stock_dict[sCode].update({"저가": abs(int(k))})

    def trade(self, sCode):
        my_stock = self.mystock_dict[sCode]
        now_stock = self.portfolio_stock_dict[sCode]

        if sCode in self.mystock_dict.keys() and sCode not in self.jango_dict.keys():
            self.sell_mystock(sCode, my_stock, now_stock)

        elif sCode in self.jango_dict.keys():
            self.sell_jango(sCode)
        else:
            self.buy_new_stock(sCode, my_stock, now_stock)

    def sell_mystock(self, sCode, my_stock, now_stock):
        my_fluctuation = (now_stock["현재가"] - my_stock["매입가"]) / my_stock["매입가"] * 100
        if my_stock["매매가능수량"] > 0 and (my_fluctuation > 5 or my_fluctuation < -5):
            order_success = self.dynamicCall(
                "SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                ["신규매도", self.portfolio_stock_dict[sCode]["주문용스크린번호"], self.account_num, 2,
                sCode, my_stock["매매가능수량"], 0, self.typeCode.SENDTYPE["거래구분"]["시장가"], ""])
            if order_success == 0:
                print("매도주문 전달 성공")
                # TODO: 매도 주문 검증 절차
                del my_stock
            else:
                print("매도주문 전달 실패(%s)" % errors(order_success))

    def sell_jango(self, sCode):
        print("%s %s" % ("잔고 내 신규매도를 함", sCode))
        # TODO: 잔고 주식 매도

    def buy_new_stock(self, sCode, my_stock, now_stock):
        if now_stock["등락율"] > 2.0 and sCode not in self.jango_dict:
            print("%s %s" % ("신규매수를 함", sCode))
            # TODO: 신규 매수

    def cancel_mystock_not_concluded(self, sCode):
        not_selling_list = list(self.mystock_not_concluded_dict)

        for order_num in not_selling_list:
            code = self.mystock_not_concluded_dict[order_num]["종목코드"]
            my_order_price = self.mystock_not_concluded_dict[order_num]["주문가격"]
            not_conclusion_quantity = self.mystock_not_concluded_dict[order_num]["미체결수량"]
            selling_classification = self.mystock_not_concluded_dict[order_num]["매도수구분"]

            if selling_classification == "매수" and not_conclusion_quantity > 0 and self.portfolio_stock_dict[sCode]["(최우선)매도호가"] > my_order_price:
                print("%s %s" % ("매수취소 한다", sCode))
                # Todo: 미체결 주식 매수 취소
            elif not_conclusion_quantity == 0:
                del self.mystock_not_concluded_dict[order_num]

            if selling_classification == "매도" and not_conclusion_quantity > 0 and self.portfolio_stock_dict[sCode]["(최우선)매수호가"] < my_order_price:
                print("%s %s" % ("매도취소 한다", sCode))
                # Todo: 미체결 주식 매도 취소
            elif not_conclusion_quantity == 0:
                del self.mystock_not_concluded_dict[order_num]

    def chejan_slot(self, sGubun, nItemCnt, sFidList):
        sGubun = int(sGubun)

        if sGubun == "0":
            print("주문체결")
        elif sGubun == "1":
            print("잔고")

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
        self.dynamicCall("CommRqData(String, String, int, String)", "계좌평가잔고내역요청", "opw00018", sPrevNext, self.SCREEN_MY_INFO)
        self.account_detail_event_loop.exec_()

    def signal_account_detail_mystock_not_concluded(self, sPrevNext="0"):
        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(QString, QString)", "체결구분", "1")
        self.dynamicCall("SetInputValue(QString, QString)", "매매구분", "0")
        self.dynamicCall("CommRqData(String, String, int, String)", "실시간미체결요청", "opt10075", sPrevNext, self.SCREEN_MY_INFO)
        self.account_detail_event_loop.exec_()

    def real_signal_market_start_time(self):
        self.dynamicCall("SetRealReg(QString, QString, QString, QString)", self.SCREEN_REAL_START_STOP, "",
                         self.typeCode.REALTYPE["장시작시간"]["장운영구분"], "0")

    def real_signal_stock_conclusion(self):
        for code in self.portfolio_stock_dict.keys():
            screen_num = self.portfolio_stock_dict[code]["스크린번호"]
            fids = self.typeCode.REALTYPE["주식체결"]["체결시간"]
            self.dynamicCall("SetRealReg(QString, QString, QString, QString)", screen_num, code, fids, "1")

    def get_acount_num(self):
        account_list = self.dynamicCall("GetLoginInfo(String)", "ACCNO")
        return account_list.split(":")[0][:-1]

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
            order_status = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문상태")  # 접수, 확인, 체결
            order_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문수량")
            order_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문가격")
            order_classification = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문구분")  # -매도, +매수, 정정주문, 취소주문
            not_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "미체결수량")
            ok_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "체결량")
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
        self.total_profit_loss_rate = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "총수익률(%)")

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

    def set_analysis_data(self, sRQName, sTrCode):
        row_size = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)

        for i in range(row_size):
            data = []

            current_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "현재가")
            value = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "거래량")
            trading_value = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "거래대금")
            date = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "일자")
            start_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "시가")
            high_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "고가")
            low_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "저가")

            data.append("")
            data.append(int(current_price.strip()))
            data.append(int(value.strip()))
            data.append(int(trading_value.strip()))
            data.append(date.strip())
            data.append(int(start_price.strip()))
            data.append(int(high_price.strip()))
            data.append(int(low_price.strip()))
            data.append("")
            self.analysis_data.append(data)

    def analyze_chart(self):
        code_list = self.get_code_list_by_market(self.KOSDAQ_NUM)

        for idx, code in enumerate(code_list):
            self.dynamicCall("DisconnectRealData(QString)", self.SCREEN_CALCULATION_STOCK)
            print("%s / %s : 코스닥 주식 정보 : %s 업데이트 중..." % (idx+1, len(code_list), code))
            self.signal_bars_of_day(code)

    def get_code_list_by_market(self, market_code):
        code_list = self.dynamicCall("GetCodeListByMarket(QString)", market_code)
        return code_list.split(";")[:-1]

    def signal_bars_of_day(self, code=None, date=None, sPrevNext="0"):
        QTest.qWait(1600)
        self.dynamicCall("SetInputValue(QString, QString)", "종목코드", code)
        self.dynamicCall("SetInputValue(QString, QString)", "수정주가구분", "1")

        if date != None:
            self.dynamicCall("SetInputValue(QString, QString)", "기준일자", date)

        self.dynamicCall("CommRqData(QString, QString, int, QString", "주식일봉차트조회", "opt10081", sPrevNext, self.SCREEN_CALCULATION_STOCK)  # Tr서버로 전송 - Transaction
        self.analysis_event_loop.exec_()

    def set_portfolio_stock_dict(self):
        combined_list = list(self.mystock_dict.keys()) + list(utils.read_stock_info().keys())
        combined_list += [self.mystock_not_concluded_dict[order_no]["종목코드"] for order_no in list(self.mystock_not_concluded_dict.keys())]
        set_list = list(set(combined_list))

        for idx, code in enumerate(set_list):
            screen_num = int(self.SCREEN_REAL_STOCK)
            selling_screen_num = int(self.SCREEN_REAL_SELLING_STOCK)

            if (idx % 50) == 0:
                screen_num += 1
                self.SCREEN_REAL_STOCK = str(screen_num)

            if (idx % 50) == 0:
                selling_screen_num += 1
                self.SCREEN_REAL_SELLING_STOCK = str(selling_screen_num)

            if code in self.portfolio_stock_dict.keys():
                self.portfolio_stock_dict[code].update({"스크린번호": str(self.SCREEN_REAL_STOCK)})
                self.portfolio_stock_dict[code].update({"주문용스크린번호": str(self.SCREEN_REAL_SELLING_STOCK)})
            else:
                self.portfolio_stock_dict.update({code: {"스크린번호": str(self.SCREEN_REAL_STOCK), "주문용스크린번호": str(self.SCREEN_REAL_SELLING_STOCK)}})
