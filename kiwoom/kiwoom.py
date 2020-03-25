from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from PyQt5.QtTest import *
from config.returnCode import ReturnCode
from config import utils
from config.typeCode import Type as Type
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

        self.deposit = None
        self.deposit_ok = None
        self.total_buy_money = None
        self.total_profit_loss_rate = None
        self.mystock_dict = {}
        self.mystock_not_concluded_dict = {}
        self.analyzed_stocks = {}
        self.portfolio_stock_dict = {}
        self.jango_dict = {}

        self.analysis_data = []

        # 계좌 관련 변수
        self.account_num = None
        self.use_money = 0

        # eventloop
        self.login_event_loop = QEventLoop()
        self.request_loop = QEventLoop()

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
        print(ReturnCode.CAUSE[errCode])
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

            if sPrevNext == Type.NUM["다음"]:
                self.signal_account_detail_mystock(sPrevNext=Type.NUM["다음"])
            else:
                print("총매입금액: %s" % int(self.total_buy_money))
                print("총수익률(%%): %s" % float(self.total_profit_loss_rate))
                print("계좌 내 종목 개수: %s" % len(self.mystock_dict))
                print("계좌 내 종목 내역: %s" % self.mystock_dict)

        elif sRQName == "실시간미체결요청":
            self.set_mystock_not_concluded_details(sRQName, sTrCode)
            print("계좌 내 미체결 종목 개수: %s" % len(self.mystock_not_concluded_dict))
            print("계좌 내 미체결 종목 내역: %s" % self.mystock_not_concluded_dict)

        elif sRQName == "주식일봉차트조회":
            self.set_analysis_data(sRQName, sTrCode)
            code = self.get_comm_data(sTrCode, sRQName, 0, "종목코드")

            if sPrevNext == Type.NUM["다음"]:
                self.signal_bars_of_day(code, sPrevNext=Type.NUM["다음"])
            else:
                if granvileLaw.is_possible_4th_law(self.analysis_data):
                    code_nm = self.get_master_code_name(code)
                    utils.save_stock_info(code, code_nm, str(self.analysis_data[0][1]))

                self.analysis_data.clear()

        self.request_loop.exit()

    def realdata_slot(self, sCode, sRealType, sRealData):
        if sRealType == "장시작시간":
            self.print_market_status(sCode, sRealType)

        elif sRealType == "주식체결":
            self.update_portfolio_by_real_stock_conclusion(sCode, sRealType)
            self.trade(sCode)
            self.cancel_mystock_not_concluded(sCode)

    def print_market_status(self, sCode, sRealType):
        status_code = self.get_comm_real_data(sCode, sRealType, "장운영구분")

        if status_code == Type.NUM["시작전"]:
            print("장 시작 전")
        elif status_code == Type.NUM["시작"]:
            print("장 시작")
        elif status_code == Type.NUM["동시호가"]:
            print("장 종료, 동시호가로 넘어감")
        elif status_code == Type.NUM["종료"]:
            print("3시 30분 장 종료")

    def update_portfolio_by_real_stock_conclusion(self, sCode, sRealType):
        if sCode not in self.portfolio_stock_dict:
            self.portfolio_stock_dict.update({sCode: {}})

        portfolio_stock_dict_item = self.portfolio_stock_dict[sCode]
        cols = ['체결시간', '현재가', '전일대비', '등락율', '(최우선)매도호가',
                '(최우선)매수호가', '거래량', '누적거래량', '고가', '시가', '저가']
        for col in cols:
            col_data = self.get_comm_real_data(sCode, sRealType, col)
            if col == '체결시간':
                portfolio_stock_dict_item.update({col: col_data})
            elif col == '등락율':
                portfolio_stock_dict_item.update({col: float(col_data)})
            else:
                portfolio_stock_dict_item.update({col: abs(int(col_data))})

    def trade(self, sCode):
        if sCode in self.mystock_dict.keys() and sCode not in self.jango_dict.keys():
            self.sell_before_stock(sCode, self.mystock_dict[sCode], self.portfolio_stock_dict[sCode])
        elif sCode in self.jango_dict.keys():
            self.sell_today_stock(sCode, self.jango_dict[sCode], self.portfolio_stock_dict[sCode])

        if sCode not in self.jango_dict.keys():
            self.buy_today_stock(sCode, self.portfolio_stock_dict[sCode])

    def sell_before_stock(self, sCode, my_stock, now_stock):
        my_fluctuation = (now_stock["현재가"] - my_stock["매입가"]) / my_stock["매입가"] * 100
        if my_stock["매매가능수량"] > 0 and (my_fluctuation > 2 or my_fluctuation < -2):
            self.sell_order_process(sCode, my_stock, my_stock["매매가능수량"])

    def sell_today_stock(self, sCode, jango_dict, now_stock):
        my_fluctuation = (now_stock["현재가"] - jango_dict["매입단가"]) / jango_dict["매입단가"] * 100
        if jango_dict["보유수량"] > 0 and (my_fluctuation > 2 or my_fluctuation < -2):
            self.sell_order_process(sCode, jango_dict, jango_dict["보유수량"])

    def sell_order_process(self, sCode, dict_item, quantity):
        order_success = self.send_order("신규매도", self.portfolio_stock_dict[sCode]["주문용스크린번호"], self.account_num, 2,
                                        sCode, quantity, 0, Type.SEND["거래구분"]["시장가"], "")
        if order_success == ReturnCode.OP_ERR_NONE:
            print("매도주문 전달 성공", sCode)
            # TODO: 매도 주문 검증 절차
            del dict_item
        else:
            print("매도주문 전달 실패(%s)" % ReturnCode.CAUSE[order_success])

    def sell_order_cancel_process(self, sCode):
        order_success = self.send_order("매도취소", self.portfolio_stock_dict[sCode]["주문용스크린번호"], self.account_num, 4,
                                        sCode, 0, 0, Type.SEND["거래구분"]["시장가"], "")
        if order_success == ReturnCode.OP_ERR_NONE:
            print("매도취소 성공", sCode)
            # TODO: 매도 취소 검증 절차
        else:
            print("매도취소 실패(%s)" % ReturnCode.CAUSE[order_success])

    def buy_today_stock(self, sCode, portfolio_stock):
        if portfolio_stock["등락율"] > 2.0:
            buy_quantity = (self.use_money // len(self.analyzed_stocks)) // portfolio_stock['(최우선)매도호가']
            print(buy_quantity, portfolio_stock['(최우선)매도호가'])
            self.buy_order_process(sCode, int(buy_quantity), portfolio_stock)

    def buy_order_process(self, sCode, quantity, portfolio_stock):
        order_success = self.send_order("신규매수", self.portfolio_stock_dict[sCode]["주문용스크린번호"], self.account_num, 1,
                                        sCode, quantity, portfolio_stock['(최우선)매도호가'], Type.SEND["거래구분"]["시장가"], "")
        if order_success == ReturnCode.OP_ERR_NONE:
            print("매수주문 전달 성공", sCode)
            # TODO: 매수 주문 검증 절차
        else:
            print("매수주문 전달 실패(%s)" % ReturnCode.CAUSE[order_success])

    def buy_order_cancel_process(self, sCode):
        order_success = self.send_order("매수취소", self.portfolio_stock_dict[sCode]["주문용스크린번호"], self.account_num, 3,
                                        sCode, 0, 0, Type.SEND["거래구분"]["시장가"], "")
        if order_success == ReturnCode.OP_ERR_NONE:
            print("매수취소 성공", sCode)
            # TODO: 매수 취소 검증 절차
        else:
            print("매수취소 실패(%s)" % ReturnCode.CAUSE[order_success])

    def cancel_mystock_not_concluded(self, sCode):
        not_selling_list = list(self.mystock_not_concluded_dict)

        for order_num in not_selling_list:
            code = self.mystock_not_concluded_dict[order_num]["종목코드"]
            my_order_price = self.mystock_not_concluded_dict[order_num]["(최우선)매도호가"]
            not_conclusion_quantity = self.mystock_not_concluded_dict[order_num]["미체결수량"]
            order_classification = self.mystock_not_concluded_dict[order_num]["주문구분"]

            if order_classification == "매수" and not_conclusion_quantity > 0 and self.portfolio_stock_dict[code]["(최우선)매도호가"] > my_order_price:
                print('취소할 주문번호: ', order_num, self.portfolio_stock_dict[code]["(최우선)매도호가"], my_order_price, '미체결 리스트: ', not_selling_list)
                self.buy_order_cancel_process(code)
            elif not_conclusion_quantity == 0:
                del self.mystock_not_concluded_dict[order_num]

            # if order_classification == "매도" and not_conclusion_quantity > 0 and self.portfolio_stock_dict[sCode]["(최우선)매수호가"] < my_order_price:
            #     print("%s %s" % ("매도취소 한다", sCode))
            #     self.sell_order_cancel_process(sCode)
            # elif not_conclusion_quantity == 0:
            #     del self.mystock_not_concluded_dict[order_num]

    def chejan_slot(self, sGubun, nItemCnt, sFidList):
        if sGubun == Type.NUM["주문미체결"]:
            self.update_mystock_not_concluded_by_real()

        elif sGubun == Type.NUM["잔고"]:
            self.update_jango_by_real()

    def update_mystock_not_concluded_by_real(self):
        order_number = self.get_chejan_data(Type.REAL['주문체결']['주문번호'])
        if order_number not in self.mystock_not_concluded_dict.keys():
            self.mystock_not_concluded_dict.update({order_number: {}})

        mystock_not_concluded_dict_item = self.mystock_not_concluded_dict[order_number]
        cols = ['종목코드', '종목명', '원주문번호', '주문번호', '주문상태',
                '주문수량', '주문가격', '미체결수량', '주문구분', '주문/체결시간',
                '체결가', '체결량', '현재가', '(최우선)매도호가', '(최우선)매수호가']

        for col in cols:
            col_data = self.get_chejan_data(Type.REAL['주문체결'][col])
            # print(col, col_data)
            if col == '종목코드':
                mystock_not_concluded_dict_item.update({col: col_data[1:].strip()})
            elif col == '주문구분':
                mystock_not_concluded_dict_item.update({col: col_data.strip().lstrip("+").lstrip("-")})
            elif col in ['체결가', '체결량']:
                mystock_not_concluded_dict_item.update({col: 0 if col_data == "" else int(col_data)})
            elif col in ['주문수량', '주문가격', '미체결수량', '현재가', '(최우선)매도호가', '(최우선)매수호가']:
                mystock_not_concluded_dict_item.update({col: abs(int(col_data))})
            else:
                mystock_not_concluded_dict_item.update({col: col_data.strip()})
        print('추가된 미체결: ', order_number, mystock_not_concluded_dict_item)

    def update_jango_by_real(self, code=None):
        if code is None:
            code = self.get_chejan_data(Type.REAL['잔고']['종목코드'])[1:]

        if code not in self.jango_dict.keys():
            self.jango_dict.update({code: {}})

        jango_dict_item = self.jango_dict[code]
        cols = ['현재가', '종목코드', '종목명', '보유수량', '주문가능수량', '매입단가',
                '총매입가', '매도매수구분', '(최우선)매도호가', '(최우선)매수호가']

        for col in cols:
            col_data = self.get_chejan_data(Type.REAL['잔고'][col])

            if col == '보유수량' and col_data == 0:
                del self.jango_dict[code]
                self.set_real_remove(self.portfolio_stock_dict[code]["스크린번호"], code)
                return

            if col in ['현재가', '(최우선)매도호가', '(최우선)매수호가']:
                jango_dict_item.update({col: abs(int(col_data.strip()))})
            elif col in ['보유수량', '주문가능수량', '매입단가', '총매입가']:
                jango_dict_item.update({col: int(col_data.strip())})
            elif col == '매도매수구분':
                jango_dict_item.update({col: col_data.strip().lstrip("+").lstrip("-")})
            else:
                jango_dict_item.update({col: col_data.strip()})
        print('추가된 잔고: ', jango_dict_item)

    def signal_login_commConnect(self):
        self.dynamicCall('CommConnect()')
        self.login_event_loop.exec_()

    def signal_account_detail_info(self):
        self.set_input_value("계좌번호", self.account_num)
        self.set_input_value("비밀번호", "0000")
        self.set_input_value("비밀번호입력매체구분", "00")
        self.set_input_value("조회구분", Type.NUM["추정조회"])
        self.comm_rq_data("예수금상세현황요청", "opw00001", Type.NUM["처음"], self.SCREEN_MY_INFO)

    def signal_account_detail_mystock(self, sPrevNext=Type.NUM["처음"]):
        self.set_input_value("계좌번호", self.account_num)
        self.set_input_value("비밀번호", "0000")
        self.set_input_value("비밀번호입력매체구분", "00")
        self.set_input_value("조회구분", Type.NUM["추정조회"])
        self.comm_rq_data("계좌평가잔고내역요청", "opw00018", sPrevNext, self.SCREEN_MY_INFO)

    def signal_account_detail_mystock_not_concluded(self, sPrevNext=Type.NUM["처음"]):
        self.set_input_value("계좌번호", self.account_num)
        self.set_input_value("체결구분", Type.NUM["미체결"])
        self.set_input_value("매매구분", Type.NUM["전체"])
        self.comm_rq_data("실시간미체결요청", "opt10075", sPrevNext, self.SCREEN_MY_INFO)

    def real_signal_market_start_time(self):
        self.set_real_reg(self.SCREEN_REAL_START_STOP, "",
                         Type.REAL["장시작시간"]["장운영구분"], Type.NUM["새로등록"])

    def real_signal_stock_conclusion(self):
        for code in self.portfolio_stock_dict.keys():
            screen_num = self.portfolio_stock_dict[code]["스크린번호"]
            fids = Type.REAL["주식체결"]["체결시간"]
            self.set_real_reg(screen_num, code, fids, Type.NUM["추가등록"])

    def get_acount_num(self):
        account_list = self.dynamicCall("GetLoginInfo(String)", "ACCNO")
        return account_list.split(":")[0][:-1]

    def set_deposit_details(self, sRQName, sTrCode):
        self.deposit = self.get_comm_data(sTrCode, sRQName, 0, "예수금")
        self.deposit_ok = self.get_comm_data(sTrCode, sRQName, 0, "출금가능금액")
        self.use_money = int(self.deposit) * self.USE_MONEY_PERCENT

    def set_mystock_not_concluded_details(self, sRQName, sTrCode):  # 미체결 요청
        row_size = self.get_repeat_cnt(sTrCode, sRQName)

        for i in range(row_size):
            order_no = self.get_comm_data(sTrCode, sRQName, i, '주문번호')
            if order_no not in self.mystock_not_concluded_dict:
                self.mystock_not_concluded_dict[order_no] = {}

            mystock_not_concluded_dict_item = self.mystock_not_concluded_dict[order_no]
            cols = ['종목코드', '종목명', '주문번호', '주문상태', '주문수량',
                    '주문가격', '주문구분', '미체결수량', '체결량']

            for col in cols:
                col_data = self.get_comm_data(sTrCode, sRQName, i, col)

                if col in ['주문수량', '주문가격', '미체결수량', '체결량']:
                    mystock_not_concluded_dict_item.update({col: int(col_data)})
                elif col == '주문구분':
                    mystock_not_concluded_dict_item.update({col: col_data.lstrip("+").lstrip("-")})
                else:
                    mystock_not_concluded_dict_item.update({col: col_data})

    def set_mystock_details(self, sRQName, sTrCode, sPrevNext):
        self.total_buy_money = self.get_comm_data(sTrCode, sRQName, 0, "총매입금액")
        self.total_profit_loss_rate = self.get_comm_data(sTrCode, sRQName, 0, "총수익률(%)")

        row_size = self.get_repeat_cnt(sTrCode, sRQName)

        for i in range(row_size):
            code = self.get_comm_data(sTrCode, sRQName, i, '종목번호')[1:]
            if code not in self.mystock_dict:
                self.mystock_dict.update({code: {}})

            mystock_dict = self.mystock_dict[code]
            cols = ['종목명', '보유수량', '매입가', '수익률(%)', '현재가', '매입금액', '매매가능수량']

            for col in cols:
                col_data = self.get_comm_data(sTrCode, sRQName, i, col)

                if col == '종목명':
                    mystock_dict.update({col: col_data})
                elif col == '현재가':
                    mystock_dict.update({col: abs(int(col_data.strip()))})
                elif col == '수익률(%)':
                    mystock_dict.update({col: float(col_data.strip())})
                else:
                    mystock_dict.update({col: int(col_data.strip())})

    def set_analysis_data(self, sRQName, sTrCode):
        row_size = self.get_repeat_cnt(sTrCode, sRQName)

        for i in range(row_size):
            data = ['']
            cols = ['현재가', '거래량', '거래대금', '일자', '시가', '고가', '저가']
            for col in cols:
                col_data = self.get_comm_data(sTrCode, sRQName, i, col)
                if col == '일자':
                    data.append(col_data)
                else:
                    data.append(int(col_data))
            data.append('')
            self.analysis_data.append(data)

    def analyze_chart(self):
        code_list = self.get_code_list_by_market(self.KOSDAQ_NUM)

        for idx, code in enumerate(code_list):
            self.disconnect_real_data(self.SCREEN_CALCULATION_STOCK)
            print("%s / %s : 코스닥 주식 정보 : %s 업데이트 중..." % (idx+1, len(code_list), code))
            self.signal_bars_of_day(code)

    def signal_bars_of_day(self, code=None, date=None, sPrevNext=Type.NUM["처음"]):
        QTest.qWait(200)
        self.set_input_value("종목코드", code)
        self.set_input_value("수정주가구분", Type.NUM["수정후"])

        if date is None:
            self.set_input_value("기준일자", date)

        self.comm_rq_data("주식일봉차트조회", "opt10081", sPrevNext, self.SCREEN_CALCULATION_STOCK)  # Tr서버로 전송 - Transaction
        self.request_loop.exec_()

    def set_portfolio_stock_dict(self):
        self.analyzed_stocks = utils.read_stock_info()
        combined_list = list(self.analyzed_stocks.keys())
        combined_list += list(self.mystock_dict.keys())
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

    #################################################################
    # 조회 관련 메서드                                        #
    # 시세조회, 관심종목 조회, 조건검색 등 이들의 합산 조회 횟수가 1초에 5회까지 허용 #
    #################################################################
    def set_input_value(self, id, value):
        self.dynamicCall("SetInputValue(QString, QString)", id, value)

    def get_comm_data(self, sTrCode, sRQName, i, column_nm):
        ret = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, column_nm)
        return ret.strip()

    def comm_rq_data(self, request_name, tr_code, inquiry, screen_no):
        self.dynamicCall("CommRqData(QString, QString, int, QString)", request_name, tr_code, inquiry, screen_no)
        self.request_loop.exec_()

    def get_repeat_cnt(self, tr_code, request_name):
        ret = self.dynamicCall("GetRepeatCnt(QString, QString)", tr_code, request_name)
        return ret

    ###############################################################
    # 실시간 데이터 처리 관련 메서드                           #
    ###############################################################
    def get_comm_real_data(self, code, sRealType, column_nm):
        value = self.dynamicCall("GetCommRealData(QString, int)", code, Type.REAL[sRealType][column_nm])
        return value

    def disconnect_real_data(self, screen_no):
        self.dynamicCall("DisconnectRealData(QString)", screen_no)

    def set_real_reg(self, screen_no, codes, fids, real_reg_type):
        self.dynamicCall("SetRealReg(QString, QString, QString, QString)", screen_no, codes, fids, real_reg_type)

    def set_real_remove(self, screen_no, code):
        self.dynamicCall("SetRealRemove(QString, QString)", screen_no, code)

    ###############################################################
    # 주문과 잔고처리 관련 메서드                              #
    # 1초에 5회까지 주문 허용                                          #
    ###############################################################
    def send_order(self, request_name, screen_no, account_no, order_type, code, qty, price, hoga_type, origin_order_no):
        QTest.qWait(200)
        return_code = self.dynamicCall("SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                                       [request_name, screen_no, account_no, order_type, code, qty, price, hoga_type, origin_order_no])
        return return_code

    def get_chejan_data(self, nFid):
        ret = self.dynamicCall("GetChejanData(int)", nFid)
        return ret

    ###############################################################
    # 그 외 메서드                                               #
    ###############################################################
    def get_master_code_name(self, code):
        code_nm = self.dynamicCall("GetMasterCodeName(QString)", code)
        return code_nm.strip()

    def get_code_list_by_market(self, market_code):
        code_list = self.dynamicCall("GetCodeListByMarket(QString)", market_code)
        return code_list.split(";")[:-1]
