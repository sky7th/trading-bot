import os


def save_stock_info(code, code_nm, price):
    f = open("files/condition_stock.txt", "a", encoding="utf8")
    f.write("%s\t%s\t%s\t" % (code, code_nm, price))
    f.close()


def delete_stock_info():
    if os.path.isfile('files/condition_stock.txt'):
        os.remove('files/condition_stock.txt')


def read_stock_info():
    portfolio_stock_dict = {}

    if os.path.exists("files/condition_stock.txt"):
        f = open("files/condition_stock.txt", "r", encoding="utf8")

        lines = f.readlines()
        for line in lines:
            if line != "":
                lst = line.split("\t")
                stock_code = lst[0]
                stock_name = lst[1]
                stock_price = abs(int(lst[2].split("\n")[0]))
                portfolio_stock_dict.update({stock_code: {"종목명": stock_name, "현재가": stock_price}})
        f.close()

    return portfolio_stock_dict
