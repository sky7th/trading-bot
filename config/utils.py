def save_information_of_item_in_text_file(price, code, code_nm):
    f = open("files/condition_stock.txt", "a", encoding="utf8")
    f.write("%s\t%s\t%s\t" % (code, code_nm, price))
    f.close()