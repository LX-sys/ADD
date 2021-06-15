import string

# az = list(string.ascii_uppercase)
# print(az)
# print(len(az))

# 字母表(生成的表示与Excel同步)
def letterAable(number:int=702)->list:
    az = list(string.ascii_uppercase)
    temp = list()
    if number <=26:
        return az
    # 倍数
    multiple = number // 26
    for wk in az[:multiple]:
        for k in az:
            temp.append(wk+k)

    return (az + temp)[:number]


# c = letterAable(30)
# print(c)
# print(len(c))