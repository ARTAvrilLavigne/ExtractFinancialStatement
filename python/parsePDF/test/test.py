# -*- coding:utf-8 -*-

def is_chinese(string):
    for ch in string:
        if u'\u4e00' <= ch <= u'\u9fff':
            return True

    return False

def isSpecialCharacter(str):
    # string = "~!@#$%^&*()_+-*/<>,.[]\/"
    string = ",.-"
    for i in string:
        if i in str:
            return True
    return False

# print(is_chinese("资产合计123)"))
# print(isSpecialCharacter("1.1"))
# print(isSpecialCharacter("1,1"))
# print(isSpecialCharacter("-1.1"))
# print(isSpecialCharacter("-2,50"))
# print(isSpecialCharacter("-"))
# print(isSpecialCharacter("10"))

if (True and False):
    print('hhhh')
else:
    print('ffff')
