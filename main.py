import os
import cv2
import xlwings as xw


# 转换单个数字
def num2abc(num):
    """ 给定索引num(1开始), 返回excel中对应的ABC列名, 此函数和进制转换不一样
        excel列名和数字转换规则: A->1, Z->26, 系数是从1-26(不是0-25, 这是和26进制不一样的地方)
            列名转数字: 先将每位的字母转成该位的数字, 如ABZ -> [1, 2, 26], 然后计算加权和, 1*26**2 + 2*26 + 26
            数字转列名: 根据列名转数字的写法可知, 整体是按照进制转换去做, 但是需要处理余数为0时, 要改成26, 并让除数减1
        ps: excel列名整体上来说就是进制转换, 唯一不同的是, 它的系数是1-26, 而不是0-25, 按照拆分后的带幂次数字去处理即可
    """
    abc_num = ''
    while num > 0:
        d, r = num // 26, num % 26
        if r == 0:  # 当余数是0时, 此位的值要改成26, 这是和进制转换不同的地方
            d, r = d - 1, 26
        abc_num += chr(r + ord('A') - 1)
        num = d
    return abc_num[::-1]


if __name__ == '__main__':
    print('[+]  \033[36m请将需要转化的文件拖拽到此处\033[0m')
    dirpath = input('[+]  \033[35m图片路径：\033[0m')
    img_cv   = cv2.imread(dirpath)#读取数据
    excel_path = dirpath + ".xlsx"
    print('[+]  \033[35m生成文件将存放在此处：\033[0m{}'.format(excel_path    ))
    # 当前App下新建一个Book， visible参数控制创建文件时可见的属性
    app=xw.App(visible=True, add_book=False)
    wb=app.books.add()
    sht = wb.sheets[0]
    sht.range("A1", "{0}{1}".format(num2abc(img_cv.shape[1]), img_cv.shape[0])).row_height = 10
    sht.range("A1", "{0}{1}".format(num2abc(img_cv.shape[1]), img_cv.shape[0])).column_width = 1
    for i in range(1, img_cv.shape[0]):
        print("当前进度{}/{}".format(i, img_cv.shape[0]))
        for j in range(1, img_cv.shape[1]): 
            sht.range("{0}{1}".format(num2abc(j), i)).color = img_cv[i-1, j-1][::-1]
    wb.save(excel_path)
    wb.close()
    #结束进程
    app.quit()
