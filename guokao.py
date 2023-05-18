# encoding:utf-8
import xlrd
import xlwt
import re
'''
如果无法打开 xlsx
pip uninstall xlrd
pip install xlrd==1.2.0
'''
def check(row_value, major, edu, year, p_edu, p_major, p_others, p_city):

    c = True
    if p_city:
        city = row_value[p_city]
        c = check_v(city, '广东')
    if edu:
        ed = row_value[p_edu]
        c = c and check_v(ed, edu)
    if major:
        ma = row_value[p_major]
        c = c and check_v(ma, major)
        # 该岗位所需专业
    if p_others:
        others = row_value[p_others]
        c = c and check_v(others, '仅限应届毕业生')
        # 该岗位所需年限

    if c:
        #print(value)
        print(row_value)
        print()
        print()
        return True
    else:
        return False


# 检查是否满足学历要求
def check_v(value, goal):
    pat = re.compile(goal)
    if re.search(pat, value):
        return True
    return False


# 根据条件筛选出职位
def filter_guo_kao(file, major, edu, year):
    data = xlrd.open_workbook(file)
    # 将表格数据读取到data中
    res = ''
    output = xlwt.Workbook(encoding='utf-8')
    for sheet in data.sheets():
        output_sheet = output.add_sheet(sheet.name)
        # 筛选出来的文件中也添加这些子表格
        for col in range(sheet.ncols):
            # 添加第二行的列信息
            output_sheet.row(0).write(col, sheet.cell(1,col).value)
        #print(type(sheet), sheet)
        output_row = 1
        cnt = 0
        p_edu, p_major, p_fresh, p_city, p_others = 0, 0, 0, 0, 0
        for row in range(sheet.nrows):
            # sheet指的是xlsx的一个板块、单元
            row_value = sheet.row_values(row)

            if '部门名称' in row_value:
                #print(row_value)
                p_fresh = row_value.index('基层工作最低年限')
                p_major = row_value.index('专业')
                p_city = row_value.index('工作地点')
                p_others = row_value.index('备注')
                # d = row_value.index('招考单位')
                #print(a)
            choosed = check(row_value, major, None, year, None, p_major, p_others, p_city)
            # 是否满足三个条件（专业、学历、基层限制）

            if choosed:
                # 满足则输出到文件中
                res += str(row_value)
                res += '\n\n'
                for col in range(sheet.ncols):
                    output_sheet.row(output_row).write(col, sheet.cell(row, col).value)
                output_sheet.flush_row_data()
                output_row += 1
    s = file[0:6]
    txtname = s + '.txt'
    with open(txtname, 'w', encoding="utf-8") as f:
        f.write(res)
    # 要是想要xls的话就取消下面的注释
    '''
    filename = s + '.xls'
    output.save(filename)
    '''

if __name__ == '__main__':

    #filter_guo_kao('2021国考公务员.xls', None, None, '应届')
    # 不知道为什么筛选不了2021的 懒得改了（开摆）

    filter_guo_kao('2022国考公务员.xls', '计算机', None, '应届')
    # 因为岗位数量太少了 直接筛选应届计算机的 没有筛选学历 到时候自己筛选一下