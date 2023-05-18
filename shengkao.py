# encoding:utf-8
import xlrd
import xlwt
import re

def check(row_value, major, edu, year, p_edu, p_major, p_fresh):
    ma = row_value[p_major]
    # 该岗位所需专业
    ed = row_value[p_edu]
    # 该岗位所需学历
    fresh = row_value[p_fresh]
    # 该岗位所需年限
    loc= row_value[0]
    # 该岗位位置
    #check_major(ma, major) and check_edu(ed, edu) and
    if check_major(ma, major) and check_edu(ed, edu) and checkSpecial(fresh, year):
        #print(value)
        print(row_value)
        print()
        print()
        return True
    else:
        return False

def check_major(value, major):
    # 检查是否满足专业要求
    pat = re.compile(major)
    if re.search(pat, value):
        return True
    return False

# 检查是否满足学历要求
def check_edu(value, edu):
    pat = re.compile(edu)
    if re.search(pat, value):
        return True
    return False

# 检查基层年限设置
def checkSpecial(value, year):
    pat = re.compile(year)
    if re.search(pat, value):
        return False
    return True
# 上面这个return True和其它两个不同 要改的话注意

# 根据条件筛选出职位
def filterTitle(file, major, edu, year):
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
        p_edu, p_major, p_fresh = 0, 0, 0
        for row in range(sheet.nrows):
            # sheet指的是xlsx的一个板块、单元
            row_value = sheet.row_values(row)

            if '录用人数' in row_value:
                p_fresh = row_value.index('是否限应届毕业生报考')
                p_edu = row_value.index('学历')
                p_major = row_value.index('研究生专业\n名称及代码')

                # d = row_value.index('招考单位')
                #print(a)
            choosed = check(row_value, major, edu, year, p_edu, p_major, p_fresh)
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
    filterTitle('2021省考公务员.xls', '计算机技术', '研究生', '否')
    filterTitle('2022省考公务员.xls', '计算机技术', '研究生', '否')
    filterTitle('2023省考公务员.xls', '计算机技术', '研究生', '否')