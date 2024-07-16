import pandas as pd

def sheet1():
    blank = pd.MultiIndex.from_product([
        [' '],
        [' '],
        [' ']
        ])
    columns_1 = pd.MultiIndex.from_product([
        [' '],
        [' '],
        ['บริษัทประกันชีวิต']
    ])
    columns_2 = pd.MultiIndex.from_product([
        ['เบี้ยประกันชีวิต'],
        ['ประเภทสามัญ', 'ประเภทอุตสาหกรรม', 'ประเภทกลุ่ม'],
        ['ปีแรก', 'ปีต่อไป']
    ])
    add1 = pd.MultiIndex.from_product([
        ['เบี้ยประกันชีวิต'],
        ['การรับประกันภัย'],
        ['อุบัติเหตุส่วนบุคคล']
    ])
    add2 = pd.MultiIndex.from_product([
        ['ค้าจ้างหรือค่าบำเหน็จในเดือนนี้'],
        ['การรับประกันภัย'],
        ['อุบัติเหตุส่วนบุคคล']
    ])
    columns_3 = pd.MultiIndex.from_product([
        ['ค้าจ้างหรือค่าบำเหน็จในเดือนนี้'],
        ['ประเภทสามัญ', 'ประเภทอุตสาหกรรม', 'ประเภทกลุ่ม'],
        ['ปีแรก', 'ปีต่อไป']
    ])
    columns_4 = pd.MultiIndex.from_product([
        [' '],
        [' '],
        ['ค่าจ้างหรือค่าบำเหน็จรับทั้งสิ้น ในเดือนนี้','ค่าจ้างหรือค่าบำเหน็จค้างรับ ณ วันสิ้นเดือน','เบี้ยประกันภัยค้างนำส่ง ณ วันสิ้นเดือน']
    ])
    joined_columns = blank.append([columns_1,columns_2,add1,columns_3,add2,columns_4])
    return joined_columns

def sheet2():
    # blank = pd.MultiIndex.from_product([[' ']])
    # columns_1 = pd.DataFrame.from_product([[' ','แผนความคุ้มครอง', 'ชื่อ', 'ประเภท', 'ความคุ้มครอง', 'rtu']])
    # joined_columns = blank.append([columns_1])
    data = ['ประเภท','แผนความคุ้มครอง', 'ชื่อ', 'ความคุ้มครอง', 'เบี้ยประกันเพิ่มต่อเนื่อง']
    return data


def sheet3():
    blank = pd.MultiIndex.from_product([
        [' ']
    ])
    columns_1 = pd.MultiIndex.from_product([
        ['แผนประกัน','เบี้ยประกัน','ความคุ้มครอง']
    ])
    joined_columns = blank.append([columns_1])
    return joined_columns