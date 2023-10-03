import re
from decimal import *

def validate_form(form_data):

    invalid_fields = []
    bol_return = True

    # 資料規則
    for key, value in form_data.items():
        # 表單號碼
        if key in ['order_number','tvm_order_number']:
            if len(value) != 12 or not value.isdigit():
                bol_return = False
                invalid_fields.append(key)
                continue
        # 訂單號碼 自動補0至12碼
        if key in ['production_number','tvm_production_number']:
            if len(value) != 12 and not value.isdigit():
                bol_return = False
                invalid_fields.append(key)
                continue

            form_data[key] = value.zfill(12)
        # 剛種材料名稱
        if key in ['material_name','tvm_material_name'] and len(value) > 8:
            bol_return = False
            invalid_fields.append(key)
            continue
        #數量
        # 验证数量是否为4位数字
        if key in ['quantity', 'tvm_quantity'] and (len(value) > 4 or not value.isdigit()):
            bol_return = False
            invalid_fields.append(key)
            continue
        #工序
        if key in ['pro_no', 'tvm_pro_no'] and (not len(value) == 2 or not value.isdigit()):
            bol_return = False
            invalid_fields.append(key)
            continue
        #熱處理 研磨
        if key in ['z_heat', 'tvm_z_heat', 'grind', 'tvm_grind']:
            if not value:
                value = 0

            if not (-1 <= float(value) <= 1):
                bol_return = False
                invalid_fields.append(key)
                continue
            form_data[key] = format_value(value,0)
        # 倒角
        if key == 'chamfering':
            if not value:
                value = 0
            if not (0 <= float(value) <= 9.9):
                bol_return = False
                invalid_fields.append(key)
                continue

            form_data[key] = value
        # 切削條件 粗切磁力 精切磁力
        if key in ['cut_parameter', 'tvm_cut_parameter', 'tvm_fin_magn', 'tvm_rou_magn']:
            if not value:
                bol_return = False
                invalid_fields.append(key)
                continue
            if  not (1 <= int(value) <= 8):
                bol_return = False
                invalid_fields.append(key)
                continue
        if key in ['material_length','material_width','material_height','finished_length','finished_width','finished_height','tvm_material_length','tvm_material_width','tvm_material_height', 'tvm_single_cut', 'tvm_double_cut']:
            if not value:
                if key == 'tvm_double_cut':
                    value = 0
                else:
                    bol_return = False
                    invalid_fields.append(key)
                    continue

            value = format_value(value)
            form_data[key] = value

        if key in ['bale_check', 'packet_check'] and value == True:
            form_data[key+'box'] = "☑"

    if bol_return == True:
        form_data['fw'] = calc_data(form_data['finished_width'],form_data['z_heat'],form_data['grind'])
        form_data['fl'] = calc_data(form_data['finished_length'],form_data['z_heat'],form_data['grind'])
        form_data['fh_single'] = calc_data(form_data['tvm_single_cut'],form_data['tvm_z_heat'],form_data['tvm_grind'])
        form_data['fh_double'] = calc_data(form_data['tvm_double_cut'],form_data['tvm_z_heat'],form_data['tvm_grind'])

    return bol_return, invalid_fields, form_data

def format_value(value, type=1):
    # 小數點補零至所需格式
    if not value:
        return value
    try:
        value = format(Decimal(value), '.3f')
        if '.' not in value:
            value += '.000'
        if type == 1:
            value = value.rjust(8, '0')

        return value
    except ValueError:
        return False

def calc_data(value, z_heat, grind):
    result = Decimal(value) + Decimal(z_heat) + Decimal(grind)
    return format_value(result.quantize(Decimal('0.000')))