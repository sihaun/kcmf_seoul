from copy import copy
import pandas as pd

def perfectcopy(target, source, value=None):
    target.font = copy(source.font)
    target.border = copy(source.border)
    target.fill = copy(source.fill)
    target.number_format = copy(source.number_format)
    target.protection = copy(source.protection)
    target.alignment = copy(source.alignment)

    if value is None:
        target.value = source.value
    else:
        target.value = value

def paste_equipment(manager_name, config_data, ws_source, ws_target, source_pointer, target_pointer):
    register_path = ws_source[f"D{source_pointer}"].value
    month = ws_source[f"E{source_pointer}"].value
    borrow_date = ws_source[f"F{source_pointer}"].value
    borrow_day = ws_source[f"G{source_pointer}"].value
    return_date = ws_source[f"H{source_pointer}"].value
    return_day = ws_source[f"I{source_pointer}"].value
    register_name = ws_source[f"J{source_pointer}"].value
    party_name = ws_source[f"K{source_pointer}"].value
    gender = ws_source[f"L{source_pointer}"].value
    city = ws_source[f"M{source_pointer}"].value
    address = ws_source[f"N{source_pointer}"].value
    age = ws_source[f"O{source_pointer}"].value
    use_object = ws_source[f"P{source_pointer}"].value
    detailed_task = ws_source[f"Q{source_pointer}"].value
    performance_indicator = ws_source[f"R{source_pointer}"].value
    vulnerable = ws_source[f"S{source_pointer}"].value
    equipment_level = ws_source[f"T{source_pointer}"].value
    equipment_group = ws_source[f"U{source_pointer}"].value
    model_name = ws_source[f"V{source_pointer}"].value
    quantity = ws_source[f"W{source_pointer}"].value
    period = ws_source[f"X{source_pointer}"].value
    usage = ws_source[f"Y{source_pointer}"].value

    if performance_indicator == "None":
        performance_indicator = "해당없음"

    if vulnerable == "None":
        performance_indicator = "해당없음"

    group = config_data[model_name]["equipment_group"]
    name = config_data[model_name]["model_name"]

    perfectcopy(ws_target[f"A{target_pointer}"], ws_target[f"A2"], "서울")
    perfectcopy(ws_target[f"B{target_pointer}"], ws_target[f"B2"], manager_name)
    perfectcopy(ws_target[f"C{target_pointer}"], ws_target[f"C2"], register_path)
    perfectcopy(ws_target[f"D{target_pointer}"], ws_target[f"D2"], month)
    perfectcopy(ws_target[f"E{target_pointer}"], ws_target[f"E2"], borrow_date)
    perfectcopy(ws_target[f"F{target_pointer}"], ws_target[f"F2"], borrow_day)
    perfectcopy(ws_target[f"G{target_pointer}"], ws_target[f"G2"], return_date)
    perfectcopy(ws_target[f"H{target_pointer}"], ws_target[f"H2"], return_day)
    perfectcopy(ws_target[f"I{target_pointer}"], ws_target[f"I2"], register_name)
    perfectcopy(ws_target[f"J{target_pointer}"], ws_target[f"J2"], party_name)
    perfectcopy(ws_target[f"K{target_pointer}"], ws_target[f"K2"], gender)
    perfectcopy(ws_target[f"L{target_pointer}"], ws_target[f"L2"], "")
    perfectcopy(ws_target[f"M{target_pointer}"], ws_target[f"M2"], "")
    perfectcopy(ws_target[f"N{target_pointer}"], ws_target[f"N2"], city)
    perfectcopy(ws_target[f"O{target_pointer}"], ws_target[f"O2"], address)
    perfectcopy(ws_target[f"P{target_pointer}"], ws_target[f"P2"], "")
    perfectcopy(ws_target[f"Q{target_pointer}"], ws_target[f"Q2"], "")
    perfectcopy(ws_target[f"R{target_pointer}"], ws_target[f"R2"], f"{age}대")
    perfectcopy(ws_target[f"S{target_pointer}"], ws_target[f"S2"], "제작")
    perfectcopy(ws_target[f"T{target_pointer}"], ws_target[f"T2"], performance_indicator)
    perfectcopy(ws_target[f"U{target_pointer}"], ws_target[f"U2"], vulnerable)
    perfectcopy(ws_target[f"V{target_pointer}"], ws_target[f"V2"], "해당없음")
    perfectcopy(ws_target[f"W{target_pointer}"], ws_target[f"W2"], equipment_level)
    perfectcopy(ws_target[f"X{target_pointer}"], ws_target[f"X2"], group)
    perfectcopy(ws_target[f"Y{target_pointer}"], ws_target[f"Y2"], name)
    perfectcopy(ws_target[f"Z{target_pointer}"], ws_target[f"Z2"], quantity)
    perfectcopy(ws_target[f"AA{target_pointer}"], ws_target[f"AA2"], period)
    perfectcopy(ws_target[f"AB{target_pointer}"], ws_target[f"AB2"], usage)

def xls2xlsx(xls_path, xlsx_path):
    df = pd.read_excel(xls_path, engine="xlrd")
    df.to_excel(xlsx_path, index=False)
