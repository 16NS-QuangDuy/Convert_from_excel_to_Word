import pandas as pd

def birth(df, name):
    result = df[df["Tên"] == name]

    # Kiểm tra nếu tìm thấy kết quả
    if not result.empty:
        year_of_birth = result["Ngày sinh"].values[0]
        return year_of_birth
    else:
        return ""

def death(df, name):
    result = df[df["Tên"] == name]

    # Kiểm tra nếu tìm thấy kết quả
    if not result.empty:
        death_of_birth = result["Ngày giỗ"].values[0]
        return death_of_birth
    else:
        return ""

def place_of_residence (df, name):
    result = df[df["Tên"] == name]

    # Kiểm tra nếu tìm thấy kết quả
    if not result.empty:
        death_of_birth = result["Thường trú"].values[0]
        return death_of_birth
    else:
        return ""

def partner (df, row):
    if row["Giới tính"] == "Nam":  # Nếu là nam, tìm tất cả vợ
        husband_name = row["Tên"]
        wives = df[df["Tên chồng"] == husband_name]
        return wives
    elif row["Giới tính"] == "Nữ":
        husbands_name = row["Tên chồng"]
        husbands = df[df["Tên"] == husbands_name]
        return husbands

def all_child (df_down, row):
    if row["Giới tính"] == "Nam":
        father_name = row["Tên"]
        children = df_down[df_down["Tên bố"] == father_name]["Tên"]
        return children
    else:
        mother_name = row["Tên"]
        children = df_down[df_down["Tên mẹ"] == mother_name]["Tên"]
        return children

def count_male (df_down, row):
    if row["Giới tính"] == "Nam":
        father_name = row["Tên"]
        children = df_down[df_down["Tên bố"] == father_name]
    else:
        mother_name = row["Tên"]
        children = df_down[df_down["Tên mẹ"] == mother_name]
    return children[children["Giới tính"] == "Nam"].shape[0]