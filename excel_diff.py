# -*- coding: utf-8 -*-
"""
Created on Sun Sep 20 18:42:44 2020

@author: Cihat Özeray

Iki excel dosyasından belirtilen sheetlerdeki (default 0) değerlerin arada
"/" olacak şekilde bir birine eklenmesi (ortak olanlar hariç) ve elde edilen
data'nın yeni bir excel dosyası olarak kaydedilmesi

Uyarı: Içınde eleman olmayan bir sheet istendiğinde hata vermekte

cmd örneği:
python excel_diff.py -file1 file1.xlsx -sheet1 Sheet2 -file2 file2.xlsx

"""

import sys
import pandas as pd
import numpy as np


def read_files(file1, sheet1, file2, sheet2):
    """
    içinde bulunduran bir liste oluşturur
    """
    assert isinstance(file1, str)
    assert isinstance(file2, str)
    df1 = pd.read_excel(file1, sheet_name=sheet1, keep_default_na=False, dtype=(str), header=None)
    df2 = pd.read_excel(file2, sheet_name=sheet2, keep_default_na=False, dtype=(str), header=None)
    df_list = [df1, df2]
    return df_list


def process(df_list):
    """
    Ilk dosyanın elemanları ile ikinci dosyanın elemanları arasına "/" ekler
    her ikisinde de bulunan elemanlara eklemez!
    """

    def shape_union(arr1, arr2):
        """
        Herhangi bir dosyadan eleman bulunduran alanın şeklini hesaplar
        """
        shape1 = arr1.shape
        shape2 = arr2.shape
        max_rows = max(shape1[0], shape2[0])
        max_columns = max(shape1[1], shape2[1])
        union_shape = (max_rows, max_columns)
        return union_shape


    # data framelerin ayrılması
    df1 = df_list[0]
    df2 = df_list[1]
    # dataframeler arasında karşılaştırma yapılabilmesi için aynı boyuta getirilmeleri
    # birleşimlerinin boyutunda tahta oluşturulması
    empty_array = np.full(shape_union(df1, df2), "")
    df_board = pd.DataFrame(empty_array)
    # dataframe'lerin tahta boyutuna uyarlanmaları
    df1_boarded = df1.add(df_board, fill_value="")
    df2_boarded = df2.add(df_board, fill_value="")
    # 2.dosyada bulunup ilkinde bulunmayan hücrelerin ayrılması
    df_only2 = df2_boarded[df2_boarded != df1_boarded]
    df_only2.fillna("", inplace=True)
    # her iki dosyada da aynı olan hücrelerin ayrılması
    df_same = df1_boarded[df1_boarded == df2_boarded]
    df_same.fillna("", inplace=True)

    #  araya "/" eklenebilmesi için yeni bir dataframe oluşturulması
    notation_array = np.full(shape_union(df1, df2), "/")
    df_notation = pd.DataFrame(notation_array)
    # notation iki dosyada da aynı olan hücrelere eklenmeyeceği için filtreleme uygulaması
    df_notation = df_notation[df_same == ""]
    df_notation.fillna("", inplace=True)

    # ilk dosyanın, kesişim kümesinde olmayan hücrelerine notasyon eklenmesi
    temp = df1_boarded.add(df_notation)
    # notasyonu eklenmiş ilk dosyaya sadece ikinci dosyada bulunan elemanların eklenmesi
    df_total = temp.add(df_only2)
    # boşluklara gereksiz eklenmiş olan notasyonların silinmesi
    df_total = df_total.replace("/", "")
    return df_total


def style_dataframe(df_to_style):
    """
    Girilen dataFrame'i Dataframe.style kullanarak style eder
    """
    def color_cells(val):
        """
        Change backround color of cells
        """
        if val.find("/") != -1:
            if val[0] == "/":
                return "background-color: yellow"
            if val[-1] == "/":
                return "background-color: orange"
            return "background-color: red"
        return "background-color: white"
    return df_to_style.style.applymap(color_cells)


def write_excel(df_to_write, output_file_name, engine='openpyxl'):
    """
    Dataframe'i, çalışma konumunda bir excel dosyası olarak kaydeder
    """
    assert isinstance(output_file_name, str)
    df_to_write.to_excel(output_file_name, engine=engine)


def get_file_name_sheet_name_parameters():
    """
    Dosya adı parametrelerini sys.argv'dan okuyup dictionary'ye kaydeder
    """
    settings = {}

    param = "-file1"
    param_index = sys.argv.index(param)
    file1 = sys.argv[param_index + 1]
    settings["file1"] = file1

    param = "-file2"
    param_index = sys.argv.index(param)
    file1 = sys.argv[param_index + 1]
    settings["file2"] = file1

    param = "-sheet1"
    if param in sys.argv:
        param_index = sys.argv.index(param)
        sheet1 = sys.argv[param_index + 1]
    else:
        sheet1 = 0
    settings["sheet1"] = sheet1

    param = "-sheet2"
    if param in sys.argv:
        param_index = sys.argv.index(param)
        sheet2 = sys.argv[param_index + 1]
    else:
        sheet2 = 0
    settings["sheet2"] = sheet2

    return settings


def main():
    """
    dosya parametrelerini alır
    dosyaları dataframe olarak okur
    dataframe'ler üzerinde istenen işlemleri gerçekleştirir (Part:1)
    dataframe'stillerini düzenler (Part:2)
    dataframe'in düzenlenmiş halini yeni bir dosya olarak kaydeder
    """
    parameters = get_file_name_sheet_name_parameters()
    df_list = read_files(parameters["file1"], parameters["sheet1"],\
                         parameters["file2"], parameters["sheet2"])
    df_out = process(df_list)
    df_out = style_dataframe(df_out)
    write_excel(df_out, "file.out.xlsx")


if __name__ == "__main__":
    main()
    