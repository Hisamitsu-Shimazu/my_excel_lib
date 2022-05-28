import openpyxl
import pandas as pd
import os
import re
import yaml

path_to_config = './config/'

class MyExcelLib():
    def __init__(self, file_path, file_name):
        self._file_path = file_path
        self._file_name = file_name

        if os.path.exists(file_path + file_name):
            print(f'load {self._file_path+self._file_name}')
            self._book = openpyxl.load_workbook(self._file_path + self._file_name)
        else:
            print('create new book.')
            self._book = openpyxl.Workbook()

    def book_info(self):
        """ブック情報の確認"""
        print(self._book.sheetnames)

    def save_book(self):
        """ブックの保存"""
        print(f'save {self._file_path+self._file_name}')
        self._book.save(self._file_path + self._file_name)
    
    def set_sheet_name(self, name, index):
        """シート名の変更"""
        self._book.worksheets[index].title = name

    def add_sheet(self, title, index=None):
        """シートの追加
        Args:
            title (str): 追加するシート名
            index (int): 何番目に追加するか、デフォルトは末尾
        """
        self._book.create_sheet(title, index)

    def remove_worksheet(self, title):
        """シートの削除
        Args:
            title (str): 削除するシート名
        """
        self._book.remove(title)

    def set_value(self, value, sheet, position, style='basic'):
        """セルへ値の設定"""
        position = get_cell_position(position)
        cell = self._book[sheet].cell(*position)
        cell.value = value
        self.set_style(sheet, position, style)

    def set_hyperlink_to_cell(self, value, to_sheet, to_position, sheet, position, style='hyperlink'):
        """同ブックセルへのリンク"""
        if type(to_position) == tuple:
            to_position = conv_tuple_str(to_position)
        link = f'{self._file_name}\#{to_sheet}!{to_position}'
        position = get_cell_position(position)
        cell = self._book[sheet].cell(*position)
        cell.value = value
        cell.hyperlink = link
        self.set_style(sheet, position, style)

    def get_value(self, sheet, position):
        """セルの値の取得"""
        position = get_cell_position(position)
        cell = self._book[sheet].cell(*position)
        return cell.value
    
    # 書式
    def set_style(self, sheet, position, style):
        """セルに対してスタイルの適用"""
        position = get_cell_position(position)
        cell = self._book[sheet].cell(*position)

        if type(style) == str:
            style = import_config('style.yml')[style]

        if 'font' in style.keys():
            cell.font = openpyxl.styles.Font(**style['font'])
        if 'alignment' in style.keys():
            cell.alignment = openpyxl.styles.alignment.Alignment(**style['alignment'])
        if 'fill' in style.keys():
            cell.fill = openpyxl.styles.PatternFill(**style['fill'])
        if 'border' in style.keys():
            border_style = {}
            for side in style['border'].keys():
                border_style[side] = openpyxl.styles.Side(**style['border'][side])
            if 'diagonal' in border_style.keys():
                border_style['diagonalDown'] = True
            cell.border = openpyxl.styles.Border(**border_style)
    
    def concat_cells(self, sheet, start_position, end_position):
        """セルの結合"""
        start_position = get_cell_position(start_position)
        end_position = get_cell_position(end_position)
        self._book[sheet].merge_cells(
            start_row=start_position[0], 
            start_column=start_position[1],
            end_row=end_position[0],
            end_column=end_position[1])

    def set_df(self, df, sheet, position):
        """dfの配置"""
        position = get_cell_position(position)
        
        # df左上
        self.set_value('', sheet, position, 'df_head')
        for i, c in enumerate(df.columns):
            self.set_value(c, sheet, (position[0], position[1]+(i+1)), 'df_head')
        for i, idx in enumerate(df.index):
            self.set_value(idx, sheet, (position[0]+i+1, position[1]), 'df_value')
        for r in range(df.shape[0]):
            for c in range(df.shape[1]):
                self.set_value(df.iat[r,c], sheet, (position[0]+1+r, position[1]+1+c), 'df_value')

def import_config(file_name, path_to_config=path_to_config):
    """設定ファイル(yml)を読み込んで返す
    Args:
        file_name(str): 読み込み対象のファイル名
    Return:
        config(dict): 読み込んだ設定値
    """
    with open(path_to_config + file_name, 'r') as yml:
        config = yaml.load(yml,  Loader=yaml.SafeLoader)
    return config

def get_cell_position(position):
    """セル位置を文字列で渡された場合に数値のタプルで返す"""
    if type(position) == str:
        position = conv_str_tuple(position)
    return (position[0], position[1])

def conv_str_tuple(pos_str:str)->tuple:
    """セル'英字+数字'->(行番号, 列番号)
    列最小値が0でないため単純な26進数変換を適用できない"""
    al_num = lambda x: 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.index(x) + 1
    
    row = int(re.sub("\D", "", pos_str))
    col_str = re.sub("\d", "", pos_str)
    col = sum([al_num(s)*(26**i) for i, s in enumerate(col_str[::-1])])
    return (row, col)

def conv_tuple_str(pos_tuple:tuple)->str:
    """セル(行番号, 列番号)->'英字+数字'
    列最小値が0でないため単純な26進数変換を適用できない"""
    row, col = pos_tuple
    li = []
    while True:
        sub_flag = 0
        mod = col % 26
        if mod==0:
            sub_flag=1
            mod = 26
        li.append('ABCDEFGHIJKLMNOPQRSTUVWXYZ'[mod-1])
        col = col // 26 - sub_flag
        if col == 0:
            break
    col_str = ''.join(li[::-1])
    pos_str = col_str + str(row)
    return pos_str