import xlrd
import datetime
import math

from matplotlib.pyplot import figure
#from matplotlib.pyplot import show
import matplotlib.patches as patches
import matplotlib.pyplot as plt
import matplotlib.ticker as tick #目盛り操作に必要なライブラリを読み込みます

import glob

def get_position(GP_filename):
    """
    ファイルから所定のデータを出力する

    Parameters
    ----------
    GP_filename:str   ファイル名称

    Returns
    ----------
    GP_result:list
    GP_result[0]        シリアル番号
    GP_result[n][0]     該当レイアウト
    GP_result[n][1][0]  X方向の位置データ×4か所
    GP_result[n][1][1]  Y方向の位置データ×4か所
    GP_result[n][2]     描画するラインの色
    GP_result[n][3]     描画するドットの色

    ----------
    """
    #ファイルを取得する
    wb = xlrd.open_workbook(GP_filename)
    #シリアル番号と日付を取得する
    sheet = wb.sheet_by_name("HEADER")
    GP_serial = sheet.cell_value(33,10)
    temp_1 = sheet.cell_value(35,10)
    temp_2 = datetime.datetime(*xlrd.xldate_as_tuple(temp_1, wb.datemode))
    GP_serial = GP_serial+"@"+str(temp_2.year)+"/"+str(temp_2.month)+"/"+str(temp_2.day)


    #モジュールデータを取得する
    sheet = wb.sheet_by_name("COMPARISON TO DESIGN POSITION")
    #対象モジュールの行位置を取得する
    GP_row_index = 0
    for row in range(sheet.nrows):
        temp = sheet.cell(row, 1).value
        if "BLOCK 1-10" in temp:
            GP_row_index = row +10
            break
    #対象モジュールのデータを取得する
    GP_DEFAULT_NUMBER = 100
    GP_data = [[GP_DEFAULT_NUMBER]*8 for i in range(50)]
    GP_col_index = 0
    GP_module_id = sheet.row_values(GP_row_index)

    #各方向の計算式を設定する
    GP_cal_x = lambda a: math.sin((math.pi*a/165888))*512
    GP_cal_y = lambda a: a*15/2048
    GP_cal_z = lambda a: a/256
    #詳細データを取得する
    for row in GP_module_id[6:]:
        if row == "":
            break
        #モジュール名称
        GP_data[GP_col_index][0] = row
        #モジュールID
        GP_data[GP_col_index][1] = sheet.cell_value(GP_row_index+1,GP_col_index+6)
        #X方向規定値
        GP_data[GP_col_index][2] = GP_cal_x(sheet.cell_value(GP_row_index+2,GP_col_index+6))
        #Y方向規定値
        GP_data[GP_col_index][3] = GP_cal_y(sheet.cell_value(GP_row_index+4,GP_col_index+6))
        #Z方向規定値
        GP_data[GP_col_index][4] = GP_cal_z(sheet.cell_value(GP_row_index+6,GP_col_index+6))
        #X方向結果
        GP_data[GP_col_index][5] = GP_cal_x(sheet.cell_value(GP_row_index+12,GP_col_index+6))
        #Y方向結果
        GP_data[GP_col_index][6] = GP_cal_y(sheet.cell_value(GP_row_index+14,GP_col_index+6))
        #Z方向結果
        GP_data[GP_col_index][7] = GP_cal_z(sheet.cell_value(GP_row_index+16,GP_col_index+6))
        #次のモジュールに移動させる
        GP_col_index = GP_col_index +1

    #結果をまとめる
    #結果のリストを生成
    GP_result = [0]*7
    #シリアル番号を保存
    GP_result[0] = GP_serial
    #まずは高さの場合分け準備
    GP_max_z = max(GP_data, key = lambda x: x[4])[4]
    GP_min_z = min(GP_data, key = lambda x: x[4])[4]
    #結果算出用配列の準備
    GP_temp = [[0]*6 for i in range(6)]
    #該当するレイアウト
    GP_temp[0][0] = "TOP-CSB"
    GP_temp[1][0] = "TOP-EXT"
    GP_temp[2][0] = "LOW-CSB"
    GP_temp[3][0] = "LOW-EXT"
    GP_temp[4][0] = "MID-CSB"
    GP_temp[5][0] = "MID-EXT"
    #ラインの色
    GP_temp[0][3] = "y-"
    GP_temp[1][3] = "y-"
    GP_temp[2][3] = "b-"
    GP_temp[3][3] = "b-"
    GP_temp[4][3] = "g-"
    GP_temp[5][3] = "g-"
    #ドットの色
    GP_temp[0][4] = "^y"
    GP_temp[1][4] = "sy"
    GP_temp[2][4] = "^b"
    GP_temp[3][4] = "sb"
    GP_temp[4][4] = "^g"
    GP_temp[5][4] = "sg"
    #X方向位置の差分
    GP_temp[0][1] = [0, 0, 0, 0]
    GP_temp[1][1] = [0, 0, 0, 0]
    GP_temp[2][1] = [0, 0, 0, 0]
    GP_temp[3][1] = [0, 0, 0, 0]
    GP_temp[4][1] = [0, 0, 0, 0]
    GP_temp[5][1] = [0, 0, 0, 0]
    #Y方向位置の差分
    GP_temp[0][2] = [0, 0, 0, 0]
    GP_temp[1][2] = [0, 0, 0, 0]
    GP_temp[2][2] = [0, 0, 0, 0]
    GP_temp[3][2] = [0, 0, 0, 0]
    GP_temp[4][2] = [0, 0, 0, 0]
    GP_temp[5][2] = [0, 0, 0, 0]
    for row in sorted(GP_data, key = lambda x: x[3]):
        if row[4]==GP_min_z and row[2]>0:
            #TOP-IN側の処理
            #X方向の差分
            GP_temp[0][1][GP_temp[0][5]] = round(row[5]-row[2], 2)
            #Y方向の差分
            GP_temp[0][2][GP_temp[0][5]] = round(row[6]-row[3], 2)
            #カウンタを増やす
            GP_temp[0][5] = GP_temp[0][5]+1
        elif row[4]==GP_min_z and row[2]<0:
            #TOP-EXT側の処理
            #X方向の差分
            GP_temp[1][1][GP_temp[1][5]] = round(row[5]-row[2], 2)
            #Y方向の差分
            GP_temp[1][2][GP_temp[1][5]] = round(row[6]-row[3], 2)
            #カウンタを増やす
            GP_temp[1][5] = GP_temp[1][5]+1
        elif row[4]==GP_max_z and row[2]>0:
            #LOW-CSB側の処理
            #X方向の差分
            GP_temp[2][1][GP_temp[2][5]] = round(row[5]-row[2], 2)
            #Y方向の差分
            GP_temp[2][2][GP_temp[2][5]] = round(row[6]-row[3], 2)
            #カウンタを増やす
            GP_temp[2][5] = GP_temp[2][5]+1
        elif row[4]==GP_max_z and row[2]<0:
            #LOW-EX側の処理
            #X方向の差分
            GP_temp[3][1][GP_temp[3][5]] = round(row[5]-row[2], 2)
            #Y方向の差分
            GP_temp[3][2][GP_temp[3][5]] = round(row[6]-row[3], 2)
            #カウンタを増やす
            GP_temp[3][5] = GP_temp[3][5]+1
        elif row[4]!=GP_DEFAULT_NUMBER and row[2]>0:
            #MID-CSB側の処理
            #X方向の差分
            GP_temp[4][1][GP_temp[4][5]] = round(row[5]-row[2], 2)
            #Y方向の差分
            GP_temp[4][2][GP_temp[4][5]] = round(row[6]-row[3], 2)
            #カウンタを増やす
            GP_temp[4][5] = GP_temp[4][5]+1
        elif row[4]!=GP_DEFAULT_NUMBER and row[2]<0:
            #MID-EX側の処理
            #X方向の差分
            GP_temp[5][1][GP_temp[5][5]] = round(row[5]-row[2], 2)
            #Y方向の差分
            GP_temp[5][2][GP_temp[5][5]] = round(row[6]-row[3], 2)
            #カウンタを増やす
            GP_temp[5][5] = GP_temp[5][5]+1

    GP_result[5] = GP_temp[0][0:5]      #"TOP-CSB"
    GP_result[6] = GP_temp[1][0:5]      #"TOP-EXT"
    GP_result[1] = GP_temp[2][0:5]      #"LOW-CSB"
    GP_result[2] = GP_temp[3][0:5]      #"LOW-EXT"
    GP_result[3] = GP_temp[4][0:5]      #"MID-CSB"
    GP_result[4] = GP_temp[5][0:5]      #"MID-EXT"
    return(GP_result)


def plot_position(PP_data):
    """
    データからプロット表示する

    Parameters
    ----------
    PP_data:list
        PP_data[0]        シリアル番号
        PP_data[n][0]     モージュール名称
        PP_data[n][1][0]  X方向の位置データ×4か所
        PP_data[n][1][1]  Y方向の位置データ×4か所
        PP_data[n][2]     描画するラインの色
        PP_data[n][3]     描画するドットの色
    Returns
    ----------
    """
    #Figureオブジェクトを作成
    fig = figure()
    fig.suptitle("COMPARISON TO DESIGN POSITION @1-10", fontweight="bold")

    #figに属するAxesオブジェクトを作成
    ax = fig.add_subplot(1, 1, 1)
    #ラインとドット、凡例を描画
    for row in PP_data[1:]:
        if row[1] != [0, 0, 0, 0] and row[2] != [0, 0, 0, 0]:
            #ラインを表示
            ax.plot(row[1], row[2], row[3])
            #ドットを表示
            ax.plot(row[1], row[2], row[4], label=row[0])
            #凡例を表示
            ax.legend(loc="best", fontsize=8)

    #円を描画
    c = patches.Circle(xy=(0, 0), radius=5, ec="red", fill=False, linestyle="solid", linewidth = 2)
    ax.add_patch(c)
    c = patches.Circle(xy=(0, 0), radius=2.5, ec="orange", fill=False, linestyle="dashed", linewidth = 2)
    ax.add_patch(c)

    #ラベルの表示
    ax.set_title(PP_data[0])
    ax.set_xlabel("X-axis")
    ax.set_ylabel("Y-axis")
    #目盛を作成
    ax.set_xlim(-8, 8)
    ax.set_ylim(-8, 8)
    #縦横比率を同じにする
    ax.set_aspect(1.0/ax.get_data_ratio())

    #目盛線を描画
    plt.grid(which = "major", color="black", linewidth = 1)
    #補助目盛線を描画
    #0.5刻みにで小目盛り(minor locator)表示
    plt.gca().xaxis.set_minor_locator(tick.MultipleLocator(0.5))
    plt.gca().yaxis.set_minor_locator(tick.MultipleLocator(0.5))
    plt.grid(which = "minor")

    #表示
    fig.show()

def main():
    print("ver 1.2:2018/11/22  akihiro.teramoto@tel.com")
    filename = glob.glob("*.xls")
    print(filename)
    for row in filename:
        if row != []:
            print("プロットするファイル：", row)
            plot_data = get_position(row)
            plot_position(plot_data)
        else:
            print("ファイルがありません")


if __name__ == "__main__":
    main()
