import xlrd
import math

import matplotlib.pyplot as plt
import matplotlib.patches as patches
import matplotlib.ticker as tick # 目盛り操作に必要なライブラリを読み込みます

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
    #シリアル番号を取得する
    sheet = wb.sheet_by_name("HEADER")
    GP_serial = sheet.cell_value(33,10)

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
    #一括パラメータ設定
    #plt.rcParams["font.family"] ="sans-serif"#使用するフォント
    #plt.rcParams["xtick.direction"] = "in"#x軸の目盛線が内向き("in")か外向き("out")か双方向か("inout")
    #plt.rcParams["ytick.direction"] = "in"#y軸の目盛線が内向き("in")か外向き("out")か双方向か("inout")
    #plt.rcParams["xtick.major.width"] = 1.0#x軸主目盛り線の線幅
    #plt.rcParams["ytick.major.width"] = 1.0#y軸主目盛り線の線幅
    plt.rcParams["font.size"] = 10 #フォントの大きさ
    #plt.rcParams["axes.linewidth"] = 1.0# 軸の線幅edge linewidth。囲みの太さ

    #軸のラベルを作成する
    plt.title(PP_data[0]+"\nCOMPARISON TO DESIGN POSITION\n 1-10")
    plt.xlabel("X-axis")
    plt.ylabel("Y-axis")

    #描画枠を決定する（X軸方向、Y軸方向）
    #"equal"は縦横軸を同じにする
    plt.axis([-8, 8, -8, 8])
    plt.axes().set_aspect("equal")

    #目盛線を描画
    plt.grid(which="major",color="black", linewidth = 1)
    #補助目盛線を描画
    #0.5刻みにで小目盛り(minor locator)表示
    plt.gca().xaxis.set_minor_locator(tick.MultipleLocator(0.5))
    plt.gca().yaxis.set_minor_locator(tick.MultipleLocator(0.5))
    plt.grid(which="minor")

    #円を描画
    c = patches.Circle(xy=(0, 0), radius=5, ec="red", fill=False, linestyle="solid", linewidth = 2)
    plt.axes().add_patch(c)
    c = patches.Circle(xy=(0, 0), radius=2.5, ec="orange", fill=False, linestyle="dashed", linewidth = 2)
    plt.axes().add_patch(c)

    for row in PP_data[1:]:
        #ラインとドット、凡例を描画
        if row[1] != [0, 0, 0, 0] and row[2] != [0, 0, 0, 0]:
            plt.plot(row[1], row[2], row[3])
            plt.plot(row[1], row[2], row[4], label=row[0])
            plt.legend()


    #表示する
    plt.show()


def main():
    print("ver 1.0:2018/11/21  akihiro.teramoto@tel.com")
    filename = glob.glob("*.xls")
    for row in filename:
        plot_data = get_position(row)
        plot_position(plot_data)


if __name__ == "__main__":
    main()
