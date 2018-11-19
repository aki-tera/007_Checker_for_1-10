import matplotlib.pyplot as plt
import matplotlib.patches as patches
import matplotlib.ticker as tick # 目盛り操作に必要なライブラリを読み込みます


def plot_position(PP_data):
    """
    データからプロット表示する

    Parameters
    ----------
    PP_data:list
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
    plt.title("COMPARISON TO DESIGN POSITION\n 1-10")
    plt.xlabel("X-axis")
    plt.ylabel("Y-axis")

    #描画枠を決定する（X軸方向、Y軸方向）
    #"equal"は縦横軸を同じにする
    plt.axis([-8, 8, -8, 8])
    plt.axes().set_aspect("equal")

    #目盛線を描画
    plt.grid(which="major",color="black")
    #補助目盛線を描画
    #0.5刻みにで小目盛り(minor locator)表示
    plt.gca().xaxis.set_minor_locator(tick.MultipleLocator(0.5))
    plt.gca().yaxis.set_minor_locator(tick.MultipleLocator(0.5))
    plt.grid(which="minor")

    #円を描画
    c = patches.Circle(xy=(0, 0), radius=5, ec="gray", fill=False, linestyle="dashed", linewidth = 0.5)
    plt.axes().add_patch(c)
    c = patches.Circle(xy=(0, 0), radius=2.5, ec="gray", fill=False, linestyle="dashed", linewidth = 0.5)
    plt.axes().add_patch(c)

    temp = 0
    for row in PP_data:
        #ラインとドット、凡例を描画
        plt.plot(row[1][0], row[1][1], row[2], label=row[0])
        plt.plot(row[1][0], row[1][1], row[3])
        plt.legend()


    #表示する
    plt.show()


def main():
    plot_data = [["1-41", [[4, 0, 5, 1],[1, 2, 3, 4]], "b-", "^b"], ["1-1", [[-1, -2, -3, -4],[-4, 0, 3, 5]], "y-", "^y"]]
    plot_position(plot_data)


if __name__ == "__main__":
    main()
