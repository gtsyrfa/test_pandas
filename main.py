import pandas as pd
import time


def save_to_exc(df, filename):
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    df.to_excel(writer, startrow=0, header=True)
    writer.save()


def main():
    df_o = pd.read_excel("orders.xlsx", parse_dates=True, index="OrderId")
    df_ol = pd.read_excel("order_lines.xlsx")
    df_new = pd.merge(df_o, df_ol)
    # Просто магия, даже не пришлось указывать,
    # по какому ключу объединять (машины скоро победят)
    # print(df_new.info())
    grouped = df_new.groupby(["ProductId"])
    # Считаем количество сгруппированных элементов и сортируем
    # в контексте задачи вместо "OrderId" можно взять любое поле.
    results = grouped["OrderId"].count()
    # переименовываем поле для дальшнейшего удобства
    results.name = "Count"
    results = results.sort_values(ascending=False)
    # склоадываем Series "results" с DF "grouped["Price"].sum()"
    results = pd.merge(
                          results,
                          grouped["Price"].sum(),
                          left_index=True,
                          right_index=True
                        )
    # Добавляем поле
    results["avg_price"] = results["Price"]/results["Count"]
    # Сохраняем результат в экселевский файл
    save_to_exc(results, "resultfile.xlsx")

if __name__ == "__main__":
    start_time = time.time()
    main()
    print(time.time() - start_time)
