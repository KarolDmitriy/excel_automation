import pandas as pd
import matplotlib.pyplot as plt
from pathlib import Path


def merge_excel_files(folder_path="data", output_file="summary.xlsx"):
    # Список всех Excel-файлов
    files = list(Path(folder_path).glob("*.xlsx"))
    if not files:
        print("В папке 'data' нет Excel-файлов.")
        return

    all_data = []
    for file in files:
        df = pd.read_excel(file)
        df['Файл'] = file.name  # отметим, откуда данные
        all_data.append(df)

    # Объединяем все таблицы
    combined = pd.concat(all_data, ignore_index=True)
    combined.to_excel(output_file, index=False)
    print(f"Файлы объединены и сохранены в {output_file}")

    # Если есть колонка "Сумма" — строим график
    if "Сумма" in combined.columns:
        plt.figure(figsize=(8, 4))
        combined.groupby("Файл")["Сумма"].sum().plot(kind="bar")
        plt.title("Сумма по файлам")
        plt.ylabel("₽")
        plt.tight_layout()
        plt.savefig("summary_chart.png")
        print("График сохранён как summary_chart.png")


if __name__ == "__main__":
    merge_excel_files()
