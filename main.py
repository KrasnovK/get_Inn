import csv

import requests
import pandas as pd

FILE = "result.csv"


def template():
    file_name = "template.xlsx"
    sheet = "Sheet1"
    df = pd.read_excel(io=file_name, sheet_name=sheet)
    return df


def get_request_id(fam, nam, otch, bdate, doctype, docno, docdt=None):
    url = "https://service.nalog.ru/inn-new-proc.do"
    data = {
        "fam": fam,
        "nam": nam,
        "otch": otch,
        "bdate": bdate,
        "doctype": doctype,
        "docno": docno,
        "docdt": docdt,
        "c": "find",
        "captcha": "",
        "captchaToken": "",
    }
    resp = requests.post(url=url, data=data)
    resp.raise_for_status()
    return resp.json()["requestId"]


def get_inn():
    url = "https://service.nalog.ru/inn-new-proc.json"
    df = template()
    requestId_list = []
    for i in df.values:
        print(i)
        data = {
            "c": "get",
            "requestId": get_request_id(
                fam=i[0],
                nam=i[1],
                otch=i[2],
                bdate=i[3],
                doctype=str(i[4]),
                docno=str(i[5]),
            ),
        }
        resp = requests.post(url=url, data=data)
        if "inn" in resp.json():
            correct_inn = resp.json()["inn"]
            requestId_list.append(
                {
                    "fam": i[0],
                    "nam": i[1],
                    "otch": i[2],
                    "bdate": i[3],
                    "doctype": i[4],
                    "docno": i[5],
                    "inn": i[6],
                    "correct_inn": correct_inn,
                    "bench_inn": int(correct_inn) == i[6],
                }
            )
        else:
            requestId_list.append(
                {
                    "fam": i[0],
                    "nam": i[1],
                    "otch": i[2],
                    "bdate": i[3],
                    "doctype": i[4],
                    "docno": i[5],
                    "inn": i[6],
                    "correct_inn": "Кривые паспортные данные",
                    "bench_inn": None,
                }
            )
    return requestId_list


def save_file(data, path):
    with open(path, "w", newline="", encoding="utf-16", errors="ignore") as file:
        writer = csv.writer(file, delimiter=";")
        writer.writerow(
            [
                "Фамилия",
                "Имя",
                "Отчество",
                "Дата рождения",
                "Документ",
                "Номер документа",
                "ИНН из системы",
                "Корректный ИНН",
                "Проверка ИНН",
            ]
        )
        for row in data:
            writer.writerow(
                [
                    row["fam"],
                    row["nam"],
                    row["otch"],
                    row["bdate"],
                    row["doctype"],
                    row["docno"],
                    row["inn"],
                    row["correct_inn"],
                    row["bench_inn"],
                ]
            )


def main():
    price_list = []
    price_list.extend(get_inn())
    save_file(price_list, FILE)
    print("Выполнено")


main()
