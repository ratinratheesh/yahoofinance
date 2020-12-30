from calendar import monthrange
from tkinter import *
import pandas as pd
import requests
import io,openpyxl


def getting_date_range(year, month):
    dates = []
    for day in range(1, monthrange(year, month)[1] + 1):
        date = f"{year:04d}-{month:02d}-{day:02d}"
        dates.append(date)
    return dates


def download_file(range_val):

    req_url = "https://query1.finance.yahoo.com/v7/finance/download/%5EBSESN"
    parameters = {
        'events' : "history",
        'includeAdjustedClose': True,
        'interval': '1d',
        'range': range_val
    }
    resp = requests.get(req_url, params= parameters).content
    df = pd.read_csv(io.StringIO(resp.decode('utf-8')))

    return df

def create_excel():
    range_list = ['1d', '5d', '3mo', 'ytd']
    writer = pd.ExcelWriter('yahoo_bsesn.xlsx', engine='xlsxwriter')
    for each_range in range_list:
        df = download_file(each_range)
        df.to_excel(writer, sheet_name=each_range,index=False)

    writer.save()


def filter_data():
    month_dict = {'1': 'Janauary', '2': 'February', '3': 'March', '4': 'April', '5': 'May', '6': 'June', '7': 'July',
                  '8': 'August', '9': 'September', '10': 'October', '11': 'November', '12': 'December'}
    df = pd.read_excel('yahoo_bsesn.xlsx',sheet_name='ytd',engine='openpyxl')
    mon = month.get()
    month_list = mon.split(",")
    writer = pd.ExcelWriter('Filter.xlsx', engine='xlsxwriter')
    for each_month in month_list:
        dates = getting_date_range(2020, int(each_month))
        data_filter = df.loc[(df['Date'].ge(dates[0])) & (df['Date'].le(dates[-1]))]
        data_filter.to_excel(writer, sheet_name=month_dict[each_month],index=False)
    writer.save()
    top.destroy()


top = Tk()
top.geometry("350x250")
month_label = Label(top, text="Enter Month For Filter", font='14').place(x=90, y=50)
month = StringVar(top)
month_entry = Entry(top, textvariable=month).place(x=100, y=100)
Button(top, text="Submit", activebackground="pink", activeforeground="blue", command=filter_data).place(x=140, y=150)

top.mainloop()
