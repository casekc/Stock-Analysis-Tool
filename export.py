
import pandas as pd
import os

net_income_2d_list = [['Net Income', '$59,972,000', '$76,033,000', '$40,269,000', '$34,343,000'], ['Net Income', '-$2,722,000', '$33,364,000', '$21,331,000', '$11,588,000'], ['Net Income', '$12,556,000', '$5,519,000', '$721,000', '-$862,000']]
ticker_list = ['goog', 'amzn', 'tsla']

for i in net_income_2d_list:
    del i[0]
print(net_income_2d_list)
