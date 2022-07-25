import xlwings as xw

import cv2
import numpy as np
import os

wb = xw.Book('path to excel file.xlsx (to file)')
sheet = wb.sheets['Sheet1']

All_col = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA', 'AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT','AU','AV','AW','AX','AY','AZ','BA','BB','BC','BD','BE','BF','BG','BH','BI','BJ','BK','BL','BM','BN','BO','BP','BQ','BR','BS','BT','BU','BV','BW','BX','BY','BZ','CA','CB','CC','CD','CE','CF','CG','CH']

path = 'path to folder containing images (to folder not to file)'

for f in os.listdir(path):
    img = cv2.imread(path + '\\' + f)
    img = cv2.resize(img, (31, 23), interpolation = cv2.INTER_AREA)
    
    ny = 1
    for c in img:
        nx = 0
        for xc in c:
            col = str(All_col[nx]) + str(ny)
            sheet.range(col).color = xc
            nx += 1
        ny += 1

