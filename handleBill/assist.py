
# coding=utf-8

# 此程序输出来源于企查查的信息，包括：身份信息，股东信息，变更记录信息
# 以json格式输出

import docx
import re
import yaml
import os
import time

from handleBill import docx_enhanced

current_path = os.path.dirname(os.path.realpath(__file__))
f = open(current_path+"\\config.yml", encoding="utf-8")
config = yaml.load(f, Loader=yaml.FullLoader)
f.close()

def getTick(timeStr):
  timeArray = time.strptime(timeStr, "%Y-%m-%d")
  return int(time.mktime(timeArray))

def handelArray(arr):
  ret = [[item[3],item[4],item[2],item[6]] for item in arr];
  # 结算列表的长度
  n = len(ret)
  # 外层循环控制从头走到尾的次数
  for j in range(n - 1):
    # 用一个count记录一共交换的次数，可以排除已经是排好的序列
    count = 0
    # 内层循环控制走一次的过程
    for i in range(0, n - 1 - j):
      # 如果前一个元素大于后一个元素，则交换两个元素（升序）
      if getTick(ret[i][0]) > getTick(ret[i + 1][0]) or (getTick(ret[i][0]) == getTick(ret[i + 1][0]) and ret[i + 1][2] == config['partner']):
        # 交换元素
        ret[i], ret[i + 1] = ret[i + 1], ret[i]
        # 记录交换的次数
        count += 1
    # count == 0 代表没有交换，序列已经有序
    if 0 == count:
      break
  pre = ret[0]
  for j in range(1, n):
    if ret[j][0] == pre[0]:
      ret[j][0] = ''
      continue
    pre = ret[j]

  return ret

class CustomError(Exception):
  def __init__(self,ErrorInfo):
    super().__init__(self) #初始化父类
    self.errorinfo=ErrorInfo
  def __str__(self):
    return self.errorinfo