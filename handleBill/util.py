from enum import Enum

class CONFIG_COL(Enum):
  '''配置文件列常量'''
  caseName = 0
  fileName = 1
  handleType = 2
  contractNum = 3
  versionNum = 4
  handleMethod = 5
  lawyerRate = 6
  currency = 7
  specifiedDate = 8

class WORK_COL(Enum):
  '''业务工作时间列常量'''
  customerName = 0
  caseName = 1
  lawyer = 2
  workDate = 3
  content = 4
  workDuration = 5
  billDuration = 6
