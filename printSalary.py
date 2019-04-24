import config
import yaml
import os
from xlrd import open_workbook


BASE_PATH = os.path.dirname(os.path.abspath(__file__))
DATA_PATH = os.path.join(BASE_PATH, 'DATA')
WORK_DAYS = 21.75
PRT_WGFILE = os.path.join(DATA_PATH, 'WAGECAL_INFO.txt')


class ExcelReader:

    def __init__(self, excel, sheet = 0, title_line = True):
        if os.path.exists(DATA_PATH+os.path.sep+excel.replace('.py', '.xlsx')):
            self.excel = DATA_PATH+os.path.sep+excel.replace('.py', '.xlsx')
            print(DATA_PATH+os.path.sep+excel.replace('.py', '.xlsx'))
        elif os.path.exists(excel):
            self.excel = excel
            #print(DATA_PATH)
        else:
            raise FileNotFoundError('文件不存在!')
        self.sheet = sheet
        self.title_line = title_line
        self._data = list()

    @property
    def data(self):
        if not self._data:
            workbook = open_workbook(self.excel)
            if type(self.sheet) not in  [int, str]:
                raise SheetTypeError('Please pass in <type int> or <type str>, not {0}'.format(type(self.sheet)))
            elif type(self.sheet) == int:
                s = workbook.sheet_by_index(self.sheet)
            else:
                s = workbook.sheet_by_name(self.sheet)

            if self.title_line:
                 title = s.row_values(0)
                 for row in range(1, s.nrows):
                     self._data.append(dict(zip(title, s.row_values(row))))
            else:
                for row in range(0, s.nrows):
                    self._data.append(s.row_values(row))
        return self._data

class Wage_Calculation:
    def __init__(self, sal_data):
        if not sal_data:
            print('Not find any salary data, please input first!')
            return()
        self.job_num = sal_data['job_num']
        self.name = sal_data['name']
        self.base_pay = sal_data['base_pay']
        self.overtime_days = sal_data['overtime_days']
        self.sick_days = sal_data['sick_days']
        self.leave_days = sal_data['leave_days']
        self.merit_pay = sal_data['merit_pay']
        self.sal_data = sal_data

    def wage_cal(self):
        prt_wage = ('员工编号', '员工姓名', '基本工资',  '加班天数', '病假天数', '事假天数', '绩效奖金', '应付工资')

        gross_pay = self.base_pay-self.sick_days*(0.5)*(self.base_pay/WORK_DAYS)-self.leave_days*(self.base_pay/WORK_DAYS)+(self.base_pay/WORK_DAYS)*self.overtime_days*2+self.merit_pay
        gross_pay = round(gross_pay*100)/100
        #print(gross_pay)
        prt_item = ''
        for item in prt_wage:
            prt_item = prt_item + ('{:^16s}'.format(item)+'\t')
        prt_val = ''
        for value in self.sal_data.values():
            if type(value) in [float]:
                prt_val = prt_val + str(('{:^20.2f}'.format(value)+'\t'))
            else:
                prt_val = prt_val + ('{:^20s}'.format(value)+'\t')
        prt_val = prt_val + str(('{:^20.2f}'.format(gross_pay)+'\t'))     

        wagecal_str = prt_item+'\n'+prt_val
        return(wagecal_str)

def save_salary(wrtfile_path, wrtinfo):
    if(os.path.exists(wrtfile_path)):
        os.remove(wrtfile_path)
    wf = open(wrtfile_path, 'w')
    if(wrtinfo != ''):
        try:
            wf.write(wrtinfo)
        except IOError as res:
            print(wrtfile_path+' 文件写入错误' +res)
                    
    wf.close()
    
if __name__=='__main__':
    
    excelrd = ExcelReader(os.path.basename(os.path.realpath(__file__)), title_line = True)
    emplayee_nums = len(excelrd.data)
    save_wageinfo = ''

    for i in range(0, emplayee_nums):
        save_wageinfo =save_wageinfo + Wage_Calculation(excelrd.data[i]).wage_cal()+'\n'

    save_salary(PRT_WGFILE, save_wageinfo)

    
    
