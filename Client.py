import pandas as pd
import openpyxl


class Client ():
    def __init__(self):
        self.wb = openpyxl.load_workbook(r'WorkCycle.xlsx')
        self.ws_names = self.wb.sheetnames
        self.ws_names_to_show = list()
        for i in range(len(self.ws_names)):
            if (i % 3) == 0 :
                self.ws_names_to_show.append(self.ws_names[i])
        self.Main_Page()

    def Main_Page(self):
        id_to_task = int(input('하고 싶은 작업을 선택하여 숫자로 입력해 주세요 : 0-종료, 1-예정 사역 확인, 2-사역 면제 쿠폰 사용\n'))
        while (id_to_task!=0 and id_to_task!=1 and id_to_task!=2):
            id_to_task = int(input('입력 에러! 숫자를 1 또는 2로 입력해 주세요 : 0-종료, 1-예정 사역 확인, 2-사역 면제 쿠폰 사용\n'))
        if id_to_task ==1 :
            self.Confirm()
        elif id_to_task==2 :
            self.Request()
        else :
            return 0;


    def Confirm(self):
        print("사역 목록 입니다 : ", end='')
        len_wsts=len(self.ws_names_to_show)
        for i in range(len_wsts):
            print(i+1,end='-')
            if i == len_wsts:
                print(self.ws_names_to_show[i])
            else:
                print(self.ws_names_to_show[i],end=' ')
        print()
        id_to_search = int(input('확인 하고 싶은 사역의 번호를 적어주세요 : '))
        while (id_to_search<1 or id_to_search>len_wsts):
            id_to_search = int(input('입력 에러! 범위에 맞는 숫자 값을 입력해주세요 : '))
        print(pd.read_excel('WorkCycle.xlsx',sheet_name=str(self.ws_names_to_show[id_to_search-1]),usecols=[0,2]))
        self.Main_Page()

    def Request(self):
        df = pd.read_excel('WorkCycle.xlsx', header=None, index_col=None)
        print(df)
        self.Main_Page()

def main():
    Client()
if __name__ =="__main__":
    main()