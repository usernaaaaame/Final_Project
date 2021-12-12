import openpyxl
import hashlib
class Manager ():
    def __init__(self):
        self.wb = openpyxl.load_workbook(r'WorkCycle.xlsx')
        self.wbpass= openpyxl.load_workbook(r'pass.xlsx')
        self.ws_names = self.wb.sheetnames
        self.ws_names_to_show = list()
        for i in range(len(self.ws_names)):
            if (i % 2) == 0:
                self.ws_names_to_show.append(self.ws_names[i])
        print('비밀번호를 입력해주세요 :',end ='')
        password  = input(" ")
        Hashed_password = self.Hash_String(password)
        self.SecurityCheck(Hashed_password)

    def Hash_String(self, password : str):
        password = password.encode("UTF-8")
        MD5=hashlib.md5()
        MD5.update(password)
        Enctext=MD5.hexdigest()
        Enctext=Enctext.upper()
        return Enctext

    def SecurityCheck(self, password : str):
        cnt=1
        self.pass_list=list()
        self.passSheet = self.wbpass['Sheet1']
        for i in self.passSheet.rows:
            self.pass_list.append(i[0].value)

        while self.Find_First_Index(self.pass_list,password)==None:
            if(cnt<5):
                print('비밀번호를 다시 확인해주세요. 5회 오류시 프로그램이 종료됩니다.\n 현재 오류 횟수 :' + str(cnt),end='회')
                password = self.Hash_String(input('\n'))
                cnt = cnt+1
            else:
                print('5회 오류! 프로그램이 종료됩니다.')
                return 0
        print('접속완료!')
        self.Main_Page()


    def Main_Page(self):
        id_to_task = int(input('하고 싶은 작업을 선택하여 숫자로 입력해 주세요 : 0-종료, 1-예정 사역 확인, 2-면제 요청 확인, 3-관리자 추가\n'))
        while (id_to_task <0 or id_to_task >3):
            id_to_task = int(input('입력 에러! 0에서 3사이의 값을 입력해 주세요 : 0-종료, 1-예정 사역 확인, 2-면제 요청 확인, 3-관리자 추가\n'))
        if id_to_task == 1:
            self.Check_Cycle()
        elif id_to_task == 2:
            self.Check_Request()
        elif id_to_task == 3:
            self.Add_manager()
        elif id_to_task == 0:
            return 0

    # def Check_Cycle(self):
    #
    # def Check_Request(self):

    def Add_manager(self):
        print("새로 추가될 관리자가 사용할 비밀번호를 입력해주세요 :", end='')
        password=input()
        password=self.Hash_String(password)
        while (self.Find_First_Index(self.pass_list,password)!=None):
            print('이미 사용하고 있는 비밀번호 입니다. 다른 비밀번호를 입력해주세요.',end=' ')
            password = self.Hash_String(input('만약 메인 화면으로 돌아가고 싶다면 main을 입력해주세요\n'))
            if(password=='FAD58DE7366495DB4650CFEFAC2FCD61'):
                self.Main_Page()

        row_to_input = len(self.pass_list)
        if ((self.Find_First_Index(self.pass_list,None)!=None) and (row_to_input >self.Find_First_Index(self.pass_list,None))):
            row_to_input=self.Find_First_Index(self.pass_list,None)
        self.passSheet.cell(row=row_to_input+1,column=1).value=password
        self.wbpass.save('pass.xlsx')
        print("새로운 관리자로 접속하려면 프로그램 종료 후 다시 실행해주세요!")
        next_work = int(input('다음 작업을 선택해주세요 : 0-종료, 1-메인화면 \n'))
        while (next_work < 0 or next_work > 1):
            next_work = int(input('입력 에러! 범위에 맞는 숫자 값을 입력해주세요 \n'))
        if (next_work == 1):
            self.Main_Page()
        else:
            return 0


    def Find_First_Index(self, list1: list, val):  # 리스트 중 특정 값 갖는 배열의 인덱스 반환
        for i in range(len(list1)):
            if list1[i] == val:
                return i


def main():
    Manager()
if __name__ =="__main__":
    main()