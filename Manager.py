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
        self.Security(Hashed_password)

    def Hash_String(self, password : str):
        password = password.encode("UTF-8")
        MD5=hashlib.md5()
        MD5.update(password)
        Enctext=MD5.hexdigest()
        Enctext=Enctext.upper()
        return Enctext

    def Security(self, password : str):
        if
    def MainPage(self):

    def Check_Cycle(self):

    def Manage(self):



def main():
    Manager()
if __name__ =="__main__":
    main()