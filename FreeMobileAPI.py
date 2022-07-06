import pycurl
import re

class FreeMobileAPI:
    def __init__(self):
        self.__handler = pycurl.Curl()
        self.__handler.setopt(pycurl.SSL_VERIFYPEER, False)
        self.__handler.setopt(pycurl.SSL_VERIFYHOST, False)
        self.__handler.setopt(pycurl.CUSTOMREQUEST, "POST")
        self.__handler.setopt(pycurl.TIMEOUT, 3600)
        self.__handler.setopt(pycurl.MAXREDIRS, 10)

        self.__isConnected = False


    def __del__(self):
        self.__handler.close()
        self.__cookies = None
        self.__isConnected = False


    @property
    def IsConnected(self):
        return self.__isConnected


    def Login(self, UserId, Password):
        self.__handler.setopt(pycurl.URL, "https://mobile.free.fr/account/")
        self.__handler.setopt(pycurl.POSTFIELDS, "login-ident={userId}&login-pwd={password}&bt-login=1".format(userId=UserId, password=Password))        
        self.__handler.setopt(pycurl.HTTPHEADER, ("Content-type: application/x-www-form-urlencoded", "Cache-control: no-cache"))
        self.__handler.setopt(pycurl.COOKIEJAR, "")

        result = self.__handler.perform_rs()

        if self.__handler.getinfo(pycurl.HTTP_CODE) == 302:
            self.__cookies = self.__handler.getinfo(pycurl.INFO_COOKIELIST)[0]
            self.__isConnected = True

        return self.IsConnected


    def Logout(self):
        self.__handler.setopt(pycurl.URL, "https://mobile.free.fr/account/?logout=user")
        self.__handler.setopt(pycurl.CUSTOMREQUEST, "GET")
        self.__handler.setopt(pycurl.POSTFIELDS, "")        
        self.__handler.setopt(pycurl.COOKIE, self.__cookies)

        result = self.__handler.perform_rs()

        if self.__handler.getinfo(pycurl.HTTP_CODE) == 302:
            self.__isConnected = False
            self.__cookies = None

        return not self.IsConnected

    """
    Enable Filter    
    """
    def EnableFilter(self, Enabled: bool, Direction: int=1):
        """
        *** Enable Filter ***
        :param bool Enabled:
        :param int Direction: (default 1) {0:In, 1:Out}
        """
        self.__handler.setopt(pycurl.URL, "https://mobile.free.fr/account/mes-services/filtres?action=default&dir={direction}&val={value}".format(direction=Direction, value=int(Enabled)))
        self.__handler.setopt(pycurl.CUSTOMREQUEST, "GET")
        self.__handler.setopt(pycurl.COOKIE, self.__cookies)

        result = self.__handler.perform_rs()

        return True if self.__handler.getinfo(pycurl.HTTP_CODE) == 302 else False
        

    def GetFilterIds(self):
        self.__handler.setopt(pycurl.URL, "https://mobile.free.fr/account/mes-services/filtres")
        self.__handler.setopt(pycurl.CUSTOMREQUEST, "GET")
        self.__handler.setopt(pycurl.COOKIE, self.__cookies)

        result = self.__handler.perform_rs()
               
        filterOut = False
        filterIds = {"In":[], "Out":[]}

        if self.__handler.getinfo(pycurl.HTTP_CODE) == 200:
            for id in result.split("\n"):
                if re.search("f-rules__list--out", id):
                    filterOut = True

                if re.search("data-id", id):
                    if filterOut == False:
                        filterIds["In"].append(id.strip().split(" ")[4].split("=")[1][:-1].replace('"', ''))
                    else:
                        filterIds["Out"].append(id.strip().split(" ")[4].split("=")[1][:-1].replace('"', ''))

        return filterIds
    

    def AddFilter(self, Params):       
        self.__handler.setopt(pycurl.URL, "https://mobile.free.fr/account/mes-services/filtres?action=save")
        self.__handler.setopt(pycurl.CUSTOMREQUEST, "POST")        
        self.__handler.setopt(pycurl.POSTFIELDS, Params)
        self.__handler.setopt(pycurl.HTTPHEADER, ("Content-type: application/x-www-form-urlencoded", "Cache-control: no-cache"))
        self.__handler.setopt(pycurl.COOKIE, self.__cookies)

        result = self.__handler.perform_rs()

        return True if self.__handler.getinfo(pycurl.HTTP_CODE) == 302 else False

        
    def DeleteFilter(self, RuleId):
        self.__handler.setopt(pycurl.URL, "https://mobile.free.fr/account/mes-services/filtres?action=delete&id={ruleid}".format(ruleid=RuleId))
        self.__handler.setopt(pycurl.CUSTOMREQUEST, "GET")        
        self.__handler.setopt(pycurl.POSTFIELDS, "")
        self.__handler.setopt(pycurl.COOKIE, self.__cookies)

        result = self.__handler.perform_rs()

        return True if self.__handler.getinfo(pycurl.HTTP_CODE) == 302 else False



