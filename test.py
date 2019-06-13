import win32com.client


def alm_login(username, password, domain, project):
    td.InitConnection(url)
    td.Login(username, password)
    td.Connect(domain, project)
    if td.Connected:
        print("connect successfully")



if __name__ == "__main__":
    url = "http://15.83.240.100/qcbin"
    td = win32com.client.Dispatch("TDApiOle80.TDConnection")
    try:
        alm_login("chen.si", "P@ssw0rd", "TEST", "Test_WES")
    except Exception as e:
        print(e)
    finally:
        td.Disconnect()

