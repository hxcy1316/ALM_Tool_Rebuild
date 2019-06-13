import win32com.client

url = "http://15.83.240.100/qcbin"
username = "chen.si"
password = "P@ssw0rd"
domain = "TEST"
project = "Test_WES"

testData = {}
td = win32com.client.Dispatch("TDApiOle80.TDConnection")
td.InitConnection(url)
td.Login(username, password)
td.Connect(domain, project)
if td.Connected:
    print("connect successfully")
    td.Disconnect()
else:
    print("Not connected")