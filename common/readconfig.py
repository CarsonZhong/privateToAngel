import configparser

class ReadConfig:
    """定义一个读取配置文件的类"""

    def __init__(self):
        self.cf = configparser.ConfigParser()
        self.cf.read("config.ini")


    def get_map(self, param):
        value = self.cf.get("map info", param)
        return value

    def get_style(self, param):
        value = self.cf.get("style info", param)
        return value

    def get_srcxls(self, param):
        value = self.cf.get("srcxls info", param)
        return value

if __name__ == '__main__':
    test = ReadConfig()

    mapxlsname = test.get_map("mapxlsname")
    nameclos = test.get_map("nameclos")
    mapidclos = test.get_map("mapidclos")

    stylexlsname = test.get_style("stylexlsname")
    dataRow = test.get_style("dataRow")
    ColStart = test.get_style("ColStart")

    srcxlsname = test.get_srcxls("srcxlsname")
    srcRowStart = test.get_srcxls("srcRowStart")
    endrowpara = test.get_srcxls("endrowpara")
    
    print(endrowpara)