'''
    PowerPoint Touch Assist v1.1
    读取配置文件
    Author: RinLit_233OuO @bilibili
    Version:  1.1
'''
import configparser as config
path = 'config.ini'

#选项及对应的dpi
dpi_dict = {
    '0':1,
    '1':1.25,
    '2':1.5,
    '3':1.75,
    '4':2
}

conf = config.ConfigParser()
#读取config
def read_conf(Section = 'General', Key = str):
    data = config.ConfigParser()
    try:
        with open(path, 'r', encoding='utf-8') as configfile:
            data.read_file(configfile)
    except FileNotFoundError:
        return None
    except Exception:
        return None

    if Section in data and Key in data[Section]:
        return data[Section][Key]
    else:
        return None
#写入config
def write_conf(section, key, value):
    data = config.ConfigParser()
    try:
        with open(path, 'r', encoding='utf-8') as configfile:
            data.read_file(configfile)
    except FileNotFoundError:
        pass
    except Exception:
        print(Exception)

    if section not in data:
        data.add_section(section)
    
    data.set(section, key, value)

    with open(path, 'w', encoding='utf-8') as configfile:
        data.write(configfile)