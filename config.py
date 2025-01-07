import configparser as config
import json
import os

from loguru import logger

PATH = 'config.ini'

# DPI列表
dpi_list = ['1', '1.25', '1.5', '1.75', '25', '2.5', '3']
dpi_list_text = ['100%', '125%', '150%', '175%', '200%', '250%', '300%']

conf = config.ConfigParser()


# 读取config
def read_conf(section='General', key=''):
    data = config.ConfigParser()
    try:
        with open(PATH, 'r', encoding='utf-8') as configfile:
            data.read_file(configfile)
    except FileNotFoundError:
        return None
    except Exception:
        return None

    if section in data and key in data[section]:
        return data[section][key]
    else:
        return None


# 写入config
def write_conf(section, key, value):
    data = config.ConfigParser()
    try:
        with open(PATH, 'r', encoding='utf-8') as configfile:
            data.read_file(configfile)
    except FileNotFoundError:
        pass
    except Exception:
        print(Exception)

    if section not in data:
        data.add_section(section)

    data.set(section, key, value)

    with open(PATH, 'w', encoding='utf-8') as configfile:
        data.write(configfile)


def check_config():  # 确认配置
    conf = config.ConfigParser()
    with open(f'config/default_conf.json', 'r', encoding='utf-8') as file:  # 加载默认配置
        default_conf = json.load(file)

    if not os.path.exists('config.ini'):  # 如果配置文件不存在，则copy默认配置文件
        conf.read_dict(default_conf)
        with open(PATH, 'w', encoding='utf-8') as configfile:
            conf.write(configfile)
        logger.info("配置文件不存在，已创建并写入默认配置。")
    else:
        with open(PATH, 'r', encoding='utf-8') as configfile:
            conf.read_file(configfile)

        if conf['Misc']['ver'] != default_conf['Misc']['ver']:  # 如果配置文件版本不同，则更新配置文件
            logger.info(f"配置文件版本不同，将重新适配")
            try:
                for section, options in default_conf.items():
                    if section not in conf:
                        conf[section] = options
                    else:
                        for key, value in options.items():
                            if key not in conf[section]:
                                conf[section][key] = str(value)
                conf.set('Misc', 'ver', default_conf['Misc']['ver'])
                with open(PATH, 'w', encoding='utf-8') as configfile:
                    conf.write(configfile)
                logger.info(f"配置文件已更新")
            except Exception as e:
                logger.error(f"配置文件更新失败: {e}")


check_config()


if __name__ == '__main__':
    print(read_conf('General', 'dpi'))