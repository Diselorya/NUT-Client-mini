import time  
import win32com.client  
import nut2 as nut
import os
import configparser as cp
import ctypes
import pandas as pd
import datetime
import math

# 将当前工作目录设置为脚本所在的路径
# os.chdir(os.path.dirname(os.path.abspath(__file__)))

# # NUT服务器配置  
# nut_host = '192.168.2.10'  
# nut_port = 3493  
# nut_user = 'upsmon'  
# nut_password = 'secret'  

# # UPS设备配置  
# ups_device = 'ups' 
# shutdown_when = 90
  
# # Windows服务器配置  
# server_name = 'server_name'  
# service_name = 'service_name'  

config = cp.ConfigParser()
config_dir = os.path.dirname(os.path.abspath(__file__))
print(config_dir)

# 打包时注掉这句
# os.chdir(config_dir)

config.read('nutconfig.ini')


# NUT服务器配置  
nut_host = config.get('NUT_SERVER', 'host')
nut_port = config.getint('NUT_SERVER', 'port')
nut_user = config.get('NUT_SERVER', 'user')
nut_password = config.get('NUT_SERVER', 'password')
  
# UPS设备配置  
ups_device = config.get('UPS_DEVICE', 'device')

# 关机设置
shutdown_when = config.getint('SHUTDOWN', 'shutdown_when')
wait_seconds = config.getint('SHUTDOWN', 'wait_seconds')
ori_shutdown_when = shutdown_when
  
# Windows服务器配置  
server_name = config.get('WINDOWS_SERVER', 'server_name')
service_name = config.get('WINDOWS_SERVER', 'service_name')
  
def shutdown_server():  
    # 关闭Windows服务器服务  
    # win32com.client.Dispatch('WbemScripting.SWbemLocator').ConnectServer(server_name, 'root\cimv2')  
    # win32com.client.Dispatch('WbemScripting.SWbemService').OpenService(service_name)  
    # win32com.client.Dispatch('WbemScripting.SWbemService').StopService()  
    params = f'shutdown /s /t {wait_seconds}'
    ctypes.windll.shell32.ShellExecuteW(None, "runas", "cmd.exe", "/c " + params, None, 1)
    print(f"将于 {wait_seconds} 秒后关闭 Windows 服务器，如需取消请使用命令：'shutdown /a'")  
    print(f'注意：只有以管理员模式运行才能自动关机。')  


def seconds_to_hms(seconds):
    hours, remainder = divmod(seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"


# 在代码中增加写入 log 的功能
def write_to_log(message):
    with open('status.log', 'w', encoding='utf-8') as file:
        file.write(message + '\n')

  
def get_nut_status():  
    while True:
        try:
            # 连接到NUT服务器并获取UPS状态  
            client = nut.PyNUTClient(host=nut_host, port=nut_port, login=nut_user, password=nut_password)  
            status = client.list_vars(ups_device)  
            # print(status)
            return status  
        except Exception as e:
            print(f"获取 NUT 服务器状态失败: {e}. 1 分钟后重试...")
            time.sleep(60)
    # example = {'battery.charge': '100', 'battery.charge.low': '10', 'battery.mfr.date': '2023/11/14',
    #  'battery.runtime': '1283', 'battery.runtime.low': '120', 'battery.type': 'PbAc',
    #  'battery.voltage': '13.6', 'battery.voltage.nominal': '12.0', 'device.mfr': 'American Power Conversion', 
    # 'device.model': 'Back-UPS BK650M2-CH', 'device.serial': 'n2201001020', 'device.type': 'ups', 
    # 'driver.name': 'usbhid-ups', 'driver.parameter.pollfreq': '30', 'driver.parameter.pollinterval': '5', 
    # 'driver.parameter.port': 'auto', 'driver.parameter.synchronous': 'no',
    #  'driver.version': 'DSM7-2-1-NewModel-repack-64570-230831', 'driver.version.data': 'APC HID 0.96',
    #  'driver.version.internal': '0.41', 'input.sensitivity': 'high',
    #  'input.transfer.high': '256', 'input.transfer.low': '196', 
    # 'input.transfer.reason': 'input voltage out of range',
    #  'input.voltage': '222.0', 'input.voltage.nominal': '220', 
    # 'ups.beeper.status': 'enabled', 'ups.delay.shutdown': '20', 'ups.firmware': '2333A237-292804G ',
    #  'ups.load': '30', 'ups.mfr': 'American Power Conversion', 'ups.mfr.date': '2004/06/08', 
    # 'ups.model': 'Back-UPS BK650M2-CH', 'ups.productid': '0002', 'ups.realpower.nominal': '390',
    #  'ups.serial': 'n2201001020', 'ups.status': 'OL', 'ups.test.result': 'Done and passed',
    #  'ups.timer.reboot': '0', 'ups.timer.shutdown': '-1', 'ups.vendorid': '051d'}
  
def main():  
    global shutdown_when
    global ori_shutdown_when

    # 如果启动的时候电量就低于关机电量
    # 获取UPS状态  
    ups_status = get_nut_status()  
    battery_percent = int(ups_status['battery.charge'])
    # 检查UPS电源剩余百分比  
    if battery_percent <= shutdown_when:
        shutdown_when = max(5, battery_percent - (100 - shutdown_when))

    while True:  
        # 获取UPS状态  
        ups_status = get_nut_status()  
        battery_percent = int(ups_status['battery.charge'])

        # 如果电量回上去了，恢复关机电量值
        if battery_percent > ori_shutdown_when:
            shutdown_when = ori_shutdown_when
  
        # 检查UPS电源剩余百分比  
        if battery_percent < shutdown_when:  
            shutdown_server()
            break
        else:  
            os.system('cls' if os.name == 'nt' else 'clear')

            #print(ups_status)
            
            text = f"检查时间：{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}（检查周期： {wait_seconds} 秒）\n"
            text += f"UPS 型号：{ups_status['device.mfr']} {ups_status['device.model']}\n"
            text += (f"UPS 状态：{ups_status['ups.status']}，UPS 测试结果：{ups_status['ups.test.result']}\n")
            text += ('\n')
            text += (f"UPS 剩余电量：{battery_percent}%，剩余可使用时间：{seconds_to_hms(int(ups_status['battery.runtime']))}\n")
            text += (f"输入电压：{ups_status['input.voltage']}V/{ups_status['input.voltage.nominal']}V，电池电压：{ups_status['battery.voltage']}V/{ups_status['battery.voltage.nominal']}V\n")
            text += (f"功率：{round(int(ups_status['ups.realpower.nominal'])*int(ups_status['ups.load'])/100,2)}W/{ups_status['ups.realpower.nominal']}W，负载：{ups_status['ups.load']}%\n")
            text += ('\n')

            # 将 status 转换为 DataFrame
            # status_df = pd.DataFrame(ups_status.items(), columns=['参数', '数值'])
            # print(status_df)
            text += (f'UPS电源剩余百分比大于 {shutdown_when}%（{ori_shutdown_when}%），无需关闭 Windows 服务器。\n')  
            text += (f'注意：只有以管理员模式运行才能自动关机。\n')  

            print(text)
            write_to_log(text)
  
        # 等待一分钟后再次检查  
        time.sleep(wait_seconds)  
  
if __name__ == "__main__":  
    main()