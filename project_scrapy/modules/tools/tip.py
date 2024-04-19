import sys

def tip(platform_name):
    platform_name = platform_name.split('_')[0]
    if platform_name == 'jd':
        tip_title = f'请确保已经完成以下准备：\n1.关闭VPN；\n2.删除data/{platform_name}文件夹内所有的xlsx文件；\n3.删除data/{platform_name}文件夹下merge文件夹及其内部所有文件；\n4.在模拟浏览器上登录京东；\n5.保持模拟浏览器最大化；\n【Y/N】：'
    elif platform_name == 'tb':
        tip_title = f'请确保已经完成以下准备：\n1.关闭VPN；\n2.删除data/{platform_name}文件夹内所有的xlsx文件；\n3.删除data/{platform_name}文件夹下merge文件夹及其内部所有文件；\n4.在模拟浏览器上登录淘宝；\n5.保持模拟浏览器最大化并且无遮挡；\n【Y/N】：'
    elif platform_name == 'pdd':
        tip_title = f'请确保已经完成以下准备：\n1.关闭VPN；\n2.删除data/{platform_name}文件夹内所有的xlsx文件；\n3.删除data/{platform_name}文件夹下merge文件夹及其内部所有文件；\n【Y/N】：'
    elif platform_name == '1688':
        tip_title = f'请确保已经完成以下准备：\n1.关闭VPN；\n2.删除data/{platform_name}文件夹内所有的xlsx文件；\n3.删除data/{platform_name}文件夹下merge文件夹及其内部所有文件；\n4.在模拟浏览器上登录1688；\n5.保持模拟浏览器最大化并且无遮挡；\n【Y/N】：'
    else:
        tip_title = '暂无此平台'
    ready_or_not = input(tip_title)
    if ready_or_not != 'Y':
        sys.exit()