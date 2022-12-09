import os
path = os.getcwd()+'\core\chromedriver.exe'
# print(path)
# cmd = 'setx Path=%Path%;{}'.format(path)
cmd = 'set -m path=%path%;{}'.format(path)
print(cmd)
# res = os.popen(cmd)
# output_str = res.read()   # 获得输出字符串
# print(output_str)
