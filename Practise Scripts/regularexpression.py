# encoding: UTF-8
import re
 
# 将正则表达式编译成Pattern对象
pattern = re.compile(r'[A-Z0-9]+')
 
# 使用Pattern匹配文本，获得匹配结果，无法匹配时将返回None
match = pattern.match('箱号')
print(match)
 
if match:
    # 使用Match获得分组信息
    print(match.group())
    
### 输出 ###
# hello