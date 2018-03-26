# encoding=utf-8
import jieba
import re

def check(text):
    def __isNum(word):
        pattern = re.compile(r'[0-9]+$')
        if(pattern.match(word)):
            return True
        numList=["一","二","三","四","五","六","七","八","九","十"]
        for tmpWord in word:
            if tmpWord in numList:
                return True
        return False
                
    
    def __checkNum(Num):
        return __isNum(seg_list[seg_list.index(word)-1])
        
    seg_list = list(jieba.cut(test))#分词
    
    print("seg_list:"+", ".join(seg_list))
    
    keyword1 = ["组合","*","，"]#这组里面如果存在无脑判断是多品
    for word in keyword1:
        if word in seg_list:
            return True
                
                
    keyword2 = ["盒装","支","盒","袋","包","个","罐","罐装","瓶"] #这组里面如果存在还需判断之前之后的内容
    for word in keyword2:
        if word in seg_list:
            if(__checkNum(word)):
                return True

    keyword3 =["+"]#这组还要check之前会不会是d
    for word in keyword3:
        if word in seg_list:
            if not seg_list[seg_list.index(word)-1].upper() == "D":
                return True
                
    keyword2 = ["支","盒","袋","包","个","罐","瓶"]
    big_regex = re.compile('|'.join(map(re.escape, keyword2)))
    print(text)
    text = big_regex.sub("|", text)
    print(text)
    seg_list = list(jieba.cut(text))#分词
    print("seg_list:"+", ".join(seg_list))
    for word in seg_list:
        if __isNum(word):
            return True
    return False
    
            
test = "Aerogard无香无刺激防蚊喷雾135ml 滚珠50ml"
if check(test):
    print("yeah")
else:
    print("no")