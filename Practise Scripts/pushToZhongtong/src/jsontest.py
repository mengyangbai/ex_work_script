import json

class ZhongtongDto:
    #initialise
    def __init__(self,id,logisitcsId,optDate,optMan,remark,weight,zone):
        self.id=id
        self.logisitcsId=logisitcsId
        self.optDate=optDate
        self.optMan=optMan
        self.remark=remark
        self.weight=weight
        self.zone=zone
        self.platformSource=10647
        self.warehouseCode="au003"
        

if __name__ == '__main__':        
    test=ZhongtongDto(150,"120002327209","2017-06-21 12:12","白孟阳","测试",1.5,"0")
    testJson=json.dumps(test, default=lambda obj: obj.__dict__)
    print(testJson)