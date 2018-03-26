import requests
import json
import hashlib
import base64

def getcode(data):
    '''根据拼接后的json生成校验码
       XML需要是unicode
    '''
    data = data + "A2651F123E6CDB113D"
    temp = data.encode("utf-8")
    md5 = hashlib.md5()
    md5.update(temp)
    md5str = md5.digest()  # 16位
    b64str = base64.b64encode(md5str)
    return b64str

class ZhongtongDto:
    #initialise
    def __init__(self,id,logisticsId,optDate,optMan,remark,weight,zone):
        self.id=id
        self.logisticsId=logisticsId
        self.optDate=optDate
        self.optMan=optMan
        self.remark=remark
        self.weight=weight
        self.zone=zone
        self.platformSource=10647
        self.warehouseCode="au003"
        
        
if __name__=="__main__":
    test=ZhongtongDto(150,"120002327209","2017-06-21 12:12","白孟阳","测试",1.5,"0")
    data=json.dumps(test, default=lambda obj: obj.__dict__)
    print(data)
    headers={"Content-Type":"application/x-www-form-urlencoded","charset":"GBK"}
    payload = {'data': data, 'msg_type': 'zto.logistics.tracksInfo',"data_digest":getcode(data),"company_id":"AZEWE0952147E106"}
    r = requests.post("http://intltest.zto.cn/api/import/init",headers=headers ,data=payload)
    print(r.text)