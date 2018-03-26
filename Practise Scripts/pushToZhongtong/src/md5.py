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
    
