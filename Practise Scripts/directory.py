import os
FILE_PATH='test'
if not os.path.isdir(FILE_PATH):
        print('dir not exists')
        os.mkdir(FILE_PATH)