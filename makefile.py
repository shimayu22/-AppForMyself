import time

for i in range(10):
    a = str(i)
    filename = a.zfill(6)
    with open("{}.txt".format(filename),"w") as f:
        f.write(a)
    
    time.sleep(1)
