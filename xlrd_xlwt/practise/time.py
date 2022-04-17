import datetime

now_time = datetime.datetime.now()

print(now_time)

now_time = now_time.strftime("%Y/%m/%d %H:%M:%S")

now_time =datetime.datetime.strptime(now_time,r"%Y/%m/%d %H:%M:%S")

print("time of now:",now_time)

print(type(now_time))

end_time = "2021/9/19 22:12:46"

print(type(end_time))

end_time = datetime.datetime.strptime(end_time,r"%Y/%m/%d %H:%M:%S")

print(end_time)

print(type(end_time))

index = end_time -now_time

print("index = ",index)




