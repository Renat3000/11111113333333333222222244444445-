# -*- coding: utf-8 -*-
import vk_api
import xlwt
import time;
import imp; imp.reload(time)
login, password = 'тут логин от вк', 'тут пароль'
vk = vk_api.VkApi(login, password)
try:
    vk.authorization()
except vk_api.AuthorizationError as error_msg:
    print(error_msg)
users = []
audios = []
temp = []
wb = xlwt.Workbook()
ws = wb.add_sheet('Test')

me = "тут id"

f = open('C://vk_api/settings2.txt', 'r')
settings = []
for line in f:
    settings.append(line)
offset = 0
city = 73
country = 1
sex = int(settings[0])
age_from = int(settings[1])
age_to = int(settings[2])
birth_day = int(settings[3])
group = settings[4].lower()
has_photo = 1
if age_from == age_to == 0:
        response = vk.method("users.search", {"city":city,"country":country,"sex":sex,"birth_day":birth_day,"has_photo":has_photo,"offset":offset,"count":"1000"})
        for item in response["items"]:
            users.append(item["id"])

else:
            response = vk.method("users.search", {"city":city,"country":country,"sex":sex,"age_from":age_from,"age_to":age_to,"birth_day":birth_day,"has_photo":has_photo,"offset":offset,"count":"1000"})
            for item in response["items"]:
                users.append(item["id"])
ws.write(0, 0, "VK_id")
ws.write(0, 1, "Artist")
ws.write(0, 2, "Count")
count = 1
save = "C://vk_api/"+str(age_from)+"_"+str(sex)+"_"+str(birth_day)+"_"+group+".xls"

temp = []

for user in users:
    time.sleep(1.5)
    try:
        response = vk.method("audio.get", {"owner_id": user, "count": "6000"})
        for item in response["items"]:
            temp.append(item["artist"].lower())
        print(user, group, temp.count(group))
        ws.write(count, 0, user)
        ws.write(count, 1, group)
        ws.write(count, 2, temp.count(group))
        wb.save(save)
    except:
        print (user, "Access Denied")
        ws.write(count, 0, user)
        ws.write(count, 1, group)
        ws.write(count, 2, "Access Denied")
        wb.save(save)
    temp = []
    count += 1
save = "C://vk_api/"+str(age_from)+"_"+str(sex)+"_"+str(birth_day)+"_"+group+".xls",
wb.save(save)
f.close()
