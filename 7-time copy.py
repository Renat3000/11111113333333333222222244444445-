# -*- coding: utf-8 -*-
import vk_api
import xlwt
import time;
import imp; imp.reload(time)


login, password = 'login', 'parole'
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

me = "id"

f = open('C://vk_api/settings1.txt', 'r')
settings = []
for line in f:
    settings.append(line)
city = 73
country = 1
sex = int(settings[0])
age_from = int(settings[1])
age_to = int(settings[2])
birth_day = int(settings[3])
group =  settings[4]
has_photo = 1
offset = 0
if age_from == age_to == 0:
        response = vk.method("users.search", {"city":city, "country":country, "sex": sex, "birth_day": birth_day,"has_photo": has_photo, "offset": offset, "count": "1000"})
        if response['items']:
            for i in range(len(response["items"])):
                users.append(response['items'][i]["id"])

else:
            response = vk.method("users.search", {"city":city, "country":country, "sex": sex, "age_from": age_from, "age_to": age_to, "birth_day": birth_day,"has_photo": has_photo, "offset": offset, "count": "1000"})
            for i in range(len(response["items"])):
                users.append(response['items'][i]["id"])
ws.write(0, 0, "VK_id")
ws.write(0, 1, "Artist")
ws.write(0, 2, "Count")
count = 1
save = "C://vk_api/18_1_2"+group+".xls"

for user in users:
    try:
        response = vk.method("audio.get", {"owner_id": user, "count": "6000"})
        if response['items']:
            time.sleep(4)
            for i in range(len(response["items"])):
                temp.append((response['items'][i]["artist"]))
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
save = "C://vk_api/18_1_2"+group+".xls",
wb.save(save)
f.close()