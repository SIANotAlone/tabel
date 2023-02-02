
from app import App

print("Перед використанням додатка обов`язково дайте доступ на редагуваняя пошті: \nСкопіюйте пошту: school86@testproj-332512.iam.gserviceaccount.com ")
link = input("Введіть посилання на Google таблицю: ")


# list = input("Введіть назву листа для першого семестру: ")
# list2 = input("Введіть назву листа для другого семестру: ")
# #grades = input("Введіть діапазон з оцінками")
# if list == "": list ="І семестр"
# if list2 =="": list2="ІІ семестр"

#app = App(link="https://docs.google.com/spreadsheets/d/1bqq8MriakhOIG3fS9RSWVp83h71FnOW7Cvy-VYQ3uas/edit#gid=438916944", list="І семестр", range="A1:AP39", list2="ІІ семестр")
app = App(link=link, list="І семестр", range="A1:AP38", list2="ІІ семестр")


if __name__ == "__main__":
    app.start()