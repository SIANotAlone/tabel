
from app import App


link = input("Введіть посилання на Google таблицю: ")


# list = input("Введіть назву листа для першого семестру: ")
# list2 = input("Введіть назву листа для другого семестру: ")
# #grades = input("Введіть діапазон з оцінками")
# if list == "": list ="І семестр"
# if list2 =="": list2="ІІ семестр"

#app = App(link="https://docs.google.com/spreadsheets/d/1bqq8MriakhOIG3fS9RSWVp83h71FnOW7Cvy-VYQ3uas/edit#gid=438916944", list="І семестр", range="A1:AP39", list2="ІІ семестр")
app = App(link=link, list="І семестр", range="A1:AP39", list2="ІІ семестр")


if __name__ == "__main__":
    app.start()