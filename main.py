
from app import App



app = App(link="https://docs.google.com/spreadsheets/d/1bqq8MriakhOIG3fS9RSWVp83h71FnOW7Cvy-VYQ3uas/edit#gid=438916944", list="І семестр", range="A1:AP39", list2="ІІ семестр")

if __name__ == "__main__":
    app.start()