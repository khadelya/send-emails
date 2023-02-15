import win32com.client as win32
from string import Template
import pandas

SUBJECT = "Subject text"


def get_users_info(excel_file_name):
    excel_data_df = pandas.read_excel(excel_file_name)
    name = excel_data_df["Имя"].tolist()
    login = excel_data_df["Логин"].tolist()
    mail_to = excel_data_df["Почта"].tolist()
    password = excel_data_df["Пароль"].tolist()
    indices = excel_data_df.index
    return indices, name, login, mail_to, password


def get_template(file_name):
    with open(file_name, "r", encoding="utf-8") as file:
        msg_template = file.read()
    return Template(msg_template)


def main():
    outlook = win32.Dispatch("outlook.application")
    indices, name, login, mail_to, password = get_users_info("user_data.xlsx")
    message = get_template("template.txt")
    for i in indices:
        mail = outlook.CreateItem(0)
        mail.To = mail_to[i]
        mail.Subject = SUBJECT
        mail.Body = message.substitute(login=login[i], pswd=password[i])
        mail.Send()


if __name__ == "__main__":
    main()
