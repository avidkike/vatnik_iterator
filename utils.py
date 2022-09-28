import datetime
import time
from time import strptime
import vk_api
import xlsxwriter


def export_to_xls(data_list, user_id):
    if data_list.__len__() == 0:
        print('No data to convert!')
        return

    now = datetime.datetime.now()
    dt_string = now.strftime("%Y%m%d-%H%M%S")

    workbook = xlsxwriter.Workbook(f"{user_id}_{dt_string}.xlsx")
    worksheet = workbook.add_worksheet()

    row = 0

    for line in data_list:

        col = 0

        for field in line:
            worksheet.write(row, col, line.get(field))
            col += 1
        row += 1

    workbook.close()


def authorize():
    try:
        with open(r'login.txt', 'r', encoding='utf-8', errors='ignore') as login_file:
            login = login_file.readlines()[0]
    except FileNotFoundError:
        login = input("\nEnter login: ")

    try:
        with open(r'password.txt', 'r', encoding='utf-8', errors='ignore') as password_file:
            password = password_file.readlines()[0]
    except FileNotFoundError:
        password = input("\nEnter password: ")

    vk_session = vk_api.VkApi(login, password, captcha_handler=captcha_handler)

    try:
        vk_session.auth()
    except vk_api.exceptions.Captcha as captcha:
        sid = captcha.sid  # Get sid
        print(sid)
        captcha.get_url()  # Get ref of captcha image
        captcha.get_image()  # Get captcha image (jpg)

    try:
        vk_session.auth(token_only=True)
    except vk_api.AuthError as error_msg:
        print(error_msg)
        return

    open('login.txt', mode='w').write(login)
    open('password.txt', mode='w').write(password)

    return vk_session.get_api()


def get_reference(post, comment):

    if len(comment['parents_stack']):
        reference = f"https://vk.com/wall{post['owner_id']}_{comment['post_id']}?" \
                    f"reply={comment['id']}&thread={comment['parents_stack'][0]}"
    else:
        reference = f"https://vk.com/wall{post['owner_id']}_{comment['post_id']}?reply={comment['id']}"
    return reference


def filter_comments(comments, user, post):
    user_comments = list()

    for comment in comments:

        if user['id'] != comment['from_id']:
            continue

        user_comments.append({
            'reply_text': comment['text'],
            'group_id': -post['owner_id'],
            'post_id': comment['post_id'],
            'post_date': post['date'],
            'post_text': post['text'],
            'reply_id': comment['id'],
            'reference': get_reference(post, comment),
            'user_id': comment['from_id'],
            'reply_date': comment['date']
        })

        print('Comment: ' + comment['text'])
    return user_comments


def define_user(vk):
    user = dict()
    user_id = input('Enter user ID:\n')
    try:
        user = vk.users.get(user_ids=user_id, fields='city,connections,contacts,country,domain,education,screen_name')
        print(user)
    except vk_api.AccessDenied as error_msg:
        print(error_msg)
    return user


def define_date(from_date):
    date = datetime.datetime.fromtimestamp(from_date)
    try:
        ans = strptime(input(f"\nDefault from-date is {date.date().strftime('%d.%m.%Y').__str__()} ({from_date})"
                             f" \nInput new date to change it, otherwise input 'n'\n"), '%d.%m.%Y')

        from_date = int(time.mktime(ans))
        print(f"New timestamp is {from_date}")
    except ValueError:
        print("Date was not changed")
    return from_date


def define_groups(vk, user):
    user_groups = set()
    group_set = set()
    try:
        with open(r'groups.txt', 'r', encoding='utf-8', errors='ignore') as groups_file:
            file_list = groups_file.readlines()
            if len(file_list):
                group_list = vk.groups.getById(group_ids=','.join(file_list))
                for group in group_list:
                    group_set.add(group['id'])
    except FileNotFoundError:
        open('groups.txt', mode='w')

    ans = input("\nDo you want to add user profile groups?\n")
    if ans == 'Y' or ans == 'y':
        try:
            user_groups = set(vk.groups.get(user_id=user['id'])['items'])
        except vk_api.exceptions.ApiError as error_msg:
            print(error_msg)
            user_groups = set()

    group_set.update(user_groups)

    return group_set


def captcha_handler(captcha):
    key = input("Enter captcha code {0}: ".format(captcha.get_url())).strip()
    return captcha.try_again(key)
